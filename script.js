// ─── Состояние приложения (изолировано, без глобальных переменных) ───────────
const AppState = {
    csvContent: "",
    rejectedCsvContent: "",
    modifiedRows: [],
    rejectedRows: [],
    reset() {
        this.csvContent = "";
        this.rejectedCsvContent = "";
        this.modifiedRows = [];
        this.rejectedRows = [];
    }
};

// ─── DOM ─────────────────────────────────────────────────────────────────────
const dropzone        = document.getElementById('dropzone');
const fileInput       = document.getElementById('fileInput');
const reportContainer = document.getElementById('report-container');

// ─── Утилиты ─────────────────────────────────────────────────────────────────

// Экранирование HTML — защита от XSS при вставке данных из файла в innerHTML
function esc(str) {
    return String(str ?? "")
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

// Скачивание CSV: BOM для совместимости с Excel + освобождение ObjectURL
function download(content, filename) {
    const blob = new Blob(["\ufeff", content], { type: 'text/csv;charset=utf-8;' });
    const url  = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href     = url;
    link.download = filename;
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 100); // освобождаем память
}

// Показ ошибки прямо в UI вместо блокирующего alert()
function showError(message) {
    reportContainer.style.display = 'block';
    reportContainer.innerHTML = `
        <div class="error-banner">
            <span class="error-icon">⚠️</span>
            <span>${esc(message)}</span>
        </div>`;
}

// ─── Drag & Drop ──────────────────────────────────────────────────────────────
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(evt => {
    dropzone.addEventListener(evt, e => { e.preventDefault(); e.stopPropagation(); });
});
dropzone.addEventListener('dragover',  () => dropzone.classList.add('dragover'));
dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
dropzone.addEventListener('drop', e => {
    dropzone.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
});
fileInput.addEventListener('change', e => {
    if (e.target.files[0]) handleFile(e.target.files[0]);
});

// ─── Чтение файла ─────────────────────────────────────────────────────────────
function handleFile(file) {
    if (!file.name.match(/\.(xlsx|xls|csv)$/i)) {
        showError("Неподдерживаемый формат. Загрузите файл Excel (.xlsx, .xls) или .csv");
        return;
    }

    const reader = new FileReader();
    reader.onload = e => {
        try {
            const data = new Uint8Array(e.target.result);
            // cellDates:true — SheetJS распознаёт даты как JS Date-объекты
            // НЕ передаём raw:true здесь, т.к. это конфликтует с cellDates
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];

            // Нормализация дат → ДД.ММ.ГГГГ
            for (const addr in sheet) {
                if (addr[0] === '!') continue;
                const cell = sheet[addr];

                if (cell.v instanceof Date) {
                    // JS Date (пришёл через cellDates:true)
                    const d = cell.v;
                    const day   = String(d.getDate()).padStart(2, '0');
                    const month = String(d.getMonth() + 1).padStart(2, '0');
                    const year  = d.getFullYear();
                    cell.v = `${day}.${month}.${year}`;
                    cell.w = cell.v;
                    cell.t = 's';
                } else if (typeof cell.v === 'string' && /^\d{4}-\d{2}-\d{2}/.test(cell.v)) {
                    // ISO-строка "2026-03-30" или "2026-03-30T00:00:00"
                    const parts = cell.v.split('T')[0].split('-');
                    cell.v = `${parts[2]}.${parts[1]}.${parts[0]}`;
                    cell.w = cell.v;
                    cell.t = 's';
                }
            }

            // raw:false — числа форматируются в строки (ИНН/телефон без экспоненты)
            const json = XLSX.utils.sheet_to_json(sheet, {
                header: 1,
                defval: "",
                raw: false
            });
            processData(json);
        } catch (err) {
            showError(`Ошибка при чтении файла: ${err.message}`);
        }
    };
    reader.onerror = () => showError("Не удалось прочитать файл.");
    reader.readAsArrayBuffer(file);
}

// ─── Обработка данных ─────────────────────────────────────────────────────────
function processData(data) {
    if (data.length === 0) { showError("Файл пустой."); return; }

    AppState.reset();

    const headers = data[0];
    const rows    = data.slice(1);

    const findIdx = name =>
        headers.findIndex(h => String(h ?? "").trim().toLowerCase() === name.toLowerCase());

    const colNameIdx    = findIdx("Название компании");
    const colInnIdx     = findIdx("ИНН");
    const colPhoneIdx   = findIdx("Контактный телефон");
    const colContactIdx = findIdx("Имя контакта");
    const colEmailIdx   = findIdx("Рабочий e-mail");

    // Проверка обязательных колонок
    const required = [
        [colNameIdx,    "Название компании"],
        [colInnIdx,     "ИНН"],
        [colPhoneIdx,   "Контактный телефон"],
        [colContactIdx, "Имя контакта"],
    ];
    const missing = required.filter(([idx]) => idx === -1).map(([, name]) => `"${name}"`);
    if (missing.length > 0) {
        showError(`В файле не найдены обязательные колонки: ${missing.join(', ')}`);
        return;
    }

    const validRows = [headers];

    // Regex — объявляем один раз, не внутри цикла
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    // Телефон: от 7 до 20 символов, только цифры/пробелы/+/скобки/дефис
    const phoneRegex = /^[\d\s+\-().]{7,20}$/;
    const quoteRegex = /["'«»„"]/g;
    const llcRegex   = /ОБЩЕСТВО\s+С\s+ОГРАНИЧЕНН[А-ЯЁA-Z]*\s+ОТВЕТСТВЕННОСТ[А-ЯЁA-Z]*/giu;
    const ipRegex    = /ИНДИВИДУАЛЬН[А-ЯЁA-Z]*\s+ПРЕДПРИНИМАТЕЛ[А-ЯЁA-Z]*/giu;

    rows.forEach((originalRow, index) => {
        const rowNum = index + 2; // +1 заголовок, +1 Excel 1-based
        const row = originalRow.map(c => String(c ?? "").trim());

        // Пропускаем полностью пустые строки
        if (row.every(c => c === "")) return;

        const nameVal    = row[colNameIdx];
        const innVal     = row[colInnIdx];
        const phoneVal   = row[colPhoneIdx];
        const contactVal = row[colContactIdx];
        const emailVal   = colEmailIdx !== -1 ? row[colEmailIdx] : "";

        // ── Валидация ───────────────────────────────────────────────────────
        const errors = [];
        if (!nameVal)  errors.push("Пустое название компании");
        if (!innVal)   errors.push("Нет ИНН");
        if (!phoneVal && !contactVal) errors.push("Нет контактных данных (телефон и имя оба пусты)");
        if (phoneVal && !phoneRegex.test(phoneVal)) errors.push("Некорректный формат телефона");
        if (emailVal && !emailRegex.test(emailVal)) errors.push("Некорректный e-mail");

        if (errors.length > 0) {
            AppState.rejectedRows.push({ rowNum, reason: errors.join("; "), rowData: row });
            return;
        }

        // ── Клининг ─────────────────────────────────────────────────────────
        const cleanName = nameVal
            .replace(quoteRegex, '')
            .replace(llcRegex, 'ООО')
            .replace(ipRegex, 'ИП')
            .toUpperCase()
            .replace(/\s{2,}/g, ' ')
            .trim();

        const cleanInn = innVal.replace(/["'«»„"\s]/g, '');

        const nameChanged = cleanName !== nameVal;
        const innChanged  = cleanInn  !== innVal;

        if (nameChanged || innChanged) {
            row[colNameIdx] = cleanName;
            row[colInnIdx]  = cleanInn;

            // Логируем каждое изменение отдельно — не теряем изменения ИНН
            const changes = [];
            if (nameChanged) changes.push({ field: 'Название', oldVal: nameVal, newVal: cleanName });
            if (innChanged)  changes.push({ field: 'ИНН',      oldVal: innVal,  newVal: cleanInn  });

            AppState.modifiedRows.push({ rowNum, changes, rowData: [...row] });
        }

        validRows.push(row);
    });

    // CSV формируем до renderUI, чтобы кнопка скачать была готова сразу
    const ws = XLSX.utils.aoa_to_sheet(validRows);
    AppState.csvContent = XLSX.utils.sheet_to_csv(ws, { FS: ";" });

    if (AppState.rejectedRows.length > 0) {
        const rejData = AppState.rejectedRows.map(r => {
            const obj = { "Номер строки": r.rowNum, "Причина отклонения": r.reason };
            headers.forEach((h, i) => { obj[h] = r.rowData[i]; });
            return obj;
        });
        const wsRej = XLSX.utils.json_to_sheet(rejData);
        AppState.rejectedCsvContent = XLSX.utils.sheet_to_csv(wsRej, { FS: ";" });
    }

    renderUI(headers, validRows.length - 1);
}

// ─── Рендер UI ────────────────────────────────────────────────────────────────
function renderUI(headers, successCount) {
    const { modifiedRows, rejectedRows } = AppState;
    const hasModified = modifiedRows.length > 0;
    const hasRejected = rejectedRows.length > 0;

    // Динамический статус — не захардкожен в HTML
    const statusIcon = hasRejected ? '⚠️' : '✅';
    const statusText = hasRejected
        ? `Проверка завершена: ${rejectedRows.length} ${plural(rejectedRows.length, 'строка отклонена', 'строки отклонены', 'строк отклонено')}`
        : 'Все строки прошли валидацию';

    reportContainer.style.display = 'block';
    reportContainer.innerHTML = `
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label">Принято</div>
                <div class="stat-value" style="color:var(--success)">${successCount}</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Исправлено</div>
                <div class="stat-value" style="color:var(--warning)">${modifiedRows.length}</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Отклонено</div>
                <div class="stat-value" style="color:var(--danger)">${rejectedRows.length}</div>
            </div>
        </div>

        <div class="status-message ${hasRejected ? 'has-errors' : ''}">${statusIcon} ${esc(statusText)}</div>

        <div class="main-action">
            <button id="downloadBtn" class="btn-primary">
                <svg class="icon-sm" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3"/>
                </svg>
                Скачать CSV
            </button>
        </div>

        <div class="view-controls">
            ${hasModified ? '<button id="showModifiedBtn" class="btn-outline">Показать исправления</button>' : ''}
            ${hasRejected ? '<button id="showRejectedBtn" class="btn-outline">Показать отклонённые</button>' : ''}
        </div>

        ${hasModified ? `
        <div id="modified-details" class="details-section" style="display:none;">
            <h4>📝 Лог исправлений:</h4>
            <div class="table-wrapper">${buildModifiedTable(headers)}</div>
        </div>` : ''}

        ${hasRejected ? `
        <div id="rejected-details" class="details-section" style="display:none;">
            <div class="section-header">
                <h4>❌ Отклонённые строки:</h4>
                <button id="downloadRejectedBtn" class="btn-secondary btn-small">
                    <svg class="icon-xs" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3"/>
                    </svg>
                    Скачать отклонённые (.csv)
                </button>
            </div>
            <div class="table-wrapper">${buildRejectedTable(headers)}</div>
        </div>` : ''}
    `;

    // Обработчики навешиваем после рендера
    document.getElementById('downloadBtn')
        ?.addEventListener('click', () => download(AppState.csvContent, 'cleaned_data.csv'));
    document.getElementById('downloadRejectedBtn')
        ?.addEventListener('click', () => download(AppState.rejectedCsvContent, 'rejected_rows.csv'));

    setupTabs();
}

// ─── Таблицы ──────────────────────────────────────────────────────────────────

function buildModifiedTable(headers) {
    const { modifiedRows } = AppState;
    if (modifiedRows.length === 0) return '<p class="empty-msg">Нет исправлений.</p>';

    // Заголовок: Строка | Поле | Было | Стало | ...колонки данных
    let html = `<table><thead><tr>
        <th>Строка</th><th>Поле</th><th>Было</th><th>Стало</th>
        ${headers.map(h => `<th>${esc(h)}</th>`).join('')}
    </tr></thead><tbody>`;

    modifiedRows.forEach(item => {
        item.changes.forEach((change, ci) => {
            html += '<tr>';
            // Номер строки и данные строки — через rowspan, чтобы не дублировать
            if (ci === 0) {
                html += `<td rowspan="${item.changes.length}">${esc(item.rowNum)}</td>`;
            }
            html += `
                <td>${esc(change.field)}</td>
                <td><span class="old-val">${esc(change.oldVal)}</span></td>
                <td><span class="new-val">${esc(change.newVal)}</span></td>`;
            if (ci === 0) {
                item.rowData.forEach(c => { html += `<td>${esc(c)}</td>`; });
            }
            html += '</tr>';
        });
    });

    return html + '</tbody></table>';
}

function buildRejectedTable(headers) {
    const { rejectedRows } = AppState;
    if (rejectedRows.length === 0) return '<p class="empty-msg">Нет отклонённых строк.</p>';

    let html = `<table><thead><tr>
        <th>Строка</th><th>Причина</th>
        ${headers.map(h => `<th>${esc(h)}</th>`).join('')}
    </tr></thead><tbody>`;

    rejectedRows.forEach(item => {
        html += `<tr>
            <td>${esc(item.rowNum)}</td>
            <td><span class="reason-badge">${esc(item.reason)}</span></td>
            ${item.rowData.map(c => `<td>${esc(c)}</td>`).join('')}
        </tr>`;
    });

    return html + '</tbody></table>';
}

// ─── Табы ─────────────────────────────────────────────────────────────────────
function setupTabs() {
    const mBtn = document.getElementById('showModifiedBtn');
    const rBtn = document.getElementById('showRejectedBtn');
    const mBox = document.getElementById('modified-details');
    const rBox = document.getElementById('rejected-details');

    function resetTabs() {
        if (mBox) mBox.style.display = 'none';
        if (rBox) rBox.style.display = 'none';
        mBtn?.classList.remove('active');
        rBtn?.classList.remove('active');
    }

    mBtn?.addEventListener('click', () => {
        const isActive = mBtn.classList.contains('active');
        resetTabs();
        if (!isActive) { mBox.style.display = 'block'; mBtn.classList.add('active'); }
    });

    rBtn?.addEventListener('click', () => {
        const isActive = rBtn.classList.contains('active');
        resetTabs();
        if (!isActive) { rBox.style.display = 'block'; rBtn.classList.add('active'); }
    });
}

// ─── Утилита: склонение числительных ─────────────────────────────────────────
function plural(n, one, few, many) {
    const mod10  = n % 10;
    const mod100 = n % 100;
    if (mod10 === 1 && mod100 !== 11) return one;
    if (mod10 >= 2 && mod10 <= 4 && (mod100 < 10 || mod100 >= 20)) return few;
    return many;
}
