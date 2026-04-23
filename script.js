const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const downloadBtn = document.getElementById('downloadBtn');
const reportContainer = document.getElementById('report-container');

let csvContent = ""; 
let rejectedCsvContent = "";
let modifiedRows = []; 
let rejectedRows = []; 

// --- Файловая логика ---
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(e => {
    dropzone.addEventListener(e, (ev) => { ev.preventDefault(); ev.stopPropagation(); });
});
dropzone.addEventListener('dragover', () => dropzone.classList.add('dragover'));
dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
dropzone.addEventListener('drop', (e) => {
    dropzone.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files[0]) handleFile(e.target.files[0]);
});

function handleFile(file) {
    if (!file.name.match(/\.(xlsx|xls|csv)$/)) { alert("Нужен Excel файл!"); return; }
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        // Включаем cellDates: true, чтобы библиотека сама распознала даты
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        // --- УМНАЯ ОБРАБОТКА ДАТ ---
        for (let z in sheet) {
            if (z[0] === '!') continue; 
            const cell = sheet[z];

            // 1. Если это объект даты (JS Date)
            if (cell.v instanceof Date) {
                const d = cell.v;
                // Форматируем в ДД.ММ.ГГГГ (с учетом часового пояса)
                const day = String(d.getDate()).padStart(2, '0');
                const month = String(d.getMonth() + 1).padStart(2, '0');
                const year = d.getFullYear();
                cell.v = `${day}.${month}.${year}`;
                cell.t = 's'; 
            } 
            // 2. Если дата пришла строкой типа "2026-03-30" (как в твоем примере)
            else if (typeof cell.v === 'string' && /^\d{4}-\d{2}-\d{2}/.test(cell.v)) {
                const parts = cell.v.split('T')[0].split('-');
                cell.v = `${parts[2]}.${parts[1]}.${parts[0]}`;
                cell.t = 's';
            }
        }

        const json = XLSX.utils.sheet_to_json(sheet, { 
            header: 1, 
            defval: "",
            raw: true // Важно для ИНН и Телефонов
        });
        processData(json);
    };
    reader.readAsArrayBuffer(file);
}

function processData(data) {
    if (data.length === 0) return;
    const headers = data[0];
    const rows = data.slice(1);

    const findIdx = (name) => headers.findIndex(h => String(h || "").trim().toLowerCase() === name.toLowerCase());

    const colNameIdx = findIdx("Название компании");
    const colInnIdx = findIdx("ИНН");
    const colPhoneIdx = findIdx("Контактный телефон");
    const colContactIdx = findIdx("Имя контакта");
    const colEmailIdx = findIdx("Рабочий e-mail");

    if (colNameIdx === -1 || colInnIdx === -1 || colPhoneIdx === -1 || colContactIdx === -1) {
        let missing = [];
        if (colNameIdx === -1) missing.push('"Название компании"');
        if (colInnIdx === -1) missing.push('"ИНН"');
        if (colPhoneIdx === -1) missing.push('"Контактный телефон"');
        if (colContactIdx === -1) missing.push('"Имя контакта"');
        alert(`Ошибка: В файле не найдены колонки: ${missing.join(', ')}`);
        return;
    }

    let validRows = [headers];
    modifiedRows = [];
    rejectedRows = [];

    rows.forEach((originalRow, index) => {
        const rowNum = index + 2;
        let row = originalRow.map(c => String(c || "").trim());
        if (row.every(c => c === "")) return;

        const nameVal = row[colNameIdx];
        const innVal = row[colInnIdx];
        const phoneVal = row[colPhoneIdx];
        const contactVal = row[colContactIdx];
        const emailVal = colEmailIdx !== -1 ? row[colEmailIdx] : "";

        // ВАЛИДАЦИЯ
        let errors = [];
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!nameVal) errors.push("Пустое название");
        if (!innVal) errors.push("Нет ИНН");
        if (!phoneVal && !contactVal) errors.push("Нет контактных данных (тел/имя)");
        if (emailVal && !emailRegex.test(emailVal)) errors.push("Некорректный email");

        if (errors.length > 0) {
            rejectedRows.push({ rowNum, reason: errors.join(", "), rowData: row });
            return;
        }

        // КЛИНИНГ
        const quoteRegex = /["'«»„“]/g;
        // Используем "мягкие" паттерны по основам слов, чтобы срабатывать даже при
        // небольших вариациях/опечатках (например, в окончаниях) в исходном файле.
        // Важно: \w не покрывает кириллицу, поэтому используем явный диапазон букв.
        const llcRegex = /ОБЩЕСТВО\s+С\s+ОГРАНИЧЕНН[А-ЯЁA-Z]*\s+ОТВЕТСТВЕННОСТ[А-ЯЁA-Z]*/giu;
        const ipRegex = /ИНДИВИДУАЛЬН[А-ЯЁA-Z]*\s+ПРЕДПРИНИМАТЕЛ[А-ЯЁA-Z]*/giu;

        let cleanName = nameVal
            .replace(quoteRegex, '')
            .replace(llcRegex, 'ООО')
            .replace(ipRegex, 'ИП')
            .toUpperCase()
            .replace(/\s{2,}/g, ' ')
            .trim();
        let cleanInn = innVal.replace(/["'«»„“\s]/g, '');

        if (cleanName !== nameVal || cleanInn !== innVal) {
            row[colNameIdx] = cleanName;
            row[colInnIdx] = cleanInn;
            modifiedRows.push({ rowNum, oldVal: nameVal, newVal: cleanName, rowData: [...row] });
        }
        validRows.push(row);
    });

    renderUI(headers, validRows.length - 1);
    
    // Формируем CSV с разделителем ";"
    const ws = XLSX.utils.aoa_to_sheet(validRows);
    csvContent = XLSX.utils.sheet_to_csv(ws, { FS: ";" });

    if (rejectedRows.length > 0) {
        const rejData = rejectedRows.map(r => {
            let obj = { "Номер строки": r.rowNum, "Причина": r.reason };
            headers.forEach((h, i) => obj[h] = r.rowData[i]);
            return obj;
        });
        const wsRejected = XLSX.utils.json_to_sheet(rejData);
        rejectedCsvContent = XLSX.utils.sheet_to_csv(wsRejected, { FS: ";" });
    }
}

function renderUI(headers, successCount) {
    reportContainer.style.display = 'block';
    document.getElementById('stats-summary').innerHTML = `
        <div class="stat-card"><div class="stat-label">В итоге</div><div class="stat-value" style="color:var(--success)">${successCount}</div></div>
        <div class="stat-card"><div class="stat-label">Изменено</div><div class="stat-value" style="color:var(--warning)">${modifiedRows.length}</div></div>
        <div class="stat-card"><div class="stat-label">Удалено</div><div class="stat-value" style="color:var(--danger)">${rejectedRows.length}</div></div>
    `;
    document.getElementById('showModifiedBtn').style.display = modifiedRows.length > 0 ? 'inline-block' : 'none';
    document.getElementById('showRejectedBtn').style.display = rejectedRows.length > 0 ? 'inline-block' : 'none';
    resetTabs();
    fillTable('modified-table-container', headers, modifiedRows, true);
    fillTable('rejected-table-container', headers, rejectedRows, false);
}

function fillTable(containerId, headers, data, isMod) {
    const container = document.getElementById(containerId);
    let html = `<table><thead><tr><th>Номер строки</th><th>${isMod ? 'Было ➔ Стало' : 'Причина'}</th>`;
    headers.forEach(h => html += `<th>${h}</th>`);
    html += `</tr></thead><tbody>`;
    data.forEach(item => {
        html += `<tr><td>${item.rowNum}</td>`;
        html += isMod ? `<td><span class="old-val">${item.oldVal}</span> <span class="new-val">➔ ${item.newVal}</span></td>` : `<td><span class="reason-badge">${item.reason}</span></td>`;
        item.rowData.forEach(c => html += `<td>${c}</td>`);
        html += `</tr>`;
    });
    container.innerHTML = html + `</tbody></table>`;
}

// Управление табами
const mBtn = document.getElementById('showModifiedBtn'), rBtn = document.getElementById('showRejectedBtn');
const mBox = document.getElementById('modified-details'), rBox = document.getElementById('rejected-details');
function resetTabs() { mBox.style.display = rBox.style.display = 'none'; mBtn.classList.remove('active'); rBtn.classList.remove('active'); }
mBtn.onclick = () => { const act = mBtn.classList.contains('active'); resetTabs(); if(!act) { mBox.style.display = 'block'; mBtn.classList.add('active'); }};
rBtn.onclick = () => { const act = rBtn.classList.contains('active'); resetTabs(); if(!act) { rBox.style.display = 'block'; rBtn.classList.add('active'); }};

// Скачивание
function download(content, name) {
    const blob = new Blob(["\ufeff", content], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = name;
    link.click();
}
downloadBtn.onclick = () => download(csvContent, 'cleaned_data.csv');
document.getElementById('downloadRejectedBtn').onclick = () => download(rejectedCsvContent, 'rejected_rows.csv');
