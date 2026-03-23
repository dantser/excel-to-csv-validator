const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const downloadBtn = document.getElementById('downloadBtn');
const reportContainer = document.getElementById('report-container');

let csvContent = ""; 
let modifiedRows = []; // Список измененных
let rejectedRows = []; // Список удаленных

// --- Инициализация ---
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
    if (!file.name.match(/\.(xlsx|xls)$/)) { alert("Нужен Excel файл!"); return; }
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        processData(json);
    };
    reader.readAsArrayBuffer(file);
}

// --- Обработка ---
function processData(data) {
    const headers = data[0];
    const rows = data.slice(1);
    const colIndex = headers.findIndex(h => String(h || "").trim().toLowerCase() === "название компании");

    if (colIndex === -1) { alert("Колонка 'Название компании' не найдена!"); return; }

    let validRowsForExport = [headers];
    modifiedRows = [];
    rejectedRows = [];

    rows.forEach((originalRow, index) => {
        const rowNum = index + 2;
        // 1. Глобальный Trim
        let row = originalRow.map(c => String(c || "").trim());

        if (row.every(c => c === "")) return; // Пропуск пустых

        const originalName = row[colIndex];

        // 2. Валидация на пустоту
        if (!originalName) {
            rejectedRows.push({ rowNum, reason: "Пустое название", rowData: row });
            return;
        }

        // 3. CAPS + Убираем кавычки
        let cleanName = originalName.replace(/["'«»„“]/g, '').toUpperCase();

        if (cleanName !== originalName) {
            modifiedRows.push({ rowNum, oldVal: originalName, newVal: cleanName, rowData: [...row] });
            row[colIndex] = cleanName;
        }

        validRowsForExport.push(row);
    });

    renderUI(headers, validRowsForExport.length - 1);
    
    // Готовим CSV
    const ws = XLSX.utils.aoa_to_sheet(validRowsForExport);
    csvContent = XLSX.utils.sheet_to_csv(ws);
}

function renderUI(headers, successCount) {
    reportContainer.style.display = 'block';
    
    // Статистика
    const stats = document.getElementById('stats-summary');
    stats.innerHTML = `
        <div class="stat-card"><div class="stat-label">В итоге</div><div class="stat-value" style="color:var(--success)">${successCount}</div></div>
        <div class="stat-card"><div class="stat-label">Изменено</div><div class="stat-value" style="color:var(--warning)">${modifiedRows.length}</div></div>
        <div class="stat-card"><div class="stat-label">Удалено</div><div class="stat-value" style="color:var(--danger)">${rejectedRows.length}</div></div>
    `;

    // Кнопки управления
    document.getElementById('showModifiedBtn').style.display = modifiedRows.length > 0 ? 'inline-block' : 'none';
    document.getElementById('showRejectedBtn').style.display = rejectedRows.length > 0 ? 'inline-block' : 'none';

    // Рендерим таблицы
    fillTable('modified-table-container', headers, modifiedRows, true);
    fillTable('rejected-table-container', headers, rejectedRows, false);

    reportContainer.scrollIntoView({ behavior: 'smooth' });
}

function fillTable(containerId, headers, data, isModified) {
    const container = document.getElementById(containerId);
    let html = `<table><thead><tr><th>Стр.</th><th>${isModified ? 'Было ➔ Стало' : 'Причина'}</th>`;
    headers.forEach(h => html += `<th>${h}</th>`);
    html += `</tr></thead><tbody>`;

    data.forEach(item => {
        html += `<tr><td>${item.rowNum}</td>`;
        if (isModified) {
            html += `<td><span class="old-val">${item.oldVal}</span> <span class="new-val">➔ ${item.newVal}</span></td>`;
        } else {
            html += `<td><span class="reason-badge">${item.reason}</span></td>`;
        }
        item.rowData.forEach(c => html += `<td>${c}</td>`);
        html += `</tr>`;
    });
    container.innerHTML = html + `</tbody></table>`;
}

// Кнопки переключения
document.getElementById('showModifiedBtn').onclick = function() {
    this.classList.toggle('active');
    const el = document.getElementById('modified-details');
    el.style.display = el.style.display === 'none' ? 'block' : 'none';
};

document.getElementById('showRejectedBtn').onclick = function() {
    this.classList.toggle('active');
    const el = document.getElementById('rejected-details');
    el.style.display = el.style.display === 'none' ? 'block' : 'none';
};

downloadBtn.onclick = () => {
    const blob = new Blob(["\ufeff", csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `cleaned_caps_data.csv`;
    link.click();
};
