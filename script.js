const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const downloadBtn = document.getElementById('downloadBtn');
const reportContainer = document.getElementById('report-container');

let csvContent = ""; 
let modifiedRows = []; 
let rejectedRows = []; 

// --- Drag-and-Drop ---
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
        const json = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1, defval: "" });
        processData(json);
    };
    reader.readAsArrayBuffer(file);
}

function processData(data) {
    if (data.length === 0) return;
    const headers = data[0];
    const rows = data.slice(1);

    const colNameIndex = headers.findIndex(h => String(h || "").trim().toLowerCase() === "название компании");
    const colInnIndex = headers.findIndex(h => String(h || "").trim().toLowerCase() === "инн");

    if (colNameIndex === -1 || colInnIndex === -1) {
        let missing = [];
        if (colNameIndex === -1) missing.push('"Название компании"');
        if (colInnIndex === -1) missing.push('"ИНН"');
        alert(`Ошибка: В файле не найдены колонки: ${missing.join(', ')}`);
        return;
    }

    let validRowsForExport = [headers];
    modifiedRows = [];
    rejectedRows = [];

    rows.forEach((originalRow, index) => {
        const rowNum = index + 2;
        let row = originalRow.map(c => String(c || "").trim());
        if (row.every(c => c === "")) return;

        const originalName = row[colNameIndex];
        const innValue = row[colInnIndex];

        let errors = [];
        if (!originalName) errors.push("Пустое название");
        if (!innValue) errors.push("Отсутствует ИНН");

        if (errors.length > 0) {
            rejectedRows.push({ rowNum, reason: errors.join(", "), rowData: row });
            return;
        }

        const quoteRegex = /["'«»„“]/g;
        let cleanName = originalName.replace(quoteRegex, '').toUpperCase();
        let cleanInn = innValue.replace(/["'«»„“\s]/g, '');

        if (cleanName !== originalName || cleanInn !== innValue) {
            modifiedRows.push({ 
                rowNum, 
                oldVal: originalName + (cleanInn !== innValue ? ` (ИНН: ${innValue})` : ""), 
                newVal: cleanName + (cleanInn !== innValue ? ` (ИНН: ${cleanInn})` : ""), 
                rowData: [...row] 
            });
            row[colNameIndex] = cleanName;
            row[colInnIndex] = cleanInn;
        }
        validRowsForExport.push(row);
    });

    renderUI(headers, validRowsForExport.length - 1);
    const ws = XLSX.utils.aoa_to_sheet(validRowsForExport);
    csvContent = XLSX.utils.sheet_to_csv(ws);
}

function renderUI(headers, successCount) {
    reportContainer.style.display = 'block';
    const stats = document.getElementById('stats-summary');
    stats.innerHTML = `
        <div class="stat-card"><div class="stat-label">В итоге</div><div class="stat-value" style="color:var(--success)">${successCount}</div></div>
        <div class="stat-card"><div class="stat-label">Изменено</div><div class="stat-value" style="color:var(--warning)">${modifiedRows.length}</div></div>
        <div class="stat-card"><div class="stat-label">Удалено</div><div class="stat-value" style="color:var(--danger)">${rejectedRows.length}</div></div>
    `;

    document.getElementById('showModifiedBtn').style.display = modifiedRows.length > 0 ? 'inline-block' : 'none';
    document.getElementById('showRejectedBtn').style.display = rejectedRows.length > 0 ? 'inline-block' : 'none';

    // Сброс состояния табов при новой загрузке
    hideAllDetails();

    fillTable('modified-table-container', headers, modifiedRows, true);
    fillTable('rejected-table-container', headers, rejectedRows, false);
}

function fillTable(containerId, headers, data, isModified) {
    const container = document.getElementById(containerId);
    let html = `<table><thead><tr><th>Номер строки</th><th>${isModified ? 'Было ➔ Стало' : 'Причина'}</th>`;
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

// --- ЛОГИКА ТАБОВ ---

const modBtn = document.getElementById('showModifiedBtn');
const rejBtn = document.getElementById('showRejectedBtn');
const modBox = document.getElementById('modified-details');
const rejBox = document.getElementById('rejected-details');

function hideAllDetails() {
    modBox.style.display = 'none';
    rejBox.style.display = 'none';
    modBtn.classList.remove('active');
    rejBtn.classList.remove('active');
}

modBtn.onclick = function() {
    const isActive = this.classList.contains('active');
    hideAllDetails();
    if (!isActive) {
        modBox.style.display = 'block';
        this.classList.add('active');
    }
};

rejBtn.onclick = function() {
    const isActive = this.classList.contains('active');
    hideAllDetails();
    if (!isActive) {
        rejBox.style.display = 'block';
        this.classList.add('active');
    }
};

downloadBtn.onclick = () => {
    const blob = new Blob(["\ufeff", csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `export_${new Date().getTime()}.csv`;
    link.click();
};
