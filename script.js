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

    const colNameIdx = headers.findIndex(h => String(h || "").trim().toLowerCase() === "название компании");
    const colInnIdx = headers.findIndex(h => String(h || "").trim().toLowerCase() === "инн");

    if (colNameIdx === -1 || colInnIdx === -1) {
        alert("Колонки 'Название компании' и 'ИНН' обязательны!");
        return;
    }

    let validRows = [headers];
    modifiedRows = [];
    rejectedRows = [];

    rows.forEach((originalRow, index) => {
        const rowNum = index + 2;
        let row = originalRow.map(c => String(c || "").trim());
        if (row.every(c => c === "")) return;

        const originalName = row[colNameIdx];
        const innValue = row[colInnIdx];

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
            modifiedRows.push({ rowNum, oldVal: originalName, newVal: cleanName, rowData: [...row] });
            row[colNameIdx] = cleanName;
            row[colInnIdx] = cleanInn;
        }
        validRows.push(row);
    });

    renderUI(headers, validRows.length - 1);
    
    // Основной CSV
    const ws = XLSX.utils.aoa_to_sheet(validRows);
    csvContent = XLSX.utils.sheet_to_csv(ws);

    // CSV ошибок
    if (rejectedRows.length > 0) {
        const rejData = rejectedRows.map(r => {
            let obj = { "Номер строки": r.rowNum, "Причина": r.reason };
            headers.forEach((h, i) => obj[h] = r.rowData[i]);
            return obj;
        });
        rejectedCsvContent = XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(rejData));
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

// --- Логика табов ---
const mBtn = document.getElementById('showModifiedBtn'), rBtn = document.getElementById('showRejectedBtn');
const mBox = document.getElementById('modified-details'), rBox = document.getElementById('rejected-details');

function resetTabs() { mBox.style.display = rBox.style.display = 'none'; mBtn.classList.remove('active'); rBtn.classList.remove('active'); }

mBtn.onclick = () => { const act = mBtn.classList.contains('active'); resetTabs(); if(!act) { mBox.style.display = 'block'; mBtn.classList.add('active'); }};
rBtn.onclick = () => { const act = rBtn.classList.contains('active'); resetTabs(); if(!act) { rBox.style.display = 'block'; rBtn.classList.add('active'); }};

// --- Скачивание ---
function download(content, name) {
    const blob = new Blob(["\ufeff", content], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = name;
    link.click();
}

downloadBtn.onclick = () => download(csvContent, 'cleaned_data.csv');
document.getElementById('downloadRejectedBtn').onclick = () => download(rejectedCsvContent, 'rejected_rows.csv');
