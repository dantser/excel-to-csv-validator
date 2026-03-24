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
        const workbook = XLSX.read(data, { type: 'array', cellNF: true, cellText: true });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        // УМНЫЙ ПАРСЕР ДАТ
        for (let z in sheet) {
            if (z[0] === '!') continue; 
            const cell = sheet[z];
            
            // Если ячейка содержит число (n) и имеет формат даты (z)
            if (cell.t === 'n' && cell.z && /m|d|y/.test(cell.z.toLowerCase())) {
                try {
                    // Используем парсер SSF для извлечения компонентов даты из числа
                    const d = XLSX.SSF.parse_date_code(cell.v);
                    if (d) {
                        // Собираем строку DD.MM.YYYY вручную
                        const day = String(d.d).padStart(2, '0');
                        const month = String(d.m).padStart(2, '0');
                        const year = d.y;
                        cell.v = `${day}.${month}.${year}`;
                        cell.t = 's'; // Меняем тип на строку (string)
                    }
                } catch (err) {
                    // Если дата совсем битая, оставляем как есть или пустую строку
                    cell.v = cell.w || "";
                }
            }
        }

        const json = XLSX.utils.sheet_to_json(sheet, { 
            header: 1, 
            defval: "",
            raw: true // Оставляем ИНН и телефоны "чистыми" числами
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

        let errors = [];
        if (!nameVal) errors.push("Пустое название");
        if (!innVal) errors.push("Нет ИНН");
        if (!phoneVal && !contactVal) errors.push("Нет контактных данных");

        if (errors.length > 0) {
            rejectedRows.push({ rowNum, reason: errors.join(", "), rowData: row });
            return;
        }

        const quoteRegex = /["'«»„“]/g;
        let cleanName = nameVal.replace(quoteRegex, '').toUpperCase();
        let cleanInn = innVal.replace(/["'«»„“\s]/g, '');

        if (cleanName !== nameVal || cleanInn !== innVal) {
            modifiedRows.push({ rowNum, oldVal: nameVal, newVal: cleanName, rowData: [...row] });
            row[colNameIdx] = cleanName;
            row[colInnIdx] = cleanInn;
        }
        validRows.push(row);
    });

    renderUI(headers, validRows.length - 1);
    
    // Экспорт с разделителем ";"
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

const mBtn = document.getElementById('showModifiedBtn'), rBtn = document.getElementById('showRejectedBtn');
const mBox = document.getElementById('modified-details'), rBox = document.getElementById('rejected-details');
function resetTabs() { mBox.style.display = rBox.style.display = 'none'; mBtn.classList.remove('active'); rBtn.classList.remove('active'); }
mBtn.onclick = () => { const act = mBtn.classList.contains('active'); resetTabs(); if(!act) { mBox.style.display = 'block'; mBtn.classList.add('active'); }};
rBtn.onclick = () => { const act = rBtn.classList.contains('active'); resetTabs(); if(!act) { rBox.style.display = 'block'; rBtn.classList.add('active'); }};

function download(content, name) {
    const blob = new Blob(["\ufeff", content], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = name;
    link.click();
}
downloadBtn.onclick = () => download(csvContent, 'cleaned_data.csv');
document.getElementById('downloadRejectedBtn').onclick = () => download(rejectedCsvContent, 'rejected_rows.csv');
