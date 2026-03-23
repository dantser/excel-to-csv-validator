const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const downloadBtn = document.getElementById('downloadBtn');
const reportContainer = document.getElementById('report-container');

// Обработка перетаскивания
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    dropzone.addEventListener(eventName, e => {
        e.preventDefault();
        e.stopPropagation();
    }, false);
});

['dragenter', 'dragover'].forEach(eventName => {
    dropzone.addEventListener(eventName, () => dropzone.classList.add('dragover'), false);
});

['dragleave', 'drop'].forEach(eventName => {
    dropzone.addEventListener(eventName, () => dropzone.classList.remove('dragover'), false);
});

dropzone.addEventListener('drop', e => {
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
});

fileInput.addEventListener('change', e => {
    if (e.target.files[0]) handleFile(e.target.files[0]);
});

function handleFile(file) {
    if (!file.name.match(/\.(xlsx|xls)$/)) {
        alert("Пожалуйста, выберите Excel файл");
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
        validateAndConvert(jsonData);
    };
    reader.readAsArrayBuffer(file);
}

function validateAndConvert(data) {
    const headers = data[0];
    const rows = data.slice(1);
    const campaignColName = "Название кампании";
    const campaignColIndex = headers.indexOf(campaignColName);

    if (campaignColIndex === -1) {
        alert(`Колонка "${campaignColName}" не найдена!`);
        return;
    }

    let validRows = [headers];
    let rejectedData = [];
    let modifiedCount = 0;

    rows.forEach((row, index) => {
        const isCompletelyEmpty = row.every(cell => String(cell).trim() === "");
        if (isCompletelyEmpty) return;

        let campaignValue = String(row[campaignColIndex] || "").trim();

        if (!campaignValue) {
            rejectedData.push({ rowNum: index + 2, reason: "Пустое название", rowData: row });
            return;
        }

        if (campaignValue.match(/["'«»„“]/)) {
            campaignValue = campaignValue.replace(/["'«»„“]/g, '');
            row[campaignColIndex] = campaignValue;
            modifiedCount++;
        }

        validRows.push(row);
    });

    displayResults(headers, rejectedData, validRows.length - 1, modifiedCount);

    const ws = XLSX.utils.aoa_to_sheet(validRows);
    csvContent = XLSX.utils.sheet_to_csv(ws);
    
    reportContainer.style.display = 'block';
    // Скролл к отчету
    reportContainer.scrollIntoView({ behavior: 'smooth' });
}

function displayResults(headers, rejected, success, modified) {
    const statsGrid = document.getElementById('stats-summary');
    statsGrid.innerHTML = `
        <div class="stat-card">
            <div class="stat-label">Готово к экспорту</div>
            <div class="stat-value" style="color: var(--success)">${success}</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">Исправлено (кавычки)</div>
            <div class="stat-value" style="color: var(--warning)">${modified}</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">Ошибок</div>
            <div class="stat-value" style="color: var(--danger)">${rejected.length}</div>
        </div>
    `;

    const details = document.getElementById('report-details');
    if (rejected.length > 0) {
        let html = `<div class="table-wrapper"><table><thead><tr><th>Строка</th><th>Причина</th>`;
        headers.forEach(h => html += `<th>${h}</th>`);
        html += `</tr></thead><tbody>`;
        rejected.forEach(item => {
            html += `<tr>
                <td>${item.rowNum}</td>
                <td><span class="reason-badge">${item.reason}</span></td>
                ${item.rowData.map(cell => `<td>${cell}</td>`).join('')}
            </tr>`;
        });
        html += `</tbody></table></div>`;
        details.innerHTML = html;
    } else {
        details.innerHTML = "";
    }
}

downloadBtn.addEventListener('click', () => {
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "campaigns_cleaned.csv";
    link.click();
});
