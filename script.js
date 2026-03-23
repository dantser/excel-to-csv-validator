const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const downloadBtn = document.getElementById('downloadBtn');
const reportContainer = document.getElementById('report-container');

let csvContent = ""; 

// --- Настройка Drag-and-Drop ---
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    dropzone.addEventListener(eventName, e => { e.preventDefault(); e.stopPropagation(); }, false);
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
        alert("Пожалуйста, выберите файл Excel");
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

// --- Основная логика ---
function validateAndConvert(data) {
    if (data.length === 0) return;

    const headers = data[0];
    const rows = data.slice(1);
    
    // Ищем колонку "Название компании"
    const targetHeader = "название компании";
    const colIndex = headers.findIndex(h => String(h || "").trim().toLowerCase() === targetHeader);

    if (colIndex === -1) {
        alert(`Ошибка: Колонка "Название компании" не найдена!`);
        return;
    }

    let validRows = [headers];
    let rejectedData = [];
    let modifiedCount = 0;

    rows.forEach((originalRow, index) => {
        const rowNum = index + 2;

        // 1. Глобальный Trim (чистим пробелы во всей строке)
        let row = originalRow.map(cell => (cell !== null && cell !== undefined) ? String(cell).trim() : "");

        // 2. Игнорируем полностью пустые строки
        if (row.every(cell => cell === "")) return;

        let companyValue = row[colIndex];

        // 3. Проверка на обязательность
        if (!companyValue) {
            rejectedData.push({ rowNum, reason: "Пустое название", rowData: row });
            return;
        }

        // 4. Очистка и принудительный CAPS LOCK
        const quoteRegex = /["'«»„“]/g;
        
        // Удаляем кавычки и переводим В ВЕРХНИЙ РЕГИСТР
        let cleanName = companyValue.replace(quoteRegex, '').toUpperCase();

        // Проверяем, изменилось ли что-то (для статистики)
        if (cleanName !== companyValue) {
            row[colIndex] = cleanName;
            modifiedCount++;
        }

        // Если после всех чисток строка стала пустой (например, были одни кавычки)
        if (!cleanName) {
            rejectedData.push({ rowNum, reason: "Некорректное название", rowData: row });
            return;
        }

        validRows.push(row);
    });

    displayResults(headers, rejectedData, validRows.length - 1, modifiedCount);

    const ws = XLSX.utils.aoa_to_sheet(validRows);
    csvContent = XLSX.utils.sheet_to_csv(ws);
    
    reportContainer.style.display = 'block';
    reportContainer.scrollIntoView({ behavior: 'smooth' });
    downloadBtn.style.display = validRows.length > 1 ? 'inline-flex' : 'none';
}

function displayResults(headers, rejected, success, modified) {
    const statsGrid = document.getElementById('stats-summary');
    statsGrid.innerHTML = `
        <div class="stat-card">
            <div class="stat-label">В итоговом CSV</div>
            <div class="stat-value" style="color: var(--success)">${success}</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">Приведено к CAPS/Без кавычек</div>
            <div class="stat-value" style="color: var(--warning)">${modified}</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">Ошибки (удалены)</div>
            <div class="stat-value" style="color: var(--danger)">${rejected.length}</div>
        </div>
    `;

    const details = document.getElementById('report-details');
    if (rejected.length > 0) {
        let html = `<h4>Детали исключенных строк:</h4><div class="table-wrapper"><table><thead><tr><th>Стр.</th><th>Причина</th>`;
        headers.forEach(h => html += `<th>${h}</th>`);
        html += `</tr></thead><tbody>`;
        rejected.forEach(item => {
            html += `<tr><td>${item.rowNum}</td><td><span class="reason-badge">${item.reason}</span></td>`;
            item.rowData.forEach(cell => html += `<td>${cell}</td>`);
            html += `</tr>`;
        });
        html += `</tbody></table></div>`;
        details.innerHTML = html;
    } else {
        details.innerHTML = `<p style="color: var(--success); text-align: center; padding: 20px;">✅ Проверка пройдена. Все названия теперь в ВЕРХНЕМ РЕГИСТРЕ.</p>`;
    }
}

downloadBtn.addEventListener('click', () => {
    const blob = new Blob(["\ufeff", csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `EXPORT_CAPS_${new Date().toISOString().slice(0,10)}.csv`;
    link.click();
});
