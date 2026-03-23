document.getElementById('fileInput').addEventListener('change', handleFile);
const downloadBtn = document.getElementById('downloadBtn');
let csvContent = "";

function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
        
        validateAndConvert(jsonData);
    };
    reader.readAsArrayBuffer(file);
}

function validateAndConvert(data) {
    const headers = data[0];
    const rows = data.slice(1);
    const reportArea = document.getElementById('report');
    const campaignColName = "Название кампании";
    const campaignColIndex = headers.indexOf(campaignColName);

    if (campaignColIndex === -1) {
        reportArea.innerHTML = `<p class="error">❌ Ошибка: Колонка "${campaignColName}" не найдена!</p>`;
        return;
    }

    let validRows = [headers];
    let rejectedData = []; // Сюда сохраняем причину + саму строку

    rows.forEach((row, index) => {
        const rowNum = index + 2;
        const isCompletelyEmpty = row.every(cell => String(cell).trim() === "");
        
        if (isCompletelyEmpty) return; // Пустые строки просто игнорируем

        const campaignValue = String(row[campaignColIndex] || "").trim();
        let errorReason = "";

        // Проверка 1: Наличие названия
        if (!campaignValue) {
            errorReason = "Отсутствует название кампании";
        } 
        // Проверка 2: Кавычки
        else if (campaignValue.includes('"') || campaignValue.includes("'")) {
            errorReason = "Обнаружены кавычки";
        }

        if (errorReason) {
            rejectedData.push({ rowNum, reason: errorReason, rowData: row });
        } else {
            validRows.push(row);
        }
    });

    renderDetailedReport(headers, rejectedData, validRows.length - 1);

    const ws = XLSX.utils.aoa_to_sheet(validRows);
    csvContent = XLSX.utils.sheet_to_csv(ws);
    downloadBtn.style.display = validRows.length > 1 ? 'block' : 'none';
}

function renderDetailedReport(headers, rejectedData, successCount) {
    const reportArea = document.getElementById('report');
    let html = `<h3>Статистика:</h3>`;
    html += `<p>✅ Успешно: <strong>${successCount}</strong></p>`;

    if (rejectedData.length > 0) {
        html += `<p style="color: #d93025;">❌ Исключено строк: <strong>${rejectedData.length}</strong></p>`;
        html += `<div class="table-wrapper"><table><thead><tr><th>Стр.</th><th>Причина</th>`;
        
        // Добавляем заголовки из файла в таблицу отчета
        headers.forEach(h => html += `<th>${h}</th>`);
        html += `</tr></thead><tbody>`;

        rejectedData.forEach(item => {
            html += `<tr>
                <td>${item.rowNum}</td>
                <td class="reason-cell">${item.reason}</td>`;
            item.rowData.forEach(cell => {
                html += `<td>${cell}</td>`;
            });
            html += `</tr>`;
        });

        html += `</tbody></table></div>`;
    } else {
        html += `<p style="color: green;">Все строки прошли проверку!</p>`;
    }
    
    reportArea.innerHTML = html;
}
