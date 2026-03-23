// Элементы интерфейса
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const downloadBtn = document.getElementById('downloadBtn');
const reportContainer = document.getElementById('report-container');

let csvContent = ""; // Здесь будет храниться готовый CSV

// --- Настройка Drag-and-Drop ---
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

// --- Обработка файла ---
function handleFile(file) {
    if (!file.name.match(/\.(xlsx|xls)$/)) {
        alert("Пожалуйста, выберите файл Excel (.xlsx или .xls)");
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        
        // Читаем всё как массив массивов. defval: "" заменяет пустые ячейки на пустые строки
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
        validateAndConvert(jsonData);
    };
    reader.readAsArrayBuffer(file);
}

// --- Основная логика валидации и конвертации ---
function validateAndConvert(data) {
    if (data.length === 0) {
        alert("Файл пуст!");
        return;
    }

    const headers = data[0];
    const rows = data.slice(1);
    
    // 1. Умный поиск колонки "Название компании" (игнорируем регистр и пробелы в заголовке)
    const targetHeader = "название компании";
    const campaignColIndex = headers.findIndex(h => 
        String(h || "").trim().toLowerCase() === targetHeader
    );

    if (campaignColIndex === -1) {
        const actualHeaders = headers.map(h => `"${h}"`).join(', ');
        alert(`Ошибка: Колонка "Название компании" не найдена.\n\nЗаголовки в вашем файле: ${actualHeaders}`);
        return;
    }

    let validRows = [headers]; // Начинаем с заголовков
    let rejectedData = [];     // Для строк с ошибками
    let modifiedCount = 0;     // Сколько строк почистили от кавычек

    rows.forEach((originalRow, index) => {
        const rowNum = index + 2; // Номер строки для пользователя

        // 2. Глобальный Trim для всей строки (удаляем пробелы в начале/конце каждой ячейки)
        let row = originalRow.map(cell => (cell !== null && cell !== undefined) ? String(cell).trim() : "");

        // 3. Проверка: полностью пустая строка (после Trim)
        const isCompletelyEmpty = row.every(cell => cell === "");
        if (isCompletelyEmpty) return; // Просто пропускаем её

        let companyName = row[campaignColIndex];

        // 4. Проверка: обязательно наличие названия компании
        if (!companyName) {
            rejectedData.push({ 
                rowNum, 
                reason: "Пустое название компании", 
                rowData: row 
            });
            return;
        }

        // 5. Автоматическое удаление кавычек
        const quoteRegex = /["'«»„“]/g;
        if (quoteRegex.test(companyName)) {
            companyName = companyName.replace(quoteRegex, '');
            row[campaignColIndex] = companyName; // Обновляем в строке
            modifiedCount++;
        }

        // Если название компании после удаления кавычек стало пустым
        if (!companyName.trim()) {
            rejectedData.push({ 
                rowNum, 
                reason: "Название состояло только из кавычек", 
                rowData: row 
            });
            return;
        }

        // Если всё ок — добавляем в итоговый набор
        validRows.push(row);
    });

    // --- Вывод результатов ---
    displayResults(headers, rejectedData, validRows.length - 1, modifiedCount);

    // Подготовка CSV
    const ws = XLSX.utils.aoa_to_sheet(validRows);
    csvContent = XLSX.utils.sheet_to_csv(ws);
    
    // Показываем блок отчета
    reportContainer.style.display = 'block';
    reportContainer.scrollIntoView({ behavior: 'smooth' });
    
    // Кнопка скачивания доступна, если есть данные (кроме заголовка)
    downloadBtn.style.display = validRows.length > 1 ? 'inline-flex' : 'none';
}

// --- Отрисовка статистики и таблицы ошибок ---
function displayResults(headers, rejected, success, modified) {
    const statsGrid = document.getElementById('stats-summary');
    statsGrid.innerHTML = `
        <div class="stat-card">
            <div class="stat-label">Готово к экспорту</div>
            <div class="stat-value" style="color: var(--success)">${success}</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">Очищено от кавычек</div>
            <div class="stat-value" style="color: var(--warning)">${modified}</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">Ошибок (исключено)</div>
            <div class="stat-value" style="color: var(--danger)">${rejected.length}</div>
        </div>
    `;

    const details = document.getElementById('report-details');
    if (rejected.length > 0) {
        let html = `
            <h4>Строки, которые не вошли в CSV:</h4>
            <div class="table-wrapper">
                <table>
                    <thead>
                        <tr>
                            <th>Стр.</th>
                            <th>Причина</th>
                            ${headers.map(h => `<th>${h}</th>`).join('')}
                        </tr>
                    </thead>
                    <tbody>
        `;

        rejected.forEach(item => {
            html += `
                <tr>
                    <td>${item.rowNum}</td>
                    <td><span class="reason-badge">${item.reason}</span></td>
                    ${item.rowData.map(cell => `<td>${cell}</td>`).join('')}
                </tr>
            `;
        });

        html += `</tbody></table></div>`;
        details.innerHTML = html;
    } else {
        details.innerHTML = `<p style="color: var(--success); text-align: center; padding: 20px;">
            🎉 Все строки прошли валидацию и готовы к выгрузке!
        </p>`;
    }
}

// --- Скачивание файла ---
downloadBtn.addEventListener('click', () => {
    if (!csvContent) return;
    
    // Добавляем BOM для корректного открытия в Excel (чтобы не было кракозябр)
    const blob = new Blob(["\ufeff", csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    
    const link = document.createElement("a");
    link.href = url;
    link.download = `cleaned_data_${new Date().toLocaleDateString()}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
});
