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

        // Берем первый лист
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Превращаем в массив массивов (строки и колонки)
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

        validateAndConvert(jsonData);
    };
    reader.readAsArrayBuffer(file);
}

function validateAndConvert(data) {
    const reportArea = document.getElementById('report');
    let errors = [];
    let validRows = [];

    data.forEach((row, index) => {
        // Проверяем, если вся строка пустая
        const isEmpty = row.every(cell => cell.toString().trim() === "");
        
        if (isEmpty) {
            errors.push(`Строка ${index + 1}: Полностью пустая`);
        } else {
            // Проверяем на наличие пустых ячеек внутри строки
            const hasEmptyCell = row.some(cell => cell.toString().trim() === "");
            if (hasEmptyCell) {
                errors.push(`Строка ${index + 1}: Есть незаполненные поля`);
            }
            validRows.push(row);
        }
    });

    // Вывод отчета
    if (errors.length > 0) {
        reportArea.innerHTML = "<h3>Результаты проверки:</h3><ul>" + 
            errors.map(err => `<li>⚠️ ${err}</li>`).join('') + "</ul>";
    } else {
        reportArea.innerHTML = "<p style='color: green;'>✅ Все строки заполнены корректно!</p>";
    }

    // Создаем CSV
    const ws = XLSX.utils.aoa_to_sheet(validRows);
    csvContent = XLSX.utils.sheet_to_csv(ws);
    
    downloadBtn.style.display = 'block';
}

downloadBtn.addEventListener('click', () => {
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.setAttribute("download", "converted_data.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
});
