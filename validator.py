import pandas as pd

def validate_excel_data(df: pd.DataFrame):
    report = {
        "is_valid": True,
        "errors": []
    }
    
    # 1. Проверка на полностью пустые строки
    empty_rows = df[df.isnull().all(axis=1)].index.tolist()
    if empty_rows:
        report["errors"].append({
            "type": "empty_rows",
            "message": f"Найдены полностью пустые строки: {list(map(lambda x: x+2, empty_rows))}" 
            # +2 т.к. в Excel индекс начинается с 1 и есть заголовок
        })

    # 2. Проверка критических ячеек (пример: обязательная колонка 'ID')
    # Проверяем все колонки на наличие NaN
    for col in df.columns:
        null_indices = df[df[col].isnull()].index.tolist()
        if null_indices:
            report["is_valid"] = False
            report["errors"].append({
                "column": col,
                "type": "missing_values",
                "rows": [i + 2 for i in null_indices]
            })
            
    return report
