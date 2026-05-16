import pandas as pd
from typing import Optional


def validate_excel_data(
    df: pd.DataFrame,
    required_columns: Optional[list[str]] = None,
) -> dict:
    """
    Валидирует DataFrame из Excel.

    Параметры
    ---------
    df : pd.DataFrame
        Данные для проверки.
    required_columns : list[str] | None
        Список обязательных колонок. Если None — проверяются все колонки.

    Возвращает
    ----------
    dict с ключами:
        is_valid : bool   — True, если ошибок нет
        errors   : list   — список словарей с описанием каждой ошибки
    """
    report: dict = {"is_valid": True, "errors": []}

    def add_error(error: dict) -> None:
        """Добавляет ошибку и сразу выставляет is_valid = False."""
        report["is_valid"] = False
        report["errors"].append(error)

    # 1. Проверка наличия обязательных колонок
    if required_columns:
        missing_cols = [c for c in required_columns if c not in df.columns]
        if missing_cols:
            add_error({
                "type": "missing_columns",
                "message": f"В файле отсутствуют обязательные колонки: {missing_cols}",
                "columns": missing_cols,
            })
            # Дальнейшие проверки бессмысленны без нужных колонок
            return report

    # 2. Проверка на полностью пустые строки
    empty_rows = df[df.isnull().all(axis=1)].index.tolist()
    if empty_rows:
        # +2: Excel-индекс начинается с 1, строка 1 — заголовок
        excel_rows = [i + 2 for i in empty_rows]
        add_error({
            "type": "empty_rows",
            "message": f"Найдены полностью пустые строки: {excel_rows}",
            "rows": excel_rows,
        })

    # 3. Проверка пропущенных значений в обязательных (или всех) колонках
    cols_to_check = required_columns if required_columns else df.columns.tolist()
    for col in cols_to_check:
        null_indices = df[df[col].isnull()].index.tolist()
        if null_indices:
            add_error({
                "type": "missing_values",
                "column": col,
                "message": f"Пустые значения в колонке «{col}»",
                "rows": [i + 2 for i in null_indices],
            })

    # 4. Проверка дублирующихся строк
    duplicate_mask = df.duplicated(keep=False)
    if duplicate_mask.any():
        dup_groups = (
            df[duplicate_mask]
            .groupby(list(df.columns), dropna=False)
            .apply(lambda g: [i + 2 for i in g.index.tolist()])
            .tolist()
        )
        add_error({
            "type": "duplicate_rows",
            "message": "Найдены полностью дублирующиеся строки",
            "groups": dup_groups,
        })

    return report
