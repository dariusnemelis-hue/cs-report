import pandas as pd
from collections import defaultdict
from managers import MANAGER_REFS

IGNORE_REFS = {
    "maxat",
    "wramaccount1",
    "wramaccount3",
    "wramaccount5",
    "amarendra",
    "maintenance2",
    "maintenance",
}

def run(file_path: str) -> dict[str, int]:
    """
    Считает кол-во мерчантов с оборотом >= 5000 по каждому менеджеру.
    Использует тот же файл "Кастомная выгрузка mini.xlsx".
    Дата в D1 игнорируется — читаем начиная со строки 2 (header=0).
    Дата привязки (столбец B) тоже не важна — считаем всех мерчей без фильтра по месяцу.
    """
    df = pd.read_excel(file_path, sheet_name=0, header=0)

    result = defaultdict(int)

    for _, row in df.iterrows():
        try:
            ref = row.iloc[0]       # A — реф менеджера
            turnover = row.iloc[3]  # D — оборот

            if not isinstance(ref, str) or pd.isna(ref):
                continue
            ref = ref.strip().lower()
            if ref in IGNORE_REFS:
                continue

            if pd.isna(turnover):
                continue
            t = float(str(turnover).replace(",", "."))
            if t >= 5000:
                result[ref] += 1

        except (IndexError, ValueError, TypeError):
            continue

    return {ref: result.get(ref, 0) for ref in MANAGER_REFS}