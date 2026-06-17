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
    Считает кол-во отключившихся мерчантов по каждому менеджеру.
    Отключился = в прошлом месяце (столбец D) был оборот > 0,
                 а в расчётном месяце (столбец E) — 0 или пусто.
    Первая строка — заголовок (header=0), D1 и E1 это даты-названия, пропускаются автоматически.
    """
    df = pd.read_excel(file_path, sheet_name=0, header=0)

    result = defaultdict(int)

    for _, row in df.iterrows():
        try:
            ref = row.iloc[0]       # A — реф менеджера
            prev_turnover = row.iloc[3]   # D — оборот прошлого месяца
            curr_turnover = row.iloc[4]   # E — оборот расчётного месяца

            if not isinstance(ref, str) or pd.isna(ref):
                continue
            ref = ref.strip().lower()
            if ref in IGNORE_REFS:
                continue

            # Прошлый месяц: был оборот
            if pd.isna(prev_turnover):
                continue
            prev = float(str(prev_turnover).replace(",", "."))
            if prev <= 0:
                continue

            # Расчётный месяц: 0 или пусто = отключился
            if pd.isna(curr_turnover):
                result[ref] += 1
            else:
                curr = float(str(curr_turnover).replace(",", "."))
                if curr == 0:
                    result[ref] += 1

        except (IndexError, ValueError, TypeError):
            continue

    return {ref: result.get(ref, 0) for ref in MANAGER_REFS}