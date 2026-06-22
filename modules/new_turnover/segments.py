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

# Сегменты: (название, мин включительно, макс не включительно)
# None в максимуме = нет верхней границы
SEGMENTS = [
    ("starter",    0,      1000),
    ("growth",     1000,   2000),
    ("mid",        2000,   4000),
    ("large",      4000,   8000),
    ("vip",        8000,   16000),
    ("enterprise", 16000,  None),
]

def get_segment(turnover: float) -> str:
    """Определяет сегмент мерчанта по обороту. 0 и пустые идут в Starter."""
    for name, min_val, max_val in SEGMENTS:
        if max_val is None:
            if turnover >= min_val:
                return name
        else:
            if min_val <= turnover < max_val:
                return name
    return None

def run(file_path: str) -> dict[str, dict[str, int]]:
    """
    Считает количество мерчантов в каждом сегменте для каждого менеджера.
    Используется оборот за расчётный месяц (столбец D).
    Дата привязки не важна — считаем весь портфель менеджера.

    Возвращает:
    {
        "kazakov": {"starter": 3, "growth": 1, "mid": 0, ...},
        "lucas":   {"starter": 0, "growth": 2, ...},
        ...
    }
    """
    df = pd.read_excel(file_path, sheet_name=0, header=0)

    # Инициализируем счётчики для всех менеджеров и всех сегментов
    result = {
        ref: {seg[0]: 0 for seg in SEGMENTS}
        for ref in MANAGER_REFS
    }

    for _, row in df.iterrows():
        try:
            ref      = row.iloc[0]  # A — реф менеджера
            turnover = row.iloc[3]  # D — оборот за месяц

            if not isinstance(ref, str) or pd.isna(ref):
                continue
            ref = ref.strip().lower()
            if ref not in result:  # интересуют только наши менеджеры
                continue
            if ref in IGNORE_REFS:
                continue

            if pd.isna(turnover):
                continue
            t = float(str(turnover).replace(",", "."))

            # Если ячейка пустая — оборот 0, идёт в Starter
            if pd.isna(turnover):
                t = 0.0
            else:
                t = float(str(turnover).replace(",", "."))

            segment = get_segment(t)
            if segment:
                result[ref][segment] += 1

        except (IndexError, ValueError, TypeError):
            continue

    print(f"[OK] Сегменты: обработано {len(result)} менеджеров")
    return result