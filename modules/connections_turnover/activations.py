import re
from collections import defaultdict
import pandas as pd
from managers import MANAGER_REFS

IGNORE_REFS = {
    "maxat",
    "wramaccount1",
    "wramaccount3",
    "wramaccount5",
    "amarena",
    "maintenance2",
    "maintenance",
}


def get_month_from_file(filepath: str) -> str:
    """Месяц расчёта из ячейки D1."""
    df = pd.read_excel(filepath, sheet_name=0, header=None, nrows=1)
    val = df.iloc[0, 3]  # D1
    if pd.isna(val):
        return None
    if hasattr(val, "month"):
        return f"{val.year}-{str(val.month).zfill(2)}"
    s = str(val).strip()
    m = re.search(r"(\d{4})[-/](\d{1,2})", s)
    if m:
        return f"{m.group(1)}-{m.group(2).zfill(2)}"
    m = re.search(r"\d{2}\.(\d{2})\.\d{4}", s)
    if m:
        return f"2026-{m.group(1)}"
    return None


def load_activations(filepath: str) -> tuple[str, dict[str, dict]]:
    """
    Структура (заголовок в строке 1):
    A: ref менеджера
    B: дата привязки менеджера
    C: мерчант
    D: оборот за месяц
    """
    month = get_month_from_file(filepath)
    if not month:
        return None, {}

    df = pd.read_excel(filepath, sheet_name=0, header=0)
    result = defaultdict(lambda: {"total": 0, "activated": 0})

    for _, row in df.iterrows():
        try:
            ref = row.iloc[0]       # A
            date_val = row.iloc[1]  # B
            turnover = row.iloc[3]  # D

            if not isinstance(ref, str) or pd.isna(ref):
                continue
            ref = ref.strip().lower()
            if ref in IGNORE_REFS:
                continue

            # Месяц привязки
            bind_month = None
            if pd.notna(date_val):
                if hasattr(date_val, "month"):
                    bind_month = f"{date_val.year}-{str(date_val.month).zfill(2)}"
                elif isinstance(date_val, str):
                    m = re.search(r"(\d{4})[-/](\d{1,2})", date_val)
                    if m:
                        bind_month = f"{m.group(1)}-{m.group(2).zfill(2)}"

            if bind_month != month:
                continue

            result[ref]["total"] += 1

            if pd.notna(turnover):
                t = float(str(turnover).replace(",", "."))
                if t >= 100:
                    result[ref]["activated"] += 1
        except (IndexError, ValueError, TypeError):
            continue

    return month, dict(result)


def run(file_path: str) -> dict[str, float]:
    month, data = load_activations(file_path)
    if not data:
        print(f"[WARN] Нет данных по активациям (месяц: {month})")
        return {}
    print(f"[OK] Месяц: {month}, менеджеров: {len(data)}")

    total_conn = defaultdict(int)
    total_act = defaultdict(int)
    for ref, info in data.items():
        total_conn[ref] += info["total"]
        total_act[ref] += info["activated"]

    result = {}
    for ref in MANAGER_REFS:
        t = total_conn.get(ref, 0)
        result[ref] = round(total_act.get(ref, 0) / t * 100, 2) if t > 0 else 0.0
    return result