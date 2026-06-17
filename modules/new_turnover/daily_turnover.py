import re
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

def get_month_from_file(filepath: str) -> str:
    """Расчётный месяц из ячейки D1 (формат 2026-04-01 — берём только год-месяц)."""
    df = pd.read_excel(filepath, sheet_name=0, header=None, nrows=1)
    val = df.iloc[0, 3]
    if pd.isna(val):
        return None
    if hasattr(val, "month"):
        return f"{val.year}-{str(val.month).zfill(2)}"
    s = str(val).strip()
    m = re.search(r"(\d{4})[-/](\d{1,2})", s)
    if m:
        return f"{m.group(1)}-{m.group(2).zfill(2)}"
    return None

def run(file_path: str) -> dict[str, float]:
    """
    Считает средний дневной оборот новых мерчантов (подключённых в расчётном месяце).
    Столбцы D, E, F... — обороты по дням (D=день1, E=день2, ...).
    Средний = сумма всех дневных оборотов всех новых мерчей / кол-во дней в месяце.
    """
    month = get_month_from_file(file_path)
    if not month:
        print("[WARN] daily_turnover: не удалось определить расчётный месяц")
        return {ref: 0.0 for ref in MANAGER_REFS}

    # Читаем без заголовка чтобы точно знать индексы столбцов
    df_raw = pd.read_excel(file_path, sheet_name=0, header=None)

    # Строка 0 — заголовки, данные начинаются со строки 1
    # Столбцы: 0=ref, 1=дата привязки, 2=мерчант, 3..N=обороты по дням
    day_cols = list(range(3, df_raw.shape[1]))  # индексы столбцов с оборотами
    days_count = len(day_cols)

    df_data = df_raw.iloc[1:].reset_index(drop=True)  # убираем строку заголовков

    total_turnover = defaultdict(float)
    merchant_count = defaultdict(int)

    for _, row in df_data.iterrows():
        try:
            ref = row.iloc[0]       # A
            date_val = row.iloc[1]  # B — дата привязки

            if not isinstance(ref, str) or pd.isna(ref):
                continue
            ref = ref.strip().lower()
            if ref in IGNORE_REFS:
                continue

            # Только новые мерчанты — привязанные в расчётном месяце
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

            # Суммируем все дневные обороты этого мерчанта
            daily_sum = 0.0
            for col_idx in day_cols:
                val = row.iloc[col_idx]
                if pd.notna(val):
                    daily_sum += float(str(val).replace(",", "."))

            total_turnover[ref] += daily_sum
            merchant_count[ref] += 1

        except (IndexError, ValueError, TypeError):
            continue

    # Средний дневной оборот = общий оборот всех новых мерчей / кол-во дней
    result = {}
    for ref in MANAGER_REFS:
        total = total_turnover.get(ref, 0.0)
        result[ref] = round(total / days_count, 2) if days_count > 0 else 0.0

    print(f"[OK] Средний дневной оборот: месяц {month}, дней в файле: {days_count}")
    return result