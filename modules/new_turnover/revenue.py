import re
import pandas as pd
from collections import defaultdict
from managers import MANAGER_REFS

# Технические и служебные рефы — игнорируем
IGNORE_REFS = {
    "maxat",
    "wramaccount1",
    "wramaccount3",
    "wramaccount5",
    "amarendra",
    "maintenance2",
    "maintenance",
}


def get_month_from_file(filepath: str) -> str | None:
    """
    Определяет расчётный месяц из ячейки O1 файла.
    O1 — это 15-й столбец (индекс 14), первая строка.
    Возвращает строку формата "YYYY-MM", например "2026-04".
    """
    df = pd.read_excel(filepath, sheet_name=0, header=None, nrows=1)
    val = df.iloc[0, 14]  # O1

    if pd.isna(val):
        return None

    # Если pandas распознал как дату — берём год и месяц напрямую
    if hasattr(val, "month"):
        return f"{val.year}-{str(val.month).zfill(2)}"

    # Парсим строку — ищем паттерн YYYY-MM или YYYY/MM
    s = str(val).strip()
    m = re.search(r"(\d{4})[-/](\d{1,2})", s)
    if m:
        return f"{m.group(1)}-{m.group(2).zfill(2)}"

    # Запасной вариант: паттерн DD.MM.YYYY
    m = re.search(r"\d{2}\.(\d{2})\.(\d{4})", s)
    if m:
        return f"{m.group(2)}-{m.group(1)}"

    return None


def run(file_path: str) -> tuple[dict[str, float], dict[str, float]]:
    """
    Считает два показателя доходности для каждого менеджера.

    Доходность считается по формуле:
        доходность мерча = оборот (столбец O) * комиссия (столбец M)
    Комиссия хранится как дробное число (например 0.015 = 1.5%).

    1. new_revenue   — доходность только с новых мерчей,
                       то есть привязанных к менеджеру в расчётном месяце (столбец B).
                       Показывает сколько компания заработала на свежих подключениях.

    2. total_revenue — доходность со всего портфеля менеджера за месяц,
                       без фильтра по дате привязки.
                       Общий вклад менеджера в доходность компании.

    Возвращает кортеж (new_revenue, total_revenue).

    Структура входного файла "Кастомная выгрузка max.xlsx":
        Строка 1   — заголовки (пропускается через header=0)
        O1         — дата расчётного месяца
        Столбец A  — реф менеджера
        Столбец B  — дата привязки мерчанта к менеджеру
        Столбцы C-L — не используются
        Столбец M  — комиссия мерча (дробное число, например 0.015)
        Столбец N  — не используется
        Столбец O  — оборот мерча за расчётный месяц
    """
    month = get_month_from_file(file_path)
    if not month:
        print("[WARN] revenue: не удалось определить расчётный месяц")
        empty = {ref: 0.0 for ref in MANAGER_REFS}
        return empty, empty

    # Читаем данные — первая строка заголовок
    df = pd.read_excel(file_path, sheet_name=0, header=0)

    new_revenue   = defaultdict(float)  # доходность только новых мерчей
    total_revenue = defaultdict(float)  # доходность всего портфеля

    for _, row in df.iterrows():
        try:
            ref      = row.iloc[0]   # A — реф менеджера
            date_val = row.iloc[1]   # B — дата привязки мерчанта
            comm     = row.iloc[12]  # M — комиссия (индекс 12)
            turnover = row.iloc[14]  # O — оборот (индекс 14)

            # Пропускаем строки без рефа
            if not isinstance(ref, str) or pd.isna(ref):
                continue

            ref = ref.strip().lower()

            # Пропускаем служебные аккаунты
            if ref in IGNORE_REFS:
                continue

            # Пропускаем если нет оборота или комиссии
            if pd.isna(turnover) or pd.isna(comm):
                continue

            t = float(str(turnover).replace(",", "."))
            c = float(str(comm).replace(",", "."))

            # Доходность этого мерча = оборот * комиссия
            revenue = t * (c / 100)

            # Всегда в общую доходность — весь портфель
            total_revenue[ref] += revenue

            # Определяем месяц привязки мерчанта
            bind_month = None
            if pd.notna(date_val):
                if hasattr(date_val, "month"):
                    bind_month = f"{date_val.year}-{str(date_val.month).zfill(2)}"
                elif isinstance(date_val, str):
                    m = re.search(r"(\d{4})[-/](\d{1,2})", date_val)
                    if m:
                        bind_month = f"{m.group(1)}-{m.group(2).zfill(2)}"

            # В новую доходность — только мерчи привязанные в расчётном месяце
            if bind_month == month:
                new_revenue[ref] += revenue

        except (IndexError, ValueError, TypeError):
            continue

    print(f"[OK] Доходность: месяц {month}, менеджеров: {len([r for r in MANAGER_REFS if total_revenue.get(r, 0) > 0])}")

    return (
        {ref: round(new_revenue.get(ref, 0.0), 2)   for ref in MANAGER_REFS},
        {ref: round(total_revenue.get(ref, 0.0), 2) for ref in MANAGER_REFS},
    )