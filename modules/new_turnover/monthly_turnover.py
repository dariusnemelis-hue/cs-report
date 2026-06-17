import re
import pandas as pd
from collections import defaultdict
from managers import MANAGER_REFS

# Технические и служебные рефы, которые не являются менеджерами — игнорируем
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
    """
    Определяет расчётный месяц из ячейки D1 файла.
    D1 содержит дату в виде datetime или строки (например 2026-04-01).
    Возвращает строку формата "YYYY-MM", например "2026-04".
    """
    df = pd.read_excel(filepath, sheet_name=0, header=None, nrows=1)
    val = df.iloc[0, 3]  # D1

    if pd.isna(val):
        return None

    # Если pandas распознал как дату — берём год и месяц напрямую
    if hasattr(val, "month"):
        return f"{val.year}-{str(val.month).zfill(2)}"

    # Иначе парсим строку — ищем паттерн YYYY-MM или YYYY/MM
    s = str(val).strip()
    m = re.search(r"(\d{4})[-/](\d{1,2})", s)
    if m:
        return f"{m.group(1)}-{m.group(2).zfill(2)}"

    # Запасной вариант: паттерн DD.MM.YYYY
    m = re.search(r"\d{2}\.(\d{2})\.\d{4}", s)
    if m:
        return f"2026-{m.group(1)}"

    return None


def run(file_path: str) -> tuple[dict[str, float], dict[str, float]]:
    """
    За один проход по файлу считает два показателя для каждого менеджера:

    1. new_turnover   — оборот только новых мерчантов,
                        то есть тех, кто был привязан в расчётном месяце (столбец B).
                        Используется чтобы видеть сколько принесли именно свежие подключения.

    2. total_turnover — оборот ВСЕХ мерчантов менеджера за расчётный месяц,
                        без фильтра по дате привязки.
                        Общая картина портфеля.

    Оба словаря: { ref_менеджера: сумма_оборота }
    Возвращает кортеж (new_turnover, total_turnover).

    Структура входного файла "Кастомная выгрузка mini.xlsx":
        Строка 1  — заголовки (пропускается через header=0)
        D1        — дата расчётного месяца
        Столбец A — реф менеджера
        Столбец B — дата привязки мерчанта к менеджеру
        Столбец C — название мерчанта (не используем)
        Столбец D — оборот мерчанта за расчётный месяц
    """
    # Определяем расчётный месяц из шапки файла
    month = get_month_from_file(file_path)
    if not month:
        print("[WARN] monthly_turnover: не удалось определить расчётный месяц")
        empty = {ref: 0.0 for ref in MANAGER_REFS}
        return empty, empty

    # Читаем данные, первая строка — заголовок
    df = pd.read_excel(file_path, sheet_name=0, header=0)

    new_turnover   = defaultdict(float)  # оборот только новых мерчей
    total_turnover = defaultdict(float)  # оборот всех мерчей

    for _, row in df.iterrows():
        try:
            ref      = row.iloc[0]  # A — реф менеджера
            date_val = row.iloc[1]  # B — дата привязки мерчанта
            turnover = row.iloc[3]  # D — оборот за расчётный месяц

            # Пропускаем строки без рефа
            if not isinstance(ref, str) or pd.isna(ref):
                continue

            ref = ref.strip().lower()

            # Пропускаем служебные аккаунты
            if ref in IGNORE_REFS:
                continue

            # Пропускаем строки без оборота
            if pd.isna(turnover):
                continue

            t = float(str(turnover).replace(",", "."))

            # Добавляем в общий оборот — без условий, все мерчанты
            total_turnover[ref] += t

            # Определяем месяц привязки мерчанта
            bind_month = None
            if pd.notna(date_val):
                if hasattr(date_val, "month"):
                    # pandas распознал как дату
                    bind_month = f"{date_val.year}-{str(date_val.month).zfill(2)}"
                elif isinstance(date_val, str):
                    # парсим строку
                    m = re.search(r"(\d{4})[-/](\d{1,2})", date_val)
                    if m:
                        bind_month = f"{m.group(1)}-{m.group(2).zfill(2)}"

            # В новый оборот добавляем только если мерчант привязан в расчётном месяце
            if bind_month == month:
                new_turnover[ref] += t

        except (IndexError, ValueError, TypeError):
            continue

    print(f"[OK] Обороты: месяц {month}")

    # Округляем до 2 знаков и возвращаем только нужных менеджеров из MANAGER_REFS
    return (
        {ref: round(new_turnover.get(ref, 0.0), 2)   for ref in MANAGER_REFS},
        {ref: round(total_turnover.get(ref, 0.0), 2) for ref in MANAGER_REFS},
    )