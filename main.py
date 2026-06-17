#!/usr/bin/env python3
"""
CS Report Generator — оркестратор
"""
import sys
from pathlib import Path

BASE_DIR = Path(__file__).parent.resolve()
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

sys.path.insert(0, str(BASE_DIR))

from modules.connections_turnover.connections import run as run_connections
from modules.connections_turnover.activations import run as run_activations
from modules.connections_turnover.from_five_thousand import run as run_from_five_thousand
from modules.connections_turnover.disconnections import run as run_disconnections
from modules.new_turnover.monthly_turnover import run as run_monthly_turnover
from modules.new_turnover.daily_turnover import run as run_daily_turnover
from report_builder.builder import create_report


def main():
    data = {}

    # === МОДУЛЬ 1: Подключения ===
    conn_file = INPUT_DIR / "Кол-во подключений.xlsx"
    if conn_file.exists():
        data["connections"] = run_connections(str(conn_file))
        print(f"[OK] Подключения: {len(data['connections'])} менеджеров")
    else:
        print(f"[WARN] Не найден: {conn_file.name}")
        data["connections"] = {}

    # === МОДУЛЬ 2: Активации ===
    act_file = INPUT_DIR / "Кастомная выгрузка mini.xlsx"
    if act_file.exists():
        data["activation"] = run_activations(str(act_file))
        print(f"[OK] Активации: {len(data['activation'])} менеджеров")
    else:
        print(f"[WARN] Не найден: {act_file.name}")
        data["activation"] = {}

    # === МОДУЛЬ 3: Мерчи от 5000 ===
    if act_file.exists():
        data["from_five_thousand"] = run_from_five_thousand(str(act_file))
        print(f"[OK] От 5000: {len(data['from_five_thousand'])} менеджеров")
    else:
        data["from_five_thousand"] = {}

    # === МОДУЛЬ 4: Отключения ===
    disc_file = INPUT_DIR / "Кастомная выгрузка 2 мес.xlsx"
    if disc_file.exists():
        data["disconnections"] = run_disconnections(str(disc_file))
        print(f"[OK] Отключения: {len(data['disconnections'])} менеджеров")
    else:
        print(f"[WARN] Не найден: {disc_file.name}")
        data["disconnections"] = {}

    # === МОДУЛЬ 5: Обороты (новые + общий) ===
    if act_file.exists():
        data["monthly_turnover"], data["total_turnover"] = run_monthly_turnover(str(act_file))
    else:
        print(f"[WARN] Не найден: {act_file.name}")
        data["monthly_turnover"] = {}
        data["total_turnover"] = {}

    # === МОДУЛЬ 6: Средний дневной оборот новых мерчей ===
    daily_file = INPUT_DIR / "Кастомная выгрузка по дням.xlsx"
    if daily_file.exists():
        data["daily_turnover"] = run_daily_turnover(str(daily_file))
    else:
        print(f"[WARN] Не найден: {daily_file.name}")
        data["daily_turnover"] = {}

    # === СБОРКА ===
    wb = create_report(data)
    output_file = OUTPUT_DIR / "cs_report.xlsx"
    wb.save(str(output_file))
    print(f"[OK] Отчёт сохранён: {output_file}")


if __name__ == "__main__":
    main()