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

from modules.connections_turnover import run_connections, run_activations
from report_builder import create_report


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

    # === СБОРКА ===
    wb = create_report(data)
    output_file = OUTPUT_DIR / "cs_report.xlsx"
    wb.save(str(output_file))
    print(f"[OK] Отчёт сохранён: {output_file}")


if __name__ == "__main__":
    main()