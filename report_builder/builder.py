from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from managers import MANAGER_REFS, MANAGER_NAMES

# --- Константы стилей прямо здесь ---
FILL_HEADER_DARK = PatternFill(start_color="2B4162", end_color="2B4162", fill_type="solid")
FILL_HEADER_LIGHT = PatternFill(start_color="3D5A80", end_color="3D5A80", fill_type="solid")
FILL_GREEN = PatternFill(start_color="D7E8CA", end_color="D7E8CA", fill_type="solid")
FILL_BLUE = PatternFill(start_color="D9E4F5", end_color="D9E4F5", fill_type="solid")
FILL_GRAY = PatternFill(start_color="4A4A4A", end_color="4A4A4A", fill_type="solid")

FONT_BOLD_WHITE = Font(bold=True, color="FFFFFF", name="Nunito", size=10)
FONT_NORMAL_BLACK = Font(name="Nunito", size=10)
ALIGN_CENTER = Alignment(horizontal="center", vertical="bottom", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="bottom")


# --- Конец стилей ---


def create_report(data: dict) -> Workbook:
    """
    data — словарь со всеми данными от модулей:
    {
        "connections": {"Vitalijus": 0, "Kazakov": 6, ...},
        # future: "turnover": {...}, "historical": {...}, ...
    }

    builder НЕ считает — он только размещает готовые данные в ячейки.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "CS Report"

    n = len(MANAGER_REFS)

    # === ЗАГОЛОВОК: строка 1 ===
    # A1 — Менеджер
    ws.merge_cells("A1:A2")
    ws["A1"] = "Менеджер"
    ws["A1"].fill = FILL_HEADER_DARK
    ws["A1"].font = FONT_BOLD_WHITE
    ws["A1"].alignment = ALIGN_CENTER

    # B1 — Реф
    ws.merge_cells("B1:B2")
    ws["B1"] = "Реф"
    ws["B1"].fill = FILL_HEADER_DARK
    ws["B1"].font = FONT_BOLD_WHITE
    ws["B1"].alignment = ALIGN_CENTER

    # C1 — Подключения и отток
    ws.merge_cells("C1:D1")
    ws["C1"] = "Подключения и отток"
    ws["C1"].fill = FILL_HEADER_DARK
    ws["C1"].font = FONT_BOLD_WHITE
    ws["C1"].alignment = ALIGN_CENTER

    # === ПОДЗАГОЛОВКИ: строка 2 ===
    ws["C2"] = "Новые подкл."
    ws["C2"].fill = FILL_HEADER_LIGHT
    ws["C2"].font = FONT_BOLD_WHITE
    ws["C2"].alignment = ALIGN_CENTER

    ws["D2"] = "Активаций +%"
    ws["D2"].fill = FILL_HEADER_LIGHT
    ws["D2"].font = FONT_BOLD_WHITE
    ws["D2"].alignment = ALIGN_CENTER

    # === ДАННЫЕ: строки 3..n+2 ===
    connections = data.get("connections", {})

    for row_idx, (ref, name) in enumerate(zip(MANAGER_REFS, MANAGER_NAMES), start=3):
        # A — Имя
        cell_a = ws.cell(row=row_idx, column=1, value=name)
        cell_a.fill = FILL_GRAY
        cell_a.font = FONT_NORMAL_BLACK
        cell_a.alignment = ALIGN_LEFT

        # B — Реф
        cell_b = ws.cell(row=row_idx, column=2, value=ref)
        cell_b.fill = FILL_HEADER_DARK
        cell_b.font = FONT_BOLD_WHITE
        cell_b.alignment = ALIGN_CENTER

        # C — Новые подкл.
        cell_c = ws.cell(row=row_idx, column=3, value=connections.get(ref, 0))
        cell_c.fill = FILL_BLUE
        cell_c.font = FONT_NORMAL_BLACK
        cell_c.alignment = ALIGN_CENTER

        # D — Активаций +%
        activation_data = data.get("activation", {})
        act_pct = activation_data.get(ref, 0.0)
        cell_d = ws.cell(row=row_idx, column=4, value=f"{act_pct:.2f}%")
        cell_d.fill = FILL_BLUE
        cell_d.font = FONT_NORMAL_BLACK
        cell_d.alignment = ALIGN_CENTER

    # === ИТОГО: строка n+3 ===
    total_row = n + 3
    total_conn = sum(connections.get(ref, 0) for ref in MANAGER_REFS)
    ws.cell(row=total_row, column=3).value = total_conn
    ws.cell(row=total_row, column=3).fill = FILL_HEADER_DARK
    ws.cell(row=total_row, column=3).font = FONT_BOLD_WHITE

    # Итог для столбца D
    activation_data = data.get("activation", {})
    avg_activation = sum(activation_data.get(ref, 0) for ref in MANAGER_REFS) / len(MANAGER_REFS)
    ws.cell(row=total_row, column=4).value = f"{avg_activation:.2f}%"
    ws.cell(row=total_row, column=4).fill = FILL_HEADER_DARK
    ws.cell(row=total_row, column=4).font = FONT_BOLD_WHITE

    # === ШИРИНА СТОЛБЦОВ ===
    ws.column_dimensions["A"].width = 28.0
    ws.column_dimensions["B"].width = 12.0
    ws.column_dimensions["C"].width = 14.0
    ws.column_dimensions["D"].width = 18.0

    return wb