from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from managers import MANAGER_REFS, MANAGER_NAMES

# --- Константы стилей ---
FILL_HEADER_DARK  = PatternFill(start_color="2B4162", end_color="2B4162", fill_type="solid")
FILL_HEADER_LIGHT = PatternFill(start_color="3D5A80", end_color="3D5A80", fill_type="solid")
FILL_GREEN        = PatternFill(start_color="D7E8CA", end_color="D7E8CA", fill_type="solid")
FILL_BLUE         = PatternFill(start_color="D9E4F5", end_color="D9E4F5", fill_type="solid")
FILL_GRAY         = PatternFill(start_color="E8EDF5", end_color="E8EDF5", fill_type="solid")

FONT_BOLD_WHITE   = Font(bold=True, color="FFFFFF", name="Nunito", size=10)
FONT_NORMAL_BLACK = Font(name="Nunito", size=10)
ALIGN_CENTER      = Alignment(horizontal="center", vertical="bottom", wrap_text=True)
ALIGN_LEFT        = Alignment(horizontal="left", vertical="bottom")
# --- Конец стилей ---


def create_report(data: dict) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "CS Report"

    n = len(MANAGER_REFS)
    connections    = data.get("connections", {})
    activation_data = data.get("activation", {})
    five_k_data    = data.get("from_five_thousand", {})
    disconnections_data = data.get("disconnections", {})
    monthly_turnover_data = data.get("monthly_turnover", {})
    daily_turnover_data = data.get("daily_turnover", {})
    total_turnover_data = data.get("total_turnover", {})

    # === ЗАГОЛОВКИ: строка 1 ===
    ws.merge_cells("A1:A2")
    ws["A1"] = "Менеджер"
    ws["A1"].fill = FILL_HEADER_DARK
    ws["A1"].font = FONT_BOLD_WHITE
    ws["A1"].alignment = ALIGN_CENTER

    ws.merge_cells("B1:B2")
    ws["B1"] = "Реф"
    ws["B1"].fill = FILL_HEADER_DARK
    ws["B1"].font = FONT_BOLD_WHITE
    ws["B1"].alignment = ALIGN_CENTER

    ws.merge_cells("C1:D1")
    ws["C1"] = "Подключения и отток"
    ws["C1"].fill = FILL_HEADER_DARK
    ws["C1"].font = FONT_BOLD_WHITE
    ws["C1"].alignment = ALIGN_CENTER

    ws.merge_cells("F1:F2")
    ws["F1"] = "Мерчи от 5000"
    ws["F1"].fill = FILL_HEADER_DARK
    ws["F1"].font = FONT_BOLD_WHITE
    ws["F1"].alignment = ALIGN_CENTER

    ws.merge_cells("G1:G2")
    ws["G1"] = "Отключения"
    ws["G1"].fill = FILL_HEADER_DARK
    ws["G1"].font = FONT_BOLD_WHITE
    ws["G1"].alignment = ALIGN_CENTER

    ws.merge_cells("H1:H2")
    ws["H1"] = "Новый оборот"
    ws["H1"].fill = FILL_HEADER_DARK
    ws["H1"].font = FONT_BOLD_WHITE
    ws["H1"].alignment = ALIGN_CENTER

    ws.merge_cells("I1:I2")
    ws["I1"] = "Ср. дневной оборот"
    ws["I1"].fill = FILL_HEADER_DARK
    ws["I1"].font = FONT_BOLD_WHITE
    ws["I1"].alignment = ALIGN_CENTER

    ws.merge_cells("J1:J2")
    ws["J1"] = "Общий оборот"
    ws["J1"].fill = FILL_HEADER_DARK
    ws["J1"].font = FONT_BOLD_WHITE
    ws["J1"].alignment = ALIGN_CENTER

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
    for row_idx, (ref, name) in enumerate(zip(MANAGER_REFS, MANAGER_NAMES), start=3):
        cell_a = ws.cell(row=row_idx, column=1, value=name)
        cell_a.fill = FILL_GRAY
        cell_a.font = FONT_NORMAL_BLACK
        cell_a.alignment = ALIGN_LEFT

        cell_b = ws.cell(row=row_idx, column=2, value=ref)
        cell_b.fill = FILL_HEADER_DARK
        cell_b.font = FONT_BOLD_WHITE
        cell_b.alignment = ALIGN_CENTER

        cell_c = ws.cell(row=row_idx, column=3, value=connections.get(ref, 0))
        cell_c.fill = FILL_BLUE
        cell_c.font = FONT_NORMAL_BLACK
        cell_c.alignment = ALIGN_CENTER

        cell_d = ws.cell(row=row_idx, column=4, value=activation_data.get(ref, 0.0))
        cell_d.number_format = '0.00"%"'
        cell_d.fill = FILL_BLUE
        cell_d.font = FONT_NORMAL_BLACK
        cell_d.alignment = ALIGN_CENTER

        cell_f = ws.cell(row=row_idx, column=6, value=five_k_data.get(ref, 0))
        cell_f.fill = FILL_GREEN
        cell_f.font = FONT_NORMAL_BLACK
        cell_f.alignment = ALIGN_CENTER

        cell_g = ws.cell(row=row_idx, column=7, value=disconnections_data.get(ref, 0))
        cell_g.fill = FILL_BLUE
        cell_g.font = FONT_NORMAL_BLACK
        cell_g.alignment = ALIGN_CENTER

        cell_h = ws.cell(row=row_idx, column=8, value=monthly_turnover_data.get(ref, 0.0))
        cell_h.number_format = '#,##0.00'
        cell_h.fill = FILL_GREEN
        cell_h.font = FONT_NORMAL_BLACK
        cell_h.alignment = ALIGN_CENTER

        cell_i = ws.cell(row=row_idx, column=9, value=daily_turnover_data.get(ref, 0.0))
        cell_i.number_format = '#,##0.00'
        cell_i.fill = FILL_GREEN
        cell_i.font = FONT_NORMAL_BLACK
        cell_i.alignment = ALIGN_CENTER

        cell_j = ws.cell(row=row_idx, column=10, value=total_turnover_data.get(ref, 0.0))
        cell_j.number_format = '#,##0.00'
        cell_j.fill = FILL_GREEN
        cell_j.font = FONT_NORMAL_BLACK
        cell_j.alignment = ALIGN_CENTER

    # === ИТОГО: строка n+3 ===
    total_row = n + 3

    cell_total_c = ws.cell(row=total_row, column=3, value=sum(connections.get(ref, 0) for ref in MANAGER_REFS))
    cell_total_c.fill = FILL_HEADER_DARK
    cell_total_c.font = FONT_BOLD_WHITE
    cell_total_c.alignment = ALIGN_CENTER

    avg_act = sum(activation_data.get(ref, 0) for ref in MANAGER_REFS) / len(MANAGER_REFS)
    cell_total_d = ws.cell(row=total_row, column=4, value=avg_act)
    cell_total_d.number_format = '0.00"%"'
    cell_total_d.fill = FILL_HEADER_DARK
    cell_total_d.font = FONT_BOLD_WHITE
    cell_total_d.alignment = ALIGN_CENTER

    cell_total_f = ws.cell(row=total_row, column=6, value=sum(five_k_data.get(ref, 0) for ref in MANAGER_REFS))
    cell_total_f.fill = FILL_HEADER_DARK
    cell_total_f.font = FONT_BOLD_WHITE
    cell_total_f.alignment = ALIGN_CENTER

    cell_total_g = ws.cell(row=total_row, column=7, value=sum(disconnections_data.get(ref, 0) for ref in MANAGER_REFS))
    cell_total_g.fill = FILL_HEADER_DARK
    cell_total_g.font = FONT_BOLD_WHITE
    cell_total_g.alignment = ALIGN_CENTER

    cell_total_h = ws.cell(row=total_row, column=8,
    value=sum(monthly_turnover_data.get(ref, 0.0) for ref in MANAGER_REFS))
    cell_total_h.number_format = '#,##0.00'
    cell_total_h.fill = FILL_HEADER_DARK
    cell_total_h.font = FONT_BOLD_WHITE
    cell_total_h.alignment = ALIGN_CENTER

    avg_daily = sum(daily_turnover_data.get(ref, 0.0) for ref in MANAGER_REFS) / len(MANAGER_REFS)
    cell_total_i = ws.cell(row=total_row, column=9, value=avg_daily)
    cell_total_i.number_format = '#,##0.00'
    cell_total_i.fill = FILL_HEADER_DARK
    cell_total_i.font = FONT_BOLD_WHITE
    cell_total_i.alignment = ALIGN_CENTER

    cell_total_j = ws.cell(row=total_row, column=10,
    value=sum(total_turnover_data.get(ref, 0.0) for ref in MANAGER_REFS))
    cell_total_j.number_format = '#,##0.00'
    cell_total_j.fill = FILL_HEADER_DARK
    cell_total_j.font = FONT_BOLD_WHITE
    cell_total_j.alignment = ALIGN_CENTER

    # === ШИРИНА СТОЛБЦОВ ===
    ws.column_dimensions["A"].width = 28.0
    ws.column_dimensions["B"].width = 12.0
    ws.column_dimensions["C"].width = 14.0
    ws.column_dimensions["D"].width = 18.0
    ws.column_dimensions["E"].width = 4.0
    ws.column_dimensions["F"].width = 16.0
    ws.column_dimensions["G"].width = 14.0
    ws.column_dimensions["H"].width = 18.0
    ws.column_dimensions["I"].width = 20.0
    ws.column_dimensions["J"].width = 18.0

    return wb