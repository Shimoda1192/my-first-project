import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "QDC_FULL基本機能確認"

HOURS_PER_DAY = 8

COLOR_HEADER_DARK = "1F3864"
COLOR_HEADER_MID  = "2E75B6"
COLOR_SECTION_BG  = "D6E4F0"
COLOR_SUBSEC_BG   = "EBF3FB"
COLOR_ALT_ROW     = "F2F9FF"
COLOR_WHITE       = "FFFFFF"
COLOR_MANUAL_H    = "C6EFCE"
COLOR_MANUAL_D    = "E2EFDA"
COLOR_AUTO_H      = "FFEB9C"
COLOR_AUTO_D      = "FFF2CC"
COLOR_SUBTOTAL_BG = "B8CCE4"  # medium blue for subtotal rows
COLOR_TOTAL_BG    = "FFF2CC"

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def border(bottom_style="thin"):
    thin = Side(style="thin", color="AAAAAA")
    bot  = Side(style=bottom_style, color="333333" if bottom_style == "medium" else "AAAAAA")
    return Border(left=thin, right=thin, top=thin, bottom=bot)

def mfont(bold=False, color="000000", size=10):
    return Font(bold=bold, color=color, size=size, name="Meiryo")

def align(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

# Column widths
ws.column_dimensions["A"].width = 14
ws.column_dimensions["B"].width = 22
ws.column_dimensions["C"].width = 22
ws.column_dimensions["D"].width = 14
ws.column_dimensions["E"].width = 12
ws.column_dimensions["F"].width = 14
ws.column_dimensions["G"].width = 12

# ── Row 1: Title ────────────────────────────────────────────────
ws.merge_cells("A1:G1")
c = ws["A1"]
c.value = "QDC_FULL 基本機能確認（QFC, QDC）"
c.font = Font(bold=True, color=COLOR_WHITE, size=13, name="Meiryo")
c.fill = fill(COLOR_HEADER_DARK)
c.alignment = align()
c.border = border()
ws.row_dimensions[1].height = 26

# ── Row 2: Column headers ────────────────────────────────────────
for col, label in enumerate(
    ["大分類", "中分類", "詳細項目",
     "手動操作\n(h)", "手動操作\n(日)",
     "自動操作\n(h)", "自動操作\n(日)"], start=1
):
    c = ws.cell(row=2, column=col, value=label)
    c.font = Font(bold=True, color=COLOR_WHITE, size=10, name="Meiryo")
    c.fill = fill(COLOR_HEADER_MID)
    c.alignment = align(wrap=True)
    c.border = border()
ws.row_dimensions[2].height = 32

# ── Data definition ──────────────────────────────────────────────
# Groups: (group_label, cat1, cat2, [items])
# item: (detail, manual_h, auto_h)
groups = [
    (
        "① 実験準備時間 小計",
        "実験時間", "実験準備時間(h)",
        [
            ("データ整理\n（UPロード時間、データ確認）", 1.2, 1.2),
            ("スタンバイ\n（C/D載せ・降ろし）",           3.0, 3.0),
            ("計測準備\n（設置、設定）",                   9.5, 9.5),
            ("DTC取得",                                     1.2, 1.2),
        ],
    ),
    (
        "② 実験評価時間 小計",
        "実験時間", "実験評価時間(h)",
        [
            ("CD実験・作業",  8.0, 8.0),
            ("定置実験・作業", 6.0, 6.0),
        ],
    ),
    (
        "③ 波形作成・クライテリア確認時間 小計",
        "解析時間", "波形作成、クライテリア確認時間(h)",
        [
            ("", 16.0, 16.0),
        ],
    ),
]

# ── Build rows ───────────────────────────────────────────────────
# Row layout:
#   data rows for group 1 → subtotal row ①
#   data rows for group 2 → subtotal row ②
#   data rows for group 3 → subtotal row ③
#   TOTAL row

ROW_START = 3
current_row = ROW_START

# Track which rows belong to each section for merging A column
section_row_ranges = {}  # cat1 -> [first_row, last_row_before_total]
subcat_row_ranges  = {}  # (cat1, cat2) -> [first_row, last_data_row]

def write_value_cells(row, manual_h, auto_h, bold=False, bg_override=None):
    for col, val, fmt, bg in [
        (4, manual_h,                             "0.0",  bg_override or COLOR_MANUAL_H),
        (5, round(manual_h / HOURS_PER_DAY, 2),  "0.00", bg_override or COLOR_MANUAL_D),
        (6, auto_h,                               "0.0",  bg_override or COLOR_AUTO_H),
        (7, round(auto_h   / HOURS_PER_DAY, 2),  "0.00", bg_override or COLOR_AUTO_D),
    ]:
        c = ws.cell(row=row, column=col, value=val)
        c.font = mfont(bold=bold)
        c.fill = fill(bg)
        c.alignment = align()
        c.number_format = fmt
        c.border = border()

alt = 0
for group_label, cat1, cat2, items in groups:
    group_first_row = current_row

    # ── Data rows ─────────────────────────────────────────────
    for detail, manual_h, auto_h in items:
        ws.row_dimensions[current_row].height = 36
        row_bg = COLOR_ALT_ROW if alt % 2 == 0 else COLOR_WHITE
        alt += 1

        # A: 大分類 (value written only in first row; merge later)
        ca = ws.cell(row=current_row, column=1, value=None)
        ca.fill = fill(COLOR_SECTION_BG)
        ca.alignment = align(wrap=True)
        ca.border = border()

        # B: 中分類 (value written only in first row; merge later)
        cb = ws.cell(row=current_row, column=2, value=None)
        cb.fill = fill(COLOR_SUBSEC_BG)
        cb.alignment = align(wrap=True)
        cb.border = border()

        # C: 詳細項目
        cc = ws.cell(row=current_row, column=3, value=detail or None)
        cc.font = mfont()
        cc.fill = fill(row_bg)
        cc.alignment = align("left", wrap=True)
        cc.border = border()

        write_value_cells(current_row, manual_h, auto_h)
        current_row += 1

    data_last_row = current_row - 1

    # ── Subtotal row ───────────────────────────────────────────
    ws.row_dimensions[current_row].height = 24
    subtotal_h_manual = sum(it[1] for it in items)
    subtotal_h_auto   = sum(it[2] for it in items)

    # A: keep section color (will be part of A merge)
    ca = ws.cell(row=current_row, column=1, value=None)
    ca.fill = fill(COLOR_SECTION_BG)
    ca.border = border()

    # B+C merged: subtotal label
    ws.merge_cells(f"B{current_row}:C{current_row}")
    cb = ws.cell(row=current_row, column=2, value=group_label)
    cb.font = mfont(bold=True, color="1F3864")
    cb.fill = fill(COLOR_SUBTOTAL_BG)
    cb.alignment = align(wrap=True)
    cb.border = border()
    ws.cell(row=current_row, column=3).border = border()

    write_value_cells(current_row, subtotal_h_manual, subtotal_h_auto,
                      bold=True, bg_override=COLOR_SUBTOTAL_BG)

    section_row_ranges.setdefault(cat1, [current_row, current_row])
    section_row_ranges[cat1][0] = min(section_row_ranges[cat1][0], group_first_row)
    section_row_ranges[cat1][1] = current_row

    subcat_row_ranges[(cat1, cat2)] = [group_first_row, data_last_row]

    current_row += 1

# ── Write 大分類 / 中分類 values and merges ──────────────────────
for cat1, (r_first, r_last) in section_row_ranges.items():
    ws.cell(row=r_first, column=1).value = cat1
    ws.cell(row=r_first, column=1).font = mfont(bold=True)
    if r_first < r_last:
        ws.merge_cells(f"A{r_first}:A{r_last}")

for (cat1, cat2), (r_first, r_last) in subcat_row_ranges.items():
    ws.cell(row=r_first, column=2).value = cat2
    ws.cell(row=r_first, column=2).font = mfont(bold=True)
    if r_first < r_last:
        ws.merge_cells(f"B{r_first}:B{r_last}")
    # For 波形作成 row (no detail), merge B+C in the data row itself
    if groups[[g[0] for g in groups].index(
        next(g[0] for g in groups if g[1] == cat1 and g[2] == cat2)
    )][3][0][0] == "":  # detail is empty
        ws.merge_cells(f"B{r_first}:C{r_first}")

# ── TOTAL row ────────────────────────────────────────────────────
TOTAL_ROW = current_row
ws.row_dimensions[TOTAL_ROW].height = 28

total_manual_h = sum(it[1] for g in groups for it in g[3])
total_auto_h   = sum(it[2] for g in groups for it in g[3])

ws.merge_cells(f"A{TOTAL_ROW}:C{TOTAL_ROW}")
ct = ws[f"A{TOTAL_ROW}"]
ct.value = "TOTAL"
ct.font = Font(bold=True, size=11, name="Meiryo")
ct.fill = fill(COLOR_TOTAL_BG)
ct.alignment = align()
ct.border = border("medium")

for col, val, fmt in [
    (4, total_manual_h,                         "0.0"),
    (5, round(total_manual_h / HOURS_PER_DAY, 2), "0.00"),
    (6, total_auto_h,                           "0.0"),
    (7, round(total_auto_h   / HOURS_PER_DAY, 2), "0.00"),
]:
    c = ws.cell(row=TOTAL_ROW, column=col, value=val)
    c.font = Font(bold=True, size=11, name="Meiryo")
    c.fill = fill(COLOR_TOTAL_BG)
    c.alignment = align()
    c.number_format = fmt
    c.border = border("medium")

# ── Legend ───────────────────────────────────────────────────────
LEG_ROW = TOTAL_ROW + 2
ws.merge_cells(f"A{LEG_ROW}:G{LEG_ROW}")
ws[f"A{LEG_ROW}"].value = "凡例：1日 = 8時間換算"
ws[f"A{LEG_ROW}"].font = Font(italic=True, size=9, color="666666", name="Meiryo")
ws[f"A{LEG_ROW}"].alignment = align("left")

ws.sheet_view.showGridLines = False

wb.save("/home/user/my-first-project/QDC_FULL_schedule.xlsx")
print("Excel saved.")
