"""
create_database.py
------------------
Builds the example Excel workbook used by the rig planning app.
The workbook contains four sheets:
    1. Rigs              - master list of rigs and capabilities
    2. RigSchedule       - which rig is busy on which dates
    3. ReamerCapability  - which reamer sizes each rig can run
    4. Rods              - rod inventory by size and availability
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import date

wb = Workbook()

# ---- Styling helpers ----
header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
header_fill = PatternFill("solid", start_color="1F4E78")
center      = Alignment(horizontal="center", vertical="center")
thin        = Side(border_style="thin", color="BFBFBF")
border      = Border(left=thin, right=thin, top=thin, bottom=thin)


def style_header(ws, ncols):
    for col in range(1, ncols + 1):
        c = ws.cell(row=1, column=col)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center
        c.border = border
    ws.row_dimensions[1].height = 22


# =====================================================================
# Sheet 1: Rigs
# =====================================================================
ws1 = wb.active
ws1.title = "Rigs"
ws1.append(["RigName", "Location", "MaxHoleLength_m", "Status"])
rigs = [
    ["RIG-A", "Field North", 3500, "Active"],
    ["RIG-B", "Field North", 2500, "Active"],
    ["RIG-C", "Field South", 4000, "Active"],
    ["RIG-D", "Field South", 1800, "Active"],
    ["RIG-E", "Field East",  3000, "Active"],
]
for r in rigs:
    ws1.append(r)
style_header(ws1, 4)
for col, w in zip("ABCD", [14, 14, 18, 12]):
    ws1.column_dimensions[col].width = w


# =====================================================================
# Sheet 2: RigSchedule - dates each rig is BUSY (unavailable)
# =====================================================================
ws2 = wb.create_sheet("RigSchedule")
ws2.append(["RigName", "BusyDate", "Job"])
schedule = [
    ["RIG-A", date(2026, 6, 1),  "Job-100"],
    ["RIG-A", date(2026, 6, 2),  "Job-100"],
    ["RIG-A", date(2026, 6, 3),  "Job-100"],
    ["RIG-B", date(2026, 6, 5),  "Job-101"],
    ["RIG-B", date(2026, 6, 6),  "Job-101"],
    ["RIG-C", date(2026, 6, 1),  "Job-102"],
    ["RIG-C", date(2026, 6, 10), "Job-103"],
    ["RIG-D", date(2026, 6, 8),  "Job-104"],
    ["RIG-E", date(2026, 6, 2),  "Job-105"],
    ["RIG-E", date(2026, 6, 3),  "Job-105"],
]
for r in schedule:
    ws2.append(r)
style_header(ws2, 3)
for col, w in zip("ABC", [14, 14, 14]):
    ws2.column_dimensions[col].width = w


# =====================================================================
# Sheet 3: ReamerCapability - which reamer sizes each rig can handle
# Sizes are in inches: 8.5, 12.25, 17.5
# =====================================================================
ws3 = wb.create_sheet("ReamerCapability")
ws3.append(["RigName", "ReamerSize_in"])
capability = [
    ["RIG-A", 8.5],
    ["RIG-A", 12.25],
    ["RIG-A", 17.5],   # RIG-A can do all three
    ["RIG-B", 8.5],
    ["RIG-B", 12.25],
    ["RIG-C", 12.25],
    ["RIG-C", 17.5],
    ["RIG-D", 8.5],
    ["RIG-E", 8.5],
    ["RIG-E", 12.25],
    ["RIG-E", 17.5],
]
for r in capability:
    ws3.append(r)
style_header(ws3, 2)
for col, w in zip("AB", [14, 16]):
    ws3.column_dimensions[col].width = w


# =====================================================================
# Sheet 4: Rods - inventory by size (length per rod = 9 m typical)
# =====================================================================
ws4 = wb.create_sheet("Rods")
ws4.append(["RodSize_in", "LengthPerRod_m", "AvailableCount"])
rods = [
    [3.5, 9, 400],   # standard drill pipe
    [4.5, 9, 250],   # heavyweight
    [5.5, 9, 120],   # large diameter
]
for r in rods:
    ws4.append(r)
style_header(ws4, 3)
for col, w in zip("ABC", [14, 18, 18]):
    ws4.column_dimensions[col].width = w


# =====================================================================
# Sheet 5: ReamerRodMapping - which rod size pairs with which reamer
# =====================================================================
ws5 = wb.create_sheet("ReamerRodMapping")
ws5.append(["ReamerSize_in", "RequiredRodSize_in"])
mapping = [
    [8.5,   3.5],
    [12.25, 4.5],
    [17.5,  5.5],
]
for r in mapping:
    ws5.append(r)
style_header(ws5, 2)
for col, w in zip("AB", [16, 20]):
    ws5.column_dimensions[col].width = w


out = "rig_database.xlsx"
wb.save(out)
print(f"Saved: {out}")
