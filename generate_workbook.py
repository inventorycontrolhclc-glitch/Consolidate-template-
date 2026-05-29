#!/usr/bin/env python3
"""
High Cultured Stock Consolidation Template Generator
Generates: High_Cultured_Stock_Consolidation_Template.xlsx
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import (
    Font, Fill, PatternFill, Alignment, Border, Side,
    GradientFill
)
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.worksheet.dimensions import ColumnDimension
import datetime

OUTPUT_FILE = "/home/user/Consolidate-template-/High_Cultured_Stock_Consolidation_Template.xlsx"

# ─── COLOR PALETTE ────────────────────────────────────────────────────────────
C_DARK_BG    = "1A1A2E"   # dark navy
C_ACCENT     = "E94560"   # red accent
C_MID        = "16213E"   # deep blue
C_LIGHT_BG   = "0F3460"   # medium blue header
C_WHITE      = "FFFFFF"
C_LIGHT_GREY = "F5F5F5"
C_GREY_LINE  = "D9D9D9"
C_GREEN      = "00B050"
C_ORANGE     = "FF8C00"
C_RED        = "FF0000"
C_YELLOW     = "FFFF00"
C_LIGHT_GREEN= "E2EFDA"
C_LIGHT_ORANGE="FFF2CC"
C_LIGHT_RED  = "FFCCCC"
C_HEADER_FONT= "FFFFFF"
C_SHEET_TAB_DARK = "1A1A2E"

# ─── STORES DATA ──────────────────────────────────────────────────────────────
STORES = [
    ("CS1",  "BERJAYA TIMES SQUARE (LOWER GROUND)",  "CS1-BERJAYA TIMES SQUARE (LOWER GROUND)"),
    ("CS3",  "SUNWAY PYRAMID",                        "CS3-SUNWAY PYRAMID"),
    ("CS4",  "MELAKA DATARAN PAHLAWAN",               "CS4-MELAKA DATARAN PAHLAWAN"),
    ("CS6",  "FAHRENHEIT 88",                         "CS6-FAHRENHEIT 88"),
    ("CS7",  "KOTA BHARU AEON MALL",                  "CS7-KOTA BHARU AEON MALL"),
    ("CS9",  "BERJAYA TIMES SQUARE (1ST FLOOR)",      "CS9-BERJAYA TIMES SQUARE (1ST FLOOR)"),
    ("CS10", "JOHOR PARADIGM MALL",                   "CS10-JOHOR PARADIGM MALL"),
    ("CS11", "KUANTAN EAST COAST MALL",               "CS11-KUANTAN EAST COAST MALL"),
    ("CS12", "PENANG QUEENSBAY MALL",                 "CS12-PENANG QUEENSBAY MALL"),
    ("CS13", "TERENGGANU KTCC MALL",                  "CS13-TERENGGANU KTCC MALL"),
    ("CS14", "ALOR SETAR AMAN CENTRAL",               "CS14-ALOR SETAR AMAN CENTRAL"),
    ("CS15", "JOHOR PARADIGM MALL (GROUND FLOOR)",    "CS15-JOHOR PARADIGM MALL (GROUND FLOOR)"),
    ("CS16", "SERI MANJUNG AEON MALL",                "CS16-SERI MANJUNG AEON MALL"),
    ("CS17", "TERENGGANU MAYANG MALL",                "CS17-TERENGGANU MAYANG MALL"),
    ("CS18", "KUCHING AEON MALL",                     "CS18-KUCHING AEON MALL"),
    ("CS19", "CENTRAL I-CITY MALL",                   "CS19-CENTRAL I-CITY MALL"),
    ("CS20", "ÆON MALL IPOH STATION 18",         "CS20-ÆON MALL IPOH STATION 18"),
    ("CS21", "ÆON MALL BANDARAY MELAKA",         "CS21-ÆON MALL BANDARAY MELAKA"),
    ("CS22", "1 UTAMA SHOPPING CENTRE",               "CS22-1 UTAMA SHOPPING CENTRE"),
]

# ─── SKU MASTER DATA ──────────────────────────────────────────────────────────
SKU_DATA = [
    ("HCLC-MENS-T-0001",   "Oversized Tee",   "Tee",   "S",         35.00, 89.00,  "2024-01-15", "Active"),
    ("HCLC-MENS-T-0001",   "Oversized Tee",   "Tee",   "M",         35.00, 89.00,  "2024-01-15", "Active"),
    ("HCLC-MENS-T-0001",   "Oversized Tee",   "Tee",   "L",         35.00, 89.00,  "2024-01-15", "Active"),
    ("HCLC-MENS-T-0001",   "Oversized Tee",   "Tee",   "XL",        35.00, 89.00,  "2024-01-15", "Active"),
    ("HCLC-MENS-PNT-0001", "Baggy Pants",     "Pants", "S",         55.00, 149.00, "2024-02-01", "Active"),
    ("HCLC-MENS-PNT-0001", "Baggy Pants",     "Pants", "M",         55.00, 149.00, "2024-02-01", "Active"),
    ("HCLC-MENS-PNT-0001", "Baggy Pants",     "Pants", "L",         55.00, 149.00, "2024-02-01", "Active"),
    ("HCLC-MENS-PNT-0001", "Baggy Pants",     "Pants", "XL",        55.00, 149.00, "2024-02-01", "Active"),
    ("HCLC-MENS-CAP-0001", "Logo Cap",        "Cap",   "FREE SIZE", 25.00, 69.00,  "2024-03-10", "Active"),
]

SIZES_ALL = ["S", "M", "L", "XL", "FREE SIZE", "XXL", "28", "30", "32", "34"]
SKUS_LIST = sorted(set(r[0] for r in SKU_DATA))
STORE_CODES = [s[0] for s in STORES]

# ─── STYLE HELPERS ────────────────────────────────────────────────────────────
def make_font(bold=False, size=11, color=C_DARK_BG, italic=False):
    return Font(name="Calibri", bold=bold, size=size, color=color, italic=italic)

def make_fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def make_border(style="thin"):
    s = Side(style=style)
    return Border(left=s, right=s, top=s, bottom=s)

def make_align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def header_style(ws, row, col, value, bg=C_LIGHT_BG, fg=C_WHITE, bold=True, size=11, wrap=False):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(name="Calibri", bold=bold, size=size, color=fg)
    cell.fill = make_fill(bg)
    cell.alignment = make_align("center", "center", wrap)
    cell.border = make_border()
    return cell

def data_cell(ws, row, col, value=None, bold=False, align="left", size=10, bg=None, fg=C_DARK_BG):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(name="Calibri", bold=bold, size=size, color=fg)
    if bg:
        cell.fill = make_fill(bg)
    cell.alignment = make_align(align, "center")
    cell.border = make_border("thin")
    return cell

def freeze_and_filter(ws, freeze_cell="A2", filter_ref=None):
    ws.freeze_panes = freeze_cell
    if filter_ref:
        ws.auto_filter.ref = filter_ref

def add_table(ws, ref, name, style="TableStyleMedium2"):
    tbl = Table(displayName=name, ref=ref)
    tbl.tableStyleInfo = TableStyleInfo(name=style, showFirstColumn=False,
                                        showLastColumn=False, showRowStripes=True,
                                        showColumnStripes=False)
    ws.add_table(tbl)

def col_widths(ws, widths: dict):
    for col_letter, w in widths.items():
        ws.column_dimensions[col_letter].width = w

def set_tab_color(ws, hex_color):
    ws.sheet_properties.tabColor = hex_color


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 1: README
# ══════════════════════════════════════════════════════════════════════════════
def build_readme(wb):
    ws = wb.create_sheet("README")
    set_tab_color(ws, "1A1A2E")
    ws.sheet_view.showGridLines = False

    # Title banner
    ws.merge_cells("A1:H1")
    c = ws["A1"]
    c.value = "HIGH CULTURED STOCK CONSOLIDATION SYSTEM"
    c.font = Font(name="Calibri", bold=True, size=20, color=C_WHITE)
    c.fill = make_fill(C_DARK_BG)
    c.alignment = make_align("center", "center")
    ws.row_dimensions[1].height = 45

    ws.merge_cells("A2:H2")
    c = ws["A2"]
    c.value = "User Guide & Instructions"
    c.font = Font(name="Calibri", bold=False, size=13, color="AAAAAA", italic=True)
    c.fill = make_fill(C_MID)
    c.alignment = make_align("center", "center")
    ws.row_dimensions[2].height = 28

    readme_sections = [
        ("", ""),
        ("HOW TO USE THIS WORKBOOK", ""),
        ("", ""),
        ("STEP 1 — Opening_Stock", "Paste or enter the current stock quantity for each SKU/Size/Store before any transfer or adjustment begins. This is your baseline."),
        ("", ""),
        ("STEP 2 — Transfer_Entry", "Enter all stock transfer movements between concept stores. Set the Status to 'Completed' when the physical stock has moved. Only Completed transfers affect the final balance."),
        ("", ""),
        ("STEP 3 — Stock_Adjustment", "Use this sheet to manually add or remove stock (e.g. new deliveries, damages, write-offs). Choose ADD or MINUS from the Adjustment_Type dropdown."),
        ("", ""),
        ("STEP 4 — Stock_Balance", "Review the final calculated stock balance per SKU/Size/Store. This sheet uses SUMIFS formulas — do NOT manually edit it."),
        ("", ""),
        ("STEP 5 — Consolidation_Planner", "Use this sheet to see which stores are overstocked or understocked, and get suggested transfer quantities."),
        ("", ""),
        ("STEP 6 — Dashboard", "Review the executive summary with total stock figures and charts."),
        ("", ""),
        ("", ""),
        ("WORKED EXAMPLES", ""),
        ("", ""),
        ("Example 1 — Store-to-Store Transfer", ""),
        ("Scenario:", "CS1 (Berjaya Times Square Lower Ground) wants to send 3 pcs of Oversized Tee (Size M) to CS3 (Sunway Pyramid)."),
        ("Action:", "Go to Transfer_Entry. Fill in Date, SKU = HCLC-MENS-T-0001, Size = M, From_Store_Code = CS1, To_Store_Code = CS3, Transfer_Qty = 3. Set Status = Completed."),
        ("Result:", "Stock_Balance will show CS1 minus 3 pcs and CS3 plus 3 pcs automatically."),
        ("", ""),
        ("Example 2 — Adding Stock from Warehouse", ""),
        ("Scenario:", "The warehouse delivers 5 pcs of Baggy Pants (Size L) to CS10 (Johor Paradigm Mall)."),
        ("Action:", "Go to Stock_Adjustment. Fill in Date, SKU = HCLC-MENS-PNT-0001, Size = L, Store_Code = CS10, Adjustment_Type = ADD, Adjustment_Qty = 5."),
        ("Result:", "Stock_Balance will show CS10 increased by 5 pcs for that SKU/Size."),
        ("", ""),
        ("Example 3 — Removing Damaged Stock", ""),
        ("Scenario:", "2 pcs of Logo Cap (FREE SIZE) at CS12 (Penang Queensbay Mall) are damaged and must be removed."),
        ("Action:", "Go to Stock_Adjustment. Fill in Date, SKU = HCLC-MENS-CAP-0001, Size = FREE SIZE, Store_Code = CS12, Adjustment_Type = MINUS, Adjustment_Qty = 2, Reason = Damaged."),
        ("Result:", "Stock_Balance will show CS12 reduced by 2 pcs for that SKU/Size."),
        ("", ""),
        ("", ""),
        ("IMPORTANT NOTES", ""),
        ("Do NOT edit formula cells", "Columns with grey/shaded backgrounds in Stock_Balance, Stock_Movement_Log, and Consolidation_Planner are formula-driven. Editing them will break the calculations."),
        ("Transaction-based logic", "This workbook uses a transaction ledger approach. Never overwrite Opening_Stock after operations have begun — instead use Stock_Adjustment ADD/MINUS."),
        ("Status matters", "Only Transfer_Entry rows with Status = Completed will be counted in Stock_Balance. Pending/Approved/Cancelled transfers do NOT affect stock."),
        ("Data Validation", "Dropdown lists are set on SKU, Size, Store_Code, and Status columns. Always select from the list to avoid formula errors."),
    ]

    row = 3
    for label, desc in readme_sections:
        ws.row_dimensions[row].height = 20
        if label in ("HOW TO USE THIS WORKBOOK", "WORKED EXAMPLES", "IMPORTANT NOTES"):
            ws.merge_cells(f"A{row}:H{row}")
            c = ws.cell(row=row, column=1, value=label)
            c.font = Font(name="Calibri", bold=True, size=13, color=C_WHITE)
            c.fill = make_fill(C_ACCENT)
            c.alignment = make_align("left", "center")
            ws.row_dimensions[row].height = 28
        elif label.startswith("STEP") or label.startswith("Example") and "—" in label:
            c1 = ws.cell(row=row, column=1, value=label)
            c1.font = Font(name="Calibri", bold=True, size=11, color=C_WHITE)
            c1.fill = make_fill(C_LIGHT_BG)
            c1.alignment = make_align("left", "center")
            ws.merge_cells(f"B{row}:H{row}")
            c2 = ws.cell(row=row, column=2, value=desc)
            c2.font = Font(name="Calibri", size=10, color=C_DARK_BG)
            c2.fill = make_fill("EBF3FB")
            c2.alignment = make_align("left", "center", wrap=True)
            ws.row_dimensions[row].height = 32
        elif label and label not in ("", "WORKED EXAMPLES"):
            c1 = ws.cell(row=row, column=1, value=label)
            c1.font = Font(name="Calibri", bold=True, size=10, color=C_LIGHT_BG)
            c1.fill = make_fill(C_LIGHT_GREY)
            c1.alignment = make_align("left", "center")
            ws.merge_cells(f"B{row}:H{row}")
            c2 = ws.cell(row=row, column=2, value=desc)
            c2.font = Font(name="Calibri", size=10, color="444444")
            c2.fill = make_fill(C_WHITE)
            c2.alignment = make_align("left", "center", wrap=True)
            ws.row_dimensions[row].height = 28
        row += 1

    # Column widths
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 18
    for col in ["C","D","E","F","G","H"]:
        ws.column_dimensions[col].width = 18

    ws.merge_cells("B3:H3")
    ws.freeze_panes = "A3"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 2: Store_List
# ══════════════════════════════════════════════════════════════════════════════
def build_store_list(wb):
    ws = wb.create_sheet("Store_List")
    set_tab_color(ws, "0F3460")

    headers = ["Store_Code", "Store_Name", "Full_Store_Name"]
    for i, h in enumerate(headers, 1):
        header_style(ws, 1, i, h, bg=C_LIGHT_BG, size=11)

    for r, (code, name, full) in enumerate(STORES, 2):
        data_cell(ws, r, 1, code, align="center")
        data_cell(ws, r, 2, name)
        data_cell(ws, r, 3, full)
        # alternate row shading
        if r % 2 == 0:
            for c in range(1, 4):
                ws.cell(row=r, column=c).fill = make_fill("EBF3FB")

    last_row = len(STORES) + 1
    add_table(ws, f"A1:C{last_row}", "Store_List_Table", "TableStyleMedium2")

    col_widths(ws, {"A": 14, "B": 45, "C": 52})
    freeze_and_filter(ws, "A2", f"A1:C{last_row}")


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 3: SKU_Master
# ══════════════════════════════════════════════════════════════════════════════
def build_sku_master(wb):
    ws = wb.create_sheet("SKU_Master")
    set_tab_color(ws, "0F3460")

    headers = ["SKU", "Item_Name", "Category", "Size", "Cost", "Retail_Price", "Launch_Date", "Status"]
    for i, h in enumerate(headers, 1):
        header_style(ws, 1, i, h, bg=C_LIGHT_BG)

    status_vals = ["Active", "Discontinued", "Clearance", "Incoming"]

    for r, row_data in enumerate(SKU_DATA, 2):
        sku, name, cat, size, cost, retail, launch, status = row_data
        data_cell(ws, r, 1, sku, align="center")
        data_cell(ws, r, 2, name)
        data_cell(ws, r, 3, cat, align="center")
        data_cell(ws, r, 4, size, align="center")
        c_cost = data_cell(ws, r, 5, cost, align="right")
        c_cost.number_format = '"RM "#,##0.00'
        c_retail = data_cell(ws, r, 6, retail, align="right")
        c_retail.number_format = '"RM "#,##0.00'
        data_cell(ws, r, 7, launch, align="center")
        data_cell(ws, r, 8, status, align="center")
        if r % 2 == 0:
            for c in range(1, 9):
                if ws.cell(row=r, column=c).fill.fgColor.rgb == "00000000":
                    ws.cell(row=r, column=c).fill = make_fill("F5F5F5")

    last_row = len(SKU_DATA) + 1

    # Status dropdown
    dv_status = DataValidation(type="list", formula1='"Active,Discontinued,Clearance,Incoming"',
                               showDropDown=False, showErrorMessage=True,
                               error="Select from list", errorTitle="Invalid Status")
    ws.add_data_validation(dv_status)
    dv_status.sqref = f"H2:H{last_row + 100}"

    add_table(ws, f"A1:H{last_row}", "SKU_Master_Table", "TableStyleMedium2")
    col_widths(ws, {"A": 24, "B": 22, "C": 12, "D": 12, "E": 12, "F": 14, "G": 14, "H": 14})
    freeze_and_filter(ws, "A2", f"A1:H{last_row}")


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 4: Opening_Stock
# ══════════════════════════════════════════════════════════════════════════════
def build_opening_stock(wb):
    ws = wb.create_sheet("Opening_Stock")
    set_tab_color(ws, "00B050")

    headers = ["SKU", "Item_Name", "Size", "Store_Code", "Store_Name", "Opening_Qty"]
    for i, h in enumerate(headers, 1):
        header_style(ws, 1, i, h, bg=C_LIGHT_BG)

    # Sample opening stock data
    sample_os = [
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "S",         "CS1",  5),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "M",         "CS1",  8),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "L",         "CS1",  6),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "XL",        "CS1",  3),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "S",         "CS3",  4),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "M",         "CS3",  7),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "L",         "CS3",  5),
        ("HCLC-MENS-PNT-0001", "Baggy Pants",     "M",         "CS1",  3),
        ("HCLC-MENS-PNT-0001", "Baggy Pants",     "L",         "CS1",  4),
        ("HCLC-MENS-PNT-0001", "Baggy Pants",     "M",         "CS3",  2),
        ("HCLC-MENS-PNT-0001", "Baggy Pants",     "L",         "CS3",  6),
        ("HCLC-MENS-CAP-0001", "Logo Cap",        "FREE SIZE", "CS1",  10),
        ("HCLC-MENS-CAP-0001", "Logo Cap",        "FREE SIZE", "CS3",  8),
        ("HCLC-MENS-CAP-0001", "Logo Cap",        "FREE SIZE", "CS10", 1),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "M",         "CS10", 2),
        ("HCLC-MENS-PNT-0001", "Baggy Pants",     "L",         "CS12", 0),
    ]

    sku_list_str = ",".join(SKUS_LIST)
    size_list_str = ",".join(SIZES_ALL)
    store_code_str = ",".join(STORE_CODES)

    # Data validations
    dv_sku = DataValidation(type="list",
                            formula1=f"SKU_Master_Table[SKU]",
                            showDropDown=False, showErrorMessage=True,
                            error="Select valid SKU", errorTitle="Invalid SKU")
    dv_size = DataValidation(type="list",
                             formula1='"S,M,L,XL,XXL,FREE SIZE,28,30,32,34"',
                             showDropDown=False, showErrorMessage=True,
                             error="Select valid Size", errorTitle="Invalid Size")
    dv_store = DataValidation(type="list",
                              formula1="Store_List_Table[Store_Code]",
                              showDropDown=False, showErrorMessage=True,
                              error="Select valid Store Code", errorTitle="Invalid Store")
    dv_qty = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0",
                            showErrorMessage=True,
                            error="Must be a whole number >= 0", errorTitle="Invalid Qty")

    ws.add_data_validation(dv_sku)
    ws.add_data_validation(dv_size)
    ws.add_data_validation(dv_store)
    ws.add_data_validation(dv_qty)

    for r, row_data in enumerate(sample_os, 2):
        sku, name, size, store_code, qty = row_data
        data_cell(ws, r, 1, sku, align="center")
        # Item_Name auto-fill formula
        fn = ws.cell(row=r, column=2)
        fn.value = f'=IFERROR(XLOOKUP(A{r},SKU_Master_Table[SKU],SKU_Master_Table[Item_Name],""),"")'
        fn.font = Font(name="Calibri", size=10, color="444444", italic=True)
        fn.fill = make_fill("F0F0F0")
        fn.alignment = make_align("left", "center")
        fn.border = make_border()
        data_cell(ws, r, 3, size, align="center")
        data_cell(ws, r, 4, store_code, align="center")
        # Store_Name XLOOKUP
        sn = ws.cell(row=r, column=5)
        sn.value = f'=IFERROR(XLOOKUP(D{r},Store_List_Table[Store_Code],Store_List_Table[Store_Name],""),"")'
        sn.font = Font(name="Calibri", size=10, color="444444", italic=True)
        sn.fill = make_fill("F0F0F0")
        sn.alignment = make_align("left", "center")
        sn.border = make_border()
        data_cell(ws, r, 6, qty, align="center")

    # Extend validations to 500 rows
    dv_sku.sqref = "A2:A500"
    dv_size.sqref = "C2:C500"
    dv_store.sqref = "D2:D500"
    dv_qty.sqref = "F2:F500"

    last_row = len(sample_os) + 1
    add_table(ws, f"A1:F{last_row}", "Opening_Stock_Table", "TableStyleMedium7")
    col_widths(ws, {"A": 24, "B": 22, "C": 12, "D": 14, "E": 40, "F": 14})
    freeze_and_filter(ws, "A2", f"A1:F{last_row}")


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 5: Transfer_Entry
# ══════════════════════════════════════════════════════════════════════════════
def build_transfer_entry(wb):
    ws = wb.create_sheet("Transfer_Entry")
    set_tab_color(ws, "E94560")

    headers = [
        "Date", "Transfer_ID", "SKU", "Item_Name", "Size",
        "From_Store_Code", "From_Store_Name", "To_Store_Code", "To_Store_Name",
        "Transfer_Qty", "Reason", "Requested_By", "Status", "Remarks", "Validation_Check"
    ]
    for i, h in enumerate(headers, 1):
        header_style(ws, 1, i, h, bg=C_ACCENT, size=10)

    # Sample transfers
    sample_transfers = [
        (datetime.date(2024, 4, 1), "TRF-0001", "HCLC-MENS-T-0001", "M", "CS1", "CS3", 2, "Balance stock", "Ahmad", "Completed", ""),
        (datetime.date(2024, 4, 2), "TRF-0002", "HCLC-MENS-CAP-0001", "FREE SIZE", "CS1", "CS10", 3, "Low stock at CS10", "Siti", "Completed", ""),
        (datetime.date(2024, 4, 3), "TRF-0003", "HCLC-MENS-PNT-0001", "L", "CS3", "CS12", 2, "Consolidation", "Ahmad", "Pending", "Awaiting approval"),
    ]

    # Data validations
    dv_sku = DataValidation(type="list", formula1="SKU_Master_Table[SKU]",
                            showDropDown=False, showErrorMessage=True,
                            error="Select valid SKU", errorTitle="Invalid SKU")
    dv_size = DataValidation(type="list", formula1='"S,M,L,XL,XXL,FREE SIZE,28,30,32,34"',
                             showDropDown=False, showErrorMessage=True,
                             error="Select valid Size", errorTitle="Invalid Size")
    dv_from = DataValidation(type="list", formula1="Store_List_Table[Store_Code]",
                             showDropDown=False, showErrorMessage=True,
                             error="Select valid Store", errorTitle="Invalid Store")
    dv_to = DataValidation(type="list", formula1="Store_List_Table[Store_Code]",
                           showDropDown=False, showErrorMessage=True,
                           error="Select valid Store", errorTitle="Invalid Store")
    dv_qty = DataValidation(type="whole", operator="greaterThan", formula1="0",
                            showErrorMessage=True,
                            error="Must be a whole number > 0", errorTitle="Invalid Qty")
    dv_status = DataValidation(type="list",
                               formula1='"Pending,Approved,Completed,Cancelled"',
                               showDropDown=False, showErrorMessage=True,
                               error="Select valid Status", errorTitle="Invalid Status")

    ws.add_data_validation(dv_sku)
    ws.add_data_validation(dv_size)
    ws.add_data_validation(dv_from)
    ws.add_data_validation(dv_to)
    ws.add_data_validation(dv_qty)
    ws.add_data_validation(dv_status)

    dv_sku.sqref    = "C2:C500"
    dv_size.sqref   = "E2:E500"
    dv_from.sqref   = "F2:F500"
    dv_to.sqref     = "H2:H500"
    dv_qty.sqref    = "J2:J500"
    dv_status.sqref = "M2:M500"

    for r, row_data in enumerate(sample_transfers, 2):
        date, tid, sku, size, from_code, to_code, qty, reason, reqby, status, remarks = row_data
        data_cell(ws, r, 1, date, align="center")
        ws.cell(row=r, column=1).number_format = "DD/MM/YYYY"
        data_cell(ws, r, 2, tid, align="center")
        data_cell(ws, r, 3, sku, align="center")

        # Item_Name auto-fill (col 4, D)
        fn = ws.cell(row=r, column=4)
        fn.value = f'=IFERROR(XLOOKUP(C{r},SKU_Master_Table[SKU],SKU_Master_Table[Item_Name],""),"")'
        fn.font = Font(name="Calibri", size=10, italic=True, color="444444")
        fn.fill = make_fill("F0F0F0")
        fn.alignment = make_align("left", "center")
        fn.border = make_border()

        data_cell(ws, r, 5, size, align="center")
        data_cell(ws, r, 6, from_code, align="center")

        # From_Store_Name auto-fill (col 7, G)
        fn2 = ws.cell(row=r, column=7)
        fn2.value = f'=IFERROR(XLOOKUP(F{r},Store_List_Table[Store_Code],Store_List_Table[Store_Name],""),"")'
        fn2.font = Font(name="Calibri", size=10, italic=True, color="444444")
        fn2.fill = make_fill("F0F0F0")
        fn2.alignment = make_align("left", "center")
        fn2.border = make_border()

        data_cell(ws, r, 8, to_code, align="center")

        # To_Store_Name auto-fill (col 9, I)
        fn3 = ws.cell(row=r, column=9)
        fn3.value = f'=IFERROR(XLOOKUP(H{r},Store_List_Table[Store_Code],Store_List_Table[Store_Name],""),"")'
        fn3.font = Font(name="Calibri", size=10, italic=True, color="444444")
        fn3.fill = make_fill("F0F0F0")
        fn3.alignment = make_align("left", "center")
        fn3.border = make_border()

        data_cell(ws, r, 10, qty, align="center")
        data_cell(ws, r, 11, reason)
        data_cell(ws, r, 12, reqby, align="center")
        data_cell(ws, r, 13, status, align="center")
        data_cell(ws, r, 14, remarks)

        # Validation_Check formula (col 15, O)
        vc = ws.cell(row=r, column=15)
        vc.value = (
            f'=IF(OR(C{r}="",E{r}="",F{r}="",H{r}="",J{r}=""),"ERROR: Missing Info",'
            f'IF(F{r}=H{r},"ERROR: Same Store",'
            f'IF(AND(J{r}>0,SUMIFS(Opening_Stock_Table[Opening_Qty],Opening_Stock_Table[SKU],C{r},Opening_Stock_Table[Size],E{r},Opening_Stock_Table[Store_Code],F{r})'
            f'+SUMIFS(Transfer_Entry_Table[Transfer_Qty],Transfer_Entry_Table[SKU],C{r},Transfer_Entry_Table[Size],E{r},Transfer_Entry_Table[To_Store_Code],F{r},Transfer_Entry_Table[Status],"Completed")'
            f'-SUMIFS(Transfer_Entry_Table[Transfer_Qty],Transfer_Entry_Table[SKU],C{r},Transfer_Entry_Table[Size],E{r},Transfer_Entry_Table[From_Store_Code],F{r},Transfer_Entry_Table[Status],"Completed")'
            f'+SUMIFS(Stock_Adjustment_Table[Adjustment_Qty],Stock_Adjustment_Table[SKU],C{r},Stock_Adjustment_Table[Size],E{r},Stock_Adjustment_Table[Store_Code],F{r},Stock_Adjustment_Table[Adjustment_Type],"ADD")'
            f'-SUMIFS(Stock_Adjustment_Table[Adjustment_Qty],Stock_Adjustment_Table[SKU],C{r},Stock_Adjustment_Table[Size],E{r},Stock_Adjustment_Table[Store_Code],F{r},Stock_Adjustment_Table[Adjustment_Type],"MINUS")'
            f'<J{r},"ERROR: Insufficient Stock","OK")))'
        )
        vc.font = Font(name="Calibri", size=10, bold=True, color=C_WHITE)
        vc.fill = make_fill("005B96")
        vc.alignment = make_align("center", "center")
        vc.border = make_border()

        # Status-based row coloring
        if status == "Completed":
            for c in [1,2,3,5,6,8,10,11,12,13,14]:
                ws.cell(row=r, column=c).fill = make_fill("E2EFDA")
        elif status == "Cancelled":
            for c in [1,2,3,5,6,8,10,11,12,13,14]:
                ws.cell(row=r, column=c).fill = make_fill("FFCCCC")

    last_row = len(sample_transfers) + 1
    add_table(ws, f"A1:O{last_row}", "Transfer_Entry_Table", "TableStyleMedium3")

    col_widths(ws, {
        "A": 13, "B": 14, "C": 24, "D": 20, "E": 10,
        "F": 16, "G": 36, "H": 14, "I": 36,
        "J": 14, "K": 22, "L": 16, "M": 12, "N": 20, "O": 22
    })
    freeze_and_filter(ws, "A2", f"A1:O{last_row}")

    # Conditional formatting for Validation_Check column O
    from openpyxl.formatting.rule import CellIsRule
    ok_fill = make_fill("00B050")
    err_fill = make_fill(C_RED)

    ws.conditional_formatting.add(
        f"O2:O500",
        CellIsRule(operator="equal", formula=['"OK"'],
                   fill=make_fill(C_GREEN), font=Font(color=C_WHITE, bold=True))
    )
    ws.conditional_formatting.add(
        f"O2:O500",
        CellIsRule(operator="notEqual", formula=['"OK"'],
                   fill=make_fill("FFCCCC"), font=Font(color=C_RED, bold=True))
    )


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 6: Stock_Adjustment
# ══════════════════════════════════════════════════════════════════════════════
def build_stock_adjustment(wb):
    ws = wb.create_sheet("Stock_Adjustment")
    set_tab_color(ws, "FF8C00")

    headers = [
        "Date", "Adjustment_ID", "SKU", "Item_Name", "Size",
        "Store_Code", "Store_Name", "Adjustment_Type", "Adjustment_Qty",
        "Reason", "PIC", "Remarks", "Validation_Check"
    ]
    for i, h in enumerate(headers, 1):
        header_style(ws, 1, i, h, bg=C_ORANGE, fg=C_WHITE)

    sample_adj = [
        (datetime.date(2024, 4, 5), "ADJ-0001", "HCLC-MENS-CAP-0001", "FREE SIZE", "CS12", "ADD",   5, "New delivery from warehouse", "Manager A", ""),
        (datetime.date(2024, 4, 6), "ADJ-0002", "HCLC-MENS-T-0001",   "L",         "CS3",  "MINUS", 1, "Damaged item",               "Manager B", "Item soiled"),
        (datetime.date(2024, 4, 7), "ADJ-0003", "HCLC-MENS-PNT-0001", "M",         "CS1",  "ADD",   4, "Warehouse replenishment",    "Manager A", ""),
    ]

    dv_sku   = DataValidation(type="list", formula1="SKU_Master_Table[SKU]",
                              showDropDown=False, showErrorMessage=True)
    dv_size  = DataValidation(type="list", formula1='"S,M,L,XL,XXL,FREE SIZE,28,30,32,34"',
                              showDropDown=False, showErrorMessage=True)
    dv_store = DataValidation(type="list", formula1="Store_List_Table[Store_Code]",
                              showDropDown=False, showErrorMessage=True)
    dv_type  = DataValidation(type="list", formula1='"ADD,MINUS"',
                              showDropDown=False, showErrorMessage=True,
                              error="Select ADD or MINUS", errorTitle="Invalid Type")
    dv_qty   = DataValidation(type="whole", operator="greaterThan", formula1="0",
                              showErrorMessage=True,
                              error="Must be whole number > 0", errorTitle="Invalid Qty")

    ws.add_data_validation(dv_sku)
    ws.add_data_validation(dv_size)
    ws.add_data_validation(dv_store)
    ws.add_data_validation(dv_type)
    ws.add_data_validation(dv_qty)

    dv_sku.sqref   = "C2:C500"
    dv_size.sqref  = "E2:E500"
    dv_store.sqref = "F2:F500"
    dv_type.sqref  = "H2:H500"
    dv_qty.sqref   = "I2:I500"

    for r, row_data in enumerate(sample_adj, 2):
        date, aid, sku, size, store_code, adj_type, qty, reason, pic, remarks = row_data
        data_cell(ws, r, 1, date, align="center")
        ws.cell(row=r, column=1).number_format = "DD/MM/YYYY"
        data_cell(ws, r, 2, aid, align="center")
        data_cell(ws, r, 3, sku, align="center")

        fn = ws.cell(row=r, column=4)
        fn.value = f'=IFERROR(XLOOKUP(C{r},SKU_Master_Table[SKU],SKU_Master_Table[Item_Name],""),"")'
        fn.font = Font(name="Calibri", size=10, italic=True, color="444444")
        fn.fill = make_fill("F0F0F0")
        fn.alignment = make_align("left", "center")
        fn.border = make_border()

        data_cell(ws, r, 5, size, align="center")
        data_cell(ws, r, 6, store_code, align="center")

        sn = ws.cell(row=r, column=7)
        sn.value = f'=IFERROR(XLOOKUP(F{r},Store_List_Table[Store_Code],Store_List_Table[Store_Name],""),"")'
        sn.font = Font(name="Calibri", size=10, italic=True, color="444444")
        sn.fill = make_fill("F0F0F0")
        sn.alignment = make_align("left", "center")
        sn.border = make_border()

        data_cell(ws, r, 8, adj_type, align="center")
        if adj_type == "ADD":
            ws.cell(row=r, column=8).fill = make_fill("E2EFDA")
            ws.cell(row=r, column=8).font = Font(name="Calibri", size=10, bold=True, color="00B050")
        else:
            ws.cell(row=r, column=8).fill = make_fill("FFCCCC")
            ws.cell(row=r, column=8).font = Font(name="Calibri", size=10, bold=True, color=C_RED)

        data_cell(ws, r, 9, qty, align="center")
        data_cell(ws, r, 10, reason)
        data_cell(ws, r, 11, pic, align="center")
        data_cell(ws, r, 12, remarks)

        # Validation_Check col 13
        vc = ws.cell(row=r, column=13)
        vc.value = (
            f'=IF(OR(C{r}="",E{r}="",F{r}="",H{r}="",I{r}=""),"ERROR: Missing Info",'
            f'IF(AND(H{r}="MINUS",'
            f'SUMIFS(Opening_Stock_Table[Opening_Qty],Opening_Stock_Table[SKU],C{r},Opening_Stock_Table[Size],E{r},Opening_Stock_Table[Store_Code],F{r})'
            f'+SUMIFS(Transfer_Entry_Table[Transfer_Qty],Transfer_Entry_Table[SKU],C{r},Transfer_Entry_Table[Size],E{r},Transfer_Entry_Table[To_Store_Code],F{r},Transfer_Entry_Table[Status],"Completed")'
            f'-SUMIFS(Transfer_Entry_Table[Transfer_Qty],Transfer_Entry_Table[SKU],C{r},Transfer_Entry_Table[Size],E{r},Transfer_Entry_Table[From_Store_Code],F{r},Transfer_Entry_Table[Status],"Completed")'
            f'+SUMIFS(Stock_Adjustment_Table[Adjustment_Qty],Stock_Adjustment_Table[SKU],C{r},Stock_Adjustment_Table[Size],E{r},Stock_Adjustment_Table[Store_Code],F{r},Stock_Adjustment_Table[Adjustment_Type],"ADD")'
            f'-SUMIFS(Stock_Adjustment_Table[Adjustment_Qty],Stock_Adjustment_Table[SKU],C{r},Stock_Adjustment_Table[Size],E{r},Stock_Adjustment_Table[Store_Code],F{r},Stock_Adjustment_Table[Adjustment_Type],"MINUS")'
            f'<I{r}),"ERROR: Insufficient Stock","OK"))'
        )
        vc.font = Font(name="Calibri", size=10, bold=True, color=C_WHITE)
        vc.fill = make_fill("005B96")
        vc.alignment = make_align("center", "center")
        vc.border = make_border()

    last_row = len(sample_adj) + 1
    add_table(ws, f"A1:M{last_row}", "Stock_Adjustment_Table", "TableStyleMedium4")

    col_widths(ws, {
        "A": 13, "B": 14, "C": 24, "D": 20, "E": 10,
        "F": 14, "G": 36, "H": 16, "I": 16,
        "J": 28, "K": 14, "L": 20, "M": 22
    })
    freeze_and_filter(ws, "A2", f"A1:M{last_row}")

    # Conditional formatting validation col
    ws.conditional_formatting.add(
        "M2:M500",
        CellIsRule(operator="equal", formula=['"OK"'],
                   fill=make_fill(C_GREEN), font=Font(color=C_WHITE, bold=True))
    )
    ws.conditional_formatting.add(
        "M2:M500",
        CellIsRule(operator="notEqual", formula=['"OK"'],
                   fill=make_fill("FFCCCC"), font=Font(color=C_RED, bold=True))
    )


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 7: Stock_Movement_Log
# ══════════════════════════════════════════════════════════════════════════════
def build_movement_log(wb):
    ws = wb.create_sheet("Stock_Movement_Log")
    set_tab_color(ws, "16213E")

    headers = [
        "Date", "Movement_Type", "Reference_ID", "SKU", "Item_Name",
        "Size", "Store_Code", "Store_Name", "Qty_In", "Qty_Out", "Net_Qty",
        "Reason", "Remarks"
    ]
    for i, h in enumerate(headers, 1):
        header_style(ws, 1, i, h, bg=C_MID)

    # Static sample movement rows (formula-driven union would require VBA or Power Query;
    # we build static examples and explain in header note)
    note = ws.cell(row=1, column=1)

    # Opening stock entries
    movements = []
    # Build from sample opening stock
    sample_os = [
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "S",         "CS1",  "BERJAYA TIMES SQUARE (LOWER GROUND)",  5),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "M",         "CS1",  "BERJAYA TIMES SQUARE (LOWER GROUND)",  8),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "L",         "CS1",  "BERJAYA TIMES SQUARE (LOWER GROUND)",  6),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "XL",        "CS1",  "BERJAYA TIMES SQUARE (LOWER GROUND)",  3),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "S",         "CS3",  "SUNWAY PYRAMID",                        4),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "M",         "CS3",  "SUNWAY PYRAMID",                        7),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "L",         "CS3",  "SUNWAY PYRAMID",                        5),
        ("HCLC-MENS-PNT-0001", "Baggy Pants",     "M",         "CS1",  "BERJAYA TIMES SQUARE (LOWER GROUND)",  3),
        ("HCLC-MENS-PNT-0001", "Baggy Pants",     "L",         "CS1",  "BERJAYA TIMES SQUARE (LOWER GROUND)",  4),
        ("HCLC-MENS-PNT-0001", "Baggy Pants",     "M",         "CS3",  "SUNWAY PYRAMID",                        2),
        ("HCLC-MENS-PNT-0001", "Baggy Pants",     "L",         "CS3",  "SUNWAY PYRAMID",                        6),
        ("HCLC-MENS-CAP-0001", "Logo Cap",        "FREE SIZE", "CS1",  "BERJAYA TIMES SQUARE (LOWER GROUND)", 10),
        ("HCLC-MENS-CAP-0001", "Logo Cap",        "FREE SIZE", "CS3",  "SUNWAY PYRAMID",                        8),
        ("HCLC-MENS-CAP-0001", "Logo Cap",        "FREE SIZE", "CS10", "JOHOR PARADIGM MALL",                   1),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "M",         "CS10", "JOHOR PARADIGM MALL",                   2),
        ("HCLC-MENS-PNT-0001", "Baggy Pants",     "L",         "CS12", "PENANG QUEENSBAY MALL",                 0),
    ]

    sample_transfers = [
        (datetime.date(2024, 4, 1), "TRF-0001", "HCLC-MENS-T-0001",   "Oversized Tee", "M",         "CS1", "BERJAYA TIMES SQUARE (LOWER GROUND)", "CS3",  "SUNWAY PYRAMID",          2, "Balance stock",         ""),
        (datetime.date(2024, 4, 2), "TRF-0002", "HCLC-MENS-CAP-0001", "Logo Cap",      "FREE SIZE", "CS1", "BERJAYA TIMES SQUARE (LOWER GROUND)", "CS10", "JOHOR PARADIGM MALL",     3, "Low stock at CS10",     ""),
    ]

    sample_adj = [
        (datetime.date(2024, 4, 5), "ADJ-0001", "HCLC-MENS-CAP-0001", "Logo Cap",    "FREE SIZE", "CS12", "PENANG QUEENSBAY MALL",                "ADD",   5, "New delivery",   ""),
        (datetime.date(2024, 4, 6), "ADJ-0002", "HCLC-MENS-T-0001",   "Oversized Tee","L",        "CS3",  "SUNWAY PYRAMID",                        "MINUS", 1, "Damaged item",   "Item soiled"),
        (datetime.date(2024, 4, 7), "ADJ-0003", "HCLC-MENS-PNT-0001", "Baggy Pants", "M",         "CS1",  "BERJAYA TIMES SQUARE (LOWER GROUND)",  "ADD",   4, "WH replenishment",""),
    ]

    row = 2
    mv_fill_map = {
        "OPENING":          "EBF3FB",
        "TRANSFER IN":      "E2EFDA",
        "TRANSFER OUT":     "FFCCCC",
        "ADJUSTMENT ADD":   "E2EFDA",
        "ADJUSTMENT MINUS": "FFCCCC",
    }

    # Opening rows
    for sku, name, size, store_code, store_name, qty in sample_os:
        if qty == 0:
            continue
        mv_type = "OPENING"
        cells_data = [
            (1, "01/01/2024"),
            (2, mv_type),
            (3, "OPENING"),
            (4, sku),
            (5, name),
            (6, size),
            (7, store_code),
            (8, store_name),
            (9, qty),       # Qty_In
            (10, 0),        # Qty_Out
            (11, qty),      # Net_Qty
            (12, "Opening Balance"),
            (13, ""),
        ]
        bg = mv_fill_map[mv_type]
        for col, val in cells_data:
            c = ws.cell(row=row, column=col, value=val)
            c.font = Font(name="Calibri", size=10, color=C_DARK_BG)
            c.fill = make_fill(bg)
            c.alignment = make_align("center" if col in [1,2,3,6,7,9,10,11] else "left", "center")
            c.border = make_border()
        row += 1

    # Transfer OUT rows (Completed only)
    for date, tid, sku, name, size, from_code, from_name, to_code, to_name, qty, reason, remarks in sample_transfers:
        # Transfer OUT from sending store
        mv_type = "TRANSFER OUT"
        bg = mv_fill_map[mv_type]
        for col, val in [
            (1, date.strftime("%d/%m/%Y")), (2, mv_type), (3, tid),
            (4, sku), (5, name), (6, size), (7, from_code), (8, from_name),
            (9, 0), (10, qty), (11, -qty), (12, reason), (13, remarks)
        ]:
            c = ws.cell(row=row, column=col, value=val)
            c.font = Font(name="Calibri", size=10, color=C_DARK_BG)
            c.fill = make_fill(bg)
            c.alignment = make_align("center" if col in [1,2,3,6,7,9,10,11] else "left", "center")
            c.border = make_border()
        row += 1

        # Transfer IN to receiving store
        mv_type = "TRANSFER IN"
        bg = mv_fill_map[mv_type]
        for col, val in [
            (1, date.strftime("%d/%m/%Y")), (2, mv_type), (3, tid),
            (4, sku), (5, name), (6, size), (7, to_code), (8, to_name),
            (9, qty), (10, 0), (11, qty), (12, reason), (13, remarks)
        ]:
            c = ws.cell(row=row, column=col, value=val)
            c.font = Font(name="Calibri", size=10, color=C_DARK_BG)
            c.fill = make_fill(bg)
            c.alignment = make_align("center" if col in [1,2,3,6,7,9,10,11] else "left", "center")
            c.border = make_border()
        row += 1

    # Adjustment rows
    for date, aid, sku, name, size, store_code, store_name, adj_type, qty, reason, remarks in sample_adj:
        mv_type = "ADJUSTMENT ADD" if adj_type == "ADD" else "ADJUSTMENT MINUS"
        bg = mv_fill_map[mv_type]
        qty_in  = qty if adj_type == "ADD" else 0
        qty_out = qty if adj_type == "MINUS" else 0
        net     = qty_in - qty_out
        for col, val in [
            (1, date.strftime("%d/%m/%Y")), (2, mv_type), (3, aid),
            (4, sku), (5, name), (6, size), (7, store_code), (8, store_name),
            (9, qty_in), (10, qty_out), (11, net), (12, reason), (13, remarks)
        ]:
            c = ws.cell(row=row, column=col, value=val)
            c.font = Font(name="Calibri", size=10, color=C_DARK_BG)
            c.fill = make_fill(bg)
            c.alignment = make_align("center" if col in [1,2,3,6,7,9,10,11] else "left", "center")
            c.border = make_border()
        row += 1

    # Note row
    note_cell = ws.cell(row=row + 1, column=1,
        value="NOTE: This log is pre-populated from sample data. In production, maintain this log manually or via macro as transactions are entered.")
    note_cell.font = Font(name="Calibri", size=9, italic=True, color="888888")
    ws.merge_cells(f"A{row+1}:M{row+1}")

    last_data_row = row - 1
    add_table(ws, f"A1:M{last_data_row}", "Stock_Movement_Log_Table", "TableStyleMedium2")

    col_widths(ws, {
        "A": 13, "B": 18, "C": 14, "D": 24, "E": 20,
        "F": 10, "G": 12, "H": 36, "I": 10, "J": 10,
        "K": 10, "L": 28, "M": 20
    })
    freeze_and_filter(ws, "A2", f"A1:M{last_data_row}")


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 8: Stock_Balance
# ══════════════════════════════════════════════════════════════════════════════
def build_stock_balance(wb):
    ws = wb.create_sheet("Stock_Balance")
    set_tab_color(ws, "00B050")

    headers = [
        "SKU", "Item_Name", "Size", "Store_Code", "Store_Name",
        "Opening_Qty", "Transfer_In_Qty", "Transfer_Out_Qty",
        "Adjustment_Add_Qty", "Adjustment_Minus_Qty", "Final_Qty"
    ]
    for i, h in enumerate(headers, 1):
        header_style(ws, 1, i, h, bg=C_GREEN, fg=C_WHITE)

    # Generate all combos from sample data
    combos = set()
    sample_os = [
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "S",         "CS1"),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "M",         "CS1"),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "L",         "CS1"),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "XL",        "CS1"),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "S",         "CS3"),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "M",         "CS3"),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "L",         "CS3"),
        ("HCLC-MENS-PNT-0001", "Baggy Pants",     "M",         "CS1"),
        ("HCLC-MENS-PNT-0001", "Baggy Pants",     "L",         "CS1"),
        ("HCLC-MENS-PNT-0001", "Baggy Pants",     "M",         "CS3"),
        ("HCLC-MENS-PNT-0001", "Baggy Pants",     "L",         "CS3"),
        ("HCLC-MENS-CAP-0001", "Logo Cap",        "FREE SIZE", "CS1"),
        ("HCLC-MENS-CAP-0001", "Logo Cap",        "FREE SIZE", "CS3"),
        ("HCLC-MENS-CAP-0001", "Logo Cap",        "FREE SIZE", "CS10"),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "M",         "CS10"),
        ("HCLC-MENS-PNT-0001", "Baggy Pants",     "L",         "CS12"),
        # From transfers - add CS3 transfer entries too
        ("HCLC-MENS-CAP-0001", "Logo Cap",        "FREE SIZE", "CS12"),
    ]

    store_name_map = {s[0]: s[1] for s in STORES}

    for combo in sample_os:
        combos.add(combo)

    combo_list = sorted(combos, key=lambda x: (x[0], x[2], x[3]))

    for r, (sku, name, size, store_code) in enumerate(combo_list, 2):
        store_name = store_name_map.get(store_code, "")

        data_cell(ws, r, 1, sku, align="center")
        data_cell(ws, r, 2, name)
        data_cell(ws, r, 3, size, align="center")
        data_cell(ws, r, 4, store_code, align="center")
        data_cell(ws, r, 5, store_name)

        # Formula cells - shaded grey (formula-driven)
        formula_bg = "F0F8FF"
        formula_font = Font(name="Calibri", size=10, italic=True, color="1A5276")

        # Opening_Qty - col 6 (F)
        f6 = ws.cell(row=r, column=6)
        f6.value = f'=IFERROR(SUMIFS(Opening_Stock_Table[Opening_Qty],Opening_Stock_Table[SKU],A{r},Opening_Stock_Table[Size],C{r},Opening_Stock_Table[Store_Code],D{r}),0)'
        f6.font = formula_font; f6.fill = make_fill(formula_bg)
        f6.alignment = make_align("center","center"); f6.border = make_border()
        f6.number_format = "0"

        # Transfer_In_Qty - col 7 (G)
        f7 = ws.cell(row=r, column=7)
        f7.value = f'=IFERROR(SUMIFS(Transfer_Entry_Table[Transfer_Qty],Transfer_Entry_Table[SKU],A{r},Transfer_Entry_Table[Size],C{r},Transfer_Entry_Table[To_Store_Code],D{r},Transfer_Entry_Table[Status],"Completed"),0)'
        f7.font = formula_font; f7.fill = make_fill(formula_bg)
        f7.alignment = make_align("center","center"); f7.border = make_border()
        f7.number_format = "0"

        # Transfer_Out_Qty - col 8 (H)
        f8 = ws.cell(row=r, column=8)
        f8.value = f'=IFERROR(SUMIFS(Transfer_Entry_Table[Transfer_Qty],Transfer_Entry_Table[SKU],A{r},Transfer_Entry_Table[Size],C{r},Transfer_Entry_Table[From_Store_Code],D{r},Transfer_Entry_Table[Status],"Completed"),0)'
        f8.font = formula_font; f8.fill = make_fill(formula_bg)
        f8.alignment = make_align("center","center"); f8.border = make_border()
        f8.number_format = "0"

        # Adjustment_Add_Qty - col 9 (I)
        f9 = ws.cell(row=r, column=9)
        f9.value = f'=IFERROR(SUMIFS(Stock_Adjustment_Table[Adjustment_Qty],Stock_Adjustment_Table[SKU],A{r},Stock_Adjustment_Table[Size],C{r},Stock_Adjustment_Table[Store_Code],D{r},Stock_Adjustment_Table[Adjustment_Type],"ADD"),0)'
        f9.font = formula_font; f9.fill = make_fill(formula_bg)
        f9.alignment = make_align("center","center"); f9.border = make_border()
        f9.number_format = "0"

        # Adjustment_Minus_Qty - col 10 (J)
        f10 = ws.cell(row=r, column=10)
        f10.value = f'=IFERROR(SUMIFS(Stock_Adjustment_Table[Adjustment_Qty],Stock_Adjustment_Table[SKU],A{r},Stock_Adjustment_Table[Size],C{r},Stock_Adjustment_Table[Store_Code],D{r},Stock_Adjustment_Table[Adjustment_Type],"MINUS"),0)'
        f10.font = formula_font; f10.fill = make_fill(formula_bg)
        f10.alignment = make_align("center","center"); f10.border = make_border()
        f10.number_format = "0"

        # Final_Qty - col 11 (K)
        f11 = ws.cell(row=r, column=11)
        f11.value = f'=F{r}+G{r}-H{r}+I{r}-J{r}'
        f11.font = Font(name="Calibri", size=11, bold=True, color=C_DARK_BG)
        f11.fill = make_fill("DDEEFF")
        f11.alignment = make_align("center","center"); f11.border = make_border()
        f11.number_format = "0"

    last_row = len(combo_list) + 1
    add_table(ws, f"A1:K{last_row}", "Stock_Balance_Table", "TableStyleMedium7")

    # Conditional formatting on Final_Qty (col K)
    red_fill   = make_fill(C_LIGHT_RED)
    orange_fill = make_fill(C_LIGHT_ORANGE)
    green_fill = make_fill(C_LIGHT_GREEN)

    ws.conditional_formatting.add(
        f"K2:K{last_row}",
        CellIsRule(operator="equal", formula=["0"],
                   fill=red_fill, font=Font(bold=True, color=C_RED))
    )
    ws.conditional_formatting.add(
        f"K2:K{last_row}",
        CellIsRule(operator="between", formula=["1","2"],
                   fill=orange_fill, font=Font(bold=True, color=C_ORANGE))
    )
    ws.conditional_formatting.add(
        f"K2:K{last_row}",
        CellIsRule(operator="greaterThanOrEqual", formula=["10"],
                   fill=green_fill, font=Font(bold=True, color=C_GREEN))
    )

    col_widths(ws, {
        "A": 24, "B": 20, "C": 12, "D": 14, "E": 36,
        "F": 14, "G": 16, "H": 17, "I": 18, "J": 20, "K": 12
    })
    freeze_and_filter(ws, "A2", f"A1:K{last_row}")


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 9: Consolidation_Planner
# ══════════════════════════════════════════════════════════════════════════════
def build_consolidation_planner(wb):
    ws = wb.create_sheet("Consolidation_Planner")
    set_tab_color(ws, "9B59B6")

    headers = [
        "SKU", "Item_Name", "Size", "Store_Code", "Store_Name",
        "Final_Qty", "Min_Stock_Level", "Max_Stock_Level",
        "Stock_Status", "Suggested_Action", "Suggested_Transfer_Qty"
    ]
    for i, h in enumerate(headers, 1):
        header_style(ws, 1, i, h, bg="6C3483", fg=C_WHITE)

    # Same combos as stock balance
    combo_list = [
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "S",         "CS1"),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "M",         "CS1"),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "L",         "CS1"),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "XL",        "CS1"),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "S",         "CS3"),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "M",         "CS3"),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "L",         "CS3"),
        ("HCLC-MENS-PNT-0001", "Baggy Pants",     "M",         "CS1"),
        ("HCLC-MENS-PNT-0001", "Baggy Pants",     "L",         "CS1"),
        ("HCLC-MENS-PNT-0001", "Baggy Pants",     "M",         "CS3"),
        ("HCLC-MENS-PNT-0001", "Baggy Pants",     "L",         "CS3"),
        ("HCLC-MENS-CAP-0001", "Logo Cap",        "FREE SIZE", "CS1"),
        ("HCLC-MENS-CAP-0001", "Logo Cap",        "FREE SIZE", "CS3"),
        ("HCLC-MENS-CAP-0001", "Logo Cap",        "FREE SIZE", "CS10"),
        ("HCLC-MENS-T-0001",   "Oversized Tee",   "M",         "CS10"),
        ("HCLC-MENS-PNT-0001", "Baggy Pants",     "L",         "CS12"),
        ("HCLC-MENS-CAP-0001", "Logo Cap",        "FREE SIZE", "CS12"),
    ]

    store_name_map = {s[0]: s[1] for s in STORES}
    formula_bg = "F5EEF8"
    formula_font = Font(name="Calibri", size=10, italic=True, color="6C3483")

    for r, (sku, name, size, store_code) in enumerate(combo_list, 2):
        store_name = store_name_map.get(store_code, "")
        data_cell(ws, r, 1, sku, align="center")
        data_cell(ws, r, 2, name)
        data_cell(ws, r, 3, size, align="center")
        data_cell(ws, r, 4, store_code, align="center")
        data_cell(ws, r, 5, store_name)

        # Final_Qty - pull from Stock_Balance (formula)
        f6 = ws.cell(row=r, column=6)
        f6.value = (
            f'=IFERROR(SUMIFS(Opening_Stock_Table[Opening_Qty],Opening_Stock_Table[SKU],A{r},Opening_Stock_Table[Size],C{r},Opening_Stock_Table[Store_Code],D{r})'
            f'+SUMIFS(Transfer_Entry_Table[Transfer_Qty],Transfer_Entry_Table[SKU],A{r},Transfer_Entry_Table[Size],C{r},Transfer_Entry_Table[To_Store_Code],D{r},Transfer_Entry_Table[Status],"Completed")'
            f'-SUMIFS(Transfer_Entry_Table[Transfer_Qty],Transfer_Entry_Table[SKU],A{r},Transfer_Entry_Table[Size],C{r},Transfer_Entry_Table[From_Store_Code],D{r},Transfer_Entry_Table[Status],"Completed")'
            f'+SUMIFS(Stock_Adjustment_Table[Adjustment_Qty],Stock_Adjustment_Table[SKU],A{r},Stock_Adjustment_Table[Size],C{r},Stock_Adjustment_Table[Store_Code],D{r},Stock_Adjustment_Table[Adjustment_Type],"ADD")'
            f'-SUMIFS(Stock_Adjustment_Table[Adjustment_Qty],Stock_Adjustment_Table[SKU],A{r},Stock_Adjustment_Table[Size],C{r},Stock_Adjustment_Table[Store_Code],D{r},Stock_Adjustment_Table[Adjustment_Type],"MINUS"),0)'
        )
        f6.font = formula_font; f6.fill = make_fill(formula_bg)
        f6.alignment = make_align("center","center"); f6.border = make_border()
        f6.number_format = "0"

        # Min_Stock_Level - editable (default 3)
        data_cell(ws, r, 7, 3, align="center")

        # Max_Stock_Level - editable (default 8)
        data_cell(ws, r, 8, 8, align="center")

        # Stock_Status formula
        f9 = ws.cell(row=r, column=9)
        f9.value = f'=IF(F{r}>H{r},"OVERSTOCK",IF(F{r}<G{r},"LOW STOCK","BALANCED"))'
        f9.font = Font(name="Calibri", size=10, bold=True)
        f9.fill = make_fill(formula_bg)
        f9.alignment = make_align("center","center"); f9.border = make_border()

        # Suggested_Action formula
        f10 = ws.cell(row=r, column=10)
        f10.value = f'=IF(I{r}="OVERSTOCK","Send Out",IF(I{r}="LOW STOCK","Request Stock","No Action"))'
        f10.font = Font(name="Calibri", size=10, bold=True)
        f10.fill = make_fill(formula_bg)
        f10.alignment = make_align("center","center"); f10.border = make_border()

        # Suggested_Transfer_Qty formula
        f11 = ws.cell(row=r, column=11)
        f11.value = f'=IF(I{r}="OVERSTOCK",F{r}-H{r},IF(I{r}="LOW STOCK",G{r}-F{r},0))'
        f11.font = Font(name="Calibri", size=10, bold=True)
        f11.fill = make_fill(formula_bg)
        f11.alignment = make_align("center","center"); f11.border = make_border()
        f11.number_format = "0"

    last_row = len(combo_list) + 1

    # Conditional formatting Stock_Status
    ws.conditional_formatting.add(
        f"I2:I{last_row}",
        CellIsRule(operator="equal", formula=['"OVERSTOCK"'],
                   fill=make_fill("FFF2CC"), font=Font(bold=True, color=C_ORANGE))
    )
    ws.conditional_formatting.add(
        f"I2:I{last_row}",
        CellIsRule(operator="equal", formula=['"LOW STOCK"'],
                   fill=make_fill(C_LIGHT_RED), font=Font(bold=True, color=C_RED))
    )
    ws.conditional_formatting.add(
        f"I2:I{last_row}",
        CellIsRule(operator="equal", formula=['"BALANCED"'],
                   fill=make_fill(C_LIGHT_GREEN), font=Font(bold=True, color=C_GREEN))
    )

    # Suggested_Action
    ws.conditional_formatting.add(
        f"J2:J{last_row}",
        CellIsRule(operator="equal", formula=['"Send Out"'],
                   fill=make_fill("FFF2CC"), font=Font(bold=True, color=C_ORANGE))
    )
    ws.conditional_formatting.add(
        f"J2:J{last_row}",
        CellIsRule(operator="equal", formula=['"Request Stock"'],
                   fill=make_fill(C_LIGHT_RED), font=Font(bold=True, color=C_RED))
    )
    ws.conditional_formatting.add(
        f"J2:J{last_row}",
        CellIsRule(operator="equal", formula=['"No Action"'],
                   fill=make_fill(C_LIGHT_GREEN), font=Font(bold=True, color=C_GREEN))
    )

    add_table(ws, f"A1:K{last_row}", "Consolidation_Planner_Table", "TableStyleMedium9")

    col_widths(ws, {
        "A": 24, "B": 20, "C": 12, "D": 14, "E": 36,
        "F": 12, "G": 16, "H": 16, "I": 14, "J": 18, "K": 22
    })
    freeze_and_filter(ws, "A2", f"A1:K{last_row}")


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 10: Dashboard
# ══════════════════════════════════════════════════════════════════════════════
def build_dashboard(wb):
    ws = wb.create_sheet("Dashboard")
    set_tab_color(ws, "E94560")
    ws.sheet_view.showGridLines = False

    # Title
    ws.merge_cells("A1:P1")
    c = ws["A1"]
    c.value = "HIGH CULTURED — STOCK CONSOLIDATION DASHBOARD"
    c.font = Font(name="Calibri", bold=True, size=22, color=C_WHITE)
    c.fill = make_fill(C_DARK_BG)
    c.alignment = make_align("center","center")
    ws.row_dimensions[1].height = 50

    ws.merge_cells("A2:P2")
    c2 = ws["A2"]
    c2.value = f"As of: {datetime.date.today().strftime('%d %B %Y')}   |   Prepared by: Merchandising Team"
    c2.font = Font(name="Calibri", size=11, color="AAAAAA", italic=True)
    c2.fill = make_fill(C_MID)
    c2.alignment = make_align("center","center")
    ws.row_dimensions[2].height = 24

    # ── KPI Tiles ──────────────────────────────────────────────────────────
    ws.row_dimensions[3].height = 10  # spacer

    kpi_configs = [
        # (start_col, end_col, label, formula, bg, fg)
        (1,  3,  "TOTAL OPENING STOCK",
         '=IFERROR(SUM(Opening_Stock_Table[Opening_Qty]),0)', "1A1A2E", C_WHITE),
        (4,  6,  "TOTAL TRANSFER IN",
         '=IFERROR(SUMIF(Transfer_Entry_Table[Status],"Completed",Transfer_Entry_Table[Transfer_Qty]),0)', "0F3460", C_WHITE),
        (7,  9,  "TOTAL TRANSFER OUT",
         '=IFERROR(SUMIF(Transfer_Entry_Table[Status],"Completed",Transfer_Entry_Table[Transfer_Qty]),0)', C_ACCENT, C_WHITE),
        (10, 12, "TOTAL ADJ ADD",
         '=IFERROR(SUMIF(Stock_Adjustment_Table[Adjustment_Type],"ADD",Stock_Adjustment_Table[Adjustment_Qty]),0)', "117A65", C_WHITE),
        (13, 15, "TOTAL ADJ MINUS",
         '=IFERROR(SUMIF(Stock_Adjustment_Table[Adjustment_Type],"MINUS",Stock_Adjustment_Table[Adjustment_Qty]),0)', "922B21", C_WHITE),
    ]

    for start_col, end_col, label, formula, bg, fg in kpi_configs:
        # Label row
        ws.merge_cells(start_row=4, start_column=start_col, end_row=4, end_column=end_col)
        lbl = ws.cell(row=4, column=start_col, value=label)
        lbl.font = Font(name="Calibri", bold=True, size=9, color="AAAAAA")
        lbl.fill = make_fill(C_DARK_BG)
        lbl.alignment = make_align("center","center")
        ws.row_dimensions[4].height = 18

        # Value row
        ws.merge_cells(start_row=5, start_column=start_col, end_row=6, end_column=end_col)
        val = ws.cell(row=5, column=start_col, value=formula)
        val.font = Font(name="Calibri", bold=True, size=28, color=fg)
        val.fill = make_fill(bg)
        val.alignment = make_align("center","center")
        ws.row_dimensions[5].height = 36
        ws.row_dimensions[6].height = 36

    # Total Final Stock - large center tile
    ws.merge_cells("A7:P7")
    c7 = ws["A7"]
    c7.value = "TOTAL FINAL STOCK"
    c7.font = Font(name="Calibri", bold=True, size=10, color="AAAAAA")
    c7.fill = make_fill(C_DARK_BG)
    c7.alignment = make_align("center","center")
    ws.row_dimensions[7].height = 18

    ws.merge_cells("A8:P9")
    c8 = ws["A8"]
    c8.value = (
        '=IFERROR(SUM(Opening_Stock_Table[Opening_Qty])'
        '+SUMIF(Transfer_Entry_Table[Status],"Completed",Transfer_Entry_Table[Transfer_Qty])'
        '-SUMIF(Transfer_Entry_Table[Status],"Completed",Transfer_Entry_Table[Transfer_Qty])'
        '+SUMIF(Stock_Adjustment_Table[Adjustment_Type],"ADD",Stock_Adjustment_Table[Adjustment_Qty])'
        '-SUMIF(Stock_Adjustment_Table[Adjustment_Type],"MINUS",Stock_Adjustment_Table[Adjustment_Qty]),0)'
    )
    c8.font = Font(name="Calibri", bold=True, size=36, color=C_WHITE)
    c8.fill = make_fill(C_LIGHT_BG)
    c8.alignment = make_align("center","center")
    ws.row_dimensions[8].height = 40
    ws.row_dimensions[9].height = 40

    # ── Stock Status Counts ─────────────────────────────────────────────────
    ws.row_dimensions[10].height = 10

    status_kpis = [
        (1,  5,  "OVERSTOCK SKU-STORES",
         '=IFERROR(COUNTIF(Consolidation_Planner_Table[Stock_Status],"OVERSTOCK"),0)', "7D6608", "F9E79F"),
        (6,  10, "LOW STOCK SKU-STORES",
         '=IFERROR(COUNTIF(Consolidation_Planner_Table[Stock_Status],"LOW STOCK"),0)', "922B21", "FADBD8"),
        (11, 15, "BALANCED SKU-STORES",
         '=IFERROR(COUNTIF(Consolidation_Planner_Table[Stock_Status],"BALANCED"),0)', "117A65", "D5F5E3"),
    ]

    for start_col, end_col, label, formula, fg, bg in status_kpis:
        ws.merge_cells(start_row=11, start_column=start_col, end_row=11, end_column=end_col)
        lbl = ws.cell(row=11, column=start_col, value=label)
        lbl.font = Font(name="Calibri", bold=True, size=9, color="888888")
        lbl.fill = make_fill(C_DARK_BG)
        lbl.alignment = make_align("center","center")
        ws.row_dimensions[11].height = 18

        ws.merge_cells(start_row=12, start_column=start_col, end_row=13, end_column=end_col)
        val = ws.cell(row=12, column=start_col, value=formula)
        val.font = Font(name="Calibri", bold=True, size=30, color=fg)
        val.fill = make_fill(bg)
        val.alignment = make_align("center","center")
        ws.row_dimensions[12].height = 36
        ws.row_dimensions[13].height = 36

    # ── Summary Table ────────────────────────────────────────────────────────
    ws.row_dimensions[14].height = 10

    ws.merge_cells("A15:H15")
    st = ws["A15"]
    st.value = "STOCK SUMMARY BY STORE"
    st.font = Font(name="Calibri", bold=True, size=12, color=C_WHITE)
    st.fill = make_fill(C_LIGHT_BG)
    st.alignment = make_align("left","center")
    ws.row_dimensions[15].height = 24

    sum_headers = ["Store_Code", "Store_Name", "Opening", "Trans In", "Trans Out", "Adj Add", "Adj Minus", "Final Qty"]
    for i, h in enumerate(sum_headers, 1):
        c = ws.cell(row=16, column=i, value=h)
        c.font = Font(name="Calibri", bold=True, size=9, color=C_WHITE)
        c.fill = make_fill(C_MID)
        c.alignment = make_align("center","center")
        c.border = make_border()
    ws.row_dimensions[16].height = 18

    store_name_map = {s[0]: s[1] for s in STORES}
    stores_with_data = ["CS1", "CS3", "CS10", "CS12"]

    for r_idx, store_code in enumerate(stores_with_data, 17):
        store_name = store_name_map.get(store_code, "")
        ws.cell(row=r_idx, column=1, value=store_code).font = Font(name="Calibri", size=9, bold=True)
        ws.cell(row=r_idx, column=2, value=store_name).font = Font(name="Calibri", size=9)
        ws.cell(row=r_idx, column=1).alignment = make_align("center","center")
        ws.cell(row=r_idx, column=2).alignment = make_align("left","center")

        # Opening
        c3 = ws.cell(row=r_idx, column=3)
        c3.value = f'=IFERROR(SUMIF(Opening_Stock_Table[Store_Code],A{r_idx},Opening_Stock_Table[Opening_Qty]),0)'
        c3.alignment = make_align("center","center"); c3.number_format = "0"

        # Trans In
        c4 = ws.cell(row=r_idx, column=4)
        c4.value = f'=IFERROR(SUMPRODUCT((Transfer_Entry_Table[To_Store_Code]=A{r_idx})*(Transfer_Entry_Table[Status]="Completed")*Transfer_Entry_Table[Transfer_Qty]),0)'
        c4.alignment = make_align("center","center"); c4.number_format = "0"

        # Trans Out
        c5 = ws.cell(row=r_idx, column=5)
        c5.value = f'=IFERROR(SUMPRODUCT((Transfer_Entry_Table[From_Store_Code]=A{r_idx})*(Transfer_Entry_Table[Status]="Completed")*Transfer_Entry_Table[Transfer_Qty]),0)'
        c5.alignment = make_align("center","center"); c5.number_format = "0"

        # Adj Add
        c6 = ws.cell(row=r_idx, column=6)
        c6.value = f'=IFERROR(SUMPRODUCT((Stock_Adjustment_Table[Store_Code]=A{r_idx})*(Stock_Adjustment_Table[Adjustment_Type]="ADD")*Stock_Adjustment_Table[Adjustment_Qty]),0)'
        c6.alignment = make_align("center","center"); c6.number_format = "0"

        # Adj Minus
        c7 = ws.cell(row=r_idx, column=7)
        c7.value = f'=IFERROR(SUMPRODUCT((Stock_Adjustment_Table[Store_Code]=A{r_idx})*(Stock_Adjustment_Table[Adjustment_Type]="MINUS")*Stock_Adjustment_Table[Adjustment_Qty]),0)'
        c7.alignment = make_align("center","center"); c7.number_format = "0"

        # Final
        c8r = ws.cell(row=r_idx, column=8)
        c8r.value = f'=C{r_idx}+D{r_idx}-E{r_idx}+F{r_idx}-G{r_idx}'
        c8r.font = Font(name="Calibri", size=9, bold=True, color=C_DARK_BG)
        c8r.fill = make_fill("DDEEFF")
        c8r.alignment = make_align("center","center"); c8r.number_format = "0"

        for col in range(1, 9):
            ws.cell(row=r_idx, column=col).border = make_border()
            if col not in [1, 2, 8]:
                ws.cell(row=r_idx, column=col).font = Font(name="Calibri", size=9)
            if r_idx % 2 == 0:
                if col not in [8]:
                    ws.cell(row=r_idx, column=col).fill = make_fill("F0F0F0")
        ws.row_dimensions[r_idx].height = 18

    # ── Bar Chart: Final Stock by Store ─────────────────────────────────────
    chart_start_row = 17
    chart_end_row   = 17 + len(stores_with_data) - 1

    chart1 = BarChart()
    chart1.type = "col"
    chart1.title = "Final Stock by Store"
    chart1.style = 10
    chart1.y_axis.title = "Qty"
    chart1.x_axis.title = "Store"
    chart1.height = 12
    chart1.width  = 20

    data_ref = Reference(ws, min_col=8, min_row=16, max_row=chart_end_row)
    cats_ref = Reference(ws, min_col=1, min_row=chart_start_row, max_row=chart_end_row)
    chart1.add_data(data_ref, titles_from_data=True)
    chart1.set_categories(cats_ref)
    chart1.shape = 4
    ws.add_chart(chart1, "I15")

    # ── Column widths for Dashboard ──────────────────────────────────────────
    for col_letter in ["A","B","C","D","E","F","G","H"]:
        ws.column_dimensions[col_letter].width = 14
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["I"].width = 4
    for col_letter in ["J","K","L","M","N","O","P"]:
        ws.column_dimensions[col_letter].width = 10

    ws.freeze_panes = "A3"


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════
def main():
    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    print("Building README...")
    build_readme(wb)

    print("Building Store_List...")
    build_store_list(wb)

    print("Building SKU_Master...")
    build_sku_master(wb)

    print("Building Opening_Stock...")
    build_opening_stock(wb)

    print("Building Transfer_Entry...")
    build_transfer_entry(wb)

    print("Building Stock_Adjustment...")
    build_stock_adjustment(wb)

    print("Building Stock_Movement_Log...")
    build_movement_log(wb)

    print("Building Stock_Balance...")
    build_stock_balance(wb)

    print("Building Consolidation_Planner...")
    build_consolidation_planner(wb)

    print("Building Dashboard...")
    build_dashboard(wb)

    print(f"Saving to {OUTPUT_FILE}...")
    wb.save(OUTPUT_FILE)
    print("Done!")

if __name__ == "__main__":
    main()
