from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "tbl_pnl_dashboard"

headers = [
    "DUPLOAD_DATE",
    "FLAG",
    "FREQUENCY",
    "CM_ADTV",
    "FUTURES_ADTV",
    "OPTIONS_ADTV"
]

for col, header in enumerate(headers, start=1):
    ws.cell(row=1, column=col).value = header

wb.save("tbl_pnl_dashboard_template.xlsx")

print("Excel template created successfully!")