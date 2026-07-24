# EXCEL_OPERATION_SCRIPT
# Name: Test_Update_Status
# Version: 1.0

from openpyxl import load_workbook

FILE_NAME = "input.xlsx"

wb = load_workbook(FILE_NAME)
ws = wb["Sheet1"]

ws["B2"] = "Teszt"
ws["C2"] = 123

wb.save(FILE_NAME)
