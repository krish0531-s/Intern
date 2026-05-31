from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws["A1"] = "Hello"
wb.save("test.xlsx")

print("Success")