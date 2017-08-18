import openpyxl
from openpyxl import Workbook
xl = openpyxl.load_workbook('cfay.xlsx')
print type(xl)
print xl.get_sheet_names()

sheet1 = xl.get_sheet_by_name(xl.get_sheet_names()[0])
print sheet1

sheet2 = xl.get_sheet_by_name(xl.get_sheet_names()[1])
print sheet2

sheet1['A2'] = 50
sheet1['A3'] = 3
sheet1['A4'] = "=SUM(A2, A3)"

print sheet1['A4'].value
xl.save("cfay.xlsx")