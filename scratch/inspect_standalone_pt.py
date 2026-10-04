import win32com.client as win32
import os

wb_path = os.path.abspath('PWC_Switzerland_Virtual_Case.xlsm')
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(wb_path)

ws = wb.Worksheets('03_CallCenter_Cockpit')
for co in ws.ChartObjects():
    ch = co.Chart
    try:
        pt = ch.PivotLayout.PivotTable
        print(f'{co.Name} -> PT Name: {pt.Name}, Parent: {type(pt.Parent)}')
    except Exception as e:
        print(f'{co.Name} -> Error: {e}')

wb.Close(False)
excel.Quit()
