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
        addr = pt.TableRange2.Address if pt.TableRange2 else 'None'
        print(f'{co.Name} -> PT: {pt.Name}, Sheet: {pt.Parent.Name}, Range: {addr}')
    except Exception as e:
        print(f'{co.Name} -> No PT: {e}')

print('\nAll PivotTables in wb:')
for ws_i in wb.Worksheets:
    print(f'Sheet {ws_i.Name}: count = {ws_i.PivotTables().Count}')
    for pt_i in ws_i.PivotTables():
        print(f'  PT: {pt_i.Name}')

wb.Close(False)
excel.Quit()
