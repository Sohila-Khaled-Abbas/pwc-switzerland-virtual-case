import win32com.client as win32
import os

wb_path = os.path.abspath('PWC_Switzerland_Virtual_Case.xlsm')
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(wb_path)

cockpits = ['04_CustomerRetention_Cockpit', '05_DiversityInclusion_Cockpit']

for cname in cockpits:
    ws = wb.Worksheets(cname)
    print(f'=== {cname} ===')
    for col_letter, col_idx in [('AA', 27), ('AB', 28), ('AC', 29), ('AD', 30), ('AE', 31)]:
        cell = ws.Cells(65, col_idx)
        print(f'  {col_letter}65: Formula={cell.Formula}, Value={cell.Value}, Text={cell.Text}')

wb.Close(False)
excel.Quit()
