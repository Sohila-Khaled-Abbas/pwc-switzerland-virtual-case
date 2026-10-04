import win32com.client as win32
import os

wb_path = os.path.abspath('PWC_Switzerland_Virtual_Case.xlsm')
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(wb_path)

ws = wb.Worksheets('Staging_Pivots')
print('UsedRange:', ws.UsedRange.Address)

# Print non-empty cells
for r in range(1, 35):
    row_vals = [ws.Cells(r, c).Value for c in range(1, 15)]
    if any(v is not None for v in row_vals):
        print(f'Row {r}: {[v for v in row_vals if v is not None]}')

wb.Close(False)
excel.Quit()
