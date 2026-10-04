import win32com.client as win32
import os

wb_path = os.path.abspath('PWC_Switzerland_Virtual_Case.xlsm')
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(wb_path)

ws = wb.Worksheets('02_Metadata_&_KPI_Catalog')

print('Shapes on 02_Metadata_&_KPI_Catalog:')
for sh in ws.Shapes:
    txt = ''
    try:
        txt = sh.TextFrame2.TextRange.Text.replace('\r', ' ').replace('\n', ' ')
    except Exception:
        pass
    print(f'  Name: {sh.Name}, Top: {sh.Top:.1f}, Bottom: {(sh.Top + sh.Height):.1f}, Left: {sh.Left:.1f}, Width: {sh.Width:.1f}, Text: {txt[:40]}')

print('\nNon-empty cell rows:')
for r in range(1, 40):
    val_b = ws.Cells(r, 2).Value
    val_c = ws.Cells(r, 3).Value
    if val_b or val_c:
        print(f'  Row {r}: B={val_b}, C={val_c}')

wb.Close(False)
excel.Quit()
