import win32com.client as win32
import os

wb_path = os.path.abspath('PWC_Switzerland_Virtual_Case.xlsm')
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(wb_path)

ws = wb.Worksheets('00_Home_Portal')
print('=== 00_Home_Portal Shapes ===')
for sh in ws.Shapes:
    txt = ''
    try:
        txt = sh.TextFrame2.TextRange.Text.replace('\r', ' ').replace('\n', ' ')
    except:
        pass
    print(f'  Name: {sh.Name}, Top: {sh.Top:.1f}, Left: {sh.Left:.1f}, Width: {sh.Width:.1f}, Height: {sh.Height:.1f}, Text: {txt[:40]}')

wb.Close(False)
excel.Quit()
