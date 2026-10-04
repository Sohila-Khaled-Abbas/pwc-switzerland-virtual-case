import win32com.client as win32
import os

wb_path = os.path.abspath('PWC_Switzerland_Virtual_Case.xlsm')
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(wb_path)

sheets = ['03_CallCenter_Cockpit', '04_CustomerRetention_Cockpit', '05_DiversityInclusion_Cockpit']

for sname in sheets:
    ws = wb.Worksheets(sname)
    print(f'=== {sname} Shapes ===')
    for sh in ws.Shapes:
        txt = ''
        try:
            txt = sh.TextFrame2.TextRange.Text.replace('\r', ' ').replace('\n', ' ')
        except Exception:
            pass
        if 'slicer' in sh.Name.lower() or 'slot' in sh.Name.lower() or 'slicer' in txt.lower() or 'slot' in txt.lower() or 'filter' in sh.Name.lower():
            print(f'  Name: {sh.Name}, Left: {sh.Left}, Top: {sh.Top}, Width: {sh.Width}, Height: {sh.Height}, Text: {txt[:40]}')

wb.Close(False)
excel.Quit()
