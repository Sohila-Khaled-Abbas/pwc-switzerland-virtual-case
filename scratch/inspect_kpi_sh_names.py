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
    print(f'=== {cname} KPI Shapes ===')
    for sh in ws.Shapes:
        if any(k in sh.Name.lower() for k in ['value', 'ban', 'val']):
            txt = ''
            try:
                txt = sh.TextFrame2.TextRange.Text.replace('\r', ' ').replace('\n', ' ')
            except:
                pass
            print(f'  Shape: {sh.Name}, Text: {txt}, Top: {sh.Top:.1f}, Left: {sh.Left:.1f}')

wb.Close(False)
excel.Quit()
