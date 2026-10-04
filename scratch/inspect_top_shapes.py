import win32com.client as win32
import os

wb_path = os.path.abspath('PWC_Switzerland_Virtual_Case.xlsm')
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(wb_path)

sheets = [
    '00_Home_Portal',
    '01_Business_Domains',
    '02_Metadata_&_KPI_Catalog',
    '03_CallCenter_Cockpit',
    '04_CustomerRetention_Cockpit',
    '05_DiversityInclusion_Cockpit'
]

for sname in sheets:
    ws = wb.Worksheets(sname)
    print(f'=== {sname} Top Shapes (Top < 80) ===')
    for sh in ws.Shapes:
        if sh.Top < 80:
            txt = ''
            try:
                txt = sh.TextFrame2.TextRange.Text.replace('\r', ' ').replace('\n', ' ')
            except Exception:
                pass
            print(f'  Name: {sh.Name}, Left: {sh.Left:.1f}, Top: {sh.Top:.1f}, Width: {sh.Width:.1f}, Height: {sh.Height:.1f}, Text: {txt[:30]}')

wb.Close(False)
excel.Quit()
