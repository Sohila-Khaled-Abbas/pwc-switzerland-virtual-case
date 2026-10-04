import win32com.client as win32
import os

wb_path = os.path.abspath('PWC_Switzerland_Virtual_Case.xlsm')
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(wb_path)

for ws in wb.Worksheets:
    print(f'=== Sheet: {ws.Name} ===')
    for co in ws.ChartObjects():
        ch = co.Chart
        pt_name = 'None'
        try:
            pt_name = ch.PivotLayout.PivotTable.Name
        except Exception:
            pass
        title = ch.ChartTitle.Text if ch.HasTitle else 'NoTitle'
        print(f'  Chart: {co.Name}, Title: {title}, PT: {pt_name}')
        if pt_name == 'None':
            for s in ch.SeriesCollection():
                try:
                    print(f'    Series: {s.Name}, Formula: {s.Formula[:60]}...')
                except Exception:
                    pass

wb.Close(False)
excel.Quit()
