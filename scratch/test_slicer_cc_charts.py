import win32com.client as win32
import os

wb_path = os.path.abspath('PWC_Switzerland_Virtual_Case.xlsm')
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(wb_path)
    ws_cc = wb.Worksheets('03_CallCenter_Cockpit')
    ws_stage = wb.Worksheets('Staging_Pivots')
    pt_agent = ws_stage.PivotTables('pt_Agent')
    
    sc = wb.SlicerCaches.Add2(Source=pt_agent, SourceField='[DimTopic].[Topic].[Topic]', Name='test_sc')
    
    print('Connecting test_sc to charts on 03_CallCenter_Cockpit:')
    for co in ws_cc.ChartObjects():
        try:
            pt = co.Chart.PivotLayout.PivotTable
            sc.PivotTables.AddPivotTable(pt)
            print(f'  Connected to {co.Name} (PT: {pt.Name})!')
        except Exception as e:
            print(f'  Could not connect to {co.Name}: {e}')
            
    sc.Delete()
    wb.Close(False)
except Exception as e:
    print('Error:', e)
finally:
    excel.Quit()
