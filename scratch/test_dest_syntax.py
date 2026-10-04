import win32com.client as win32
import os

wb_path = os.path.abspath('PWC_Switzerland_Virtual_Case.xlsm')
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(wb_path)

ws_stage = wb.Worksheets('Staging_Pivots')
pt_agent = ws_stage.PivotTables('pt_Agent')
pc = pt_agent.PivotCache()

print(f'PivotCache Connection: {pc.Connection}')

# Try various TableDestination formats
for dest in ['Staging_Pivots!J3', 'Staging_Pivots!R3C10', ws_stage.Range('J3')]:
    try:
        pt = pc.CreatePivotTable(dest, 'TestPT')
        print(f'SUCCESS with {dest}: {pt.Name}')
        pt.TableRange2.Clear()
        break
    except Exception as e:
        print(f'FAILED with {dest}: {e}')

wb.Close(False)
excel.Quit()
