import win32com.client as win32
import os

wb_path = os.path.abspath('PWC_Switzerland_Virtual_Case.xlsm')
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(wb_path)
    ws_stage = wb.Worksheets('Staging_Pivots')
    conn = wb.Connections('ThisWorkbookDataModel')
    
    cols = ['J3', 'N3', 'R3', 'V3', 'Z3', 'AD3', 'AH3', 'AL3']
    pts = []
    for i, dest in enumerate(cols):
        pc = wb.PivotCaches().Create(SourceType=2, SourceData=conn)
        pt = pc.CreatePivotTable(TableDestination=ws_stage.Range(dest), TableName=f'test_pt_{i}')
        pts.append(pt)
        print(f'Successfully created {pt.Name} at {dest}!')
        
    for pt in pts:
        pt.TableRange2.Clear()
    wb.Close(False)
    print('ALL 8 PIVOT TABLES CREATED SIMULTANEOUSLY WITHOUT CONFLICT!')
except Exception as e:
    print('Error:', e)
finally:
    excel.Quit()
