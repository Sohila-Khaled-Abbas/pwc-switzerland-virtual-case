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
    
    # Create PT 1
    pc1 = wb.PivotCaches().Create(SourceType=2, SourceData=conn)
    pt1 = pc1.CreatePivotTable(TableDestination=ws_stage.Range('J3'), TableName='pt1')
    print('PT1 created successfully!')
    
    # Try with a fresh PivotCache pc2
    pc2 = wb.PivotCaches().Create(SourceType=2, SourceData=conn)
    pt2 = pc2.CreatePivotTable(TableDestination=ws_stage.Range('J15'), TableName='pt2')
    print('PT2 created successfully with fresh PivotCache!')
    
    # Clean up
    pt1.TableRange2.Clear()
    pt2.TableRange2.Clear()
    wb.Close(False)
except Exception as e:
    print('Error:', e)
finally:
    excel.Quit()
