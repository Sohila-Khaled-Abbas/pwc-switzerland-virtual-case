import win32com.client as win32
import os
import traceback

wb_path = os.path.abspath('PWC_Switzerland_Virtual_Case.xlsm')
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(wb_path)
    ws_stage = wb.Worksheets('Staging_Pivots')
    
    conn = wb.Connections('ThisWorkbookDataModel')
    pc = wb.PivotCaches().Create(SourceType=2, SourceData=conn)
    pt = pc.CreatePivotTable(TableDestination=ws_stage.Range('J3'), TableName='test_pt_CH')
    pt.CubeFields('[DimContract].[Contract]').Orientation = 1
    pt.CubeFields('[Measures].[Total Customers]').Orientation = 4
    
    print(f'PT created at {pt.TableRange2.Address}')
    
    # Try AddChart2
    try:
        sh = ws_stage.Shapes.AddChart2(201, 51, 100, 100, 400, 250) # xlColumnClustered
        ch = sh.Chart
        ch.SetSourceData(pt.TableRange2)
        print('AddChart2 success!')
        sh.Delete()
    except Exception as e:
        print('AddChart2 failed:')
        traceback.print_exc()

    pt.TableRange2.Clear()
    wb.Close(False)
except Exception as e:
    traceback.print_exc()
finally:
    excel.Quit()
