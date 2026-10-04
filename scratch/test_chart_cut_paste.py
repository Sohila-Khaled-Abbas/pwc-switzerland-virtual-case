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
    ws_target = wb.Worksheets('04_CustomerRetention_Cockpit')
    
    conn = wb.Connections('ThisWorkbookDataModel')
    pc = wb.PivotCaches().Create(SourceType=2, SourceData=conn)
    pt = pc.CreatePivotTable(TableDestination=ws_stage.Range('J3'), TableName='test_pt_CH')
    pt.CubeFields('[DimContract].[Contract]').Orientation = 1
    pt.CubeFields('[Measures].[Total Customers]').Orientation = 4
    
    ws_stage.Activate()
    ws_stage.Range('J4').Select()
    sh = ws_stage.Shapes.AddChart2(201, 51, 100, 100, 400, 250)
    
    # Cut and paste to target
    sh.Cut()
    ws_target.Activate()
    ws_target.Paste()
    
    # Check pasted shape
    pasted_sh = ws_target.Shapes(ws_target.Shapes.Count)
    print('Pasted shape name:', pasted_sh.Name)
    ch = pasted_sh.Chart
    print('Attached PT:', ch.PivotLayout.PivotTable.Name)
    
    # Clean up
    pasted_sh.Delete()
    pt.TableRange2.Clear()
    wb.Close(False)
    print('SUCCESS: PivotChart cut & paste verified!')
except Exception as e:
    traceback.print_exc()
finally:
    excel.Quit()
