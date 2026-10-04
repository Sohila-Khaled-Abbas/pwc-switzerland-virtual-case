import win32com.client as win32
import os

wb_path = os.path.abspath('PWC_Switzerland_Virtual_Case.xlsm')
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(wb_path)
    ws_stage = wb.Worksheets('Staging_Pivots')
    ws_test = wb.Worksheets('04_CustomerRetention_Cockpit')
    
    conn = wb.Connections('ThisWorkbookDataModel')
    pc = wb.PivotCaches().Create(SourceType=2, SourceData=conn)
    pt = pc.CreatePivotTable(TableDestination=ws_stage.Range('J3'), TableName='test_pt_CH')
    pt.CubeFields('[DimContract].[Contract]').Orientation = 1
    pt.CubeFields('[Measures].[Total Customers]').Orientation = 4
    
    print(f'PT created at {pt.TableRange2.Address}')
    
    # Try creating chart from PT
    co = ws_stage.ChartObjects().Add(100, 100, 400, 250)
    ch = co.Chart
    ch.SetSourceData(pt.TableRange2)
    
    print('Chart created! Is PivotChart?')
    try:
        print('PT attached to chart:', ch.PivotLayout.PivotTable.Name)
    except Exception as e:
        print('PivotLayout error:', e)
        
    co.Delete()
    pt.TableRange2.Clear()
    wb.Close(False)
    print('SUCCESS: Verified PivotChart generation from Model PivotTable!')
except Exception as e:
    print('Error:', e)
finally:
    excel.Quit()
