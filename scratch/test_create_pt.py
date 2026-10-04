import win32com.client
import os

wb_path = os.path.abspath("PWC_Switzerland_Virtual_Case.xlsm")
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(wb_path)
    ws_stage = wb.Worksheets("Staging_Pivots")
    
    # Check if pt_CH_Contract can be created
    conn = wb.Connections("ThisWorkbookDataModel")
    pc = wb.PivotCaches().Create(SourceType=2, SourceData=conn)
    print("Created PivotCache from ThisWorkbookDataModel!")
    
    pt = pc.CreatePivotTable(TableDestination=ws_stage.Range("J3"), TableName="test_pt_CH_Contract")
    print(f"Created PivotTable: {pt.Name}")
    
    # Add Contract to Rows
    pt.CubeFields("[DimContract].[Contract]").Orientation = 1 # xlRowField
    print("Added Contract to Rows")
    
    # Add Total Customers to Values
    pt.CubeFields("[Measures].[Total Customers]").Orientation = 4 # xlDataField
    print("Added Total Customers to DataField")
    
    # Add Churn Rate % to Values
    pt.CubeFields("[Measures].[Churn Rate %]").Orientation = 4 # xlDataField
    print("Added Churn Rate % to DataField")
    
    # Delete test pivot table so we don't pollute
    ws_stage.Range("J1:N12").Clear()
    print("Cleaned up test range.")
    
    wb.Close(False)
    print("SUCCESS: PivotTable creation from Data Model verified!")
except Exception as e:
    print(f"Error: {e}")
finally:
    excel.Quit()
