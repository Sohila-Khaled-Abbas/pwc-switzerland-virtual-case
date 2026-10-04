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

print(f'PivotCache SourceType: {pc.SourceType}')

# Try creating a new PivotTable at J3 for Contract Risk
dest_range = ws_stage.Range('J3')
pt_contract = pc.CreatePivotTable(TableDestination=dest_range, TableName='pt_CH_Contract')

# Add Row Field: [DimContract].[Contract]
cube_field_contract = pt_contract.CubeFields('[DimContract].[Contract]')
cube_field_contract.Orientation = 1 # xlRowField

# Add Measure: [Measures].[Total Customers]
cube_meas1 = pt_contract.CubeFields('[Measures].[Total Customers]')
cube_meas1.Orientation = 4 # xlDataField

# Add Measure: [Measures].[Churn Rate %]
cube_meas2 = pt_contract.CubeFields('[Measures].[Churn Rate %]')
cube_meas2.Orientation = 4 # xlDataField

print(f'Created pt_CH_Contract at {pt_contract.TableRange2.Address}')

# Check values
for r in range(3, 8):
    vals = [ws_stage.Cells(r, c).Value for c in range(10, 14)]
    print(f'Row {r}: {vals}')

wb.Close(False)
excel.Quit()
