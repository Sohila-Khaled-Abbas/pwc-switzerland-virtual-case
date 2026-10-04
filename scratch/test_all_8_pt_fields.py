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
    
    configs = [
        ('pt_CH_Contract', 'J3', '[DimContract].[Contract]', 1, None, 0, ['[Measures].[Total Customers]', '[Measures].[Churn Rate %]']),
        ('pt_CH_Tenure', 'N3', '[Fact_Churn].[Tenure_Cohort]', 1, None, 0, ['[Measures].[Churn Rate %]']),
        ('pt_CH_Payment', 'R3', '[Fact_Churn].[PaymentMethod]', 1, None, 0, ['[Measures].[Churn Rate %]']),
        ('pt_CH_Service', 'V3', '[Fact_Churn].[InternetService]', 1, None, 0, ['[Measures].[Total Customers]', '[Measures].[Churn Rate %]']),
        ('pt_DI_Funnel', 'Z3', '[Dim_CareerLadder].[Base_Job_Level]', 1, '[Fact_Employees].[Gender]', 2, ['[Measures].[Total Employees]']),
        ('pt_DI_Parity', 'AD3', '[DimDepartment].[Department]', 1, None, 0, ['[Measures].[Female Representation %]']),
        ('pt_DI_Promo', 'AH3', '[Dim_CareerLadder].[Base_Job_Level]', 1, None, 0, ['[Measures].[Female Promotion %]']),
        ('pt_DI_Rating', 'AL3', '[Fact_Employees].[FY20_Rating]', 1, None, 0, ['[Measures].[Turnover Rate %]']),
    ]
    
    pts = []
    for pt_name, dest, r_field, r_orient, c_field, c_orient, meas_list in configs:
        pc = wb.PivotCaches().Create(SourceType=2, SourceData=conn)
        pt = pc.CreatePivotTable(TableDestination=ws_stage.Range(dest), TableName=pt_name)
        if r_field:
            pt.CubeFields(r_field).Orientation = r_orient
        if c_field:
            pt.CubeFields(c_field).Orientation = c_orient
        for m in meas_list:
            pt.CubeFields(m).Orientation = 4
        pts.append(pt)
        print(f'Successfully built {pt.Name} at {dest} (Range: {pt.TableRange2.Address})')
        
    for pt in pts:
        pt.TableRange2.Clear()
    wb.Close(False)
    print('SUCCESS: All 8 PivotTables configured with DAX measures and CubeFields!')
except Exception as e:
    print('Error:', e)
finally:
    excel.Quit()
