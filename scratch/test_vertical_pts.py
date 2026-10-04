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
    
    configs = [
        ('pt_CH_Contract', 'J3', '[DimContract].[Contract]', ['[Measures].[Total Customers]', '[Measures].[Churn Rate %]']),
        ('pt_CH_Tenure', 'J15', '[Fact_Churn].[Tenure_Cohort]', ['[Measures].[Churn Rate %]']),
        ('pt_CH_Payment', 'J27', '[Fact_Churn].[PaymentMethod]', ['[Measures].[Churn Rate %]']),
        ('pt_CH_Service', 'J39', '[Fact_Churn].[InternetService]', ['[Measures].[Total Customers]', '[Measures].[Churn Rate %]']),
        ('pt_DI_Funnel', 'J51', '[Dim_CareerLadder].[Base_Job_Level]', ['[Measures].[Total Employees]']),
        ('pt_DI_Parity', 'J63', '[DimDepartment].[Department]', ['[Measures].[Female Representation %]']),
        ('pt_DI_Promo', 'J75', '[Dim_CareerLadder].[Base_Job_Level]', ['[Measures].[Female Promotion %]']),
        ('pt_DI_Rating', 'J87', '[Fact_Employees].[FY20_Rating]', ['[Measures].[Turnover Rate %]']),
    ]
    
    for pt_name, cell_dest, row_field, meas_list in configs:
        try:
            pt = pc.CreatePivotTable(TableDestination=ws_stage.Range(cell_dest), TableName=pt_name)
            pt.CubeFields(row_field).Orientation = 1
            for m in meas_list:
                pt.CubeFields(m).Orientation = 4
            print(f'SUCCESS: Created {pt_name} at {cell_dest} (Range: {pt.TableRange2.Address})')
        except Exception as e:
            print(f'FAILED: {pt_name} at {cell_dest}: {e}')
            
    wb.Close(False)
except Exception as e:
    traceback.print_exc()
finally:
    excel.Quit()
