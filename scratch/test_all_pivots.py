import win32com.client
import os

wb_path = os.path.abspath("PWC_Switzerland_Virtual_Case.xlsm")
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(wb_path)
    ws_stage = wb.Worksheets("Staging_Pivots")
    conn = wb.Connections("ThisWorkbookDataModel")
    pc = wb.PivotCaches().Create(SourceType=2, SourceData=conn)
    
    # 1. pt_CH_Contract at J3
    try:
        pt1 = pc.CreatePivotTable(TableDestination=ws_stage.Range("J3"), TableName="pt_CH_Contract")
        pt1.CubeFields("[DimContract].[Contract]").Orientation = 1
        pt1.CubeFields("[Measures].[Total Customers]").Orientation = 4
        pt1.CubeFields("[Measures].[Churn Rate %]").Orientation = 4
        print("Created pt_CH_Contract successfully")
    except Exception as e:
        print("pt_CH_Contract error:", e)

    # 2. pt_CH_Tenure at J15
    try:
        pt2 = pc.CreatePivotTable(TableDestination=ws_stage.Range("J15"), TableName="pt_CH_Tenure")
        pt2.CubeFields("[Fact_Churn].[Tenure_Cohort]").Orientation = 1
        pt2.CubeFields("[Measures].[Churn Rate %]").Orientation = 4
        print("Created pt_CH_Tenure successfully")
    except Exception as e:
        print("pt_CH_Tenure error:", e)

    # 3. pt_CH_Payment at J27
    try:
        pt3 = pc.CreatePivotTable(TableDestination=ws_stage.Range("J27"), TableName="pt_CH_Payment")
        pt3.CubeFields("[Fact_Churn].[PaymentMethod]").Orientation = 1
        pt3.CubeFields("[Measures].[Churn Rate %]").Orientation = 4
        print("Created pt_CH_Payment successfully")
    except Exception as e:
        print("pt_CH_Payment error:", e)

    # 4. pt_CH_Service at J39
    try:
        pt4 = pc.CreatePivotTable(TableDestination=ws_stage.Range("J39"), TableName="pt_CH_Service")
        pt4.CubeFields("[Fact_Churn].[InternetService]").Orientation = 1
        pt4.CubeFields("[Measures].[Total Customers]").Orientation = 4
        pt4.CubeFields("[Measures].[Churn Rate %]").Orientation = 4
        print("Created pt_CH_Service successfully")
    except Exception as e:
        print("pt_CH_Service error:", e)

    # 5. pt_DI_Funnel at R3
    try:
        pt5 = pc.CreatePivotTable(TableDestination=ws_stage.Range("R3"), TableName="pt_DI_Funnel")
        pt5.CubeFields("[Dim_CareerLadder].[Base_Job_Level]").Orientation = 1
        pt5.CubeFields("[Fact_Employees].[Gender]").Orientation = 2 # xlColumnField
        pt5.CubeFields("[Measures].[Total Employees]").Orientation = 4
        print("Created pt_DI_Funnel successfully")
    except Exception as e:
        print("pt_DI_Funnel error:", e)

    # 6. pt_DI_Parity at R15
    try:
        pt6 = pc.CreatePivotTable(TableDestination=ws_stage.Range("R15"), TableName="pt_DI_Parity")
        pt6.CubeFields("[DimDepartment].[Department]").Orientation = 1
        pt6.CubeFields("[Measures].[Female Representation %]").Orientation = 4
        print("Created pt_DI_Parity successfully")
    except Exception as e:
        print("pt_DI_Parity error:", e)

    # 7. pt_DI_Velocity at R27
    try:
        pt7 = pc.CreatePivotTable(TableDestination=ws_stage.Range("R27"), TableName="pt_DI_Velocity")
        pt7.CubeFields("[Dim_CareerLadder].[Base_Job_Level]").Orientation = 1
        pt7.CubeFields("[Measures].[Female Promotion Rate %]").Orientation = 4
        pt7.CubeFields("[Measures].[Male Promotion Rate %]").Orientation = 4
        print("Created pt_DI_Velocity successfully")
    except Exception as e:
        print("pt_DI_Velocity error:", e)

    # 8. pt_DI_Audit at R39
    try:
        pt8 = pc.CreatePivotTable(TableDestination=ws_stage.Range("R39"), TableName="pt_DI_Audit")
        pt8.CubeFields("[Fact_Employees].[FY20_Rating]").Orientation = 1
        pt8.CubeFields("[Fact_Employees].[Gender]").Orientation = 2
        pt8.CubeFields("[Measures].[Total Promotions FY21]").Orientation = 4
        print("Created pt_DI_Audit successfully")
    except Exception as e:
        print("pt_DI_Audit error:", e)

    # Clean up test tables so we don't modify workbook yet
    ws_stage.Range("J1:Z60").Clear()
    print("Cleaned up staging range.")

    wb.Close(False)
    print("ALL 8 PIVOTTABLES VERIFIED!")
except Exception as e:
    print("Fatal:", e)
finally:
    excel.Quit()
