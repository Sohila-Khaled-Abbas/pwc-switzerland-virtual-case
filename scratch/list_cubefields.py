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
    pt = pc.CreatePivotTable(TableDestination=ws_stage.Range("J3"), TableName="inspect_fields_pt")
    
    print("=== TOTAL CUBEFIELDS ===", pt.CubeFields.Count)
    churn_fields = []
    emp_fields = []
    measures = []
    
    for cf in pt.CubeFields:
        name = cf.Name
        if "Fact_Churn" in name or "DimContract" in name:
            churn_fields.append(name)
        elif "Employees" in name or "Career" in name or "PRA" in name or "Department" in name:
            emp_fields.append(name)
        elif "Measures" in name:
            measures.append(name)
            
    print("\n--- CHURN FIELDS ---")
    for f in churn_fields: print(" ", f)
    
    print("\n--- EMPLOYEE FIELDS ---")
    for f in emp_fields: print(" ", f)
    
    ws_stage.Range("J1:N10").Clear()
    wb.Close(False)
except Exception as e:
    print(e)
finally:
    excel.Quit()
