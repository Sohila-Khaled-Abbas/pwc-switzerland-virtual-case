import win32com.client
import os

wb_path = os.path.abspath("PWC_Switzerland_Virtual_Case.xlsm")
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(wb_path)
    
    # Check DAX measures or evaluate expressions using CUBEVALUE in scratch cells
    ws = wb.Worksheets("Staging_Pivots")
    test_cell = ws.Range("Z100")
    
    queries = [
        ("[Measures].[At-Risk MRR]", '=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[At-Risk MRR]")'),
        ("[Measures].[Total Monthly Charges]", '=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Total Monthly Charges]")'),
        ("[Measures].[Total Lifetime Charges]", '=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Total Lifetime Charges]")'),
        ("[Measures].[Financial Churn Rate %]", '=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Financial Churn Rate %]")'),
        ("Churned TotalCharges", '=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Total Lifetime Charges]", "[Fact_Churn].[Churn].&[Yes]")'),
    ]
    
    for name, formula in queries:
        test_cell.Formula = formula
        excel.Calculate()
        val = test_cell.Value
        print(f"{name}: {val}")
        
    test_cell.Clear()
    
    # Check Home Portal shape texts
    ws_home = wb.Worksheets("00_Home_Portal")
    print("\n--- Home Portal Ticker Shapes ---")
    for shp in ws_home.Shapes:
        if "Ticker_" in shp.Name:
            try:
                print(f"{shp.Name}: '{shp.TextFrame2.TextRange.Text.strip()}'")
            except: pass
            
    # Check Business Domains shape texts
    ws_dom = wb.Worksheets("01_Business_Domains")
    print("\n--- Business Domains Shapes ---")
    for shp in ws_dom.Shapes:
        if "Metric_" in shp.Name or "Value_" in shp.Name or "Card_" in shp.Name:
            try:
                print(f"{shp.Name}: '{shp.TextFrame2.TextRange.Text.strip()}'")
            except: pass

    wb.Close(False)
except Exception as e:
    print(e)
finally:
    excel.Quit()
