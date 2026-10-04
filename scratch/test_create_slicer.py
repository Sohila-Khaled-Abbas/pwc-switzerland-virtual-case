import win32com.client
import os

wb_path = os.path.abspath("PWC_Switzerland_Virtual_Case.xlsm")
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(wb_path)
    ws_stage = wb.Worksheets("Staging_Pivots")
    ws_cc = wb.Worksheets("03_CallCenter_Cockpit")
    
    pt = ws_stage.PivotTables("pt_Agent")
    print(f"Using PivotTable: {pt.Name}")
    
    sc = wb.SlicerCaches.Add2(Source=pt, SourceField="[DimTopic].[Topic].[Topic]", Name="test_sc_Topic")
    print(f"Created SlicerCache: {sc.Name}")
    
    sl = sc.Slicers.Add(SlicerDestination=ws_cc, Name="test_sl_Topic", Caption="Inquiry Topic", Top=280.0, Left=36.0, Width=206.0, Height=140.0)
    print(f"Created Slicer on sheet: {sl.Name}, Caption: {sl.Caption}")
    
    # Clean up
    sl.Delete()
    print("Slicer deleted cleanly.")

    wb.Close(False)
    print("SUCCESS: Real Slicer creation verified!")
except Exception as e:
    print(f"Error: {e}")
finally:
    excel.Quit()
