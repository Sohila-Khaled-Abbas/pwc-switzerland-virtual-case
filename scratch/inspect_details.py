import win32com.client
import os

wb_path = os.path.abspath("PWC_Switzerland_Virtual_Case.xlsm")
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(wb_path)
    
    print("=== PIVOT TABLES ON STAGING_PIVOTS ===")
    ws_stage = wb.Worksheets("Staging_Pivots")
    for pt in ws_stage.PivotTables():
        print(f"PivotTable: {pt.Name}, Range: {pt.TableRange2.Address}")
        row_fields = [f.Name for f in pt.RowFields]
        col_fields = [f.Name for f in pt.ColumnFields]
        data_fields = [f.Name for f in pt.DataFields]
        print(f"  Rows: {row_fields}, Cols: {col_fields}, Data: {data_fields}")

    print("\n=== CHARTS ON COCKPITS ===")
    for s_name in ["03_CallCenter_Cockpit", "04_CustomerRetention_Cockpit", "05_DiversityInclusion_Cockpit"]:
        ws = wb.Worksheets(s_name)
        print(f"\nSheet: {s_name} (Charts: {ws.ChartObjects().Count}, Pivots: {ws.PivotTables().Count})")
        for co in ws.ChartObjects():
            ch = co.Chart
            has_pivot = False
            pt_name = None
            try:
                if ch.PivotLayout:
                    has_pivot = True
                    pt_name = ch.PivotLayout.PivotTable.Name
            except Exception:
                pass
            print(f"  ChartObject: '{co.Name}', Type: {ch.ChartType}, HasPivot: {has_pivot}, PivotTable: {pt_name}")
            if not has_pivot:
                # check series
                sc_count = ch.SeriesCollection().Count
                print(f"    Non-Pivot Chart! Series count: {sc_count}")
                for s_idx in range(1, sc_count + 1):
                    s = ch.SeriesCollection(s_idx)
                    try:
                        print(f"      Series {s_idx}: Name='{s.Name}', Formula='{s.Formula}'")
                    except Exception as se:
                        print(f"      Series {s_idx}: {se}")

    print("\n=== CUBE / HELPER CELLS ===")
    for s_name in ["03_CallCenter_Cockpit", "04_CustomerRetention_Cockpit", "05_DiversityInclusion_Cockpit"]:
        ws = wb.Worksheets(s_name)
        print(f"\nHelper cells on {s_name}:")
        for col in ["AA", "AB", "AC", "AD", "AE"]:
            h_lbl = ws.Range(f"{col}64").Value
            h_val = ws.Range(f"{col}65").Value
            h_form = ws.Range(f"{col}65").Formula
            print(f"  {col}64: '{h_lbl}' | {col}65: '{h_val}' | Formula: {h_form}")

    print("\n=== SLICER CACHES IN WORKBOOK ===")
    print(f"Total SlicerCaches: {wb.SlicerCaches.Count}")
    for sc in wb.SlicerCaches:
        print(f"  Cache: Name='{sc.Name}', SourceName='{sc.SourceName}'")
        for s in sc.Slicers:
            print(f"    Slicer: Name='{s.Name}', Caption='{s.Caption}', Parent='{s.Parent.Name}'")
        for pt in sc.PivotTables:
            print(f"    Connected Pivot: '{pt.Name}' on '{pt.Parent.Name}'")

    print("\n=== REVENUE / ARR METRIC INVESTIGATION ===")
    # Calculate At-Risk MRR and ARR from Fact_Churn
    ws_ch = wb.Worksheets("04_CustomerRetention_Cockpit")
    # Let's evaluate CUBEVALUE or check values
    mrr_val = ws_ch.Range("AC65").Value
    print(f"AC65 on 04_CustomerRetention_Cockpit: {mrr_val}")
    
    wb.Close(False)
except Exception as e:
    print(f"Error: {e}")
finally:
    excel.Quit()
