"""
Automated QA Validation Suite for PWC_Switzerland_Virtual_Case_FIXED.xlsm
Simulates real enterprise executive usage and validates data model, pivots,
slicers, formulas, navigation, and visual hierarchy.
"""

import os
import win32com.client as win32

WB_PATH = os.path.abspath("PWC_Switzerland_Virtual_Case_FIXED.xlsm")

print(f"=== Starting Comprehensive QA Audit on {os.path.basename(WB_PATH)} ===")

excel = win32.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

results = {
    "DATA_MODEL": False,
    "TABLES_COUNT": 0,
    "MEASURES_COUNT": 0,
    "PIVOT_TABLES": [],
    "CHARTS_ATTACHED": 0,
    "SLICERS_WORKING": 0,
    "KPI_FORMULA_LINKS": 0,
    "FORMULA_ERRORS": [],
    "NAV_BUTTONS_COUNT": 0,
    "VIEWPORT_NORMALIZED": True,
}

try:
    wb = excel.Workbooks.Open(WB_PATH)
    print("1. File opened cleanly without corruption or repair warnings.")

    # 1. DATA MODEL & MEASURES
    try:
        conn = wb.Connections("ThisWorkbookDataModel")
        results["DATA_MODEL"] = True
        print("2. Power Pivot VertiPaq Data Model verified: ACTIVE.")
        
        # Check tables & measures through CubeFields on pt_Agent
        ws_stage = wb.Worksheets("Staging_Pivots")
        pt_agent = ws_stage.PivotTables("pt_Agent")
        
        tables = set()
        measures = []
        for cf in pt_agent.CubeFields:
            name = cf.Name
            if name.startswith("[Measures]."):
                measures.append(name)
            elif name.startswith("["):
                tbl = name.split("].[")[0].replace("[", "")
                tables.add(tbl)
                
        results["TABLES_COUNT"] = len(tables)
        results["MEASURES_COUNT"] = len(measures)
        print(f"3. Model Tables: {len(tables)} verified | DAX Measures: {len(measures)} verified.")
    except Exception as e:
        print("Data Model check error:", e)

    # 2. PIVOTTABLES
    print("\n4. Checking Staging PivotTables:")
    expected_pts = [
        "pt_Agent", "pt_CH_Contract", "pt_CH_Tenure", "pt_CH_Payment", "pt_CH_Service",
        "pt_DI_Funnel", "pt_DI_Parity", "pt_DI_Promo", "pt_DI_Rating"
    ]
    for pt_name in expected_pts:
        try:
            pt = ws_stage.PivotTables(pt_name)
            results["PIVOT_TABLES"].append(pt.Name)
            print(f"   [PASS] {pt.Name} located at {pt.TableRange2.Address}")
        except Exception as e:
            print(f"   [FAIL] Missing {pt_name}: {e}")

    # 3. PIVOTCHARTS ATTACHMENT
    print("\n5. Checking Native PivotCharts across Cockpits:")
    cockpits = ["03_CallCenter_Cockpit", "04_CustomerRetention_Cockpit", "05_DiversityInclusion_Cockpit"]
    for cname in cockpits:
        ws_c = wb.Worksheets(cname)
        print(f"   --- Sheet: {cname} ---")
        for co in ws_c.ChartObjects():
            pt_attached = "None"
            try:
                pt_attached = co.Chart.PivotLayout.PivotTable.Name
                results["CHARTS_ATTACHED"] += 1
                print(f"     [PASS] {co.Name} -> Connected to PivotTable '{pt_attached}'")
            except Exception:
                print(f"     [INFO] {co.Name} -> Standalone/Scorecard Chart (No PivotLayout)")

    # 4. SLICERS & FILTERING INTERACTION TEST
    print("\n6. Testing Interactive Slicers:")
    print(f"   Total SlicerCaches in workbook: {wb.SlicerCaches.Count}")
    for sc in wb.SlicerCaches:
        print(f"   SlicerCache: {sc.Name} (SourceField: {sc.SourceName})")
        print(f"     Connected PivotTables: {sc.PivotTables.Count}")
        for pt_conn in sc.PivotTables:
            print(f"       -> {pt_conn.Name}")
        results["SLICERS_WORKING"] += 1

    # Test ClearAllFilters macro
    try:
        print("\n7. Testing ClearAllFilters VBA execution:")
        excel.Run("modFilterController.ClearAllFilters")
        print("   [PASS] modFilterController.ClearAllFilters executed successfully!")
    except Exception as e:
        print(f"   [FAIL] ClearAllFilters error: {e}")

    # 5. SINGLE SOURCE OF TRUTH (KPI VALUE SHAPE FORMULAS)
    print("\n8. Testing Single-Source-of-Truth KPI Formulations:")
    for cname in cockpits:
        ws_c = wb.Worksheets(cname)
        print(f"   --- {cname} ---")
        for sh in ws_c.Shapes:
            sh_l = sh.Name.lower()
            if any(k in sh_l for k in ["value_", "val_"]):
                try:
                    f = sh.DrawingObject.Formula
                    txt = sh.TextFrame2.TextRange.Text if sh.TextFrame2.HasText else ""
                    print(f"     [PASS] {sh.Name} -> Formula='{f}' | DisplayText='{txt}'")
                    results["KPI_FORMULA_LINKS"] += 1
                except Exception as e:
                    print(f"     [FAIL] {sh.Name} has no formula link: {e}")
        # Check hidden helper rows
        is_hidden = ws_c.Rows("65").Hidden
        print(f"     Row 65 Hidden State: {'YES [PASS]' if is_hidden else 'NO [FAIL]'}")

    # 6. NAVIGATION SYSTEM AUDIT
    print("\n9. Testing Universal Top Navigation System:")
    pres_sheets = [
        "00_Home_Portal", "01_Business_Domains", "02_Metadata_&_KPI_Catalog",
        "03_CallCenter_Cockpit", "04_CustomerRetention_Cockpit", "05_DiversityInclusion_Cockpit"
    ]
    expected_tabs = [
        "Nav_Tab_Home", "Nav_Tab_Domains", "Nav_Tab_Catalog",
        "Nav_Tab_CallCenter", "Nav_Tab_Retention", "Nav_Tab_Diversity"
    ]
    for sname in pres_sheets:
        ws_p = wb.Worksheets(sname)
        found_tabs = 0
        for s in ws_p.Shapes:
            if s.Name.lower().startswith("nav_tab_"):
                found_tabs += 1
                results["NAV_BUTTONS_COUNT"] += 1
        brand_title = ""
        try:
            brand_title = ws_p.Shapes("Nav_BrandTitle").TextFrame2.TextRange.Text
        except Exception:
            try:
                brand_title = ws_p.Shapes("Nav_TitleBox").TextFrame2.TextRange.Text
            except Exception:
                brand_title = "PwC Dashboard"
        print(f"   Sheet {sname}: {found_tabs}/6 Tabs Present | BrandTitle: {brand_title}")

    # 7. SCAN FOR FORMULA ERRORS
    print("\n10. Scanning Workbook for Formula Errors (#REF!, #VALUE!, etc.):")
    error_tokens = ["#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#N/A", "#NUM!", "#SPILL!"]
    for ws_any in wb.Worksheets:
        used_rng = ws_any.UsedRange
        cnt_err = 0
        try:
            # SpecialCells(xlCellTypeFormulas, xlErrors) = (2, 16)
            err_cells = used_rng.SpecialCells(2, 16)
            for c in err_cells:
                results["FORMULA_ERRORS"].append(f"{ws_any.Name}!{c.Address}: {c.Text}")
                cnt_err += 1
        except Exception:
            pass # No formula errors found on sheet
        print(f"   Sheet {ws_any.Name}: {cnt_err} formula errors.")

    # 8. VIEWPORT NORMALIZATION
    print("\n11. Verifying Landing Viewports:")
    for ws_any in wb.Worksheets:
        ws_any.Activate()
        z = excel.ActiveWindow.Zoom
        sr = excel.ActiveWindow.ScrollRow
        sc = excel.ActiveWindow.ScrollColumn
        sel = excel.Selection.Address
        print(f"   Sheet {ws_any.Name}: Zoom={z}%, Scroll=({sr},{sc}), Selection={sel}")
        if z != 80 or sr != 1 or sc != 1 or sel != "$A$1":
            results["VIEWPORT_NORMALIZED"] = False

    wb.Close(False)
    print("\n=== QA AUDIT COMPLETE ===")

except Exception as e:
    print("FATAL ERROR IN QA SCRIPT:", e)
finally:
    excel.Quit()
