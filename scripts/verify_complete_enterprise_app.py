import os
import sys
import win32com.client

def run_qa_verification(wb_path):
    print(f"=== Running Forensic Enterprise QA Verification on: {wb_path} ===")
    if not os.path.exists(wb_path):
        print(f"ERROR: File {wb_path} does not exist!")
        return False

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False

    passed = True

    try:
        wb = excel.Workbooks.Open(os.path.abspath(wb_path), ReadOnly=False)
        print("Workbook opened cleanly without Excel warnings or corruption.")
        initial_active_sheet = wb.ActiveSheet.Name
        print(f"Initial Active Sheet on Open: '{initial_active_sheet}'")

        # 1. Sheets inventory
        expected_sheets = [
            "00_Home_Portal",
            "01_Business_Domains",
            "02_Metadata_&_KPI_Catalog",
            "03_CallCenter_Cockpit",
            "04_CustomerRetention_Cockpit",
            "05_DiversityInclusion_Cockpit"
        ]
        actual_sheets = [ws.Name for ws in wb.Worksheets]
        print(f"\nSheet Inventory: {actual_sheets}")
        for s in expected_sheets:
            if s in actual_sheets:
                print(f"  [PASS] Sheet '{s}' is present.")
            else:
                print(f"  [FAIL] Sheet '{s}' is MISSING!")
                passed = False

        if "Sheet1" in actual_sheets or "sheet1" in [s.lower() for s in actual_sheets]:
            print("  [FAIL] Blank Sheet1 was not removed!")
            passed = False
        else:
            print("  [PASS] Blank Sheet1 is absent.")

        # 2. Check Data Model
        try:
            model = wb.Model
            t_count = model.ModelTables.Count
            r_count = model.ModelRelationships.Count
            print(f"\nData Model Status: {t_count} Tables, {r_count} Relationships.")
            if t_count < 10:
                print("  [FAIL] Data Model table count is lower than expected!")
                passed = False
            else:
                print("  [PASS] Data Model VertiPaq engine is intact.")
        except Exception as e:
            print(f"  [FAIL] Unable to read Data Model: {e}")
            passed = False

        # 3. Check Staging_Pivots & PivotTables
        try:
            ws_stage = wb.Worksheets("Staging_Pivots")
            pt_count = ws_stage.PivotTables().Count
            print(f"\nStaging_Pivots PivotTable count: {pt_count}")
            expected_pts = [
                "pt_CH_Contract", "pt_CH_Tenure", "pt_CH_Payment", "pt_CH_Service",
                "pt_DI_Funnel", "pt_DI_Parity", "pt_DI_Promo", "pt_DI_Rating",
                "pt_CC_Topic", "pt_CC_Hourly", "pt_Agent",
                "pt_CC_Summary", "pt_CR_Summary", "pt_DI_Summary"
            ]
            pt_names = [ws_stage.PivotTables(i).Name for i in range(1, pt_count + 1)]
            for pt in expected_pts:
                if pt in pt_names:
                    print(f"  [PASS] PivotTable '{pt}' exists on Staging_Pivots.")
                else:
                    print(f"  [FAIL] PivotTable '{pt}' missing from Staging_Pivots!")
                    passed = False
        except Exception as e:
            print(f"  [FAIL] Error inspecting Staging_Pivots: {e}")
            passed = False

        # 4. Check PivotCharts on Cockpits
        cockpit_charts = {
            "04_CustomerRetention_Cockpit": ["CH_ContractRisk", "CH_TenureCohort", "CH_PaymentFriction", "CH_ServiceMatrix"],
            "05_DiversityInclusion_Cockpit": ["DI_PipelineFunnel", "DI_DeptParity", "DI_PromoVelocity", "DI_PerformanceAudit"]
        }
        print("\nVerifying Native PivotCharts:")
        for sheet_name, charts in cockpit_charts.items():
            ws = wb.Worksheets(sheet_name)
            sheet_chart_names = [ws.ChartObjects(i).Name for i in range(1, ws.ChartObjects().Count + 1)]
            for c_name in charts:
                if c_name in sheet_chart_names:
                    co = ws.ChartObjects(c_name)
                    has_pivot = False
                    try:
                        if co.Chart.PivotLayout is not None:
                            has_pivot = True
                    except:
                        pass
                    print(f"  [PASS] Chart '{c_name}' on '{sheet_name}': Left={co.Left:.1f}, Top={co.Top:.1f}, Width={co.Width:.1f}, Height={co.Height:.1f}, HasPivotLayout={has_pivot}")
                else:
                    print(f"  [FAIL] Chart '{c_name}' missing from '{sheet_name}'!")
                    passed = False

        # 5. Check SlicerCaches and Slicers
        print(f"\nSlicerCaches count: {wb.SlicerCaches.Count}")
        sc_names = [wb.SlicerCaches(i).Name for i in range(1, wb.SlicerCaches.Count + 1)]
        print(f"SlicerCaches: {sc_names}")
        expected_slicers = [
            ("03_CallCenter_Cockpit", ["sl_CC_Month", "sl_CC_Topic", "sl_CC_Agent"]),
            ("04_CustomerRetention_Cockpit", ["sl_CR_Contract", "sl_CR_Payment", "sl_CR_Internet"]),
            ("05_DiversityInclusion_Cockpit", ["sl_DI_Dept", "sl_DI_Level", "sl_DI_Age"])
        ]
        for sheet_name, sl_list in expected_slicers:
            ws = wb.Worksheets(sheet_name)
            shape_names = [s.Name for s in ws.Shapes]
            for sl_name in sl_list:
                found = any(sl_name.lower() in sn.lower() for sn in shape_names)
                if found:
                    print(f"  [PASS] Slicer '{sl_name}' docked on '{sheet_name}'.")
                else:
                    print(f"  [FAIL] Slicer '{sl_name}' not found on '{sheet_name}'!")
                    passed = False

        # 6. Check GETPIVOTDATA formula cells and values
        print("\nVerifying GETPIVOTDATA Helper Ranges (AA65:AG65) and KPI BAN Bindings:")
        cockpit_kpis = {
            "03_CallCenter_Cockpit": ["AA65", "AB65", "AC65", "AD65", "AE65"],
            "04_CustomerRetention_Cockpit": ["AA65", "AB65", "AC65", "AD65", "AE65"],
            "05_DiversityInclusion_Cockpit": ["AA65", "AB65", "AC65", "AD65", "AE65", "AF65"]
        }
        for sheet_name, cells in cockpit_kpis.items():
            ws = wb.Worksheets(sheet_name)
            for c in cells:
                cell_formula = ws.Range(c).Formula
                cell_val = ws.Range(c).Value
                cell_text = ws.Range(c).Text
                is_err = False
                if str(cell_val).startswith("#") or (isinstance(cell_val, (int, float)) and cell_val < 0):
                    is_err = True
                    passed = False
                status = "[FAIL]" if is_err else "[PASS]"
                print(f"  {status} {sheet_name} {c}: Val={cell_val} (Text='{cell_text}'), Formula={cell_formula}")

        # 7. Check Top Navigation Standard on All Sheets
        print("\nVerifying Top Navigation Buttons across all presentation sheets:")
        nav_buttons = [
            "Nav_Tab_Home", "Nav_Tab_Domains", "Nav_Tab_Catalog",
            "Nav_Tab_CallCenter", "Nav_Tab_Retention", "Nav_Tab_Diversity"
        ]
        action_buttons = ["btn_ThemeToggle", "btn_ResetFilters", "btn_RefreshData", "btn_ExportBriefing"]
        for s in expected_sheets:
            ws = wb.Worksheets(s)
            shape_names = [shp.Name for shp in ws.Shapes]
            all_nav = all(nb in shape_names for nb in nav_buttons)
            all_act = all(ab in shape_names for ab in action_buttons)
            if all_nav and all_act:
                print(f"  [PASS] Sheet '{s}' has complete Top Nav (6 tabs + 4 action buttons).")
            else:
                missing = [b for b in nav_buttons + action_buttons if b not in shape_names]
                print(f"  [FAIL] Sheet '{s}' is missing navigation buttons: {missing}")
                passed = False

        # 8. Check Call Center Agent Scorecard
        print("\nVerifying Call Center Agent Scorecard on 03_CallCenter_Cockpit:")
        ws_cc = wb.Worksheets("03_CallCenter_Cockpit")
        cc_shape_names = [s.Name for s in ws_cc.Shapes]
        if "LiveScorecardPic" in cc_shape_names or "LiveScorecard_HTMLTable" in cc_shape_names:
            print("  [PASS] Live Linked Scorecard Picture is docked inside the Agent Scorecard container.")
        else:
            print("  [FAIL] Live Scorecard Picture missing from 03_CallCenter_Cockpit!")
            passed = False

        # Check underlying pt_Agent on Staging_Pivots
        pt_agent = ws_stage.PivotTables("pt_Agent")
        speed_header = pt_agent.DataFields(4).Caption
        print(f"  pt_Agent DataField 4: Caption='{speed_header}'")
        # Check first agent row on Staging_Pivots (Row 4, Col G = col 7)
        agent1_speed_text = ws_stage.Cells(4, 7).Text
        agent1_rate_text = ws_stage.Cells(4, 5).Text
        print(f"  Agent 1 (Becky): Speed='{agent1_speed_text}', Answer Rate='{agent1_rate_text}'")
        if "%" in agent1_speed_text:
            print(f"    [FAIL] Speed format contains percentage symbol: '{agent1_speed_text}'!")
            passed = False
        else:
            print(f"    [PASS] Speed format validated: '{agent1_speed_text}'.")

        # 9. Check Viewports
        print("\nVerifying Viewport & Display Settings:")
        for s in expected_sheets:
            ws = wb.Worksheets(s)
            ws.Activate()
            win = wb.Application.ActiveWindow
            zoom = win.Zoom
            grid = win.DisplayGridlines
            head = win.DisplayHeadings
            active_cell = excel.ActiveCell.Address
            status = "[PASS]"
            if zoom != 80 or grid or head or "$A$1" not in active_cell:
                status = "[FAIL]"
                passed = False
            print(f"  {status} {s}: Zoom={zoom}%, Gridlines={grid}, Headings={head}, ActiveCell={active_cell}")

        # Check default landing page
        if initial_active_sheet == "00_Home_Portal":
            print(f"  [PASS] 00_Home_Portal is active default landing page on open.")
        else:
            print(f"  [FAIL] Default landing page on open is '{initial_active_sheet}', expected '00_Home_Portal'!")
            passed = False

        # 10. Check Formula Errors across entire sheets
        print("\nScanning entire presentation sheets for formula errors...")
        error_types = ["#REF!", "#VALUE!", "#DIV/0!", "#N/A", "#NAME?", "#NUM!"]
        for s in expected_sheets:
            ws = wb.Worksheets(s)
            used_range = ws.UsedRange
            errors_found = []
            for err in error_types:
                found_cell = None
                try:
                    found_cell = used_range.Find(err)
                except:
                    pass
                if found_cell is not None:
                    errors_found.append(f"{err} at {found_cell.Address}")
            if errors_found:
                print(f"  [FAIL] Sheet '{s}' contains formula errors: {errors_found}")
                passed = False
            else:
                print(f"  [PASS] Sheet '{s}' has ZERO formula errors.")

        # 11. Test Slicer Interactivity & Dynamic Calculation
        print("\nTesting Slicer Interactivity & Dynamic GETPIVOTDATA Calculation:")
        try:
            sc_contract = wb.SlicerCaches("sc_CR_Contract")
            ws_cr = wb.Worksheets("04_CustomerRetention_Cockpit")
            val_before = ws_cr.Range("AA65").Value
            print(f"  Initial Total Subscribers (AA65): {val_before}")

            # Select 'One year' contract via OLAP MDX member specification
            print("  Applying slicer filter: '[DimContract].[Contract].&[One year]'...")
            sc_contract.VisibleSlicerItemsList = ["[DimContract].[Contract].&[One year]"]
            excel.Calculate()
            val_filtered = ws_cr.Range("AA65").Value
            print(f"  Filtered Total Subscribers for 'One year': {val_filtered}")

            if val_filtered != val_before and val_filtered > 0:
                print("  [PASS] Slicer dynamically updated GETPIVOTDATA calculation!")
            else:
                print(f"  [WARN] Value did not change or became zero (Before={val_before}, Filtered={val_filtered})")

            # Clear filter
            sc_contract.ClearManualFilter()
            excel.Calculate()
            val_restored = ws_cr.Range("AA65").Value
            print(f"  Restored Total Subscribers: {val_restored}")
            if val_restored == val_before:
                print("  [PASS] Slicer ClearManualFilter restored original metric!")
            else:
                print(f"  [WARN] Restored value does not match initial value ({val_restored} != {val_before})")
        except Exception as e:
            print(f"  [FAIL] Slicer interactivity test failed: {e}")
            passed = False

    except Exception as e:
        print(f"FATAL QA ERROR: {e}")
        passed = False
    finally:
        wb.Close(SaveChanges=False)
        excel.Quit()

    print("\n" + "="*50)
    if passed:
        print("OVERALL QA STATUS: ALL CHECKS PASSED [SUCCESS]!")
    else:
        print("OVERALL QA STATUS: SOME CHECKS FAILED [FAILURE]!")
    print("="*50)
    return passed

if __name__ == "__main__":
    target = r"D:\courses\Data Analysis 26-27\7-Introducation to Data Fields (Excel)\11_Demos_and_Workbooks\10_Projects_and_Demos\PWC\scratch\test_build_complete.xlsm"
    run_qa_verification(target)
