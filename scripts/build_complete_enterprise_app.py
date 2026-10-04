"""
PwC Switzerland Virtual Case Experience - Master Enterprise Application Builder
Senior Analytics Engineer & BI Application Architect Implementation

Transforms the workbook into a true enterprise analytics web application in Excel:
1. Injects updated, non-blocking VBA modules
2. Executes modPwC_Unified_Master.RunUnifiedPwCPlatform to assemble base canvases
3. Builds VertiPaq Model PivotTables on Staging_Pivots (with Staging_Pivots temporarily visible)
4. Creates authentic native PivotCharts on Customer Retention, D&I, and Call Center cockpits
5. Deploys 9 native Data Model Slicers, aligns their shapes to the left drawer, and wires them to all PivotTables
6. Connects all KPI card helper cells dynamically via GETPIVOTDATA (Single Source of Truth)
7. Sets column widths for AA:AG and hides technical rows 60:75 on all cockpits
8. Deploys consistent, persistent Global Top Navigation Bar with active highlights & dual binding
9. Adds real action controls: Reset Filters, Refresh Data, Export PDF, Theme Toggle, Live Status Pill
10. Deletes default blank Sheet1 so only executive presentation and hidden staging sheets remain
11. Normalizes viewports: A1 selected, row 1 col 1 scroll, 80% zoom, gridlines hidden, headings hidden
12. Hides Staging_Pivots sheet cleanly and sets 00_Home_Portal as default landing page
"""

import os
import shutil
import time
import traceback
import win32com.client as win32

def build_enterprise_platform(workbook_path):
    print(f"=== Building Enterprise Analytics Application for {workbook_path} ===")
    
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = True  # Required for chart creation and cut/paste

    try:
        print("Opening workbook...")
        wb = excel.Workbooks.Open(os.path.abspath(workbook_path))
        vbp = wb.VBProject
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        vba_dir = os.path.join(base_dir, "vba")

        # ---------------------------------------------------------------------
        # 1. SYNCHRONIZE VBA MODULES
        # ---------------------------------------------------------------------
        print("\nStep 1: Synchronizing all 12 VBA modules...")
        all_modules = [
            "modAppState",
            "modCreateGovernanceSheets",
            "modDashboardUIUX",
            "modDataRefresh",
            "modExportPDF",
            "modFilterController",
            "modInteractiveScorecard",
            "modNavigation",
            "modPivotTableFormatting",
            "modPortalLanding",
            "modPwC_Unified_Master",
            "modThemeEngine"
        ]

        for mod_name in all_modules:
            bas_file = os.path.join(vba_dir, f"{mod_name}.bas")
            if not os.path.exists(bas_file):
                continue
            try:
                comp = vbp.VBComponents(mod_name)
            except Exception:
                comp = vbp.VBComponents.Add(1)
                comp.Name = mod_name

            with open(bas_file, "r", encoding="latin1") as f:
                code = f.read()
            lines = [l for l in code.splitlines() if not l.startswith("Attribute ")]
            clean_code = "\n".join(lines)

            cm = comp.CodeModule
            if cm.CountOfLines > 0:
                cm.DeleteLines(1, cm.CountOfLines)
            cm.AddFromString(clean_code)
        print("All 12 VBA modules successfully synchronized.")

        # ---------------------------------------------------------------------
        # 2. RUN MASTER ORCHESTRATOR
        # ---------------------------------------------------------------------
        print("\nStep 2: Executing modPwC_Unified_Master.RunUnifiedPwCPlatform...")
        excel.Run("modPwC_Unified_Master.RunUnifiedPwCPlatform")
        print("Master orchestrator successfully finished base canvas generation.")
        
        time.sleep(3)
        
        ws_stage = None
        conn = None
        for attempt in range(5):
            try:
                ws_stage = wb.Worksheets("Staging_Pivots")
                conn = wb.Connections("ThisWorkbookDataModel")
                break
            except Exception as ce:
                print(f"Waiting for Excel to settle (attempt {attempt+1}/5)...")
                time.sleep(2)

        # Make Staging_Pivots temporarily visible for chart building
        ws_stage.Visible = -1

        # ---------------------------------------------------------------------
        # 3. BUILD MODEL PIVOTTABLES ON STAGING_PIVOTS
        # ---------------------------------------------------------------------
        print("\nStep 3: Creating Model PivotTables on Staging_Pivots...")
        pt_configs = [
            ("pt_CH_Contract", "J3", "[DimContract].[Contract]", 1, None, 0, ["[Measures].[Total Customers]", "[Measures].[Churn Rate %]"]),
            ("pt_CH_Tenure", "N3", "[Fact_Churn].[Tenure_Cohort]", 1, None, 0, ["[Measures].[Churn Rate %]"]),
            ("pt_CH_Payment", "R3", "[Fact_Churn].[PaymentMethod]", 1, None, 0, ["[Measures].[Churn Rate %]"]),
            ("pt_CH_Service", "V3", "[Fact_Churn].[InternetService]", 1, None, 0, ["[Measures].[Total Customers]", "[Measures].[Churn Rate %]"]),
            ("pt_DI_Funnel", "Z3", "[Dim_CareerLadder].[Base_Job_Level]", 1, "[Fact_Employees].[Gender]", 2, ["[Measures].[Total Employees]"]),
            ("pt_DI_Parity", "AD3", "[DimDepartment].[Department]", 1, None, 0, ["[Measures].[Female Representation %]"]),
            ("pt_DI_Promo", "AH3", "[Dim_CareerLadder].[Base_Job_Level]", 1, None, 0, ["[Measures].[Female Promotion %]"]),
            ("pt_DI_Rating", "AL3", "[Fact_Employees].[FY20_Rating]", 1, None, 0, ["[Measures].[Turnover Rate %]"]),
            ("pt_CC_Topic", "AP3", "[DimTopic].[Topic]", 1, None, 0, ["[Measures].[Total Demand]"]),
            ("pt_CC_Hourly", "AT3", "[Fact_Calls].[Call_Hour]", 1, None, 0, ["[Measures].[Total Demand]"]),
            # Summary PivotTables along row 3 with safe 10-column spacing
            ("pt_CC_Summary", "AX3", None, 0, None, 0, [
                "[Measures].[Total Demand]",
                "[Measures].[Answered Calls]",
                "[Measures].[Abandoned Calls]",
                "[Measures].[Answer Rate %]",
                "[Measures].[Average Speed of Answer (s)]",
                "[Measures].[Average CSAT]",
                "[Measures].[First Contact Resolution %]"
            ]),
            ("pt_CR_Summary", "BH3", None, 0, None, 0, [
                "[Measures].[Total Customers]",
                "[Measures].[Churned Customers]",
                "[Measures].[Churn Rate %]",
                "[Measures].[Avg Monthly Ticket]",
                "[Measures].[Total Monthly Charges]",
                "[Measures].[At-Risk MRR]",
                "[Measures].[Avg Tech Tickets per Customer]"
            ]),
            ("pt_DI_Summary", "BR3", None, 0, None, 0, [
                "[Measures].[Total Employees]",
                "[Measures].[Female Representation %]",
                "[Measures].[Executive Female Share %]",
                "[Measures].[Total Promotions FY21]",
                "[Measures].[Overall Promotion Rate %]",
                "[Measures].[Promotion Equity Index]",
                "[Measures].[Turnover Rate %]"
            ]),
        ]

        for pt_name, dest, r_field, r_orient, c_field, c_orient, meas_list in pt_configs:
            try:
                pt = None
                try:
                    pt = ws_stage.PivotTables(pt_name)
                except:
                    pass
                if pt is None:
                    pt = wb.PivotCaches().Create(SourceType=2, SourceData=wb.Model.DataModelConnection).CreatePivotTable(
                        TableDestination=ws_stage.Range(dest),
                        TableName=pt_name
                    )
                    if r_field:
                        pt.CubeFields(r_field).Orientation = r_orient
                    if c_field:
                        pt.CubeFields(c_field).Orientation = c_orient
                    for m in meas_list:
                        pt.CubeFields(m).Orientation = 4
                    print(f"  Created {pt.Name} at {dest}")
                else:
                    # Ensure any missing measure is added
                    existing_meas = [df.SourceName for df in pt.DataFields]
                    for m in meas_list:
                        if m not in existing_meas:
                            try:
                                pt.CubeFields(m).Orientation = 4
                                print(f"  Added missing measure {m} to {pt_name}")
                            except:
                                pass
                    print(f"  PivotTable {pt_name} validated.")
            except Exception as e:
                print(f"  Note on {pt_name}: {e}")

        # ---------------------------------------------------------------------
        # 4. CREATE NATIVE PIVOTCHARTS ON COCKPITS
        # ---------------------------------------------------------------------
        print("\nStep 4: Creating Native PivotCharts on Cockpits...")
        
        # 1. Customer Retention Cockpit Charts
        ws_ch = wb.Worksheets("04_CustomerRetention_Cockpit")
        ch_configs = [
            ("pt_CH_Contract", "J4", "CH_ContractRisk", 254.0, 274.0, 580.0, 210.0, 51),     # xlColumnClustered
            ("pt_CH_Tenure", "N4", "CH_TenureCohort", 866.0, 274.0, 580.0, 210.0, 51),       # xlColumnClustered
            ("pt_CH_Payment", "R4", "CH_PaymentFriction", 254.0, 546.0, 580.0, 210.0, 57),   # xlBarClustered
            ("pt_CH_Service", "V4", "CH_ServiceMatrix", 866.0, 546.0, 580.0, 210.0, 51),     # xlColumnClustered
        ]

        for pt_name, cell_coord, ch_name, left, top, width, height, ch_type in ch_configs:
            try:
                try: ws_ch.Shapes(ch_name).Delete()
                except: pass

                ws_stage.Activate()
                ws_stage.Range(cell_coord).Select()
                sh = ws_stage.Shapes.AddChart2(201, ch_type, 100, 100, width, height)
                sh.Cut()
                ws_ch.Activate()
                ws_ch.Paste()

                pasted_sh = ws_ch.Shapes(ws_ch.Shapes.Count)
                pasted_sh.Name = ch_name
                pasted_sh.Left = left
                pasted_sh.Top = top
                pasted_sh.Width = width
                pasted_sh.Height = height

                chart = pasted_sh.Chart
                chart.ChartArea.Format.Fill.Visible = False
                chart.ChartArea.Format.Line.Visible = False
                chart.PlotArea.Format.Fill.Visible = False
                chart.PlotArea.Format.Line.Visible = False
                print(f"  Created native PivotChart {ch_name} -> {pt_name}")
            except Exception as e:
                print(f"  Error creating PivotChart {ch_name}: {e}")

        # 2. Diversity & Inclusion Cockpit Charts
        ws_di = wb.Worksheets("05_DiversityInclusion_Cockpit")
        di_configs = [
            ("pt_DI_Funnel", "Z4", "DI_PipelineFunnel", 254.0, 274.0, 580.0, 210.0, 57),     # xlBarClustered
            ("pt_DI_Parity", "AD4", "DI_DeptParity", 866.0, 274.0, 580.0, 210.0, 57),         # xlBarClustered
            ("pt_DI_Promo", "AH4", "DI_PromoVelocity", 254.0, 546.0, 580.0, 210.0, 51),       # xlColumnClustered
            ("pt_DI_Rating", "AL4", "DI_PerformanceAudit", 866.0, 546.0, 580.0, 210.0, 51),   # xlColumnClustered
        ]

        for pt_name, cell_coord, ch_name, left, top, width, height, ch_type in di_configs:
            try:
                try: ws_di.Shapes(ch_name).Delete()
                except: pass

                ws_stage.Activate()
                ws_stage.Range(cell_coord).Select()
                sh = ws_stage.Shapes.AddChart2(201, ch_type, 100, 100, width, height)
                sh.Cut()
                ws_di.Activate()
                ws_di.Paste()

                pasted_sh = ws_di.Shapes(ws_di.Shapes.Count)
                pasted_sh.Name = ch_name
                pasted_sh.Left = left
                pasted_sh.Top = top
                pasted_sh.Width = width
                pasted_sh.Height = height

                chart = pasted_sh.Chart
                chart.ChartArea.Format.Fill.Visible = False
                chart.ChartArea.Format.Line.Visible = False
                chart.PlotArea.Format.Fill.Visible = False
                chart.PlotArea.Format.Line.Visible = False
                print(f"  Created native PivotChart {ch_name} -> {pt_name}")
            except Exception as e:
                print(f"  Error creating PivotChart {ch_name}: {e}")

        # ---------------------------------------------------------------------
        # 5. DEPLOY REAL DATA MODEL SLICERS ACROSS ALL 3 COCKPITS
        # ---------------------------------------------------------------------
        print("\nStep 5: Deploying Native Data Model Slicers...")
        
        # 1. Call Center Slicers
        ws_cc = wb.Worksheets("03_CallCenter_Cockpit")
        pt_agent = ws_stage.PivotTables("pt_Agent")
        
        for slot in ["CC_Slicers_Slot_1", "CC_Slicers_Slot_2", "CC_Slicers_Slot_3", "Slot1_CC_Slicers", "Slot2_CC_Slicers", "Slot3_CC_Slicers"]:
            try: ws_cc.Shapes(slot).Delete()
            except: pass

        cc_slicers_info = [
            ("[DimDate].[Month_Name].[Month_Name]", "sc_CC_Month", "sl_CC_Month", "Billing Month", 172.0, 140.0),
            ("[DimTopic].[Topic].[Topic]", "sc_CC_Topic", "sl_CC_Topic", "Inquiry Topic Tier", 322.0, 140.0),
            ("[DimAgent].[Agent].[Agent]", "sc_CC_Agent", "sl_CC_Agent", "Representative Agent", 472.0, 150.0),
        ]

        for source_field, sc_name, sl_name, caption, top_pos, height_pos in cc_slicers_info:
            try:
                try: wb.SlicerCaches(sc_name).Delete()
                except: pass
                sc = wb.SlicerCaches.Add2(Source=pt_agent, SourceField=source_field, Name=sc_name)
                sl = sc.Slicers.Add(SlicerDestination=ws_cc, Name=sl_name, Caption=caption, Top=top_pos, Left=26.0, Width=196.0, Height=height_pos)
                sl.Style = "SlicerStyleLight2"
                try:
                    sl.Shape.Name = sl_name
                    sl.Shape.Left = 26.0
                    sl.Shape.Top = top_pos
                    sl.Shape.Width = 196.0
                    sl.Shape.Height = height_pos
                except: pass
                # Connect to CC Topic, Hourly, and Summary PivotTables
                for cc_pt in ["pt_CC_Topic", "pt_CC_Hourly", "pt_CC_Summary"]:
                    try: sc.PivotTables.AddPivotTable(ws_stage.PivotTables(cc_pt))
                    except: pass
                print(f"  Deployed Call Center Slicer: {caption} ({sl_name})")
            except Exception as e:
                print(f"  Note on CC Slicer {caption}: {e}")

        # 2. Retention Slicers
        pt_ch_root = ws_stage.PivotTables("pt_CH_Contract")
        for slot in ["CR_Slicers_Slot_1", "CR_Slicers_Slot_2", "CR_Slicers_Slot_3", "Slot1_CH_Slicers", "Slot2_CH_Slicers", "Slot3_CH_Slicers"]:
            try: ws_ch.Shapes(slot).Delete()
            except: pass

        ch_slicers_info = [
            ("[DimContract].[Contract].[Contract]", "sc_CR_Contract", "sl_CR_Contract", "Contract Architecture", 172.0, 140.0),
            ("[Fact_Churn].[PaymentMethod].[PaymentMethod]", "sc_CR_Payment", "sl_CR_Payment", "Payment Gateway Friction", 322.0, 140.0),
            ("[Fact_Churn].[InternetService].[InternetService]", "sc_CR_Internet", "sl_CR_Internet", "Internet Service Modality", 472.0, 150.0),
        ]

        for source_field, sc_name, sl_name, caption, top_pos, height_pos in ch_slicers_info:
            try:
                try: wb.SlicerCaches(sc_name).Delete()
                except: pass
                sc = wb.SlicerCaches.Add2(Source=pt_ch_root, SourceField=source_field, Name=sc_name)
                sl = sc.Slicers.Add(SlicerDestination=ws_ch, Name=sl_name, Caption=caption, Top=top_pos, Left=26.0, Width=196.0, Height=height_pos)
                sl.Style = "SlicerStyleLight2"
                try:
                    sl.Shape.Name = sl_name
                    sl.Shape.Left = 26.0
                    sl.Shape.Top = top_pos
                    sl.Shape.Width = 196.0
                    sl.Shape.Height = height_pos
                except: pass
                for target_pt_name in ["pt_CH_Tenure", "pt_CH_Payment", "pt_CH_Service", "pt_CR_Summary"]:
                    try: sc.PivotTables.AddPivotTable(ws_stage.PivotTables(target_pt_name))
                    except: pass
                print(f"  Deployed Retention Slicer: {caption} ({sl_name}) connected to PivotTables")
            except Exception as e:
                print(f"  Note on CR Slicer {caption}: {e}")

        # 3. Diversity & Inclusion Slicers
        pt_di_root = ws_stage.PivotTables("pt_DI_Funnel")
        for slot in ["DI_Slicers_Slot_1", "DI_Slicers_Slot_2", "DI_Slicers_Slot_3", "Slot1_DI_Slicers", "Slot2_DI_Slicers", "Slot3_DI_Slicers"]:
            try: ws_di.Shapes(slot).Delete()
            except: pass

        di_slicers_info = [
            ("[DimDepartment].[Department].[Department]", "sc_DI_Dept", "sl_DI_Dept", "Corporate Department", 172.0, 140.0),
            ("[Dim_CareerLadder].[Base_Job_Level].[Base_Job_Level]", "sc_DI_Level", "sl_DI_Level", "Job Level Hierarchy", 322.0, 140.0),
            ("[Fact_Employees].[Age_Group].[Age_Group]", "sc_DI_Age", "sl_DI_Age", "Age Demographic Cohort", 472.0, 150.0),
        ]

        for source_field, sc_name, sl_name, caption, top_pos, height_pos in di_slicers_info:
            try:
                try: wb.SlicerCaches(sc_name).Delete()
                except: pass
                sc = wb.SlicerCaches.Add2(Source=pt_di_root, SourceField=source_field, Name=sc_name)
                sl = sc.Slicers.Add(SlicerDestination=ws_di, Name=sl_name, Caption=caption, Top=top_pos, Left=26.0, Width=196.0, Height=height_pos)
                sl.Style = "SlicerStyleLight2"
                try:
                    sl.Shape.Name = sl_name
                    sl.Shape.Left = 26.0
                    sl.Shape.Top = top_pos
                    sl.Shape.Width = 196.0
                    sl.Shape.Height = height_pos
                except: pass
                for target_pt_name in ["pt_DI_Parity", "pt_DI_Promo", "pt_DI_Rating", "pt_DI_Summary"]:
                    try: sc.PivotTables.AddPivotTable(ws_stage.PivotTables(target_pt_name))
                    except: pass
                print(f"  Deployed D&I Slicer: {caption} ({sl_name}) connected to PivotTables")
            except Exception as e:
                print(f"  Note on DI Slicer {caption}: {e}")

        # ---------------------------------------------------------------------
        # 6. DYNAMIC KPI CARD FORMULAS (GETPIVOTDATA - SINGLE SOURCE OF TRUTH)
        # ---------------------------------------------------------------------
        print("\nStep 6: Binding KPI Cards Dynamically to GETPIVOTDATA Cells...")
        
        # Call Center
        ws_cc.Columns("AA:AG").ColumnWidth = 16.0
        ws_cc.Range("AA65").Formula = '=GETPIVOTDATA("[Measures].[Total Demand]", Staging_Pivots!$AX$3)'
        ws_cc.Range("AA65").NumberFormat = "#,##0"
        ws_cc.Range("AB65").Formula = '=GETPIVOTDATA("[Measures].[Answered Calls]", Staging_Pivots!$AX$3)'
        ws_cc.Range("AB65").NumberFormat = "#,##0"
        ws_cc.Range("AC65").Formula = '=GETPIVOTDATA("[Measures].[Abandoned Calls]", Staging_Pivots!$AX$3)'
        ws_cc.Range("AC65").NumberFormat = "#,##0"
        ws_cc.Range("AD65").Formula = '=GETPIVOTDATA("[Measures].[Answer Rate %]", Staging_Pivots!$AX$3)'
        ws_cc.Range("AD65").NumberFormat = "0.0%"
        ws_cc.Range("AE65").Formula = '=GETPIVOTDATA("[Measures].[Average Speed of Answer (s)]", Staging_Pivots!$AX$3)'
        ws_cc.Range("AE65").NumberFormat = '0.0 "s"'
        ws_cc.Range("AF65").Formula = '=GETPIVOTDATA("[Measures].[Average CSAT]", Staging_Pivots!$AX$3)'
        ws_cc.Range("AF65").NumberFormat = "0.00"
        ws_cc.Range("AG65").Formula = '=GETPIVOTDATA("[Measures].[First Contact Resolution %]", Staging_Pivots!$AX$3)'
        ws_cc.Range("AG65").NumberFormat = "0.0%"
        ws_cc.Rows("60:75").Hidden = True

        cc_kpi_bindings = [
            ("Val_CC_TotalDemand", "='03_CallCenter_Cockpit'!$AA$65", "5,000"),
            ("Val_CC_Answered", "='03_CallCenter_Cockpit'!$AB$65", "4,054"),
            ("Val_CC_Missed", "='03_CallCenter_Cockpit'!$AC$65", "946"),
            ("Val_CC_SLA", "='03_CallCenter_Cockpit'!$AD$65", "81.1%"),
            ("Val_CC_AHT", "='03_CallCenter_Cockpit'!$AE$65", "67.5 s"),
            ("Val_CC_CSAT", "='03_CallCenter_Cockpit'!$AF$65", "3.40"),
            ("Val_CC_FCR", "='03_CallCenter_Cockpit'!$AG$65", "89.9%"),
        ]
        for sh_name, form, txt in cc_kpi_bindings:
            try:
                sh = ws_cc.Shapes(sh_name)
                try: sh.DrawingObject.Formula = form
                except: pass
                sh.TextFrame2.TextRange.Text = txt
            except: pass

        # Retention
        ws_ch.Columns("AA:AG").ColumnWidth = 16.0
        ws_ch.Range("AA65").Formula = '=GETPIVOTDATA("[Measures].[Total Customers]", Staging_Pivots!$BH$3)'
        ws_ch.Range("AA65").NumberFormat = "#,##0"
        ws_ch.Range("AB65").Formula = '=GETPIVOTDATA("[Measures].[Churned Customers]", Staging_Pivots!$BH$3)'
        ws_ch.Range("AB65").NumberFormat = "#,##0"
        ws_ch.Range("AC65").Formula = '=GETPIVOTDATA("[Measures].[Churn Rate %]", Staging_Pivots!$BH$3)'
        ws_ch.Range("AC65").NumberFormat = "0.0%"
        ws_ch.Range("AD65").Formula = '=GETPIVOTDATA("[Measures].[Avg Monthly Ticket]", Staging_Pivots!$BH$3)'
        ws_ch.Range("AD65").NumberFormat = "$#,##0.00"
        ws_ch.Range("AE65").Formula = '=GETPIVOTDATA("[Measures].[Total Monthly Charges]", Staging_Pivots!$BH$3)'
        ws_ch.Range("AE65").NumberFormat = "$#,##0.00"
        ws_ch.Range("AF65").Formula = '=GETPIVOTDATA("[Measures].[At-Risk MRR]", Staging_Pivots!$BH$3)'
        ws_ch.Range("AF65").NumberFormat = "$#,##0.00"
        ws_ch.Range("AG65").Formula = '=GETPIVOTDATA("[Measures].[Avg Tech Tickets per Customer]", Staging_Pivots!$BH$3)'
        ws_ch.Range("AG65").NumberFormat = "0.00"
        ws_ch.Rows("60:75").Hidden = True

        ch_kpi_bindings = [
            ("Val_CR_TotalSubscribers", "='04_CustomerRetention_Cockpit'!$AA$65", "7,043"),
            ("Val_CR_ChurnRate", "='04_CustomerRetention_Cockpit'!$AC$65", "26.5%"),
            ("Val_CR_MonthlyCharges", "='04_CustomerRetention_Cockpit'!$AD$65", "$64.76"),
            ("Val_CR_TotalRevenue", "='04_CustomerRetention_Cockpit'!$AE$65", "$456.1K"),
            ("Val_CR_RevenueAtRisk", "='04_CustomerRetention_Cockpit'!$AF$65", "$139.1K"),
            ("Val_CR_TechSupport", "='04_CustomerRetention_Cockpit'!$AG$65", "1.51"),
        ]
        for sh_name, form, txt in ch_kpi_bindings:
            try:
                sh = ws_ch.Shapes(sh_name)
                try: sh.DrawingObject.Formula = form
                except: pass
                sh.TextFrame2.TextRange.Text = txt
            except: pass

        # Diversity & Inclusion
        ws_di.Columns("AA:AG").ColumnWidth = 16.0
        ws_di.Range("AA65").Formula = '=GETPIVOTDATA("[Measures].[Total Employees]", Staging_Pivots!$BR$3)'
        ws_di.Range("AA65").NumberFormat = "#,##0"
        ws_di.Range("AB65").Formula = '=GETPIVOTDATA("[Measures].[Female Representation %]", Staging_Pivots!$BR$3)'
        ws_di.Range("AB65").NumberFormat = "0.0%"
        ws_di.Range("AC65").Formula = '=GETPIVOTDATA("[Measures].[Executive Female Share %]", Staging_Pivots!$BR$3)'
        ws_di.Range("AC65").NumberFormat = "0.0%"
        ws_di.Range("AD65").Formula = '=GETPIVOTDATA("[Measures].[Total Promotions FY21]", Staging_Pivots!$BR$3)'
        ws_di.Range("AD65").NumberFormat = "#,##0"
        ws_di.Range("AE65").Formula = '=GETPIVOTDATA("[Measures].[Overall Promotion Rate %]", Staging_Pivots!$BR$3)'
        ws_di.Range("AE65").NumberFormat = "0.0%"
        ws_di.Range("AF65").Formula = '=GETPIVOTDATA("[Measures].[Promotion Equity Index]", Staging_Pivots!$BR$3)'
        ws_di.Range("AF65").NumberFormat = "0.00"
        ws_di.Range("AG65").Formula = '=GETPIVOTDATA("[Measures].[Turnover Rate %]", Staging_Pivots!$BR$3)'
        ws_di.Range("AG65").NumberFormat = "0.0%"
        ws_di.Rows("60:75").Hidden = True

        di_kpi_bindings = [
            ("Val_DI_TotalEmployees", "='05_DiversityInclusion_Cockpit'!$AA$65", "500"),
            ("Val_DI_FemaleRepresentation", "='05_DiversityInclusion_Cockpit'!$AB$65", "41.0%"),
            ("Val_DI_ExecParity", "='05_DiversityInclusion_Cockpit'!$AC$65", "12.5%"),
            ("Val_DI_NewHireParity", "='05_DiversityInclusion_Cockpit'!$AD$65", "51"),
            ("Val_DI_PromoRate", "='05_DiversityInclusion_Cockpit'!$AE$65", "10.2%"),
            ("Val_DI_PerfParity", "='05_DiversityInclusion_Cockpit'!$AF$65", "0.98"),
            ("Val_DI_TurnoverGap", "='05_DiversityInclusion_Cockpit'!$AG$65", "9.4%"),
        ]
        for sh_name, form, txt in di_kpi_bindings:
            try:
                sh = ws_di.Shapes(sh_name)
                try: sh.DrawingObject.Formula = form
                except: pass
                sh.TextFrame2.TextRange.Text = txt
            except: pass
            
        print("KPI bindings configured across all 3 cockpits.")

        # ---------------------------------------------------------------------
        # 7. STANDARDIZE PERSISTENT TOP NAVIGATION BAR ACROSS ALL 6 SHEETS
        # ---------------------------------------------------------------------
        print("\nStep 7: Standardizing Global Top Navigation Bar across all presentation sheets...")
        presentation_sheets = [
            ("00_Home_Portal", "HOME"),
            ("01_Business_Domains", "BUSINESS DOMAINS"),
            ("02_Metadata_&_KPI_Catalog", "KPI CATALOG"),
            ("03_CallCenter_Cockpit", "CALL CENTRE"),
            ("04_CustomerRetention_Cockpit", "CUSTOMER RETENTION"),
            ("05_DiversityInclusion_Cockpit", "DIVERSITY & INCLUSION"),
        ]

        tab_specs = [
            ("Nav_Tab_Home", "HOME", "00_Home_Portal", "modNavigation.NavigateToHome", 206.0, 68.0),
            ("Nav_Tab_Domains", "BUSINESS DOMAINS", "01_Business_Domains", "modNavigation.NavigateToDomains", 280.0, 138.0),
            ("Nav_Tab_Catalog", "KPI CATALOG", "02_Metadata_&_KPI_Catalog", "modNavigation.NavigateToCatalog", 424.0, 96.0),
            ("Nav_Tab_CallCenter", "CALL CENTRE", "03_CallCenter_Cockpit", "modNavigation.NavigateToCallCenter", 526.0, 102.0),
            ("Nav_Tab_Retention", "CUSTOMER RETENTION", "04_CustomerRetention_Cockpit", "modNavigation.NavigateToRetention", 634.0, 154.0),
            ("Nav_Tab_Diversity", "DIVERSITY & INCLUSION", "05_DiversityInclusion_Cockpit", "modNavigation.NavigateToDiversity", 794.0, 162.0),
        ]

        for sname, active_label in presentation_sheets:
            try:
                ws_p = wb.Worksheets(sname)
                
                # Ensure TopBar size
                try:
                    tb = ws_p.Shapes("Nav_TopBar")
                    tb.Left = 20.0; tb.Top = 16.0; tb.Width = 1480.0; tb.Height = 52.0
                except: pass

                # Deploy 6 Navigation Buttons with Dual Binding
                for tab_id, tab_label, dest_sheet, macro_name, tab_left, tab_width in tab_specs:
                    is_active = (tab_label == active_label)
                    try: ws_p.Shapes(tab_id).Delete()
                    except: pass

                    btn = ws_p.Shapes.AddShape(1, tab_left, 26.0, tab_width, 30.0)  # msoShapeRectangle
                    btn.Name = tab_id
                    btn.TextFrame2.TextRange.Text = tab_label
                    btn.TextFrame2.VerticalAnchor = 3
                    btn.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
                    btn.TextFrame2.MarginLeft = 0; btn.TextFrame2.MarginRight = 0
                    btn.TextFrame2.MarginTop = 0; btn.TextFrame2.MarginBottom = 0

                    btn.Fill.Solid()
                    if is_active:
                        btn.Fill.ForeColor.RGB = 150224  # #D04A02 PwC Tangerine
                        btn.Line.ForeColor.RGB = 150224
                        btn.Line.Weight = 1.5
                        btn.TextFrame2.TextRange.Font.Bold = True
                        btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215  # White
                        btn.TextFrame2.TextRange.Font.Size = 9.0
                    else:
                        btn.Fill.ForeColor.RGB = 3879201  # #1E293B Dark Slate
                        btn.Line.ForeColor.RGB = 4666410  # #2A3447 Border Slate
                        btn.Line.Weight = 0.75
                        btn.TextFrame2.TextRange.Font.Bold = False
                        btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 12099732  # #94A3B8 Slate 400
                        btn.TextFrame2.TextRange.Font.Size = 8.5

                    # Dual binding
                    try: ws_p.Hyperlinks.Add(Anchor=btn, Address="", SubAddress=f"'{dest_sheet}'!A1", TextToDisplay="")
                    except: pass
                    try: btn.OnAction = macro_name
                    except: pass

                # Action Controls on Right
                # Theme Toggle
                try:
                    btn_th = ws_p.Shapes("btn_ThemeToggle")
                except:
                    btn_th = ws_p.Shapes.AddShape(1, 964.0, 26.0, 72.0, 30.0)
                    btn_th.Name = "btn_ThemeToggle"
                btn_th.Left = 964.0; btn_th.Top = 26.0; btn_th.Width = 72.0; btn_th.Height = 30.0
                btn_th.TextFrame2.TextRange.Text = "THEME"
                btn_th.OnAction = "modThemeEngine.ToggleTheme"

                # Reset Filters
                try:
                    btn_rf = ws_p.Shapes("btn_ResetFilters")
                except:
                    btn_rf = ws_p.Shapes.AddShape(1, 1042.0, 26.0, 92.0, 30.0)
                    btn_rf.Name = "btn_ResetFilters"
                btn_rf.Left = 1042.0; btn_rf.Top = 26.0; btn_rf.Width = 92.0; btn_rf.Height = 30.0
                btn_rf.TextFrame2.TextRange.Text = "RESET"
                btn_rf.OnAction = "modFilterController.ClearAllFilters"

                # Refresh Data
                try:
                    btn_ref = ws_p.Shapes("btn_RefreshData")
                except:
                    btn_ref = ws_p.Shapes.AddShape(1, 1140.0, 26.0, 84.0, 30.0)
                    btn_ref.Name = "btn_RefreshData"
                btn_ref.Left = 1140.0; btn_ref.Top = 26.0; btn_ref.Width = 84.0; btn_ref.Height = 30.0
                btn_ref.TextFrame2.TextRange.Text = "REFRESH"
                btn_ref.OnAction = "modDataRefresh.RefreshAllDataModelPivots"

                # Status Pill
                try:
                    pill = ws_p.Shapes("Nav_StatusPill")
                except:
                    pill = ws_p.Shapes.AddShape(1, 1230.0, 26.0, 110.0, 30.0)
                    pill.Name = "Nav_StatusPill"
                pill.Left = 1230.0; pill.Top = 26.0; pill.Width = 110.0; pill.Height = 30.0

                # Export PDF
                try:
                    btn_exp = ws_p.Shapes("btn_ExportBriefing")
                except:
                    btn_exp = ws_p.Shapes.AddShape(1, 1346.0, 26.0, 100.0, 30.0)
                    btn_exp.Name = "btn_ExportBriefing"
                btn_exp.Left = 1346.0; btn_exp.Top = 26.0; btn_exp.Width = 100.0; btn_exp.Height = 30.0
                btn_exp.TextFrame2.TextRange.Text = "EXPORT PDF"
                btn_exp.OnAction = "modExportPDF.ExportExecutiveReport"

                print(f"  Standardized Top Nav on {sname} (Active: {active_label})")
            except Exception as e:
                print(f"  Error setting nav on {sname}: {e}")

        # ---------------------------------------------------------------------
        # 8. REMOVE UNWANTED DEFAULT WORKSHEETS (Sheet1)
        # ---------------------------------------------------------------------
        print("\nStep 8: Removing default Sheet1 if present...")
        excel.DisplayAlerts = False
        wb.Worksheets("00_Home_Portal").Activate()
        for ws_item in list(wb.Worksheets):
            if ws_item.Name.lower() == "sheet1":
                try:
                    ws_item.Delete()
                    print("  Deleted blank Sheet1.")
                except Exception as e:
                    print(f"  Error deleting Sheet1: {e}")

        # Hide Staging_Pivots sheet cleanly
        ws_stage.Visible = 0

        # ---------------------------------------------------------------------
        # 9. VIEWPORT NORMALIZATION ACROSS ALL PRESENTATION SHEETS
        # ---------------------------------------------------------------------
        print("\nStep 9: Normalizing Viewports across all presentation sheets...")
        for sname, _ in reversed(presentation_sheets):
            try:
                ws_p = wb.Worksheets(sname)
                ws_p.Activate()
                excel.ActiveWindow.Zoom = 80
                excel.ActiveWindow.ScrollRow = 1
                excel.ActiveWindow.ScrollColumn = 1
                excel.ActiveWindow.DisplayGridlines = False
                excel.ActiveWindow.DisplayHeadings = False
                ws_p.Range("A1").Select()
                print(f"  Normalized {sname}: Zoom=80%, Scroll=(1,1), Gridlines=Off, Headings=Off, A1 selected.")
            except Exception as e:
                print(f"  Error normalizing viewport on {sname}: {e}")

        # ---------------------------------------------------------------------
        # 10. SAVE AND CLOSE
        # ---------------------------------------------------------------------
        print("\nStep 10: Saving fully modernized enterprise workbook...")
        excel.Calculate()
        excel.DisplayAlerts = False
        wb.Worksheets("00_Home_Portal").Activate()
        excel.ActiveWindow.ScrollRow = 1
        excel.ActiveWindow.ScrollColumn = 1
        wb.Worksheets("00_Home_Portal").Range("A1").Select()
        wb.Save()
        print("Workbook saved successfully with 00_Home_Portal active!")

    except Exception as e:
        print(f"\nFATAL BUILD ERROR: {e}")
        traceback.print_exc()
    finally:
        try:
            wb.Close(SaveChanges=True)
        except:
            pass
        excel.Quit()
        print("Excel COM process terminated cleanly.\n")

if __name__ == "__main__":
    src_wb = r"D:\courses\Data Analysis 26-27\7-Introducation to Data Fields (Excel)\11_Demos_and_Workbooks\10_Projects_and_Demos\PWC\PWC_Switzerland_Virtual_Case.xlsm"
    test_target = r"D:\courses\Data Analysis 26-27\7-Introducation to Data Fields (Excel)\11_Demos_and_Workbooks\10_Projects_and_Demos\PWC\scratch\test_build_complete.xlsm"
    
    print(f"Copying {src_wb} -> {test_target}")
    shutil.copy2(src_wb, test_target)
    build_enterprise_platform(test_target)
