"""
PwC Switzerland Virtual Case Experience - Enterprise Automation Script
Applies all workbook updates:
1. Injects updated VBA modules (Theme Engine, Portal Landing, Navigation, Scorecard, UI/UX)
2. Builds executive Landing Page ("00_Home_Portal") with hero banner, KPI ticker, and interactive cards
3. Fixes Agent Scorecard formatting (Speed format 0.0 "s", hides AutoFilter dropdown, calibrates box fit)
4. Upgrades visual chart styling and converts dashed DockZones to solid modern borders
5. Connects Theme Switcher & Home Portal navigation buttons across all cockpits
"""

import os
import shutil
import win32com.client

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKBOOK_PATH = os.path.join(BASE_DIR, "PWC_Switzerland_Virtual_Case.xlsm")
BACKUP_PATH = os.path.join(BASE_DIR, "PWC_Switzerland_Virtual_Case_backup.xlsm")
VBA_DIR = os.path.join(BASE_DIR, "vba")

def backup_workbook():
    if not os.path.exists(BACKUP_PATH):
        print(f"Creating safety backup: {BACKUP_PATH}")
        shutil.copy2(WORKBOOK_PATH, BACKUP_PATH)

def update_vba_module(vbp, mod_name, file_path):
    try:
        comp = vbp.VBComponents(mod_name)
    except Exception:
        comp = vbp.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        comp.Name = mod_name

    with open(file_path, "r", encoding="latin1") as f:
        code = f.read()

    # Filter out Attribute lines because CodeModule manages attributes
    lines = [l for l in code.splitlines() if not l.startswith("Attribute ")]
    clean_code = "\n".join(lines)

    cm = comp.CodeModule
    if cm.CountOfLines > 0:
        cm.DeleteLines(1, cm.CountOfLines)
    cm.AddFromString(clean_code)
    print(f"Successfully injected code into component: {mod_name} ({cm.CountOfLines} lines)")

def apply_updates():
    backup_workbook()
    
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False

    try:
        print(f"Opening workbook: {WORKBOOK_PATH}")
        wb = excel.Workbooks.Open(WORKBOOK_PATH)
        vbp = wb.VBProject

        # 1. Update VBA Modules safely via CodeModule
        modules_to_update = [
            "modThemeEngine",
            "modPortalLanding",
            "modNavigation",
            "modInteractiveScorecard",
            "modPivotTableFormatting",
            "modDashboardUIUX"
        ]

        for mod_name in modules_to_update:
            bas_file = os.path.join(VBA_DIR, f"{mod_name}.bas")
            if not os.path.exists(bas_file):
                print(f"Warning: {bas_file} not found, skipping.")
                continue
            update_vba_module(vbp, mod_name, bas_file)

        # 2. Build Executive Portal Landing Page ("00_Home_Portal")
        print("Executing modPortalLanding.BuildExecutivePortal...")
        try:
            excel.Run("modPortalLanding.BuildExecutivePortal")
            print("Successfully built 00_Home_Portal!")
        except Exception as e:
            print(f"Error building 00_Home_Portal: {e}")

        # 3. Format Staging_Pivots & Dock Scorecard
        print("Calibrating Staging_Pivots and pt_Agent...")
        try:
            ws_staging = wb.Worksheets("Staging_Pivots")
            for pt in ws_staging.PivotTables():
                if pt.Name == "pt_Agent":
                    pt.DisplayFieldCaptions = False
                    pt.RowGrand = False
                    pt.ColumnGrand = False
                    
                    # Ensure Avg Speed field number format is 0.0 "s"
                    for df in pt.DataFields:
                        if "Speed" in df.Name or "Speed" in df.Caption:
                            df.NumberFormat = '0.0 "s"'
                            print(f"Enforced NumberFormat on {df.Name}: 0.0 \"s\"")
            
            # Format staging cells C3:H11 explicitly
            col_widths = {'C': 12.5, 'D': 10.5, 'E': 11.5, 'F': 10.5, 'G': 12.0, 'H': 10.5}
            for col_letter, width in col_widths.items():
                ws_staging.Columns(col_letter).ColumnWidth = width

            ws_staging.Rows(3).RowHeight = 24.0
            for r in range(4, 12):
                ws_staging.Rows(r).RowHeight = 22.5

            # Enforce cell number format on G4:G11
            ws_staging.Range("G4:G11").NumberFormat = '0.0 "s"'
            print("Configured Staging_Pivots geometry and speed format.")
        except Exception as e:
            print(f"Staging_Pivots error: {e}")

        # Run Scorecard Docker to sync linked picture and SVG icon
        print("Executing modInteractiveScorecard.BuildAndDockInteractiveScorecard...")
        try:
            excel.Run("modInteractiveScorecard.BuildAndDockInteractiveScorecard")
            print("Scorecard docked successfully!")
        except Exception as e:
            print(f"Scorecard dock error: {e}")

        # 4. Modernize Chart Styling & DockZones across Cockpits
        cockpits = ["03_CallCenter_Cockpit", "04_CustomerRetention_Cockpit", "05_DiversityInclusion_Cockpit"]
        for sheet_name in cockpits:
            try:
                ws = wb.Worksheets(sheet_name)
                print(f"Modernizing {sheet_name}...")

                # Convert all DockZones to solid #E2E8F0 borders
                for shp in ws.Shapes:
                    if "DockZone_" in shp.Name:
                        shp.Fill.Solid()
                        shp.Fill.ForeColor.RGB = 16777215  # #FFFFFF
                        shp.Line.ForeColor.RGB = 15790306  # #E2E8F0
                        shp.Line.Weight = 0.75
                        shp.Line.DashStyle = 1  # msoLineSolid
                        # Clear placeholder watermark text if any
                        if shp.TextFrame2.HasText:
                            shp.TextFrame2.DeleteText()

                # Add or update Theme Switcher button on cockpit header
                btn_exists = False
                for shp in ws.Shapes:
                    if "ThemeToggle" in shp.Name or "btn_Theme" in shp.Name:
                        btn_exists = True
                        shp.OnAction = "modThemeEngine.ToggleDashboardTheme"
                        break
                
                if not btn_exists:
                    btn = ws.Shapes.AddShape(1, 1080, 24, 100, 28)  # msoShapeRoundedRectangle = 1
                    btn.Name = "btn_ThemeToggle"
                    btn.Fill.Solid()
                    btn.Fill.ForeColor.RGB = 2762511  # #0F172A
                    btn.Line.ForeColor.RGB = 1481168  # #D04A02
                    btn.Line.Weight = 1
                    btn.OnAction = "modThemeEngine.ToggleDashboardTheme"
                    btn.TextFrame2.TextRange.Text = "DARK MODE"
                    btn.TextFrame2.TextRange.Font.Name = "Segoe UI"
                    btn.TextFrame2.TextRange.Font.Size = 8.0
                    btn.TextFrame2.TextRange.Font.Bold = True
                    btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
                    btn.TextFrame2.TextRange.ParagraphFormat.Alignment = 2  # Center

                # Add Home Portal Navigation Button on cockpit header
                home_btn_exists = False
                for shp in ws.Shapes:
                    if "btn_HomePortal" in shp.Name:
                        home_btn_exists = True
                        break
                
                if not home_btn_exists:
                    btn_home = ws.Shapes.AddShape(1, 970, 24, 100, 28)
                    btn_home.Name = "btn_HomePortal"
                    btn_home.Fill.Solid()
                    btn_home.Fill.ForeColor.RGB = 16185078  # #F1F5F9
                    btn_home.Line.ForeColor.RGB = 15790306  # #E2E8F0
                    btn_home.Line.Weight = 1
                    ws.Hyperlinks.Add(
                        Anchor=btn_home,
                        Address="",
                        SubAddress="'00_Home_Portal'!A1",
                        ScreenTip="Return to Executive Portal"
                    )
                    btn_home.TextFrame2.TextRange.Text = "Home Portal"
                    btn_home.TextFrame2.TextRange.Font.Name = "Segoe UI"
                    btn_home.TextFrame2.TextRange.Font.Size = 8.0
                    btn_home.TextFrame2.TextRange.Font.Bold = True
                    btn_home.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 2762511
                    btn_home.TextFrame2.TextRange.ParagraphFormat.Alignment = 2

                # Format Charts on 03_CallCenter_Cockpit
                if sheet_name == "03_CallCenter_Cockpit":
                    for ch_obj in ws.ChartObjects():
                        ch = ch_obj.Chart
                        ch.ChartArea.Format.Fill.Visible = False
                        ch.ChartArea.Format.Line.Visible = False
                        ch.PlotArea.Format.Fill.Visible = False
                        ch.PlotArea.Format.Line.Visible = False
                        try:
                            ch.Axes(1).TickLabels.Font.Name = "Segoe UI"
                            ch.Axes(1).TickLabels.Font.Size = 8.0
                            ch.Axes(1).TickLabels.Font.Color = 6579300  # #64748B
                            ch.Axes(1).Format.Line.ForeColor.RGB = 15790306
                        except Exception:
                            pass
                        try:
                            ch.Axes(2).TickLabels.Font.Name = "Segoe UI"
                            ch.Axes(2).TickLabels.Font.Size = 8.0
                            ch.Axes(2).TickLabels.Font.Color = 6579300
                            ch.Axes(2).Format.Line.ForeColor.RGB = 15790306
                            ch.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB = 16382457
                        except Exception:
                            pass
                        if ch.HasLegend:
                            ch.Legend.Format.Fill.Visible = False
                            ch.Legend.Format.Line.Visible = False
                            ch.Legend.Font.Name = "Segoe UI"
                            ch.Legend.Font.Size = 8.0
                            ch.Legend.Font.Color = 6579300

            except Exception as e:
                print(f"Error updating {sheet_name}: {e}")

        # Set 00_Home_Portal as the active view
        try:
            ws_home = wb.Worksheets("00_Home_Portal")
            ws_home.Activate()
            ws_home.Range("A1").Select()
        except Exception:
            pass

        print("Saving workbook...")
        wb.Save()
        print("Workbook successfully saved!")

        wb.Close(SaveChanges=True)
    finally:
        excel.Quit()
        print("Excel process closed.")

if __name__ == "__main__":
    apply_updates()
