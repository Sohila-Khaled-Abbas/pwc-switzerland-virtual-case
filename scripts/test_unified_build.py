import os
import shutil
import win32com.client as win32

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKBOOK_SRC = os.path.join(BASE_DIR, "PWC_Switzerland_Virtual_Case.xlsm")
WORKBOOK_TEST = os.path.join(BASE_DIR, "scratch", "test_build_full.xlsm")
VBA_DIR = os.path.join(BASE_DIR, "vba")

print(f"Creating test copy: {WORKBOOK_TEST}")
shutil.copy2(WORKBOOK_SRC, WORKBOOK_TEST)

excel = win32.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.ScreenUpdating = False

try:
    print(f"Opening test workbook: {WORKBOOK_TEST}")
    wb = excel.Workbooks.Open(WORKBOOK_TEST)
    vbp = wb.VBProject

    # 1. Update all 12 VBA Modules
    print("\n--- Synchronizing all 12 VBA modules into VBProject ---")
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
        bas_file = os.path.join(VBA_DIR, f"{mod_name}.bas")
        if not os.path.exists(bas_file):
            print(f"  Warning: {bas_file} not found!")
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
        print(f"  Updated {mod_name} ({cm.CountOfLines} lines)")

    print("\n--- Running modPwC_Unified_Master.RunUnifiedPwCPlatform ---")
    excel.Run("modPwC_Unified_Master.RunUnifiedPwCPlatform")
    print("Execution completed without unhandled exceptions!")

    print("\n--- Worksheets present after execution ---")
    for ws in wb.Worksheets:
        print(f"  Sheet: {ws.Name}, Visible: {ws.Visible}, Shapes: {ws.Shapes.Count}, PivotTables: {ws.PivotTables().Count}")

    wb.Save()
    print("Test build saved successfully.")
    wb.Close(SaveChanges=True)

except Exception as e:
    print("Error during test build:", e)
    import traceback
    traceback.print_exc()
finally:
    excel.Quit()
    print("Excel closed.")
