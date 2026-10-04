import win32com.client
import os

wb_path = os.path.abspath("PWC_Switzerland_Virtual_Case_backup.xlsm")
if not os.path.exists(wb_path):
    print(f"File not found: {wb_path}")
    exit(0)

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(wb_path)
    print("=== INSPECTING PWC_Switzerland_Virtual_Case_backup.xlsm ===")
    for ws in wb.Worksheets:
        print(f"Sheet '{ws.Name}' has {ws.PivotTables().Count} PivotTables, {ws.ChartObjects().Count} Charts:")
        for pt in ws.PivotTables():
            print(f"  - PT: {pt.Name} at {pt.TableRange2.Address if hasattr(pt, 'TableRange2') else 'N/A'}")
        for co in ws.ChartObjects():
            has_pivot = False
            pt_name = None
            try:
                if co.Chart.PivotLayout:
                    has_pivot = True
                    pt_name = co.Chart.PivotLayout.PivotTable.Name
            except Exception: pass
            print(f"  - Chart: '{co.Name}', HasPivot: {has_pivot}, PT: {pt_name}")

    wb.Close(False)
except Exception as e:
    print(f"Error: {e}")
finally:
    excel.Quit()
