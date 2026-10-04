import win32com.client
import os

wb_path = os.path.abspath("PWC_Switzerland_Virtual_Case.xlsm")
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(wb_path)
    for ws in wb.Worksheets:
        print(f"Sheet '{ws.Name}' has {ws.PivotTables().Count} PivotTables:")
        for pt in ws.PivotTables():
            print(f"  - {pt.Name} at {pt.TableRange2.Address if hasattr(pt, 'TableRange2') else 'N/A'}")
    
    print("\nPivotCaches in Workbook:")
    print(f"Total PivotCaches: {wb.PivotCaches().Count}")
    for idx, pc in enumerate(wb.PivotCaches(), 1):
        try:
            print(f"  Cache {idx}: SourceType={pc.SourceType}")
        except Exception as e:
            print(f"  Cache {idx}: {e}")

    wb.Close(False)
except Exception as e:
    print(f"Error: {e}")
finally:
    excel.Quit()
