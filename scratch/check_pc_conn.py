import win32com.client
import os

wb_path = os.path.abspath("PWC_Switzerland_Virtual_Case.xlsm")
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(wb_path)
    print("=== PIVOT CACHES INSPECTION ===")
    for i in range(1, wb.PivotCaches().Count + 1):
        pc = wb.PivotCaches(i)
        try:
            conn = getattr(pc, "Connection", "N/A")
            cmd = getattr(pc, "CommandText", "N/A")
            print(f"PivotCache {i}: SourceType={pc.SourceType}, Connection={conn}, CommandText={cmd}")
        except Exception as e:
            print(f"PivotCache {i}: Error {e}")
            
    print("\nWorkbook Connections:")
    for c in wb.Connections:
        print(f"Connection: '{c.Name}', Type={c.Type}, Description='{c.Description}'")
        
    wb.Close(False)
except Exception as e:
    print(e)
finally:
    excel.Quit()
