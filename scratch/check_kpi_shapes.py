import win32com.client
import os

wb_path = os.path.abspath("PWC_Switzerland_Virtual_Case.xlsm")
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(wb_path)
    for s_name in ["03_CallCenter_Cockpit", "04_CustomerRetention_Cockpit", "05_DiversityInclusion_Cockpit"]:
        ws = wb.Worksheets(s_name)
        print(f"\n--- Checking KPI Shapes on {s_name} ---")
        for shp in ws.Shapes:
            if "Card_" in shp.Name or "Value_" in shp.Name or "Label_" in shp.Name or "Subtext_" in shp.Name:
                form = ""
                try:
                    form = shp.DrawingObject.Formula
                except:
                    pass
                txt = ""
                try:
                    txt = shp.TextFrame2.TextRange.Text.replace("\r", " ").replace("\n", " ")
                except:
                    pass
                print(f"Shape '{shp.Name}': Formula='{form}', Text='{txt}'")
    wb.Close(False)
except Exception as e:
    print(e)
finally:
    excel.Quit()
