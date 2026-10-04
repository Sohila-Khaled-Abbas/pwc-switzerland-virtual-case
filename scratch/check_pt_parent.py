import win32com.client
import os

wb_path = os.path.abspath("PWC_Switzerland_Virtual_Case.xlsm")
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(wb_path)
    ws = wb.Worksheets("03_CallCenter_Cockpit")
    for co in ws.ChartObjects():
        try:
            pt = co.Chart.PivotLayout.PivotTable
            print(f"Chart: {co.Name} -> PivotTable: {pt.Name} on Sheet: '{pt.Parent.Name}', Range: {pt.TableRange2.Address}")
        except Exception as e:
            print(f"Chart: {co.Name} -> No PivotTable ({e})")
            
    wb.Close(False)
except Exception as e:
    print(e)
finally:
    excel.Quit()
