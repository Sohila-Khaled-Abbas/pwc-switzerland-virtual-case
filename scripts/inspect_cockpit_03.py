import win32com.client as win32
import os

excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb_path = os.path.abspath('PWC_Switzerland_Virtual_Case.xlsm')
    wb = excel.Workbooks.Open(wb_path)
    ws = wb.Worksheets('03_CallCenter_Cockpit')

    print("=== SHAPES ON 03_CallCenter_Cockpit ===")
    for shp in ws.Shapes:
        formula = ''
        try:
            formula = shp.DrawingObject.Formula
        except:
            pass
        txt = ''
        try:
            txt = shp.TextFrame.Characters().Text
        except:
            pass
        print(f"Name: {shp.Name} | Type: {shp.Type} | Left: {shp.Left:.1f} | Top: {shp.Top:.1f} | W: {shp.Width:.1f} | H: {shp.Height:.1f} | Text: {repr(txt[:30])} | Formula: {formula}")

    print("\n=== CHARTOBJECTS ===")
    for co in ws.ChartObjects():
        title = "None"
        try:
            if co.Chart.HasTitle:
                title = co.Chart.ChartTitle.Text
        except:
            pass
        print(f"Chart Name: {co.Name} | Left: {co.Left:.1f} | Top: {co.Top:.1f} | W: {co.Width:.1f} | H: {co.Height:.1f} | Title: {title}")

    print("\n=== SLICERS ===")
    for sc in wb.SlicerCaches:
        for sl in sc.Slicers:
            if sl.Parent.Name == ws.Name:
                print(f"Slicer: {sl.Name} | Caption: {sl.Caption} | Field: {sc.SourceName} | Left: {sl.Left:.1f} | Top: {sl.Top:.1f} | W: {sl.Width:.1f} | H: {sl.Height:.1f}")

    wb.Close(SaveChanges=False)
finally:
    excel.Quit()
