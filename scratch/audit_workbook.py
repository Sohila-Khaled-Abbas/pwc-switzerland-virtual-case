import win32com.client
import os
import json

wb_path = os.path.abspath("PWC_Switzerland_Virtual_Case.xlsm")
print(f"Auditing: {wb_path}")

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

report = {}

try:
    wb = excel.Workbooks.Open(wb_path)
    report["Worksheets"] = [ws.Name for ws in wb.Worksheets]
    
    # 1. Model Tables & Measures
    model_info = {"Tables": [], "Measures": []}
    try:
        for t in wb.Model.ModelTables:
            model_info["Tables"].append(t.Name)
    except Exception as e:
        model_info["Tables_Error"] = str(e)

    try:
        for m in wb.Model.ModelMeasures:
            m_dict = {"Name": m.Name}
            try: m_dict["Formula"] = m.Formula
            except: pass
            try: m_dict["Description"] = m.Description
            except: pass
            model_info["Measures"].append(m_dict)
    except Exception as e:
        model_info["Measures_Error"] = str(e)
    report["Model"] = model_info

    # 2. Slicer Caches
    slicers_info = []
    try:
        for sc in wb.SlicerCaches:
            sc_data = {
                "Name": sc.Name,
                "SourceName": getattr(sc, "SourceName", ""),
                "Slicers": [s.Name for s in sc.Slicers],
                "PivotTables": [pt.Name for pt in sc.PivotTables]
            }
            slicers_info.append(sc_data)
    except Exception as e:
        slicers_info.append({"Error": str(e)})
    report["SlicerCaches"] = slicers_info

    # 3. Sheets Details
    sheets_info = {}
    for ws in wb.Worksheets:
        s_data = {
            "Visible": ws.Visible,
            "ShapesCount": ws.Shapes.Count,
            "ChartObjectsCount": ws.ChartObjects().Count,
            "PivotTablesCount": ws.PivotTables().Count,
            "PivotTables": [],
            "ChartObjects": [],
            "Shapes": []
        }
        
        # PivotTables
        for pt in ws.PivotTables():
            try:
                s_data["PivotTables"].append({
                    "Name": pt.Name,
                    "TableRange": pt.TableRange2.Address if hasattr(pt, "TableRange2") else None
                })
            except Exception as pe:
                s_data["PivotTables"].append({"Error": str(pe)})

        # ChartObjects
        for co in ws.ChartObjects():
            try:
                ch = co.Chart
                has_pivot = False
                pivot_name = None
                try:
                    if ch.PivotLayout:
                        has_pivot = True
                        pivot_name = ch.PivotLayout.PivotTable.Name
                except Exception:
                    pass
                s_data["ChartObjects"].append({
                    "Name": co.Name,
                    "Left": co.Left,
                    "Top": co.Top,
                    "Width": co.Width,
                    "Height": co.Height,
                    "HasPivotLayout": has_pivot,
                    "PivotTable": pivot_name,
                    "ChartType": ch.ChartType
                })
            except Exception as ce:
                s_data["ChartObjects"].append({"Name": co.Name, "Error": str(ce)})

        # Shapes summary (sample top 30)
        for i in range(1, min(ws.Shapes.Count + 1, 45)):
            try:
                shp = ws.Shapes(i)
                txt = ""
                try:
                    txt = shp.TextFrame2.TextRange.Text[:40]
                except:
                    pass
                s_data["Shapes"].append({
                    "Name": shp.Name,
                    "Type": shp.Type,
                    "Left": round(shp.Left, 1),
                    "Top": round(shp.Top, 1),
                    "Width": round(shp.Width, 1),
                    "Height": round(shp.Height, 1),
                    "Text": txt
                })
            except Exception:
                pass

        # Helper cells AA64:AE66
        try:
            helper_cells = {}
            for col in ["AA", "AB", "AC", "AD", "AE"]:
                for row in [64, 65, 66]:
                    c_addr = f"{col}{row}"
                    val = ws.Range(c_addr).Value
                    form = ws.Range(c_addr).Formula
                    if val is not None or (form and form != ""):
                        helper_cells[c_addr] = {"Value": str(val), "Formula": form}
            s_data["HelperCells"] = helper_cells
        except Exception as e:
            s_data["HelperCellsError"] = str(e)

        sheets_info[ws.Name] = s_data
    report["Sheets"] = sheets_info

    # 4. VBA Components
    vba_info = []
    try:
        for c in wb.VBProject.VBComponents:
            vba_info.append({
                "Name": c.Name,
                "Type": c.Type,
                "Lines": c.CodeModule.CountOfLines
            })
    except Exception as e:
        vba_info.append({"Error": str(e)})
    report["VBA"] = vba_info

    wb.Close(False)
except Exception as e:
    report["FatalError"] = str(e)
finally:
    excel.Quit()

with open("scratch/audit_report.json", "w", encoding="utf-8") as f:
    json.dump(report, f, indent=2)

print("Audit complete! Report written to scratch/audit_report.json")
