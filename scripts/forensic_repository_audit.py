import os
import json
import hashlib
import win32com.client

def run_forensic_audit():
    root = os.path.abspath(".")
    report = {
        "timestamp": "2026-10-04",
        "filesystem": {},
        "workbook_forensics": {},
        "discrepancies": []
    }
    
    # 1. Filesystem inventory
    wb_files = [f for f in os.listdir(root) if f.endswith(('.xlsx', '.xlsm', '.xlsb', '.xls'))]
    report["filesystem"]["workbook_files"] = wb_files
    
    vba_files = [f for f in os.listdir(os.path.join(root, "vba")) if f.endswith(('.bas', '.cls', '.frm'))] if os.path.exists("vba") else []
    report["filesystem"]["vba_files"] = vba_files
    
    pq_files = [f for f in os.listdir(os.path.join(root, "power_query")) if f.endswith('.m')] if os.path.exists("power_query") else []
    report["filesystem"]["power_query_files"] = pq_files
    
    dax_files = [f for f in os.listdir(os.path.join(root, "dax")) if f.endswith('.dax')] if os.path.exists("dax") else []
    report["filesystem"]["dax_files"] = dax_files
    
    dashboards_files = os.listdir(os.path.join(root, "dashboards")) if os.path.exists("dashboards") else []
    report["filesystem"]["dashboards_files"] = dashboards_files
    
    data_files = os.listdir(os.path.join(root, "data")) if os.path.exists("data") else []
    report["filesystem"]["data_files"] = data_files
    
    scripts_files = os.listdir(os.path.join(root, "scripts")) if os.path.exists("scripts") else []
    report["filesystem"]["scripts_files"] = scripts_files

    # 2. Check for duplicate VBA procedures across .bas files
    proc_defs = {}
    for vf in vba_files:
        vp = os.path.join(root, "vba", vf)
        with open(vp, "r", encoding="utf-8", errors="ignore") as f:
            for lnum, line in enumerate(f, 1):
                sl = line.strip()
                if (sl.startswith("Sub ") or sl.startswith("Public Sub ") or sl.startswith("Private Sub ") or
                    sl.startswith("Function ") or sl.startswith("Public Function ") or sl.startswith("Private Function ")):
                    parts = sl.split()
                    proc_name = ""
                    for p in parts[1:]:
                        if p.startswith("("):
                            break
                        if "(" in p:
                            proc_name = p.split("(")[0]
                            break
                        if p not in ["Sub", "Function"]:
                            proc_name = p
                            break
                    if proc_name:
                        proc_defs.setdefault(proc_name, []).append((vf, lnum))
                        
    duplicate_procs = {p: locs for p, locs in proc_defs.items() if len(locs) > 1}
    report["vba_analysis"] = {
        "total_unique_procedures": len(proc_defs),
        "duplicate_procedures": duplicate_procs
    }

    # 3. Workbook deep inspection via Excel COM
    wb_name = "PWC_Switzerland_Virtual_Case.xlsm"
    wb_path = os.path.join(root, wb_name)
    
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    
    try:
        wb = excel.Workbooks.Open(wb_path, UpdateLinks=False, ReadOnly=True)
        wf = {}
        wf["name"] = wb.Name
        wf["sheets"] = []
        wf["hidden_sheets"] = []
        wf["very_hidden_sheets"] = []
        
        for ws in wb.Worksheets:
            if ws.Visible == -1: # xlSheetVisible
                wf["sheets"].append(ws.Name)
            elif ws.Visible == 0: # xlSheetHidden
                wf["hidden_sheets"].append(ws.Name)
            else: # xlSheetVeryHidden
                wf["very_hidden_sheets"].append(ws.Name)
                
        # Connections
        wf["connections"] = []
        try:
            for c in wb.Connections:
                wf["connections"].append({
                    "name": c.Name,
                    "type": str(c.Type),
                    "description": getattr(c, "Description", "")
                })
        except Exception as e:
            wf["connections_error"] = str(e)
            
        # Data Model
        wf["data_model"] = {"tables": [], "measures": [], "relationships": []}
        try:
            for t in wb.Model.ModelTables:
                wf["data_model"]["tables"].append(t.Name)
        except Exception as e:
            wf["data_model"]["tables_error"] = str(e)
            
        try:
            for m in wb.Model.ModelMeasures:
                tbl_name = ""
                try:
                    tbl_name = m.AssociatedTable.Name
                except:
                    tbl_name = ""
                wf["data_model"]["measures"].append({
                    "name": m.Name,
                    "table": tbl_name,
                    "formula": str(getattr(m, "Formula", ""))
                })
        except Exception as e:
            wf["data_model"]["measures_error"] = str(e)
            
        try:
            for r in wb.Model.ModelRelationships:
                wf["data_model"]["relationships"].append({
                    "from_table": r.ForeignKeyTable.Name,
                    "from_column": r.ForeignKeyColumn.Name,
                    "to_table": r.PrimaryKeyTable.Name,
                    "to_column": r.PrimaryKeyColumn.Name,
                    "active": r.Active
                })
        except Exception as e:
            wf["data_model"]["relationships_error"] = str(e)
            
        # SlicerCaches
        wf["slicer_caches"] = []
        try:
            for sc in wb.SlicerCaches:
                sc_info = {
                    "name": sc.Name,
                    "source_name": getattr(sc, "SourceName", ""),
                    "slicers": [],
                    "pivot_tables": [pt.Name for pt in sc.PivotTables]
                }
                for s in sc.Slicers:
                    sc_info["slicers"].append({
                        "name": s.Name,
                        "caption": s.Caption,
                        "parent_sheet": s.Shape.Parent.Name if hasattr(s, "Shape") else ""
                    })
                wf["slicer_caches"].append(sc_info)
        except Exception as e:
            wf["slicer_caches_error"] = str(e)
            
        # Named ranges
        wf["names"] = []
        try:
            for n in wb.Names:
                wf["names"].append({
                    "name": n.Name,
                    "refers_to": n.RefersTo,
                    "visible": n.Visible
                })
        except Exception as e:
            wf["names_error"] = str(e)
            
        # VBA inside workbook
        wf["vba_components"] = []
        try:
            for c in wb.VBProject.VBComponents:
                wf["vba_components"].append({
                    "name": c.Name,
                    "type": c.Type,
                    "line_count": c.CodeModule.CountOfLines
                })
        except Exception as e:
            wf["vba_components_error"] = str(e)
            
        # Per sheet inspection: Tables, PivotTables, Charts, Shapes, Formulas, Errors
        sheet_details = {}
        formula_errors = []
        shape_names_seen = {}
        
        for ws in wb.Worksheets:
            sd = {
                "name": ws.Name,
                "visible": ws.Visible,
                "list_objects": [],
                "pivot_tables": [],
                "charts": [],
                "shapes": [],
                "error_cells": [],
                "hidden_rows": [],
                "hidden_cols": []
            }
            
            # ListObjects (Excel Tables)
            for lo in ws.ListObjects:
                sd["list_objects"].append({
                    "name": lo.Name,
                    "range": lo.Range.Address
                })
                
            # PivotTables
            for pt in ws.PivotTables():
                try:
                    sd["pivot_tables"].append({
                        "name": pt.Name,
                        "source_data": getattr(pt, "SourceData", ""),
                        "range": pt.TableRange2.Address if hasattr(pt, "TableRange2") else ""
                    })
                except Exception as pte:
                    sd["pivot_tables"].append({"name": pt.Name, "error": str(pte)})
                    
            # Charts
            for co in ws.ChartObjects():
                try:
                    ch = co.Chart
                    pt_parent = None
                    try:
                        if ch.PivotLayout:
                            pt_parent = ch.PivotLayout.PivotTable.Name
                    except:
                        pass
                    sd["charts"].append({
                        "name": co.Name,
                        "title": ch.ChartTitle.Text if ch.HasTitle else "",
                        "left": round(co.Left, 1),
                        "top": round(co.Top, 1),
                        "width": round(co.Width, 1),
                        "height": round(co.Height, 1),
                        "is_pivot_chart": pt_parent is not None,
                        "pivot_table": pt_parent
                    })
                except Exception as che:
                    sd["charts"].append({"name": co.Name, "error": str(che)})
                    
            # Shapes & duplicate shape names
            for s_idx in range(1, ws.Shapes.Count + 1):
                try:
                    shp = ws.Shapes(s_idx)
                    s_name = shp.Name
                    shape_names_seen.setdefault(s_name, []).append((ws.Name, s_idx))
                    
                    # Macro callback / OnAction
                    on_act = ""
                    try: on_act = shp.OnAction
                    except: pass
                    
                    # Hyperlink
                    hl_addr = ""
                    hl_sub = ""
                    try:
                        if shp.Hyperlink:
                            hl_addr = shp.Hyperlink.Address
                            hl_sub = shp.Hyperlink.SubAddress
                    except: pass
                    
                    # Text
                    txt = ""
                    try: txt = shp.TextFrame2.TextRange.Text[:50]
                    except: pass
                    
                    sd["shapes"].append({
                        "name": s_name,
                        "type": shp.Type,
                        "left": round(shp.Left, 1),
                        "top": round(shp.Top, 1),
                        "width": round(shp.Width, 1),
                        "height": round(shp.Height, 1),
                        "on_action": on_act,
                        "hyperlink_sub": hl_sub,
                        "text": txt
                    })
                except Exception as se:
                    pass
                    
            # Scan used range for formula errors (#REF!, #VALUE!, #DIV/0!, #N/A, #NAME?, #NUM!)
            try:
                ur = ws.UsedRange
                # SpecialCells for formulas with errors
                try:
                    err_cells = ur.SpecialCells(3, 16) # xlCellTypeFormulas, xlErrors
                    for cell in err_cells:
                        val_str = str(cell.Text)
                        formula_errors.append({
                            "sheet": ws.Name,
                            "cell": cell.Address,
                            "formula": cell.Formula,
                            "value": val_str
                        })
                        sd["error_cells"].append({
                            "cell": cell.Address,
                            "formula": cell.Formula,
                            "value": val_str
                        })
                except:
                    # No error formula cells found via SpecialCells
                    pass
            except Exception as e:
                sd["scan_error"] = str(e)
                
            sheet_details[ws.Name] = sd
            
        wf["sheet_details"] = sheet_details
        wf["formula_errors"] = formula_errors
        wf["duplicate_shape_names"] = {name: locs for name, locs in shape_names_seen.items() if len(locs) > 1}
        
        report["workbook_forensics"] = wf
        wb.Close(False)
    except Exception as e:
        report["workbook_error"] = str(e)
    finally:
        excel.Quit()
        
    out_file = os.path.join(root, "scratch", "forensic_audit_full.json")
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2)
        
    print(f"Forensic audit written to {out_file}")
    print(f"Total Sheets: {len(report['workbook_forensics'].get('sheets', []))} visible, {len(report['workbook_forensics'].get('hidden_sheets', []))} hidden")
    print(f"Model Tables: {len(report['workbook_forensics'].get('data_model', {}).get('tables', []))}")
    print(f"Model Measures: {len(report['workbook_forensics'].get('data_model', {}).get('measures', []))}")
    print(f"Model Relationships: {len(report['workbook_forensics'].get('data_model', {}).get('relationships', []))}")
    print(f"Slicer Caches: {len(report['workbook_forensics'].get('slicer_caches', []))}")
    print(f"Formula Errors: {len(report['workbook_forensics'].get('formula_errors', []))}")
    print(f"Duplicate Procedures across .bas: {len(report['vba_analysis']['duplicate_procedures'])}")

if __name__ == "__main__":
    run_forensic_audit()
