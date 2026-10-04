"""
VBA MCP Server - Excel VBA code direct read/write tool
======================================================
Claude Code can directly read/write VBA code in .xlsm files.
No manual export/import needed.

Tools:
  - vba_list_modules    : List all VBA modules in .xlsm
  - vba_read_module     : Read a specific module's code
  - vba_read_all        : Read all modules at once
  - vba_write_module    : Write/replace module code (auto-save)
  - vba_create_module   : Create a new module
  - vba_delete_module   : Delete a module
  - vba_backup          : Backup all VBA code to text files
  - vba_list_workbooks  : List .xlsm files in a directory

Requirements:
  - Windows + Excel (local install)
  - Excel: Trust access to VBA project object model
  - pip install mcp pywin32
"""

import os
import sys
import json
import time
import glob
import traceback
from datetime import datetime
from pathlib import Path
from contextlib import contextmanager
from typing import Optional

from mcp.server.fastmcp import FastMCP

mcp = FastMCP("vba-mcp-server")

VB_COMP_TYPES = {
    1: "StandardModule",
    2: "ClassModule",
    3: "UserForm",
    100: "Document",
}

VB_COMP_TYPE_EXT = {
    1: ".bas",
    2: ".cls",
    3: ".frm",
    100: ".cls",
}


def _get_excel_app(visible=False):
    import win32com.client
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = visible
    excel.DisplayAlerts = False
    return excel


def _find_open_workbook(excel, file_path):
    abs_path = os.path.abspath(file_path).lower()
    for wb in excel.Workbooks:
        try:
            if os.path.abspath(wb.FullName).lower() == abs_path:
                return wb
        except Exception:
            continue
    return None


@contextmanager
def open_workbook(file_path, save_on_close=False):
    abs_path = os.path.abspath(file_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"File not found: {abs_path}")
    if not abs_path.lower().endswith(('.xlsm', '.xlsb', '.xls', '.xla', '.xlam')):
        raise ValueError(f"Not a VBA-capable file: {abs_path}")

    excel = _get_excel_app()
    wb = _find_open_workbook(excel, abs_path)
    already_open = wb is not None

    if not already_open:
        wb = excel.Workbooks.Open(abs_path)

    try:
        yield wb
    finally:
        if save_on_close and wb:
            wb.Save()
        if not already_open and wb:
            wb.Close(SaveChanges=save_on_close)


def _module_info(comp):
    comp_type = comp.Type
    code_module = comp.CodeModule
    line_count = code_module.CountOfLines if code_module else 0
    return {
        "name": comp.Name,
        "type": VB_COMP_TYPES.get(comp_type, f"Unknown({comp_type})"),
        "type_id": comp_type,
        "line_count": line_count,
    }


def _trust_access_hint():
    return (
        "Enable 'Trust access to the VBA project object model' in Excel.\n"
        "File > Options > Trust Center > Trust Center Settings > Macro Settings > Check the box."
    )


@mcp.tool()
def vba_list_modules(file_path: str) -> str:
    """List all VBA modules in an .xlsm file."""
    try:
        with open_workbook(file_path) as wb:
            modules = []
            for comp in wb.VBProject.VBComponents:
                modules.append(_module_info(comp))
            return json.dumps(modules, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e), "hint": _trust_access_hint()}, ensure_ascii=False)


@mcp.tool()
def vba_read_module(file_path: str, module_name: str) -> str:
    """Read source code of a specific VBA module."""
    try:
        with open_workbook(file_path) as wb:
            try:
                comp = wb.VBProject.VBComponents(module_name)
            except Exception:
                available = [c.Name for c in wb.VBProject.VBComponents]
                return json.dumps({
                    "error": f"Module '{module_name}' not found",
                    "available_modules": available
                }, ensure_ascii=False)

            cm = comp.CodeModule
            if cm.CountOfLines == 0:
                return f"# {module_name} (empty module)\n"
            code = cm.Lines(1, cm.CountOfLines)
            return f"' === Module: {module_name} ({VB_COMP_TYPES.get(comp.Type, 'Unknown')}) ===\n{code}"
    except Exception as e:
        return json.dumps({"error": str(e), "hint": _trust_access_hint()}, ensure_ascii=False)


@mcp.tool()
def vba_read_all(file_path: str) -> str:
    """Read ALL VBA modules at once. Useful for full codebase review."""
    try:
        with open_workbook(file_path) as wb:
            result = []
            for comp in wb.VBProject.VBComponents:
                cm = comp.CodeModule
                header = f"\n{'='*60}\n' Module: {comp.Name}\n' Type: {VB_COMP_TYPES.get(comp.Type, 'Unknown')}\n' Lines: {cm.CountOfLines}\n{'='*60}"
                if cm.CountOfLines > 0:
                    code = cm.Lines(1, cm.CountOfLines)
                    result.append(f"{header}\n{code}")
                else:
                    result.append(f"{header}\n' (empty)")
            return "\n".join(result)
    except Exception as e:
        return json.dumps({"error": str(e), "hint": _trust_access_hint()}, ensure_ascii=False)


@mcp.tool()
def vba_write_module(file_path: str, module_name: str, code: str, create_if_missing: bool = True) -> str:
    """Write/replace VBA module code. Auto-saves the workbook."""
    try:
        with open_workbook(file_path, save_on_close=True) as wb:
            comp = None
            for c in wb.VBProject.VBComponents:
                if c.Name == module_name:
                    comp = c
                    break

            if comp is None:
                if not create_if_missing:
                    return json.dumps({"error": f"Module '{module_name}' not found"}, ensure_ascii=False)
                comp = wb.VBProject.VBComponents.Add(1)
                comp.Name = module_name

            cm = comp.CodeModule
            if cm.CountOfLines > 0:
                cm.DeleteLines(1, cm.CountOfLines)
            if code.strip():
                cm.AddFromString(code)

            return json.dumps({
                "success": True,
                "module": module_name,
                "lines_written": cm.CountOfLines,
                "message": f"Module '{module_name}' updated ({cm.CountOfLines} lines)"
            }, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e), "hint": _trust_access_hint()}, ensure_ascii=False)


@mcp.tool()
def vba_create_module(file_path: str, module_name: str, code: str = "", module_type: str = "standard") -> str:
    """Create a new VBA module."""
    type_map = {"standard": 1, "class": 2}
    comp_type = type_map.get(module_type.lower(), 1)
    try:
        with open_workbook(file_path, save_on_close=True) as wb:
            for c in wb.VBProject.VBComponents:
                if c.Name == module_name:
                    return json.dumps({"error": f"Module '{module_name}' already exists"}, ensure_ascii=False)
            comp = wb.VBProject.VBComponents.Add(comp_type)
            comp.Name = module_name
            if code.strip():
                comp.CodeModule.AddFromString(code)
            return json.dumps({
                "success": True,
                "module": module_name,
                "type": VB_COMP_TYPES.get(comp_type),
                "message": f"Module '{module_name}' created"
            }, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e), "hint": _trust_access_hint()}, ensure_ascii=False)


@mcp.tool()
def vba_delete_module(file_path: str, module_name: str) -> str:
    """Delete a VBA module (Sheet/ThisWorkbook modules cannot be deleted, only cleared)."""
    try:
        with open_workbook(file_path, save_on_close=True) as wb:
            comp = None
            for c in wb.VBProject.VBComponents:
                if c.Name == module_name:
                    comp = c
                    break
            if comp is None:
                return json.dumps({"error": f"Module '{module_name}' not found"}, ensure_ascii=False)
            if comp.Type == 100:
                return json.dumps({"error": f"'{module_name}' is a Document module and cannot be deleted. Use vba_write_module to clear its code."}, ensure_ascii=False)
            wb.VBProject.VBComponents.Remove(comp)
            return json.dumps({"success": True, "message": f"Module '{module_name}' deleted"}, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e), "hint": _trust_access_hint()}, ensure_ascii=False)


@mcp.tool()
def vba_backup(file_path: str, backup_dir: Optional[str] = None) -> str:
    """Backup all VBA modules to text files."""
    try:
        abs_path = os.path.abspath(file_path)
        if backup_dir is None:
            backup_dir = os.path.join(os.path.dirname(abs_path), "VBA_Backup")
        os.makedirs(backup_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest_dir = os.path.join(backup_dir, timestamp)
        os.makedirs(dest_dir, exist_ok=True)

        exported = []
        with open_workbook(file_path) as wb:
            for comp in wb.VBProject.VBComponents:
                ext = VB_COMP_TYPE_EXT.get(comp.Type, ".txt")
                out_path = os.path.join(dest_dir, f"{comp.Name}{ext}")
                comp.Export(out_path)
                exported.append({
                    "name": comp.Name,
                    "type": VB_COMP_TYPES.get(comp.Type, "Unknown"),
                    "file": out_path,
                    "lines": comp.CodeModule.CountOfLines,
                })

        manifest = {"source_file": abs_path, "backup_time": timestamp, "modules": exported}
        with open(os.path.join(dest_dir, "_manifest.json"), 'w', encoding='utf-8') as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2)

        return json.dumps({
            "success": True,
            "backup_dir": dest_dir,
            "file_count": len(exported),
            "files": exported,
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


@mcp.tool()
def vba_list_workbooks(directory: str, recursive: bool = False) -> str:
    """List .xlsm files in a directory."""
    try:
        abs_dir = os.path.abspath(directory)
        if not os.path.isdir(abs_dir):
            return json.dumps({"error": f"Directory not found: {abs_dir}"}, ensure_ascii=False)
        pattern = os.path.join(abs_dir, "**", "*.xlsm") if recursive else os.path.join(abs_dir, "*.xlsm")
        files = glob.glob(pattern, recursive=recursive)
        result = []
        for f in sorted(files):
            stat = os.stat(f)
            result.append({
                "path": f,
                "name": os.path.basename(f),
                "size_kb": round(stat.st_size / 1024, 1),
                "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
            })
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


if __name__ == "__main__":
    mcp.run(transport="stdio")
