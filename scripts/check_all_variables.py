import os
import re

vba_dir = r"d:\courses\Data Analysis 26-27\7-Introducation to Data Fields (Excel)\11_Demos_and_Workbooks\10_Projects_and_Demos\PWC\vba"

# Global built-in keywords and VBA constants
BUILTINS = {
    'thisworkbook', 'activeworkbook', 'application', 'activesheet', 'err',
    'vbcrlf', 'vbcritical', 'vbinformation', 'vbexclamation', 'xlsheethidden',
    'vbtextcompare', 'vbyesno', 'vbyes', 'vbno', 'vbokonly', 'true', 'false',
    'nothing', 'empty', 'null', 'vbtab', 'vbred', 'vbgreen', 'vbblue',
    'xlup', 'xldown', 'xltoright', 'xltoleft', 'xlcenter', 'xltop', 'xlbottom',
    'msoanchormiddle', 'msoanchorcenter', 'msoaligncenter', 'msotrue', 'msofalse',
    'xlcalculationautomatic', 'xlcalculationmanual', 'cint', 'clng', 'cstr', 'cdbl',
    'cdate', 'csng', 'cvar', 'cbool', 'cbyte', 'mid', 'left', 'right', 'instr',
    'instrrev', 'len', 'trim', 'ltrim', 'rtrim', 'ucase', 'lcase', 'replace',
    'split', 'join', 'ubound', 'lbound', 'isnumeric', 'isdate', 'isempty', 'isnull',
    'isarray', 'iserror', 'isobject', 'format', 'formatpercent', 'formatcurrency',
    'val', 'str', 'now', 'date', 'time', 'timer', 'dateserial', 'timeserial',
    'year', 'month', 'day', 'hour', 'minute', 'second', 'weekday', 'monthname',
    'dir', 'msgbox', 'inputbox', 'doevents', 'beep', 'rgb', 'qbcolor', 'array',
    'environ', 'round', 'int', 'fix', 'abs', 'sgn', 'sqr', 'exp', 'log', 'sin', 'cos', 'tan',
    'curdir', 'filedatetime', 'filelen', 'getattr', 'me', 'activecell', 'selection'
}

# Module names for cross-module calls
MODULE_NAMES = {
    'modappstate', 'modcreategovernancesheets', 'moddashboarduiux',
    'moddatarefresh', 'modexportpdf', 'modfiltercontroller',
    'modinteractivescorecard', 'modnavigation', 'modpivottableformatting',
    'modportallanding', 'modpwc_unified_master', 'modthemeengine'
}

# Gather all public sub/function names across all modules
ALL_GLOBAL_PROCS = set()
for fname in os.listdir(vba_dir):
    if fname.endswith('.bas'):
        with open(os.path.join(vba_dir, fname), 'r', encoding='utf-8') as f:
            for m in re.finditer(r'(?:Public\s+)?(?:Sub|Function)\s+([A-Za-z0-9_]+)', f.read(), re.IGNORECASE):
                ALL_GLOBAL_PROCS.add(m.group(1).lower())

all_issues = []

for fname in sorted(os.listdir(vba_dir)):
    if not fname.endswith('.bas'):
        continue
    fpath = os.path.join(vba_dir, fname)
    with open(fpath, 'r', encoding='utf-8') as f:
        text = f.read()

    # Find module-level constants and variables
    module_vars = set(m.lower() for m in re.findall(r'^(?:Private|Public|Dim|Const)\s+([A-Za-z0-9_]+)', text, re.MULTILINE | re.IGNORECASE))
    
    # Check each sub/function
    pattern = re.compile(r'(?:Public\s+|Private\s+)?(Sub|Function)\s+([A-Za-z0-9_]+)\s*\((.*?)\)([\s\S]*?)End\s+(?:Sub|Function)', re.IGNORECASE)
    
    for match in pattern.finditer(text):
        kind, name, args_str, body = match.groups()
        
        # Extract argument names
        args = set()
        for arg in args_str.split(','):
            arg = arg.strip()
            if not arg:
                continue
            arg = re.sub(r'^(?:Optional\s+)?(?:ByVal\s+|ByRef\s+)?', '', arg, flags=re.IGNORECASE)
            arg_name = arg.split()[0]
            args.add(arg_name.lower())
            
        dimmed = set(a.lower() for a in args)
        dimmed.update(module_vars)
        dimmed.update(BUILTINS)
        dimmed.update(MODULE_NAMES)
        dimmed.update(ALL_GLOBAL_PROCS)
        dimmed.add(name.lower())
        
        # Extract Dim statements in body
        for m in re.finditer(r'Dim\s+([^:\n\r]+)', body, re.IGNORECASE):
            decls = m.group(1).split(',')
            for d in decls:
                parts = d.strip().split()
                if parts:
                    dimmed.add(parts[0].lower())
                    
        # For Each var
        for m in re.finditer(r'For\s+Each\s+([A-Za-z0-9_]+)', body, re.IGNORECASE):
            dimmed.add(m.group(1).lower())
        for m in re.finditer(r'For\s+([A-Za-z0-9_]+)\s*=', body, re.IGNORECASE):
            dimmed.add(m.group(1).lower())

        # Check Set <var> = ...
        for m in re.finditer(r'Set\s+([A-Za-z0-9_]+)\s*=', body, re.IGNORECASE):
            var_name = m.group(1).lower()
            if var_name not in dimmed:
                all_issues.append((fname, name, f"Set target '{m.group(1)}' not declared!"))

        # Check <var> = ... (assignments)
        for line in body.splitlines():
            line_str = line.strip()
            if line_str.startswith("'") or line_str.startswith("Rem ") or not line_str:
                continue
            # Look for: Set x = y or x = y
            m_set = re.match(r'Set\s+([A-Za-z0-9_]+)\s*=', line_str, re.IGNORECASE)
            if m_set:
                v = m_set.group(1).lower()
                if v not in dimmed:
                    all_issues.append((fname, name, f"Variable in Set '{m_set.group(1)}' not declared!"))
                # Check right hand side: e.g. Set x = wb.Worksheets(...)
                rhs = line_str[m_set.end():].strip()
                m_rhs_obj = re.match(r'([A-Za-z0-9_]+)\.', rhs)
                if m_rhs_obj:
                    obj = m_rhs_obj.group(1).lower()
                    if obj not in dimmed:
                        all_issues.append((fname, name, f"Object '{m_rhs_obj.group(1)}' in RHS '{line_str}' not declared!"))

if all_issues:
    print(f"Found {len(all_issues)} issues:")
    for f, n, err in all_issues:
        print(f"  {f} :: {n} -> {err}")
else:
    print("Zero undeclared variable issues found across all modules!")
