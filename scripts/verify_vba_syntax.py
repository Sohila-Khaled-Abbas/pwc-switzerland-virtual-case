import re

def validate_vba_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        raw_lines = f.readlines()
        
    # Reconstruct logical lines (joining lines ending with _)
    logical_lines = []
    current_line = ""
    current_start_idx = 1
    
    for idx, line in enumerate(raw_lines, 1):
        s = line.rstrip()
        stripped = s.strip()
        
        # Check comment
        if stripped.startswith("'"):
            continue
            
        if current_line == "":
            current_start_idx = idx
            
        if s.endswith("_"):
            current_line += s[:-1] + " "
        else:
            current_line += s
            logical_lines.append((current_start_idx, current_line.strip()))
            current_line = ""
            
    stack = []
    errors = 0
    
    for start_idx, line in logical_lines:
        if not line or line.startswith("'"):
            continue
            
        lower = line.lower()
        
        # Multi-line If check
        if lower.startswith("if ") and lower.endswith(" then"):
            stack.append(("if", start_idx, line))
        elif lower == "end if":
            if not stack or stack[-1][0] != "if":
                print(f"Error: Mismatched 'End If' at line {start_idx}, stack top: {stack[-1] if stack else None}")
                errors += 1
            else:
                stack.pop()
        elif lower.startswith("with ") and not lower.endswith(" end with"):
            stack.append(("with", start_idx, line))
        elif lower == "end with":
            if not stack or stack[-1][0] != "with":
                print(f"Error: Mismatched 'End With' at line {start_idx}, stack top: {stack[-1] if stack else None}")
                errors += 1
            else:
                stack.pop()
        elif lower.startswith("select case "):
            stack.append(("select", start_idx, line))
        elif lower == "end select":
            if not stack or stack[-1][0] != "select":
                print(f"Error: Mismatched 'End Select' at line {start_idx}, stack top: {stack[-1] if stack else None}")
                errors += 1
            else:
                stack.pop()
        elif lower.startswith("for each ") or (lower.startswith("for ") and " to " in lower):
            stack.append(("for", start_idx, line))
        elif lower.startswith("next"):
            if not stack or stack[-1][0] != "for":
                print(f"Error: Mismatched 'Next' at line {start_idx}, stack top: {stack[-1] if stack else None}")
                errors += 1
            else:
                stack.pop()
        elif any(lower.startswith(p) for p in ["sub ", "public sub ", "private sub "]):
            stack.append(("sub", start_idx, line))
        elif lower == "end sub":
            if not stack or stack[-1][0] != "sub":
                print(f"Error: Mismatched 'End Sub' at line {start_idx}, stack top: {stack[-1] if stack else None}")
                errors += 1
            else:
                stack.pop()
        elif any(lower.startswith(p) for p in ["function ", "public function ", "private function "]):
            stack.append(("function", start_idx, line))
        elif lower == "end function":
            if not stack or stack[-1][0] != "function":
                print(f"Error: Mismatched 'End Function' at line {start_idx}, stack top: {stack[-1] if stack else None}")
                errors += 1
            else:
                stack.pop()
                
    print(f"File: {filepath}")
    print(f"Total raw lines: {len(raw_lines)}")
    print(f"Total logical lines: {len(logical_lines)}")
    print(f"Total validation errors: {errors}")
    print(f"Unclosed blocks remaining: {len(stack)}")
    for kind, lnum, ltext in stack:
        print(f"  Unclosed {kind} from line {lnum}: {ltext[:60]}")

if __name__ == '__main__':
    validate_vba_file('vba/modPwC_Unified_Master.bas')
