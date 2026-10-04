import os
import re

def main():
    print("--- 1. Sanitizing existing modules for pure ASCII ---")
    vba_dir = 'vba'
    
    # Common replacements for non-ascii
    replacements = [
        (b'\xe2\x80\x94', b'--'),      # em-dash
        (b'\xe2\x80\x93', b'-'),       # en-dash
        (b'\xe2\x80\x9c', b'"'),       # left dquote
        (b'\xe2\x80\x9d', b'"'),       # right dquote
        (b'\xe2\x80\x98', b"'"),       # left squote
        (b'\xe2\x80\x99', b"'"),       # right squote
        (b'\xe2\x96\xb2', b'ChrW(9650)'), # triangle up
        (b'\xe2\x96\xbc', b'ChrW(9660)'), # triangle down
        (b'\xc2\xa0', b' '),          # non-breaking space
        (b'\xe2\x80\xa2', b'*'),       # bullet
        (b'\xc2\xb7', b'-'),           # middle dot
    ]
    
    for fname in ['modAppState.bas', 'modNavigation.bas', 'modDataRefresh.bas', 
                  'modFilterController.bas', 'modExportPDF.bas', 
                  'modCreateGovernanceSheets.bas', 'modPortalLanding.bas',
                  'modInteractiveScorecard.bas', 'modPivotTableFormatting.bas',
                  'modThemeEngine.bas']:
        path = os.path.join(vba_dir, fname)
        with open(path, 'rb') as f:
            content = f.read()
        for k, v in replacements:
            content = content.replace(k, v)
        # Check any remaining non-ascii
        rem = [b for b in content if b > 127]
        if rem:
            print(f"Warning: {fname} still has {len(rem)} non-ascii bytes")
            # Replace any other byte > 127 with space or ASCII
            content = bytes([b if b <= 127 else ord('?') for b in content])
        with open(path, 'wb') as f:
            f.write(content)
        print(f"Sanitized {fname}: {len(content)} bytes")

    # In modThemeEngine.bas, convert Public Const to Private Const
    theme_path = os.path.join(vba_dir, 'modThemeEngine.bas')
    with open(theme_path, 'r', encoding='latin-1') as f:
        theme_code = f.read()
    theme_code = theme_code.replace('Public Const LIGHT_', 'Private Const LIGHT_')
    theme_code = theme_code.replace('Public Const DARK_', 'Private Const DARK_')
    with open(theme_path, 'w', encoding='latin-1') as f:
        f.write(theme_code)
    print("Scoped modThemeEngine constants to Private Const.")

if __name__ == '__main__':
    main()
