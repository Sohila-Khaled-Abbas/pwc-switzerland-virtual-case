import os

def sanitize_file(filepath):
    with open(filepath, 'rb') as fp:
        content = fp.read()
        
    replacements = [
        (b'\xe2\x80\x94', b'--'),
        (b'\xe2\x80\x93', b'-'),
        (b'\xe2\x80\x9c', b'"'),
        (b'\xe2\x80\x9d', b'"'),
        (b'\xe2\x80\x98', b"'"),
        (b'\xe2\x80\x99', b"'"),
        (b'\xe2\x96\xb2', b'▲'),
        (b'\xe2\x96\xbc', b'▼'),
        (b'\xc2\xa0', b' '),
        (b'\xe2\x80\xa2', b'*'),
        (b'\xc2\xb7', b'-'),
    ]
    
    for k, v in replacements:
        content = content.replace(k, v)
        
    # Check any remaining non-ascii
    rem = [b for b in content if b > 127]
    print(f"{os.path.basename(filepath)}: length={len(content)}, non-ascii={len(rem)}")
    return content

if __name__ == '__main__':
    vba_dir = 'vba'
    for f in sorted(os.listdir(vba_dir)):
        if f.endswith('.bas'):
            p = os.path.join(vba_dir, f)
            sanitized = sanitize_file(p)
