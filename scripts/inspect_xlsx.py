import zipfile
import re

def inspect_xml_blocks():
    with zipfile.ZipFile('PWC_Switzerland_Virtual_Case.xlsx', 'r') as z:
        data = z.read('xl/model/item.data')
    
    # search for <...
    for m in re.finditer(b'<([A-Za-z0-9_]+)[^>]*>(.*?)</\\1>', data):
        tag = m.group(1).decode(errors='ignore')
        content = m.group(2)[:100].decode(errors='ignore')
        if len(tag) > 3:
            print(f"Tag: <{tag}> -> {content[:60]}")

if __name__ == '__main__':
    inspect_xml_blocks()
