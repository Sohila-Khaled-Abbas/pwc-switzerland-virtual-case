import zipfile
import xml.etree.ElementTree as ET

wb_path = r'd:\courses\Data Analysis 26-27\7-Introducation to Data Fields (Excel)\11_Demos_and_Workbooks\10_Projects_and_Demos\PWC\PWC_Switzerland_Virtual_Case.xlsm'

with zipfile.ZipFile(wb_path, 'r') as z:
    wb_xml = z.read('xl/workbook.xml')
    root = ET.fromstring(wb_xml)
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    sheets = root.findall('main:sheets/main:sheet', ns)
    print('Sheets in workbook:')
    for s in sheets:
        print(f"  Id: {s.attrib.get('sheetId')}, Name: {s.attrib.get('name')}, r:id: {s.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')}")

    rels_xml = z.read('xl/_rels/workbook.xml.rels')
    rels_root = ET.fromstring(rels_xml)
    rel_map = {r.attrib['Id']: r.attrib['Target'] for r in rels_root}

    for s in sheets:
        s_name = s.attrib.get('name')
        rid = s.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        target = 'xl/' + rel_map[rid]
        sheet_xml = z.read(target)
        s_root = ET.fromstring(sheet_xml)
        
        sheetViews = s_root.findall('.//main:sheetView', ns)
        views_info = []
        for v in sheetViews:
            top_left = v.attrib.get('topLeftCell', 'A1')
            zoom = v.attrib.get('zoomScale', '100')
            pane = v.find('main:pane', ns)
            pane_info = ''
            if pane is not None:
                pane_info = f" Pane(xSplit={pane.attrib.get('xSplit')}, ySplit={pane.attrib.get('ySplit')}, topLeftCell={pane.attrib.get('topLeftCell')})"
            views_info.append(f"zoom={zoom}%, topLeft={top_left}{pane_info}")
            
        # Check drawings
        drawings = s_root.findall('.//main:drawing', ns)
        draw_rids = [d.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id') for d in drawings]
        
        print(f'Sheet "{s_name}" ({target}): {views_info}, drawings={draw_rids}')
