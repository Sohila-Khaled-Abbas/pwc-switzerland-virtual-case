import win32com.client as win32
import os

wb_path = os.path.abspath('PWC_Switzerland_Virtual_Case.xlsm')
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(wb_path)

cockpits = ['03_CallCenter_Cockpit', '04_CustomerRetention_Cockpit', '05_DiversityInclusion_Cockpit']

for cname in cockpits:
    ws = wb.Worksheets(cname)
    print(f'=== {cname} ===')
    # Print CUBE cells AA65:AE65
    print('  CUBE Formulas & Values (Row 65):')
    for col_letter, col_idx in [('AA', 27), ('AB', 28), ('AC', 29), ('AD', 30), ('AE', 31)]:
        cell = ws.Cells(65, col_idx)
        print(f'    {col_letter}65: Formula={cell.Formula}, Value={cell.Value}, Text={cell.Text}')
    
    # Print KPI shapes
    print('  KPI Value Shapes:')
    for sh in ws.Shapes:
        if any(k in sh.Name.lower() for k in ['ban', 'kpi_val', 'value_']):
            txt = ''
            try:
                txt = sh.TextFrame2.TextRange.Text.replace('\r', ' ').replace('\n', ' ')
            except:
                pass
            print(f'    Shape: {sh.Name}, Text: {txt}')

wb.Close(False)
excel.Quit()
