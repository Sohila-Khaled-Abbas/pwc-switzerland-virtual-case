import win32com.client as win32
import os

wb_path = os.path.abspath('PWC_Switzerland_Virtual_Case_FIXED.xlsm')
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(wb_path)

ws_cc = wb.Worksheets('03_CallCenter_Cockpit')
for sh in ws_cc.Shapes:
    sh_l = sh.Name.lower()
    if 'speed' in sh_l or 'value' in sh_l:
        txt = sh.TextFrame2.TextRange.Text if sh.TextFrame2.HasText else ''
        print(f'Shape: {sh.Name}, Text: {txt}')

wb.Close(False)
excel.Quit()
