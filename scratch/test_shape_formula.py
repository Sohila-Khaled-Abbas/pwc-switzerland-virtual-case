import win32com.client as win32
import os

wb_path = os.path.abspath('PWC_Switzerland_Virtual_Case.xlsm')
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(wb_path)
    ws = wb.Worksheets('03_CallCenter_Cockpit')
    
    sh = ws.Shapes('Value_CC_TotalDemand')
    print('Current Text:', sh.TextFrame2.TextRange.Text)
    
    # In Excel COM, setting Formula on DrawingObject:
    sh.DrawingObject.Formula = '=$AA$65'
    print('Set Formula to =$AA$65 successfully!')
    print('New Text after formula link:', sh.TextFrame2.TextRange.Text)
    
    wb.Close(False)
    print('SUCCESS: Verified Shape cell-link formula!')
except Exception as e:
    print('Error:', e)
finally:
    excel.Quit()
