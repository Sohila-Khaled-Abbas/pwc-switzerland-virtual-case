import win32com.client as win32
import os

wb_path = os.path.abspath('PWC_Switzerland_Virtual_Case_FIXED.xlsm')
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(wb_path)
    ws_cc = wb.Worksheets('03_CallCenter_Cockpit')
    ws_cc.Shapes('Value_CC_ASA').DrawingObject.Formula = '=$AD$65'
    print('Successfully linked Value_CC_ASA -> =$AD$65!')
    wb.Save()
    wb.Close(SaveChanges=True)
    print('Saved successfully!')
except Exception as e:
    print('Error:', e)
finally:
    excel.Quit()
