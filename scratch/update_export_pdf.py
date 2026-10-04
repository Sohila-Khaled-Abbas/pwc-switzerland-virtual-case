import win32com.client as win32
import os

wb_path = os.path.abspath('PWC_Switzerland_Virtual_Case_FIXED.xlsm')
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(wb_path)
    comp = wb.VBProject.VBComponents('modExportPDF')
    comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
    comp.CodeModule.AddFromFile(os.path.abspath('vba/modExportPDF.bas'))
    print('Updated modExportPDF in VBProject successfully!')
    
    # Test running both
    excel.Run('modExportPDF.ExportActiveDashboardPDF')
    print('SUCCESS: ExportActiveDashboardPDF ran without error!')
    
    wb.Save()
    wb.Close(SaveChanges=True)
except Exception as e:
    print('Error:', e)
finally:
    excel.Quit()
