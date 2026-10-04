import win32com.client as win32
import os

wb_path = os.path.abspath('PWC_Switzerland_Virtual_Case.xlsm')
excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(wb_path)

conn = wb.Connections('ThisWorkbookDataModel')
pc = wb.PivotCaches().Create(SourceType=2, SourceData=conn)
pt = wb.Worksheets('Staging_Pivots').PivotTables('pt_Agent')

for cf in pt.CubeFields:
    name = cf.Name
    if any(k in name for k in ['Date', 'Topic', 'Agent', 'Call']):
        print(name)

wb.Close(False)
excel.Quit()
