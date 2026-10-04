Attribute VB_Name = "modExportPDF"
Option Explicit

' ==============================================================================
' PwC Switzerland Digital Accelerator -- Publication-Grade PDF Generator
' Exports active dashboard canvas formatted for A4 Landscape executive briefing
' Connected to modAppState and modDashboardUIUX
' ==============================================================================

Public Sub ExportActiveDashboardPDF()
    ExportExecutiveReport
End Sub

Public Sub ExportExecutiveReport()
    Dim ws As Worksheet
    Dim exportPath As String
    Dim cleanSheetName As String
    Dim fileName As String
    
    On Error GoTo ErrorHandler
    
    Set ws = ActiveSheet
    exportPath = ThisWorkbook.Path & "\"
    cleanSheetName = Replace(ws.Name, " ", "_")
    fileName = exportPath & "PwC_" & cleanSheetName & "_Executive_Report_" & Format(Now, "YYYYMMDD_HHMM") & ".pdf"
    
    modAppState.FreezeAppState
    Application.StatusBar = "Generating publication-grade PDF report for " & ws.Name & "..."
    
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintGridlines = False
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
    End With
    
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=fileName, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
        
    modAppState.RestoreAppState
    
    If Application.UserControl Then
        MsgBox "Executive PDF Report successfully generated for '" & ws.Name & "'!" & vbCrLf & vbCrLf & _
               "File saved at:" & vbCrLf & fileName, vbInformation, "PwC PDF Publisher"
    End If
    Exit Sub

ErrorHandler:
    modAppState.RestoreAppState
    If Application.UserControl Then
        MsgBox "Export Failed: " & Err.Description, vbCritical, "PwC PDF Publisher Error"
    End If
End Sub
