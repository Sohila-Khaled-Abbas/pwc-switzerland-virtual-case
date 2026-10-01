Attribute VB_Name = "modExportPDF"
Option Explicit

' ==============================================================================
' PwC Switzerland Digital Accelerator — Publication-Grade PDF Generator
' Exports active dashboard canvas formatted for A4 Landscape executive briefing
' ==============================================================================

Public Sub ExportExecutiveReport()
    Dim ws As Worksheet
    Dim exportPath As String
    Dim fileName As String
    
    On Error GoTo ErrorHandler
    
    Set ws = ActiveSheet
    exportPath = ThisWorkbook.Path & "\"
    fileName = exportPath & "PwC_Call_Center_Executive_Report_" & Format(Now, "YYYYMMDD_HHMM") & ".pdf"
    
    modAppState.FreezeAppState
    Application.StatusBar = "Generating publication-grade PDF report..."
    
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintGridlines = False
    End With
    
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=fileName, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
        
    modAppState.RestoreAppState
    MsgBox "Executive Report successfully generated at:" & vbCrLf & fileName, vbInformation, "PwC PDF Publisher"
    Exit Sub

ErrorHandler:
    modAppState.RestoreAppState
    MsgBox "Export Failed: " & Err.Description, vbCritical, "PwC PDF Publisher Error"
End Sub
