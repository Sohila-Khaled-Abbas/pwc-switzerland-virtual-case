Attribute VB_Name = "modNavigation"
Option Explicit

' ==============================================================================
' PwC Switzerland Digital Accelerator - View Navigation Controller
' Delivers instantaneous, flicker-free transitions between executive cockpits
' ==============================================================================

Public Sub NavigateToHomePortal()
    NavigateToSheet "00_Home_Portal"
End Sub

Public Sub NavigateToCallCenter()
    NavigateToSheet "03_CallCenter_Cockpit"
End Sub

Public Sub NavigateToRetention()
    NavigateToSheet "04_CustomerRetention_Cockpit"
End Sub

Public Sub NavigateToDiversity()
    NavigateToSheet "05_DiversityInclusion_Cockpit"
End Sub

Public Sub NavigateToDomains()
    NavigateToSheet "01_Business_Domains"
End Sub

Public Sub NavigateToCatalog()
    NavigateToSheet "02_Metadata_&_KPI_Catalog"
End Sub

Private Sub NavigateToSheet(ByVal targetSheetName As String)
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set ws = ThisWorkbook.Worksheets(targetSheetName)
    ws.Visible = xlSheetVisible
    ws.Activate
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ws.Range("A1").Select
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Navigation Error: Unable to locate sheet '" & targetSheetName & "'." & vbCrLf & Err.Description, vbExclamation, "PwC Navigation"
End Sub
