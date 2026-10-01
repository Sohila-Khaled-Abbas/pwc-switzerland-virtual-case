Attribute VB_Name = "modNavigation"
Option Explicit

' ==============================================================================
' PwC Switzerland Digital Accelerator — View Navigation Controller
' Delivers instantaneous, flicker-free transitions between reporting views
' ==============================================================================

Public Sub NavigateToDashboard()
    On Error GoTo ErrorHandler
    modAppState.FreezeAppState
    
    Sheets("Dashboard").Visible = xlSheetVisible
    Sheets("Dashboard").Activate
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Range("D5").Select
    
    modAppState.RestoreAppState
    Exit Sub

ErrorHandler:
    modAppState.RestoreAppState
    MsgBox "Navigation Error: " & Err.Description, vbExclamation, "PwC Navigation"
End Sub

Public Sub NavigateToAgentDetail()
    On Error GoTo ErrorHandler
    modAppState.FreezeAppState
    
    Sheets("Agent_Detail").Visible = xlSheetVisible
    Sheets("Agent_Detail").Activate
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Range("D5").Select
    
    modAppState.RestoreAppState
    Exit Sub

ErrorHandler:
    modAppState.RestoreAppState
    MsgBox "Navigation Error: " & Err.Description, vbExclamation, "PwC Navigation"
End Sub
