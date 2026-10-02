Attribute VB_Name = "modFilterController"
Option Explicit

' ==============================================================================
' PwC Switzerland Digital Accelerator — Slicer & Filter State Controller
' Resets all multi-dimensional slicers and table filters across the workbook
' Connected to modAppState and modDashboardUIUX
' ==============================================================================

Public Sub ClearAllFilters()
    Dim sc As SlicerCache
    Dim ws As Worksheet
    Dim pt As PivotTable
    
    On Error GoTo ErrorHandler
    
    modAppState.FreezeAppState
    Application.StatusBar = "Resetting all interactive dashboard filters and slicers..."
    
    ' 1. Clear all Slicer selections in workbook
    For Each sc In ThisWorkbook.SlicerCaches
        On Error Resume Next
        sc.ClearManualFilter
        On Error GoTo ErrorHandler
    Next sc
    
    ' 2. Clear PivotTable filter fields if any were filtered directly
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            On Error Resume Next
            pt.ClearAllFilters
            On Error GoTo ErrorHandler
        Next pt
    Next ws
    
    modAppState.RestoreAppState
    Application.StatusBar = "All filters successfully reset."
    
    If Application.UserControl Then
        MsgBox "All dashboard slicers and filters have been successfully cleared!", _
               vbInformation, "PwC Filter Controller"
    End If
    Exit Sub

ErrorHandler:
    modAppState.RestoreAppState
    If Application.UserControl Then
        MsgBox "Filter Reset Error: " & Err.Description, vbCritical, "PwC Filter Controller"
    End If
End Sub
