Attribute VB_Name = "modFilterController"
Option Explicit

' ==============================================================================
' PwC Switzerland Digital Accelerator — Slicer State Controller
' Resets all multi-dimensional slicers across the entire workbook
' ==============================================================================

Public Sub ClearAllFilters()
    Dim sc As SlicerCache
    On Error GoTo ErrorHandler
    
    modAppState.FreezeAppState
    Application.StatusBar = "Resetting all interactive dashboard filters..."
    
    For Each sc In ThisWorkbook.SlicerCaches
        sc.ClearManualFilter
    Next sc
    
    modAppState.RestoreAppState
    Application.StatusBar = "All filters successfully reset."
    Exit Sub

ErrorHandler:
    modAppState.RestoreAppState
    MsgBox "Filter Reset Error: " & Err.Description, vbCritical, "PwC Filter Controller"
End Sub
