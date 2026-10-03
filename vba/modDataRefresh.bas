Attribute VB_Name = "modDataRefresh"
Option Explicit

' ==============================================================================
' PwC Switzerland Digital Accelerator — Synchronous Data Refresh Layer
' Refreshes analytical PivotTables and CUBE calculations in sequence
' Connected to modAppState, modDashboardUIUX, and modFilterController
' ==============================================================================

Public Sub RefreshPipelineSynchronously()
    Dim startTime As Double
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim wsActive As Worksheet
    Dim shStatus As Shape
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    modAppState.FreezeAppState
    Application.StatusBar = "Refreshing VertiPaq Tabular Engine & Analytical Pivots..."
    
    ' 1. Refresh all PivotTables across all worksheets
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            On Error Resume Next
            pt.Update
            On Error GoTo ErrorHandler
        Next pt
    Next ws
    
    ' 2. Recalculate all CUBEVALUE and dashboard formulas
    Application.StatusBar = "Recalculating CUBE metrics..."
    Application.CalculateFull
    
    ' 3. Update Status Indicator in Active Dashboard
    On Error Resume Next
    Set wsActive = ActiveSheet
    If Not wsActive Is Nothing Then
        Set shStatus = wsActive.Shapes("Nav_StatusPill")
        If Not shStatus Is Nothing Then
            shStatus.TextFrame2.TextRange.Text = "[LIVE] REFRESHED: " & Format(Now, "HH:MM")
        End If
    End If
    On Error GoTo ErrorHandler
    
    modAppState.RestoreAppState
    Application.StatusBar = "Analytical engine successfully refreshed."
    
    If Application.UserControl Then
        MsgBox "Data Model PivotTables and KPI metrics successfully refreshed in " & Round(Timer - startTime, 2) & " seconds!", _
               vbInformation, "PwC Data Refresh"
    End If
    Exit Sub

ErrorHandler:
    modAppState.RestoreAppState
    If Application.UserControl Then
        MsgBox "Refresh Failed: " & Err.Description, vbCritical, "PwC Data Refresh Error"
    End If
End Sub
