Attribute VB_Name = "modDataRefresh"
Option Explicit

' ==============================================================================
' PwC Switzerland Digital Accelerator — Synchronous Data Refresh Layer
' Refreshes background connections and analytical PivotCaches in sequence
' ==============================================================================

Public Sub RefreshPipelineSynchronously()
    Dim startTime As Double
    Dim conn As WorkbookConnection
    Dim pc As PivotCache
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    modAppState.FreezeAppState
    Application.StatusBar = "Refreshing Power Query ETL pipeline..."
    
    ' 1. Refresh background model connections synchronously
    For Each conn In ThisWorkbook.Connections
        If conn.Type = xlConnectionTypeOLEDB Or conn.Type = xlConnectionTypeODBC Or conn.Type = xlConnectionTypeMODEL Then
            conn.OLEDBConnection.BackgroundQuery = False
            conn.Refresh
        End If
    Next conn
    
    ' 2. Refresh downstream analytical PivotCaches
    Application.StatusBar = "Updating analytical PivotCaches..."
    For Each pc In ThisWorkbook.PivotCaches
        pc.Refresh
    Next pc
    
    ' 3. Log refresh metadata to dashboard header
    On Error Resume Next
    Sheets("Dashboard").Range("W2").Value = "Last Refreshed: " & Format(Now, "YYYY-MM-DD HH:MM")
    On Error GoTo ErrorHandler
    
    modAppState.RestoreAppState
    MsgBox "Pipeline refresh successfully completed in " & Round(Timer - startTime, 2) & " seconds!", vbInformation, "PwC Data Refresh"
    Exit Sub

ErrorHandler:
    modAppState.RestoreAppState
    MsgBox "Refresh Failed: " & Err.Description, vbCritical, "PwC Data Refresh Error"
End Sub
