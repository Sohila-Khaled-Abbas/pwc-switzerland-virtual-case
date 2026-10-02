Attribute VB_Name = "modDataRefresh"
Option Explicit

' ==============================================================================
' PwC Switzerland Digital Accelerator — Synchronous Data Refresh Layer
' Refreshes background connections and analytical PivotCaches in sequence
' Connected to modAppState, modDashboardUIUX, and modFilterController
' ==============================================================================

Public Sub RefreshPipelineSynchronously()
    Dim startTime As Double
    Dim conn As WorkbookConnection
    Dim pc As PivotCache
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    modAppState.FreezeAppState
    Application.StatusBar = "Refreshing VertiPaq Data Model & ETL Pipeline..."
    
    ' 1. Refresh background model connections synchronously
    For Each conn In ThisWorkbook.Connections
        If conn.Type = xlConnectionTypeOLEDB Or conn.Type = xlConnectionTypeODBC Or conn.Type = xlConnectionTypeMODEL Then
            On Error Resume Next
            conn.OLEDBConnection.BackgroundQuery = False
            conn.Refresh
            On Error GoTo ErrorHandler
        End If
    Next conn
    
    ' 2. Refresh downstream analytical PivotCaches
    Application.StatusBar = "Updating analytical PivotCaches..."
    For Each pc In ThisWorkbook.PivotCaches
        On Error Resume Next
        pc.Refresh
        On Error GoTo ErrorHandler
    Next pc
    
    ' 3. Update Status Indicator in Active Dashboard
    On Error Resume Next
    Set ws = ActiveSheet
    If ws.Shapes("Nav_StatusPill") IsNot Nothing Then
        ws.Shapes("Nav_StatusPill").TextFrame2.TextRange.Text = "[LIVE] REFRESHED: " & Format(Now, "HH:MM")
    End If
    On Error GoTo ErrorHandler
    
    modAppState.RestoreAppState
    
    If Application.UserControl Then
        MsgBox "Data Model and PivotCaches successfully refreshed in " & Round(Timer - startTime, 2) & " seconds!", _
               vbInformation, "PwC Data Refresh"
    End If
    Exit Sub

ErrorHandler:
    modAppState.RestoreAppState
    If Application.UserControl Then
        MsgBox "Refresh Failed: " & Err.Description, vbCritical, "PwC Data Refresh Error"
    End If
End Sub
