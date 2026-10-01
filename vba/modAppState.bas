Attribute VB_Name = "modAppState"
Option Explicit

' ==============================================================================
' PwC Switzerland Digital Accelerator — Application State Shield
' Provides deterministic screen freezing and fail-safe environment restoration
' ==============================================================================

Private m_OriginalScreenUpdating As Boolean
Private m_OriginalEnableEvents   As Boolean
Private m_OriginalCalculation    As XlCalculation
Private m_OriginalDisplayAlerts  As Boolean
Private m_IsFrozen               As Boolean

Public Sub FreezeAppState(Optional ByVal ManualCalc As Boolean = False)
    On Error Resume Next
    If Not m_IsFrozen Then
        m_OriginalScreenUpdating = Application.ScreenUpdating
        m_OriginalEnableEvents = Application.EnableEvents
        m_OriginalCalculation = Application.Calculation
        m_OriginalDisplayAlerts = Application.DisplayAlerts
        
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        If ManualCalc Then Application.Calculation = xlCalculationManual
        Application.DisplayAlerts = False
        m_IsFrozen = True
    End If
End Sub

Public Sub RestoreAppState()
    On Error Resume Next
    If m_IsFrozen Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = m_OriginalCalculation
        Application.DisplayAlerts = True
        Application.StatusBar = False
        m_IsFrozen = False
    End If
End Sub
