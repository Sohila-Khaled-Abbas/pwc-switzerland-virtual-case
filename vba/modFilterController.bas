Attribute VB_Name = "modFilterController"
Option Explicit

' ==============================================================================
' PwC Switzerland Digital Accelerator -- Slicer & Filter State Controller
' Module: modFilterController
' Creates, docks, styles, and resets multi-dimensional VertiPaq Data Model
' Slicers across all three analytical cockpits:
'   - 03_CallCenter_Cockpit        : Month, Inquiry Topic, Representative Agent
'   - 04_CustomerRetention_Cockpit : Contract, Payment Method, Internet Service
'   - 05_DiversityInclusion_Cockpit: Department, Job Level Hierarchy, Gender
' Pure 7-bit ASCII encoding.
' ==============================================================================

Private Const SHEET_CC As String = "03_CallCenter_Cockpit"
Private Const SHEET_RETENTION As String = "04_CustomerRetention_Cockpit"
Private Const SHEET_DIVERSITY As String = "05_DiversityInclusion_Cockpit"
Private Const SHEET_STAGING_PIVOTS As String = "Staging_Pivots"

' ==============================================================================
' 1. Slicer Provisioning & Docking Engine
' ==============================================================================
Public Function AddDataModelSlicer(wsTarget As Worksheet, ptSource As PivotTable, _
                                   ByVal primaryField As String, _
                                   ByVal scName As String, ByVal slName As String, ByVal captionText As String, _
                                   ByVal leftPos As Single, ByVal topPos As Single, _
                                   ByVal w As Single, ByVal h As Single) As Slicer
    On Error Resume Next
    Dim wb As Workbook: Set wb = wsTarget.Parent
    Dim sc As SlicerCache
    Dim sl As Slicer
    Dim existingSl As Slicer
    
    ' 1. Locate existing SlicerCache or provision from Data Model PivotTable
    Set sc = wb.SlicerCaches(scName)
    If sc Is Nothing Then
        Set sc = wb.SlicerCaches.Add2(ptSource, primaryField, scName)
    End If
    
    If sc Is Nothing Then Exit Function
    
    ' Ensure Source PivotTable is connected to this cache
    If Not ptSource Is Nothing Then
        On Error Resume Next
        sc.PivotTables.AddPivotTable ptSource
        On Error GoTo 0
    End If
    
    ' 2. Check if Slicer visual already exists in SlicerCache
    Set existingSl = Nothing
    On Error Resume Next
    Set existingSl = sc.Slicers(slName)
    On Error GoTo 0
    
    If Not existingSl Is Nothing Then
        Set sl = existingSl
        With sl
            .Top = topPos
            .Left = leftPos
            .Width = w
            .Height = h
            .Caption = captionText
            .NumberOfColumns = 1
            .Style = "SlicerStyleLight2"
        End With
        Set AddDataModelSlicer = sl
        Exit Function
    End If
    
    ' 3. Add Slicer visual object to target worksheet inside sidebar slot
    On Error Resume Next
    Set sl = sc.Slicers.Add(SlicerDestination:=wsTarget, Name:=slName, Caption:=captionText, _
                            Top:=topPos, Left:=leftPos, Width:=w, Height:=h)
    
    If Not sl Is Nothing Then
        With sl
            .Top = topPos
            .Left = leftPos
            .Width = w
            .Height = h
            .Caption = captionText
            .NumberOfColumns = 1
            .Style = "SlicerStyleLight2"
        End With
    End If
    
    Set AddDataModelSlicer = sl
End Function

' ==============================================================================
' 2. Cockpit-Specific Slicer Deployment Engines
' ==============================================================================
Public Sub DeployCallCenterSlicers(ws As Worksheet, Optional ByVal ptSource As PivotTable = Nothing)
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    
    If ptSource Is Nothing Then
        Dim wsStg As Worksheet
        Set wsStg = ws.Parent.Worksheets(SHEET_STAGING_PIVOTS)
        If wsStg Is Nothing Then
            Set wsStg = ws.Parent.Worksheets.Add(After:=ws.Parent.Worksheets(ws.Parent.Worksheets.Count))
            wsStg.Name = SHEET_STAGING_PIVOTS
            wsStg.Visible = xlSheetHidden
        End If
        Set ptSource = modInteractiveScorecard.EnsureOrBuildAgentPivotTable(wsStg)
    End If
    If ptSource Is Nothing Then Exit Sub
    
    ' Slot 1: Billing / Calendar Month
    AddDataModelSlicer ws, ptSource, "[DimDate].[Month]", _
                       "sc_CC_Month", "sl_CC_Month", "Billing Month", 26, 172, 196, 140
    modDashboardUIUX.SafeDeleteShape ws, "CC_Slicers_Slot_1"
    
    ' Slot 2: Inquiry Topic Tier
    AddDataModelSlicer ws, ptSource, "[DimTopic].[Topic]", _
                       "sc_CC_Topic", "sl_CC_Topic", "Inquiry Topic Tier", 26, 322, 196, 140
    modDashboardUIUX.SafeDeleteShape ws, "CC_Slicers_Slot_2"
    
    ' Slot 3: Representative Agent
    AddDataModelSlicer ws, ptSource, "[DimAgent].[Agent]", _
                       "sc_CC_Agent", "sl_CC_Agent", "Representative Agent", 26, 472, 196, 150
    modDashboardUIUX.SafeDeleteShape ws, "CC_Slicers_Slot_3"
End Sub

Public Sub DeployRetentionSlicers(ws As Worksheet, Optional ByVal ptSource As PivotTable = Nothing)
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    
    If ptSource Is Nothing Then
        Dim wsStg As Worksheet
        Set wsStg = ws.Parent.Worksheets(SHEET_STAGING_PIVOTS)
        If wsStg Is Nothing Then
            Set wsStg = ws.Parent.Worksheets.Add(After:=ws.Parent.Worksheets(ws.Parent.Worksheets.Count))
            wsStg.Name = SHEET_STAGING_PIVOTS
            wsStg.Visible = xlSheetHidden
        End If
        Set ptSource = modInteractiveScorecard.EnsureOrBuildAgentPivotTable(wsStg)
    End If
    If ptSource Is Nothing Then Exit Sub
    
    ' Slot 1: Contract Architecture
    AddDataModelSlicer ws, ptSource, "[DimContract].[Contract]", _
                       "sc_CR_Contract", "sl_CR_Contract", "Contract Architecture", 26, 172, 196, 140
    modDashboardUIUX.SafeDeleteShape ws, "CR_Slicers_Slot_1"
    
    ' Slot 2: Payment Gateway Friction
    AddDataModelSlicer ws, ptSource, "[Fact_Churn].[PaymentMethod]", _
                       "sc_CR_Payment", "sl_CR_Payment", "Payment Gateway Friction", 26, 322, 196, 140
    modDashboardUIUX.SafeDeleteShape ws, "CR_Slicers_Slot_2"
    
    ' Slot 3: Internet Service Modality
    AddDataModelSlicer ws, ptSource, "[Fact_Churn].[InternetService]", _
                       "sc_CR_Internet", "sl_CR_Internet", "Internet Service Modality", 26, 472, 196, 150
    modDashboardUIUX.SafeDeleteShape ws, "CR_Slicers_Slot_3"
End Sub

Public Sub DeployDiversitySlicers(ws As Worksheet, Optional ByVal ptSource As PivotTable = Nothing)
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    
    If ptSource Is Nothing Then
        Dim wsStg As Worksheet
        Set wsStg = ws.Parent.Worksheets(SHEET_STAGING_PIVOTS)
        If wsStg Is Nothing Then
            Set wsStg = ws.Parent.Worksheets.Add(After:=ws.Parent.Worksheets(ws.Parent.Worksheets.Count))
            wsStg.Name = SHEET_STAGING_PIVOTS
            wsStg.Visible = xlSheetHidden
        End If
        Set ptSource = modInteractiveScorecard.EnsureOrBuildAgentPivotTable(wsStg)
    End If
    If ptSource Is Nothing Then Exit Sub
    
    ' Slot 1: Corporate Department
    AddDataModelSlicer ws, ptSource, "[DimDepartment].[Department]", _
                       "sc_DI_Dept", "sl_DI_Dept", "Corporate Department", 26, 172, 196, 140
    modDashboardUIUX.SafeDeleteShape ws, "DI_Slicers_Slot_1"
    
    ' Slot 2: Job Level Hierarchy
    AddDataModelSlicer ws, ptSource, "[Fact_Employees].[Job_Level_Baseline]", _
                       "sc_DI_JobLevel", "sl_DI_JobLevel", "Job Level Hierarchy", 26, 322, 196, 140
    modDashboardUIUX.SafeDeleteShape ws, "DI_Slicers_Slot_2"
    
    ' Slot 3: Gender Parity Demographics
    AddDataModelSlicer ws, ptSource, "[Fact_Employees].[Gender]", _
                       "sc_DI_Gender", "sl_DI_Gender", "Gender Demographics", 26, 472, 196, 150
    modDashboardUIUX.SafeDeleteShape ws, "DI_Slicers_Slot_3"
End Sub

' ==============================================================================
' 3. Master Deployer for All Slicers
' ==============================================================================
Public Sub DeployAllCockpitSlicers()
    On Error Resume Next
    Dim wb As Workbook: Set wb = ThisWorkbook: If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim wsStg As Worksheet, pt As PivotTable
    
    Set wsStg = wb.Worksheets(SHEET_STAGING_PIVOTS)
    If wsStg Is Nothing Then
        Set wsStg = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsStg.Name = SHEET_STAGING_PIVOTS
        wsStg.Visible = xlSheetHidden
    End If
    
    Set pt = modInteractiveScorecard.EnsureOrBuildAgentPivotTable(wsStg)
    If pt Is Nothing Then Exit Sub
    
    Dim wsCC As Worksheet, wsCR As Worksheet, wsDI As Worksheet
    Set wsCC = wb.Worksheets(SHEET_CC)
    Set wsCR = wb.Worksheets(SHEET_RETENTION)
    Set wsDI = wb.Worksheets(SHEET_DIVERSITY)
    
    If Not wsCC Is Nothing Then Call DeployCallCenterSlicers(wsCC, pt)
    If Not wsCR Is Nothing Then Call DeployRetentionSlicers(wsCR, pt)
    If Not wsDI Is Nothing Then Call DeployDiversitySlicers(wsDI, pt)
End Sub

' ==============================================================================
' 4. Filter & Slicer State Reset
' ==============================================================================
Public Sub ClearAllFilters()
    Dim sc As SlicerCache
    Dim ws As Worksheet
    Dim pt As PivotTable
    
    On Error Resume Next
    modAppState.FreezeAppState
    Application.StatusBar = "Resetting all interactive dashboard filters and slicers..."
    
    ' 1. Clear all Slicer selections across workbook
    For Each sc In ThisWorkbook.SlicerCaches
        sc.ClearManualFilter
    Next sc
    
    ' 2. Clear PivotTable filter fields if any were filtered directly
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.ClearAllFilters
        Next pt
    Next ws
    
    modAppState.RestoreAppState
    Application.StatusBar = "All filters successfully reset."
    
    If Application.UserControl Then
        MsgBox "All dashboard slicers and filters have been successfully cleared!", _
               vbInformation, "PwC Filter Controller"
    End If
End Sub
