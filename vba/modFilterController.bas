Attribute VB_Name = "modFilterController"
Option Explicit

' ==============================================================================
' PwC Switzerland Digital Accelerator -- Slicer & Filter State Controller
' Creates, docks and resets multi-dimensional VertiPaq Slicers across cockpits
' Connected to modAppState, modDashboardUIUX, and modInteractiveScorecard
' ==============================================================================

Private Const SHEET_CC As String = "03_CallCenter_Cockpit"
Private Const SHEET_RETENTION As String = "04_CustomerRetention_Cockpit"
Private Const SHEET_DIVERSITY As String = "05_DiversityInclusion_Cockpit"
Private Const SHEET_STAGING_PIVOTS As String = "Staging_Pivots"

' ==============================================================================
' 1. Slicer Provisioning & Docking Helper
' ==============================================================================
Public Function AddDataModelSlicer(wsTarget As Worksheet, ptSource As PivotTable, _
                                   ByVal primaryField As String, ByVal fallbackField As String, _
                                   ByVal scName As String, ByVal slName As String, ByVal captionText As String, _
                                   ByVal leftPos As Single, ByVal topPos As Single, _
                                   ByVal w As Single, ByVal h As Single) As Slicer
    On Error Resume Next
    Dim wb As Workbook: Set wb = wsTarget.Parent
    Dim sc As SlicerCache
    Dim sl As Slicer
    
    ' 1. Locate or create the SlicerCache from the Data Model PivotTable
    Set sc = wb.SlicerCaches(scName)
    If sc Is Nothing Then
        Set sc = wb.SlicerCaches.Add2(Source:=ptSource, SourceField:=primaryField, Name:=scName)
        If sc Is Nothing And Len(fallbackField) > 0 Then
            Set sc = wb.SlicerCaches.Add2(Source:=ptSource, SourceField:=fallbackField, Name:=scName)
        End If
    End If
    
    If sc Is Nothing Then Exit Function
    
    ' 2. Connect ptSource if not already connected
    On Error Resume Next
    sc.PivotTables.AddPivotTable ptSource
    
    ' 3. Check if Slicer already exists on destination sheet
    Dim shp As Shape
    For Each shp In wsTarget.Shapes
        If shp.Type = msoSlicer Then
            If shp.Name = slName Then
                Set sl = shp.Slicer
                Exit For
            End If
        End If
    Next shp
    
    ' 4. Add or position the Slicer visual object
    If sl Is Nothing Then
        Set sl = sc.Slicers.Add(SlicerDestination:=wsTarget, Name:=slName, Caption:=captionText, _
                                Top:=topPos, Left:=leftPos, Width:=w, Height:=h)
    Else
        sl.Top = topPos
        sl.Left = leftPos
        sl.Width = w
        sl.Height = h
    End If
    
    ' 5. Apply PwC Executive Visual Slicer Style
    If Not sl Is Nothing Then
        sl.Caption = captionText
        sl.NumberOfColumns = 1
        On Error Resume Next
        sl.Style = "SlicerStyleLight2"
    End If
    
    Set AddDataModelSlicer = sl
End Function

' ==============================================================================
' 2. Cockpit-Specific Slicer Deployment Engines
' ==============================================================================
Public Sub DeployCallCenterSlicers(ws As Worksheet, ptSource As PivotTable)
    On Error Resume Next
    If ws Is Nothing Or ptSource Is Nothing Then Exit Sub
    
    ' Slot 1: Billing / Calendar Month
    AddDataModelSlicer ws, ptSource, "[DimDate].[Month].[Month]", "[DimDate].[Month]", _
                       "sc_CC_Month", "sl_CC_Month", "Billing Month", 46, 170, 176, 138
    modDashboardUIUX.SafeDeleteShape ws, "CC_Slicers_Slot_1"
    
    ' Slot 2: Inquiry Topic
    AddDataModelSlicer ws, ptSource, "[DimTopic].[Topic].[Topic]", "[DimTopic].[Topic]", _
                       "sc_CC_Topic", "sl_CC_Topic", "Inquiry Topic Tier", 46, 318, 176, 138
    modDashboardUIUX.SafeDeleteShape ws, "CC_Slicers_Slot_2"
    
    ' Slot 3: Representative Agent
    AddDataModelSlicer ws, ptSource, "[DimAgent].[Agent].[Agent]", "[DimAgent].[Agent]", _
                       "sc_CC_Agent", "sl_CC_Agent", "Representative Agent", 46, 466, 176, 148
    modDashboardUIUX.SafeDeleteShape ws, "CC_Slicers_Slot_3"
End Sub

Public Sub DeployRetentionSlicers(ws As Worksheet, ptSource As PivotTable)
    On Error Resume Next
    If ws Is Nothing Or ptSource Is Nothing Then Exit Sub
    
    ' Slot 1: Contract Architecture
    AddDataModelSlicer ws, ptSource, "[DimContract].[Contract].[Contract]", "[DimContract].[Contract]", _
                       "sc_CR_Contract", "sl_CR_Contract", "Contract Architecture", 46, 170, 176, 138
    modDashboardUIUX.SafeDeleteShape ws, "CR_Slicers_Slot_1"
    
    ' Slot 2: Payment Gateway Friction
    AddDataModelSlicer ws, ptSource, "[Fact_Churn].[PaymentMethod].[PaymentMethod]", "[Fact_Churn].[PaymentMethod]", _
                       "sc_CR_Payment", "sl_CR_Payment", "Payment Gateway Friction", 46, 318, 176, 138
    modDashboardUIUX.SafeDeleteShape ws, "CR_Slicers_Slot_2"
    
    ' Slot 3: Internet Service Modality
    AddDataModelSlicer ws, ptSource, "[Fact_Churn].[InternetService].[InternetService]", "[Fact_Churn].[InternetService]", _
                       "sc_CR_Internet", "sl_CR_Internet", "Internet Service Modality", 46, 466, 176, 148
    modDashboardUIUX.SafeDeleteShape ws, "CR_Slicers_Slot_3"
End Sub

Public Sub DeployDiversitySlicers(ws As Worksheet, ptSource As PivotTable)
    On Error Resume Next
    If ws Is Nothing Or ptSource Is Nothing Then Exit Sub
    
    ' Slot 1: Corporate Department
    AddDataModelSlicer ws, ptSource, "[DimDepartment].[Department].[Department]", "[DimDepartment].[Department]", _
                       "sc_DI_Dept", "sl_DI_Dept", "Corporate Department", 46, 170, 176, 138
    modDashboardUIUX.SafeDeleteShape ws, "DI_Slicers_Slot_1"
    
    ' Slot 2: Job Level Hierarchy
    AddDataModelSlicer ws, ptSource, "[Fact_Employees].[JobLevel].[JobLevel]", "[Fact_Employees].[JobLevel]", _
                       "sc_DI_JobLevel", "sl_DI_JobLevel", "Job Level Hierarchy", 46, 318, 176, 138
    modDashboardUIUX.SafeDeleteShape ws, "DI_Slicers_Slot_2"
    
    ' Slot 3: Gender Parity
    AddDataModelSlicer ws, ptSource, "[Fact_Employees].[Gender].[Gender]", "[Fact_Employees].[Gender]", _
                       "sc_DI_Gender", "sl_DI_Gender", "Gender Demographics", 46, 466, 176, 148
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
    
    ' 1. Clear all Slicer selections in workbook
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
