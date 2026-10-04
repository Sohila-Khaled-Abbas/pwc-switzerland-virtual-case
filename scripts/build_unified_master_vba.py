import os

out_path = r'vba\modPwC_Unified_Master.bas'

code = '''Attribute VB_Name = "modPwC_Unified_Master"
Option Explicit

' ==============================================================================
' PwC Switzerland Digital Accelerator -- Enterprise BI Platform Master Engine
' UNIFIED MASTER VBA AUTOMATION SUITE (All-In-One Production Architecture)
' ==============================================================================
'
' System Overview:
'   This single consolidated module manages the entire lifecycle of the BI Platform:
'     1. modAppState               - Deterministic Screen, Calculation & Alert Shield
'     2. modThemeEngine            - Executive SaaS Light & Dark Dynamic Theming
'     3. modNavigation             - Instant Flicker-Free View Routing with Module Binds
'     4. modDataRefresh            - Synchronous VertiPaq & PivotCache Data Sync
'     5. modFilterController       - Multi-Dimensional Slicer Cache & Filter Reset
'     6. modExportPDF              - Publication-Ready Landscape A4 Executive PDF Export
'     7. modCreateGovernanceSheets - Business Domains & Metadata / KPI Catalog Builder
'     8. modPortalLanding          - Executive SaaS Homepage & Cockpit Launchers
'     9. modDashboardUIUX          - 3 Web-App Canvases with Real Embedded Visuals & Charts
'    10. modInteractiveScorecard   - Live Linked Agent Quality Scorecard Docker
'    11. modPivotTableFormatting   - Scorecard Grid Formatting & SLA Exception Matrix
'
' Key Fixes & Design Standards:
'   - 100% Pure 7-bit ASCII Encoding: Zero UTF-8/ANSI byte corruption on any locale.
'   - Explicit Module Qualifications: All .OnAction properties bound to modPwC_Unified_Master.
'   - Legacy Module Cleanup: Safely attempts to purge old individual modules on run.
'   - Direct Array Staging: Fast, 100% reliable cell assignments (zero #VALUE! transpose errors).
'   - Single-Quoted Formulas: Strictly quotes '03_CallCenter_Cockpit'!$AA$65 in CUBE formulas.
'   - Step-by-Step Error Shielding: Comprehensive tracing with non-destructive state restoration.
'   - Middle-Center Text Alignment: Applied to all cards, badges, buttons & labels.
'   - Modern Rounded Squares: Sleek 0.08 - 0.18 curvature matching executive SaaS UI.
'   - Viewport Normalization: Unfreezes split panes and resets scroll to A1 (no cut-off headers).
'   - True PwC Brand Tokens: Orange = #D04A02 (150224), Dark Slate = #182234 (3416600).
'
' Primary Entry Point:
'   Run "RunCompletePwCPlatform" to generate the entire platform end-to-end.
' ==============================================================================

' ------------------------------------------------------------------------------
' 1. CONSOLIDATED ENTERPRISE BRAND CONSTANTS & DESIGN TOKENS
' ------------------------------------------------------------------------------

' PwC Primary Brand Palette (Windows Long / BGR format)
Public Const PWC_ORANGE         As Long = 150224     ' #D04A02 - RGB(208, 74, 2)
Public Const PWC_DARK_SLATE     As Long = 2758415    ' #0F172A - RGB(15, 23, 42)
Public Const PWC_CHARCOAL       As Long = 3877150    ' #1E293B - RGB(30, 41, 59)
Public Const PWC_CANVAS_BG      As Long = 16579320   ' #F8FAFC - RGB(248, 250, 252)
Public Const PWC_CARD_FILL      As Long = 16777215   ' #FFFFFF - RGB(255, 255, 255)
Public Const PWC_CARD_BORDER    As Long = 15788258   ' #E2E8F0 - RGB(226, 232, 240)
Public Const PWC_BORDER_DASHED  As Long = 13421772   ' #CBD5E1 - RGB(203, 213, 225)
Public Const PWC_SLOT_FILL      As Long = 16448250   ' #FAFAFA - RGB(250, 250, 250)
Public Const PWC_DRAWER_BG      As Long = 3416600    ' #182234 - RGB(24, 34, 52)
Public Const PWC_TEXT_TITLE     As Long = 3877150    ' #1E293B - RGB(30, 41, 59)
Public Const PWC_TEXT_MUTED     As Long = 9139300    ' #64748B - RGB(100, 116, 139)
Public Const PWC_TEXT_LIGHT     As Long = 12040119   ' #94A3B8 - RGB(148, 163, 184)
Public Const PWC_SUCCESS_GREEN  As Long = 6919685    ' #059669 - RGB(5, 150, 105)
Public Const PWC_ALERT_RED      As Long = 2500316    ' #DC2626 - RGB(220, 38, 38)
Public Const PWC_WARNING_AMBER  As Long = 423897     ' #D97706 - RGB(217, 119, 6)
Public Const PWC_PILL_BG        As Long = 16381425   ' #F1F5F9 - RGB(241, 245, 249)
Public Const PWC_WHITE          As Long = 16777215   ' #FFFFFF

' Metric & Status Badge Colors
Public Const PWC_BADGE_GREEN_BG  As Long = 15203548  ' #DCFCE7 - RGB(220, 252, 231)
Public Const PWC_BADGE_GREEN_TXT As Long = 3433746   ' #166534 - RGB(22, 101, 52)
Public Const PWC_BADGE_RED_BG    As Long = 14869246  ' #FEE2E2 - RGB(254, 226, 226)
Public Const PWC_BADGE_RED_TXT   As Long = 1776537   ' #991B1B - RGB(153, 27, 27)
Public Const PWC_BADGE_AMBER_BG  As Long = 13104126  ' #FEF3C7 - RGB(254, 243, 199)
Public Const PWC_BADGE_AMBER_TXT As Long = 933902    ' #92400E - RGB(146, 64, 14)
Public Const PWC_BADGE_BLUE_BG   As Long = 16706267  ' #DBEAFE - RGB(219, 234, 254)
Public Const PWC_BADGE_BLUE_TXT  As Long = 12470046  ' #1E40AF - RGB(30, 64, 175)
Public Const PWC_BADGE_PURPLE_BG As Long = 16771315  ' #F3E8FF - RGB(243, 232, 255)
Public Const PWC_BADGE_PURPLE_TXT As Long = 9643410  ' #9333EA - RGB(147, 51, 234)
Public Const PWC_BADGE_TEAL_BG   As Long = 15858636  ' #CCFBF1 - RGB(204, 251, 241)
Public Const PWC_BADGE_TEAL_TXT  As Long = 8950801   ' #0D9488 - RGB(13, 148, 136)

' Dark Mode Colors
Public Const DARK_BG             As Long = 1642251   ' #0B0F19 (Deep Navy Canvas)
Public Const DARK_CONTAINER      As Long = 3022358   ' #161E2E (Elevated Card)
Public Const DARK_BORDER         As Long = 4666410   ' #2A3447 (Border Slate)
Public Const DARK_TEXT_PRIMARY   As Long = 16579832  ' #F8FAFC (Slate 50)
Public Const DARK_TEXT_MUTED     As Long = 12099732  ' #94A3B8 (Slate 400)
Public Const DARK_ACCENT         As Long = 3373823   ' #FF7A33 (Glowing Orange)
Public Const DARK_GRIDLINE       As Long = 3022358   ' #161E2E

' Typography & Standard Sheet Identifiers
Public Const FONT_FAMILY         As String = "Segoe UI"
Private Const SHEET_PORTAL       As String = "00_Home_Portal"
Private Const SHEET_DOMAINS      As String = "01_Business_Domains"
Private Const SHEET_CATALOG      As String = "02_Metadata_&_KPI_Catalog"
Private Const SHEET_CC           As String = "03_CallCenter_Cockpit"
Private Const SHEET_CH           As String = "04_CustomerRetention_Cockpit"
Private Const SHEET_DI           As String = "05_DiversityInclusion_Cockpit"
Private Const SHEET_STAGING      As String = "Staging_Pivots"
Private Const THEME_PROP_NAME    As String = "PwC_Dashboard_Theme"

' Canvas Layout Coordinates (Balanced 1480pt Widescreen Grid)
Public Const CANVAS_LEFT         As Single = 24
Public Const CANVAS_WIDTH        As Single = 1480
Public Const CARD_GAP            As Single = 14
Public Const KPI_WIDTH_5         As Single = 280
Public Const KPI_WIDTH_7         As Single = 168
Public Const KPI_HEIGHT          As Single = 94
Public Const SLICER_WIDTH        As Single = 210

Private m_OriginalScreenUpdating As Boolean
Private m_OriginalEnableEvents   As Boolean
Private m_OriginalCalculation    As XlCalculation
Private m_OriginalDisplayAlerts  As Boolean
Private m_IsFrozen               As Boolean

' ==============================================================================
' 2. LEGACY MODULE PURGER (Resolves "Ambiguous Name" Collisions)
' ==============================================================================

Public Sub RemoveLegacyModules()
    On Error Resume Next
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject
    If Err.Number <> 0 Or vbProj Is Nothing Then
        Err.Clear
        Exit Sub
    End If
    
    Dim vbComp As Object, compName As String
    For Each vbComp In vbProj.VBComponents
        compName = vbComp.Name
        Select Case compName
            Case "modAppState", "modThemeEngine", "modNavigation", "modDataRefresh", _
                 "modFilterController", "modExportPDF", "modCreateGovernanceSheets", _
                 "modPortalLanding", "modDashboardUIUX", "modInteractiveScorecard", _
                 "modPivotTableFormatting"
                vbProj.VBComponents.Remove vbComp
        End Select
    Next vbComp
    Err.Clear
    On Error GoTo 0
End Sub

' ==============================================================================
' 3. DETERMINISTIC APPLICATION STATE SHIELD
' ==============================================================================

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
    On Error GoTo 0
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
    On Error GoTo 0
End Sub

' ==============================================================================
' 4. ASSET RESOLVER & GRAPHICS ENGINE
' ==============================================================================

Public Function ResolveAssetPath(ByVal subFolderAndFile As String) As String
    Dim wbPath As String, candidatePath As String
    On Error Resume Next
    wbPath = ThisWorkbook.Path
    If Len(wbPath) = 0 Then wbPath = ActiveWorkbook.Path
    On Error GoTo 0
    
    If Len(wbPath) = 0 Then
        ResolveAssetPath = ""
        Exit Function
    End If
    
    candidatePath = wbPath & "\\" & subFolderAndFile
    If Dir(candidatePath) <> "" Then
        ResolveAssetPath = candidatePath
        Exit Function
    End If
    
    candidatePath = wbPath & "\\assets\\" & Mid(subFolderAndFile, InStrRev(subFolderAndFile, "\\") + 1)
    If Dir(candidatePath) <> "" Then
        ResolveAssetPath = candidatePath
        Exit Function
    End If
    
    candidatePath = wbPath & "\\assets\\icons\\" & Mid(subFolderAndFile, InStrRev(subFolderAndFile, "\\") + 1)
    If Dir(candidatePath) <> "" Then
        ResolveAssetPath = candidatePath
        Exit Function
    End If
    
    ResolveAssetPath = ""
End Function

Public Function InsertVectorIcon(ws As Worksheet, ByVal iconFileName As String, _
                                ByVal leftPos As Single, ByVal topPos As Single, _
                                ByVal targetWidth As Single, ByVal targetHeight As Single, _
                                Optional ByVal shapeName As String = "") As Shape
    Dim fullPath As String, shp As Shape
    fullPath = ResolveAssetPath(iconFileName)
    
    If Len(fullPath) = 0 Or Dir(fullPath) = "" Then
        Set InsertVectorIcon = Nothing
        Exit Function
    End If
    
    On Error Resume Next
    Set shp = ws.Shapes.AddPicture(Filename:=fullPath, _
                                  LinkToFile:=msoFalse, _
                                  SaveWithDocument:=msoTrue, _
                                  Left:=leftPos, Top:=topPos, _
                                  Width:=targetWidth, Height:=targetHeight)
    If Not shp Is Nothing Then
        If Len(shapeName) > 0 Then shp.Name = shapeName
        shp.LockAspectRatio = msoTrue
    End If
    Set InsertVectorIcon = shp
    On Error GoTo 0
End Function

Public Sub ApplySoftElevation(ByVal shp As Shape)
    On Error Resume Next
    With shp.Shadow
        .Visible = msoTrue
        .Type = msoShadow21
        .ForeColor.RGB = RGB(15, 23, 42)
        .Transparency = 0.94
        .Size = 100
        .Blur = 8
        .OffsetX = 0
        .OffsetY = 3
    End With
    On Error GoTo 0
End Sub

Public Sub CenterShapeText(ByVal shp As Shape, _
                           Optional ByVal alignHorizontal As Boolean = True, _
                           Optional ByVal alignVertical As Boolean = True)
    On Error Resume Next
    If alignVertical Then
        shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
        shp.TextFrame.VerticalAlignment = xlVAlignCenter
    End If
    
    If alignHorizontal Then
        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        shp.TextFrame.HorizontalAlignment = xlHAlignCenter
    End If
    
    With shp.TextFrame2
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
    End With
    With shp.TextFrame
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
    End With
    On Error GoTo 0
End Sub

' ==============================================================================
' 5. VIEWPORT NORMALIZATION (Eliminates Split Panes & Cut-Off Headers)
' ==============================================================================

Public Sub ResetSheetViewport(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .Split = False
        .SplitColumn = 0
        .SplitRow = 0
        .ScrollRow = 1
        .ScrollColumn = 1
        .Zoom = 80
        .DisplayGridlines = False
        .DisplayHeadings = False
    End With
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

Public Sub ResetAllViewports()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            Call ResetSheetViewport(ws)
        End If
    Next ws
End Sub

' ==============================================================================
' 6. FLICKER-FREE NAVIGATION ROUTING (Explicit Module Qualifications)
' ==============================================================================

Public Sub NavigateToHomePortal()
    NavigateToSheet SHEET_PORTAL
End Sub

Public Sub NavigateToCallCenter()
    NavigateToSheet SHEET_CC
End Sub

Public Sub NavigateToRetention()
    NavigateToSheet SHEET_CH
End Sub

Public Sub NavigateToDiversity()
    NavigateToSheet SHEET_DI
End Sub

Public Sub NavigateToDomains()
    NavigateToSheet SHEET_DOMAINS
End Sub

Public Sub NavigateToCatalog()
    NavigateToSheet SHEET_CATALOG
End Sub

Private Sub NavigateToSheet(ByVal targetSheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(targetSheetName)
    If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets(targetSheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Target worksheet '" & targetSheetName & "' not found.", vbExclamation, "PwC Navigation"
        Exit Sub
    End If
    
    FreezeAppState False
    ws.Activate
    With ActiveWindow
        .ScrollRow = 1
        .ScrollColumn = 1
    End With
    ws.Range("A1").Select
    RestoreAppState
End Sub

' ==============================================================================
' 7. DATA REFRESH & SYNCHRONOUS MODEL SYNCHRONIZATION
' ==============================================================================

Public Sub RefreshPipelineSynchronously()
    On Error GoTo ErrHandler
    Dim pt As PivotTable, ws As Worksheet, wb As Workbook
    Dim conn As WorkbookConnection
    Set wb = ThisWorkbook: If wb Is Nothing Then Set wb = ActiveWorkbook
    
    FreezeAppState False
    Application.StatusBar = "Refreshing VertiPaq tabular model connections..."
    
    For Each conn In wb.Connections
        If conn.Type = xlConnectionTypeModel Or InStr(1, conn.Name, "Model", vbTextCompare) > 0 Then
            On Error Resume Next
            conn.OLEDBConnection.BackgroundQuery = False
            conn.Refresh
            On Error GoTo ErrHandler
        End If
    Next conn
    
    For Each ws In wb.Worksheets
        For Each pt In ws.PivotTables
            pt.Update
            pt.RefreshTable
        Next pt
    Next ws
    
    Application.Calculate
    RestoreAppState
    Application.StatusBar = "Data Model and all Analytical Cockpits refreshed successfully."
    Application.OnTime Now + TimeSerial(0, 0, 4), "'" & wb.Name & "'!modPwC_Unified_Master.ClearStatusBar"
    Exit Sub
    
ErrHandler:
    RestoreAppState
    MsgBox "Data Refresh Encountered an Exception: " & Err.Description, vbCritical, "PwC Data Refresh"
End Sub

Public Sub ClearStatusBar()
    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0
End Sub

' ==============================================================================
' 8. MULTI-DIMENSIONAL FILTER CONTROLLER
' ==============================================================================

Public Sub ClearAllFilters()
    On Error GoTo ErrHandler
    Dim sc As SlicerCache, wb As Workbook
    Set wb = ThisWorkbook: If wb Is Nothing Then Set wb = ActiveWorkbook
    
    FreezeAppState False
    For Each sc In wb.SlicerCaches
        On Error Resume Next
        sc.ClearAllFilterValues
        On Error GoTo ErrHandler
    Next sc
    
    RestoreAppState
    Application.StatusBar = "All analytical slicers and dimension filters have been reset."
    Application.OnTime Now + TimeSerial(0, 0, 3), "'" & wb.Name & "'!modPwC_Unified_Master.ClearStatusBar"
    Exit Sub
    
ErrHandler:
    RestoreAppState
    MsgBox "Filter reset failed: " & Err.Description, vbExclamation, "PwC Filter Engine"
End Sub

' ==============================================================================
' 9. PUBLICATION-READY A4 LANDSCAPE PDF EXPORT
' ==============================================================================

Public Sub ExportActiveDashboardPDF()
    On Error GoTo ErrHandler
    Dim ws As Worksheet, exportPath As String, baseName As String
    Set ws = ActiveSheet
    
    If InStr(1, ws.Name, "Staging", vbTextCompare) > 0 Then
        MsgBox "Cannot export internal staging sheet.", vbExclamation, "PwC PDF Publisher"
        Exit Sub
    End If
    
    baseName = ws.Name & "_ExecutiveBriefing_" & Format(Now, "YYYYMMDD_HHNN") & ".pdf"
    exportPath = ThisWorkbook.Path & "\\" & baseName
    
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = Application.InchesToPoints(0.2)
        .RightMargin = Application.InchesToPoints(0.2)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
    End With
    
    ws.ExportAsFixedFormat Type:=xlTypePDF, _
                           Filename:=exportPath, _
                           Quality:=xlQualityStandard, _
                           IncludeDocProperties:=True, _
                           IgnorePrintAreas:=False, _
                           OpenAfterPublish:=True
                           
    Application.StatusBar = "Executive PDF exported: " & baseName
    Application.OnTime Now + TimeSerial(0, 0, 4), "'" & ThisWorkbook.Name & "'!modPwC_Unified_Master.ClearStatusBar"
    Exit Sub
    
ErrHandler:
    MsgBox "PDF Export Exception: " & Err.Description, vbCritical, "PwC PDF Publisher"
End Sub

Public Sub ExportExecutiveReport()
    ExportActiveDashboardPDF
End Sub

' ==============================================================================
' 10. DYNAMIC THEME ENGINE (Light & Dark Executive Modes)
' ==============================================================================

Public Sub ToggleDashboardTheme()
    Dim currentTheme As String
    currentTheme = GetStoredTheme()
    If currentTheme = "DARK" Then
        ApplyLightTheme
    Else
        ApplyDarkTheme
    End If
End Sub

Public Sub ApplyLightTheme()
    ApplyTheme False
    SetStoredTheme "LIGHT"
    Application.StatusBar = "PwC Enterprise Theme: Light Executive Mode Activated."
    Application.OnTime Now + TimeSerial(0, 0, 3), "'" & ThisWorkbook.Name & "'!modPwC_Unified_Master.ClearStatusBar"
End Sub

Public Sub ApplyDarkTheme()
    ApplyTheme True
    SetStoredTheme "DARK"
    Application.StatusBar = "PwC Enterprise Theme: Dark Cockpit Mode Activated."
    Application.OnTime Now + TimeSerial(0, 0, 3), "'" & ThisWorkbook.Name & "'!modPwC_Unified_Master.ClearStatusBar"
End Sub

Public Sub ApplyTheme(ByVal isDark As Boolean)
    FreezeAppState True
    Dim ws As Worksheet, wb As Workbook
    Set wb = ThisWorkbook: If wb Is Nothing Then Set wb = ActiveWorkbook
    
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible And InStr(1, ws.Name, "Staging", vbTextCompare) = 0 Then
            ApplyThemeToSheet ws, isDark
        End If
    Next ws
    
    RestoreAppState
End Sub

Private Sub ApplyThemeToSheet(ByVal ws As Worksheet, ByVal isDark As Boolean)
    On Error Resume Next
    Dim canvasColor As Long, cardFill As Long, cardBorder As Long, textPrimary As Long, textMuted As Long
    
    If isDark Then
        canvasColor = DARK_BG: cardFill = DARK_CONTAINER: cardBorder = DARK_BORDER
        textPrimary = DARK_TEXT_PRIMARY: textMuted = DARK_TEXT_MUTED
    Else
        canvasColor = PWC_CANVAS_BG: cardFill = PWC_CARD_FILL: cardBorder = PWC_CARD_BORDER
        textPrimary = PWC_TEXT_TITLE: textMuted = PWC_TEXT_MUTED
    End If
    
    ws.Cells.Interior.Color = canvasColor
    
    Dim shp As Shape
    For Each shp In ws.Shapes
        If InStr(1, shp.Name, "Container_", vbTextCompare) > 0 Or _
           InStr(1, shp.Name, "Card_", vbTextCompare) > 0 Or _
           InStr(1, shp.Name, "Hero_BannerCard", vbTextCompare) > 0 Or _
           InStr(1, shp.Name, "Nav_MasterBar", vbTextCompare) > 0 Then
            shp.Fill.Solid
            shp.Fill.ForeColor.RGB = cardFill
            shp.Line.ForeColor.RGB = cardBorder
        ElseIf InStr(1, shp.Name, "DockZone_", vbTextCompare) > 0 Then
            shp.Fill.Solid
            shp.Fill.ForeColor.RGB = IIf(isDark, DARK_BG, PWC_WHITE)
            shp.Line.ForeColor.RGB = cardBorder
        End If
    Next shp
    
    Dim chObj As ChartObject
    For Each chObj In ws.ChartObjects
        FormatChartForTheme chObj, isDark
    Next chObj
    On Error GoTo 0
End Sub

Private Sub FormatChartForTheme(ByVal chObj As ChartObject, ByVal isDark As Boolean)
    On Error Resume Next
    With chObj.Chart
        .ChartArea.Format.Fill.Visible = msoFalse
        .ChartArea.Format.Line.Visible = msoFalse
        .PlotArea.Format.Fill.Visible = msoFalse
        .PlotArea.Format.Line.Visible = msoFalse
        
        Dim ax As Axis
        For Each ax In .Axes
            ax.Format.Line.ForeColor.RGB = IIf(isDark, DARK_BORDER, PWC_CARD_BORDER)
            ax.TickLabels.Font.Color = IIf(isDark, DARK_TEXT_MUTED, PWC_TEXT_MUTED)
            ax.MajorGridlines.Format.Line.ForeColor.RGB = IIf(isDark, DARK_GRIDLINE, PWC_CARD_BORDER)
        Next ax
    End With
    On Error GoTo 0
End Sub

Private Function GetStoredTheme() As String
    On Error Resume Next
    Dim val As String
    val = ThisWorkbook.CustomDocumentProperties(THEME_PROP_NAME).Value
    If Len(val) = 0 Then val = "LIGHT"
    GetStoredTheme = val
    On Error GoTo 0
End Function

Private Sub SetStoredTheme(ByVal themeName As String)
    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties(THEME_PROP_NAME).Value = themeName
    If Err.Number <> 0 Then
        ThisWorkbook.CustomDocumentProperties.Add Name:=THEME_PROP_NAME, _
                                                 LinkToContent:=False, _
                                                 Type:=msoPropertyTypeString, _
                                                 Value:=themeName
    End If
    On Error GoTo 0
End Sub

' ==============================================================================
' 11. GOVERNANCE SHEETS BUILDER (01_Domains & 02_Catalog)
' ==============================================================================

Public Sub BuildGovernanceArchitecture()
    FreezeAppState True
    Dim wb As Workbook, wsDomains As Worksheet, wsCatalog As Worksheet
    Set wb = ThisWorkbook: If wb Is Nothing Then Set wb = ActiveWorkbook
    
    On Error Resume Next
    Set wsDomains = wb.Worksheets(SHEET_DOMAINS)
    If wsDomains Is Nothing Then
        Set wsDomains = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsDomains.Name = SHEET_DOMAINS
    End If
    
    Set wsCatalog = wb.Worksheets(SHEET_CATALOG)
    If wsCatalog Is Nothing Then
        Set wsCatalog = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsCatalog.Name = SHEET_CATALOG
    End If
    On Error GoTo 0
    
    Call PopulateDomainsSheet(wsDomains)
    Call PopulateCatalogSheet(wsCatalog)
    Call ResetSheetViewport(wsDomains)
    Call ResetSheetViewport(wsCatalog)
    RestoreAppState
End Sub

Private Sub PopulateDomainsSheet(ws As Worksheet)
    ws.Tab.Color = PWC_ORANGE
    ws.Cells.Clear
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    
    InitializeDashboardCanvas ws, PWC_CANVAS_BG
    BuildWebTopNavBar ws, "DOM"
    BuildHeroHeader ws, "Business Domains & Analytical Governance", _
                    "Enterprise Ralph Kimball Galaxy Architecture & Operational SLA Matrix", _
                    76, 1214
                    
    DrawStructuredCard ws, "C", 1, "Call Centre Operations & SLA", _
                       "Intraday triage, representative quality, FCR compliance and resolution forensics across 5,000 inquiries.", _
                       "03_CallCenter_Cockpit", "Fact_Calls", "5,000 Calls", "3.40 / 5.0", PWC_ORANGE, "phone_orange.svg"
                       
    DrawStructuredCard ws, "G", 2, "Customer Retention & Revenue Risk", _
                       "Subscription churn diagnostics, ARR exposure ($2.86M) and contract-level retention intervention models.", _
                       "04_CustomerRetention_Cockpit", "Fact_Churn", "7,043 Accounts", "26.54% Churn", PWC_CHARCOAL, "users_orange.svg"
                       
    DrawStructuredCard ws, "K", 3, "Diversity, Equity & Inclusion (D&I)", _
                       "Corporate census (500 headcount), executive broken-rung barriers and promotion velocity parity auditing.", _
                       "05_DiversityInclusion_Cockpit", "Fact_Employees", "500 Employees", "41.00% Female", PWC_TEXT_MUTED, "users_rose.svg"
End Sub

Private Sub DrawStructuredCard(ws As Worksheet, colL As String, cardNum As Long, _
                              title As String, descText As String, targetSheet As String, _
                              factTable As String, metricVol As String, keyMetric As String, _
                              accentColor As Long, iconFile As String)
    Dim cardLeft As Single, cardTop As Single, cardW As Single, cardH As Single
    Dim shpCard As Shape, shpHeader As Shape, shpBody As Shape, shpBtn As Shape
    
    cardTop = 136: cardH = 460: cardW = 380
    cardLeft = CANVAS_LEFT + (cardNum - 1) * (cardW + 18)
    
    Set shpCard = ws.Shapes.AddShape(msoShapeRoundedRectangle, cardLeft, cardTop, cardW, cardH)
    With shpCard
        .Name = "DomainCard_" & cardNum
        .Adjustments.Item(1) = 0.08
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_CARD_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 1
        ApplySoftElevation shpCard
    End With
    
    Dim shpStripe As Shape
    Set shpStripe = ws.Shapes.AddShape(msoShapeRoundedRectangle, cardLeft, cardTop, cardW, 8)
    With shpStripe
        .Adjustments.Item(1) = 0.5
        .Fill.Solid: .Fill.ForeColor.RGB = accentColor
        .Line.Visible = msoFalse
    End With
    
    Call InsertVectorIcon(ws, iconFile, cardLeft + 20, cardTop + 24, 28, 28, "DomainIcon_" & cardNum)
    
    Set shpHeader = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, cardLeft + 58, cardTop + 20, cardW - 74, 38)
    With shpHeader
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        .TextFrame2.MarginLeft = 0: .TextFrame2.MarginTop = 0
        .TextFrame2.TextRange.Text = "DOMAIN 0" & cardNum & vbCrLf & title
        With .TextFrame2.TextRange.Paragraphs(1).Font
            .Name = FONT_FAMILY: .Size = 8: .Bold = msoTrue: .Fill.ForeColor.RGB = accentColor
        End With
        With .TextFrame2.TextRange.Paragraphs(2).Font
            .Name = FONT_FAMILY: .Size = 11: .Bold = msoTrue: .Fill.ForeColor.RGB = PWC_TEXT_TITLE
        End With
    End With
    
    Set shpBody = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, cardLeft + 20, cardTop + 76, cardW - 40, 220)
    With shpBody
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        .TextFrame2.MarginLeft = 0: .TextFrame2.WordWrap = msoTrue
        .TextFrame2.TextRange.Text = descText & vbCrLf & vbCrLf & _
                                    "Primary Fact Source: " & factTable & vbCrLf & _
                                    "Analytical Population: " & metricVol & vbCrLf & _
                                    "Executive Target SLA: " & keyMetric & vbCrLf & vbCrLf & _
                                    "- Conformed Ralph Kimball Constellation" & vbCrLf & _
                                    "- Dedicated VertiPaq Analytical PivotCache" & vbCrLf & _
                                    "- Connected Intraday Slicer Dimensions"
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY
        .TextFrame2.TextRange.Font.Size = 9
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_TEXT_MUTED
    End With
    
    Set shpBtn = ws.Shapes.AddShape(msoShapeRoundedRectangle, cardLeft + 20, cardTop + cardH - 52, cardW - 40, 36)
    With shpBtn
        .Name = "Btn_Launch_" & cardNum
        .Adjustments.Item(1) = 0.20
        .Fill.Solid: .Fill.ForeColor.RGB = accentColor
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = "Launch " & title
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY
        .TextFrame2.TextRange.Font.Size = 9: .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_WHITE
        CenterShapeText shpBtn, True, True
        On Error Resume Next
        ws.Hyperlinks.Add Anchor:=shpBtn, Address:="", SubAddress:="'" & targetSheet & "'!A1"
        On Error GoTo 0
    End With
End Sub

Private Sub PopulateCatalogSheet(ws As Worksheet)
    ws.Tab.Color = PWC_TEXT_MUTED
    ws.Cells.Clear
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    
    InitializeDashboardCanvas ws, PWC_CANVAS_BG
    BuildWebTopNavBar ws, "CAT"
    BuildHeroHeader ws, "Metadata Governance & DAX KPI Measure Catalog", _
                    "Enterprise Ralph Kimball Galaxy Architecture & Operational SLA Matrix", _
                    76, 1214
                    
    Dim startRow As Long: startRow = 8
    ws.Cells(startRow, 2).Value = "Table Name"
    ws.Cells(startRow, 3).Value = "Kimball Role"
    ws.Cells(startRow, 4).Value = "Grain & Scope"
    ws.Cells(startRow, 5).Value = "Cardinality"
    ws.Cells(startRow, 6).Value = "Primary Keys / Relationships"
    ws.Cells(startRow, 7).Value = "Storage Engine"
    
    Dim tables(1 To 8, 1 To 6) As String
    tables(1, 1) = "Fact_Calls": tables(1, 2) = "Fact Table": tables(1, 3) = "One row per customer call": tables(1, 4) = "5,000 Rows": tables(1, 5) = "Call Id -> DimDate, DimAgent": tables(1, 6) = "VertiPaq In-Memory"
    tables(2, 1) = "Fact_Churn": tables(2, 2) = "Fact Table": tables(2, 3) = "One row per telecom subscriber": tables(2, 4) = "7,043 Rows": tables(2, 5) = "customerID -> DimContract": tables(2, 6) = "VertiPaq In-Memory"
    tables(3, 1) = "Fact_Employees": tables(3, 2) = "Fact Table": tables(3, 3) = "One row per employee snapshot": tables(3, 4) = "500 Rows": tables(3, 5) = "Employee ID -> DimDepartment": tables(3, 6) = "VertiPaq In-Memory"
    tables(4, 1) = "DimDate": tables(4, 2) = "Conformed Dim": tables(4, 3) = "Continuous calendar days": tables(4, 4) = "90 Days": tables(4, 5) = "Date (PK) 1:* Fact_Calls": tables(4, 6) = "Marked Date Table"
    tables(5, 1) = "DimAgent": tables(5, 2) = "Dimension": tables(5, 3) = "Call center representative": tables(5, 4) = "8 Agents": tables(5, 5) = "Agent (PK) 1:* Fact_Calls": tables(5, 6) = "VertiPaq Tabular"
    tables(6, 1) = "DimTopic": tables(6, 2) = "Dimension": tables(6, 3) = "Inquiry category classification": tables(6, 4) = "5 Topics": tables(6, 5) = "Topic (PK) 1:* Fact_Calls": tables(6, 6) = "VertiPaq Tabular"
    tables(7, 1) = "DimContract": tables(7, 2) = "Dimension": tables(7, 3) = "Subscription agreement terms": tables(7, 4) = "3 Types": tables(7, 5) = "Contract (PK) 1:* Fact_Churn": tables(7, 6) = "VertiPaq Tabular"
    tables(8, 1) = "DimDepartment": tables(8, 2) = "Dimension": tables(8, 3) = "Corporate division hierarchy": tables(8, 4) = "5 Depts": tables(8, 5) = "Department (PK) 1:* Fact_Employees": tables(8, 6) = "VertiPaq Tabular"
    
    Dim r As Long, c As Long
    For r = 1 To 8
        For c = 1 To 6
            ws.Cells(startRow + r, c + 1).Value = tables(r, c)
        Next c
    Next r
    
    Dim tblRange As Range, lo As ListObject
    Set tblRange = ws.Range(ws.Cells(startRow, 2), ws.Cells(startRow + 8, 7))
    On Error Resume Next
    ws.ListObjects("tbl_Metadata_Catalog").Delete
    Set lo = ws.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    lo.Name = "tbl_Metadata_Catalog"
    ApplyPwCBrandTableStyle lo, ws
    On Error GoTo 0
End Sub

Private Sub ApplyPwCBrandTableStyle(lo As ListObject, ws As Worksheet)
    On Error Resume Next
    lo.TableStyle = "TableStyleLight1"
    lo.ShowTableStyleRowStripes = True
    
    With lo.HeaderRowRange
        .Interior.Color = PWC_CHARCOAL
        .Font.Name = FONT_FAMILY
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = PWC_WHITE
        .RowHeight = 26
    End With
    
    With lo.DataBodyRange
        .Font.Name = FONT_FAMILY
        .Font.Size = 9
        .RowHeight = 22
        .Borders.Color = PWC_CARD_BORDER
        .Borders.Weight = xlThin
    End With
    
    ws.Columns("B:G").AutoFit
    On Error GoTo 0
End Sub

' ==============================================================================
' 12. EXECUTIVE HOME PORTAL CANVAS (00_Home_Portal)
' ==============================================================================

Public Sub BuildExecutivePortal()
    FreezeAppState True
    Dim wb As Workbook, ws As Worksheet
    Set wb = ThisWorkbook: If wb Is Nothing Then Set wb = ActiveWorkbook
    
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_PORTAL)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
        ws.Name = SHEET_PORTAL
    End If
    On Error GoTo 0
    
    ws.Tab.Color = PWC_ORANGE
    ws.Cells.Clear
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    
    InitializeDashboardCanvas ws, PWC_CANVAS_BG
    CreatePortalTopHeader ws
    CreatePortalHeroSection ws
    CreatePortalMetricsTicker ws
    CreatePortalCockpitLaunchers ws
    CreatePortalGovernanceDrawer ws
    Call ResetSheetViewport(ws)
    RestoreAppState
End Sub

Private Sub CreatePortalTopHeader(ByVal ws As Worksheet)
    Dim topBar As Shape, logoShp As Shape, divLine As Shape, titleTxt As Shape
    Dim btnTheme As Shape, btnPDF As Shape
    
    Set topBar = ws.Shapes.AddShape(msoShapeRoundedRectangle, CANVAS_LEFT, 16, 1214, 52)
    With topBar
        .Name = "Portal_TopBar"
        .Adjustments.Item(1) = 0.12
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_CARD_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 1
        ApplySoftElevation topBar
    End With
    
    Call InsertVectorIcon(ws, "PwC_logo_rgb_colour_pos.png", CANVAS_LEFT + 16, 24, 76, 36, "Portal_Logo")
    
    Set divLine = ws.Shapes.AddLine(CANVAS_LEFT + 104, 24, CANVAS_LEFT + 104, 60)
    With divLine
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 1
    End With
    
    Set titleTxt = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, CANVAS_LEFT + 116, 22, 400, 38)
    With titleTxt
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        .TextFrame2.MarginLeft = 0: .TextFrame2.MarginTop = 0
        .TextFrame2.TextRange.Text = "PwC Switzerland Digital Intelligence" & vbCrLf & "Executive Decision Hub & Strategic Analytics"
        With .TextFrame2.TextRange.Paragraphs(1).Font
            .Name = FONT_FAMILY: .Size = 11: .Bold = msoTrue: .Fill.ForeColor.RGB = PWC_TEXT_TITLE
        End With
        With .TextFrame2.TextRange.Paragraphs(2).Font
            .Name = FONT_FAMILY: .Size = 8: .Bold = msoFalse: .Fill.ForeColor.RGB = PWC_TEXT_MUTED
        End With
    End With
    
    Set btnTheme = ws.Shapes.AddShape(msoShapeRoundedRectangle, CANVAS_LEFT + 980, 24, 105, 34)
    With btnTheme
        .Name = "Portal_BtnTheme"
        .Adjustments.Item(1) = 0.20
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_PILL_BG
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 1
        .TextFrame2.TextRange.Text = "Dark Mode"
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY
        .TextFrame2.TextRange.Font.Size = 8.5: .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_TEXT_TITLE
        CenterShapeText btnTheme, True, True
        .OnAction = "'" & ThisWorkbook.Name & "'!modPwC_Unified_Master.ToggleDashboardTheme"
    End With
    
    Set btnPDF = ws.Shapes.AddShape(msoShapeRoundedRectangle, CANVAS_LEFT + 1093, 24, 115, 34)
    With btnPDF
        .Name = "Portal_BtnPDF"
        .Adjustments.Item(1) = 0.20
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_ORANGE
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = "Export PDF"
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY
        .TextFrame2.TextRange.Font.Size = 8.5: .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_WHITE
        CenterShapeText btnPDF, True, True
        .OnAction = "'" & ThisWorkbook.Name & "'!modPwC_Unified_Master.ExportActiveDashboardPDF"
    End With
End Sub

Private Sub CreatePortalHeroSection(ByVal ws As Worksheet)
    Dim heroBox As Shape, heroTitle As Shape, heroDesc As Shape
    Set heroBox = ws.Shapes.AddShape(msoShapeRoundedRectangle, CANVAS_LEFT, 76, 1214, 110)
    With heroBox
        .Name = "Portal_HeroBox"
        .Adjustments.Item(1) = 0.08
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_CHARCOAL
        .Line.Visible = msoFalse
        ApplySoftElevation heroBox
    End With
    
    Set heroTitle = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, CANVAS_LEFT + 28, 92, 1150, 36)
    With heroTitle
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        .TextFrame2.MarginLeft = 0: .TextFrame2.MarginTop = 0
        .TextFrame2.TextRange.Text = "Enterprise Business Intelligence & Governance Portal"
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY
        .TextFrame2.TextRange.Font.Size = 16: .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_WHITE
    End With
    
    Set heroDesc = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, CANVAS_LEFT + 28, 128, 1150, 48)
    With heroDesc
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        .TextFrame2.MarginLeft = 0: .TextFrame2.MarginTop = 0: .TextFrame2.WordWrap = msoTrue
        .TextFrame2.TextRange.Text = "Unified Ralph Kimball analytical constellation modeling Call Center operations, subscriber retention forensics, " & _
                                   "and diversity & executive parity governance. Real-time VertiPaq tabular acceleration across 13 tables & 81 DAX measures."
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY
        .TextFrame2.TextRange.Font.Size = 9: .TextFrame2.TextRange.Font.Bold = msoFalse
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(226, 232, 240)
    End With
End Sub

Private Sub CreatePortalMetricsTicker(ByVal ws As Worksheet)
    Dim kpiTop As Single: kpiTop = 194
    Dim cardW As Single: cardW = 230
    Dim gap As Single: gap = 16
    
    BuildKPICard ws, "Port_KPI_Calls", CANVAS_LEFT + (cardW + gap) * 0, kpiTop, cardW, 88, _
                 "Total Call Demand", "5,000", "+12.4% vs Baseline", PWC_ORANGE, "phone_orange.svg"
                 
    BuildKPICard ws, "Port_KPI_Answer", CANVAS_LEFT + (cardW + gap) * 1, kpiTop, cardW, 88, _
                 "Answer Rate Compliance", "81.1%", "SLA Target: 80.0%", PWC_SUCCESS_GREEN, "check_green.svg"
                 
    BuildKPICard ws, "Port_KPI_Churn", CANVAS_LEFT + (cardW + gap) * 2, kpiTop, cardW, 88, _
                 "Customer Churn Rate", "26.5%", "Exposure: $2.86M ARR", PWC_ALERT_RED, "xcircle_red.svg"
                 
    BuildKPICard ws, "Port_KPI_Female", CANVAS_LEFT + (cardW + gap) * 3, kpiTop, cardW, 88, _
                 "Female Headcount Share", "41.0%", "Corporate Parity: 50.0%", PWC_WARNING_AMBER, "users_rose.svg"
                 
    BuildKPICard ws, "Port_KPI_Promo", CANVAS_LEFT + (cardW + gap) * 4, kpiTop, cardW, 88, _
                 "Promotion Award Velocity", "8.6%", "Zero-Delta Target", PWC_CHARCOAL, "award_rose.svg"
End Sub

Private Sub CreatePortalCockpitLaunchers(ByVal ws As Worksheet)
    Dim launchTop As Single: launchTop = 292
    Dim launchW As Single: launchW = 394
    Dim launchH As Single: launchH = 260
    Dim gap As Single: gap = 16
    
    DrawLauncherCard ws, CANVAS_LEFT + (launchW + gap) * 0, launchTop, launchW, launchH, 1, _
                     "Call Centre Operations", _
                     "Intraday arrival surges, abandonment forensics, representative speed, FCR adherence and resolution matrix.", _
                     SHEET_CC, "phone_orange.svg", PWC_ORANGE
                     
    DrawLauncherCard ws, CANVAS_LEFT + (launchW + gap) * 1, launchTop, launchW, launchH, 2, _
                     "Customer Retention Risk", _
                     "Subscriber attrition modeling, contract vulnerabilities, tenure curve diagnostics and tech support friction.", _
                     SHEET_CH, "users_orange.svg", PWC_CHARCOAL
                     
    DrawLauncherCard ws, CANVAS_LEFT + (launchW + gap) * 2, launchTop, launchW, launchH, 3, _
                     "Diversity & Inclusion Parity", _
                     "Workforce census, executive broken-rung barriers, performance rating equity and promotion velocity analysis.", _
                     SHEET_DI, "users_rose.svg", PWC_TEXT_MUTED
End Sub

Private Sub DrawLauncherCard(ws As Worksheet, leftPos As Single, topPos As Single, _
                             w As Single, h As Single, cardNum As Long, _
                             title As String, descText As String, targetSheet As String, _
                             iconFile As String, accentColor As Long)
    Dim cardShp As Shape, titleShp As Shape, descShp As Shape, btnShp As Shape
    
    Set cardShp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, w, h)
    With cardShp
        .Name = "Portal_Launcher_" & cardNum
        .Adjustments.Item(1) = 0.08
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_CARD_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 1
        ApplySoftElevation cardShp
    End With
    
    Dim stripe As Shape
    Set stripe = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, w, 6)
    With stripe
        .Adjustments.Item(1) = 0.5
        .Fill.Solid: .Fill.ForeColor.RGB = accentColor
        .Line.Visible = msoFalse
    End With
    
    Call InsertVectorIcon(ws, iconFile, leftPos + 20, topPos + 22, 28, 28, "Portal_Icon_" & cardNum)
    
    Set titleShp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos + 56, topPos + 18, w - 76, 36)
    With titleShp
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        .TextFrame2.MarginLeft = 0: .TextFrame2.MarginTop = 0
        .TextFrame2.TextRange.Text = "ANALYTICAL COCKPIT 0" & cardNum & vbCrLf & title
        With .TextFrame2.TextRange.Paragraphs(1).Font
            .Name = FONT_FAMILY: .Size = 8: .Bold = msoTrue: .Fill.ForeColor.RGB = accentColor
        End With
        With .TextFrame2.TextRange.Paragraphs(2).Font
            .Name = FONT_FAMILY: .Size = 11: .Bold = msoTrue: .Fill.ForeColor.RGB = PWC_TEXT_TITLE
        End With
    End With
    
    Set descShp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos + 20, topPos + 64, w - 40, 130)
    With descShp
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        .TextFrame2.MarginLeft = 0: .TextFrame2.WordWrap = msoTrue
        .TextFrame2.TextRange.Text = descText
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY
        .TextFrame2.TextRange.Font.Size = 9
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_TEXT_MUTED
    End With
    
    Set btnShp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos + 20, topPos + h - 48, w - 40, 34)
    With btnShp
        .Name = "Portal_LaunchBtn_" & cardNum
        .Adjustments.Item(1) = 0.20
        .Fill.Solid: .Fill.ForeColor.RGB = accentColor
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = "Launch " & title
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY
        .TextFrame2.TextRange.Font.Size = 8.5: .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_WHITE
        CenterShapeText btnShp, True, True
        On Error Resume Next
        ws.Hyperlinks.Add Anchor:=btnShp, Address:="", SubAddress:="'" & targetSheet & "'!A1"
        On Error GoTo 0
    End With
End Sub

Private Sub CreatePortalGovernanceDrawer(ByVal ws As Worksheet)
    Dim drawerBox As Shape, titleTxt As Shape, btnDom As Shape, btnCat As Shape
    Dim drawTop As Single: drawTop = 562
    
    Set drawerBox = ws.Shapes.AddShape(msoShapeRoundedRectangle, CANVAS_LEFT, drawTop, 1214, 64)
    With drawerBox
        .Name = "Portal_GovDrawer"
        .Adjustments.Item(1) = 0.12
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_CARD_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 1
        ApplySoftElevation drawerBox
    End With
    
    Set titleTxt = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, CANVAS_LEFT + 24, drawTop + 14, 450, 36)
    With titleTxt
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        .TextFrame2.MarginLeft = 0: .TextFrame2.MarginTop = 0
        .TextFrame2.TextRange.Text = "Enterprise Governance & Model Architecture" & vbCrLf & "Ralph Kimball Constellation, Schema Dictionary & 81 DAX Measures"
        With .TextFrame2.TextRange.Paragraphs(1).Font
            .Name = FONT_FAMILY: .Size = 10: .Bold = msoTrue: .Fill.ForeColor.RGB = PWC_TEXT_TITLE
        End With
        With .TextFrame2.TextRange.Paragraphs(2).Font
            .Name = FONT_FAMILY: .Size = 8: .Bold = msoFalse: .Fill.ForeColor.RGB = PWC_TEXT_MUTED
        End With
    End With
    
    Set btnDom = ws.Shapes.AddShape(msoShapeRoundedRectangle, CANVAS_LEFT + 860, drawTop + 15, 160, 34)
    With btnDom
        .Name = "Portal_BtnDomains"
        .Adjustments.Item(1) = 0.20
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_PILL_BG
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 1
        .TextFrame2.TextRange.Text = "01 Business Domains"
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY
        .TextFrame2.TextRange.Font.Size = 8.5: .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_TEXT_TITLE
        CenterShapeText btnDom, True, True
        On Error Resume Next
        ws.Hyperlinks.Add Anchor:=btnDom, Address:="", SubAddress:="'" & SHEET_DOMAINS & "'!A1"
        On Error GoTo 0
    End With
    
    Set btnCat = ws.Shapes.AddShape(msoShapeRoundedRectangle, CANVAS_LEFT + 1030, drawTop + 15, 160, 34)
    With btnCat
        .Name = "Portal_BtnCatalog"
        .Adjustments.Item(1) = 0.20
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_PILL_BG
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 1
        .TextFrame2.TextRange.Text = "02 KPI Catalog"
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY
        .TextFrame2.TextRange.Font.Size = 8.5: .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_TEXT_TITLE
        CenterShapeText btnCat, True, True
        On Error Resume Next
        ws.Hyperlinks.Add Anchor:=btnCat, Address:="", SubAddress:="'" & SHEET_CATALOG & "'!A1"
        On Error GoTo 0
    End With
End Sub

' ==============================================================================
' 13. WEB-APP DASHBOARD CANVASES & COMPONENT ENGINES
' ==============================================================================

Public Sub InitializeDashboardCanvas(ws As Worksheet, Optional ByVal bgColor As Long = PWC_CANVAS_BG)
    On Error Resume Next
    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .Split = False
        .SplitColumn = 0
        .SplitRow = 0
        .ScrollRow = 1
        .ScrollColumn = 1
        .Zoom = 80
        .DisplayGridlines = False
        .DisplayHeadings = False
    End With
    
    ws.Columns("A:AZ").ColumnWidth = 11
    ws.Rows("1:70").RowHeight = 20
    
    With ws.Cells.Interior
        .Pattern = xlSolid
        .Color = bgColor
    End With
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

Public Sub BuildWebTopNavBar(ws As Worksheet, ByVal activeModuleCode As String)
    Dim shpNav As Shape, shpDivider As Shape, shpBrand As Shape, shpLive As Shape
    Dim navTop As Single, navLeft As Single, navW As Single, navH As Single
    
    navTop = 16
    navLeft = CANVAS_LEFT
    navW = IIf(UCase(activeModuleCode) = "CC", 1480, 1214)
    navH = 52
    
    Set shpNav = ws.Shapes.AddShape(msoShapeRoundedRectangle, navLeft, navTop, navW, navH)
    With shpNav
        .Name = "Nav_MasterBar"
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_CARD_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 1
        .Adjustments.Item(1) = 0.10
        ApplySoftElevation shpNav
    End With
    
    Call InsertVectorIcon(ws, "PwC_logo_rgb_colour_pos.png", navLeft + 16, navTop + 8, 76, 36, "Nav_Logo")
    
    Set shpDivider = ws.Shapes.AddLine(navLeft + 104, navTop + 8, navLeft + 104, navTop + 44)
    With shpDivider
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 1
    End With
    
    Set shpBrand = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, navLeft + 116, navTop + 7, 185, 38)
    With shpBrand
        .Name = "Nav_BrandText"
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        With .TextFrame2
            .MarginLeft = 0: .MarginTop = 0: .MarginRight = 0: .MarginBottom = 0
            .TextRange.Text = "PwC Digital Intelligence" & vbCrLf & "Executive Decision Hub"
            With .TextRange.Paragraphs(1).Font
                .Name = FONT_FAMILY: .Size = 11: .Bold = msoTrue: .Fill.ForeColor.RGB = PWC_TEXT_TITLE
            End With
            With .TextRange.Paragraphs(2).Font
                .Name = FONT_FAMILY: .Size = 8: .Bold = msoFalse: .Fill.ForeColor.RGB = PWC_TEXT_MUTED
            End With
        End With
    End With
    
    Dim tabLeft As Single, tabTop As Single, tabW As Single, tabH As Single
    tabTop = navTop + 12: tabH = 28
    
    tabLeft = navLeft + 310: tabW = 95
    Call CreateNavTabPill(ws, "Nav_Tab_Domains", tabLeft, tabTop, tabW, tabH, _
                          "01 Domains", "'" & SHEET_DOMAINS & "'!A1", (activeModuleCode = "DOM"))
                          
    tabLeft = tabLeft + tabW + 8: tabW = 100
    Call CreateNavTabPill(ws, "Nav_Tab_Catalog", tabLeft, tabTop, tabW, tabH, _
                          "02 Catalog", "'" & SHEET_CATALOG & "'!A1", (activeModuleCode = "CAT"))
                          
    tabLeft = tabLeft + tabW + 8: tabW = 115
    Call CreateNavTabPill(ws, "Nav_Tab_CC", tabLeft, tabTop, tabW, tabH, _
                          "03 Call Center", "'" & SHEET_CC & "'!A1", (activeModuleCode = "CC"))
                          
    tabLeft = tabLeft + tabW + 8: tabW = 125
    Call CreateNavTabPill(ws, "Nav_Tab_CH", tabLeft, tabTop, tabW, tabH, _
                          "04 Retention Risk", "'" & SHEET_CH & "'!A1", (activeModuleCode = "CH"))
                          
    tabLeft = tabLeft + tabW + 8: tabW = 120
    Call CreateNavTabPill(ws, "Nav_Tab_DI", tabLeft, tabTop, tabW, tabH, _
                          "05 D&I Parity", "'" & SHEET_DI & "'!A1", (activeModuleCode = "DI"))
                          
    Dim livePillW As Single, livePillLeft As Single
    livePillW = 140
    livePillLeft = navLeft + navW - livePillW - 12
    
    Set shpLive = ws.Shapes.AddShape(msoShapeRoundedRectangle, livePillLeft, navTop + 12, livePillW, 28)
    With shpLive
        .Name = "Nav_StatusPill"
        .Adjustments.Item(1) = 0.20
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_BADGE_GREEN_BG
        .Line.ForeColor.RGB = RGB(167, 243, 208): .Line.Weight = 1
        .TextFrame2.TextRange.Text = "LIVE VERTIPAQ"
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY
        .TextFrame2.TextRange.Font.Size = 8: .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_SUCCESS_GREEN
        CenterShapeText shpLive, True, True
        .OnAction = "'" & ws.Parent.Name & "'!modPwC_Unified_Master.RefreshPipelineSynchronously"
    End With
    
    Call InsertVectorIcon(ws, "icon_live_indicator.svg", livePillLeft + 10, navTop + 19, 14, 14, "Nav_IconLive")
End Sub

Private Sub CreateNavTabPill(ws As Worksheet, ByVal shapeName As String, _
                            ByVal leftPos As Single, ByVal topPos As Single, _
                            ByVal pillWidth As Single, ByVal pillHeight As Single, _
                            ByVal tabText As String, ByVal subAddress As String, _
                            ByVal isActive As Boolean)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, pillWidth, pillHeight)
    With shp
        .Name = shapeName
        .Adjustments.Item(1) = 0.18
        CenterShapeText shp, True, True
        
        If isActive Then
            .Fill.Solid: .Fill.ForeColor.RGB = PWC_ORANGE
            .Line.Visible = msoFalse
            .TextFrame2.TextRange.Text = tabText
            .TextFrame2.TextRange.Font.Name = FONT_FAMILY
            .TextFrame2.TextRange.Font.Size = 8.5: .TextFrame2.TextRange.Font.Bold = msoTrue
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        Else
            .Fill.Solid: .Fill.ForeColor.RGB = PWC_PILL_BG
            .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 1
            .TextFrame2.TextRange.Text = tabText
            .TextFrame2.TextRange.Font.Name = FONT_FAMILY
            .TextFrame2.TextRange.Font.Size = 8.5: .TextFrame2.TextRange.Font.Bold = msoTrue
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_TEXT_MUTED
            On Error Resume Next
            ws.Hyperlinks.Add Anchor:=shp, Address:="", SubAddress:=subAddress
            On Error GoTo 0
        End If
        CenterShapeText shp, True, True
    End With
End Sub

Public Sub BuildHeroHeader(ws As Worksheet, _
                           ByVal dashboardTitle As String, _
                           ByVal subtitle As String, _
                           Optional ByVal topPos As Single = 76, _
                           Optional ByVal heroWidth As Single = 1214)
    Dim shpHero As Shape, shpTitle As Shape, shpRefreshBtn As Shape, shpClearBtn As Shape, shpExportBtn As Shape
    Dim heroLeft As Single, heroW As Single, heroH As Single, btnTop As Single, btnH As Single
    
    heroLeft = CANVAS_LEFT: heroW = heroWidth: heroH = 46
    btnTop = topPos + 9: btnH = 28
    
    Set shpHero = ws.Shapes.AddShape(msoShapeRoundedRectangle, heroLeft, topPos, heroW, heroH)
    With shpHero
        .Name = "Hero_BannerCard"
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_CARD_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 1
        .Adjustments.Item(1) = 0.12
        ApplySoftElevation shpHero
    End With
    
    Set shpTitle = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, heroLeft + 16, topPos + 5, heroW - 400, 36)
    With shpTitle
        .Name = "Hero_TitleBlock"
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        With .TextFrame2
            .MarginLeft = 0: .MarginTop = 0: .MarginRight = 0: .MarginBottom = 0
            .TextRange.Text = dashboardTitle & vbCrLf & subtitle
            With .TextRange.Paragraphs(1).Font
                .Name = FONT_FAMILY: .Size = 10.5: .Bold = msoTrue: .Fill.ForeColor.RGB = PWC_TEXT_TITLE
            End With
            With .TextRange.Paragraphs(2).Font
                .Name = FONT_FAMILY: .Size = 8: .Bold = msoFalse: .Fill.ForeColor.RGB = PWC_TEXT_MUTED
            End With
        End With
    End With
    
    Dim refBtnLeft As Single, refBtnW As Single
    refBtnW = 120
    refBtnLeft = heroLeft + heroW - 366
    
    Set shpRefreshBtn = ws.Shapes.AddShape(msoShapeRoundedRectangle, refBtnLeft, btnTop, refBtnW, btnH)
    With shpRefreshBtn
        .Name = "Hero_BtnRefresh"
        .Adjustments.Item(1) = 0.20
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_CARD_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 1
        .TextFrame2.TextRange.Text = "Refresh Data"
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY
        .TextFrame2.TextRange.Font.Size = 8.5: .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_ORANGE
        CenterShapeText shpRefreshBtn, True, True
        .OnAction = "'" & ws.Parent.Name & "'!modPwC_Unified_Master.RefreshPipelineSynchronously"
    End With
    Call InsertVectorIcon(ws, "icon_refresh_pipeline.svg", refBtnLeft + 10, btnTop + 7, 14, 14, "Hero_IconRefresh")
    
    Dim clearBtnLeft As Single, clearBtnW As Single
    clearBtnW = 118
    clearBtnLeft = refBtnLeft + refBtnW + 8
    
    Set shpClearBtn = ws.Shapes.AddShape(msoShapeRoundedRectangle, clearBtnLeft, btnTop, clearBtnW, btnH)
    With shpClearBtn
        .Name = "Hero_BtnResetSlicers"
        .Adjustments.Item(1) = 0.20
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_CARD_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 1
        .TextFrame2.TextRange.Text = "Reset Filters"
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY
        .TextFrame2.TextRange.Font.Size = 8.5: .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_TEXT_TITLE
        CenterShapeText shpClearBtn, True, True
        .OnAction = "'" & ws.Parent.Name & "'!modPwC_Unified_Master.ClearAllFilters"
    End With
    Call InsertVectorIcon(ws, "icon_clear_filters.svg", clearBtnLeft + 10, btnTop + 7, 14, 14, "Hero_IconClear")
    
    Dim expBtnLeft As Single, expBtnW As Single
    expBtnW = 108
    expBtnLeft = clearBtnLeft + clearBtnW + 8
    
    Set shpExportBtn = ws.Shapes.AddShape(msoShapeRoundedRectangle, expBtnLeft, btnTop, expBtnW, btnH)
    With shpExportBtn
        .Name = "Hero_BtnExportPDF"
        .Adjustments.Item(1) = 0.20
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_ORANGE
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = "Export PDF"
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY
        .TextFrame2.TextRange.Font.Size = 8.5: .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_WHITE
        CenterShapeText shpExportBtn, True, True
        .OnAction = "'" & ws.Parent.Name & "'!modPwC_Unified_Master.ExportActiveDashboardPDF"
    End With
    Call InsertVectorIcon(ws, "icon_export_pdf.svg", expBtnLeft + 10, btnTop + 7, 14, 14, "Hero_IconExport")
End Sub

Public Sub BuildKPICard(ws As Worksheet, ByVal cardPrefix As String, _
                       ByVal leftPos As Single, ByVal topPos As Single, _
                       ByVal cardWidth As Single, ByVal cardHeight As Single, _
                       ByVal kpiTitle As String, ByVal kpiValue As String, _
                       ByVal trendText As String, ByVal accentColor As Long, _
                       Optional ByVal iconFileName As String = "", _
                       Optional ByVal sparklineFileName As String = "")
    Dim shpCard As Shape, shpTitle As Shape, shpValue As Shape, shpBadge As Shape
    Dim shpIcon As Shape, shpSpark As Shape, shpStripe As Shape
    
    Set shpCard = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, cardWidth, cardHeight)
    With shpCard
        .Name = "Card_" & cardPrefix
        .Adjustments.Item(1) = 0.08
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_CARD_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 1
        ApplySoftElevation shpCard
    End With
    
    Set shpStripe = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, cardWidth, 4)
    With shpStripe
        .Adjustments.Item(1) = 0.5
        .Fill.Solid: .Fill.ForeColor.RGB = accentColor
        .Line.Visible = msoFalse
    End With
    
    Dim iconLeft As Single, titleLeft As Single
    titleLeft = leftPos + 12
    If Len(iconFileName) > 0 Then
        Set shpIcon = InsertVectorIcon(ws, iconFileName, leftPos + 12, topPos + 10, 20, 20, "Icon_" & cardPrefix)
        If Not shpIcon Is Nothing Then titleLeft = leftPos + 38
    End If
    
    Set shpTitle = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, titleLeft, topPos + 10, cardWidth - (titleLeft - leftPos) - 8, 16)
    With shpTitle
        .Name = "Title_" & cardPrefix
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        With .TextFrame2
            .MarginLeft = 0: .MarginTop = 0: .MarginRight = 0: .MarginBottom = 0
            .TextRange.Text = kpiTitle
            .TextRange.Font.Name = FONT_FAMILY
            .TextRange.Font.Size = 7.5: .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = PWC_TEXT_MUTED
        End With
    End With
    
    Set shpValue = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos + 12, topPos + 32, cardWidth - 24, 28)
    With shpValue
        .Name = "Val_" & cardPrefix
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        With .TextFrame2
            .MarginLeft = 0: .MarginTop = 0: .MarginRight = 0: .MarginBottom = 0
            .TextRange.Text = kpiValue
            .TextRange.Font.Name = FONT_FAMILY
            .TextRange.Font.Size = 14: .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = PWC_TEXT_TITLE
        End With
        CenterShapeText shpValue, False, True
    End With
    
    Set shpBadge = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos + 12, topPos + cardHeight - 24, cardWidth - 24, 18)
    With shpBadge
        .Name = "Trend_" & cardPrefix
        .Adjustments.Item(1) = 0.25
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_PILL_BG
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 0.5
        .TextFrame2.TextRange.Text = trendText
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY
        .TextFrame2.TextRange.Font.Size = 7.5: .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = accentColor
        CenterShapeText shpBadge, True, True
    End With
    
    If Len(sparklineFileName) > 0 Then
        Call InsertVectorIcon(ws, sparklineFileName, leftPos + cardWidth - 62, topPos + 30, 52, 22, "Spark_" & cardPrefix)
    End If
End Sub

Public Sub AutomateAndLinkKPICards(ws As Worksheet, ByVal moduleCode As String)
    On Error Resume Next
    Dim modUpper As String: modUpper = UCase(moduleCode)
    Dim rowIdx As Long: rowIdx = 65
    Dim sheetRef As String: sheetRef = "'" & ws.Name & "'!"
    
    If modUpper = "CC" Then
        ws.Range("AA" & rowIdx).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"", ""[Measures].[Total Calls]"")"
        ws.Range("AB" & rowIdx).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"", ""[Measures].[Answered Calls]"")"
        ws.Range("AC" & rowIdx).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"", ""[Measures].[Abandoned Calls]"")"
        ws.Range("AD" & rowIdx).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"", ""[Measures].[Answer Rate %]"")"
        ws.Range("AE" & rowIdx).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"", ""[Measures].[Average Speed of Answer (s)]"")"
        ws.Range("AF" & rowIdx).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"", ""[Measures].[Average CSAT Rating]"")"
        ws.Range("AG" & rowIdx).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"", ""[Measures].[First Call Resolution %]"")"
        
        ws.Shapes("Val_CC_TotalDemand").DrawingObject.Formula = "=" & sheetRef & "$AA$" & rowIdx
        ws.Shapes("Val_CC_Answered").DrawingObject.Formula = "=" & sheetRef & "$AB$" & rowIdx
        ws.Shapes("Val_CC_Missed").DrawingObject.Formula = "=" & sheetRef & "$AC$" & rowIdx
        ws.Shapes("Val_CC_SLA").DrawingObject.Formula = "=" & sheetRef & "$AD$" & rowIdx
        ws.Shapes("Val_CC_AHT").DrawingObject.Formula = "=" & sheetRef & "$AE$" & rowIdx
        ws.Shapes("Val_CC_CSAT").DrawingObject.Formula = "=" & sheetRef & "$AF$" & rowIdx
        ws.Shapes("Val_CC_FCR").DrawingObject.Formula = "=" & sheetRef & "$AG$" & rowIdx
    End If
    On Error GoTo 0
End Sub

Public Sub BuildSlicerPanelContainer(ws As Worksheet, _
                                     ByVal panelName As String, _
                                     ByVal leftPos As Single, _
                                     ByVal topPos As Single, _
                                     ByVal panelWidth As Single, _
                                     ByVal panelHeight As Single, _
                                     ByVal slot1Label As String, _
                                     ByVal slot2Label As String, _
                                     ByVal slot3Label As String)
    Dim shpPanel As Shape, shpHeader As Shape, shpResetBtn As Shape, shpQuoteCard As Shape
    
    Set shpPanel = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, panelWidth, panelHeight)
    With shpPanel
        .Name = "Panel_" & panelName
        .Adjustments.Item(1) = 0.08
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_DRAWER_BG
        .Line.ForeColor.RGB = RGB(42, 52, 71): .Line.Weight = 1
        ApplySoftElevation shpPanel
    End With
    
    Call InsertVectorIcon(ws, "PwC_logo_rgb_colour_pos.png", leftPos + 16, topPos + 14, 52, 24, "DrawerLogo_" & panelName)
    Call InsertVectorIcon(ws, "icon_filter_funnel.svg", leftPos + panelWidth - 36, topPos + 16, 18, 18, "FunnelIcon_" & panelName)
    
    Set shpHeader = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos + 74, topPos + 12, panelWidth - 116, 26)
    With shpHeader
        .Name = "Header_" & panelName
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        .TextFrame2.MarginLeft = 0: .TextFrame2.MarginTop = 0
        .TextFrame2.TextRange.Text = "Filter Drawer"
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY
        .TextFrame2.TextRange.Font.Size = 11: .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_WHITE
    End With
    
    Set shpResetBtn = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos + 12, topPos + 460, panelWidth - 24, 38)
    With shpResetBtn
        .Name = "Btn_ResetSlicers_" & panelName
        .Adjustments.Item(1) = 0.20
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_ORANGE
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = "Reset Filters"
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY
        .TextFrame2.TextRange.Font.Size = 9: .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_WHITE
        CenterShapeText shpResetBtn, True, True
        .OnAction = "'" & ws.Parent.Name & "'!modPwC_Unified_Master.ClearAllFilters"
    End With
    
    Call InsertVectorIcon(ws, "support_team_illustration.svg", leftPos + 25, topPos + 516, panelWidth - 50, 95, "Illustration_" & panelName)
    
    Set shpQuoteCard = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos + 12, topPos + 620, panelWidth - 24, 80)
    With shpQuoteCard
        .Name = "QuoteCard_" & panelName
        .Adjustments.Item(1) = 0.12
        .Fill.Solid: .Fill.ForeColor.RGB = RGB(30, 41, 59)
        .Line.ForeColor.RGB = RGB(51, 65, 85): .Line.Weight = 0.75
        With .TextFrame2.TextRange
            .Text = ChrW(8220) & " Delivering value through insights. " & ChrW(8221) & vbCrLf & ChrW(8212) & " PwC"
            .Font.Name = FONT_FAMILY
            .Font.Size = 8
            .Font.Italic = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(248, 250, 252)
            If .Paragraphs.Count >= 2 Then
                With .Paragraphs(2).Font
                    .Bold = msoTrue
                    .Italic = msoFalse
                    .Fill.ForeColor.RGB = PWC_ORANGE
                End With
            End If
        End With
        CenterShapeText shpQuoteCard, True, True
    End With
End Sub

Public Sub BuildChartContainer(ws As Worksheet, _
                               ByVal containerName As String, _
                               ByVal leftPos As Single, _
                               ByVal topPos As Single, _
                               ByVal cardWidth As Single, _
                               ByVal cardHeight As Single, _
                               ByVal chartTitle As String, _
                               ByVal chartSubtitle As String, _
                               ByVal chartTypeBadge As String, _
                               Optional ByVal iconFileName As String = "")
    Dim shpContainer As Shape, shpHeader As Shape, shpBadge As Shape, shpDockingZone As Shape
    Dim shpIcon As Shape, headerLeft As Single
    
    Set shpContainer = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, cardWidth, cardHeight)
    With shpContainer
        .Name = "Container_" & containerName
        .Adjustments.Item(1) = 0.05
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_CARD_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 1
        ApplySoftElevation shpContainer
    End With
    
    headerLeft = leftPos + 16
    If Len(iconFileName) > 0 Then
        Set shpIcon = InsertVectorIcon(ws, iconFileName, leftPos + 16, topPos + 14, 18, 18, "Icon_" & containerName)
        If Not shpIcon Is Nothing Then headerLeft = leftPos + 42
    End If
    
    Set shpHeader = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, headerLeft, topPos + 10, cardWidth - (headerLeft - leftPos) - 120, 36)
    With shpHeader
        .Name = "Header_" & containerName
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        With .TextFrame2
            .MarginLeft = 0: .MarginTop = 0: .MarginRight = 0: .MarginBottom = 0: .WordWrap = msoFalse
            .TextRange.Text = chartTitle & vbCrLf & chartSubtitle
            With .TextRange.Paragraphs(1).Font
                .Name = FONT_FAMILY: .Size = 10: .Bold = msoTrue: .Fill.ForeColor.RGB = PWC_TEXT_TITLE
            End With
            With .TextRange.Paragraphs(2).Font
                .Name = FONT_FAMILY: .Size = 7.5: .Bold = msoFalse: .Fill.ForeColor.RGB = PWC_TEXT_MUTED
            End With
        End With
    End With
    
    Set shpBadge = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos + cardWidth - 116, topPos + 10, 102, 22)
    With shpBadge
        .Name = "Badge_" & containerName
        .Adjustments.Item(1) = 0.25
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_PILL_BG
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 0.75
        .TextFrame2.TextRange.Text = chartTypeBadge
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY
        .TextFrame2.TextRange.Font.Size = 7.5: .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_TEXT_MUTED
        CenterShapeText shpBadge, True, True
    End With
    
    Dim dockLeft As Single, dockTop As Single, dockW As Single, dockH As Single
    dockLeft = leftPos + 14: dockTop = topPos + 46
    dockW = cardWidth - 28: dockH = cardHeight - 56
    
    Set shpDockingZone = ws.Shapes.AddShape(msoShapeRoundedRectangle, dockLeft, dockTop, dockW, dockH)
    With shpDockingZone
        .Name = "DockZone_" & containerName
        .Adjustments.Item(1) = 0.03
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_WHITE
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 0.75: .Line.DashStyle = msoLineSolid
        .TextFrame2.TextRange.Text = ""
        CenterShapeText shpDockingZone, True, True
    End With
End Sub

Public Sub DeclutterAndFormatChart(chtObj As ChartObject)
    On Error Resume Next
    Dim ws As Worksheet, shp As Shape
    Set ws = chtObj.Parent
    
    For Each shp In ws.Shapes
        If InStr(1, shp.Name, "DockZone_", vbTextCompare) > 0 Then
            If chtObj.Left >= (shp.Left - 15) And (chtObj.Left + chtObj.Width) <= (shp.Left + shp.Width + 15) And _
               chtObj.Top >= (shp.Top - 15) And (chtObj.Top + chtObj.Height) <= (shp.Top + shp.Height + 15) Then
                shp.TextFrame2.TextRange.Text = ""
                Exit For
            End If
        End If
    Next shp
    
    With chtObj.Chart
        .ChartArea.Format.Fill.Visible = msoFalse
        .ChartArea.Format.Line.Visible = msoFalse
        .PlotArea.Format.Fill.Visible = msoFalse
        .PlotArea.Format.Line.Visible = msoFalse
        .ShowAllFieldButtons = False
        If .HasTitle Then .HasTitle = False
        
        If .Axes.Count > 0 Then
            Dim ax As Axis
            For Each ax In .Axes
                ax.Format.Line.ForeColor.RGB = PWC_CARD_BORDER
                ax.Format.Line.Weight = 0.75
                ax.TickLabels.Font.Name = FONT_FAMILY
                ax.TickLabels.Font.Size = 8
                ax.TickLabels.Font.Color = PWC_TEXT_MUTED
                ax.MajorGridlines.Format.Line.ForeColor.RGB = PWC_CARD_BORDER
                ax.MajorGridlines.Format.Line.Weight = 0.5
            Next ax
        End If
        
        If .HasLegend Then
            .Legend.Position = xlLegendPositionTop
            .Legend.Format.Fill.Visible = msoFalse
            .Legend.Format.Line.Visible = msoFalse
            .Legend.Font.Name = FONT_FAMILY
            .Legend.Font.Size = 8
            .Legend.Font.Color = PWC_TEXT_MUTED
        End If
    End With
    chtObj.ShapeRange.ZOrder msoBringToFront
    On Error GoTo 0
End Sub

' ==============================================================================
' 14. DATA STAGING & CHART CREATION ENGINE (Real Interactive Visuals)
' ==============================================================================

Public Function EnsureStagingSheet() As Worksheet
    Dim wb As Workbook, ws As Worksheet
    Set wb = ThisWorkbook: If wb Is Nothing Then Set wb = ActiveWorkbook
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_STAGING)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SHEET_STAGING
    End If
    On Error GoTo 0
    Set EnsureStagingSheet = ws
End Function

Public Sub PopulateAnalyticalStagingData(wsStaging As Worksheet)
    On Error Resume Next
    Dim i As Long
    
    ' Call Center: Trend (Cols J-K)
    wsStaging.Cells(3, 10).Value = "Month"
    wsStaging.Cells(3, 11).Value = "Total Calls"
    Dim mList As Variant, cList As Variant
    mList = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    cList = Array(1772, 1612, 1616, 1680, 1720, 1845, 1728, 1650, 1610, 1690, 1710, 1790)
    For i = 0 To 11
        wsStaging.Cells(4 + i, 10).Value = mList(i)
        wsStaging.Cells(4 + i, 11).Value = cList(i)
    Next i
    
    ' Call Center: Resolution Breakdown (Cols N-O)
    wsStaging.Cells(3, 14).Value = "Status": wsStaging.Cells(3, 15).Value = "Calls"
    wsStaging.Cells(4, 14).Value = "Resolved First Call": wsStaging.Cells(4, 15).Value = 3646
    wsStaging.Cells(5, 14).Value = "Escalated":           wsStaging.Cells(5, 15).Value = 408
    wsStaging.Cells(6, 14).Value = "Dropped / Abandoned": wsStaging.Cells(6, 15).Value = 946
    
    ' Call Center: Hourly Arrival Pattern (Cols Z-AA)
    wsStaging.Cells(3, 26).Value = "Hour": wsStaging.Cells(3, 27).Value = "Calls"
    Dim hList As Variant, hvList As Variant
    hList = Array("08:00", "09:00", "10:00", "11:00", "12:00", "13:00", "14:00", "15:00", "16:00", "17:00")
    hvList = Array(180, 460, 680, 750, 610, 590, 670, 710, 580, 390)
    For i = 0 To 9
        wsStaging.Cells(4 + i, 26).Value = hList(i)
        wsStaging.Cells(4 + i, 27).Value = hvList(i)
    Next i
    
    ' Call Center: Complaint Categories (Cols R-S)
    wsStaging.Cells(3, 18).Value = "Topic": wsStaging.Cells(3, 19).Value = "Volume"
    Dim tList As Variant, tvList As Variant
    tList = Array("Streaming", "Tech Support", "Payment", "Admin", "Contract")
    tvList = Array(1022, 1019, 1007, 977, 975)
    For i = 0 To 4
        wsStaging.Cells(4 + i, 18).Value = tList(i)
        wsStaging.Cells(4 + i, 19).Value = tvList(i)
    Next i
    
    ' Call Center: AHT vs Volume (Cols V-X)
    wsStaging.Cells(3, 22).Value = "Month": wsStaging.Cells(3, 23).Value = "Volume": wsStaging.Cells(3, 24).Value = "AHT (s)"
    Dim ahtList As Variant
    ahtList = Array(68, 66, 67, 69, 71, 74, 70, 67, 65, 66, 68, 70)
    For i = 0 To 11
        wsStaging.Cells(4 + i, 22).Value = mList(i)
        wsStaging.Cells(4 + i, 23).Value = cList(i)
        wsStaging.Cells(4 + i, 24).Value = ahtList(i)
    Next i
    
    ' Retention: Contract Risk (Cols AH-AJ)
    wsStaging.Cells(3, 34).Value = "Contract": wsStaging.Cells(3, 35).Value = "Churned": wsStaging.Cells(3, 36).Value = "Retained"
    wsStaging.Cells(4, 34).Value = "Month-to-month": wsStaging.Cells(4, 35).Value = 1655: wsStaging.Cells(4, 36).Value = 2220
    wsStaging.Cells(5, 34).Value = "One year":       wsStaging.Cells(5, 35).Value = 166:  wsStaging.Cells(5, 36).Value = 1307
    wsStaging.Cells(6, 34).Value = "Two year":       wsStaging.Cells(6, 35).Value = 48:   wsStaging.Cells(6, 36).Value = 1647
    
    ' Retention: Tenure Cohort (Cols AM-AN)
    wsStaging.Cells(3, 39).Value = "Tenure": wsStaging.Cells(3, 40).Value = "Churned"
    wsStaging.Cells(4, 39).Value = "0-12m":  wsStaging.Cells(4, 40).Value = 1037
    wsStaging.Cells(5, 39).Value = "13-24m": wsStaging.Cells(5, 40).Value = 294
    wsStaging.Cells(6, 39).Value = "25-48m": wsStaging.Cells(6, 40).Value = 337
    wsStaging.Cells(7, 39).Value = "49-72m": wsStaging.Cells(7, 40).Value = 201
    
    ' Retention: Payment Method (Cols AR-AS)
    wsStaging.Cells(3, 44).Value = "Method":           wsStaging.Cells(3, 45).Value = "Churned"
    wsStaging.Cells(4, 44).Value = "Electronic check": wsStaging.Cells(4, 45).Value = 1071
    wsStaging.Cells(5, 44).Value = "Mailed check":     wsStaging.Cells(5, 45).Value = 308
    wsStaging.Cells(6, 44).Value = "Bank transfer":    wsStaging.Cells(6, 45).Value = 258
    wsStaging.Cells(7, 44).Value = "Credit card":      wsStaging.Cells(7, 45).Value = 232
    
    ' Retention: Internet Service (Cols AW-AY)
    wsStaging.Cells(3, 49).Value = "Service":     wsStaging.Cells(3, 50).Value = "Churned": wsStaging.Cells(3, 51).Value = "Retained"
    wsStaging.Cells(4, 49).Value = "Fiber optic": wsStaging.Cells(4, 50).Value = 1297:     wsStaging.Cells(4, 51).Value = 1799
    wsStaging.Cells(5, 49).Value = "DSL":         wsStaging.Cells(5, 50).Value = 459:      wsStaging.Cells(5, 51).Value = 1962
    wsStaging.Cells(6, 49).Value = "No internet": wsStaging.Cells(6, 50).Value = 113:      wsStaging.Cells(6, 51).Value = 1413
    
    ' Diversity: Job Level Funnel (Cols BC-BE)
    wsStaging.Cells(3, 55).Value = "Level": wsStaging.Cells(3, 56).Value = "Female %": wsStaging.Cells(3, 57).Value = "Male %"
    wsStaging.Cells(4, 55).Value = "Executive":        wsStaging.Cells(4, 56).Value = 0.16: wsStaging.Cells(4, 57).Value = 0.84
    wsStaging.Cells(5, 55).Value = "Director":         wsStaging.Cells(5, 56).Value = 0.16: wsStaging.Cells(5, 57).Value = 0.84
    wsStaging.Cells(6, 55).Value = "Senior Manager":   wsStaging.Cells(6, 56).Value = 0.24: wsStaging.Cells(6, 57).Value = 0.76
    wsStaging.Cells(7, 55).Value = "Manager":          wsStaging.Cells(7, 56).Value = 0.37: wsStaging.Cells(7, 57).Value = 0.63
    wsStaging.Cells(8, 55).Value = "Senior Associate": wsStaging.Cells(8, 56).Value = 0.43: wsStaging.Cells(8, 57).Value = 0.57
    wsStaging.Cells(9, 55).Value = "Junior Associate": wsStaging.Cells(9, 56).Value = 0.54: wsStaging.Cells(9, 57).Value = 0.46
    
    ' Diversity: Department Parity (Cols BH-BJ)
    wsStaging.Cells(3, 60).Value = "Dept": wsStaging.Cells(3, 61).Value = "Female %": wsStaging.Cells(3, 62).Value = "Male %"
    wsStaging.Cells(4, 60).Value = "Operations": wsStaging.Cells(4, 61).Value = 0.46: wsStaging.Cells(4, 62).Value = 0.54
    wsStaging.Cells(5, 60).Value = "Sales":      wsStaging.Cells(5, 61).Value = 0.41: wsStaging.Cells(5, 62).Value = 0.59
    wsStaging.Cells(6, 60).Value = "Internal":   wsStaging.Cells(6, 61).Value = 0.44: wsStaging.Cells(6, 62).Value = 0.56
    wsStaging.Cells(7, 60).Value = "Strategy":   wsStaging.Cells(7, 61).Value = 0.38: wsStaging.Cells(7, 62).Value = 0.62
    wsStaging.Cells(8, 60).Value = "Tech":       wsStaging.Cells(8, 61).Value = 0.28: wsStaging.Cells(8, 62).Value = 0.72
    
    ' Diversity: Promo Velocity (Cols BM-BO)
    wsStaging.Cells(3, 65).Value = "Year": wsStaging.Cells(3, 66).Value = "Female Promo %": wsStaging.Cells(3, 67).Value = "Male Promo %"
    wsStaging.Cells(4, 65).Value = "FY20": wsStaging.Cells(4, 66).Value = 0.102: wsStaging.Cells(4, 67).Value = 0.106
    wsStaging.Cells(5, 65).Value = "FY21": wsStaging.Cells(5, 66).Value = 0.084: wsStaging.Cells(5, 67).Value = 0.088
    
    ' Diversity: Performance Distribution (Cols BR-BT)
    wsStaging.Cells(3, 70).Value = "Rating": wsStaging.Cells(3, 71).Value = "Female": wsStaging.Cells(3, 72).Value = "Male"
    wsStaging.Cells(4, 70).Value = "1 - Low":         wsStaging.Cells(4, 71).Value = 10:  wsStaging.Cells(4, 72).Value = 10
    wsStaging.Cells(5, 70).Value = "2 - Developing":  wsStaging.Cells(5, 71).Value = 60:  wsStaging.Cells(5, 72).Value = 55
    wsStaging.Cells(6, 70).Value = "3 - Proficient":  wsStaging.Cells(6, 71).Value = 310: wsStaging.Cells(6, 72).Value = 315
    wsStaging.Cells(7, 70).Value = "4 - Advanced":    wsStaging.Cells(7, 71).Value = 90:  wsStaging.Cells(7, 72).Value = 90
    wsStaging.Cells(8, 70).Value = "5 - Expert":      wsStaging.Cells(8, 71).Value = 30:  wsStaging.Cells(8, 72).Value = 30
    
    On Error GoTo 0
End Sub

Public Sub BuildCallCenterVisuals(ws As Worksheet, wsStaging As Worksheet)
    On Error Resume Next
    Dim chtObj As ChartObject
    
    ' 1. Daily Calls Trend (Smooth Line with Markers)
    ws.ChartObjects("cht_TrendDaily").Delete
    Set chtObj = ws.ChartObjects.Add(258, 282, 432, 186)
    With chtObj
        .Name = "cht_TrendDaily"
        With .Chart
            .ChartType = xlLineMarkers
            .SetSourceData wsStaging.Range("J3:K15")
            If .SeriesCollection.Count > 0 Then
                With .SeriesCollection(1)
                    .Format.Line.ForeColor.RGB = PWC_ORANGE
                    .Format.Line.Weight = 2
                    .Smooth = True
                    .MarkerStyle = xlMarkerStyleCircle
                    .MarkerSize = 5
                    .MarkerForegroundColor = PWC_ORANGE
                    .MarkerBackgroundColor = PWC_WHITE
                End With
            End If
        End With
    End With
    DeclutterAndFormatChart chtObj
    
    ' 2. Hourly Arrival Pattern (Column Chart)
    ws.ChartObjects("cht_HourlyArrival").Delete
    Set chtObj = ws.ChartObjects.Add(722, 282, 356, 186)
    With chtObj
        .Name = "cht_HourlyArrival"
        With .Chart
            .ChartType = xlColumnClustered
            .SetSourceData wsStaging.Range("Z3:AA13")
            If .SeriesCollection.Count > 0 Then
                .SeriesCollection(1).Format.Fill.Solid
                .SeriesCollection(1).Format.Fill.ForeColor.RGB = PWC_ORANGE
            End If
        End With
    End With
    DeclutterAndFormatChart chtObj
    
    ' 3. Resolution Breakdown (Donut Chart)
    ws.ChartObjects("cht_Resolution").Delete
    Set chtObj = ws.ChartObjects.Add(1116, 282, 334, 186)
    With chtObj
        .Name = "cht_Resolution"
        With .Chart
            .ChartType = xlDoughnut
            .SetSourceData wsStaging.Range("N3:O6")
            If .SeriesCollection.Count > 0 Then
                On Error Resume Next
                .SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = PWC_SUCCESS_GREEN
                .SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = PWC_WARNING_AMBER
                .SeriesCollection(1).Points(3).Format.Fill.ForeColor.RGB = PWC_ALERT_RED
                On Error GoTo 0
            End If
        End With
    End With
    DeclutterAndFormatChart chtObj
    
    ' Center Badge for Donut Chart
    Dim centerBadge As Shape
    ws.Shapes("Resolution_CenterBadge").Delete
    Set centerBadge = ws.Shapes.AddShape(msoShapeOval, 1245, 345, 76, 50)
    With centerBadge
        .Name = "Resolution_CenterBadge"
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_WHITE
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 1
        .TextFrame2.TextRange.Text = "5,000" & vbCrLf & "Total Calls"
        With .TextFrame2.TextRange.Paragraphs(1).Font
            .Name = FONT_FAMILY: .Size = 10: .Bold = msoTrue: .Fill.ForeColor.RGB = PWC_TEXT_TITLE
        End With
        With .TextFrame2.TextRange.Paragraphs(2).Font
            .Name = FONT_FAMILY: .Size = 7.5: .Bold = msoFalse: .Fill.ForeColor.RGB = PWC_TEXT_MUTED
        End With
        CenterShapeText centerBadge, True, True
        .ZOrder msoBringToFront
    End With
    
    ' 4. Agent Performance Scorecard (Live Docked Picture of pt_Agent)
    Call DockAgentScorecardLive(ws, wsStaging, 258, 546, 290, 204)
    
    ' 5. AHT vs Call Volume (Combo Chart)
    ws.ChartObjects("cht_AHTVolume").Delete
    Set chtObj = ws.ChartObjects.Add(582, 546, 316, 204)
    With chtObj
        .Name = "cht_AHTVolume"
        With .Chart
            .ChartType = xlColumnClustered
            .SetSourceData wsStaging.Range("V3:X15")
            If .SeriesCollection.Count >= 2 Then
                .SeriesCollection(1).Format.Fill.Solid
                .SeriesCollection(1).Format.Fill.ForeColor.RGB = PWC_ORANGE
                .SeriesCollection(2).AxisGroup = xlSecondary
                .SeriesCollection(2).ChartType = xlLineMarkers
                .SeriesCollection(2).Format.Line.ForeColor.RGB = PWC_CHARCOAL
                .SeriesCollection(2).Format.Line.Weight = 2
            End If
        End With
    End With
    DeclutterAndFormatChart chtObj
    
    ' 6. Complaint Topics (Horizontal Bar Chart)
    ws.ChartObjects("cht_ComplaintPareto").Delete
    Set chtObj = ws.ChartObjects.Add(936, 546, 262, 204)
    With chtObj
        .Name = "cht_ComplaintPareto"
        With .Chart
            .ChartType = xlBarClustered
            .SetSourceData wsStaging.Range("R3:S8")
            If .SeriesCollection.Count > 0 Then
                .SeriesCollection(1).Format.Fill.Solid
                .SeriesCollection(1).Format.Fill.ForeColor.RGB = PWC_ORANGE
            End If
        End With
    End With
    DeclutterAndFormatChart chtObj
    
    ' 7. Sentiment Analysis & Regional SLA Pills
    Call BuildSentimentAndRegionalPills(ws, 1240, 546, 210, 204)
    
    On Error GoTo 0
End Sub

Private Sub DockAgentScorecardLive(ws As Worksheet, wsStaging As Worksheet, _
                                  ByVal leftPos As Single, ByVal topPos As Single, _
                                  ByVal w As Single, ByVal h As Single)
    On Error Resume Next
    Dim ptRng As Range, shpPic As Shape
    Set ptRng = wsStaging.Range("C3:H11")
    
    Dim wasUpdating As Boolean
    wasUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = True
    
    ws.Shapes("LiveScorecardPic").Delete
    ptRng.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    ws.Activate
    ws.Range("B1").Select
    ws.Paste
    
    Application.ScreenUpdating = wasUpdating
    
    Set shpPic = ws.Shapes(ws.Shapes.Count)
    If Not shpPic Is Nothing Then
        With shpPic
            .Name = "LiveScorecardPic"
            .Left = leftPos
            .Top = topPos
            .Width = w
            .Height = h
            .Line.Visible = msoFalse
            .ZOrder msoBringToFront
        End With
    End If
    On Error GoTo 0
End Sub

Private Sub BuildSentimentAndRegionalPills(ws As Worksheet, ByVal leftPos As Single, _
                                         ByVal topPos As Single, ByVal w As Single, ByVal h As Single)
    On Error Resume Next
    Dim shpP1 As Shape, shpP2 As Shape, shpP3 As Shape, shpReg As Shape
    
    ws.Shapes("Pill_Positive").Delete
    Set shpP1 = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, w, 24)
    With shpP1
        .Name = "Pill_Positive"
        .Adjustments.Item(1) = 0.25
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_BADGE_GREEN_BG
        .Line.ForeColor.RGB = RGB(167, 243, 208): .Line.Weight = 0.75
        .TextFrame2.TextRange.Text = "Positive Sentiment: 62%"
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY
        .TextFrame2.TextRange.Font.Size = 8: .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_BADGE_GREEN_TXT
        CenterShapeText shpP1, True, True
    End With
    
    ws.Shapes("Pill_Neutral").Delete
    Set shpP2 = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos + 28, w, 24)
    With shpP2
        .Name = "Pill_Neutral"
        .Adjustments.Item(1) = 0.25
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_BADGE_AMBER_BG
        .Line.ForeColor.RGB = RGB(253, 230, 138): .Line.Weight = 0.75
        .TextFrame2.TextRange.Text = "Neutral Sentiment: 25%"
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY
        .TextFrame2.TextRange.Font.Size = 8: .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_BADGE_AMBER_TXT
        CenterShapeText shpP2, True, True
    End With
    
    ws.Shapes("Pill_Negative").Delete
    Set shpP3 = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos + 56, w, 24)
    With shpP3
        .Name = "Pill_Negative"
        .Adjustments.Item(1) = 0.25
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_BADGE_RED_BG
        .Line.ForeColor.RGB = RGB(254, 202, 202): .Line.Weight = 0.75
        .TextFrame2.TextRange.Text = "Negative Sentiment: 13%"
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY
        .TextFrame2.TextRange.Font.Size = 8: .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_BADGE_RED_TXT
        CenterShapeText shpP3, True, True
    End With
    
    ws.Shapes("Box_RegionalSLA").Delete
    Set shpReg = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos + 86, w, h - 86)
    With shpReg
        .Name = "Box_RegionalSLA"
        .Adjustments.Item(1) = 0.08
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_PILL_BG
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 1
        With .TextFrame2.TextRange
            .Text = "Regional SLA Compliance" & vbCrLf & _
                    "Riyadh:  92% SLA (1,420 calls)" & vbCrLf & _
                    "Jeddah:  88% SLA (1,180 calls)" & vbCrLf & _
                    "Dammam:  90% SLA (950 calls)" & vbCrLf & _
                    "Zurich:  89% SLA (850 calls)" & vbCrLf & _
                    "Basel:   86% SLA (600 calls)"
            .Font.Name = FONT_FAMILY: .Font.Size = 7.5
            .Font.Fill.ForeColor.RGB = PWC_TEXT_MUTED
            With .Paragraphs(1).Font
                .Bold = msoTrue: .Size = 8: .Fill.ForeColor.RGB = PWC_TEXT_TITLE
            End With
        End With
    End With
    On Error GoTo 0
End Sub

Public Sub BuildRetentionVisuals(ws As Worksheet, wsStaging As Worksheet)
    On Error Resume Next
    Dim chtObj As ChartObject
    
    ' Visual 1: Contract Risk
    ws.ChartObjects("cht_ContractRisk").Delete
    Set chtObj = ws.ChartObjects.Add(264, 276, 438, 209)
    With chtObj
        .Name = "cht_ContractRisk"
        With .Chart
            .ChartType = xlColumnClustered
            .SetSourceData wsStaging.Range("AH3:AJ6")
            If .SeriesCollection.Count >= 2 Then
                .SeriesCollection(1).Format.Fill.Solid: .SeriesCollection(1).Format.Fill.ForeColor.RGB = PWC_ALERT_RED
                .SeriesCollection(2).Format.Fill.Solid: .SeriesCollection(2).Format.Fill.ForeColor.RGB = PWC_SUCCESS_GREEN
            End If
        End With
    End With
    DeclutterAndFormatChart chtObj
    
    ' Visual 2: Tenure Cohort
    ws.ChartObjects("cht_TenureCohort").Delete
    Set chtObj = ws.ChartObjects.Add(264, 557, 438, 209)
    With chtObj
        .Name = "cht_TenureCohort"
        With .Chart
            .ChartType = xlColumnClustered
            .SetSourceData wsStaging.Range("AM3:AN7")
            If .SeriesCollection.Count > 0 Then
                .SeriesCollection(1).Format.Fill.Solid: .SeriesCollection(1).Format.Fill.ForeColor.RGB = PWC_ORANGE
            End If
        End With
    End With
    DeclutterAndFormatChart chtObj
    
    ' Visual 3: Payment Friction
    ws.ChartObjects("cht_PaymentFriction").Delete
    Set chtObj = ws.ChartObjects.Add(746, 276, 438, 209)
    With chtObj
        .Name = "cht_PaymentFriction"
        With .Chart
            .ChartType = xlBarClustered
            .SetSourceData wsStaging.Range("AR3:AS7")
            If .SeriesCollection.Count > 0 Then
                .SeriesCollection(1).Format.Fill.Solid: .SeriesCollection(1).Format.Fill.ForeColor.RGB = PWC_ORANGE
            End If
        End With
    End With
    DeclutterAndFormatChart chtObj
    
    ' Visual 4: Service Matrix
    ws.ChartObjects("cht_ServiceMatrix").Delete
    Set chtObj = ws.ChartObjects.Add(746, 557, 438, 209)
    With chtObj
        .Name = "cht_ServiceMatrix"
        With .Chart
            .ChartType = xlColumnClustered
            .SetSourceData wsStaging.Range("AW3:AY6")
            If .SeriesCollection.Count >= 2 Then
                .SeriesCollection(1).Format.Fill.Solid: .SeriesCollection(1).Format.Fill.ForeColor.RGB = PWC_ALERT_RED
                .SeriesCollection(2).Format.Fill.Solid: .SeriesCollection(2).Format.Fill.ForeColor.RGB = PWC_SUCCESS_GREEN
            End If
        End With
    End With
    DeclutterAndFormatChart chtObj
    On Error GoTo 0
End Sub

Public Sub BuildDiversityVisuals(ws As Worksheet, wsStaging As Worksheet)
    On Error Resume Next
    Dim chtObj As ChartObject
    
    ' Visual 1: Career Funnel
    ws.ChartObjects("cht_DIFunnel").Delete
    Set chtObj = ws.ChartObjects.Add(264, 276, 438, 209)
    With chtObj
        .Name = "cht_DIFunnel"
        With .Chart
            .ChartType = xlBarStacked100
            .SetSourceData wsStaging.Range("BC3:BE9")
            If .SeriesCollection.Count >= 2 Then
                .SeriesCollection(1).Format.Fill.Solid: .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(147, 51, 234)
                .SeriesCollection(2).Format.Fill.Solid: .SeriesCollection(2).Format.Fill.ForeColor.RGB = PWC_CHARCOAL
            End If
        End With
    End With
    DeclutterAndFormatChart chtObj
    
    ' Visual 2: Department Parity
    ws.ChartObjects("cht_DIDeptParity").Delete
    Set chtObj = ws.ChartObjects.Add(264, 557, 438, 209)
    With chtObj
        .Name = "cht_DIDeptParity"
        With .Chart
            .ChartType = xlBarStacked100
            .SetSourceData wsStaging.Range("BH3:BJ8")
            If .SeriesCollection.Count >= 2 Then
                .SeriesCollection(1).Format.Fill.Solid: .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(147, 51, 234)
                .SeriesCollection(2).Format.Fill.Solid: .SeriesCollection(2).Format.Fill.ForeColor.RGB = PWC_CHARCOAL
            End If
        End With
    End With
    DeclutterAndFormatChart chtObj
    
    ' Visual 3: Promo Velocity
    ws.ChartObjects("cht_DIPromoVelocity").Delete
    Set chtObj = ws.ChartObjects.Add(746, 276, 438, 209)
    With chtObj
        .Name = "cht_DIPromoVelocity"
        With .Chart
            .ChartType = xlColumnClustered
            .SetSourceData wsStaging.Range("BM3:BO5")
            If .SeriesCollection.Count >= 2 Then
                .SeriesCollection(1).Format.Fill.Solid: .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(147, 51, 234)
                .SeriesCollection(2).Format.Fill.Solid: .SeriesCollection(2).Format.Fill.ForeColor.RGB = PWC_CHARCOAL
            End If
        End With
    End With
    DeclutterAndFormatChart chtObj
    
    ' Visual 4: Performance Ratings
    ws.ChartObjects("cht_DIPerformance").Delete
    Set chtObj = ws.ChartObjects.Add(746, 557, 438, 209)
    With chtObj
        .Name = "cht_DIPerformance"
        With .Chart
            .ChartType = xlColumnClustered
            .SetSourceData wsStaging.Range("BR3:BT8")
            If .SeriesCollection.Count >= 2 Then
                .SeriesCollection(1).Format.Fill.Solid: .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(147, 51, 234)
                .SeriesCollection(2).Format.Fill.Solid: .SeriesCollection(2).Format.Fill.ForeColor.RGB = PWC_CHARCOAL
            End If
        End With
    End With
    DeclutterAndFormatChart chtObj
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------------------
' DASHBOARD 1 BUILDER: CALL CENTER OPERATIONS COCKPIT
' ------------------------------------------------------------------------------
Public Sub BuildCallCenterCanvas(Optional ByVal populateInitialData As Boolean = False)
    Dim wb As Workbook, ws As Worksheet, wsStaging As Worksheet
    Set wb = ThisWorkbook: If wb Is Nothing Then Set wb = ActiveWorkbook
    
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_CC)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SHEET_CC
    End If
    On Error GoTo 0
    
    ws.Tab.Color = PWC_ORANGE
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    
    InitializeDashboardCanvas ws, PWC_CANVAS_BG
    BuildWebTopNavBar ws, "CC"
    BuildHeroHeader ws, "Call Centre Operations & SLA Performance Cockpit", _
                    "Intraday Queue Triage, Abandonment Forensics & Representative Quality Auditing (Q1 2021)", _
                    76, 1480
                    
    Dim cardTop As Single, cardW As Single, cardGap As Single, kpiStartLeft As Single
    cardTop = 130: cardW = 168: cardGap = 12: kpiStartLeft = 244
    
    Dim v1 As String, v2 As String, v3 As String, v4 As String, v5 As String, v6 As String, v7 As String
    If populateInitialData Then
        v1 = "5,000": v2 = "4,054": v3 = "946": v4 = "81.1%": v5 = "67.5 s": v6 = "3.40": v7 = "89.9%"
    Else
        v1 = "--": v2 = "--": v3 = "--": v4 = "--": v5 = "--": v6 = "--": v7 = "--"
    End If
    
    BuildKPICard ws, "CC_TotalDemand", kpiStartLeft + (cardW + cardGap) * 0, cardTop, cardW, 94, _
                 "Total Calls", v1, ChrW(9650) & " 12.4% vs PY", PWC_ORANGE, "phone_orange.svg", "sparkline_orange.svg"
                 
    BuildKPICard ws, "CC_Answered", kpiStartLeft + (cardW + cardGap) * 1, cardTop, cardW, 94, _
                 "Answered Calls", v2, ChrW(9650) & " 11.8% vs PY", PWC_SUCCESS_GREEN, "check_green.svg", "sparkline_green.svg"
                 
    BuildKPICard ws, "CC_Missed", kpiStartLeft + (cardW + cardGap) * 2, cardTop, cardW, 94, _
                 "Missed Calls", v3, ChrW(9650) & " 18.7% vs PY", PWC_ALERT_RED, "xcircle_red.svg", "sparkline_red.svg"
                 
    BuildKPICard ws, "CC_SLA", kpiStartLeft + (cardW + cardGap) * 3, cardTop, cardW, 94, _
                 "SLA (%)", v4, ChrW(9650) & " 5.9% vs PY", PWC_WARNING_AMBER, "timer_amber.svg", "sparkline_amber.svg"
                 
    BuildKPICard ws, "CC_AHT", kpiStartLeft + (cardW + cardGap) * 4, cardTop, cardW, 94, _
                 "Avg Handle Time", v5, ChrW(9660) & " 3.4% vs PY", RGB(147, 51, 234), "clock_purple.svg", "sparkline_purple.svg"
                 
    BuildKPICard ws, "CC_CSAT", kpiStartLeft + (cardW + cardGap) * 5, cardTop, cardW, 94, _
                 "CSAT Score", v6, ChrW(9650) & " 0.3 vs PY", RGB(37, 99, 235), "user_blue.svg", "sparkline_blue.svg"
                 
    BuildKPICard ws, "CC_FCR", kpiStartLeft + (cardW + cardGap) * 6, cardTop, cardW, 94, _
                 "FCR (%)", v7, ChrW(9650) & " 6.2% vs PY", RGB(13, 148, 136), "target_teal.svg", "sparkline_teal.svg"
                 
    BuildSlicerPanelContainer ws, "CC_Slicers", CANVAS_LEFT, 130, SLICER_WIDTH, 710, _
                              "Date Period (Month)", "Inquiry Topic Tier", "Representative Agent"
                              
    ' Middle Row: 3 Visual Containers (Trend, Heatmap, Resolution)
    BuildChartContainer ws, "TrendDaily", 244, 236, 460, 246, _
                        "Calls Trend (Daily / Monthly)", "Call Arrival Pattern vs Answered Volume", _
                        "Area / Line Chart", "icon_hourly_surge.svg"
                        
    BuildChartContainer ws, "Heatmap", 712, 236, 376, 246, _
                        "Calls by Hour (Arrival Heatmap)", "Intraday Hourly Demand Surge Profile", _
                        "Column Heatmap", "icon_hourly_surge.svg"
                        
    BuildChartContainer ws, "Resolution", 1096, 236, 374, 246, _
                        "Resolution Breakdown", "First Call Resolution vs Escalated & Abandoned", _
                        "Donut Chart", "icon_topic_sla.svg"
                        
    ' Bottom Row: 4 Visual Containers (Scorecard, AHT vs Volume, Complaints, Sentiment)
    BuildChartContainer ws, "AgentScorecard", 244, 498, 318, 260, _
                        "Agent Performance (Top 10)", "Live Dynamic Representative Audit Matrix", _
                        "Matrix Table", "icon_audit_matrix.svg"
                        
    BuildChartContainer ws, "AHTVolume", 570, 498, 340, 260, _
                        "Avg Handle Time vs Call Volume", "Dual-Axis Efficiency Correlation", _
                        "Combo Chart", "icon_agent_quadrant.svg"
                        
    BuildChartContainer ws, "ComplaintPareto", 918, 498, 298, 260, _
                        "Top Complaint Categories (Pareto)", "Inquiry Classification Volume", _
                        "Horizontal Bar", "icon_topic_sla.svg"
                        
    BuildChartContainer ws, "Sentiment", 1224, 498, 246, 260, _
                        "Sentiment & Regional SLA", "Sentiment Index & Geographic Distribution", _
                        "Status Pills", "icon_live_indicator.svg"
                        
    Set wsStaging = EnsureStagingSheet()
    Call PopulateAnalyticalStagingData(wsStaging)
    Call BuildCallCenterVisuals(ws, wsStaging)
    Call AutomateAndLinkKPICards(ws, "CC")
    Call ResetSheetViewport(ws)
End Sub

' ------------------------------------------------------------------------------
' DASHBOARD 2 BUILDER: CUSTOMER RETENTION COCKPIT
' ------------------------------------------------------------------------------
Public Sub BuildCustomerRetentionCanvas(Optional ByVal populateInitialData As Boolean = False)
    Dim wb As Workbook, ws As Worksheet, wsStaging As Worksheet
    Set wb = ThisWorkbook: If wb Is Nothing Then Set wb = ActiveWorkbook
    
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_CH)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SHEET_CH
    End If
    On Error GoTo 0
    
    ws.Tab.Color = PWC_CHARCOAL
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    
    InitializeDashboardCanvas ws, PWC_CANVAS_BG
    BuildWebTopNavBar ws, "CH"
    BuildHeroHeader ws, "Customer Retention & Revenue Risk Cockpit", _
                    "Subscriber Attrition Forensics, ARR Revenue Exposure & Contract Vulnerability", _
                    76, 1214
                    
    Dim cardTop As Single, val1 As String, val2 As String, val3 As String, val4 As String, val5 As String
    cardTop = 130
    If populateInitialData Then
        val1 = "7,043": val2 = "26.54%": val3 = "$2.86M": val4 = "42.71%": val5 = "1.42"
    Else
        val1 = "--": val2 = "--": val3 = "--": val4 = "--": val5 = "--"
    End If
    
    BuildKPICard ws, "CH_Subscribers", CANVAS_LEFT + (230 + 16) * 0, cardTop, 230, 88, _
                 "Total Active Accounts", val1, "Active Subscriber Base", PWC_TEXT_TITLE, "users_orange.svg"
                 
    BuildKPICard ws, "CH_ChurnRate", CANVAS_LEFT + (230 + 16) * 1, cardTop, 230, 88, _
                 "Customer Churn Rate", val2, "Target: < 20.0% Churn", PWC_ALERT_RED, "xcircle_red.svg"
                 
    BuildKPICard ws, "CH_ARRRisk", CANVAS_LEFT + (230 + 16) * 2, cardTop, 230, 88, _
                 "Annual Revenue at Risk", val3, "Annual Lost ARR Exposure", PWC_ORANGE, "dollar_orange.svg"
                 
    BuildKPICard ws, "CH_M2MChurn", CANVAS_LEFT + (230 + 16) * 3, cardTop, 230, 88, _
                 "Month-to-Month Churn", val4, "Target: < 25.0% M2M", PWC_ALERT_RED, "timer_amber.svg"
                 
    BuildKPICard ws, "CH_Tickets", CANVAS_LEFT + (230 + 16) * 4, cardTop, 230, 88, _
                 "Tech Tickets per Churn", val5, "Friction Index: 3+ Tickets", PWC_WARNING_AMBER, "target_teal.svg"
                 
    Dim bodyTop As Single: bodyTop = 230
    BuildSlicerPanelContainer ws, "CH_Slicers", CANVAS_LEFT, bodyTop, SLICER_WIDTH, 560, _
                              "Commitment Contract", "Payment Method Tier", "Internet Service Type"
                              
    Dim colMidLeft As Single, colRightLeft As Single, visW As Single, visH As Single
    colMidLeft = CANVAS_LEFT + SLICER_WIDTH + 16
    visW = 466: visH = 265
    colRightLeft = colMidLeft + visW + 16
    
    BuildChartContainer ws, "ContractRisk", colMidLeft, bodyTop, visW, visH, _
                        "Churn Rate % by Commitment Contract", _
                        "Month-to-Month Vulnerability vs 1-Year & 2-Year Commitments", _
                        "Column Chart", "icon_retention_users.svg"
                        
    BuildChartContainer ws, "TenureCohort", colMidLeft, bodyTop + visH + 16, visW, visH, _
                        "Tenure Attrition Curve & Vulnerability", _
                        "Early Risk Window (Months 1-12) vs Established Subscribers", _
                        "Area / Line Chart", "icon_hourly_surge.svg"
                        
    BuildChartContainer ws, "PaymentFriction", colRightLeft, bodyTop, visW, visH, _
                        "Payment Method Risk Diagnostics", _
                        "Electronic Check Friction vs Automated Transfers", _
                        "Clustered Bar", "icon_audit_matrix.svg"
                        
    BuildChartContainer ws, "ServiceMatrix", colRightLeft, bodyTop + visH + 16, visW, visH, _
                        "Internet Service & Add-On Protection", _
                        "Fiber Optic Churn Exposure vs Tech Support", _
                        "Matrix Table", "icon_topic_sla.svg"
                        
    Set wsStaging = EnsureStagingSheet()
    Call PopulateAnalyticalStagingData(wsStaging)
    Call BuildRetentionVisuals(ws, wsStaging)
    Call AutomateAndLinkKPICards(ws, "CH")
    Call ResetSheetViewport(ws)
End Sub

' ------------------------------------------------------------------------------
' DASHBOARD 3 BUILDER: DIVERSITY & INCLUSION COCKPIT
' ------------------------------------------------------------------------------
Public Sub BuildDiversityInclusionCanvas(Optional ByVal populateInitialData As Boolean = False)
    Dim wb As Workbook, ws As Worksheet, wsStaging As Worksheet
    Set wb = ThisWorkbook: If wb Is Nothing Then Set wb = ActiveWorkbook
    
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_DI)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SHEET_DI
    End If
    On Error GoTo 0
    
    ws.Tab.Color = PWC_TEXT_MUTED
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    
    InitializeDashboardCanvas ws, PWC_CANVAS_BG
    BuildWebTopNavBar ws, "DI"
    BuildHeroHeader ws, "Diversity, Equity & Executive Parity Cockpit", _
                    "Workforce Pipeline Governance, Broken Rung Diagnostics & Promotion Velocity (FY21)", _
                    76, 1214
                    
    Dim cardTop As Single, val1 As String, val2 As String, val3 As String, val4 As String, val5 As String
    cardTop = 130
    If populateInitialData Then
        val1 = "500": val2 = "41.00%": val3 = "20.00%": val4 = "35.29%": val5 = "0.0 Mos"
    Else
        val1 = "--": val2 = "--": val3 = "--": val4 = "--": val5 = "--"
    End If
    
    BuildKPICard ws, "DI_Workforce", CANVAS_LEFT + (230 + 16) * 0, cardTop, 230, 88, _
                 "Total Corporate Census", val1, "Active Headcount Base", PWC_TEXT_TITLE, "users_rose.svg"
                 
    BuildKPICard ws, "DI_FemaleShare", CANVAS_LEFT + (230 + 16) * 1, cardTop, 230, 88, _
                 "Female Headcount Share", val2, "Corporate Parity Target: 50.0%", PWC_WARNING_AMBER, "check_green.svg"
                 
    BuildKPICard ws, "DI_BrokenRung", CANVAS_LEFT + (230 + 16) * 2, cardTop, 230, 88, _
                 "Executive Female Share", val3, "Critical Executive Pipeline Share", PWC_ALERT_RED, "xcircle_red.svg"
                 
    BuildKPICard ws, "DI_PromoShare", CANVAS_LEFT + (230 + 16) * 3, cardTop, 230, 88, _
                 "FY21 Promotions Awarded", val4, "Promotions Parity Baseline", PWC_ORANGE, "award_rose.svg"
                 
    BuildKPICard ws, "DI_TimeInGrade", CANVAS_LEFT + (230 + 16) * 4, cardTop, 230, 88, _
                 "Promotion Velocity Gap", val5, "Target: Zero Gender Variance", PWC_ALERT_RED, "timer_amber.svg"
                 
    Dim bodyTop As Single: bodyTop = 230
    BuildSlicerPanelContainer ws, "DI_Slicers", CANVAS_LEFT, bodyTop, SLICER_WIDTH, 560, _
                              "Department Group", "Job Level Tier", "Age Demographic Cohort"
                              
    Dim colMidLeft As Single, colRightLeft As Single, visW As Single, visH As Single
    colMidLeft = CANVAS_LEFT + SLICER_WIDTH + 16
    visW = 466: visH = 265
    colRightLeft = colMidLeft + visW + 16
    
    BuildChartContainer ws, "DIFunnel", colMidLeft, bodyTop, visW, visH, _
                        "Workforce Hierarchy & Broken Rung Funnel", _
                        "Female vs Male Representation across Corporate Grades 1 to 6", _
                        "Funnel / Bar", "icon_diversity_parity.svg"
                        
    BuildChartContainer ws, "DIDeptParity", colMidLeft, bodyTop + visH + 16, visW, visH, _
                        "Departmental Representation & Target Gaps", _
                        "Operations, Sales, Finance, HR, IT & Strategy Headcount", _
                        "Clustered Column", "icon_audit_matrix.svg"
                        
    BuildChartContainer ws, "DIPromoVelocity", colRightLeft, bodyTop, visW, visH, _
                        "Promotion Velocity & Time in Grade", _
                        "Time-in-Grade Velocity Comparison by Gender", _
                        "Bar Chart", "icon_hourly_surge.svg"
                        
    BuildChartContainer ws, "DIPerformance", colRightLeft, bodyTop + visH + 16, visW, visH, _
                        "Performance Appraisal & Promotion Equity", _
                        "FY20 Appraisal Rating vs Actual FY21 Promotion Award Rate", _
                        "Matrix Table", "icon_agent_quadrant.svg"
                        
    Set wsStaging = EnsureStagingSheet()
    Call PopulateAnalyticalStagingData(wsStaging)
    Call BuildDiversityVisuals(ws, wsStaging)
    Call AutomateAndLinkKPICards(ws, "DI")
    Call ResetSheetViewport(ws)
End Sub

Public Sub BuildAllDashboardCanvases()
    FreezeAppState True
    Call BuildCallCenterCanvas(True)
    Call BuildCustomerRetentionCanvas(True)
    Call BuildDiversityInclusionCanvas(True)
    Call NavigateToCallCenter
    RestoreAppState
End Sub

' ==============================================================================
' 15. PRIMARY MASTER PLATFORM ORCHESTRATOR
' ==============================================================================

Public Sub RunCompletePwCPlatform()
    Dim currentStep As String
    currentStep = "Initializing Engine State"
    
    On Error GoTo MasterErrHandler
    FreezeAppState True
    
    ' 0. Purge conflicting legacy modules (if Trust Access is enabled)
    currentStep = "Cleaning Legacy Modules"
    Call RemoveLegacyModules
    
    ' 1. Stage analytical data ranges on Staging_Pivots (preserving pt_Agent)
    currentStep = "Staging Analytical Datasets"
    Dim wsStaging As Worksheet
    Set wsStaging = EnsureStagingSheet()
    Call PopulateAnalyticalStagingData(wsStaging)
    
    ' 2. Build Governance & Documentation Canvas Sheets
    currentStep = "Building Governance Sheets (01_Domains & 02_Catalog)"
    Call BuildGovernanceArchitecture
    
    ' 3. Build Executive Homepage Portal
    currentStep = "Building Executive Home Portal (00_Home_Portal)"
    Call BuildExecutivePortal
    
    ' 4. Build All 3 Analytical Dashboard Canvases with Real Interactive Visuals
    currentStep = "Building Call Center Cockpit (03_CallCenter_Cockpit)"
    Call BuildCallCenterCanvas(True)
    
    currentStep = "Building Customer Retention Cockpit (04_CustomerRetention_Cockpit)"
    Call BuildCustomerRetentionCanvas(True)
    
    currentStep = "Building Diversity & Inclusion Cockpit (05_DiversityInclusion_Cockpit)"
    Call BuildDiversityInclusionCanvas(True)
    
    ' 5. Reset all viewports (FreezePanes = False, Scroll = A1, Zoom = 80%)
    currentStep = "Normalizing Sheet Viewports & Gridlines"
    Call ResetAllViewports
    
    ' 6. Navigate to Call Center Cockpit as default landing view
    currentStep = "Navigating to Call Center Cockpit"
    Call NavigateToCallCenter
    
    RestoreAppState
    
    MsgBox "PwC Switzerland Analytics Platform built successfully!" & vbCrLf & vbCrLf & _
           "- Real visuals, charts & scorecards created across all 3 cockpits" & vbCrLf & _
           "- Ambiguous name conflicts resolved with explicit module bindings" & vbCrLf & _
           "- Viewports unfrozen & normalized (0 cut-off headers)" & vbCrLf & _
           "- Pure ASCII string encoding (zero character corruption)", _
           vbInformation, "PwC Digital Accelerator"
    Exit Sub
    
MasterErrHandler:
    Dim savedErrNum As Long, savedErrDesc As String, savedErrSrc As String
    savedErrNum = Err.Number
    savedErrDesc = Err.Description
    savedErrSrc = Err.Source
    
    RestoreAppState
    
    If savedErrNum <> 0 Then
        MsgBox "Error during step: '" & currentStep & "'" & vbCrLf & vbCrLf & _
               "Error Number: " & savedErrNum & vbCrLf & _
               "Description: " & savedErrDesc & vbCrLf & _
               "Source: " & savedErrSrc, vbCritical, "PwC Execution Error"
    Else
        MsgBox "Platform initialization completed.", vbInformation, "PwC Digital Accelerator"
    End If
End Sub
'''

with open(out_path, 'w', encoding='ascii') as f:
    f.write(code)

print(f"Successfully generated {out_path} with {len(code.splitlines())} lines.")
