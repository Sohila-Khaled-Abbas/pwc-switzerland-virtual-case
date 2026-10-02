Attribute VB_Name = "modDashboardUIUX"
Option Explicit

' ==============================================================================
' PwC Switzerland Digital Accelerator -- Executive Dashboard UI/UX Design System
' Web-Application Style Dashboard Architecture & Presentation Canvas Engine
' ==============================================================================
'
' Features:
'   1. Web-App SaaS Navigation Bar with embedded official PwC logo & live status
'   2. Hero Header with integrated SVG action buttons:
'      - [Refresh Data] -> modDataRefresh.RefreshPipelineSynchronously
'      - [Reset Filters] -> modFilterController.ClearAllFilters
'      - [Export PDF]    -> modExportPDF.ExportExecutiveReport
'   3. BAN KPI Metric Cards with customizable input & pure ASCII formatting (No encoding bugs)
'   4. Left Global Filter Drawer with dedicated Slicer Slots
'   5. 2x2 Grid of 4 Visual Container Cards with dashed docking zones
'   6. Transparent Chart Decluttering Engine (DeclutterAndFormatChart)
'   7. Unified Platform Integration (modAppState, modDataRefresh, modFilterController)
'
' ==============================================================================

' ------------------------------------------------------------------------------
' PwC Enterprise Brand Color Constants (RGB Values)
' ------------------------------------------------------------------------------
Public Const PWC_ORANGE         As Long = 133288     ' #D04A02 - RGB(208, 74, 2)
Public Const PWC_DARK_SLATE     As Long = 2758415    ' #0F172A - RGB(15, 23, 42)
Public Const PWC_CHARCOAL       As Long = 3877150    ' #1E293B - RGB(30, 41, 59)
Public Const PWC_CANVAS_BG      As Long = 16579320   ' #F8FAFC - RGB(248, 250, 252)
Public Const PWC_CARD_FILL      As Long = 16777215   ' #FFFFFF - RGB(255, 255, 255)
Public Const PWC_CARD_BORDER    As Long = 15790322   ' #E2E8F0 - RGB(226, 232, 240)
Public Const PWC_BORDER_DASHED  As Long = 13421772   ' #CBD5E1 - RGB(203, 213, 225)
Public Const PWC_SLOT_FILL      As Long = 16448250   ' #FAFAFA - RGB(250, 250, 250)
Public Const PWC_TEXT_TITLE     As Long = 3877150    ' #1E293B - RGB(30, 41, 59)
Public Const PWC_TEXT_MUTED     As Long = 9141108    ' #64748B - RGB(100, 116, 139)
Public Const PWC_TEXT_LIGHT     As Long = 12040119   ' #94A3B8 - RGB(148, 163, 184)
Public Const PWC_SUCCESS_GREEN  As Long = 4363781    ' #059669 - RGB(5, 150, 105)
Public Const PWC_ALERT_RED      As Long = 2500316    ' #DC2626 - RGB(220, 38, 38)
Public Const PWC_WARNING_AMBER  As Long = 422009     ' #D97706 - RGB(217, 119, 6)
Public Const PWC_PILL_BG        As Long = 16185073   ' #F1F5F9 - RGB(241, 245, 249)

' Default Enterprise Typography
Public Const FONT_FAMILY        As String = "Segoe UI"

' Layout Grid Dimensions (Modular 1214pt Canvas)
Public Const CANVAS_LEFT        As Single = 24
Public Const CANVAS_WIDTH       As Single = 1214
Public Const CARD_GAP           As Single = 16
Public Const KPI_WIDTH          As Single = 230
Public Const KPI_HEIGHT         As Single = 84
Public Const SLICER_WIDTH       As Single = 230
Public Const SLICER_HEIGHT      As Single = 554
Public Const VISUAL_WIDTH       As Single = 476
Public Const VISUAL_HEIGHT      As Single = 270

' ==============================================================================
' 1. HELPER: SAFE SVG / PNG ASSET RESOLVER
' ==============================================================================
Public Function ResolveAssetPath(ByVal subFolderAndFile As String) As String
    Dim wb As Workbook
    Dim testPath As String
    Set wb = ActiveWorkbook
    
    ' 1. Check relative to workbook
    testPath = wb.Path & "\" & subFolderAndFile
    If Dir(testPath) <> "" Then
        ResolveAssetPath = testPath
        Exit Function
    End If
    
    ' 2. Check parent directory relative
    testPath = wb.Path & "\..\" & subFolderAndFile
    If Dir(testPath) <> "" Then
        ResolveAssetPath = testPath
        Exit Function
    End If
    
    ' 3. Fallback to workspace absolute path
    testPath = "d:\courses\Data Analysis 26-27\7-Introducation to Data Fields (Excel)\11_Demos_and_Workbooks\10_Projects_and_Demos\PWC\" & subFolderAndFile
    If Dir(testPath) <> "" Then
        ResolveAssetPath = testPath
        Exit Function
    End If
    
    ResolveAssetPath = ""
End Function

Public Function InsertVectorIcon(ws As Worksheet, ByVal iconFileName As String, _
                                ByVal leftPos As Single, ByVal topPos As Single, _
                                ByVal iconWidth As Single, ByVal iconHeight As Single, _
                                ByVal shapeName As String) As Shape
    Dim fullPath As String
    Dim shp As Shape
    On Error Resume Next
    fullPath = ResolveAssetPath("assets\icons\" & iconFileName)
    If Len(fullPath) > 0 Then
        Set shp = ws.Shapes.AddPicture(fullPath, msoFalse, msoTrue, leftPos, topPos, iconWidth, iconHeight)
        If Not shp Is Nothing Then
            shp.Name = shapeName
            Set InsertVectorIcon = shp
        End If
    End If
    On Error GoTo 0
End Function

' ==============================================================================
' 2. MASTER WORKSPACE CANVAS INITIALIZER
' ==============================================================================
Public Sub InitializeDashboardCanvas(ws As Worksheet, Optional ByVal bgColor As Long = PWC_CANVAS_BG)
    On Error Resume Next
    ws.Activate
    
    ' Hide Excel spreadsheet gridlines and row/column headers for presentation-ready UI
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.Zoom = 100
    
    ' Standardize base column widths and row heights to prevent coordinate drift
    ws.Columns("A:AZ").ColumnWidth = 11
    ws.Rows("1:60").RowHeight = 20
    
    ' Flood worksheet with uniform canvas background
    With ws.Cells.Interior
        .Pattern = xlSolid
        .Color = bgColor
    End With
End Sub

' ==============================================================================
' 3. WEB-APPLICATION TOP NAVIGATION BAR (WITH EMBEDDED PWC LOGO & SVG STATUS)
' ==============================================================================
Public Sub BuildWebTopNavBar(ws As Worksheet, ByVal activeModuleCode As String)
    Dim wb As Workbook
    Dim shpNav As Shape
    Dim shpLogo As Shape
    Dim shpDivider As Shape
    Dim shpBrand As Shape
    Dim shpLive As Shape
    Dim logoPath As String
    Dim navTop As Single, navLeft As Single, navW As Single, navH As Single
    
    Set wb = ws.Parent
    navTop = 16
    navLeft = CANVAS_LEFT
    navW = CANVAS_WIDTH
    navH = 52
    
    ' 1. Master Navigation Bar Container
    Set shpNav = ws.Shapes.AddShape(msoShapeRoundedRectangle, navLeft, navTop, navW, navH)
    With shpNav
        .Name = "Nav_MasterBar"
        .Fill.Solid
        .Fill.ForeColor.RGB = PWC_CARD_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER
        .Line.Weight = 1
        .Adjustments.Item(1) = 0.12
        With .Shadow
            .Type = msoShadow21: .Visible = msoTrue: .Blur = 8: .Transparency = 0.88: .OffsetX = 0: .OffsetY = 3
        End With
    End With
    
    ' 2. Official PwC Brand Logo Insertion
    logoPath = ResolveAssetPath("assets\PwC_logo_rgb_colour_pos.png")
    If Len(logoPath) > 0 Then
        ' Embed picture permanently into workbook (SaveWithDocument = msoTrue)
        Set shpLogo = ws.Shapes.AddPicture(logoPath, msoFalse, msoTrue, navLeft + 16, navTop + 9, 53, 34)
        If Not shpLogo Is Nothing Then shpLogo.Name = "Nav_PwCLogo"
    End If
    
    ' 3. Subtle Vertical Divider
    Set shpDivider = ws.Shapes.AddShape(msoShapeRectangle, navLeft + 80, navTop + 12, 1, 28)
    With shpDivider
        .Name = "Nav_Divider"
        .Fill.Solid
        .Fill.ForeColor.RGB = PWC_CARD_BORDER
        .Line.Visible = msoFalse
    End With
    
    ' 4. Hub Brand & Subtitle
    Set shpBrand = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, navLeft + 90, navTop + 9, 210, 34)
    With shpBrand
        .Name = "Nav_BrandText"
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2
            .MarginLeft = 0: .MarginTop = 0: .MarginRight = 0: .MarginBottom = 0
            .WordWrap = msoFalse
            With .TextRange
                .Text = "PwC Digital Intelligence" & vbCrLf & "Executive Decision Hub"
                With .Paragraphs(1).Font
                    .Name = FONT_FAMILY: .Size = 11: .Bold = msoTrue
                    .Fill.ForeColor.RGB = PWC_TEXT_TITLE
                End With
                With .Paragraphs(2).Font
                    .Name = FONT_FAMILY: .Size = 8: .Bold = msoFalse
                    .Fill.ForeColor.RGB = PWC_TEXT_MUTED
                End With
            End With
        End With
    End With
    
    ' 5. Web App Navigation Tabs (Pills)
    Dim tabLeft As Single, tabTop As Single, tabW As Single, tabH As Single
    tabTop = navTop + 12
    tabH = 28
    
    ' Tab 1: Overview (Business Domains)
    tabLeft = navLeft + 310: tabW = 95
    Call CreateNavTabPill(ws, "Nav_Tab_Domains", tabLeft, tabTop, tabW, tabH, _
                          "01 Domains", "'01_Business_Domains'!A1", (activeModuleCode = "DOM"))
                          
    ' Tab 2: Data Catalog
    tabLeft = tabLeft + tabW + 8: tabW = 100
    Call CreateNavTabPill(ws, "Nav_Tab_Catalog", tabLeft, tabTop, tabW, tabH, _
                          "02 Catalog", "'02_Metadata_&_KPI_Catalog'!A1", (activeModuleCode = "CAT"))
                          
    ' Tab 3: Call Center Cockpit
    tabLeft = tabLeft + tabW + 8: tabW = 115
    Call CreateNavTabPill(ws, "Nav_Tab_CC", tabLeft, tabTop, tabW, tabH, _
                          "03 Call Center", "'03_CallCenter_Cockpit'!A1", (activeModuleCode = "CC"))
                          
    ' Tab 4: Customer Retention Cockpit
    tabLeft = tabLeft + tabW + 8: tabW = 125
    Call CreateNavTabPill(ws, "Nav_Tab_CH", tabLeft, tabTop, tabW, tabH, _
                          "04 Retention Risk", "'04_CustomerRetention_Cockpit'!A1", (activeModuleCode = "CH"))
                          
    ' Tab 5: Diversity & Inclusion Cockpit
    tabLeft = tabLeft + tabW + 8: tabW = 120
    Call CreateNavTabPill(ws, "Nav_Tab_DI", tabLeft, tabTop, tabW, tabH, _
                          "05 D&I Parity", "'05_DiversityInclusion_Cockpit'!A1", (activeModuleCode = "DI"))
                          
    ' 6. Live Model Status Indicator with pulsing live SVG icon
    Dim livePillW As Single, livePillLeft As Single
    livePillW = 140
    livePillLeft = navLeft + navW - livePillW - 12
    
    Set shpLive = ws.Shapes.AddShape(msoShapeRoundedRectangle, livePillLeft, navTop + 12, livePillW, 28)
    With shpLive
        .Name = "Nav_StatusPill"
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(236, 253, 245)  ' Soft emerald background
        .Line.ForeColor.RGB = RGB(167, 243, 208)  ' Emerald border
        .Line.Weight = 1
        .Adjustments.Item(1) = 0.5
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.MarginLeft = 16: .TextFrame.MarginRight = 2: .TextFrame.MarginTop = 0: .TextFrame.MarginBottom = 0
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 16: .MarginTop = 0: .MarginRight = 2: .MarginBottom = 0
            .WordWrap = msoFalse
            With .TextRange
                .Text = "LIVE VERTIPAQ"
                .Font.Name = FONT_FAMILY: .Font.Size = 8: .Font.Bold = msoTrue
                .Font.Fill.ForeColor.RGB = PWC_SUCCESS_GREEN
                .ParagraphFormat.Alignment = msoAlignCenter
            End With
        End With
    End With
    
    ' Embed SVG live indicator icon inside the pill
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
        .Adjustments.Item(1) = 0.25
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.MarginLeft = 0: .TextFrame.MarginRight = 0: .TextFrame.MarginTop = 0: .TextFrame.MarginBottom = 0
        If isActive Then
            .Fill.Solid
            .Fill.ForeColor.RGB = PWC_ORANGE
            .Line.Visible = msoFalse
            With .TextFrame2
                .VerticalAnchor = msoAnchorMiddle
                .MarginLeft = 0: .MarginTop = 0: .MarginRight = 0: .MarginBottom = 0
                .WordWrap = msoFalse
                With .TextRange
                    .Text = tabText
                    .Font.Name = FONT_FAMILY: .Font.Size = 8.5: .Font.Bold = msoTrue
                    .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                    .ParagraphFormat.Alignment = msoAlignCenter
                End With
            End With
        Else
            .Fill.Solid
            .Fill.ForeColor.RGB = PWC_PILL_BG
            .Line.ForeColor.RGB = PWC_CARD_BORDER
            .Line.Weight = 1
            With .TextFrame2
                .VerticalAnchor = msoAnchorMiddle
                .MarginLeft = 0: .MarginTop = 0: .MarginRight = 0: .MarginBottom = 0
                .WordWrap = msoFalse
                With .TextRange
                    .Text = tabText
                    .Font.Name = FONT_FAMILY: .Font.Size = 8.5: .Font.Bold = msoTrue
                    .Font.Fill.ForeColor.RGB = PWC_TEXT_MUTED
                    .ParagraphFormat.Alignment = msoAlignCenter
                End With
            End With
            On Error Resume Next
            ws.Hyperlinks.Add Anchor:=shp, Address:="", SubAddress:=subAddress
            On Error GoTo 0
        End If
    End With
End Sub

' ==============================================================================
' 4. EXECUTIVE HERO HEADER WITH INTEGRATED SVG ACTION BUTTONS
' ==============================================================================
Public Sub BuildHeroHeader(ws As Worksheet, _
                           ByVal dashboardTitle As String, _
                           ByVal subtitle As String, _
                           Optional ByVal topPos As Single = 76)
    Dim shpTitle As Shape
    Dim shpRefreshBtn As Shape, shpClearBtn As Shape, shpExportBtn As Shape
    Dim heroW As Single, heroLeft As Single
    heroLeft = CANVAS_LEFT
    heroW = CANVAS_WIDTH
    
    ' 1. Title & Context Description
    Set shpTitle = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, heroLeft, topPos, heroW - 380, 46)
    With shpTitle
        .Name = "Hero_TitleText"
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2
            .MarginLeft = 0: .MarginTop = 0: .MarginRight = 0: .MarginBottom = 0
            .WordWrap = msoFalse
            With .TextRange
                .Text = dashboardTitle & vbCrLf & subtitle
                With .Paragraphs(1).Font
                    .Name = FONT_FAMILY: .Size = 16: .Bold = msoTrue
                    .Fill.ForeColor.RGB = PWC_TEXT_TITLE
                End With
                With .Paragraphs(2).Font
                    .Name = FONT_FAMILY: .Size = 9: .Bold = msoFalse
                    .Fill.ForeColor.RGB = PWC_TEXT_MUTED
                End With
            End With
        End With
    End With
    
    ' Button Coordinates
    Dim btnTop As Single, btnH As Single
    btnTop = topPos + 8
    btnH = 28
    
    ' Total width: 122 + 8 + 118 + 8 + 110 = 366 pt
    Dim refBtnLeft As Single, refBtnW As Single
    refBtnW = 122
    refBtnLeft = heroLeft + heroW - 366
    
    ' 2. Web App Action Button: Refresh Data (Connected to modDataRefresh)
    Set shpRefreshBtn = ws.Shapes.AddShape(msoShapeRoundedRectangle, refBtnLeft, btnTop, refBtnW, btnH)
    With shpRefreshBtn
        .Name = "Hero_BtnRefresh"
        .Fill.Solid
        .Fill.ForeColor.RGB = PWC_CARD_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER
        .Line.Weight = 1
        .Adjustments.Item(1) = 0.25
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.MarginLeft = 16: .TextFrame.MarginRight = 2: .TextFrame.MarginTop = 0: .TextFrame.MarginBottom = 0
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 16: .MarginTop = 0: .MarginRight = 2: .MarginBottom = 0
            .WordWrap = msoFalse
            With .TextRange
                .Text = "Refresh Data"
                .Font.Name = FONT_FAMILY: .Font.Size = 8.5: .Font.Bold = msoTrue
                .Font.Fill.ForeColor.RGB = PWC_ORANGE
                .ParagraphFormat.Alignment = msoAlignCenter
            End With
        End With
        .OnAction = "modDataRefresh.RefreshPipelineSynchronously"
    End With
    Call InsertVectorIcon(ws, "icon_refresh_pipeline.svg", refBtnLeft + 10, btnTop + 7, 14, 14, "Hero_IconRefresh")
    
    ' 3. Web App Action Button: Reset Filters (Connected to modFilterController)
    Dim clearBtnLeft As Single, clearBtnW As Single
    clearBtnW = 118
    clearBtnLeft = refBtnLeft + refBtnW + 8
    
    Set shpClearBtn = ws.Shapes.AddShape(msoShapeRoundedRectangle, clearBtnLeft, btnTop, clearBtnW, btnH)
    With shpClearBtn
        .Name = "Hero_BtnResetSlicers"
        .Fill.Solid
        .Fill.ForeColor.RGB = PWC_CARD_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER
        .Line.Weight = 1
        .Adjustments.Item(1) = 0.25
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.MarginLeft = 16: .TextFrame.MarginRight = 2: .TextFrame.MarginTop = 0: .TextFrame.MarginBottom = 0
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 16: .MarginTop = 0: .MarginRight = 2: .MarginBottom = 0
            .WordWrap = msoFalse
            With .TextRange
                .Text = "Reset Filters"
                .Font.Name = FONT_FAMILY: .Font.Size = 8.5: .Font.Bold = msoTrue
                .Font.Fill.ForeColor.RGB = PWC_TEXT_TITLE
                .ParagraphFormat.Alignment = msoAlignCenter
            End With
        End With
        .OnAction = "modFilterController.ClearAllFilters"
    End With
    Call InsertVectorIcon(ws, "icon_reset_filter.svg", clearBtnLeft + 10, btnTop + 7, 14, 14, "Hero_IconReset")
    
    ' 4. Web App Action Button: Export PDF (Connected to modExportPDF)
    Dim exportBtnLeft As Single, exportBtnW As Single
    exportBtnW = 110
    exportBtnLeft = clearBtnLeft + clearBtnW + 8
    
    Set shpExportBtn = ws.Shapes.AddShape(msoShapeRoundedRectangle, exportBtnLeft, btnTop, exportBtnW, btnH)
    With shpExportBtn
        .Name = "Hero_BtnExportPDF"
        .Fill.Solid
        .Fill.ForeColor.RGB = PWC_CHARCOAL
        .Line.Visible = msoFalse
        .Adjustments.Item(1) = 0.25
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.MarginLeft = 16: .TextFrame.MarginRight = 2: .TextFrame.MarginTop = 0: .TextFrame.MarginBottom = 0
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 16: .MarginTop = 0: .MarginRight = 2: .MarginBottom = 0
            .WordWrap = msoFalse
            With .TextRange
                .Text = "Export PDF"
                .Font.Name = FONT_FAMILY: .Font.Size = 8.5: .Font.Bold = msoTrue
                .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .ParagraphFormat.Alignment = msoAlignCenter
            End With
        End With
        .OnAction = "modExportPDF.ExportExecutiveReport"
    End With
    Call InsertVectorIcon(ws, "icon_export_pdf.svg", exportBtnLeft + 10, btnTop + 7, 14, 14, "Hero_IconPDF")
End Sub

' ==============================================================================
' 5. FLOATING BAN KPI METRIC CARD (CUSTOMIZABLE INPUT - ZERO ENCODING BUGS)
' ==============================================================================
Public Sub BuildKPICard(ws As Worksheet, _
                        ByVal cardName As String, _
                        ByVal leftPos As Single, _
                        ByVal topPos As Single, _
                        ByVal cardWidth As Single, _
                        ByVal cardHeight As Single, _
                        ByVal kpiLabel As String, _
                        Optional ByVal kpiValue As String = "--", _
                        Optional ByVal targetSubtext As String = "", _
                        Optional ByVal accentColor As Long = PWC_ORANGE)
    Dim shpCard As Shape
    Dim shpAccentLine As Shape
    Dim shpText As Shape
    
    ' Ensure pure ASCII formatting for value string
    If Len(Trim(kpiValue)) = 0 Then kpiValue = "--"
    
    ' 1. Card Container with Rounded Corners & Diffused Shadow
    Set shpCard = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, cardWidth, cardHeight)
    With shpCard
        .Name = "Card_" & cardName
        .Fill.Solid
        .Fill.ForeColor.RGB = PWC_CARD_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER
        .Line.Weight = 1
        .Adjustments.Item(1) = 0.1
        With .Shadow
            .Type = msoShadow21: .Visible = msoTrue: .Blur = 6: .Transparency = 0.86: .OffsetX = 0: .OffsetY = 2
        End With
    End With
    
    ' 2. Top Color Indicator Line
    Set shpAccentLine = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos + 12, topPos + 5, cardWidth - 24, 3)
    With shpAccentLine
        .Name = "Accent_" & cardName
        .Fill.Solid
        .Fill.ForeColor.RGB = accentColor
        .Line.Visible = msoFalse
        .Adjustments.Item(1) = 0.5
    End With
    
    ' 3A. Metric Label (Fixed Header)
    Dim shpLabel As Shape, shpValue As Shape, shpSubtext As Shape
    Set shpLabel = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos + 14, topPos + 10, cardWidth - 28, 16)
    With shpLabel
        .Name = "Label_" & cardName
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        With .TextFrame2
            .MarginLeft = 0: .MarginTop = 0: .MarginRight = 0: .MarginBottom = 0
            .WordWrap = msoFalse
            With .TextRange
                .Text = UCase(kpiLabel)
                .Font.Name = FONT_FAMILY: .Font.Size = 8: .Font.Bold = msoTrue
                .Font.Fill.ForeColor.RGB = PWC_TEXT_MUTED
            End With
        End With
    End With
    
    ' 3B. Metric Value Callout (Independent Shape -- Can be Formula-Linked via Formula Bar!)
    Set shpValue = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos + 14, topPos + 24, cardWidth - 28, 36)
    With shpValue
        .Name = "Value_" & cardName
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.MarginLeft = 0: .TextFrame.MarginRight = 0: .TextFrame.MarginTop = 0: .TextFrame.MarginBottom = 0
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 0: .MarginTop = 0: .MarginRight = 0: .MarginBottom = 0
            .WordWrap = msoFalse
            With .TextRange
                .Text = kpiValue
                .Font.Name = FONT_FAMILY: .Font.Size = 22: .Font.Bold = msoTrue
                .Font.Fill.ForeColor.RGB = PWC_TEXT_TITLE
            End With
        End With
    End With
    
    ' 3C. Benchmark / SLA Target Subtext (Fixed Footer)
    Set shpSubtext = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos + 14, topPos + 60, cardWidth - 28, 16)
    With shpSubtext
        .Name = "Subtext_" & cardName
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        With .TextFrame2
            .MarginLeft = 0: .MarginTop = 0: .MarginRight = 0: .MarginBottom = 0
            .WordWrap = msoFalse
            With .TextRange
                .Text = targetSubtext
                .Font.Name = FONT_FAMILY: .Font.Size = 8: .Font.Bold = msoFalse
                .Font.Fill.ForeColor.RGB = accentColor
            End With
        End With
    End With
End Sub

' Helper Subroutine to Programmatically Set or Update KPI Card Values
Public Sub SetKPICardValue(ws As Worksheet, ByVal cardName As String, _
                           ByVal newValue As String, Optional ByVal newSubtext As String = "")
    Dim shpValue As Shape
    Dim shpSub As Shape
    Dim shpText As Shape
    On Error Resume Next
    Set shpValue = ws.Shapes("Value_" & cardName)
    If Not shpValue Is Nothing Then
        shpValue.TextFrame2.TextRange.Text = newValue
        If Len(Trim(newSubtext)) > 0 Then
            Set shpSub = ws.Shapes("Subtext_" & cardName)
            If Not shpSub Is Nothing Then shpSub.TextFrame2.TextRange.Text = newSubtext
        End If
        Exit Sub
    End If
    ' Fallback to legacy single textbox if present
    Set shpText = ws.Shapes("Text_" & cardName)
    If Not shpText Is Nothing Then
        With shpText.TextFrame2.TextRange
            .Paragraphs(2).Text = newValue & vbCrLf
            If Len(Trim(newSubtext)) > 0 And .Paragraphs.Count >= 3 Then
                .Paragraphs(3).Text = newSubtext
            End If
        End With
    End If
    On Error GoTo 0
End Sub

' ==============================================================================
' 5B. AUTOMATED DATA MODEL STAGING & FORMULA LINKING ENGINE
' Automatically provisions CUBEVALUE formulas in rows 64-65 and links the
' independent Value_<cardName> shapes dynamically without manual clicking
' ==============================================================================
Public Sub AutomateAndLinkKPICards(ws As Worksheet, ByVal moduleCode As String)
    On Error Resume Next
    Dim cols(1 To 5) As String
    cols(1) = "AA": cols(2) = "AB": cols(3) = "AC": cols(4) = "AD": cols(5) = "AE"
    
    Dim headers(1 To 5) As String
    Dim measures(1 To 5) As String
    Dim numFormats(1 To 5) As String
    Dim cardNames(1 To 5) As String
    
    Select Case UCase(moduleCode)
        Case "CC"
            headers(1) = "KPI 1: Total Calls": measures(1) = "[Measures].[Total Calls]": numFormats(1) = "#,##0": cardNames(1) = "CC_TotalDemand"
            headers(2) = "KPI 2: Answer Rate": measures(2) = "[Measures].[Answer Rate %]": numFormats(2) = "0.00%": cardNames(2) = "CC_Answered"
            headers(3) = "KPI 3: Abandonment Rate": measures(3) = "[Measures].[Abandonment Rate %]": numFormats(3) = "0.00%": cardNames(3) = "CC_Abandoned"
            headers(4) = "KPI 4: Avg Speed of Answer": measures(4) = "[Measures].[Avg Speed of Answer]": numFormats(4) = "#,##0.0 ""s""": cardNames(4) = "CC_ASA"
            headers(5) = "KPI 5: Avg CSAT Rating": measures(5) = "[Measures].[Avg CSAT Rating]": numFormats(5) = "0.00": cardNames(5) = "CC_CSAT"
            
        Case "CH"
            headers(1) = "KPI 1: Total Customers": measures(1) = "[Measures].[Total Customers]": numFormats(1) = "#,##0": cardNames(1) = "CH_Subscribers"
            headers(2) = "KPI 2: Churn Rate": measures(2) = "[Measures].[Churn Rate %]": numFormats(2) = "0.00%": cardNames(2) = "CH_ChurnRate"
            headers(3) = "KPI 3: Revenue at Risk": measures(3) = "[Measures].[Total Revenue at Risk]": numFormats(3) = "$#,##0.00": cardNames(3) = "CH_ARRRisk"
            headers(4) = "KPI 4: M2M Churn Rate": measures(4) = "[Measures].[Contract M2M Churn Rate %]": numFormats(4) = "0.00%": cardNames(4) = "CH_M2MChurn"
            headers(5) = "KPI 5: Tech Tickets / Cust": measures(5) = "[Measures].[Tech Tickets per Customer]": numFormats(5) = "0.00": cardNames(5) = "CH_Tickets"
            
        Case "DI"
            headers(1) = "KPI 1: Total Headcount": measures(1) = "[Measures].[Total Headcount]": numFormats(1) = "#,##0": cardNames(1) = "DI_Workforce"
            headers(2) = "KPI 2: Female Headcount Share": measures(2) = "[Measures].[Female Headcount Share %]": numFormats(2) = "0.00%": cardNames(2) = "DI_FemaleShare"
            headers(3) = "KPI 3: Broken Rung Gap": measures(3) = "[Measures].[Broken Rung Gap]": numFormats(3) = "+0.00%;-0.00%;0.00%": cardNames(3) = "DI_BrokenRung"
            headers(4) = "KPI 4: Female Promotion Share": measures(4) = "[Measures].[Female Promotion Share %]": numFormats(4) = "0.00%": cardNames(4) = "DI_PromoShare"
            headers(5) = "KPI 5: Time in Grade Gap": measures(5) = "[Measures].[Time in Grade Gap]": numFormats(5) = "+0.0 ""Mos"";-0.0 ""Mos"";0.0 ""Mos""": cardNames(5) = "DI_TimeInGrade"
    End Select
    
    Dim i As Integer
    Dim cellRef As String
    Dim shpValue As Shape
    
    For i = 1 To 5
        ' 1. Set Staging Header
        With ws.Range(cols(i) & "64")
            .Value = headers(i)
            .Font.Name = FONT_FAMILY
            .Font.Size = 8
            .Font.Bold = True
            .Font.Color = PWC_TEXT_MUTED
        End With
        
        ' 2. Set CUBEVALUE Formula and Format
        With ws.Range(cols(i) & "65")
            .Formula = "=CUBEVALUE(""ThisWorkbookDataModel"", """ & measures(i) & """)"
            .NumberFormat = numFormats(i)
            .Font.Name = FONT_FAMILY
            .Font.Size = 9
            .Font.Bold = False
        End With
        
        ' 3. Automatically Link Value Shape to Staging Cell
        cellRef = "='" & ws.Name & "'!$" & cols(i) & "$65"
        Set shpValue = Nothing
        Set shpValue = ws.Shapes("Value_" & cardNames(i))
        If Not shpValue Is Nothing Then
            shpValue.DrawingObject.Formula = cellRef
        End If
    Next i
    On Error GoTo 0
End Sub

' ==============================================================================
' 6. LEFT GLOBAL FILTER DRAWER (SLICER CONTAINER)
' ==============================================================================
Public Sub BuildSlicerPanelContainer(ws As Worksheet, _
                                     ByVal panelName As String, _
                                     ByVal leftPos As Single, _
                                     ByVal topPos As Single, _
                                     ByVal panelWidth As Single, _
                                     ByVal panelHeight As Single, _
                                     ByVal slot1Label As String, _
                                     ByVal slot2Label As String, _
                                     ByVal slot3Label As String)
    Dim shpPanel As Shape
    Dim shpHeader As Shape
    Dim shpSlot1 As Shape, shpSlot2 As Shape, shpSlot3 As Shape
    Dim shpFooter As Shape
    
    ' 1. Outer Container Card
    Set shpPanel = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, panelWidth, panelHeight)
    With shpPanel
        .Name = "Panel_" & panelName
        .Fill.Solid
        .Fill.ForeColor.RGB = PWC_CARD_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER
        .Line.Weight = 1
        .Adjustments.Item(1) = 0.05
        With .Shadow
            .Type = msoShadow21: .Visible = msoTrue: .Blur = 8: .Transparency = 0.88: .OffsetX = 0: .OffsetY = 3
        End With
    End With
    
    ' 2. Panel Header Text
    Set shpHeader = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos + 14, topPos + 12, panelWidth - 28, 36)
    With shpHeader
        .Name = "Header_" & panelName
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2
            .MarginLeft = 0: .MarginTop = 0: .MarginRight = 0: .MarginBottom = 0
            .WordWrap = msoFalse
            With .TextRange
                .Text = "GLOBAL FILTERS" & vbCrLf & "Interactive Slicer Drawer"
                With .Paragraphs(1).Font
                    .Name = FONT_FAMILY: .Size = 10.5: .Bold = msoTrue
                    .Fill.ForeColor.RGB = PWC_TEXT_TITLE
                End With
                With .Paragraphs(2).Font
                    .Name = FONT_FAMILY: .Size = 8: .Fill.ForeColor.RGB = PWC_TEXT_MUTED
                End With
            End With
        End With
    End With
    
    ' 3. Filter Slot 1 (Primary Dimension)
    Set shpSlot1 = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos + 12, topPos + 54, panelWidth - 24, 145)
    With shpSlot1
        .Name = "Slot1_" & panelName
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_SLOT_FILL
        .Line.ForeColor.RGB = PWC_BORDER_DASHED: .Line.Weight = 0.75: .Line.DashStyle = msoLineDash
        .Adjustments.Item(1) = 0.06
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 8: .MarginTop = 8: .MarginRight = 8: .MarginBottom = 8
            With .TextRange
                .Text = "[ Slicer Slot 1: " & slot1Label & " ]" & vbCrLf & vbCrLf & "Insert Slicer from Data Model & position here."
                .Font.Name = FONT_FAMILY: .Font.Size = 8: .Font.Fill.ForeColor.RGB = PWC_TEXT_LIGHT
                .ParagraphFormat.Alignment = msoAlignCenter
            End With
        End With
    End With
    
    ' 4. Filter Slot 2 (Secondary Dimension)
    Set shpSlot2 = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos + 12, topPos + 210, panelWidth - 24, 145)
    With shpSlot2
        .Name = "Slot2_" & panelName
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_SLOT_FILL
        .Line.ForeColor.RGB = PWC_BORDER_DASHED: .Line.Weight = 0.75: .Line.DashStyle = msoLineDash
        .Adjustments.Item(1) = 0.06
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 8: .MarginTop = 8: .MarginRight = 8: .MarginBottom = 8
            With .TextRange
                .Text = "[ Slicer Slot 2: " & slot2Label & " ]" & vbCrLf & vbCrLf & "Insert Slicer from Data Model & position here."
                .Font.Name = FONT_FAMILY: .Font.Size = 8: .Font.Fill.ForeColor.RGB = PWC_TEXT_LIGHT
                .ParagraphFormat.Alignment = msoAlignCenter
            End With
        End With
    End With
    
    ' 5. Filter Slot 3 (Tertiary Dimension)
    Set shpSlot3 = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos + 12, topPos + 366, panelWidth - 24, 135)
    With shpSlot3
        .Name = "Slot3_" & panelName
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_SLOT_FILL
        .Line.ForeColor.RGB = PWC_BORDER_DASHED: .Line.Weight = 0.75: .Line.DashStyle = msoLineDash
        .Adjustments.Item(1) = 0.06
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 8: .MarginTop = 8: .MarginRight = 8: .MarginBottom = 8
            With .TextRange
                .Text = "[ Slicer Slot 3: " & slot3Label & " ]" & vbCrLf & vbCrLf & "Insert Slicer from Data Model & position here."
                .Font.Name = FONT_FAMILY: .Font.Size = 8: .Font.Fill.ForeColor.RGB = PWC_TEXT_LIGHT
                .ParagraphFormat.Alignment = msoAlignCenter
            End With
        End With
    End With
    
    ' 6. Helper Guidance Text at bottom
    Set shpFooter = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos + 12, topPos + 510, panelWidth - 24, 34)
    With shpFooter
        .Name = "Footer_" & panelName
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        With .TextFrame2
            .MarginLeft = 0: .MarginTop = 0: .MarginRight = 0: .MarginBottom = 0
            .WordWrap = msoTrue
            With .TextRange
                .Text = "Multi-select slicers dynamically filter all visual containers via Data Model connections."
                .Font.Name = FONT_FAMILY: .Font.Size = 7.5: .Font.Fill.ForeColor.RGB = PWC_TEXT_LIGHT
            End With
        End With
    End With
End Sub

' ==============================================================================
' 7. VISUAL CONTAINER CARD (CHART DOCKING FRAME)
' ==============================================================================
Public Sub BuildChartContainer(ws As Worksheet, _
                               ByVal containerName As String, _
                               ByVal leftPos As Single, _
                               ByVal topPos As Single, _
                               ByVal cardWidth As Single, _
                               ByVal cardHeight As Single, _
                               ByVal chartTitle As String, _
                               ByVal chartSubtitle As String, _
                               ByVal chartTypeBadge As String)
    Dim shpContainer As Shape
    Dim shpHeader As Shape
    Dim shpBadge As Shape
    Dim shpDockingZone As Shape
    
    ' 1. Base Floating Card Container
    Set shpContainer = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, cardWidth, cardHeight)
    With shpContainer
        .Name = "Container_" & containerName
        .Fill.Solid
        .Fill.ForeColor.RGB = PWC_CARD_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER
        .Line.Weight = 1
        .Adjustments.Item(1) = 0.04
        With .Shadow
            .Type = msoShadow21: .Visible = msoTrue: .Blur = 8: .Transparency = 0.88: .OffsetX = 0: .OffsetY = 3
        End With
    End With
    
    ' 2. Container Header Title & Subtitle
    Set shpHeader = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos + 16, topPos + 10, cardWidth - 140, 36)
    With shpHeader
        .Name = "Header_" & containerName
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2
            .MarginLeft = 0: .MarginTop = 0: .MarginRight = 0: .MarginBottom = 0
            .WordWrap = msoFalse
            With .TextRange
                .Text = chartTitle & vbCrLf & chartSubtitle
                With .Paragraphs(1).Font
                    .Name = FONT_FAMILY: .Size = 10.5: .Bold = msoTrue
                    .Fill.ForeColor.RGB = PWC_TEXT_TITLE
                End With
                With .Paragraphs(2).Font
                    .Name = FONT_FAMILY: .Size = 8: .Bold = msoFalse
                    .Fill.ForeColor.RGB = PWC_TEXT_MUTED
                End With
            End With
        End With
    End With
    
    ' 3. Visual Type Badge (Top Right)
    Set shpBadge = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos + cardWidth - 116, topPos + 10, 102, 22)
    With shpBadge
        .Name = "Badge_" & containerName
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_PILL_BG
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 0.75
        .Adjustments.Item(1) = 0.25
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.MarginLeft = 0: .TextFrame.MarginRight = 0: .TextFrame.MarginTop = 0: .TextFrame.MarginBottom = 0
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 0: .MarginTop = 0: .MarginRight = 0: .MarginBottom = 0
            .WordWrap = msoFalse
            With .TextRange
                .Text = chartTypeBadge
                .Font.Name = FONT_FAMILY: .Font.Size = 7.5: .Font.Bold = msoTrue
                .Font.Fill.ForeColor.RGB = PWC_TEXT_MUTED
                .ParagraphFormat.Alignment = msoAlignCenter
            End With
        End With
    End With
    
    ' 4. Visual Docking Drop Zone (Watermarked Placeholder)
    Dim dockLeft As Single, dockTop As Single, dockW As Single, dockH As Single
    dockLeft = leftPos + 14
    dockTop = topPos + 48
    dockW = cardWidth - 28
    dockH = cardHeight - 60
    
    Set shpDockingZone = ws.Shapes.AddShape(msoShapeRoundedRectangle, dockLeft, dockTop, dockW, dockH)
    With shpDockingZone
        .Name = "DockZone_" & containerName
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_SLOT_FILL
        .Line.ForeColor.RGB = PWC_BORDER_DASHED: .Line.Weight = 0.75: .Line.DashStyle = msoLineDash
        .Adjustments.Item(1) = 0.04
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 12: .MarginTop = 12: .MarginRight = 12: .MarginBottom = 12
            With .TextRange
                .Text = "[ PIVOTCHART DOCKING ZONE ]" & vbCrLf & _
                        "Insert PivotChart or Table into this card boundary." & vbCrLf & _
                        "Run DeclutterAndFormatChart for transparent integration."
                With .Paragraphs(1).Font
                    .Name = FONT_FAMILY: .Size = 9: .Bold = msoTrue
                    .Fill.ForeColor.RGB = PWC_TEXT_LIGHT
                End With
                With .Paragraphs(2).Font
                    .Name = FONT_FAMILY: .Size = 8: .Fill.ForeColor.RGB = PWC_TEXT_LIGHT
                End With
                With .Paragraphs(3).Font
                    .Name = FONT_FAMILY: .Size = 7.5: .Italic = msoTrue
                    .Fill.ForeColor.RGB = PWC_TEXT_LIGHT
                End With
                .ParagraphFormat.Alignment = msoAlignCenter
            End With
        End With
    End With
End Sub

' ==============================================================================
' 8. CHART TRANSPARENCY & DECLUTTERING FORMATTER
' ==============================================================================
Public Sub DeclutterAndFormatChart(chtObj As ChartObject)
    On Error Resume Next
    With chtObj
        ' Remove external chart border and make background completely transparent
        .ShapeRange.Fill.Visible = msoFalse
        .ShapeRange.Line.Visible = msoFalse
        
        With .Chart
            .ChartArea.Format.Fill.Visible = msoFalse
            .ChartArea.Format.Line.Visible = msoFalse
            .PlotArea.Format.Fill.Visible = msoFalse
            .PlotArea.Format.Line.Visible = msoFalse
            
            ' Apply enterprise typography
            .ChartArea.Format.TextFrame2.TextRange.Font.Name = FONT_FAMILY
            
            ' Soften primary gridlines
            If .HasAxis(xlValue, xlPrimary) Then
                With .Axes(xlValue, xlPrimary)
                    If .HasMajorGridlines Then
                        .MajorGridlines.Format.Line.ForeColor.RGB = PWC_CARD_BORDER
                        .MajorGridlines.Format.Line.Weight = 0.75
                    End If
                    .Format.Line.Visible = msoFalse
                    .TickLabels.Font.Name = FONT_FAMILY
                    .TickLabels.Font.Size = 8.5
                    .TickLabels.Font.Color = PWC_TEXT_MUTED
                End With
            End If
            
            ' Soften category axis
            If .HasAxis(xlCategory, xlPrimary) Then
                With .Axes(xlCategory, xlPrimary)
                    .Format.Line.ForeColor.RGB = PWC_CARD_BORDER
                    .TickLabels.Font.Name = FONT_FAMILY
                    .TickLabels.Font.Size = 8.5
                    .TickLabels.Font.Color = PWC_TEXT_MUTED
                End With
            End If
        End With
    End With
    On Error GoTo 0
End Sub

' ==============================================================================
' 9. DASHBOARD 1: CALL CENTER OPERATIONS COCKPIT
' ==============================================================================
Public Sub BuildCallCenterCanvas(Optional ByVal populateInitialData As Boolean = False)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sheetName As String
    sheetName = "03_CallCenter_Cockpit"
    
    Set wb = ActiveWorkbook
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If
    On Error GoTo 0
    
    ws.Tab.Color = PWC_ORANGE
    
    ' Clear previous shapes
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    
    ' 1. Canvas Setup
    InitializeDashboardCanvas ws, PWC_CANVAS_BG
    
    ' 2. Web App Top Navigation Bar with Logo & Live Indicator
    BuildWebTopNavBar ws, "CC"
    
    ' 3. Executive Hero Header with Integrated SVG Action Buttons
    BuildHeroHeader ws, _
                    "Call Centre Operations & SLA Performance Cockpit", _
                    "Operational SLA Monitoring, Queue Abandonment Forensics & Agent Quality Auditing (Q1 2021)", _
                    76
    
    ' 4. Row of 5 Floating BAN KPI Metric Cards (Customizable Input)
    Dim cardTop As Single
    cardTop = 130
    
    Dim val1 As String, val2 As String, val3 As String, val4 As String, val5 As String
    If populateInitialData Then
        val1 = "5,000": val2 = "81.08%": val3 = "18.92%": val4 = "67.52 s": val5 = "3.40 / 5.0"
    Else
        val1 = "--": val2 = "--": val3 = "--": val4 = "--": val5 = "--"
    End If
    
    BuildKPICard ws, "CC_TotalDemand", CANVAS_LEFT, cardTop, KPI_WIDTH, KPI_HEIGHT, _
                 "Total Call Intake", val1, "Gross Intake Demand | 100% Logged", PWC_TEXT_TITLE
                 
    BuildKPICard ws, "CC_Answered", CANVAS_LEFT + (KPI_WIDTH + CARD_GAP) * 1, cardTop, KPI_WIDTH, KPI_HEIGHT, _
                 "Operational Answer Rate", val2, "SLA Target: >= 80.0% Connected", PWC_SUCCESS_GREEN
                 
    BuildKPICard ws, "CC_Abandoned", CANVAS_LEFT + (KPI_WIDTH + CARD_GAP) * 2, cardTop, KPI_WIDTH, KPI_HEIGHT, _
                 "Queue Abandonment Rate", val3, "SLA Threshold: <= 15.0% Dropped", PWC_ALERT_RED
                 
    BuildKPICard ws, "CC_ASA", CANVAS_LEFT + (KPI_WIDTH + CARD_GAP) * 3, cardTop, KPI_WIDTH, KPI_HEIGHT, _
                 "Avg Speed of Answer", val4, "Target: <= 60.0 Seconds Queue Wait", PWC_WARNING_AMBER
                 
    BuildKPICard ws, "CC_CSAT", CANVAS_LEFT + (KPI_WIDTH + CARD_GAP) * 4, cardTop, KPI_WIDTH, KPI_HEIGHT, _
                 "Average CSAT Rating", val5, "Service Benchmark: >= 3.50 / 5.0", PWC_ORANGE
                 
    ' 5. Left Slicer Control Panel Container
    Dim bodyTop As Single
    bodyTop = 226
    
    BuildSlicerPanelContainer ws, "CC_Slicers", CANVAS_LEFT, bodyTop, SLICER_WIDTH, SLICER_HEIGHT, _
                              "Date Period (Month)", _
                              "Inquiry Topic Tier", _
                              "Representative Agent"
                              
    ' 6. Visual Containers Grid (2 Columns x 2 Rows)
    Dim colMidLeft As Single, colRightLeft As Single
    colMidLeft = CANVAS_LEFT + SLICER_WIDTH + CARD_GAP         ' 270 pt
    colRightLeft = colMidLeft + VISUAL_WIDTH + CARD_GAP       ' 762 pt
    
    Dim rowTopPos As Single, rowBottomPos As Single
    rowTopPos = bodyTop                                       ' 226 pt
    rowBottomPos = rowTopPos + VISUAL_HEIGHT + 14             ' 510 pt
    
    ' Middle-Top: Hourly Call Volume & Queue Arrival
    BuildChartContainer ws, "CC_HourlyVolume", colMidLeft, rowTopPos, VISUAL_WIDTH, VISUAL_HEIGHT, _
                        "Intraday Demand Surge & Queue Triage", _
                        "Hourly Call Arrival Pattern (09:00 - 18:00) vs Pickup Status", _
                        "Column Chart"
                        
    ' Middle-Bottom: Topic SLA Breakdown
    BuildChartContainer ws, "CC_TopicBreakdown", colMidLeft, rowBottomPos, VISUAL_WIDTH, VISUAL_HEIGHT, _
                        "Inquiry Topic SLA Compliance & Speed", _
                        "Contract, Tech Support, Payment, Billing & Admin Volume Breakdown", _
                        "Clustered Bar"
                        
    ' Right-Top: Agent Performance Quadrant
    BuildChartContainer ws, "CC_AgentQuadrant", colRightLeft, rowTopPos, VISUAL_WIDTH, VISUAL_HEIGHT, _
                        "Representative Efficiency Matrix", _
                        "Speed of Answer vs Resolution Rate by Representative", _
                        "Scatter Plot"
                        
    ' Right-Bottom: Agent Quality & CSAT Audit
    BuildChartContainer ws, "CC_AgentScorecard", colRightLeft, rowBottomPos, VISUAL_WIDTH, VISUAL_HEIGHT, _
                        "Representative Quality & CSAT Audit", _
                        "FCR %, Answer Speed, CSAT Ratings & Assigned Tier", _
                        "Matrix Table"
                        
    ' 7. Automated Data Model CUBE Staging & Formula Linking
    Call AutomateAndLinkKPICards(ws, "CC")
    
    ws.Range("A1").Select
End Sub

' ==============================================================================
' 10. DASHBOARD 2: CUSTOMER CHURN & REVENUE RISK COCKPIT
' ==============================================================================
Public Sub BuildCustomerRetentionCanvas(Optional ByVal populateInitialData As Boolean = False)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sheetName As String
    sheetName = "04_CustomerRetention_Cockpit"
    
    Set wb = ActiveWorkbook
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If
    On Error GoTo 0
    
    ws.Tab.Color = PWC_CHARCOAL
    
    ' Clear previous shapes
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    
    ' 1. Canvas Setup
    InitializeDashboardCanvas ws, PWC_CANVAS_BG
    
    ' 2. Web App Top Navigation Bar with Logo & Live Indicator
    BuildWebTopNavBar ws, "CH"
    
    ' 3. Executive Hero Header with Integrated SVG Action Buttons
    BuildHeroHeader ws, _
                    "Customer Retention & Revenue Risk Cockpit", _
                    "Subscriber Attrition Forensics, ARR Revenue Exposure & Commitment Contract Vulnerability", _
                    76
    
    ' 4. Row of 5 Floating BAN KPI Metric Cards (Customizable Input)
    Dim cardTop As Single
    cardTop = 130
    
    Dim val1 As String, val2 As String, val3 As String, val4 As String, val5 As String
    If populateInitialData Then
        val1 = "7,043": val2 = "26.54%": val3 = "$2.86M": val4 = "42.71%": val5 = "1.42"
    Else
        val1 = "--": val2 = "--": val3 = "--": val4 = "--": val5 = "--"
    End If
    
    BuildKPICard ws, "CH_Subscribers", CANVAS_LEFT, cardTop, KPI_WIDTH, KPI_HEIGHT, _
                 "Total Active Accounts", val1, "Total Active Subscriber Portfolio", PWC_TEXT_TITLE
                 
    BuildKPICard ws, "CH_ChurnRate", CANVAS_LEFT + (KPI_WIDTH + CARD_GAP) * 1, cardTop, KPI_WIDTH, KPI_HEIGHT, _
                 "Customer Churn Rate", val2, "Operational Target: < 20.0% Churn", PWC_ALERT_RED
                 
    BuildKPICard ws, "CH_ARRRisk", CANVAS_LEFT + (KPI_WIDTH + CARD_GAP) * 2, cardTop, KPI_WIDTH, KPI_HEIGHT, _
                 "Annual Revenue at Risk", val3, "Annualized Lost ARR Exposure", PWC_ORANGE
                 
    BuildKPICard ws, "CH_M2MChurn", CANVAS_LEFT + (KPI_WIDTH + CARD_GAP) * 3, cardTop, KPI_WIDTH, KPI_HEIGHT, _
                 "Month-to-Month Churn", val4, "Target: < 25.0% Commitment Retention", PWC_ALERT_RED
                 
    BuildKPICard ws, "CH_Tickets", CANVAS_LEFT + (KPI_WIDTH + CARD_GAP) * 4, cardTop, KPI_WIDTH, KPI_HEIGHT, _
                 "Tech Tickets per Churn", val5, "Friction Index: 3+ Tickets Spikes Churn", PWC_WARNING_AMBER
                 
    ' 5. Left Slicer Control Panel Container
    Dim bodyTop As Single
    bodyTop = 226
    
    BuildSlicerPanelContainer ws, "CH_Slicers", CANVAS_LEFT, bodyTop, SLICER_WIDTH, SLICER_HEIGHT, _
                              "Commitment Contract", _
                              "Payment Method Tier", _
                              "Internet Service Type"
                              
    ' 6. Visual Containers Grid (2 Columns x 2 Rows)
    Dim colMidLeft As Single, colRightLeft As Single
    colMidLeft = CANVAS_LEFT + SLICER_WIDTH + CARD_GAP         ' 270 pt
    colRightLeft = colMidLeft + VISUAL_WIDTH + CARD_GAP       ' 762 pt
    
    Dim rowTopPos As Single, rowBottomPos As Single
    rowTopPos = bodyTop                                       ' 226 pt
    rowBottomPos = rowTopPos + VISUAL_HEIGHT + 14             ' 510 pt
    
    ' Middle-Top: Churn by Contract
    BuildChartContainer ws, "CH_ContractRisk", colMidLeft, rowTopPos, VISUAL_WIDTH, VISUAL_HEIGHT, _
                        "Churn Rate % by Commitment Contract", _
                        "Month-to-Month Vulnerability vs 1-Year & 2-Year Commitments", _
                        "Column Chart"
                        
    ' Middle-Bottom: Tenure Cohort Attrition Curve
    BuildChartContainer ws, "CH_TenureCohort", colMidLeft, rowBottomPos, VISUAL_WIDTH, VISUAL_HEIGHT, _
                        "Tenure Attrition Curve & Vulnerability", _
                        "Early Risk Window (Months 1-12) vs Established Subscribers", _
                        "Area / Line Chart"
                        
    ' Right-Top: Payment Method Friction
    BuildChartContainer ws, "CH_PaymentFriction", colRightLeft, rowTopPos, VISUAL_WIDTH, VISUAL_HEIGHT, _
                        "Payment Method Risk Diagnostics", _
                        "Electronic Check Friction vs Automated Credit Card & Bank Transfers", _
                        "Clustered Bar"
                        
    ' Right-Bottom: Service Matrix & Tech Add-ons
    BuildChartContainer ws, "CH_ServiceMatrix", colRightLeft, rowBottomPos, VISUAL_WIDTH, VISUAL_HEIGHT, _
                        "Internet Service & Add-On Protection", _
                        "Fiber Optic Churn Exposure vs Online Security & Tech Support", _
                        "Matrix Table"
                        
    ' 7. Automated Data Model CUBE Staging & Formula Linking
    Call AutomateAndLinkKPICards(ws, "CH")
    
    ws.Range("A1").Select
End Sub

' ==============================================================================
' 11. DASHBOARD 3: DIVERSITY, EQUITY & INCLUSION COCKPIT
' ==============================================================================
Public Sub BuildDiversityInclusionCanvas(Optional ByVal populateInitialData As Boolean = False)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sheetName As String
    sheetName = "05_DiversityInclusion_Cockpit"
    
    Set wb = ActiveWorkbook
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If
    On Error GoTo 0
    
    ws.Tab.Color = PWC_TEXT_MUTED
    
    ' Clear previous shapes
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    
    ' 1. Canvas Setup
    InitializeDashboardCanvas ws, PWC_CANVAS_BG
    
    ' 2. Web App Top Navigation Bar with Logo & Live Indicator
    BuildWebTopNavBar ws, "DI"
    
    ' 3. Executive Hero Header with Integrated SVG Action Buttons
    BuildHeroHeader ws, _
                    "Diversity, Equity & Executive Parity Cockpit", _
                    "Workforce Pipeline Governance, Broken Rung Diagnostics & Promotion Velocity Parity (FY21)", _
                    76
    
    ' 4. Row of 5 Floating BAN KPI Metric Cards (Customizable Input)
    Dim cardTop As Single
    cardTop = 130
    
    Dim val1 As String, val2 As String, val3 As String, val4 As String, val5 As String
    If populateInitialData Then
        val1 = "500": val2 = "41.00%": val3 = "-19.67%": val4 = "35.29%": val5 = "+5.6 Mos"
    Else
        val1 = "--": val2 = "--": val3 = "--": val4 = "--": val5 = "--"
    End If
    
    BuildKPICard ws, "DI_Workforce", CANVAS_LEFT, cardTop, KPI_WIDTH, KPI_HEIGHT, _
                 "Total Corporate Census", val1, "Active Enterprise Headcount Base", PWC_TEXT_TITLE
                 
    BuildKPICard ws, "DI_FemaleShare", CANVAS_LEFT + (KPI_WIDTH + CARD_GAP) * 1, cardTop, KPI_WIDTH, KPI_HEIGHT, _
                 "Female Headcount Share", val2, "Corporate Parity Target: 50.0%", PWC_WARNING_AMBER
                 
    BuildKPICard ws, "DI_BrokenRung", CANVAS_LEFT + (KPI_WIDTH + CARD_GAP) * 2, cardTop, KPI_WIDTH, KPI_HEIGHT, _
                 "Broken Rung Cliff Gap", val3, "Critical Manager -> Sr Mgr Pipeline Leak", PWC_ALERT_RED
                 
    BuildKPICard ws, "DI_PromoShare", CANVAS_LEFT + (KPI_WIDTH + CARD_GAP) * 3, cardTop, KPI_WIDTH, KPI_HEIGHT, _
                 "FY21 Promotions Awarded", val4, "Promotions Gender Parity Baseline", PWC_ORANGE
                 
    BuildKPICard ws, "DI_TimeInGrade", CANVAS_LEFT + (KPI_WIDTH + CARD_GAP) * 4, cardTop, KPI_WIDTH, KPI_HEIGHT, _
                 "Promotion Velocity Gap", val5, "Target: Zero Gender Velocity Variance", PWC_ALERT_RED
                 
    ' 5. Left Slicer Control Panel Container
    Dim bodyTop As Single
    bodyTop = 226
    
    BuildSlicerPanelContainer ws, "DI_Slicers", CANVAS_LEFT, bodyTop, SLICER_WIDTH, SLICER_HEIGHT, _
                              "Department Group", _
                              "Job Level Tier", _
                              "Age Demographic Cohort"
                              
    ' 6. Visual Containers Grid (2 Columns x 2 Rows)
    Dim colMidLeft As Single, colRightLeft As Single
    colMidLeft = CANVAS_LEFT + SLICER_WIDTH + CARD_GAP         ' 270 pt
    colRightLeft = colMidLeft + VISUAL_WIDTH + CARD_GAP       ' 762 pt
    
    Dim rowTopPos As Single, rowBottomPos As Single
    rowTopPos = bodyTop                                       ' 226 pt
    rowBottomPos = rowTopPos + VISUAL_HEIGHT + 14             ' 510 pt
    
    ' Middle-Top: Workforce Funnel / Broken Rung
    BuildChartContainer ws, "DI_PipelineFunnel", colMidLeft, rowTopPos, VISUAL_WIDTH, VISUAL_HEIGHT, _
                        "Workforce Hierarchy & Broken Rung Funnel", _
                        "Female vs Male Representation across Corporate Grades 1 to 6", _
                        "Funnel / Bar"
                        
    ' Middle-Bottom: Departmental Representation
    BuildChartContainer ws, "DI_DeptParity", colMidLeft, rowBottomPos, VISUAL_WIDTH, VISUAL_HEIGHT, _
                        "Departmental Representation & Target Gaps", _
                        "Operations, Sales, Finance, HR, IT & Strategy Headcount", _
                        "Clustered Column"
                        
    ' Right-Top: Promotion Velocity Gap
    BuildChartContainer ws, "DI_PromoVelocity", colRightLeft, rowTopPos, VISUAL_WIDTH, VISUAL_HEIGHT, _
                        "Promotion Velocity & Time in Grade", _
                        "Longitudinal Time-in-Grade Velocity Comparison by Gender", _
                        "Bar Chart"
                        
    ' Right-Bottom: Performance Appraisal & Equity Audit
    BuildChartContainer ws, "DI_PerformanceAudit", colRightLeft, rowBottomPos, VISUAL_WIDTH, VISUAL_HEIGHT, _
                        "Performance Appraisal & Promotion Equity", _
                        "FY20 Appraisal Rating vs Actual FY21 Promotion Award Rate", _
                        "Matrix Table"
                        
    ' 7. Automated Data Model CUBE Staging & Formula Linking
    Call AutomateAndLinkKPICards(ws, "DI")
    
    ws.Range("A1").Select
End Sub

' ==============================================================================
' 12. MASTER ONE-CLICK BUILDER: INITIALIZE ALL 3 DASHBOARD CANVASES
' ==============================================================================
Public Sub BuildAllDashboardCanvases()
    modAppState.FreezeAppState
    
    Call BuildCallCenterCanvas(False)
    Call BuildCustomerRetentionCanvas(False)
    Call BuildDiversityInclusionCanvas(False)
    
    ' Return focus to Call Center Cockpit
    On Error Resume Next
    ActiveWorkbook.Worksheets("03_CallCenter_Cockpit").Activate
    On Error GoTo 0
    
    modAppState.RestoreAppState
    
    If Application.UserControl Then
        MsgBox "PwC Web-App Style Dashboard Canvases successfully created!" & vbCrLf & vbCrLf & _
               "Created 3 Interactive Canvases:" & vbCrLf & _
               "  1. '03_CallCenter_Cockpit'" & vbCrLf & _
               "  2. '04_CustomerRetention_Cockpit'" & vbCrLf & _
               "  3. '05_DiversityInclusion_Cockpit'" & vbCrLf & vbCrLf & _
               "Key Enhancements:" & vbCrLf & _
               "  - Vector SVG Action Buttons ([Refresh Data], [Reset Filters], [Export PDF])" & vbCrLf & _
               "  - Pure ASCII formatting (Zero UTF-8/ANSI character bugs)" & vbCrLf & _
               "  - Full connectivity to modDataRefresh, modFilterController, and modExportPDF" & vbCrLf & _
               "  - 5 BAN KPI Cards formatted with clean '--' placeholders" & vbCrLf & _
               "  - Left Global Filter Drawer with dedicated Slicer Slots" & vbCrLf & _
               "  - 2x2 Grid of 4 Visual Containers with dashed docking zones", _
               vbInformation, "PwC Design System Automation"
    End If
End Sub

' ==============================================================================
' 13. MASTER SUITE RUNNER: RUN ENTIRE PWC PLATFORM END-TO-END
' ==============================================================================
Public Sub RunCompletePwCPlatform()
    modAppState.FreezeAppState
    Application.StatusBar = "Executing Complete PwC Platform Workflow..."
    
    ' Step 1: Ensure Governance sheets exist
    On Error Resume Next
    Call modCreateGovernanceSheets.BuildGovernanceArchitecture
    On Error GoTo 0
    
    ' Step 2: Build All Dashboard Presentation Canvases
    Call BuildAllDashboardCanvases
    
    ' Step 3: Refresh VertiPaq Data Model & PivotCaches
    Call modDataRefresh.RefreshPipelineSynchronously
    
    ' Step 4: Clear Filters across all Dashboards
    Call modFilterController.ClearAllFilters
    
    modAppState.RestoreAppState
    
    If Application.UserControl Then
        MsgBox "PwC Virtual Case Platform successfully initialized end-to-end!" & vbCrLf & vbCrLf & _
               "1. Governance Architecture Verified ('01_Business_Domains', '02_Metadata_&_KPI_Catalog')" & vbCrLf & _
               "2. 3 Web-App Canvases Provisioned ('03_CallCenter', '04_Retention', '05_D&I')" & vbCrLf & _
               "3. Interactive Vector SVG Action Buttons Connected" & vbCrLf & _
               "4. Data Pipeline & PivotCaches Synchronized" & vbCrLf & _
               "5. Slicers & Global Filter Drawers Ready", _
               vbInformation, "PwC Master Orchestrator"
    End If
End Sub
