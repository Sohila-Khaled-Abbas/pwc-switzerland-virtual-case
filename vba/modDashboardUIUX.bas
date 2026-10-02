Attribute VB_Name = "modDashboardUIUX"
Option Explicit

' ==============================================================================
' PwC Switzerland Digital Accelerator — Executive Dashboard UI/UX Design System
' Implements Modern Business Intelligence Card-Based Layouts & Brand Styling
' Based on Industry UI/UX Best Practices (Treating Worksheets as Presentation Canvases)
' ==============================================================================

' ------------------------------------------------------------------------------
' PwC Enterprise Brand Color Constants (RGB Values)
' ------------------------------------------------------------------------------
Public Const PWC_ORANGE         As Long = 133288     ' #D04A02 - RGB(208, 74, 2)
Public Const PWC_DARK_SLATE     As Long = 2758415    ' #0F172A - RGB(42, 23, 15) -> RGB(15, 23, 42)
Public Const PWC_CANVAS_BG      As Long = 16579320   ' #F8FAFC - RGB(248, 250, 252)
Public Const PWC_CARD_FILL      As Long = 16777215   ' #FFFFFF - RGB(255, 255, 255)
Public Const PWC_CARD_BORDER    As Long = 15790322   ' #E2E8F0 - RGB(226, 232, 240)
Public Const PWC_TEXT_TITLE     As Long = 2434341    ' #1E293B - RGB(30, 41, 59)
Public Const PWC_TEXT_MUTED     As Long = 9141108    ' #64748B - RGB(100, 116, 139)
Public Const PWC_SUCCESS_GREEN  As Long = 4363781    ' #059669 - RGB(5, 150, 105)
Public Const PWC_ALERT_RED      As Long = 2500316    ' #DC2626 - RGB(220, 38, 38)
Public Const PWC_WARNING_AMBER  As Long = 422009     ' #D97706 - RGB(217, 119, 6)
Public Const PWC_ACCENT_BLUE    As Long = 16298552   ' #38BDF8 - RGB(56, 189, 248)

' Default Enterprise Typography
Public Const FONT_FAMILY        As String = "Segoe UI"

' ==============================================================================
' 1. MASTER WORKSPACE CANVAS INITIALIZER
' Hides gridlines, headers, and applies the neutral off-white canvas background
' ==============================================================================
Public Sub InitializeDashboardCanvas(ws As Worksheet, Optional ByVal bgColor As Long = PWC_CANVAS_BG)
    On Error Resume Next
    Application.ScreenUpdating = False
    
    ws.Activate
    
    ' Hide Excel spreadsheet gridlines and row/column headers for presentation-ready UI
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.Zoom = 100
    
    ' Flood worksheet with uniform canvas background
    With ws.Cells.Interior
        .Pattern = xlSolid
        .Color = bgColor
    End With
    
    Application.ScreenUpdating = True
End Sub

' ==============================================================================
' 2. EXECUTIVE HEADER BANNER GENERATOR
' Creates a sleek navigation bar with PwC brand accent bar, title & metadata
' ==============================================================================
Public Sub BuildHeaderBanner(ws As Worksheet, _
                            ByVal dashboardTitle As String, _
                            ByVal subtitle As String, _
                            Optional ByVal topPos As Single = 15, _
                            Optional ByVal leftPos As Single = 20, _
                            Optional ByVal bannerWidth As Single = 1200, _
                            Optional ByVal bannerHeight As Single = 65)
    Dim shpBanner As Shape
    Dim shpAccent As Shape
    Dim shpTitle As Shape
    
    ' 1. Base Banner Container (Card)
    Set shpBanner = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, bannerWidth, bannerHeight)
    With shpBanner
        .Name = "Header_Banner"
        .Fill.Solid
        .Fill.ForeColor.RGB = PWC_CARD_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER
        .Line.Weight = 1
        .Adjustments.Item(1) = 0.08
        ' Soft modern drop shadow
        With .Shadow
            .Type = msoShadow21
            .Visible = msoTrue
            .Blur = 8
            .Transparency = 0.88
            .OffsetX = 0
            .OffsetY = 3
        End With
    End With
    
    ' 2. Left Brand Accent Bar (PwC Orange)
    Set shpAccent = ws.Shapes.AddShape(msoShapeRectangle, leftPos, topPos + 6, 6, bannerHeight - 12)
    With shpAccent
        .Name = "Header_AccentBar"
        .Fill.Solid
        .Fill.ForeColor.RGB = PWC_ORANGE
        .Line.Visible = msoFalse
    End With
    
    ' 3. Title & Subtitle Text Box
    Set shpTitle = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos + 22, topPos + 8, bannerWidth - 250, bannerHeight - 16)
    With shpTitle
        .Name = "Header_TitleText"
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2
            .MarginLeft = 0
            .MarginTop = 0
            .MarginRight = 0
            .MarginBottom = 0
            .WordWrap = msoFalse
            With .TextRange
                .Text = dashboardTitle & vbCrLf & subtitle
                ' Title Formatting
                With .Paragraphs(1).Font
                    .Name = FONT_FAMILY
                    .Size = 16
                    .Bold = msoTrue
                    .Fill.ForeColor.RGB = PWC_TEXT_TITLE
                End With
                ' Subtitle Formatting
                With .Paragraphs(2).Font
                    .Name = FONT_FAMILY
                    .Size = 9.5
                    .Bold = msoFalse
                    .Fill.ForeColor.RGB = PWC_TEXT_MUTED
                End With
            End With
        End With
    End With
End Sub

' ==============================================================================
' 3. FLOATING KPI SCORECARD GENERATOR
' Generates modern card components with large metrics, micro-labels & SLA pills
' ==============================================================================
Public Sub BuildKPICard(ws As Worksheet, _
                        ByVal cardName As String, _
                        ByVal leftPos As Single, _
                        ByVal topPos As Single, _
                        ByVal cardWidth As Single, _
                        ByVal cardHeight As Single, _
                        ByVal kpiLabel As String, _
                        ByVal kpiValue As String, _
                        ByVal kpiSubtext As String, _
                        Optional ByVal statusColor As Long = PWC_ORANGE)
    Dim shpCard As Shape
    Dim shpAccentLine As Shape
    Dim shpText As Shape
    
    ' 1. Card Container with Rounded Corners & Soft Drop Shadow
    Set shpCard = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, cardWidth, cardHeight)
    With shpCard
        .Name = "Card_" & cardName
        .Fill.Solid
        .Fill.ForeColor.RGB = PWC_CARD_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER
        .Line.Weight = 1
        .Adjustments.Item(1) = 0.12
        With .Shadow
            .Type = msoShadow21
            .Visible = msoTrue
            .Blur = 6
            .Transparency = 0.85
            .OffsetX = 0
            .OffsetY = 2
        End With
    End With
    
    ' 2. Top Color Indicator Line
    Set shpAccentLine = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos + 12, topPos + 6, cardWidth - 24, 3)
    With shpAccentLine
        .Name = "Accent_" & cardName
        .Fill.Solid
        .Fill.ForeColor.RGB = statusColor
        .Line.Visible = msoFalse
        .Adjustments.Item(1) = 0.5
    End With
    
    ' 3. KPI Text Content
    Set shpText = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos + 12, topPos + 14, cardWidth - 24, cardHeight - 20)
    With shpText
        .Name = "Text_" & cardName
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2
            .MarginLeft = 0
            .MarginTop = 0
            .MarginRight = 0
            .MarginBottom = 0
            .WordWrap = msoTrue
            With .TextRange
                .Text = UCase(kpiLabel) & vbCrLf & kpiValue & vbCrLf & kpiSubtext
                ' Label
                With .Paragraphs(1).Font
                    .Name = FONT_FAMILY
                    .Size = 8.5
                    .Bold = msoTrue
                    .Fill.ForeColor.RGB = PWC_TEXT_MUTED
                End With
                ' Big Value Callout
                With .Paragraphs(2).Font
                    .Name = FONT_FAMILY
                    .Size = 22
                    .Bold = msoTrue
                    .Fill.ForeColor.RGB = PWC_TEXT_TITLE
                End With
                ' Subtext / Target Status
                With .Paragraphs(3).Font
                    .Name = FONT_FAMILY
                    .Size = 8.5
                    .Bold = msoFalse
                    .Fill.ForeColor.RGB = statusColor
                End With
            End With
        End With
    End With
End Sub

' ==============================================================================
' 4. VISUAL CONTAINER CARD (CHART HOLDER)
' Creates container boxes where charts float cleanly with unified card styling
' ==============================================================================
Public Sub BuildChartContainer(ws As Worksheet, _
                               ByVal containerName As String, _
                               ByVal leftPos As Single, _
                               ByVal topPos As Single, _
                               ByVal cardWidth As Single, _
                               ByVal cardHeight As Single, _
                               ByVal chartTitle As String, _
                               Optional ByVal chartSubtitle As String = "")
    Dim shpContainer As Shape
    Dim shpHeader As Shape
    
    ' 1. Card Container
    Set shpContainer = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, cardWidth, cardHeight)
    With shpContainer
        .Name = "Container_" & containerName
        .Fill.Solid
        .Fill.ForeColor.RGB = PWC_CARD_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER
        .Line.Weight = 1
        .Adjustments.Item(1) = 0.05
        With .Shadow
            .Type = msoShadow21
            .Visible = msoTrue
            .Blur = 8
            .Transparency = 0.88
            .OffsetX = 0
            .OffsetY = 3
        End With
    End With
    
    ' 2. Container Header Title
    Set shpHeader = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos + 16, topPos + 10, cardWidth - 32, 35)
    With shpHeader
        .Name = "Header_" & containerName
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2
            .MarginLeft = 0
            .MarginTop = 0
            .MarginRight = 0
            .MarginBottom = 0
            With .TextRange
                If Len(chartSubtitle) > 0 Then
                    .Text = chartTitle & vbCrLf & chartSubtitle
                    With .Paragraphs(1).Font
                        .Name = FONT_FAMILY
                        .Size = 11.5
                        .Bold = msoTrue
                        .Fill.ForeColor.RGB = PWC_TEXT_TITLE
                    End With
                    With .Paragraphs(2).Font
                        .Name = FONT_FAMILY
                        .Size = 8.5
                        .Fill.ForeColor.RGB = PWC_TEXT_MUTED
                    End With
                Else
                    .Text = chartTitle
                    With .Paragraphs(1).Font
                        .Name = FONT_FAMILY
                        .Size = 11.5
                        .Bold = msoTrue
                        .Fill.ForeColor.RGB = PWC_TEXT_TITLE
                    End With
                End If
            End With
        End With
    End With
End Sub

' ==============================================================================
' 5. CHART TRANSPARENCY & DECLUTTERING FORMATTER
' Strips chart borders & background fills so charts seamlessly dock inside cards
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
End Sub

' ==============================================================================
' 6. ONE-CLICK DEMO: BUILD COMPLETE PWC CALL CENTRE INTELLIGENCE COCKPIT
' Builds the entire Task 1 Executive Canvas using these UI/UX principles
' ==============================================================================
Public Sub BuildFullCallCentreCockpit()
    Dim ws As Worksheet
    Dim sheetName As String
    sheetName = "CallCentre_Cockpit"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    End If
    On Error GoTo 0
    
    ' Clear previous shapes
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    
    ' 1. Canvas Setup
    InitializeDashboardCanvas ws, PWC_CANVAS_BG
    
    ' 2. Header Banner
    BuildHeaderBanner ws, _
                      "PwC Switzerland | Call Centre Operations Cockpit", _
                      "Operational SLA Monitoring, Abandonment Bottleneck Triage & Agent Performance (Q1 2021)", _
                      15, 20, 1160, 60
    
    ' 3. Row of 5 Floating KPI Cards
    Dim cardW As Single, cardH As Single, cardTop As Single
    cardW = 216
    cardH = 88
    cardTop = 90
    
    ' KPI 1: Inbound Volume
    BuildKPICard ws, "KPI1", 20, cardTop, cardW, cardH, _
                 "Total Inbound Calls", "5,000", "Target: 5,000 Inquiries", PWC_TEXT_TITLE
    
    ' KPI 2: Answer Rate
    BuildKPICard ws, "KPI2", 256, cardTop, cardW, cardH, _
                 "Calls Answered", "81.08%", "4,054 Connected | Target > 80%", PWC_SUCCESS_GREEN
                 
    ' KPI 3: Abandonment Rate
    BuildKPICard ws, "KPI3", 492, cardTop, cardW, cardH, _
                 "Abandonment Rate", "18.92%", "946 Lost In-Queue | Target < 15% 🚨", PWC_ALERT_RED
                 
    ' KPI 4: Speed of Answer
    BuildKPICard ws, "KPI4", 728, cardTop, cardW, cardH, _
                 "Avg Speed of Answer", "67.52 s", "Target: < 60.0 s ⚠️", PWC_WARNING_AMBER
                 
    ' KPI 5: Customer CSAT
    BuildKPICard ws, "KPI5", 964, cardTop, cardW, cardH, _
                 "Customer Satisfaction", "3.40 / 5.0", "4,054 Survey Responses | Target > 3.50", PWC_ORANGE
    
    ' 4. Chart / Visual Containers
    ' Top Left: Intraday Volume Heatmap
    BuildChartContainer ws, "HourlyVolume", 20, 195, 560, 280, _
                        "Intraday Inbound Call Volume & Peak Queue Surge", _
                        "Hour of Day (09:00 - 18:00) vs Daily Connection Status"
                        
    ' Top Right: Agent Performance Quadrant
    BuildChartContainer ws, "AgentQuadrant", 600, 195, 580, 280, _
                        "2D Agent Performance Quadrant Matrix", _
                        "Average Talk Duration vs Total Calls Handled"
                        
    ' Bottom Left: Topic SLA Breakdown
    BuildChartContainer ws, "TopicBreakdown", 20, 490, 560, 240, _
                        "Inquiry Topic SLA Compliance & Speed of Answer", _
                        "Streaming, Tech Support, Payment, Billing & Contract"
                        
    ' Bottom Right: Agent Scorecard Audit Table
    BuildChartContainer ws, "AgentTable", 600, 490, 580, 240, _
                        "Representative Performance & CSAT Scorecard", _
                        "Resolution Rate %, Avg Speed, CSAT and Performance Tier"
                        
    ' Select A1
    ws.Range("A1").Select
    MsgBox "PwC Executive Dashboard Canvas successfully generated!", vbInformation, "PwC Design System"
End Sub
