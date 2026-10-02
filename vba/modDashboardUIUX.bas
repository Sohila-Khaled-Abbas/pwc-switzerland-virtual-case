Attribute VB_Name = "modDashboardUIUX"
Option Explicit

' ==============================================================================
' PwC Switzerland Digital Accelerator — Executive Dashboard UI/UX Design System
' Implements Modern Business Intelligence Card-Based Layouts & Brand Styling
' Based on Industry UI/UX Best Practices (Treating Worksheets as Presentation Canvases)
'
' Features:
'   1. InitializeDashboardCanvas (Kills gridlines, headings, applies #F8FAFC canvas)
'   2. BuildHeaderBanner (Corporate Charcoal ribbon, Tangerine accent, breadcrumbs)
'   3. BuildKPICard (BAN containers with micro-labels, big metrics, status pills)
'   4. BuildChartContainer (Floating card holders with rounded corners and drop shadows)
'   5. DeclutterAndFormatChart (Strips chart background/borders for transparent docking)
'   6. Individual Canvas Builders:
'      - BuildCallCenterCanvas
'      - BuildCustomerRetentionCanvas
'      - BuildDiversityInclusionCanvas
'   7. Master Orchestrator:
'      - BuildAllDashboardCanvases
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
Public Const PWC_TEXT_TITLE     As Long = 3877150    ' #1E293B - RGB(30, 41, 59)
Public Const PWC_TEXT_MUTED     As Long = 9141108    ' #64748B - RGB(100, 116, 139)
Public Const PWC_SUCCESS_GREEN  As Long = 4363781    ' #059669 - RGB(5, 150, 105)
Public Const PWC_ALERT_RED      As Long = 2500316    ' #DC2626 - RGB(220, 38, 38)
Public Const PWC_WARNING_AMBER  As Long = 422009     ' #D97706 - RGB(217, 119, 6)
Public Const PWC_ACCENT_BLUE    As Long = 16298552   ' #38BDF8 - RGB(56, 189, 248)

' Default Enterprise Typography
Public Const FONT_FAMILY        As String = "Segoe UI"

' ==============================================================================
' 1. MASTER WORKSPACE CANVAS INITIALIZER
' Hides gridlines, headings, and applies the neutral off-white canvas background
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
' Creates a sleek navigation bar with PwC brand accent bar, title & breadcrumbs
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
                    .Size = 15
                    .Bold = msoTrue
                    .Fill.ForeColor.RGB = PWC_TEXT_TITLE
                End With
                ' Subtitle Formatting
                With .Paragraphs(2).Font
                    .Name = FONT_FAMILY
                    .Size = 9
                    .Bold = msoFalse
                    .Fill.ForeColor.RGB = PWC_TEXT_MUTED
                End With
            End With
        End With
    End With
    
    ' 4. Breadcrumb Navigation Button (Back to Business Domains)
    Dim shpNav As Shape
    Set shpNav = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos + bannerWidth - 190, topPos + 16, 175, 32)
    With shpNav
        .Name = "Header_NavDomains"
        .Fill.Solid
        .Fill.ForeColor.RGB = PWC_CHARCOAL
        .Line.Visible = msoFalse
        .Adjustments.Item(1) = 0.2
        With .TextFrame2
            .MarginLeft = 0
            .MarginTop = 0
            .MarginRight = 0
            .MarginBottom = 0
            .WordWrap = msoFalse
            With .TextRange
                .Text = "<-- Executive Hub"
                .Font.Name = FONT_FAMILY
                .Font.Size = 9.5
                .Font.Bold = msoTrue
                .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .ParagraphFormat.Alignment = msoAlignCenter
            End With
        End With
    End With
    ws.Hyperlinks.Add Anchor:=shpNav, Address:="", SubAddress:="'01_Business_Domains'!A1"
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
    Set shpText = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos + 12, topPos + 12, cardWidth - 24, cardHeight - 16)
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
                    .Size = 8
                    .Bold = msoTrue
                    .Fill.ForeColor.RGB = PWC_TEXT_MUTED
                End With
                ' Big Value Callout
                With .Paragraphs(2).Font
                    .Name = FONT_FAMILY
                    .Size = 20
                    .Bold = msoTrue
                    .Fill.ForeColor.RGB = PWC_TEXT_TITLE
                End With
                ' Subtext / Target Status
                With .Paragraphs(3).Font
                    .Name = FONT_FAMILY
                    .Size = 8
                    .Bold = msoFalse
                    .Fill.ForeColor.RGB = statusColor
                End With
            End With
        End With
    End With
End Sub

' ==============================================================================
' 4. VISUAL CONTAINER CARD (CHART / SLICER HOLDER)
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
    Set shpHeader = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos + 16, topPos + 10, cardWidth - 32, 34)
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
                        .Size = 11
                        .Bold = msoTrue
                        .Fill.ForeColor.RGB = PWC_TEXT_TITLE
                    End With
                    With .Paragraphs(2).Font
                        .Name = FONT_FAMILY
                        .Size = 8
                        .Fill.ForeColor.RGB = PWC_TEXT_MUTED
                    End With
                Else
                    .Text = chartTitle
                    With .Paragraphs(1).Font
                        .Name = FONT_FAMILY
                        .Size = 11
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
' 6. DASHBOARD 1: CALL CENTER OPERATIONS COCKPIT
' ==============================================================================
Public Sub BuildCallCenterCanvas()
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
    
    ' 2. Header Banner
    BuildHeaderBanner ws, _
                      "PwC Switzerland | Call Centre Operations Cockpit", _
                      "Operational SLA Monitoring, Abandonment Bottleneck Triage & Agent Performance (Q1 2021)", _
                      15, 20, 1180, 60
    
    ' 3. Row of 5 Floating KPI Cards (Width 224, Gap 15)
    Dim cardW As Single, cardH As Single, cardTop As Single
    cardW = 224
    cardH = 86
    cardTop = 88
    
    BuildKPICard ws, "CC_TotalDemand", 20, cardTop, cardW, cardH, _
                 "Total Inbound Calls", "5,000", "Gross Volume | 100% Captured", PWC_TEXT_TITLE
                 
    BuildKPICard ws, "CC_Answered", 259, cardTop, cardW, cardH, _
                 "Calls Answered", "81.08%", "4,054 Connected | Target > 80% [OK]", PWC_SUCCESS_GREEN
                 
    BuildKPICard ws, "CC_Abandoned", 498, cardTop, cardW, cardH, _
                 "Abandonment Rate", "18.92%", "946 Queue Drops | Target < 15% [ALERT]", PWC_ALERT_RED
                 
    BuildKPICard ws, "CC_ASA", 737, cardTop, cardW, cardH, _
                 "Avg Speed of Answer", "67.52 s", "Queue Wait Time | Target < 60s [WARN]", PWC_WARNING_AMBER
                 
    BuildKPICard ws, "CC_CSAT", 976, cardTop, cardW, cardH, _
                 "Average CSAT", "3.40 / 5.0", "4,054 Ratings | Target > 3.50 [WARN]", PWC_ORANGE
                 
    ' 4. Left Slicer Control Panel Container
    BuildChartContainer ws, "CC_Slicers", 20, 186, 224, 530, _
                        "Interactive Filter Panel", _
                        "Global Slicers: Date, Topic & Agent"
                        
    ' 5. Chart / Visual Containers (Middle & Right Columns)
    ' Middle Column (Width 465)
    BuildChartContainer ws, "CC_HourlyVolume", 259, 186, 465, 255, _
                        "Intraday Demand Surge & Queue Triage", _
                        "Hourly Inbound Calls (09:00 - 18:00) vs Daily Connection Status"
                        
    BuildChartContainer ws, "CC_TopicBreakdown", 259, 453, 465, 263, _
                        "Inquiry Topic SLA Compliance & Speed", _
                        "Streaming, Tech Support, Payment, Billing & Contract Tiers"
                        
    ' Right Column (Width 476)
    BuildChartContainer ws, "CC_AgentQuadrant", 736, 186, 464, 255, _
                        "Agent Performance Matrix", _
                        "Handle Time Efficiency vs Calls Handled per Representative"
                        
    BuildChartContainer ws, "CC_AgentScorecard", 736, 453, 464, 263, _
                        "Representative Quality & CSAT Audit", _
                        "FCR %, Speed of Answer, CSAT & Assigned Performance Tier"
                        
    ws.Range("A1").Select
End Sub

' ==============================================================================
' 7. DASHBOARD 2: CUSTOMER CHURN & RETENTION COCKPIT
' ==============================================================================
Public Sub BuildCustomerRetentionCanvas()
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
    
    ' 2. Header Banner
    BuildHeaderBanner ws, _
                      "PwC Switzerland | Customer Churn & Revenue Risk Cockpit", _
                      "Subscriber Attrition Forensics, ARR Exposure ($2.86M) & Contract Longevity (7,043 Accounts)", _
                      15, 20, 1180, 60
    
    ' 3. Row of 5 Floating KPI Cards
    Dim cardW As Single, cardH As Single, cardTop As Single
    cardW = 224
    cardH = 86
    cardTop = 88
    
    BuildKPICard ws, "CH_Subscribers", 20, cardTop, cardW, cardH, _
                 "Total Subscriber Base", "7,043", "Active Subscription Portfolio", PWC_TEXT_TITLE
                 
    BuildKPICard ws, "CH_ChurnRate", 259, cardTop, cardW, cardH, _
                 "Customer Churn Rate", "26.54%", "1,869 Terminations | Target < 20% [ALERT]", PWC_ALERT_RED
                 
    BuildKPICard ws, "CH_ARRRisk", 498, cardTop, cardW, cardH, _
                 "Annual Revenue at Risk", "$2.86M", "Annualized MRR Lost | $238.4K/Mo [ALERT]", PWC_ORANGE
                 
    BuildKPICard ws, "CH_M2MChurn", 737, cardTop, cardW, cardH, _
                 "Month-to-Month Churn", "42.71%", "1,655 Lost Accounts | Target < 25% [CRITICAL]", PWC_ALERT_RED
                 
    BuildKPICard ws, "CH_Tickets", 976, cardTop, cardW, cardH, _
                 "Tech Tickets per Churn", "1.42", "Friction Window: 3+ Tickets Spikes Churn", PWC_WARNING_AMBER
                 
    ' 4. Left Slicer Control Panel Container
    BuildChartContainer ws, "CH_Slicers", 20, 186, 224, 530, _
                        "Interactive Filter Panel", _
                        "Global Slicers: Contract, Payment & Internet"
                        
    ' 5. Chart / Visual Containers (Middle & Right Columns)
    BuildChartContainer ws, "CH_ContractRisk", 259, 186, 465, 255, _
                        "Churn Rate % by Commitment Contract", _
                        "Month-to-Month (42.7%) vs 1-Year (11.3%) vs 2-Year (2.8%)"
                        
    BuildChartContainer ws, "CH_TenureCohort", 259, 453, 465, 263, _
                        "Tenure Attrition Curve & Vulnerability", _
                        "Early Risk Cohort (Months 1-12) vs Established Subscribers"
                        
    BuildChartContainer ws, "CH_PaymentFriction", 736, 186, 464, 255, _
                        "Payment Method Risk Diagnostics", _
                        "Electronic Check Latency vs Automated Credit Card & Bank Wire"
                        
    BuildChartContainer ws, "CH_ServiceMatrix", 736, 453, 464, 263, _
                        "Internet Service & Add-On Protection", _
                        "Fiber Optic Churn vs DSL & Online Tech Support Bundles"
                        
    ws.Range("A1").Select
End Sub

' ==============================================================================
' 8. DASHBOARD 3: DIVERSITY, EQUITY & INCLUSION COCKPIT
' ==============================================================================
Public Sub BuildDiversityInclusionCanvas()
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
    
    ' 2. Header Banner
    BuildHeaderBanner ws, _
                      "PwC Switzerland | Diversity, Equity & Executive Parity Cockpit", _
                      "Workforce Pipeline Governance, Broken Rung Diagnostics & Promotion Velocity (500 Employees)", _
                      15, 20, 1180, 60
    
    ' 3. Row of 5 Floating KPI Cards
    Dim cardW As Single, cardH As Single, cardTop As Single
    cardW = 224
    cardH = 86
    cardTop = 88
    
    BuildKPICard ws, "DI_Workforce", 20, cardTop, cardW, cardH, _
                 "Total Corporate Census", "500", "Workforce Personnel Base", PWC_TEXT_TITLE
                 
    BuildKPICard ws, "DI_FemaleShare", 259, cardTop, cardW, cardH, _
                 "Female Headcount Share", "41.00%", "205 Female | Corporate Parity Target: 50% [WARN]", PWC_WARNING_AMBER
                 
    BuildKPICard ws, "DI_BrokenRung", 498, cardTop, cardW, cardH, _
                 "Broken Rung Cliff", "-19.67%", "Manager 34.3% -> Sr Mgr 14.6% [CRITICAL]", PWC_ALERT_RED
                 
    BuildKPICard ws, "DI_PromoShare", 737, cardTop, cardW, cardH, _
                 "FY21 Promotions Awarded", "35.29%", "18 Female of 51 Total Promotions [WARN]", PWC_ORANGE
                 
    BuildKPICard ws, "DI_TimeInGrade", 976, cardTop, cardW, cardH, _
                 "Promotion Velocity Lag", "+5.6 Mos", "Female Time-in-Grade Promotion Gap [ALERT]", PWC_ALERT_RED
                 
    ' 4. Left Slicer Control Panel Container
    BuildChartContainer ws, "DI_Slicers", 20, 186, 224, 530, _
                        "Interactive Filter Panel", _
                        "Global Slicers: Department, Job Level & Age"
                        
    ' 5. Chart / Visual Containers (Middle & Right Columns)
    BuildChartContainer ws, "DI_PipelineFunnel", 259, 186, 465, 255, _
                        "Workforce Hierarchy & Broken Rung Funnel", _
                        "Female vs Male Representation across Corporate Grades 1 to 6"
                        
    BuildChartContainer ws, "DI_DeptParity", 259, 453, 465, 263, _
                        "Departmental Representation & Target Gaps", _
                        "HR, Finance, Operations, Sales, Marketing & Strategy"
                        
    BuildChartContainer ws, "DI_PromoVelocity", 736, 186, 464, 255, _
                        "Average Time in Grade Prior to Promotion", _
                        "Longitudinal Promotion Velocity Comparison by Gender"
                        
    BuildChartContainer ws, "DI_PerformanceAudit", 736, 453, 464, 263, _
                        "Performance Appraisal & Promotion Equity", _
                        "FY20 Appraisal Ratings (1-4) vs Actual FY21 Promotion Rate"
                        
    ws.Range("A1").Select
End Sub

' ==============================================================================
' 9. MASTER ONE-CLICK BUILDER: INITIALIZE ALL 3 DASHBOARD CANVASES
' Builds all 3 presentation canvases with zero manual clicking
' ==============================================================================
Public Sub BuildAllDashboardCanvases()
    Dim prevScreenUpdating As Boolean
    Dim prevAlerts As Boolean
    
    prevScreenUpdating = Application.ScreenUpdating
    prevAlerts = Application.DisplayAlerts
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Call BuildCallCenterCanvas
    Call BuildCustomerRetentionCanvas
    Call BuildDiversityInclusionCanvas
    
    ' Return focus to Call Center Cockpit
    ActiveWorkbook.Worksheets("03_CallCenter_Cockpit").Activate
    
    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevAlerts
    
    MsgBox "PwC Executive Dashboard Canvases successfully generated!" & vbCrLf & vbCrLf & _
           "Created 3 Presentation Canvases:" & vbCrLf & _
           "  • '03_CallCenter_Cockpit' (5 KPI Cards + 4 Visual Containers + Slicer Panel)" & vbCrLf & _
           "  • '04_CustomerRetention_Cockpit' (5 KPI Cards + 4 Visual Containers + Slicer Panel)" & vbCrLf & _
           "  • '05_DiversityInclusion_Cockpit' (5 KPI Cards + 4 Visual Containers + Slicer Panel)" & vbCrLf & vbCrLf & _
           "All canvases styled with PwC brand colors (#1E293B, #D04A02, #F8FAFC) and ready for visual insertion.", _
           vbInformation, "PwC Design System Automation"
End Sub
