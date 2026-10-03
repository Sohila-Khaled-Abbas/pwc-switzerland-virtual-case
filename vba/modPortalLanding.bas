Attribute VB_Name = "modPortalLanding"
Option Explicit

' ==============================================================================
' PwC Switzerland Virtual Case Experience - BI Executive Platform
' Module: modPortalLanding
' Description: Executive Homepage & Web Landing Page Builder
' Generates an ultra-modern SaaS web-style portal tab ("00_Home_Portal")
' featuring hero banner, cross-enterprise KPI summaries, interactive 
' cockpit launcher cards, animated pulse effects, and theme integration.
' Pure ASCII Compatible.
' ==============================================================================

Private Const SHEET_PORTAL As String = "00_Home_Portal"

' Palette Constants
Private Const COLOR_CANVAS As Long = 16579832        ' #F8FAFC
Private Const COLOR_CARD_BG As Long = 16777215       ' #FFFFFF
Private Const COLOR_BORDER As Long = 15790306        ' #E2E8F0
Private Const COLOR_DARK_NAVY As Long = 2762511      ' #0F172A
Private Const COLOR_PWC_ORANGE As Long = 1481168     ' #D04A02
Private Const COLOR_SLATE_TEXT As Long = 9141092     ' #64748B
Private Const COLOR_EMERALD As Long = 6886917        ' #059669
Private Const COLOR_ROSE As Long = 12458173          ' #BE185D

' ==============================================================================
' Public Entry Point: Build the Executive Landing Portal
' ==============================================================================
Public Sub BuildExecutivePortal()
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Dim shp As Shape
    
    ' Optimize Excel Environment
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ' 1. Check or Create "00_Home_Portal" as the FIRST worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_PORTAL)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = SHEET_PORTAL
    Else
        ' Move to first position
        ws.Move Before:=ThisWorkbook.Worksheets(1)
        ' Clear previous shapes on portal
        For Each shp In ws.Shapes
            shp.Delete
        Next shp
    End If
    
    ' 2. Canvas Setup: Clean Gridless Surface
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ws.Cells.Interior.Color = COLOR_CANVAS
    ws.Tab.Color = COLOR_PWC_ORANGE
    
    ' 3. Top Executive Header Bar (Full Width)
    CreateTopHeader ws
    
    ' 4. Hero Section: Web SaaS Landing Banner
    CreateHeroSection ws
    
    ' 5. Cross-Enterprise Performance Ticker (4 Metrics Cards)
    CreateEnterpriseMetricsTicker ws
    
    ' 6. The 3 Interactive Cockpit Launchers (Large SaaS Cards)
    CreateCockpitLauncherCards ws
    
    ' 7. Bottom Governance & Data Architecture Quick Links
    CreateGovernanceDrawer ws
    
    ' 8. Finalize View
    ws.Range("A1").Select
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    ' Trigger Entrance Pulse Animation
    AnimatePortalBanner
    Exit Sub
ErrorHandler:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "BuildExecutivePortal Error: " & Err.Description, vbCritical, "PwC Portal Builder"
End Sub

' ==============================================================================
' Component 1: Top Navigation Bar
' ==============================================================================
Private Sub CreateTopHeader(ByVal ws As Worksheet)
    Dim bar As Shape
    Dim logoTxt As Shape
    Dim statusPill As Shape
    Dim btnTheme As Shape
    Dim btnPDF As Shape
    
    ' Header Background Strip
    Set bar = ws.Shapes.AddShape(msoShapeRectangle, 20, 16, 1220, 52)
    bar.Name = "Nav_TopBar"
    bar.Fill.Solid
    bar.Fill.ForeColor.RGB = COLOR_DARK_NAVY
    bar.Line.Visible = msoFalse
    
    ' Logo & Portal Branding
    Set logoTxt = ws.Shapes.AddShape(msoShapeRectangle, 36, 24, 380, 36)
    logoTxt.Name = "Nav_BrandLogo"
    logoTxt.Fill.Visible = msoFalse
    logoTxt.Line.Visible = msoFalse
    With logoTxt.TextFrame2
        .WordWrap = msoFalse
        .MarginLeft = 0: .MarginTop = 0
        .TextRange.Text = "pwc  |  Enterprise Business Intelligence Suite"
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 13
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.Characters(1, 3).Font.Fill.ForeColor.RGB = RGB(255, 110, 38) ' Vibrant PwC Orange
        .TextRange.Characters(1, 3).Font.Size = 15
    End With
    
    ' Live Engine Status Pill Badge
    Set statusPill = ws.Shapes.AddShape(msoShapeRoundedRectangle, 440, 26, 210, 30)
    statusPill.Name = "Nav_StatusPill"
    statusPill.Adjustments.Item(1) = 0.5
    statusPill.Fill.Solid
    statusPill.Fill.ForeColor.RGB = RGB(30, 41, 59)
    statusPill.Line.ForeColor.RGB = RGB(51, 65, 85)
    statusPill.Line.Weight = 0.75
    With statusPill.TextFrame2
        .TextRange.Text = "[ACTIVE] VertiPaq Tabular Engine v2.1"
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 8.5
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(52, 211, 153) ' Emerald Green Glow
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    
    ' Theme Switcher Button
    Set btnTheme = ws.Shapes.AddShape(msoShapeRoundedRectangle, 1000, 26, 110, 30)
    btnTheme.Name = "btn_ThemeToggle"
    btnTheme.Adjustments.Item(1) = 0.5
    btnTheme.Fill.Solid
    btnTheme.Fill.ForeColor.RGB = RGB(30, 41, 59)
    btnTheme.Line.ForeColor.RGB = RGB(255, 110, 38)
    btnTheme.Line.Weight = 1
    btnTheme.OnAction = "modThemeEngine.ToggleDashboardTheme"
    With btnTheme.TextFrame2
        .TextRange.Text = "DARK MODE"
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 8.5
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    
    ' PDF Export Briefing Button
    Set btnPDF = ws.Shapes.AddShape(msoShapeRoundedRectangle, 1120, 26, 105, 30)
    btnPDF.Name = "btn_ExportBriefing"
    btnPDF.Adjustments.Item(1) = 0.5
    btnPDF.Fill.Solid
    btnPDF.Fill.ForeColor.RGB = COLOR_PWC_ORANGE
    btnPDF.Line.Visible = msoFalse
    btnPDF.OnAction = "modExportPDF.ExportActiveDashboardPDF"
    With btnPDF.TextFrame2
        .TextRange.Text = "Export PDF"
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 8.5
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
End Sub

' ==============================================================================
' Component 2: Hero Section Banner
' ==============================================================================
Private Sub CreateHeroSection(ByVal ws As Worksheet)
    Dim heroBg As Shape
    Dim heroTitle As Shape
    Dim heroDesc As Shape
    Dim heroBadge As Shape
    
    ' Hero Container (Subtle SaaS Gradient Surface)
    Set heroBg = ws.Shapes.AddShape(msoShapeRoundedRectangle, 20, 80, 1220, 120)
    heroBg.Name = "Hero_BannerCard"
    heroBg.Adjustments.Item(1) = 0.08
    heroBg.Fill.TwoColorGradient msoGradientHorizontal, 1
    heroBg.Fill.GradientStops.Item(1).Color = RGB(15, 23, 42)    ' Slate 900
    heroBg.Fill.GradientStops.Item(2).Color = RGB(30, 41, 59)    ' Slate 800
    heroBg.Line.ForeColor.RGB = RGB(51, 65, 85)
    heroBg.Line.Weight = 1
    ApplySoftElevation heroBg
    
    ' Hero Category Pill
    Set heroBadge = ws.Shapes.AddShape(msoShapeRoundedRectangle, 40, 94, 190, 22)
    heroBadge.Name = "Hero_BadgePill"
    heroBadge.Adjustments.Item(1) = 0.5
    heroBadge.Fill.Solid
    heroBadge.Fill.ForeColor.RGB = RGB(255, 110, 38)
    heroBadge.Line.Visible = msoFalse
    With heroBadge.TextFrame2
        .TextRange.Text = "PwC DIGITAL ACCELERATOR"
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 7.5
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    
    ' Hero Title
    Set heroTitle = ws.Shapes.AddShape(msoShapeRectangle, 40, 120, 820, 36)
    heroTitle.Name = "Hero_MainTitle"
    heroTitle.Fill.Visible = msoFalse
    heroTitle.Line.Visible = msoFalse
    With heroTitle.TextFrame2
        .WordWrap = msoFalse
        .MarginLeft = 0: .MarginTop = 0
        .TextRange.Text = "Executive BI & Advanced Analytics Intelligence Portal"
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 18
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
    
    ' Hero Subtitle / Mission Statement
    Set heroDesc = ws.Shapes.AddShape(msoShapeRectangle, 40, 156, 1000, 32)
    heroDesc.Name = "Hero_Subtitle"
    heroDesc.Fill.Visible = msoFalse
    heroDesc.Line.Visible = msoFalse
    With heroDesc.TextFrame2
        .WordWrap = msoTrue
        .MarginLeft = 0: .MarginTop = 0
        .TextRange.Text = "Unified strategic governance platform connecting Call Center Telephony Operations, Customer Retention Analytics, and Diversity & Inclusion Parity Models via Ralph Kimball Star-Schema Tabular Engine."
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 9
        .TextRange.Font.Fill.ForeColor.RGB = RGB(203, 213, 225)
    End With
End Sub

' ==============================================================================
' Component 3: Cross-Enterprise KPI Ticker (4 Cards)
' ==============================================================================
Private Sub CreateEnterpriseMetricsTicker(ByVal ws As Worksheet)
    Dim cardLefts As Variant
    Dim titles As Variant
    Dim vals As Variant
    Dim subtexts As Variant
    Dim accentColors As Variant
    Dim i As Long
    Dim card As Shape
    Dim txt As Shape
    Dim bar As Shape
    
    cardLefts = Array(20, 335, 650, 965)
    titles = Array("TELEPHONY SLA COMPLIANCE", "CUSTOMER ARR PRESERVATION", "EXECUTIVE GENDER PARITY", "DATA ARCHITECTURE FIDELITY")
    vals = Array("81.1%", "$2.86M", "41.0%", "100%")
    subtexts = Array("4,054 Answered / 5,000 Total Calls", "26.5% Churn Rate across 7,043 Accounts", "28.0% Senior Leadership Representation", "13 Star-Schema Models & Lookups")
    accentColors = Array(COLOR_EMERALD, COLOR_PWC_ORANGE, COLOR_ROSE, COLOR_DARK_NAVY)
    
    For i = 0 To 3
        ' Card Background
        Set card = ws.Shapes.AddShape(msoShapeRoundedRectangle, cardLefts(i), 212, 295, 78)
        card.Name = "Ticker_Card_" & (i + 1)
        card.Adjustments.Item(1) = 0.12
        card.Fill.Solid
        card.Fill.ForeColor.RGB = COLOR_CARD_BG
        card.Line.ForeColor.RGB = COLOR_BORDER
        card.Line.Weight = 1
        ApplySoftElevation card
        
        ' Accent Left Border Indicator
        Set bar = ws.Shapes.AddShape(msoShapeRoundedRectangle, cardLefts(i) + 4, 220, 4, 62)
        bar.Name = "Ticker_Accent_" & (i + 1)
        bar.Adjustments.Item(1) = 0.5
        bar.Fill.Solid
        bar.Fill.ForeColor.RGB = accentColors(i)
        bar.Line.Visible = msoFalse
        
        ' Text Frame
        Set txt = ws.Shapes.AddShape(msoShapeRectangle, cardLefts(i) + 16, 216, 270, 70)
        txt.Name = "Ticker_Text_" & (i + 1)
        txt.Fill.Visible = msoFalse
        txt.Line.Visible = msoFalse
        With txt.TextFrame2
            .WordWrap = msoTrue
            .MarginLeft = 0: .MarginTop = 0: .MarginRight = 0: .MarginBottom = 0
            .TextRange.Text = titles(i) & vbCrLf & vals(i) & vbCrLf & subtexts(i)
            
            ' Title line
            With .TextRange.Paragraphs(1).Font
                .Name = "Segoe UI"
                .Size = 7.5
                .Bold = msoTrue
                .Fill.ForeColor.RGB = COLOR_SLATE_TEXT
            End With
            
            ' Value line
            With .TextRange.Paragraphs(2).Font
                .Name = "Segoe UI"
                .Size = 16
                .Bold = msoTrue
                .Fill.ForeColor.RGB = COLOR_DARK_NAVY
            End With
            
            ' Subtext line
            With .TextRange.Paragraphs(3).Font
                .Name = "Segoe UI"
                .Size = 7.5
                .Bold = msoFalse
                .Fill.ForeColor.RGB = COLOR_SLATE_TEXT
            End With
        End With
    Next i
End Sub

' ==============================================================================
' Component 4: The 3 Large Cockpit Launcher Cards
' ==============================================================================
Private Sub CreateCockpitLauncherCards(ByVal ws As Worksheet)
    Dim cardLefts As Variant
    Dim tags As Variant
    Dim titles As Variant
    Dim subtitles As Variant
    Dim metrics As Variant
    Dim targets As Variant
    Dim btnLabels As Variant
    Dim tagColors As Variant
    Dim i As Long
    
    cardLefts = Array(20, 435, 850)
    tags = Array("01 TELEPHONY OPERATIONS", "02 CUSTOMER RETENTION", "03 DIVERSITY & INCLUSION")
    titles = Array("Call Center Operations & Agent Audit", "Customer Retention & Revenue Risk", "Diversity, Equity & Executive Parity")
    subtitles = Array( _
        "Real-time queue triage, intraday hourly volume distributions, first-contact resolution audit, and representative efficiency scorecards with live SLA compliance metrics.", _
        "Contract vulnerability profiling, early tenure decay curve analytics, payment friction diagnostics, and tech support bundling matrices to stem ARR revenue attrition.", _
        "Comprehensive workforce census, corporate ladder broken-rung discovery, promotion velocity differentials, and performance appraisal parity audits.")
    metrics = Array( _
        "81.1% Answer Rate  |  67.5s Avg Speed  |  3.40 CSAT", _
        "7,043 Accounts  |  26.5% Churn  |  $2.86M At-Risk ARR", _
        "500 Personnel  |  41% Female  |  10.2% Promo Velocity")
    targets = Array("03_CallCenter_Cockpit", "04_CustomerRetention_Cockpit", "05_DiversityInclusion_Cockpit")
    btnLabels = Array("Launch Call Center Cockpit ->", "Launch Retention Cockpit ->", "Launch D&I Cockpit ->")
    tagColors = Array(COLOR_EMERALD, COLOR_PWC_ORANGE, COLOR_ROSE)
    
    For i = 0 To 2
        ' Card Container
        Dim card As Shape
        Set card = ws.Shapes.AddShape(msoShapeRoundedRectangle, cardLefts(i), 304, 390, 360)
        card.Name = "Card_Launcher_" & (i + 1)
        card.Adjustments.Item(1) = 0.06
        card.Fill.Solid
        card.Fill.ForeColor.RGB = COLOR_CARD_BG
        card.Line.ForeColor.RGB = COLOR_BORDER
        card.Line.Weight = 1
        ApplySoftElevation card
        
        ' Tag Pill Badge
        Dim pill As Shape
        Set pill = ws.Shapes.AddShape(msoShapeRoundedRectangle, cardLefts(i) + 20, 320, 160, 22)
        pill.Name = "Pill_Launcher_" & (i + 1)
        pill.Adjustments.Item(1) = 0.5
        pill.Fill.Solid
        pill.Fill.ForeColor.RGB = RGB(241, 245, 249)
        pill.Line.ForeColor.RGB = COLOR_BORDER
        pill.Line.Weight = 0.75
        With pill.TextFrame2
            .TextRange.Text = tags(i)
            .TextRange.Font.Name = "Segoe UI"
            .TextRange.Font.Size = 7.5
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = tagColors(i)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
        
        ' Card Title
        Dim titleTxt As Shape
        Set titleTxt = ws.Shapes.AddShape(msoShapeRectangle, cardLefts(i) + 20, 350, 350, 44)
        titleTxt.Name = "Title_Launcher_" & (i + 1)
        titleTxt.Fill.Visible = msoFalse
        titleTxt.Line.Visible = msoFalse
        With titleTxt.TextFrame2
            .WordWrap = msoTrue
            .MarginLeft = 0: .MarginTop = 0
            .TextRange.Text = titles(i)
            .TextRange.Font.Name = "Segoe UI"
            .TextRange.Font.Size = 13.5
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = COLOR_DARK_NAVY
        End With
        
        ' Card Description
        Dim descTxt As Shape
        Set descTxt = ws.Shapes.AddShape(msoShapeRectangle, cardLefts(i) + 20, 396, 350, 95)
        descTxt.Name = "Desc_Launcher_" & (i + 1)
        descTxt.Fill.Visible = msoFalse
        descTxt.Line.Visible = msoFalse
        With descTxt.TextFrame2
            .WordWrap = msoTrue
            .MarginLeft = 0: .MarginTop = 0
            .TextRange.Text = subtitles(i)
            .TextRange.Font.Name = "Segoe UI"
            .TextRange.Font.Size = 8.5
            .TextRange.Font.Fill.ForeColor.RGB = COLOR_SLATE_TEXT
        End With
        
        ' Metric Banner Box inside Card
        Dim mBox As Shape
        Set mBox = ws.Shapes.AddShape(msoShapeRoundedRectangle, cardLefts(i) + 20, 500, 350, 40)
        mBox.Name = "MetricBox_Launcher_" & (i + 1)
        mBox.Adjustments.Item(1) = 0.2
        mBox.Fill.Solid
        mBox.Fill.ForeColor.RGB = RGB(248, 250, 252)
        mBox.Line.ForeColor.RGB = COLOR_BORDER
        mBox.Line.Weight = 0.75
        With mBox.TextFrame2
            .TextRange.Text = metrics(i)
            .TextRange.Font.Name = "Segoe UI"
            .TextRange.Font.Size = 8
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = COLOR_DARK_NAVY
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
        
        ' Launch Button (Interactive Hyperlink CTA)
        Dim btnCTA As Shape
        Set btnCTA = ws.Shapes.AddShape(msoShapeRoundedRectangle, cardLefts(i) + 20, 560, 350, 40)
        btnCTA.Name = "Btn_Launch_" & (i + 1)
        btnCTA.Adjustments.Item(1) = 0.25
        btnCTA.Fill.Solid
        btnCTA.Fill.ForeColor.RGB = COLOR_DARK_NAVY
        btnCTA.Line.Visible = msoFalse
        With btnCTA.TextFrame2
            .TextRange.Text = btnLabels(i)
            .TextRange.Font.Name = "Segoe UI"
            .TextRange.Font.Size = 9.5
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
        
        ' Assign Hyperlink to Target Cockpit Sheet
        ws.Hyperlinks.Add _
            Anchor:=btnCTA, _
            Address:="", _
            SubAddress:="'" & targets(i) & "'!A1", _
            ScreenTip:="Open " & titles(i)
    Next i
End Sub

' ==============================================================================
' Component 5: Bottom Governance Drawer
' ==============================================================================
Private Sub CreateGovernanceDrawer(ByVal ws As Worksheet)
    Dim bar As Shape
    Dim txt As Shape
    Dim btnGov1 As Shape
    Dim btnGov2 As Shape
    
    ' Bottom Container
    Set bar = ws.Shapes.AddShape(msoShapeRoundedRectangle, 20, 680, 1220, 60)
    bar.Name = "Gov_Drawer"
    bar.Adjustments.Item(1) = 0.15
    bar.Fill.Solid
    bar.Fill.ForeColor.RGB = COLOR_CARD_BG
    bar.Line.ForeColor.RGB = COLOR_BORDER
    bar.Line.Weight = 1
    ApplySoftElevation bar
    
    ' Text
    Set txt = ws.Shapes.AddShape(msoShapeRectangle, 40, 692, 600, 36)
    txt.Name = "Gov_Text"
    txt.Fill.Visible = msoFalse
    txt.Line.Visible = msoFalse
    With txt.TextFrame2
        .WordWrap = msoFalse
        .MarginLeft = 0: .MarginTop = 0
        .TextRange.Text = "Enterprise Governance & Data Dictionary Repositories" & vbCrLf & "Explore official entity grains, VertiPaq tabular schema definitions, and certified KPI formulation logic."
        .TextRange.Paragraphs(1).Font.Name = "Segoe UI"
        .TextRange.Paragraphs(1).Font.Size = 9.5
        .TextRange.Paragraphs(1).Font.Bold = msoTrue
        .TextRange.Paragraphs(1).Font.Fill.ForeColor.RGB = COLOR_DARK_NAVY
        .TextRange.Paragraphs(2).Font.Name = "Segoe UI"
        .TextRange.Paragraphs(2).Font.Size = 7.5
        .TextRange.Paragraphs(2).Font.Fill.ForeColor.RGB = COLOR_SLATE_TEXT
    End With
    
    ' Button 1: Business Domains Architecture
    Set btnGov1 = ws.Shapes.AddShape(msoShapeRoundedRectangle, 820, 694, 180, 32)
    btnGov1.Name = "Btn_Gov_Domains"
    btnGov1.Adjustments.Item(1) = 0.25
    btnGov1.Fill.Solid
    btnGov1.Fill.ForeColor.RGB = RGB(241, 245, 249)
    btnGov1.Line.ForeColor.RGB = COLOR_BORDER
    btnGov1.Line.Weight = 1
    With btnGov1.TextFrame2
        .TextRange.Text = "01 Business Domains"
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 8.5
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = COLOR_DARK_NAVY
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    ws.Hyperlinks.Add Anchor:=btnGov1, Address:="", SubAddress:="'01_Business_Domains'!A1", ScreenTip:="Open 01_Business_Domains"
    
    ' Button 2: Metadata & KPI Catalog
    Set btnGov2 = ws.Shapes.AddShape(msoShapeRoundedRectangle, 1020, 694, 200, 32)
    btnGov2.Name = "Btn_Gov_Catalog"
    btnGov2.Adjustments.Item(1) = 0.25
    btnGov2.Fill.Solid
    btnGov2.Fill.ForeColor.RGB = RGB(241, 245, 249)
    btnGov2.Line.ForeColor.RGB = COLOR_BORDER
    btnGov2.Line.Weight = 1
    With btnGov2.TextFrame2
        .TextRange.Text = "02 Metadata & KPI Catalog"
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 8.5
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = COLOR_DARK_NAVY
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    ws.Hyperlinks.Add Anchor:=btnGov2, Address:="", SubAddress:="'02_Metadata_&_KPI_Catalog'!A1", ScreenTip:="Open 02_Metadata_&_KPI_Catalog"
End Sub

' ==============================================================================
' Visual Elevation Helper (Soft Modern SaaS Drop Shadow)
' ==============================================================================
Private Sub ApplySoftElevation(ByVal shp As Shape)
    On Error Resume Next
    With shp.Shadow
        .Type = msoShadow21
        .Visible = msoTrue
        .ForeColor.RGB = COLOR_DARK_NAVY
        .Transparency = 0.90
        .Size = 100
        .Blur = 8
        .OffsetX = 0
        .OffsetY = 3
    End With
End Sub

' ==============================================================================
' Dynamic Animation: Hero Banner Entrance Glow & Pulse
' ==============================================================================
Public Sub AnimatePortalBanner()
    On Error Resume Next
    If Not Application.Visible Then Exit Sub
    Dim ws As Worksheet
    Dim heroBg As Shape
    Dim step As Long
    
    Set ws = ThisWorkbook.Worksheets(SHEET_PORTAL)
    If ws Is Nothing Then Exit Sub
    
    Set heroBg = ws.Shapes("Hero_BannerCard")
    If heroBg Is Nothing Then Exit Sub
    
    ' Pulse the gradient border to simulate a sleek web application load
    For step = 1 To 3
        heroBg.Line.ForeColor.RGB = RGB(255, 110, 38) ' Pulsing Tangerine Glow
        heroBg.Line.Weight = 2
        DoEvents
        
        ' Short pause
        Dim t As Double
        t = Timer
        Do While Timer < t + 0.08: DoEvents: Loop
        
        heroBg.Line.ForeColor.RGB = RGB(51, 65, 85) ' Rest back to Slate
        heroBg.Line.Weight = 1
        DoEvents
    Next step
End Sub
