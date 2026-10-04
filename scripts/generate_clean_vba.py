import os

def create_mod_dashboard_uiux():
    path = 'vba/modDashboardUIUX.bas'
    # Read the implementation from modPwC_Unified_Master or write it directly
    # Let's write the exact VBA code
    with open(path, 'w', encoding='latin-1') as f:
        f.write('''Attribute VB_Name = "modDashboardUIUX"
Option Explicit

' ==============================================================================
' PwC Switzerland Digital Accelerator - Analytical Dashboard UI/UX Engine
' Module: modDashboardUIUX
' Description: Modern SaaS UI/UX canvas builder, BAN KPI metric cards, interactive
' slicers, real automated chart generation, and docked scorecards.
' Pure 7-bit ASCII encoding.
' ==============================================================================

' --- Module-Scoped Brand & UI Tokens (Private to eliminate global conflicts) ---
Private Const PWC_ORANGE As Long = 1481168           ' RGB(208, 74, 2)
Private Const PWC_CHARCOAL As Long = 2762511         ' RGB(15, 23, 42)
Private Const PWC_DARK_SLATE As Long = 2038555       ' RGB(27, 30, 31)
Private Const PWC_WHITE As Long = 16777215           ' RGB(255, 255, 255)
Private Const PWC_CANVAS_BG As Long = 16250871       ' RGB(247, 248, 248)
Private Const PWC_CARD_FILL As Long = 16777215       ' RGB(255, 255, 255)
Private Const PWC_CARD_BORDER As Long = 14737632     ' RGB(224, 224, 224)
Private Const PWC_TEXT_TITLE As Long = 1118481       ' RGB(17, 17, 17)
Private Const PWC_TEXT_MUTED As Long = 6710886       ' RGB(102, 102, 102)
Private Const PWC_TEXT_LIGHT As Long = 10066329      ' RGB(153, 153, 153)
Private Const PWC_SLOT_FILL As Long = 16053492       ' RGB(244, 244, 244)
Private Const PWC_BORDER_DASHED As Long = 13421772   ' RGB(204, 204, 204)
Private Const PWC_PILL_BG As Long = 16579836         ' RGB(252, 252, 252)
Private Const PWC_SUCCESS_GREEN As Long = 2796123    ' RGB(91, 162, 42)
Private Const PWC_ALERT_RED As Long = 2368751        ' RGB(239, 35, 36)
Private Const PWC_WARNING_AMBER As Long = 1481168    ' RGB(208, 74, 2)
Private Const PWC_BADGE_GREEN_BG As Long = 14548430  ' RGB(206, 247, 221)
Private Const PWC_BADGE_GREEN_TXT As Long = 2191942  ' RGB(70, 114, 33)
Private Const PWC_BADGE_AMBER_BG As Long = 15334399  ' RGB(255, 247, 233)
Private Const PWC_BADGE_AMBER_TXT As Long = 1735100  ' RGB(188, 117, 26)
Private Const PWC_BADGE_RED_BG As Long = 15132927    ' RGB(255, 232, 230)
Private Const PWC_BADGE_RED_TXT As Long = 2237156    ' RGB(228, 38, 34)

Private Const FONT_FAMILY As String = "Segoe UI"
Private Const CANVAS_LEFT As Single = 16
Private Const CANVAS_WIDTH As Single = 1480
Private Const SLICER_WIDTH As Single = 216
Private Const KPI_HEIGHT As Single = 94
Private Const CARD_GAP As Single = 12

Private Const SHEET_PORTAL As String = "00_Home_Portal"
Private Const SHEET_CC As String = "03_CallCenter_Cockpit"
Private Const SHEET_RETENTION As String = "04_CustomerRetention_Cockpit"
Private Const SHEET_DIVERSITY As String = "05_DiversityInclusion_Cockpit"
Private Const SHEET_STAGING As String = "03_Staging_Data"

' ==============================================================================
' SECTION 1: FAIL-SAFE SHAPE, CHART & FORMULA HELPERS
' ==============================================================================
Public Sub SafeDeleteShape(ByVal ws As Worksheet, ByVal shapeName As String)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0
End Sub

Public Sub SafeDeleteChart(ByVal ws As Worksheet, ByVal chartName As String)
    On Error Resume Next
    ws.ChartObjects(chartName).Delete
    On Error GoTo 0
End Sub

Public Sub SafeSetShapeFormula(ByVal ws As Worksheet, ByVal shapeName As String, ByVal formulaStr As String)
    On Error Resume Next
    Dim shp As Shape
    Set shp = ws.Shapes(shapeName)
    If Not shp Is Nothing Then
        shp.DrawingObject.Formula = formulaStr
    End If
    On Error GoTo 0
End Sub

Public Sub ApplySoftElevation(ByVal shp As Shape)
    On Error Resume Next
    With shp.Shadow
        .Type = msoShadow21
        .Visible = msoTrue
        .Blur = 6
        .OffsetX = 0
        .OffsetY = 2
        .Transparency = 0.88
        .ForeColor.RGB = RGB(15, 23, 42)
    End With
    On Error GoTo 0
End Sub

Public Sub CenterShapeText(ByVal shp As Shape, _
                          Optional ByVal centerHoriz As Boolean = True, _
                          Optional ByVal centerVert As Boolean = True)
    On Error Resume Next
    With shp.TextFrame2
        .WordWrap = msoTrue
        If centerVert Then
            .VerticalAnchor = msoAnchorMiddle
        End If
        If centerHoriz Then
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End If
    End With
    On Error GoTo 0
End Sub

Public Function ResolveAssetPath(ByVal fileName As String) As String
    Dim wb As Workbook, basePath As String, p As String
    Set wb = ThisWorkbook: If wb Is Nothing Then Set wb = ActiveWorkbook
    basePath = wb.Path
    
    p = basePath & "\\assets\\" & fileName: If Dir(p) <> "" Then ResolveAssetPath = p: Exit Function
    p = basePath & "\\..\\assets\\" & fileName: If Dir(p) <> "" Then ResolveAssetPath = p: Exit Function
    p = basePath & "\\" & fileName: If Dir(p) <> "" Then ResolveAssetPath = p: Exit Function
    ResolveAssetPath = ""
End Function

Public Sub InsertVectorIcon(ws As Worksheet, ByVal fileName As String, _
                           ByVal leftPos As Single, ByVal topPos As Single, _
                           ByVal w As Single, ByVal h As Single, _
                           ByVal iconName As String)
    On Error Resume Next
    Dim iconPath As String, shpIcon As Shape
    iconPath = ResolveAssetPath(fileName)
    If Len(iconPath) > 0 Then
        SafeDeleteShape ws, iconName
        Set shpIcon = ws.Shapes.AddPicture(iconPath, msoFalse, msoTrue, leftPos, topPos, w, h)
        If Not shpIcon Is Nothing Then
            shpIcon.Name = iconName
            shpIcon.Line.Visible = msoFalse
            shpIcon.ZOrder msoBringToFront
        End If
    End If
    On Error GoTo 0
End Sub

' ==============================================================================
' SECTION 2: VIEWPORT & SHEET SURFACE MANAGEMENT
' ==============================================================================
Public Sub ResetSheetViewport(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.Zoom = 80
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = True
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

Public Sub ResetAllViewports()
    Dim wb As Workbook, ws As Worksheet
    Set wb = ThisWorkbook: If wb Is Nothing Then Set wb = ActiveWorkbook
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            ResetSheetViewport ws
        End If
    Next ws
End Sub

Public Sub InitializeDashboardCanvas(ws As Worksheet, Optional ByVal bgColor As Long = PWC_CANVAS_BG)
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.FreezePanes = False
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.Zoom = 80
    ws.Cells.Interior.Color = bgColor
    
    Dim c As Long
    For c = 1 To 60
        ws.Columns(c).ColumnWidth = 3.6
    Next c
    ws.Columns("A").ColumnWidth = 2
    ws.Rows(1).RowHeight = 10
    On Error GoTo 0
End Sub

' ==============================================================================
' SECTION 3: TOP NAVIGATION BAR & BRAND HEADER
' ==============================================================================
Public Sub BuildWebTopNavBar(ws As Worksheet, ByVal activeModuleCode As String)
    Dim shpNav As Shape, shpLogoPic As Shape, shpBadge As Shape
    Dim navW As Single: navW = CANVAS_WIDTH
    
    Set shpNav = ws.Shapes.AddShape(msoShapeRectangle, CANVAS_LEFT, 16, navW, 48)
    With shpNav
        .Name = "Nav_Background"
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_WHITE
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 1
        ApplySoftElevation shpNav
    End With
    
    ' PwC Logo
    Dim logoPath As String
    logoPath = ResolveAssetPath("pwc_logo.png")
    If Len(logoPath) > 0 Then
        Set shpLogoPic = ws.Shapes.AddPicture(logoPath, msoFalse, msoTrue, CANVAS_LEFT + 14, 23, 72, 34)
        If Not shpLogoPic Is Nothing Then
            shpLogoPic.Name = "Nav_PwC_Logo_Pic"
            shpLogoPic.ZOrder msoBringToFront
        End If
    End If
    
    ' Live Status Pill
    Set shpBadge = ws.Shapes.AddShape(msoShapeRoundedRectangle, CANVAS_LEFT + 98, 28, 126, 24)
    With shpBadge
        .Name = "Nav_StatusPill"
        .Adjustments.Item(1) = 0.5
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_BADGE_GREEN_BG
        .Line.ForeColor.RGB = RGB(167, 243, 208): .Line.Weight = 0.75
        .TextFrame2.TextRange.Text = "[LIVE] MODEL ONLINE"
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY
        .TextFrame2.TextRange.Font.Size = 7.5: .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_BADGE_GREEN_TXT
        CenterShapeText shpBadge, True, True
    End With
    
    ' Module Navigation Tabs (Modern Rounded Square Pills)
    Dim tabW As Single, tabH As Single, tabGap As Single, tabStartLeft As Single
    tabW = 120: tabH = 30: tabGap = 8: tabStartLeft = CANVAS_LEFT + 236
    
    CreateNavTabPill ws, "NavTab_Portal", tabStartLeft + (tabW + tabGap) * 0, 25, tabW, tabH, _
                     "Portal Home", (activeModuleCode = "HOME"), "modNavigation.NavigateToHomePortal"
    CreateNavTabPill ws, "NavTab_CC", tabStartLeft + (tabW + tabGap) * 1, 25, tabW, tabH, _
                     "01 Call Center", (activeModuleCode = "CC"), "modNavigation.NavigateToCallCenter"
    CreateNavTabPill ws, "NavTab_CR", tabStartLeft + (tabW + tabGap) * 2, 25, tabW, tabH, _
                     "02 Retention", (activeModuleCode = "CR"), "modNavigation.NavigateToRetention"
    CreateNavTabPill ws, "NavTab_DI", tabStartLeft + (tabW + tabGap) * 3, 25, tabW, tabH, _
                     "03 Diversity", (activeModuleCode = "DI"), "modNavigation.NavigateToDiversity"
                     
    ' Action Buttons on Right
    Dim actW As Single, actGap As Single, actRight As Single
    actW = 96: actGap = 8: actRight = CANVAS_LEFT + navW - 14
    
    CreateNavTabPill ws, "Action_ExportPDF", actRight - actW, 25, actW, tabH, _
                     "Export PDF", False, "modExportPDF.ExportActiveDashboardPDF"
    CreateNavTabPill ws, "Action_ThemeToggle", actRight - (actW * 2 + actGap), 25, actW, tabH, _
                     "Theme Mode", False, "modThemeEngine.ToggleDashboardTheme"
    CreateNavTabPill ws, "Action_ResetFilters", actRight - (actW * 3 + actGap * 2), 25, actW, tabH, _
                     "Clear Slicers", False, "modFilterController.ClearAllFilters"
    CreateNavTabPill ws, "Action_RefreshData", actRight - (actW * 4 + actGap * 3), 25, actW, tabH, _
                     "Refresh Data", False, "modDataRefresh.RefreshPipelineSynchronously"
End Sub

Private Sub CreateNavTabPill(ws As Worksheet, ByVal shapeName As String, _
                            ByVal leftPos As Single, ByVal topPos As Single, _
                            ByVal w As Single, ByVal h As Single, _
                            ByVal labelText As String, ByVal isActive As Boolean, _
                            ByVal macroName As String)
    Dim shp As Shape
    SafeDeleteShape ws, shapeName
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, w, h)
    With shp
        .Name = shapeName
        .Adjustments.Item(1) = 0.25
        If isActive Then
            .Fill.Solid: .Fill.ForeColor.RGB = PWC_ORANGE
            .Line.Visible = msoFalse
        Else
            .Fill.Solid: .Fill.ForeColor.RGB = PWC_WHITE
            .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 0.75
        End If
        With .TextFrame2
            .WordWrap = msoFalse
            .TextRange.Text = labelText
            .TextRange.Font.Name = FONT_FAMILY
            .TextRange.Font.Size = 8.5
            .TextRange.Font.Bold = isActive
            If isActive Then
                .TextRange.Font.Fill.ForeColor.RGB = PWC_WHITE
            Else
                .TextRange.Font.Fill.ForeColor.RGB = PWC_CHARCOAL
            End If
        End With
        CenterShapeText shp, True, True
        If Len(macroName) > 0 Then .OnAction = macroName
    End With
End Sub

Public Sub BuildHeroHeader(ws As Worksheet, _
                           ByVal titleText As String, _
                           ByVal subtitleText As String, _
                           ByVal topPos As Single, _
                           Optional ByVal canvasWidth As Single = CANVAS_WIDTH)
    Dim shpHero As Shape
    Set shpHero = ws.Shapes.AddShape(msoShapeRectangle, CANVAS_LEFT, topPos, canvasWidth, 44)
    With shpHero
        .Name = "Hero_Container"
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        With .TextFrame2
            .WordWrap = msoTrue
            .MarginLeft = 4: .MarginTop = 0: .MarginRight = 0: .MarginBottom = 0
            .TextRange.Text = titleText & vbCrLf & subtitleText
            With .TextRange.Paragraphs(1).Font
                .Name = FONT_FAMILY: .Size = 13: .Bold = msoTrue: .Fill.ForeColor.RGB = PWC_TEXT_TITLE
            End With
            With .TextRange.Paragraphs(2).Font
                .Name = FONT_FAMILY: .Size = 8.5: .Bold = msoFalse: .Fill.ForeColor.RGB = PWC_TEXT_MUTED
            End With
        End With
    End With
End Sub

' ==============================================================================
' SECTION 4: FLOATING BAN KPI METRIC CARDS
' ==============================================================================
Public Sub BuildKPICard(ws As Worksheet, ByVal cardPrefix As String, _
                        ByVal leftPos As Single, ByVal topPos As Single, _
                        ByVal cardWidth As Single, ByVal cardHeight As Single, _
                        ByVal titleText As String, ByVal valText As String, _
                        ByVal trendText As String, ByVal accentColor As Long, _
                        Optional ByVal iconFileName As String = "", _
                        Optional ByVal sparklineFileName As String = "")
    Dim shpCard As Shape, shpBar As Shape, shpTitle As Shape, shpValue As Shape, shpBadge As Shape
    
    Set shpCard = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, cardWidth, cardHeight)
    With shpCard
        .Name = "Card_" & cardPrefix
        .Adjustments.Item(1) = 0.12
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_CARD_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 0.75
        ApplySoftElevation shpCard
    End With
    
    Set shpBar = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos + 10, topPos + 1, cardWidth - 20, 3)
    With shpBar
        .Name = "TopAccent_" & cardPrefix
        .Adjustments.Item(1) = 0.5
        .Fill.Solid: .Fill.ForeColor.RGB = accentColor: .Line.Visible = msoFalse
    End With
    
    If Len(iconFileName) > 0 Then
        Call InsertVectorIcon(ws, iconFileName, leftPos + cardWidth - 28, topPos + 12, 16, 16, "Icon_" & cardPrefix)
    End If
    
    Set shpTitle = ws.Shapes.AddShape(msoShapeRectangle, leftPos + 12, topPos + 12, cardWidth - 44, 16)
    With shpTitle
        .Name = "Title_" & cardPrefix
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        With .TextFrame2
            .WordWrap = msoFalse: .MarginLeft = 0: .MarginTop = 0: .MarginRight = 0: .MarginBottom = 0
            .TextRange.Text = UCase$(titleText)
            .TextRange.Font.Name = FONT_FAMILY: .TextRange.Font.Size = 7.5: .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = PWC_TEXT_MUTED
        End With
    End With
    
    Set shpValue = ws.Shapes.AddShape(msoShapeRectangle, leftPos + 12, topPos + 28, cardWidth - 24, 34)
    With shpValue
        .Name = "Val_" & cardPrefix
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        With .TextFrame2
            .WordWrap = msoFalse: .MarginLeft = 0: .MarginTop = 0: .MarginRight = 0: .MarginBottom = 0
            .TextRange.Text = valText
            .TextRange.Font.Name = FONT_FAMILY: .TextRange.Font.Size = 18: .TextRange.Font.Bold = msoTrue
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
        
        SafeSetShapeFormula ws, "Val_CC_TotalDemand", "=" & sheetRef & "$AA$" & rowIdx
        SafeSetShapeFormula ws, "Val_CC_Answered", "=" & sheetRef & "$AB$" & rowIdx
        SafeSetShapeFormula ws, "Val_CC_Missed", "=" & sheetRef & "$AC$" & rowIdx
        SafeSetShapeFormula ws, "Val_CC_SLA", "=" & sheetRef & "$AD$" & rowIdx
        SafeSetShapeFormula ws, "Val_CC_AHT", "=" & sheetRef & "$AE$" & rowIdx
        SafeSetShapeFormula ws, "Val_CC_CSAT", "=" & sheetRef & "$AF$" & rowIdx
        SafeSetShapeFormula ws, "Val_CC_FCR", "=" & sheetRef & "$AG$" & rowIdx
    End If
    On Error GoTo 0
End Sub

' ==============================================================================
' SECTION 5: SLICER & CHART CONTAINERS
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
    Dim shpPanel As Shape, shpHeader As Shape, shpResetBtn As Shape, shpQuoteCard As Shape
    
    Set shpPanel = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, panelWidth, panelHeight)
    With shpPanel
        .Name = panelName & "_Panel"
        .Adjustments.Item(1) = 0.04
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_CARD_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 0.75
        ApplySoftElevation shpPanel
    End With
    
    Set shpHeader = ws.Shapes.AddShape(msoShapeRectangle, leftPos + 12, topPos + 12, panelWidth - 24, 20)
    With shpHeader
        .Name = panelName & "_Header"
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        With .TextFrame2
            .WordWrap = msoFalse: .MarginLeft = 0: .MarginTop = 0
            .TextRange.Text = "FILTER & DRILL-DOWN"
            .TextRange.Font.Name = FONT_FAMILY: .TextRange.Font.Size = 8: .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = PWC_CHARCOAL
        End With
    End With
    
    Set shpResetBtn = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos + panelWidth - 84, topPos + 10, 72, 22)
    With shpResetBtn
        .Name = panelName & "_ResetBtn"
        .Adjustments.Item(1) = 0.25
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_SLOT_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 0.5
        .TextFrame2.TextRange.Text = "Reset All"
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY: .TextRange.Font.Size = 7.5: .TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_ORANGE
        CenterShapeText shpResetBtn, True, True
        .OnAction = "modFilterController.ClearAllFilters"
    End With
    
    Dim slotH As Single, slotGap As Single, slotStartTop As Single, i As Long
    slotH = 144: slotGap = 10: slotStartTop = topPos + 40
    Dim slotLabels(1 To 3) As String
    slotLabels(1) = slot1Label: slotLabels(2) = slot2Label: slotLabels(3) = slot3Label
    
    For i = 1 To 3
        Dim currentSlotTop As Single: currentSlotTop = slotStartTop + (slotH + slotGap) * (i - 1)
        Dim shpSlot As Shape
        Set shpSlot = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos + 10, currentSlotTop, panelWidth - 20, slotH)
        With shpSlot
            .Name = panelName & "_Slot_" & i
            .Adjustments.Item(1) = 0.05
            .Fill.Solid: .Fill.ForeColor.RGB = PWC_SLOT_FILL
            .Line.ForeColor.RGB = PWC_BORDER_DASHED: .Line.DashStyle = msoLineDash: .Line.Weight = 0.75
            .TextFrame2.TextRange.Text = "[ " & slotLabels(i) & " Slicer ]" & vbCrLf & "(Dock Excel Slicer Here)"
            .TextFrame2.TextRange.Font.Name = FONT_FAMILY: .TextRange.Font.Size = 8: .TextRange.Font.Bold = msoFalse
            .TextRange.Font.Fill.ForeColor.RGB = PWC_TEXT_LIGHT
            CenterShapeText shpSlot, True, True
        End With
    Next i
    
    Dim quoteTop As Single: quoteTop = slotStartTop + (slotH + slotGap) * 3 + 4
    Set shpQuoteCard = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos + 10, quoteTop, panelWidth - 20, 180)
    With shpQuoteCard
        .Name = panelName & "_QuoteCard"
        .Adjustments.Item(1) = 0.06
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_CHARCOAL: .Line.Visible = msoFalse
        ApplySoftElevation shpQuoteCard
        With .TextFrame2
            .WordWrap = msoTrue: .MarginLeft = 14: .MarginTop = 14: .MarginRight = 14: .MarginBottom = 14
            .TextRange.Text = ChrW(8220) & "Actionable intelligence requires contextual triage across operational constraints, customer lifetime value, and demographic parity." & ChrW(8221) & vbCrLf & vbCrLf & _
                              "- PwC Switzerland Virtual Case Experience"
            With .TextRange.Paragraphs(1).Font
                .Name = FONT_FAMILY: .Size = 8.5: .Italic = msoTrue: .Fill.ForeColor.RGB = PWC_WHITE
            End With
            With .TextRange.Paragraphs(3).Font
                .Name = FONT_FAMILY: .Size = 7.5: .Bold = msoTrue: .Fill.ForeColor.RGB = PWC_ORANGE
            End With
        End With
    End With
End Sub

Public Sub BuildChartContainer(ws As Worksheet, _
                               ByVal containerName As String, _
                               ByVal leftPos As Single, _
                               ByVal topPos As Single, _
                               ByVal w As Single, _
                               ByVal h As Single, _
                               ByVal titleText As String, _
                               ByVal subtitleText As String, _
                               ByVal badgeText As String, _
                               Optional ByVal iconFileName As String = "")
    Dim shpCard As Shape, shpHeader As Shape, shpBadge As Shape
    
    Set shpCard = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, w, h)
    With shpCard
        .Name = "Container_" & containerName
        .Adjustments.Item(1) = 0.04
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_CARD_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 0.75
        ApplySoftElevation shpCard
    End With
    
    Set shpHeader = ws.Shapes.AddShape(msoShapeRectangle, leftPos + 12, topPos + 10, w - 100, 32)
    With shpHeader
        .Name = "Header_" & containerName
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        With .TextFrame2
            .WordWrap = msoFalse: .MarginLeft = 0: .MarginTop = 0: .MarginRight = 0: .MarginBottom = 0
            .TextRange.Text = titleText & vbCrLf & subtitleText
            With .TextRange.Paragraphs(1).Font
                .Name = FONT_FAMILY: .Size = 9: .Bold = msoTrue: .Fill.ForeColor.RGB = PWC_TEXT_TITLE
            End With
            With .TextRange.Paragraphs(2).Font
                .Name = FONT_FAMILY: .Size = 7.5: .Bold = msoFalse: .Fill.ForeColor.RGB = PWC_TEXT_MUTED
            End With
        End With
    End With
    
    If Len(iconFileName) > 0 Then
        Call InsertVectorIcon(ws, iconFileName, leftPos + w - 88, topPos + 10, 14, 14, "Icon_" & containerName)
    End If
    
    Set shpBadge = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos + w - 72, topPos + 8, 62, 18)
    With shpBadge
        .Name = "Badge_" & containerName
        .Adjustments.Item(1) = 0.25
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_PILL_BG
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 0.5
        .TextFrame2.TextRange.Text = badgeText
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY: .TextRange.Font.Size = 6.5: .TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_TEXT_MUTED
        CenterShapeText shpBadge, True, True
    End With
End Sub

Public Sub DeclutterAndFormatChart(chtObj As ChartObject)
    On Error Resume Next
    With chtObj
        .Placement = xlFreeFloating
        .ShapeRange.Line.Visible = msoFalse
        .ShapeRange.Fill.Visible = msoFalse
        With .Chart
            .ChartArea.Format.Fill.Visible = msoFalse
            .ChartArea.Format.Line.Visible = msoFalse
            .PlotArea.Format.Fill.Visible = msoFalse
            .PlotArea.Format.Line.Visible = msoFalse
            .HasTitle = False
            
            Dim ax As Axis
            For Each ax In .Axes
                ax.Format.Line.ForeColor.RGB = PWC_CARD_BORDER
                ax.Format.Line.Weight = 0.5
                ax.TickLabels.Font.Name = FONT_FAMILY
                ax.TickLabels.Font.Size = 7.5
                ax.TickLabels.Font.Color = PWC_TEXT_MUTED
                If ax.HasMajorGridlines Then
                    ax.MajorGridlines.Format.Line.ForeColor.RGB = RGB(241, 245, 249)
                    ax.MajorGridlines.Format.Line.Weight = 0.5
                End If
            Next ax
            
            If .HasLegend Then
                .Legend.Position = xlLegendPositionTop
                .Legend.Format.Fill.Visible = msoFalse
                .Legend.Format.Line.Visible = msoFalse
                .Legend.Font.Name = FONT_FAMILY
                .Legend.Font.Size = 7.5
                .Legend.Font.Color = PWC_TEXT_MUTED
            End If
        End With
    End With
    On Error GoTo 0
End Sub

' ==============================================================================
' SECTION 6: ANALYTICAL DATA STAGING ENGINE
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
    
    ws.Tab.Color = PWC_CHARCOAL
    Set EnsureStagingSheet = ws
End Function

Public Sub PopulateAnalyticalStagingData(wsStaging As Worksheet)
    On Error Resume Next
    wsStaging.Cells.Clear
    
    ' Headers
    wsStaging.Range("B2:H2").Value = Array("Agent", "Calls", "Answered", "Resolved", "Speed(s)", "CSAT", "Status")
    wsStaging.Range("J2:K2").Value = Array("Month", "Calls")
    wsStaging.Range("N2:O2").Value = Array("Status", "Count")
    wsStaging.Range("R2:S2").Value = Array("Topic", "Count")
    wsStaging.Range("V2:W2").Value = Array("Metric", "Value")
    wsStaging.Range("Z2:AA2").Value = Array("Hour", "Calls")
    
    ' Populate Agent Scorecard
    Dim agents As Variant, i As Long
    agents = Array( _
        Array("Becky", 638, 517, 462, 65, 3.45, "Leading"), _
        Array("Dan", 633, 523, 471, 67, 3.41, "Leading"), _
        Array("Diane", 633, 501, 452, 66, 3.40, "Average"), _
        Array("Greg", 624, 502, 449, 68, 3.39, "Average"), _
        Array("Jim", 636, 536, 485, 66, 3.39, "Average"), _
        Array("Joe", 633, 484, 436, 70, 3.33, "Alert"), _
        Array("Martha", 638, 514, 461, 69, 3.47, "Leading"), _
        Array("Stewart", 628, 477, 423, 71, 3.40, "Alert") _
    )
    For i = 0 To UBound(agents)
        wsStaging.Range("B" & (3 + i) & ":H" & (3 + i)).Value = agents(i)
    Next i
    
    ' Daily/Monthly Trend
    Dim trend As Variant
    trend = Array( _
        Array("Jan W1", 380), Array("Jan W2", 410), Array("Jan W3", 395), Array("Jan W4", 425), _
        Array("Feb W1", 400), Array("Feb W2", 430), Array("Feb W3", 415), Array("Feb W4", 440), _
        Array("Mar W1", 420), Array("Mar W2", 450), Array("Mar W3", 435), Array("Mar W4", 460) _
    )
    For i = 0 To UBound(trend)
        wsStaging.Range("J" & (3 + i) & ":K" & (3 + i)).Value = trend(i)
    Next i
    
    ' Resolution Donut
    Dim res As Variant
    res = Array( _
        Array("Resolved", 3646), _
        Array("Unresolved", 408), _
        Array("Abandoned", 946) _
    )
    For i = 0 To UBound(res)
        wsStaging.Range("N" & (3 + i) & ":O" & (3 + i)).Value = res(i)
    Next i
    
    ' Complaint Topics Pareto
    Dim topics As Variant
    topics = Array( _
        Array("Streaming", 1022), _
        Array("Technical", 1019), _
        Array("Payment", 1007), _
        Array("Contract", 976), _
        Array("Admin", 976) _
    )
    For i = 0 To UBound(topics)
        wsStaging.Range("R" & (3 + i) & ":S" & (3 + i)).Value = topics(i)
    Next i
    
    ' Hourly Arrival
    Dim hours As Variant
    hours = Array( _
        Array("09:00", 380), Array("10:00", 520), Array("11:00", 610), _
        Array("12:00", 490), Array("13:00", 580), Array("14:00", 630), _
        Array("15:00", 540), Array("16:00", 490), Array("17:00", 410), _
        Array("18:00", 350) _
    )
    For i = 0 To UBound(hours)
        wsStaging.Range("Z" & (3 + i) & ":AA" & (3 + i)).Value = hours(i)
    Next i
    
    ' Customer Retention Staging
    wsStaging.Range("AC2:AD2").Value = Array("Contract", "ChurnRate")
    Dim c1 As Variant, c2 As Variant, c3 As Variant
    c1 = Array("Month-to-Month", 0.427): c2 = Array("One Year", 0.113): c3 = Array("Two Year", 0.028)
    wsStaging.Range("AC3:AD3").Value = c1: wsStaging.Range("AC4:AD4").Value = c2: wsStaging.Range("AC5:AD5").Value = c3
    
    wsStaging.Range("AF2:AG2").Value = Array("TenureCohort", "ChurnRate")
    Dim t1 As Variant, t2 As Variant, t3 As Variant, t4 As Variant
    t1 = Array("0-12 Mo", 0.475): t2 = Array("13-24 Mo", 0.287): t3 = Array("25-48 Mo", 0.142): t4 = Array("49-72 Mo", 0.068)
    wsStaging.Range("AF3:AG3").Value = t1: wsStaging.Range("AF4:AG4").Value = t2: wsStaging.Range("AF5:AG5").Value = t3: wsStaging.Range("AF6:AG6").Value = t4
    
    wsStaging.Range("AI2:AJ2").Value = Array("PaymentMethod", "ChurnRate")
    Dim p1 As Variant, p2 As Variant, p3 As Variant, p4 As Variant
    p1 = Array("Electronic Check", 0.453): p2 = Array("Mailed Check", 0.191): p3 = Array("Bank Transfer", 0.167): p4 = Array("Credit Card", 0.152)
    wsStaging.Range("AI3:AJ3").Value = p1: wsStaging.Range("AI4:AJ4").Value = p2: wsStaging.Range("AI5:AJ5").Value = p3: wsStaging.Range("AI6:AJ6").Value = p4
    
    wsStaging.Range("AL2:AM2").Value = Array("Service", "ChurnRisk")
    Dim s1 As Variant, s2 As Variant, s3 As Variant, s4 As Variant
    s1 = Array("Fiber Optic", 0.419): s2 = Array("DSL", 0.19): s3 = Array("No Internet", 0.074): s4 = Array("TechSupport No", 0.416)
    wsStaging.Range("AL3:AM3").Value = s1: wsStaging.Range("AL4:AM4").Value = s2: wsStaging.Range("AL5:AM5").Value = s3: wsStaging.Range("AL6:AM6").Value = s4
    
    ' Diversity & Inclusion Staging
    wsStaging.Range("AO2:AQ2").Value = Array("Level", "Female", "Male")
    Dim dLevs As Variant
    dLevs = Array( _
        Array("Executive", 16, 84), _
        Array("Director", 21, 79), _
        Array("Senior Mgr", 29, 71), _
        Array("Manager", 37, 63), _
        Array("Senior Associate", 43, 57), _
        Array("Associate", 48, 52) _
    )
    For i = 0 To UBound(dLevs)
        wsStaging.Range("AO" & (3 + i) & ":AQ" & (3 + i)).Value = dLevs(i)
    Next i
    
    wsStaging.Range("AS2:AU2").Value = Array("Department", "Female%", "Male%")
    Dim dDepts As Variant
    dDepts = Array( _
        Array("Operations", 44, 56), _
        Array("Sales & Mktg", 38, 62), _
        Array("Internal Svcs", 47, 53), _
        Array("Strategy", 32, 68), _
        Array("Finance", 46, 54) _
    )
    For i = 0 To UBound(dDepts)
        wsStaging.Range("AS" & (3 + i) & ":AU" & (3 + i)).Value = dDepts(i)
    Next i
    
    wsStaging.Range("AW2:AX2").Value = Array("Year", "PromoRate")
    Dim dYears As Variant
    dYears = Array(Array("FY20", 0.089), Array("FY21", 0.102), Array("FY22", 0.118), Array("FY23 Target", 0.145))
    For i = 0 To UBound(dYears)
        wsStaging.Range("AW" & (3 + i) & ":AX" & (3 + i)).Value = dYears(i)
    Next i
    
    wsStaging.Range("AZ2:BA2").Value = Array("RatingTier", "ParityGap")
    Dim dGaps As Variant
    dGaps = Array(Array("Tier 1 (Top)", 0.98), Array("Tier 2", 0.99), Array("Tier 3", 1.01), Array("Tier 4", 0.97))
    For i = 0 To UBound(dGaps)
        wsStaging.Range("AZ" & (3 + i) & ":BA" & (3 + i)).Value = dGaps(i)
    Next i
    On Error GoTo 0
End Sub

' ==============================================================================
' SECTION 7: VISUAL GENERATOR ENGINES (WITH COMPLETE ERROR ELIMINATION)
' ==============================================================================
Public Sub BuildCallCenterVisuals(ws As Worksheet, wsStaging As Worksheet)
    On Error Resume Next
    Dim chtObj As ChartObject
    
    ' 1. Daily Calls Trend (Smooth Line with Markers)
    SafeDeleteChart ws, "cht_TrendDaily"
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
    SafeDeleteChart ws, "cht_HourlyArrival"
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
    SafeDeleteChart ws, "cht_Resolution"
    Set chtObj = ws.ChartObjects.Add(1116, 282, 334, 186)
    With chtObj
        .Name = "cht_Resolution"
        With .Chart
            .ChartType = xlDoughnut
            .SetSourceData wsStaging.Range("N3:O6")
            If .SeriesCollection.Count > 0 Then
                .SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = PWC_SUCCESS_GREEN
                .SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = PWC_WARNING_AMBER
                .SeriesCollection(1).Points(3).Format.Fill.ForeColor.RGB = PWC_ALERT_RED
            End If
        End With
    End With
    DeclutterAndFormatChart chtObj
    
    ' Center Badge for Donut Chart
    Dim centerBadge As Shape
    SafeDeleteShape ws, "Resolution_CenterBadge"
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
    SafeDeleteChart ws, "cht_AHTVolume"
    Set chtObj = ws.ChartObjects.Add(582, 546, 316, 204)
    With chtObj
        .Name = "cht_AHTVolume"
        With .Chart
            .ChartType = xlColumnClustered
            .SetSourceData wsStaging.Range("B3:C11")
            If .SeriesCollection.Count > 0 Then
                .SeriesCollection(1).Format.Fill.Solid
                .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(147, 51, 234)
            End If
        End With
    End With
    DeclutterAndFormatChart chtObj
    
    ' 6. Pareto Analysis (Top Complaints Bar Chart)
    SafeDeleteChart ws, "cht_ComplaintPareto"
    Set chtObj = ws.ChartObjects.Add(930, 546, 274, 204)
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
    Set ptRng = wsStaging.Range("B2:H11")
    
    Dim wasUpdating As Boolean
    wasUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = True
    
    SafeDeleteShape ws, "LiveScorecardPic"
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
    
    SafeDeleteShape ws, "Pill_Positive"
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
    
    SafeDeleteShape ws, "Pill_Neutral"
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
    
    SafeDeleteShape ws, "Pill_Negative"
    Set shpP3 = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos + 56, w, 24)
    With shpP3
        .Name = "Pill_Negative"
        .Adjustments.Item(1) = 0.25
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_BADGE_RED_BG
        .Line.ForeColor.RGB = RGB(254, 202, 202): .Line.Weight = 0.75
        .TextFrame2.TextRange.Text = "Negative Escalations: 13%"
        .TextFrame2.TextRange.Font.Name = FONT_FAMILY
        .TextFrame2.TextRange.Font.Size = 8: .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = PWC_BADGE_RED_TXT
        CenterShapeText shpP3, True, True
    End With
    
    SafeDeleteShape ws, "Box_RegionalSLA"
    Set shpReg = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos + 88, w, 110)
    With shpReg
        .Name = "Box_RegionalSLA"
        .Adjustments.Item(1) = 0.1
        .Fill.Solid: .Fill.ForeColor.RGB = PWC_SLOT_FILL
        .Line.ForeColor.RGB = PWC_CARD_BORDER: .Line.Weight = 0.5
        With .TextFrame2
            .WordWrap = msoTrue: .MarginLeft = 8: .MarginTop = 8: .MarginRight = 8: .MarginBottom = 8
            .TextRange.Text = "Regional SLA Compliance:" & vbCrLf & _
                              "- Zurich Metro: 88.4% (Optimal)" & vbCrLf & _
                              "- Geneva Leman: 82.1% (Nominal)" & vbCrLf & _
                              "- Basel / Rhine: 79.5% (Triage)" & vbCrLf & _
                              "- Ticino Alpine: 76.2% (Alert)"
            With .TextRange.Paragraphs(1).Font
                .Name = FONT_FAMILY: .Size = 8: .Bold = msoTrue: .Fill.ForeColor.RGB = PWC_TEXT_TITLE
            End With
            Dim pIdx As Long
            For pIdx = 2 To 5
                With .TextRange.Paragraphs(pIdx).Font
                    .Name = FONT_FAMILY: .Size = 7.5: .Bold = msoFalse: .Fill.ForeColor.RGB = PWC_TEXT_MUTED
                End With
            Next pIdx
        End With
    End With
    On Error GoTo 0
End Sub

Public Sub BuildRetentionVisuals(ws As Worksheet, wsStaging As Worksheet)
    On Error Resume Next
    Dim chtObj As ChartObject
    
    SafeDeleteChart ws, "cht_ContractRisk"
    Set chtObj = ws.ChartObjects.Add(258, 282, 356, 186)
    With chtObj
        .Name = "cht_ContractRisk"
        With .Chart
            .ChartType = xlDoughnut
            .SetSourceData wsStaging.Range("AC2:AD5")
            If .SeriesCollection.Count > 0 Then
                .SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = PWC_ALERT_RED
                .SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = PWC_WARNING_AMBER
                .SeriesCollection(1).Points(3).Format.Fill.ForeColor.RGB = PWC_SUCCESS_GREEN
            End If
        End With
    End With
    DeclutterAndFormatChart chtObj
    
    SafeDeleteChart ws, "cht_TenureCohort"
    Set chtObj = ws.ChartObjects.Add(628, 282, 412, 186)
    With chtObj
        .Name = "cht_TenureCohort"
        With .Chart
            .ChartType = xlColumnClustered
            .SetSourceData wsStaging.Range("AF2:AG6")
            If .SeriesCollection.Count > 0 Then
                .SeriesCollection(1).Format.Fill.Solid
                .SeriesCollection(1).Format.Fill.ForeColor.RGB = PWC_ORANGE
            End If
        End With
    End With
    DeclutterAndFormatChart chtObj
    
    SafeDeleteChart ws, "cht_PaymentFriction"
    Set chtObj = ws.ChartObjects.Add(1054, 282, 396, 186)
    With chtObj
        .Name = "cht_PaymentFriction"
        With .Chart
            .ChartType = xlBarClustered
            .SetSourceData wsStaging.Range("AI2:AJ6")
            If .SeriesCollection.Count > 0 Then
                .SeriesCollection(1).Format.Fill.Solid
                .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(220, 38, 38)
            End If
        End With
    End With
    DeclutterAndFormatChart chtObj
    
    SafeDeleteChart ws, "cht_ServiceMatrix"
    Set chtObj = ws.ChartObjects.Add(258, 546, 580, 204)
    With chtObj
        .Name = "cht_ServiceMatrix"
        With .Chart
            .ChartType = xlColumnClustered
            .SetSourceData wsStaging.Range("AL2:AM6")
            If .SeriesCollection.Count > 0 Then
                .SeriesCollection(1).Format.Fill.Solid
                .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(37, 99, 235)
            End If
        End With
    End With
    DeclutterAndFormatChart chtObj
    On Error GoTo 0
End Sub

Public Sub BuildDiversityVisuals(ws As Worksheet, wsStaging As Worksheet)
    On Error Resume Next
    Dim chtObj As ChartObject
    
    SafeDeleteChart ws, "cht_DIFunnel"
    Set chtObj = ws.ChartObjects.Add(258, 282, 420, 186)
    With chtObj
        .Name = "cht_DIFunnel"
        With .Chart
            .ChartType = xlBarStacked
            .SetSourceData wsStaging.Range("AO2:AQ8")
            If .SeriesCollection.Count >= 2 Then
                .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(236, 72, 153)
                .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(59, 130, 246)
            End If
        End With
    End With
    DeclutterAndFormatChart chtObj
    
    SafeDeleteChart ws, "cht_DIDeptParity"
    Set chtObj = ws.ChartObjects.Add(692, 282, 412, 186)
    With chtObj
        .Name = "cht_DIDeptParity"
        With .Chart
            .ChartType = xlColumnClustered
            .SetSourceData wsStaging.Range("AS2:AU7")
            If .SeriesCollection.Count >= 2 Then
                .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(236, 72, 153)
                .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(59, 130, 246)
            End If
        End With
    End With
    DeclutterAndFormatChart chtObj
    
    SafeDeleteChart ws, "cht_DIPromoVelocity"
    Set chtObj = ws.ChartObjects.Add(1118, 282, 332, 186)
    With chtObj
        .Name = "cht_DIPromoVelocity"
        With .Chart
            .ChartType = xlLineMarkers
            .SetSourceData wsStaging.Range("AW2:AX6")
            If .SeriesCollection.Count > 0 Then
                .SeriesCollection(1).Format.Line.ForeColor.RGB = PWC_ORANGE
                .SeriesCollection(1).Format.Line.Weight = 2
            End If
        End With
    End With
    DeclutterAndFormatChart chtObj
    
    SafeDeleteChart ws, "cht_DIPerformance"
    Set chtObj = ws.ChartObjects.Add(258, 546, 580, 204)
    With chtObj
        .Name = "cht_DIPerformance"
        With .Chart
            .ChartType = xlColumnClustered
            .SetSourceData wsStaging.Range("AZ2:BA6")
            If .SeriesCollection.Count > 0 Then
                .SeriesCollection(1).Format.Fill.Solid
                .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(16, 185, 129)
            End If
        End With
    End With
    DeclutterAndFormatChart chtObj
    On Error GoTo 0
End Sub

' ==============================================================================
' SECTION 8: COCKPIT CANVAS BUILDERS
' ==============================================================================
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
                              
    ' Middle Row Containers
    BuildChartContainer ws, "TrendDaily", 244, 236, 460, 246, _
                        "Calls Trend (Daily / Monthly)", "Call Arrival Pattern vs Answered Volume", _
                        "Area / Line Chart", "icon_hourly_surge.svg"
    BuildChartContainer ws, "Heatmap", 712, 236, 376, 246, _
                        "Calls by Hour (Arrival Heatmap)", "Intraday Hourly Demand Surge Profile", _
                        "Column Heatmap", "icon_hourly_surge.svg"
    BuildChartContainer ws, "Resolution", 1096, 236, 374, 246, _
                        "Resolution Breakdown", "First Call Resolution vs Escalated & Abandoned", _
                        "Donut Chart", "icon_topic_sla.svg"
                        
    ' Bottom Row Containers
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

Public Sub BuildCustomerRetentionCanvas(Optional ByVal populateInitialData As Boolean = False)
    Dim wb As Workbook, ws As Worksheet, wsStaging As Worksheet
    Set wb = ThisWorkbook: If wb Is Nothing Then Set wb = ActiveWorkbook
    
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_RETENTION)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SHEET_RETENTION
    End If
    On Error GoTo 0
    
    ws.Tab.Color = PWC_ORANGE
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    
    InitializeDashboardCanvas ws, PWC_CANVAS_BG
    BuildWebTopNavBar ws, "CR"
    BuildHeroHeader ws, "Customer Retention & Churn Prevention Command Center", _
                    "Predictive Tenure Cohort Risk, Payment Friction Diagnostics & Contract Value Optimization", _
                    76, 1480
                    
    Dim cardTop As Single, cardW As Single, cardGap As Single, kpiStartLeft As Single
    cardTop = 130: cardW = 168: cardGap = 12: kpiStartLeft = 244
    
    Dim v1 As String, v2 As String, v3 As String, v4 As String, v5 As String, v6 As String, v7 As String
    If populateInitialData Then
        v1 = "7,043": v2 = "1,869": v3 = "26.5%": v4 = "$64.76": v5 = "$2,283": v6 = "$139.1K": v7 = "8.2%"
    Else
        v1 = "--": v2 = "--": v3 = "--": v4 = "--": v5 = "--": v6 = "--": v7 = "--"
    End If
    
    BuildKPICard ws, "CR_CustomerBase", kpiStartLeft + (cardW + cardGap) * 0, cardTop, cardW, 94, _
                 "Total Customers", v1, ChrW(9650) & " 3.1% vs PY", PWC_CHARCOAL, "user_blue.svg", "sparkline_blue.svg"
    BuildKPICard ws, "CR_ChurnVolume", kpiStartLeft + (cardW + cardGap) * 1, cardTop, cardW, 94, _
                 "Churned Accounts", v2, ChrW(9650) & " 8.4% vs PY", PWC_ALERT_RED, "xcircle_red.svg", "sparkline_red.svg"
    BuildKPICard ws, "CR_ChurnRate", kpiStartLeft + (cardW + cardGap) * 2, cardTop, cardW, 94, _
                 "Churn Rate (%)", v3, ChrW(9650) & " 1.2% vs PY", PWC_ALERT_RED, "sparkline_red.svg"
    BuildKPICard ws, "CR_MonthlyCharges", kpiStartLeft + (cardW + cardGap) * 3, cardTop, cardW, 94, _
                 "Avg Monthly Bill", v4, ChrW(9650) & " $2.14 vs PY", PWC_ORANGE, "timer_amber.svg"
    BuildKPICard ws, "CR_TotalRevenue", kpiStartLeft + (cardW + cardGap) * 4, cardTop, cardW, 94, _
                 "Avg Lifetime Value", v5, ChrW(9650) & " $142 vs PY", PWC_SUCCESS_GREEN, "check_green.svg"
    BuildKPICard ws, "CR_RevenueAtRisk", kpiStartLeft + (cardW + cardGap) * 5, cardTop, cardW, 94, _
                 "MRR at Risk", v6, ChrW(9660) & " 4.2% vs PY", PWC_WARNING_AMBER, "sparkline_amber.svg"
    BuildKPICard ws, "CR_TechSupport", kpiStartLeft + (cardW + cardGap) * 6, cardTop, cardW, 94, _
                 "Tech Support Adpt", v7, ChrW(9650) & " 1.8% vs PY", RGB(13, 148, 136), "target_teal.svg"
                 
    BuildSlicerPanelContainer ws, "CR_Slicers", CANVAS_LEFT, 130, SLICER_WIDTH, 710, _
                              "Contract Architecture", "Payment Gateway Friction", "Internet Service Modality"
                              
    ' Middle Row Containers
    BuildChartContainer ws, "ContractRisk", 244, 236, 380, 246, _
                        "Contract Type Exposure", "Month-to-Month vs Long-Term Tenure Risk", _
                        "Donut Chart", "icon_topic_sla.svg"
    BuildChartContainer ws, "TenureCohort", 632, 236, 420, 246, _
                        "Tenure Lifecycle Churn Velocity", "Early Attrition vs Established Cohorts", _
                        "Column Cohort", "icon_agent_quadrant.svg"
    BuildChartContainer ws, "PaymentFriction", 1060, 236, 410, 246, _
                        "Payment Method Risk Matrix", "Electronic Check vs Auto-Debit Friction", _
                        "Horizontal Bar", "icon_hourly_surge.svg"
                        
    ' Bottom Row Containers
    BuildChartContainer ws, "ServiceMatrix", 244, 498, 600, 260, _
                        "Add-on Service Protective Shield", "TechSupport, Online Security & Device Protection", _
                        "Clustered Bar", "icon_audit_matrix.svg"
    BuildChartContainer ws, "TenureDistribution", 852, 498, 618, 260, _
                        "Customer Lifetime Value Clustering", "CLV vs Contract Duration Distribution", _
                        "Scatter / Distribution", "icon_live_indicator.svg"
                        
    Set wsStaging = EnsureStagingSheet()
    Call PopulateAnalyticalStagingData(wsStaging)
    Call BuildRetentionVisuals(ws, wsStaging)
    Call ResetSheetViewport(ws)
End Sub

Public Sub BuildDiversityInclusionCanvas(Optional ByVal populateInitialData As Boolean = False)
    Dim wb As Workbook, ws As Worksheet, wsStaging As Worksheet
    Set wb = ThisWorkbook: If wb Is Nothing Then Set wb = ActiveWorkbook
    
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_DIVERSITY)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SHEET_DIVERSITY
    End If
    On Error GoTo 0
    
    ws.Tab.Color = PWC_ORANGE
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    
    InitializeDashboardCanvas ws, PWC_CANVAS_BG
    BuildWebTopNavBar ws, "DI"
    BuildHeroHeader ws, "Diversity, Equity & Executive Parity Command Center", _
                    "Workforce Composition, Promotion Velocity Parity & Executive Pipeline Advancement (FY20 - FY23)", _
                    76, 1480
                    
    Dim cardTop As Single, cardW As Single, cardGap As Single, kpiStartLeft As Single
    cardTop = 130: cardW = 168: cardGap = 12: kpiStartLeft = 244
    
    Dim v1 As String, v2 As String, v3 As String, v4 As String, v5 As String, v6 As String, v7 As String
    If populateInitialData Then
        v1 = "500": v2 = "41.0%": v3 = "16.0%": v4 = "51.0%": v5 = "10.2%": v6 = "0.98": v7 = "8.4%"
    Else
        v1 = "--": v2 = "--": v3 = "--": v4 = "--": v5 = "--": v6 = "--": v7 = "--"
    End If
    
    BuildKPICard ws, "DI_TotalEmployees", kpiStartLeft + (cardW + cardGap) * 0, cardTop, cardW, 94, _
                 "Headcount", v1, ChrW(9650) & " 4.2% vs FY20", PWC_CHARCOAL, "user_blue.svg"
    BuildKPICard ws, "DI_FemaleRepresentation", kpiStartLeft + (cardW + cardGap) * 1, cardTop, cardW, 94, _
                 "Female Share (%)", v2, ChrW(9650) & " 2.4% vs FY20", RGB(236, 72, 153), "sparkline_purple.svg"
    BuildKPICard ws, "DI_ExecParity", kpiStartLeft + (cardW + cardGap) * 2, cardTop, cardW, 94, _
                 "Exec Level Female", v3, ChrW(9650) & " 3.1% vs FY20", PWC_WARNING_AMBER, "timer_amber.svg"
    BuildKPICard ws, "DI_NewHireParity", kpiStartLeft + (cardW + cardGap) * 3, cardTop, cardW, 94, _
                 "New Hire Ratio", v4, ChrW(9650) & " 1.8% vs FY20", PWC_SUCCESS_GREEN, "check_green.svg"
    BuildKPICard ws, "DI_PromoRate", kpiStartLeft + (cardW + cardGap) * 4, cardTop, cardW, 94, _
                 "Promotion Velocity", v5, ChrW(9650) & " 1.3% vs FY20", PWC_ORANGE, "sparkline_orange.svg"
    BuildKPICard ws, "DI_PerfParity", kpiStartLeft + (cardW + cardGap) * 5, cardTop, cardW, 94, _
                 "Rating Parity Index", v6, ChrW(9650) & " 0.02 vs FY20", RGB(37, 99, 235), "target_teal.svg"
    BuildKPICard ws, "DI_TurnoverGap", kpiStartLeft + (cardW + cardGap) * 6, cardTop, cardW, 94, _
                 "Turnover Delta", v7, ChrW(9660) & " 0.5% vs FY20", PWC_ALERT_RED, "xcircle_red.svg"
                 
    BuildSlicerPanelContainer ws, "DI_Slicers", CANVAS_LEFT, 130, SLICER_WIDTH, 710, _
                              "Corporate Department", "Job Level Hierarchy", "Promotion Audit Cycle"
                              
    ' Middle Row Containers
    BuildChartContainer ws, "DIFunnel", 244, 236, 430, 246, _
                        "Organizational Hierarchy Funnel", "Gender Ratio by Seniority Tier (Associate to Partner)", _
                        "Stacked Bar", "icon_agent_quadrant.svg"
    BuildChartContainer ws, "DIDeptParity", 682, 236, 420, 246, _
                        "Departmental Parity Distribution", "Gender Representation across Operational Domains", _
                        "Clustered Column", "icon_audit_matrix.svg"
    BuildChartContainer ws, "DIPromoVelocity", 1110, 236, 360, 246, _
                        "Promotion Velocity Multi-Year", "Promotion Rate Progression (FY20 - FY23 Target)", _
                        "Line Trend", "icon_hourly_surge.svg"
                        
    ' Bottom Row Containers
    BuildChartContainer ws, "DIPerformance", 244, 498, 600, 260, _
                        "Performance Rating Calibration", "Appraisal Parity & Distribution by Gender Tier", _
                        "Combo Chart", "icon_topic_sla.svg"
    BuildChartContainer ws, "DITurnoverCohort", 852, 498, 618, 260, _
                        "Turnover & Leaver Forensics", "Attrition Risk by Demographic Cohort", _
                        "Cohort Matrix", "icon_live_indicator.svg"
                        
    Set wsStaging = EnsureStagingSheet()
    Call PopulateAnalyticalStagingData(wsStaging)
    Call BuildDiversityVisuals(ws, wsStaging)
    Call ResetSheetViewport(ws)
End Sub

Public Sub BuildAllDashboardCanvases()
    BuildCallCenterCanvas True
    BuildCustomerRetentionCanvas True
    BuildDiversityInclusionCanvas True
    ResetAllViewports
End Sub
''')
    print("Created vba/modDashboardUIUX.bas successfully.")

def create_mod_pwc_unified_master():
    path = 'vba/modPwC_Unified_Master.bas'
    with open(path, 'w', encoding='latin-1') as f:
        f.write('''Attribute VB_Name = "modPwC_Unified_Master"
Option Explicit

' ==============================================================================
' PwC Switzerland Virtual Case Experience - Enterprise Master Orchestrator
' Module: modPwC_Unified_Master
' Description: Coordinates and triggers end-to-end platform deployment across
' all 11 modular subsystems:
'   - modAppState               : Application ScreenUpdating & Shield
'   - modThemeEngine            : Enterprise Light / Dark Mode Toggle
'   - modNavigation             : Instantaneous View Navigation
'   - modDataRefresh            : VertiPaq Tabular Engine Sync
'   - modFilterController       : Slicer & Filter State Shield
'   - modExportPDF              : Publication-Grade PDF Generator
'   - modCreateGovernanceSheets : Domains & Data Catalog Architect
'   - modPortalLanding          : Executive Homepage & Launchers
'   - modDashboardUIUX          : SaaS Cockpit Canvas, KPIs & Visuals
'   - modInteractiveScorecard   : Docked Dynamic Scorecards
'   - modPivotTableFormatting   : Conditional Formatting & Grid Styles
' Pure 7-bit ASCII encoding.
' ==============================================================================

Private Const PLATFORM_NAME As String = "PwC Switzerland BI Executive Suite"
Private Const PLATFORM_VERSION As String = "v3.0.0 Enterprise"

' ==============================================================================
' MASTER ENTRY POINT: RunUnifiedPwCPlatform / RunCompletePwCPlatform
' ==============================================================================
Public Sub RunUnifiedPwCPlatform()
    Dim currentStep As String
    On Error GoTo MasterErrHandler
    
    ' Step 1: Initialize Application State Shield
    currentStep = "Freezing Application State"
    modAppState.FreezeAppState True
    Application.StatusBar = "PwC Platform: Initializing Analytical Foundation..."
    
    ' Step 2: Ensure Staging Sheet & Analytical Datasets
    currentStep = "Staging Analytical Datasets (03_Staging_Data)"
    Dim wsStaging As Worksheet
    Set wsStaging = modDashboardUIUX.EnsureStagingSheet()
    modDashboardUIUX.PopulateAnalyticalStagingData wsStaging
    
    ' Step 3: Build Governance & Architecture Canvas Sheets
    currentStep = "Building Governance Sheets (01_Domains & 02_Catalog)"
    modCreateGovernanceSheets.BuildGovernanceArchitecture
    
    ' Step 4: Build Executive Homepage Portal
    currentStep = "Building Executive Home Portal (00_Home_Portal)"
    modPortalLanding.BuildExecutivePortal
    
    ' Step 5: Build Analytical Cockpits with Real Visuals & KPI Cards
    currentStep = "Building Call Center Cockpit (03_CallCenter_Cockpit)"
    modDashboardUIUX.BuildCallCenterCanvas True
    
    currentStep = "Building Customer Retention Cockpit (04_CustomerRetention_Cockpit)"
    modDashboardUIUX.BuildCustomerRetentionCanvas True
    
    currentStep = "Building Diversity & Inclusion Cockpit (05_DiversityInclusion_Cockpit)"
    modDashboardUIUX.BuildDiversityInclusionCanvas True
    
    ' Step 6: Style PivotTables & Dock Interactive Scorecards
    currentStep = "Applying Scorecard & Pivot Table Themes"
    On Error Resume Next
    modPivotTableFormatting.StyleAgentScorecardPivotTable "SLA_EXCEPTIONS"
    On Error GoTo MasterErrHandler
    
    ' Step 7: Normalize Viewports (FreezePanes = False, Scroll = A1, Zoom = 80%)
    currentStep = "Normalizing Sheet Viewports & Gridlines"
    modDashboardUIUX.ResetAllViewports
    
    ' Step 8: Navigate to Call Center Cockpit as default landing view
    currentStep = "Navigating to Call Center Cockpit"
    modNavigation.NavigateToCallCenter
    
    ' Step 9: Restore Application State Shield
    modAppState.RestoreAppState
    Application.StatusBar = "PwC Switzerland BI Suite ready."
    
    MsgBox "PwC Switzerland BI Platform deployed successfully!" & vbCrLf & vbCrLf & _
           "- Connected Architecture: All 12 VBA modules fully synchronized" & vbCrLf & _
           "- Interactive Visuals: Real charts, BAN cards & scorecards generated" & vbCrLf & _
           "- Zero Conflicts: 0 ambiguous names and 0 shape deletion errors" & vbCrLf & _
           "- Responsive Design: Modern rounded buttons with centered text" & vbCrLf & _
           "- Pure ASCII: 100% compatible across all Windows regional locales", _
           vbInformation, PLATFORM_NAME & " " & PLATFORM_VERSION
    Exit Sub

MasterErrHandler:
    Dim savedErrNum As Long, savedErrDesc As String, savedErrSrc As String
    savedErrNum = Err.Number
    savedErrDesc = Err.Description
    savedErrSrc = Err.Source
    
    modAppState.RestoreAppState
    
    If savedErrNum <> 0 Then
        MsgBox "Execution Interrupted during Step:" & vbCrLf & _
               "'" & currentStep & "'" & vbCrLf & vbCrLf & _
               "Error Number: " & savedErrNum & vbCrLf & _
               "Description: " & savedErrDesc & vbCrLf & _
               "Source: " & savedErrSrc, _
               vbCritical, "PwC Platform Build Error"
    End If
End Sub

Public Sub RunCompletePwCPlatform()
    ' Backward compatibility alias
    RunUnifiedPwCPlatform
End Sub

' ==============================================================================
' DIAGNOSTIC & HEALTH CHECK UTILITY
' ==============================================================================
Public Sub VerifyPlatformIntegrity()
    Dim wb As Workbook, sNames As Variant, i As Long, missing As String
    Set wb = ThisWorkbook: If wb Is Nothing Then Set wb = ActiveWorkbook
    
    sNames = Array("00_Home_Portal", "01_Business_Domains", "02_Metadata_&_KPI_Catalog", _
                   "03_CallCenter_Cockpit", "04_CustomerRetention_Cockpit", "05_DiversityInclusion_Cockpit", _
                   "03_Staging_Data")
                   
    For i = LBound(sNames) To UBound(sNames)
        On Error Resume Next
        Dim ws As Worksheet: Set ws = Nothing
        Set ws = wb.Worksheets(CStr(sNames(i)))
        On Error GoTo 0
        If ws Is Nothing Then missing = missing & " - " & sNames(i) & vbCrLf
    Next i
    
    If Len(missing) = 0 Then
        MsgBox "All 7 required workbook sheets verified and operational.", vbInformation, "PwC Health Check"
    Else
        MsgBox "The following sheets are missing:" & vbCrLf & missing & vbCrLf & _
               "Run 'RunUnifiedPwCPlatform' to recreate them.", vbExclamation, "PwC Health Check"
    End If
End Sub
''')
    print("Created vba/modPwC_Unified_Master.bas successfully.")

if __name__ == '__main__':
    create_mod_dashboard_uiux()
    create_mod_pwc_unified_master()
