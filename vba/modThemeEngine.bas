Attribute VB_Name = "modThemeEngine"
Option Explicit

' ==============================================================================
' PwC Switzerland Virtual Case Experience - BI Executive Platform
' Module: modThemeEngine
' Description: Enterprise Light / Dark Mode Dynamic Theme Engine
' Provides seamless real-time visual theme switching across all dashboard
' cockpits, containers, typography, and charts with state persistence.
' Pure ASCII Compatible.
' ==============================================================================

' --- Color Palette Definitions (Windows Long / BGR format) ---
' Light Theme Palettes
Private Const LIGHT_BG As Long = 16579832           ' #F8FAFC (Slate 50)
Private Const LIGHT_CONTAINER As Long = 16777215    ' #FFFFFF (Crisp White)
Private Const LIGHT_BORDER As Long = 15790306       ' #E2E8F0 (Slate 200)
Private Const LIGHT_TEXT_PRIMARY As Long = 2762511  ' #0F172A (Deep Slate Navy)
Private Const LIGHT_TEXT_MUTED As Long = 9141092    ' #64748B (Slate 500)
Private Const LIGHT_HEADER_BG As Long = 2762511     ' #0F172A (Dark Slate)
Private Const LIGHT_HEADER_TEXT As Long = 16777215  ' #FFFFFF (Crisp White)
Private Const LIGHT_ACCENT As Long = 150224         ' #D04A02 (PwC Vibrant Tangerine Orange)
Private Const LIGHT_GRIDLINE As Long = 16382457     ' #F1F5F9 (Slate 100)
Private Const LIGHT_CARD_BG As Long = 16777215      ' #FFFFFF

' Dark Theme Palettes
Private Const DARK_BG As Long = 1642251             ' #0B0F19 (Pitch Slate Navy)
Private Const DARK_CONTAINER As Long = 3022358      ' #161E2E (Deep Elevated Card)
Private Const DARK_BORDER As Long = 4666410         ' #2A3447 (Border Slate)
Private Const DARK_TEXT_PRIMARY As Long = 16579832  ' #F8FAFC (Pure White/Slate 50)
Private Const DARK_TEXT_MUTED As Long = 12099732    ' #94A3B8 (Slate 400)
Private Const DARK_HEADER_BG As Long = 3879201      ' #1E293B (Card Header Navy)
Private Const DARK_HEADER_TEXT As Long = 16579832   ' #F8FAFC (Crisp White)
Private Const DARK_ACCENT As Long = 3373823         ' #FF7A33 (Vibrant Glowing Orange)
Private Const DARK_GRIDLINE As Long = 3022358       ' #161E2E (Subtle Dark Grid)
Private Const DARK_CARD_BG As Long = 2367258        ' #1E293B

' State Storage Property Name
Private Const THEME_PROP_NAME As String = "PwC_Dashboard_Theme"

' ==============================================================================
' Public API: Toggle Current Theme
' ==============================================================================
Public Sub ToggleDashboardTheme()
    On Error GoTo ErrorHandler
    Dim currentTheme As String
    currentTheme = GetStoredTheme()
    
    If UCase$(currentTheme) = "DARK" Then
        ApplyTheme False  ' Switch to Light
    Else
        ApplyTheme True   ' Switch to Dark
    End If
    Exit Sub
ErrorHandler:
    If Application.Visible And Application.UserControl Then
        MsgBox "Theme Toggle encountered an error: " & Err.Description, vbExclamation, "PwC Theme Engine"
    End If
End Sub

Public Sub ToggleTheme()
    ToggleDashboardTheme
End Sub

' ==============================================================================
' Public API: Explicit Light / Dark Application
' ==============================================================================
Public Sub ApplyLightTheme()
    ApplyTheme False
End Sub

Public Sub ApplyDarkTheme()
    ApplyTheme True
End Sub

Public Sub ApplyTheme(ByVal isDark As Boolean)
    On Error Resume Next
    Dim ws As Worksheet
    Dim targetSheets As Variant
    Dim i As Long
    Dim sheetName As String
    
    ' Optimize Excel Performance
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Store the current state in Workbook Custom Document Properties
    SetStoredTheme IIf(isDark, "DARK", "LIGHT")
    
    ' Dashboard sheets to apply theme to
    targetSheets = Array("00_Home_Portal", "01_Business_Domains", "02_Metadata_&_KPI_Catalog", "03_CallCenter_Cockpit", "04_CustomerRetention_Cockpit", "05_DiversityInclusion_Cockpit")
    
    For i = LBound(targetSheets) To UBound(targetSheets)
        sheetName = CStr(targetSheets(i))
        Set ws = Nothing
        Set ws = ThisWorkbook.Worksheets(sheetName)
        
        If Not ws Is Nothing Then
            ApplyThemeToSheet ws, isDark
        End If
    Next i
    
    ' Re-enable application settings
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    ' Visual notification in status bar
    Application.StatusBar = "PwC BI Suite: " & IIf(isDark, "[DARK MODE] Theme Activated", "[LIGHT MODE] Theme Activated")
    DoEvents
    Application.OnTime Now + TimeValue("00:00:03"), "modThemeEngine.ClearStatusBar"
End Sub

Public Sub ClearStatusBar()
    Application.StatusBar = False
End Sub

' ==============================================================================
' Internal Engine: Apply Theme to a Single Worksheet
' ==============================================================================
Private Sub ApplyThemeToSheet(ByVal ws As Worksheet, ByVal isDark As Boolean)
    On Error Resume Next
    Dim shp As Shape
    Dim chObj As ChartObject
    Dim bgColor As Long
    Dim containerColor As Long
    Dim borderColor As Long
    Dim textPrimary As Long
    Dim textMuted As Long
    Dim cardBgColor As Long
    
    If isDark Then
        bgColor = DARK_BG
        containerColor = DARK_CONTAINER
        borderColor = DARK_BORDER
        textPrimary = DARK_TEXT_PRIMARY
        textMuted = DARK_TEXT_MUTED
        cardBgColor = DARK_CARD_BG
    Else
        bgColor = LIGHT_BG
        containerColor = LIGHT_CONTAINER
        borderColor = LIGHT_BORDER
        textPrimary = LIGHT_TEXT_PRIMARY
        textMuted = LIGHT_TEXT_MUTED
        cardBgColor = LIGHT_CARD_BG
    End If
    
    ' 1. Canvas Background Color (Cells A1:AZ100)
    ws.Cells.Interior.Color = bgColor
    
    ' 2. Loop Through All Shapes on Sheet
    For Each shp In ws.Shapes
        ' Container Backgrounds
        If InStr(1, shp.Name, "Container_", vbTextCompare) > 0 Or InStr(1, shp.Name, "Card_", vbTextCompare) > 0 Then
            shp.Fill.Solid
            shp.Fill.ForeColor.RGB = containerColor
            shp.Line.ForeColor.RGB = borderColor
            shp.Line.Weight = 1
        
        ' DockZones (Inner Frames)
        ElseIf InStr(1, shp.Name, "DockZone_", vbTextCompare) > 0 Then
            shp.Fill.Solid
            shp.Fill.ForeColor.RGB = IIf(isDark, DARK_CONTAINER, LIGHT_CONTAINER)
            shp.Line.ForeColor.RGB = borderColor
            shp.Line.Weight = 0.75
            shp.Line.DashStyle = msoLineSolid
        
        ' Filter Drawer Panel
        ElseIf InStr(1, shp.Name, "Panel_", vbTextCompare) > 0 Then
            shp.Fill.Solid
            shp.Fill.ForeColor.RGB = containerColor
            shp.Line.ForeColor.RGB = borderColor
        
        ' Badges
        ElseIf InStr(1, shp.Name, "Badge_", vbTextCompare) > 0 Then
            shp.Fill.Solid
            shp.Fill.ForeColor.RGB = IIf(isDark, RGB(30, 41, 59), RGB(241, 245, 249))
            shp.Line.ForeColor.RGB = borderColor
            If shp.TextFrame2.HasText Then
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = IIf(isDark, RGB(203, 213, 225), RGB(71, 85, 105))
            End If
            
        ' Header Text
        ElseIf InStr(1, shp.Name, "Header_", vbTextCompare) > 0 Or InStr(1, shp.Name, "Title_", vbTextCompare) > 0 Then
            If shp.TextFrame2.HasText Then
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = textPrimary
            End If
            
        ' Subtitle & Description Text
        ElseIf InStr(1, shp.Name, "Sub_", vbTextCompare) > 0 Or InStr(1, shp.Name, "Desc_", vbTextCompare) > 0 Then
            If shp.TextFrame2.HasText Then
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = textMuted
            End If
            
        ' BAN KPI Value Text
        ElseIf InStr(1, shp.Name, "BAN_", vbTextCompare) > 0 Or InStr(1, shp.Name, "KPI_Val", vbTextCompare) > 0 Or InStr(1, shp.Name, "Value_", vbTextCompare) > 0 Then
            If shp.TextFrame2.HasText Then
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = IIf(isDark, DARK_ACCENT, LIGHT_ACCENT)
            End If
            
        ' Theme Switcher Button Itself
        ElseIf InStr(1, shp.Name, "ThemeToggle", vbTextCompare) > 0 Or InStr(1, shp.Name, "btn_Theme", vbTextCompare) > 0 Then
            shp.Fill.Solid
            shp.Fill.ForeColor.RGB = IIf(isDark, RGB(30, 41, 59), RGB(241, 245, 249))
            shp.Line.ForeColor.RGB = IIf(isDark, DARK_ACCENT, LIGHT_BORDER)
            If shp.TextFrame2.HasText Then
                shp.TextFrame2.TextRange.Text = IIf(isDark, "LIGHT MODE", "DARK MODE")
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = IIf(isDark, RGB(255, 255, 255), RGB(15, 23, 42))
                shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
                shp.TextFrame2.MarginTop = 0
                shp.TextFrame2.MarginBottom = 0
            End If
            
        ' Interactive Slicers
        ElseIf shp.Type = msoSlicer Then
            On Error Resume Next
            shp.Slicer.Style = IIf(isDark, "SlicerStyleDark2", "SlicerStyleLight2")
            On Error GoTo 0
        End If
    Next shp
    
    ' 3. Chart Objects Formatting
    For Each chObj In ws.ChartObjects
        FormatChartForTheme chObj, isDark
    Next chObj
End Sub

' ==============================================================================
' Chart Theme Adaptation Helper
' ==============================================================================
Private Sub FormatChartForTheme(ByVal chObj As ChartObject, ByVal isDark As Boolean)
    On Error Resume Next
    Dim ch As Chart
    Dim ax As Axis
    Dim textColor As Long
    Dim gridColor As Long
    
    Set ch = chObj.Chart
    textColor = IIf(isDark, DARK_TEXT_MUTED, LIGHT_TEXT_MUTED)
    gridColor = IIf(isDark, DARK_GRIDLINE, LIGHT_GRIDLINE)
    
    ' Transparent chart canvas
    ch.ChartArea.Format.Fill.Visible = msoFalse
    ch.ChartArea.Format.Line.Visible = msoFalse
    ch.PlotArea.Format.Fill.Visible = msoFalse
    ch.PlotArea.Format.Line.Visible = msoFalse
    
    ' Axis typography & lines
    For Each ax In ch.Axes
        ax.Format.Line.ForeColor.RGB = gridColor
        ax.TickLabels.Font.Color = textColor
        ax.MajorGridlines.Format.Line.ForeColor.RGB = gridColor
    Next ax
    
    ' Legend typography
    If ch.HasLegend Then
        ch.Legend.Format.Fill.Visible = msoFalse
        ch.Legend.Format.Line.Visible = msoFalse
        ch.Legend.Font.Color = textColor
    End If
End Sub

' ==============================================================================
' Persistence: Custom Document Properties
' ==============================================================================
Private Function GetStoredTheme() As String
    On Error Resume Next
    GetStoredTheme = ThisWorkbook.CustomDocumentProperties(THEME_PROP_NAME).Value
    If Err.Number <> 0 Or GetStoredTheme = "" Then
        GetStoredTheme = "LIGHT"
    End If
End Function

Private Sub SetStoredTheme(ByVal themeName As String)
    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties(THEME_PROP_NAME).Value = themeName
    If Err.Number <> 0 Then
        ThisWorkbook.CustomDocumentProperties.Add _
            Name:=THEME_PROP_NAME, _
            LinkToContent:=False, _
            Type:=msoPropertyTypeString, _
            Value:=themeName
    End If
End Sub
