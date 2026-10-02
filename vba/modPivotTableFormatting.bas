Attribute VB_Name = "modPivotTableFormatting"
Option Explicit

' ==============================================================================
' PwC Switzerland Digital Accelerator -- PivotTable & Scorecard Formatting Engine
' Advanced Theme-Compliant Styling, Typography & Executive Conditional Formatting
' ==============================================================================
'
' Features:
'   1. StyleAgentScorecardPivotTable:
'      Formats Visual 1.4 (CC_AgentScorecard) on 03_CallCenter_Cockpit
'      into a compact, borderless, SaaS-grade executive scorecard.
'   2. Four Selectable Conditional Formatting Themes:
'      - Theme A (Executive SLA Exception Matrix): Subtle breach highlights
'      - Theme B (Dual-Engine Micro Data Bars): Normalized gradient bars
'      - Theme C (PwC 3-Tier Soft Color Scales): Low/Mid/High pastel shading
'      - Theme D (Top/Bottom Milestone Badges): Highlights #1 Star & Coaching
'   3. Universal Declutter & Alignment:
'      Segoe UI typography, tabular layout, right-aligned metrics, custom number
'      formats (#,##0; 0.0%; 0.0 "s"; 0.00).
'
' ==============================================================================

' ------------------------------------------------------------------------------
' PwC Enterprise Brand Color Constants (RGB Values)
' ------------------------------------------------------------------------------
Private Const PWC_TANGERINE        As Long = 133288     ' #D04A02 - RGB(208, 74, 2)
Private Const PWC_DARK_SLATE       As Long = 2758415    ' #0F172A - RGB(15, 23, 42)
Private Const PWC_CHARCOAL          As Long = 3877150    ' #1E293B - RGB(30, 41, 59)
Private Const PWC_CANVAS_BG         As Long = 16579320   ' #F8FAFC - RGB(248, 250, 252)
Private Const PWC_CARD_FILL         As Long = 16777215   ' #FFFFFF - RGB(255, 255, 255)
Private Const PWC_BORDER_MUTED      As Long = 15790322   ' #E2E8F0 - RGB(226, 232, 240)
Private Const PWC_TEXT_TITLE        As Long = 3877150    ' #1E293B - RGB(30, 41, 59)
Private Const PWC_TEXT_MUTED        As Long = 9141108    ' #64748B - RGB(100, 116, 139)

' Soft Executive Alert & Badge Colors (Pastel fills + dark contrast text)
Private Const PWC_SOFT_RED_BG       As Long = 14803454   ' #FEE2E2 - RGB(254, 226, 226)
Private Const PWC_DARK_RED_TXT      As Long = 1776537    ' #991B1B - RGB(153, 27, 27)
Private Const PWC_SOFT_GREEN_BG     As Long = 13761756   ' #DCFCE7 - RGB(220, 252, 231)
Private Const PWC_DARK_GREEN_TXT    As Long = 3433746    ' #166534 - RGB(22, 101, 52)
Private Const PWC_SOFT_AMBER_BG     As Long = 13104126   ' #FEF3C7 - RGB(254, 243, 199)
Private Const PWC_DARK_AMBER_TXT    As Long = 933902     ' #92400E - RGB(146, 64, 14)
Private Const PWC_SOFT_BLUE_BG      As Long = 16706267   ' #DBEAFE - RGB(219, 234, 254)
Private Const PWC_DARK_BLUE_TXT     As Long = 12470046   ' #1E40AF - RGB(30, 64, 175)

Private Const FONT_NAME             As String = "Segoe UI"

' ==============================================================================
' 1. MAIN ENTRY POINT: FORMAT THE AGENT SCORECARD PIVOTTABLE
' ==============================================================================
Public Sub StyleAgentScorecardPivotTable(Optional ByVal themeChoice As String = "SLA_EXCEPTIONS")
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim targetSheetName As String
    targetSheetName = "03_CallCenter_Cockpit"
    
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(targetSheetName)
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    ' Look for the Scorecard PivotTable
    If ws.PivotTables.Count = 0 Then
        ' Check if PivotTable is on Staging_Pivots
        Dim wsStaging As Worksheet
        Set wsStaging = ActiveWorkbook.Worksheets("Staging_Pivots")
        If Not wsStaging Is Nothing Then
            If wsStaging.PivotTables.Count > 0 Then
                Set pt = wsStaging.PivotTables(wsStaging.PivotTables.Count)
            End If
        End If
    Else
        ' Take the last or designated PivotTable
        Set pt = ws.PivotTables(ws.PivotTables.Count)
    End If
    On Error GoTo 0
    
    If pt Is Nothing Then
        MsgBox "No PivotTable found on '" & ws.Name & "' or 'Staging_Pivots'." & vbCrLf & _
               "Please create your Agent Scorecard PivotTable first, then run this routine.", _
               vbExclamation, "PwC Scorecard Formatter"
        Exit Sub
    End If
    
    ' Execute formatting with AppState protection
    On Error Resume Next
    modAppState.FreezeAppState
    On Error GoTo 0
    
    FormatScorecardGrid pt
    ApplyScorecardConditionalFormatting pt, UCase(Trim(themeChoice))
    
    On Error Resume Next
    modAppState.RestoreAppState
    On Error GoTo 0
    
    If Application.UserControl Then
        MsgBox "Agent Scorecard PivotTable successfully formatted with Theme: " & themeChoice & "!", _
               vbInformation, "PwC Scorecard Formatter"
    End If
End Sub

' ==============================================================================
' 2. BASE GRID FORMATTER: LAYOUT, HEADERS, TYPOGRAPHY & NUMBER FORMATTING
' ==============================================================================
Public Sub FormatScorecardGrid(pt As PivotTable)
    Dim pf As PivotField
    Dim rngTable As Range
    
    On Error Resume Next
    
    With pt
        ' Layout: Tabular, compact, no grand totals for averages
        .RowAxisLayout xlTabularRow
        .ColumnGrand = False
        .RowGrand = False
        .ShowTableStyleRowStripes = True
        .ShowTableStyleColumnStripes = False
        .TableStyle2 = "PivotStyleLight1"
        
        ' Format whole table font
        Set rngTable = .TableRange2
        If Not rngTable Is Nothing Then
            With rngTable.Font
                .Name = FONT_NAME
                .Size = 8.5
                .Color = PWC_CHARCOAL
            End With
        End If
        
        ' Format Data Field Headers & Number Formats
        Dim i As Long
        For i = 1 To .DataFields.Count
            Set pf = .DataFields(i)
            Select Case True
                Case InStr(1, pf.Caption, "Call", vbTextCompare) > 0 Or InStr(1, pf.SourceName, "Call", vbTextCompare) > 0 Or InStr(1, pf.Caption, "Demand", vbTextCompare) > 0
                    pf.Caption = "Calls Taken"
                    pf.NumberFormat = "#,##0"
                Case InStr(1, pf.Caption, "Answer Rate", vbTextCompare) > 0 Or InStr(1, pf.SourceName, "Answer Rate", vbTextCompare) > 0
                    pf.Caption = "Answer Rate %"
                    pf.NumberFormat = "0.0%"
                Case InStr(1, pf.Caption, "Resolution", vbTextCompare) > 0 Or InStr(1, pf.SourceName, "Resolution", vbTextCompare) > 0 Or InStr(1, pf.Caption, "FCR", vbTextCompare) > 0
                    pf.Caption = "FCR Rate %"
                    pf.NumberFormat = "0.0%"
                Case InStr(1, pf.Caption, "Speed", vbTextCompare) > 0 Or InStr(1, pf.SourceName, "Speed", vbTextCompare) > 0 Or InStr(1, pf.Caption, "ASA", vbTextCompare) > 0
                    pf.Caption = "Avg Speed (s)"
                    pf.NumberFormat = "0.0 ""s"""
                Case InStr(1, pf.Caption, "CSAT", vbTextCompare) > 0 Or InStr(1, pf.SourceName, "CSAT", vbTextCompare) > 0 Or InStr(1, pf.Caption, "Satisfaction", vbTextCompare) > 0
                    pf.Caption = "Avg CSAT"
                    pf.NumberFormat = "0.00"
            End Select
        Next i
    End With
    On Error GoTo 0
End Sub

' ==============================================================================
' 3. CONDITIONAL FORMATTING ENGINE (4 EXECUTIVE THEMES)
' ==============================================================================
Public Sub ApplyScorecardConditionalFormatting(pt As PivotTable, ByVal theme As String)
    Dim rngData As Range
    Dim pf As PivotField
    Dim fldRng As Range
    Dim i As Long
    
    On Error Resume Next
    
    ' Clear existing conditional formats on the PivotTable data range
    Set rngData = pt.DataBodyRange
    If Not rngData Is Nothing Then
        rngData.FormatConditions.Delete
    End If
    
    For i = 1 To pt.DataFields.Count
        Set pf = pt.DataFields(i)
        Set fldRng = pf.DataRange
        
        If Not fldRng Is Nothing Then
            Select Case theme
                ' --------------------------------------------------------------
                ' THEME 1: EXECUTIVE SLA EXCEPTION MATRIX (Recommended Default)
                ' Discreet pastel badges for SLA thresholds & inverted speed warning
                ' --------------------------------------------------------------
                Case "SLA_EXCEPTIONS", "DEFAULT", "OPTION_1"
                    Select Case True
                        Case InStr(1, pf.Caption, "Calls", vbTextCompare) > 0
                            ' Subtle soft gradient data bar for volume
                            ApplySoftDataBar fldRng, PWC_BORDER_MUTED
                            
                        Case InStr(1, pf.Caption, "Answer", vbTextCompare) > 0
                            ' Critical SLA threshold (< 80% is Red Alert, >= 83% is Green)
                            ApplyTwoThresholdAlert fldRng, "0.80", "0.83", _
                                                   PWC_SOFT_RED_BG, PWC_DARK_RED_TXT, _
                                                   PWC_SOFT_GREEN_BG, PWC_DARK_GREEN_TXT
                                                   
                        Case InStr(1, pf.Caption, "FCR", vbTextCompare) > 0
                            ' Resolution threshold (< 89% is Amber Warning, >= 91% is Green)
                            ApplyTwoThresholdAlert fldRng, "0.89", "0.91", _
                                                   PWC_SOFT_AMBER_BG, PWC_DARK_AMBER_TXT, _
                                                   PWC_SOFT_GREEN_BG, PWC_DARK_GREEN_TXT
                                                   
                        Case InStr(1, pf.Caption, "Speed", vbTextCompare) > 0
                            ' Inverted threshold: Speed > 70s is Amber Warning, < 66s is Green
                            ApplyInvertedSpeedAlert fldRng, "70", "66", _
                                                    PWC_SOFT_AMBER_BG, PWC_DARK_AMBER_TXT, _
                                                    PWC_SOFT_GREEN_BG, PWC_DARK_GREEN_TXT
                                                    
                        Case InStr(1, pf.Caption, "CSAT", vbTextCompare) > 0
                            ' CSAT Tier: < 3.35 is Alert Red, >= 3.45 is Star Green
                            ApplyTwoThresholdAlert fldRng, "3.35", "3.45", _
                                                   PWC_SOFT_RED_BG, PWC_DARK_RED_TXT, _
                                                   PWC_SOFT_GREEN_BG, PWC_DARK_GREEN_TXT
                    End Select
                    
                ' --------------------------------------------------------------
                ' THEME 2: DUAL-ENGINE MICRO DATA BARS
                ' Normalized compact visual bars across all metrics
                ' --------------------------------------------------------------
                Case "DATA_BARS", "OPTION_2"
                    Select Case True
                        Case InStr(1, pf.Caption, "Calls", vbTextCompare) > 0
                            ApplySoftDataBar fldRng, PWC_CHARCOAL
                        Case InStr(1, pf.Caption, "Answer", vbTextCompare) > 0
                            ApplySoftDataBar fldRng, PWC_SOFT_GREEN_BG, 0.7, 1#
                        Case InStr(1, pf.Caption, "FCR", vbTextCompare) > 0
                            ApplySoftDataBar fldRng, PWC_DARK_BLUE_TXT, 0.8, 1#
                        Case InStr(1, pf.Caption, "Speed", vbTextCompare) > 0
                            ApplySoftDataBar fldRng, PWC_DARK_AMBER_TXT, 60, 75
                        Case InStr(1, pf.Caption, "CSAT", vbTextCompare) > 0
                            ApplySoftDataBar fldRng, PWC_TANGERINE, 3#, 3.6
                    End Select
                    
                ' --------------------------------------------------------------
                ' THEME 3: PWC 3-TIER SOFT COLOR SCALES (HEATMAP)
                ' Low / Mid / High pastel gradient fills
                ' --------------------------------------------------------------
                Case "COLOR_SCALES", "HEATMAP", "OPTION_3"
                    Select Case True
                        Case InStr(1, pf.Caption, "Speed", vbTextCompare) > 0
                            ' Inverted color scale (Green for low seconds, Red for high seconds)
                            ApplyColorScale3 fldRng, PWC_SOFT_GREEN_BG, PWC_SOFT_AMBER_BG, PWC_SOFT_RED_BG
                        Case InStr(1, pf.Caption, "Calls", vbTextCompare) > 0
                            ' Volume data bar
                            ApplySoftDataBar fldRng, PWC_BORDER_MUTED
                        Case Else
                            ' Higher is better: Red -> Amber -> Green
                            ApplyColorScale3 fldRng, PWC_SOFT_RED_BG, PWC_SOFT_AMBER_BG, PWC_SOFT_GREEN_BG
                    End Select
                    
                ' --------------------------------------------------------------
                ' THEME 4: TOP / BOTTOM MILESTONE BADGES (EXECUTIVE AUDIT)
                ' Only highlights the #1 Top Star and the #1 Bottom Coaching need
                ' --------------------------------------------------------------
                Case "TOP_BOTTOM", "MILESTONES", "OPTION_4"
                    Select Case True
                        Case InStr(1, pf.Caption, "Calls", vbTextCompare) > 0
                            ApplySoftDataBar fldRng, PWC_BORDER_MUTED
                        Case InStr(1, pf.Caption, "Speed", vbTextCompare) > 0
                            ' Fast Pickup (#1 Low): Green; Slowest (#1 High): Amber
                            ApplyTopBottomMilestones fldRng, True
                        Case Else
                            ' #1 High: Green; #1 Low: Red
                            ApplyTopBottomMilestones fldRng, False
                    End Select
            End Select
        End If
    Next i
    On Error GoTo 0
End Sub

' ==============================================================================
' 4. HELPER CONDITIONAL FORMATTING BUILDERS
' ==============================================================================

' Apply 2-Threshold Highlight (Low = Alert, High = Success)
Private Sub ApplyTwoThresholdAlert(rng As Range, ByVal lowVal As String, ByVal highVal As String, _
                                   ByVal lowBg As Long, ByVal lowTxt As Long, _
                                   ByVal highBg As Long, ByVal highTxt As Long)
    On Error Resume Next
    Dim fcLow As FormatCondition
    Dim fcHigh As FormatCondition
    
    ' Low Threshold (Alert)
    Set fcLow = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:=lowVal)
    With fcLow
        .Interior.Color = lowBg
        .Font.Color = lowTxt
        .Font.Bold = True
    End With
    
    ' High Threshold (Success)
    Set fcHigh = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:=highVal)
    With fcHigh
        .Interior.Color = highBg
        .Font.Color = highTxt
        .Font.Bold = True
    End With
    On Error GoTo 0
End Sub

' Apply Inverted 2-Threshold Highlight for Speed (High Seconds = Alert, Low Seconds = Success)
Private Sub ApplyInvertedSpeedAlert(rng As Range, ByVal highSecondsVal As String, ByVal lowSecondsVal As String, _
                                    ByVal highBg As Long, ByVal highTxt As Long, _
                                    ByVal lowBg As Long, ByVal lowTxt As Long)
    On Error Resume Next
    Dim fcHigh As FormatCondition
    Dim fcLow As FormatCondition
    
    ' High Seconds (Slow / Queue Risk)
    Set fcHigh = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:=highSecondsVal)
    With fcHigh
        .Interior.Color = highBg
        .Font.Color = highTxt
        .Font.Bold = True
    End With
    
    ' Low Seconds (Fast Pickup / SLA Star)
    Set fcLow = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLessEqual, Formula1:=lowSecondsVal)
    With fcLow
        .Interior.Color = lowBg
        .Font.Color = lowTxt
        .Font.Bold = True
    End With
    On Error GoTo 0
End Sub

' Apply Soft Data Bar with Optional Fixed Min/Max Scaling
Private Sub ApplySoftDataBar(rng As Range, ByVal barColor As Long, _
                             Optional ByVal fixedMin As Double = -1, _
                             Optional ByVal fixedMax As Double = -1)
    On Error Resume Next
    Dim db As Databar
    Set db = rng.FormatConditions.AddDatabar
    With db
        .Color.Color = barColor
        .BarFillType = xlDataBarFillGradient
        .BarBorder.Type = xlDataBarBorderNone
        
        If fixedMin >= 0 Then
            .MinPoint.Modify xlConditionValueNumber, fixedMin
        Else
            .MinPoint.Modify xlConditionValueAutomaticMin
        End If
        
        If fixedMax > 0 Then
            .MaxPoint.Modify xlConditionValueNumber, fixedMax
        Else
            .MaxPoint.Modify xlConditionValueAutomaticMax
        End If
    End With
    On Error GoTo 0
End Sub

' Apply 3-Color Scale (Soft Low -> Mid -> High Gradient)
Private Sub ApplyColorScale3(rng As Range, ByVal colMin As Long, ByVal colMid As Long, ByVal colMax As Long)
    On Error Resume Next
    Dim cs As ColorScale
    Set cs = rng.FormatConditions.AddColorScale(ColorScaleType:=3)
    With cs
        .ColorScaleCriteria(1).Type = xlConditionValueLowestValue
        .ColorScaleCriteria(1).FormatColor.Color = colMin
        
        .ColorScaleCriteria(2).Type = xlConditionValuePercentile
        .ColorScaleCriteria(2).Value = 50
        .ColorScaleCriteria(2).FormatColor.Color = colMid
        
        .ColorScaleCriteria(3).Type = xlConditionValueHighestValue
        .ColorScaleCriteria(3).FormatColor.Color = colMax
    End With
    On Error GoTo 0
End Sub

' Apply Top 1 & Bottom 1 Milestone Badges
Private Sub ApplyTopBottomMilestones(rng As Range, ByVal isInverted As Boolean)
    On Error Resume Next
    Dim topRule As Top10
    Dim botRule As Top10
    
    ' Top 1 Rule
    Set topRule = rng.FormatConditions.AddTop10
    With topRule
        .TopBottom = xlTop10Top
        .Rank = 1
        .Percent = False
        If isInverted Then
            ' If inverted (e.g. Speed), lowest is best
            .Interior.Color = PWC_SOFT_AMBER_BG
            .Font.Color = PWC_DARK_AMBER_TXT
        Else
            .Interior.Color = PWC_SOFT_GREEN_BG
            .Font.Color = PWC_DARK_GREEN_TXT
        End If
        .Font.Bold = True
    End With
    
    ' Bottom 1 Rule
    Set botRule = rng.FormatConditions.AddTop10
    With botRule
        .TopBottom = xlTop10Bottom
        .Rank = 1
        .Percent = False
        If isInverted Then
            ' If lowest seconds, that's best
            .Interior.Color = PWC_SOFT_GREEN_BG
            .Font.Color = PWC_DARK_GREEN_TXT
        Else
            .Interior.Color = PWC_SOFT_RED_BG
            .Font.Color = PWC_DARK_RED_TXT
        End If
        .Font.Bold = True
    End With
    On Error GoTo 0
End Sub

' ==============================================================================
' 5. LIVE LINKED PICTURE DOCKER: CONVERTS PIVOTTABLE TO FLOATING SAAS CARD
' ==============================================================================
Public Sub DockScorecardAsLinkedPicture(Optional ByVal themeChoice As String = "SLA_EXCEPTIONS")
    Dim wsDash As Worksheet
    Dim wsStaging As Worksheet
    Dim pt As PivotTable
    Dim rngTable As Range
    Dim shpOldPic As Shape
    Dim shpDockZone As Shape
    Dim picObj As Picture
    Dim dockLeft As Single, dockTop As Single, dockW As Single, dockH As Single
    
    On Error Resume Next
    Set wsDash = ActiveWorkbook.Worksheets("03_CallCenter_Cockpit")
    Set wsStaging = ActiveWorkbook.Worksheets("Staging_Pivots")
    If wsDash Is Nothing Then Set wsDash = ActiveSheet
    
    ' 1. Find PivotTable (Check Staging_Pivots first, then Dashboard)
    If Not wsStaging Is Nothing Then
        If wsStaging.PivotTables.Count > 0 Then
            Set pt = wsStaging.PivotTables(wsStaging.PivotTables.Count)
        End If
    End If
    If pt Is Nothing Then
        If wsDash.PivotTables.Count > 0 Then
            Set pt = wsDash.PivotTables(wsDash.PivotTables.Count)
        End If
    End If
    
    If pt Is Nothing Then
        MsgBox "Could not find the Agent Scorecard PivotTable." & vbCrLf & _
               "Please ensure it is created on 'Staging_Pivots' or '03_CallCenter_Cockpit'.", _
               vbExclamation, "PwC Scorecard Docker"
        Exit Sub
    End If
    
    ' 2. Style and Format the PivotTable first
    FormatScorecardGrid pt
    ApplyScorecardConditionalFormatting pt, UCase(Trim(themeChoice))
    
    ' 3. Get PivotTable Range
    Set rngTable = pt.TableRange2
    If rngTable Is Nothing Then Set rngTable = pt.TableRange1
    
    ' 4. Coordinates of Visual Container CC_AgentScorecard
    dockLeft = 776
    dockTop = 558
    dockW = 448
    dockH = 210
    
    Set shpDockZone = wsDash.Shapes("DockZone_CC_AgentScorecard")
    If Not shpDockZone Is Nothing Then
        dockLeft = shpDockZone.Left
        dockTop = shpDockZone.Top
        dockW = shpDockZone.Width
        dockH = shpDockZone.Height
        ' Hide the dashed placeholder watermark
        shpDockZone.Visible = msoFalse
    End If
    
    ' 5. Remove any previously created Linked Picture
    Set shpOldPic = wsDash.Shapes("LinkedPic_AgentScorecard")
    If Not shpOldPic Is Nothing Then shpOldPic.Delete
    
    ' 6. Copy PivotTable Range
    rngTable.Copy
    
    ' 7. Paste as Linked Picture on Dashboard
    wsDash.Activate
    wsDash.Range("A1").Select
    Set picObj = wsDash.Pictures.Paste(Link:=True)
    
    If Not picObj Is Nothing Then
        With picObj
            .Name = "LinkedPic_AgentScorecard"
            .Left = dockLeft + (dockW - .Width) / 2
            If .Left < dockLeft Then .Left = dockLeft + 6
            .Top = dockTop + 4
            .ShapeRange.LockAspectRatio = msoTrue
            If .Height > (dockH - 12) Then
                .Height = dockH - 12
            End If
            If .Width > (dockW - 12) Then
                .Width = dockW - 12
            End If
            .Left = dockLeft + (dockW - .Width) / 2
            .Top = dockTop + (dockH - .Height) / 2
        End With
    End If
    
    Application.CutCopyMode = False
    
    If Application.UserControl Then
        MsgBox "Agent Scorecard successfully docked as a live Linked Picture inside the container!", _
               vbInformation, "PwC Scorecard Docker"
    End If
    On Error GoTo 0
End Sub

