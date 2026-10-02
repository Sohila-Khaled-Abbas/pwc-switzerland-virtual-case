Attribute VB_Name = "modInteractiveScorecard"
Option Explicit

' ==============================================================================
' PwC Switzerland Digital Accelerator -- Interactive HTML Scorecard Docker
' Docks the Agent Scorecard PivotTable INSIDE the Dotted Box as a Live SaaS Widget
' With Modern HTML/CSS Visual Styling (Dark Header, Zebra Rows, Status Badges)
' ==============================================================================
'
' Key Capabilities:
'   1. BuildAndDockInteractiveScorecard:
'      - Preserves the dotted box (DockZone_CC_AgentScorecard) as the outer frame
'      - Clears the watermark placeholder text
'      - Styles the underlying PivotTable (pt_Agent) with SaaS HTML table aesthetics
'      - Inserts the live linked table INSIDE the dotted box with elegant padding
'      - 100% LIVE INTERACTIVE with all dashboard slicers in real-time!
'   2. FormatPivotTableHTMLTheme:
'      - HTML/CSS Table Design: #0F172A dark header, white bold text
'      - Alternating row zebra striping (#FFFFFF vs #F8FAFC)
'      - Soft pill badges: Green (#DCFCE7/#166534) for SLA pass, Red (#FEE2E2/#991B1B) for breach
'      - Inverted speed triage: Green for <= 66s, Amber (#FEF3C7/#92400E) for > 70s
'      - Star rating highlight for CSAT >= 3.45
'   3. ExportScorecardToHTMLFile:
'      - Generates a standalone, fully styled HTML5 report with embedded modern CSS
'
' ==============================================================================

' ------------------------------------------------------------------------------
' Brand Color Constants (RGB Values)
' ------------------------------------------------------------------------------
Private Const PWC_TANGERINE        As Long = 133288     ' #D04A02 - RGB(208, 74, 2)
Private Const PWC_DARK_SLATE       As Long = 2758415    ' #0F172A - RGB(15, 23, 42)
Private Const PWC_CHARCOAL          As Long = 3877150    ' #1E293B - RGB(30, 41, 59)
Private Const PWC_ZEBRA_BG          As Long = 16579320   ' #F8FAFC - RGB(248, 250, 252)
Private Const PWC_WHITE             As Long = 16777215   ' #FFFFFF - RGB(255, 255, 255)
Private Const PWC_BORDER_GRAY       As Long = 15790322   ' #E2E8F0 - RGB(226, 232, 240)
Private Const PWC_DOTTED_BORDER     As Long = 13421772   ' #CBD5E1 - RGB(203, 213, 225)

' HTML Status Badge Colors
Private Const PWC_BADGE_GREEN_BG    As Long = 13761756   ' #DCFCE7 - RGB(220, 252, 231)
Private Const PWC_BADGE_GREEN_TXT   As Long = 3433746    ' #166534 - RGB(22, 101, 52)
Private Const PWC_BADGE_RED_BG      As Long = 14803454   ' #FEE2E2 - RGB(254, 226, 226)
Private Const PWC_BADGE_RED_TXT     As Long = 1776537    ' #991B1B - RGB(153, 27, 27)
Private Const PWC_BADGE_AMBER_BG    As Long = 13104126   ' #FEF3C7 - RGB(254, 243, 199)
Private Const PWC_BADGE_AMBER_TXT   As Long = 933902     ' #92400E - RGB(146, 64, 14)
Private Const PWC_BADGE_GOLD_BG     As Long = 9105406    ' #FEF08A - RGB(254, 240, 138)
Private Const PWC_BADGE_GOLD_TXT    As Long = 937349     ' #854D0E - RGB(133, 77, 14)

Private Const FONT_FAMILY           As String = "Segoe UI"

' ==============================================================================
' 1. MASTER ENTRY POINT: BUILD & DOCK LIVE INTERACTIVE SCORECARD
' ==============================================================================
Public Sub BuildAndDockInteractiveScorecard()
    Dim wb As Workbook
    Dim wsDash As Worksheet
    Dim wsStaging As Worksheet
    Dim pt As PivotTable
    Dim shpDockZone As Shape
    Dim shpOldPic As Shape
    Dim picObj As Picture
    Dim rngTable As Range
    
    Set wb = ActiveWorkbook
    On Error Resume Next
    Set wsDash = wb.Worksheets("03_CallCenter_Cockpit")
    Set wsStaging = wb.Worksheets("Staging_Pivots")
    On Error GoTo 0
    
    If wsDash Is Nothing Then
        MsgBox "Dashboard sheet '03_CallCenter_Cockpit' not found!", vbCritical, "PwC Scorecard Docker"
        Exit Sub
    End If
    
    If wsStaging Is Nothing Then
        MsgBox "Staging sheet 'Staging_Pivots' not found!", vbCritical, "PwC Scorecard Docker"
        Exit Sub
    End If
    
    ' 1. Locate the PivotTable (pt_Agent or first PivotTable on Staging_Pivots)
    On Error Resume Next
    Set pt = wsStaging.PivotTables("pt_Agent")
    If pt Is Nothing Then
        If wsStaging.PivotTables.Count > 0 Then
            Set pt = wsStaging.PivotTables(1)
        End If
    End If
    On Error GoTo 0
    
    If pt Is Nothing Then
        MsgBox "No PivotTable found on 'Staging_Pivots'!" & vbCrLf & _
               "Please ensure your Agent Scorecard PivotTable is created on Staging_Pivots first.", _
               vbExclamation, "PwC Scorecard Docker"
        Exit Sub
    End If
    
    ' 2. Locate the Dotted Docking Zone on Dashboard
    On Error Resume Next
    Set shpDockZone = wsDash.Shapes("DockZone_CC_AgentScorecard")
    On Error GoTo 0
    
    If shpDockZone Is Nothing Then
        MsgBox "Dotted docking zone 'DockZone_CC_AgentScorecard' not found on dashboard!", _
               vbCritical, "PwC Scorecard Docker"
        Exit Sub
    End If
    
    ' Pause screen flicker
    On Error Resume Next
    modAppState.FreezeAppState
    On Error GoTo 0
    
    ' 3. Apply Modern HTML/CSS Visual Styling to the PivotTable
    FormatPivotTableHTMLTheme pt
    
    ' 4. PRESERVE & CLEAN THE DOTTED BOX (DO NOT REMOVE IT!)
    ' Clear the watermark text and set a clean crisp white background
    With shpDockZone
        .Fill.Solid
        .Fill.ForeColor.RGB = PWC_WHITE
        .Line.ForeColor.RGB = PWC_DOTTED_BORDER
        .Line.Weight = 0.75
        .Line.DashStyle = msoLineDash
        .TextFrame2.TextRange.Text = ""
        .Visible = msoTrue
    End With
    
    ' 5. Remove any previous linked picture if re-running
    On Error Resume Next
    Set shpOldPic = wsDash.Shapes("LiveScorecard_HTMLTable")
    If Not shpOldPic Is Nothing Then shpOldPic.Delete
    On Error GoTo 0
    
    ' 6. Copy the formatted PivotTable Range
    Set rngTable = pt.TableRange2
    If rngTable Is Nothing Then Set rngTable = pt.TableRange1
    rngTable.Copy
    
    ' 7. Paste as Live Linked Picture onto the Dashboard INSIDE the Dotted Box
    wsDash.Activate
    wsDash.Range("A1").Select
    Set picObj = wsDash.Pictures.Paste(Link:=True)
    
    ' 8. Center and Dock the Picture INSIDE the Dotted Box with 8pt Padding
    If Not picObj Is Nothing Then
        With picObj
            .Name = "LiveScorecard_HTMLTable"
            .ShapeRange.LockAspectRatio = msoTrue
            
            ' Available interior space inside the dotted box
            Dim maxW As Single, maxH As Single
            maxW = shpDockZone.Width - 16
            maxH = shpDockZone.Height - 16
            
            ' Scale to fit neatly inside the dotted box boundaries
            If .Width > maxW Then
                .Width = maxW
            End If
            If .Height > maxH Then
                .Height = maxH
            End If
            
            ' Center precisely inside the dotted box
            .Left = shpDockZone.Left + (shpDockZone.Width - .Width) / 2
            .Top = shpDockZone.Top + (shpDockZone.Height - .Height) / 2
            
            ' Bring the linked picture in front of the dotted box
            .ShapeRange.ZOrder msoBringToFront
        End With
    End If
    
    Application.CutCopyMode = False
    
    ' Restore app state
    On Error Resume Next
    modAppState.RestoreAppState
    On Error GoTo 0
    
    If Application.UserControl Then
        MsgBox "Agent Scorecard successfully docked INSIDE the dotted box!" & vbCrLf & _
               "- Dotted box preserved as the outer container frame." & vbCrLf & _
               "- HTML/CSS theme formatting applied." & vbCrLf & _
               "- 100% Live & interactive with all dashboard slicers.", _
               vbInformation, "PwC Interactive Scorecard"
    End If
End Sub

' ==============================================================================
' 2. HTML/CSS TABLE DESIGN ENGINE (APPLIED TO PIVOTTABLE CELLS)
' ==============================================================================
Public Sub FormatPivotTableHTMLTheme(pt As PivotTable)
    Dim pf As PivotField
    Dim rngTable As Range
    Dim rngData As Range
    Dim rngHeader As Range
    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = pt.Parent
    On Error Resume Next
    
    With pt
        .RowAxisLayout xlTabularRow
        .ColumnGrand = False
        .RowGrand = False
        .ShowTableStyleRowStripes = False
        .ShowTableStyleColumnStripes = False
        .TableStyle2 = ""  ' Clear built-in styles to apply pure custom HTML theme
        
        ' Format Captions & Number Formats
        For i = 1 To .DataFields.Count
            Set pf = .DataFields(i)
            Select Case True
                Case InStr(1, pf.Caption, "Call", vbTextCompare) > 0 Or InStr(1, pf.SourceName, "Call", vbTextCompare) > 0
                    pf.Caption = "Calls Taken"
                    pf.NumberFormat = "#,##0"
                Case InStr(1, pf.Caption, "Answer", vbTextCompare) > 0 Or InStr(1, pf.SourceName, "Answer", vbTextCompare) > 0
                    pf.Caption = "Answer Rate %"
                    pf.NumberFormat = "0.0%"
                Case InStr(1, pf.Caption, "Resolution", vbTextCompare) > 0 Or InStr(1, pf.SourceName, "Resolution", vbTextCompare) > 0 Or InStr(1, pf.Caption, "FCR", vbTextCompare) > 0
                    pf.Caption = "FCR Rate %"
                    pf.NumberFormat = "0.0%"
                Case InStr(1, pf.Caption, "Speed", vbTextCompare) > 0 Or InStr(1, pf.SourceName, "Speed", vbTextCompare) > 0
                    pf.Caption = "Avg Speed (s)"
                    pf.NumberFormat = "0.0 ""s"""
                Case InStr(1, pf.Caption, "CSAT", vbTextCompare) > 0 Or InStr(1, pf.SourceName, "CSAT", vbTextCompare) > 0
                    pf.Caption = "Avg CSAT"
                    pf.NumberFormat = "0.00"
            End Select
        Next i
        
        Set rngTable = .TableRange2
        If rngTable Is Nothing Then Set rngTable = .TableRange1
        
        ' 1. Typography & Global Table Formatting
        With rngTable
            .Font.Name = FONT_FAMILY
            .Font.Size = 8.5
            .Font.Color = PWC_CHARCOAL
            .VerticalAlignment = xlVAlignCenter
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Color = PWC_BORDER_GRAY
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = PWC_BORDER_GRAY
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Color = PWC_BORDER_GRAY
            .Borders(xlInsideHorizontal).Weight = xlThin
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeRight).LineStyle = xlNone
            .RowHeight = 18.5
        End With
        
        ' 2. HTML-Style Dark Header Row
        Dim headerRowIdx As Long
        headerRowIdx = pt.TableRange1.Row
        Dim colStart As Long, colCount As Long
        colStart = pt.TableRange1.Column
        colCount = pt.TableRange1.Columns.Count
        
        Set rngHeader = ws.Range(ws.Cells(headerRowIdx, colStart), ws.Cells(headerRowIdx, colStart + colCount - 1))
        With rngHeader
            .Interior.Color = PWC_DARK_SLATE        ' #0F172A Dark Slate Header
            .Font.Color = PWC_WHITE                 ' Crisp White Text
            .Font.Bold = True
            .Font.Size = 8.5
            .RowHeight = 22
            .HorizontalAlignment = xlCenter
        End With
        ' Ensure Agent header is left-aligned
        ws.Cells(headerRowIdx, colStart).HorizontalAlignment = xlLeft
        
        ' 3. HTML-Style Alternating Zebra Row Striping
        Set rngData = pt.DataBodyRange
        If Not rngData Is Nothing Then
            Dim r As Long
            Dim rowRng As Range
            For r = 1 To rngData.Rows.Count
                Set rowRng = ws.Range(ws.Cells(rngData.Rows(r).Row, colStart), _
                                      ws.Cells(rngData.Rows(r).Row, colStart + colCount - 1))
                If r Mod 2 = 0 Then
                    rowRng.Interior.Color = PWC_ZEBRA_BG ' Soft #F8FAFC
                Else
                    rowRng.Interior.Color = PWC_WHITE    ' Crisp #FFFFFF
                End If
                ' Agent name column bold
                ws.Cells(rngData.Rows(r).Row, colStart).Font.Bold = True
                ws.Cells(rngData.Rows(r).Row, colStart).HorizontalAlignment = xlLeft
            Next r
            
            ' 4. Clear and Apply Executive Status Badge Rules
            rngData.FormatConditions.Delete
            ApplyHTMLBadgeFormatting pt
        End If
    End With
    On Error GoTo 0
End Sub

' Helper: Applies Pill-Style Badge Conditional Formatting to Data Fields
Private Sub ApplyHTMLBadgeFormatting(pt As PivotTable)
    Dim pf As PivotField
    Dim fldRng As Range
    Dim i As Long
    Dim fcLow As FormatCondition, fcHigh As FormatCondition
    
    On Error Resume Next
    For i = 1 To pt.DataFields.Count
        Set pf = pt.DataFields(i)
        Set fldRng = pf.DataRange
        
        If Not fldRng Is Nothing Then
            fldRng.HorizontalAlignment = xlRight
            
            Select Case True
                ' Answer Rate %: Green Pill for >= 83%, Red Pill for < 80%
                Case InStr(1, pf.Caption, "Answer", vbTextCompare) > 0
                    Set fcLow = fldRng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0.80")
                    With fcLow
                        .Interior.Color = PWC_BADGE_RED_BG
                        .Font.Color = PWC_BADGE_RED_TXT
                        .Font.Bold = True
                    End With
                    
                    Set fcHigh = fldRng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0.83")
                    With fcHigh
                        .Interior.Color = PWC_BADGE_GREEN_BG
                        .Font.Color = PWC_BADGE_GREEN_TXT
                        .Font.Bold = True
                    End With
                    
                ' Speed of Answer: INVERTED! Green Pill for <= 66s, Amber Pill for > 70s
                Case InStr(1, pf.Caption, "Speed", vbTextCompare) > 0
                    Set fcHigh = fldRng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="70")
                    With fcHigh
                        .Interior.Color = PWC_BADGE_AMBER_BG
                        .Font.Color = PWC_BADGE_AMBER_TXT
                        .Font.Bold = True
                    End With
                    
                    Set fcLow = fldRng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLessEqual, Formula1:="66")
                    With fcLow
                        .Interior.Color = PWC_BADGE_GREEN_BG
                        .Font.Color = PWC_BADGE_GREEN_TXT
                        .Font.Bold = True
                    End With
                    
                ' CSAT: Gold Badge for >= 3.45 (Dan & Martha Stars)
                Case InStr(1, pf.Caption, "CSAT", vbTextCompare) > 0
                    Set fcHigh = fldRng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="3.45")
                    With fcHigh
                        .Interior.Color = PWC_BADGE_GOLD_BG
                        .Font.Color = PWC_BADGE_GOLD_TXT
                        .Font.Bold = True
                    End With
            End Select
        End If
    Next i
    On Error GoTo 0
End Sub

' ==============================================================================
' 3. STANDALONE HTML5 REPORT GENERATOR
' Generates an HTML5 file with embedded CSS table styling
' ==============================================================================
Public Sub ExportScorecardToHTMLFile()
    Dim wsStaging As Worksheet
    Dim pt As PivotTable
    Dim rngTable As Range
    Dim filePath As String
    Dim fNum As Integer
    Dim r As Long, c As Long
    Dim cellVal As String
    Dim html As String
    
    On Error Resume Next
    Set wsStaging = ActiveWorkbook.Worksheets("Staging_Pivots")
    Set pt = wsStaging.PivotTables("pt_Agent")
    If pt Is Nothing Then Set pt = wsStaging.PivotTables(1)
    On Error GoTo 0
    
    If pt Is Nothing Then
        MsgBox "PivotTable not found!", vbCritical
        Exit Sub
    End If
    
    Set rngTable = pt.TableRange1
    
    ' Build Modern HTML5 Table with Embedded CSS
    html = "<!DOCTYPE html>" & vbCrLf & _
           "<html lang='en'><head><meta charset='UTF-8'>" & vbCrLf & _
           "<title>PwC Agent Quality & CSAT Audit</title>" & vbCrLf & _
           "<style>" & vbCrLf & _
           "  body { font-family: 'Segoe UI', -apple-system, sans-serif; background: #F8FAFC; padding: 24px; color: #1E293B; }" & vbCrLf & _
           "  .card { background: white; border-radius: 12px; border: 1px solid #E2E8F0; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); padding: 20px; max-width: 650px; margin: auto; }" & vbCrLf & _
           "  h3 { margin: 0 0 4px 0; font-size: 16px; color: #0F172A; }" & vbCrLf & _
           "  p.sub { margin: 0 0 16px 0; font-size: 12px; color: #64748B; }" & vbCrLf & _
           "  table { width: 100%; border-collapse: collapse; font-size: 12px; }" & vbCrLf & _
           "  th { background: #0F172A; color: white; padding: 8px 12px; text-align: right; font-weight: 600; text-transform: uppercase; font-size: 10px; letter-spacing: 0.5px; }" & vbCrLf & _
           "  th:first-child { text-align: left; border-top-left-radius: 6px; }" & vbCrLf & _
           "  th:last-child { border-top-right-radius: 6px; }" & vbCrLf & _
           "  td { padding: 8px 12px; border-bottom: 1px solid #E2E8F0; text-align: right; }" & vbCrLf & _
           "  td:first-child { text-align: left; font-weight: 600; color: #0F172A; }" & vbCrLf & _
           "  tr:nth-child(even) { background-color: #F8FAFC; }" & vbCrLf & _
           "  tr:hover { background-color: #F1F5F9; transition: background 0.15s; }" & vbCrLf & _
           "  .badge-green { background: #DCFCE7; color: #166534; padding: 2px 8px; border-radius: 12px; font-weight: 600; }" & vbCrLf & _
           "  .badge-red { background: #FEE2E2; color: #991B1B; padding: 2px 8px; border-radius: 12px; font-weight: 600; }" & vbCrLf & _
           "  .badge-amber { background: #FEF3C7; color: #92400E; padding: 2px 8px; border-radius: 12px; font-weight: 600; }" & vbCrLf & _
           "  .badge-gold { background: #FEF08A; color: #854D0E; padding: 2px 8px; border-radius: 12px; font-weight: 600; }" & vbCrLf & _
           "</style></head><body>" & vbCrLf & _
           "<div class='card'>" & vbCrLf & _
           "  <h3>Representative Quality &amp; CSAT Audit</h3>" & vbCrLf & _
           "  <p class='sub'>Frontline Performance Metrics &amp; Service Level Benchmarks</p>" & vbCrLf & _
           "  <table>" & vbCrLf
           
    ' Header Row
    html = html & "    <thead><tr>" & vbCrLf
    For c = 1 To rngTable.Columns.Count
        html = html & "      <th>" & rngTable.Cells(1, c).Text & "</th>" & vbCrLf
    Next c
    html = html & "    </tr></thead><tbody>" & vbCrLf
    
    ' Data Rows
    For r = 2 To rngTable.Rows.Count
        html = html & "    <tr>" & vbCrLf
        For c = 1 To rngTable.Columns.Count
            cellVal = rngTable.Cells(r, c).Text
            If c = 1 Then
                html = html & "      <td>" & cellVal & "</td>" & vbCrLf
            ElseIf c = 3 Then ' Answer Rate
                If Val(Replace(cellVal, "%", "")) < 80# Then
                    html = html & "      <td><span class='badge-red'>" & cellVal & "</span></td>" & vbCrLf
                ElseIf Val(Replace(cellVal, "%", "")) >= 83# Then
                    html = html & "      <td><span class='badge-green'>" & cellVal & "</span></td>" & vbCrLf
                Else
                    html = html & "      <td>" & cellVal & "</td>" & vbCrLf
                End If
            ElseIf c = 5 Then ' Speed
                If Val(Replace(cellVal, "s", "")) > 70# Then
                    html = html & "      <td><span class='badge-amber'>" & cellVal & "</span></td>" & vbCrLf
                ElseIf Val(Replace(cellVal, "s", "")) <= 66# Then
                    html = html & "      <td><span class='badge-green'>" & cellVal & "</span></td>" & vbCrLf
                Else
                    html = html & "      <td>" & cellVal & "</td>" & vbCrLf
                End If
            ElseIf c = 6 Then ' CSAT
                If Val(cellVal) >= 3.45 Then
                    html = html & "      <td><span class='badge-gold'>" & cellVal & "</span></td>" & vbCrLf
                Else
                    html = html & "      <td>" & cellVal & "</td>" & vbCrLf
                End If
            Else
                html = html & "      <td>" & cellVal & "</td>" & vbCrLf
            End If
        Next c
        html = html & "    </tr>" & vbCrLf
    Next r
    
    html = html & "  </tbody></table>" & vbCrLf & "</div></body></html>"
    
    filePath = ActiveWorkbook.Path & "\agent_scorecard.html"
    fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, html
    Close #fNum
    
    If Application.UserControl Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("HTML Scorecard exported successfully to:" & vbCrLf & filePath & vbCrLf & vbCrLf & _
                      "Would you like to open it in your web browser now?", vbQuestion + vbYesNo, "PwC HTML Exporter")
        If resp = vbYes Then
            Dim shellObj As Object
            Set shellObj = CreateObject("WScript.Shell")
            shellObj.Run Chr(34) & filePath & Chr(34)
        End If
    End If
End Sub
