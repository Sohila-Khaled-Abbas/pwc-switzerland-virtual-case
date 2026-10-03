Attribute VB_Name = "modCreateGovernanceSheets"
Option Explicit

' ==============================================================================
' PwC Switzerland Virtual Case Experience - Pre-Dashboard Governance Builder
' Module: modCreateGovernanceSheets.bas
' Purpose: Automatically generates, populates, and professionally styles:
'          1. "01_Business_Domains" (Business domain briefings & strategic context)
'          2. "02_Metadata_&_KPI_Catalog" (Metadata inventory & KPI dictionary)
' Brand Palette:
'   - Primary Dark Slate: RGB(30, 41, 59)   [#1E293B]
'   - PwC Tangerine:     RGB(208, 74, 2)   [#D04A02]
'   - Muted Slate:       RGB(100, 116, 139) [#64748B]
'   - Zebra Tint:        RGB(248, 250, 252) [#F8FAFC]
'   - Card Border:       RGB(203, 213, 225) [#CBD5E1]
' Placement: Run immediately after Data Model Relationships & Power Pivot setup,
'            BEFORE creating explicit DAX measures and dashboard views.
' ==============================================================================

Public Sub BuildGovernanceArchitecture()
    Dim wb As Workbook
    Dim wsDomains As Worksheet
    Dim wsCatalog As Worksheet
    Dim s As Worksheet
    Dim lo As ListObject
    
    Set wb = ActiveWorkbook
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    ' 1. Pre-emptively unlist existing tables to prevent global namespace collisions
    For Each s In wb.Worksheets
        For Each lo In s.ListObjects
            If StrComp(lo.Name, "tbl_Metadata_Catalog", vbTextCompare) = 0 Or _
               StrComp(lo.Name, "tbl_KPI_Dictionary", vbTextCompare) = 0 Then
                lo.Unlist
            End If
        Next lo
    Next s
    
    ' 2. Reset / recreate target sheets cleanly
    On Error Resume Next
    Set wsDomains = wb.Worksheets("01_Business_Domains")
    If Not wsDomains Is Nothing Then wsDomains.Delete
    
    Set wsCatalog = wb.Worksheets("02_Metadata_&_KPI_Catalog")
    If Not wsCatalog Is Nothing Then wsCatalog.Delete
    On Error GoTo ErrorHandler
    
    ' 3. Insert sheets at the beginning of the workbook
    Set wsDomains = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    wsDomains.Name = "01_Business_Domains"
    
    Set wsCatalog = wb.Worksheets.Add(After:=wsDomains)
    wsCatalog.Name = "02_Metadata_&_KPI_Catalog"
    
    ' 4. Populate and style Domains Sheet
    Call PopulateDomainsSheet(wsDomains)
    
    ' 5. Populate and style Metadata & KPI Sheet
    Call PopulateCatalogSheet(wsCatalog)
    
    ' 6. Return focus to Domains overview
    wsDomains.Activate
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "PwC Governance Architecture successfully generated!" & vbCrLf & vbCrLf & _
           "• '01_Business_Domains': 3 Structured Executive Briefing Cards created" & vbCrLf & _
           "• '02_Metadata_&_KPI_Catalog': 12 Model Tables & 10 Core KPIs registered" & vbCrLf & vbCrLf & _
           "All formatting, column widths, and brand colors auto-fitted successfully." & vbCrLf & _
           "Next Step: Add explicit DAX measures to VertiPaq model.", _
           vbInformation, "PwC Switzerland BI Architecture"
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Private Sub PopulateDomainsSheet(ws As Worksheet)
    ws.Tab.Color = RGB(208, 74, 2) ' PwC Tangerine
    
    ' Activate sheet before setting Window properties (DisplayGridlines belongs to ActiveWindow)
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    
    ' Explicit comfortable column widths to prevent card clipping
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 54  ' Domain Card 1
    ws.Columns("C").ColumnWidth = 3   ' Spacer Gap
    ws.Columns("D").ColumnWidth = 54  ' Domain Card 2
    ws.Columns("E").ColumnWidth = 3   ' Spacer Gap
    ws.Columns("F").ColumnWidth = 54  ' Domain Card 3
    ws.Columns("G").ColumnWidth = 3   ' Right Margin
    
    ' Title Ribbon Banner across B2:F3
    ws.Range("B2:F3").Merge
    With ws.Range("B2")
        .Value = "PwC Switzerland | Client Business Domains & Strategic Context"
        .Font.Name = "Segoe UI"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(30, 41, 59)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(2).RowHeight = 22
    ws.Rows(3).RowHeight = 22
    
    ' Accent Line (PwC Tangerine) across B4:F4
    ws.Range("B4:F4").Interior.Color = RGB(208, 74, 2)
    ws.Rows(4).RowHeight = 4
    
    ' Subtitle across B5:F5
    ws.Range("B5:F5").Merge
    With ws.Range("B5")
        .Value = "Strategic context, business models, operational grains, and critical risk taxonomies for the 3 client modules."
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Italic = True
        .Font.Color = RGB(100, 116, 139)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(5).RowHeight = 22
    
    ' Navigation Jump Link at B6
    ws.Range("B6").Value = "--> Open Metadata & Executive KPI Catalog >>"
    ws.Hyperlinks.Add Anchor:=ws.Range("B6"), Address:="", SubAddress:="'02_Metadata_&_KPI_Catalog'!A1", TextToDisplay:="--> Open Metadata & Executive KPI Catalog >>"
    With ws.Range("B6")
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(208, 74, 2)
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(6).RowHeight = 22
    ws.Rows(7).RowHeight = 10 ' Spacer row
    
    ' Array of bullet points for Card 1 (Call Center)
    Dim b1 As Variant
    b1 = Array( _
        "- Source: 01 Call-Center-Dataset.xlsx (5,000 calls)", _
        "- Granularity: 1 Inbound Customer Call Attempt", _
        "- Core Mission: Balance throughput with CSAT & FCR", _
        "- Strategic Risk: 946 Abandoned Callers (Queue drops)", _
        "- Peak Window: 11:00 AM - 2:00 PM (62% Abandonment)", _
        "- Key Solution: Shift 2 morning agents to midday triage" _
    )
    Call DrawStructuredCard(ws, "B", 1, "CALL CENTER OPERATIONS", _
        "Telecommunications Customer Care Hub", b1, _
        "Abandonment < 15.0% | CSAT > 3.50", "03_CallCenter_Cockpit")
        
    ' Array of bullet points for Card 2 (Customer Retention)
    Dim b2 As Variant
    b2 = Array( _
        "- Source: 02 Churn-Dataset.xlsx (7,043 accounts)", _
        "- Granularity: 1 Customer Subscription Profile", _
        "- Core Mission: Protect Annual Recurring Revenue ($2.86M Risk)", _
        "- High-Risk Cohort: Month-to-Month Contracts (42.7% Churn)", _
        "- Friction Triggers: Electronic Check + Fiber latency", _
        "- Key Solution: 1-Year lock-in + AutoPay discount" _
    )
    Call DrawStructuredCard(ws, "D", 2, "CUSTOMER CHURN & RETENTION", _
        "Subscription Broadband & Telephony Operator", b2, _
        "Churn Rate < 20.0% | ARR Risk < $2.0M", "04_CustomerRetention_Cockpit")
        
    ' Array of bullet points for Card 3 (Diversity & Inclusion)
    Dim b3 As Variant
    b3 = Array( _
        "- Source: 03 Diversity-Inclusion-Dataset.xlsx (500 staff)", _
        "- Granularity: 1 Corporate Employee Career Snapshot", _
        "- Core Mission: Gender parity across executive ranks", _
        "- The Broken Rung: Manager (34.3% F) -> Sr Mgr (14.6% F)", _
        "- Promotion Lag: Women average +5.6 months in grade", _
        "- Key Solution: Executive sponsorship & PRA pay audits" _
    )
    Call DrawStructuredCard(ws, "F", 3, "DIVERSITY & INCLUSION", _
        "Pharma Group AG (Swiss Enterprise)", b3, _
        "Broken Rung Gap < 5.0% | Board Parity > 35%", "05_DiversityInclusion_Cockpit")
        
    ' Set fixed comfortable row heights for cards
    ws.Rows(8).RowHeight = 32  ' Card Header
    ws.Rows(9).RowHeight = 24  ' Client Subtitle
    Dim i As Long
    For i = 10 To 15
        ws.Rows(i).RowHeight = 22 ' Bullet Rows
    Next i
    ws.Rows(16).RowHeight = 26 ' Target KPI
    ws.Rows(17).RowHeight = 28 ' Action Button
End Sub

Private Sub DrawStructuredCard(ws As Worksheet, colL As String, cardNum As Long, _
                                title As String, clientName As String, bullets As Variant, _
                                kpiSummary As String, targetSheet As String)
    Dim r As Long
    
    ' 1. Card Header Row 8
    With ws.Range(colL & "8")
        .Value = "[" & Format(cardNum, "00") & "] " & title
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(30, 41, 59)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeTop).Color = RGB(208, 74, 2)
        .Borders(xlEdgeTop).Weight = xlMedium
    End With
    
    ' 2. Client Division Subtitle Row 9
    With ws.Range(colL & "9")
        .Value = "Client: " & clientName
        .Font.Name = "Segoe UI"
        .Font.Size = 9.5
        .Font.Bold = True
        .Font.Color = RGB(30, 41, 59)
        .Interior.Color = RGB(241, 245, 249)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Borders(xlEdgeBottom).Color = RGB(226, 232, 240)
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    
    ' 3. Bullet Point Rows (Rows 10 to 15) - Each bullet in its own distinct row!
    For r = 0 To UBound(bullets)
        With ws.Range(colL & (10 + r))
            .Value = bullets(r)
            .Font.Name = "Segoe UI"
            .Font.Size = 9.5
            .Font.Color = RGB(51, 65, 85)
            .WrapText = True
            If r Mod 2 = 0 Then
                .Interior.Color = RGB(255, 255, 255)
            Else
                .Interior.Color = RGB(248, 250, 252)
            End If
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .IndentLevel = 1
            .Borders(xlEdgeBottom).Color = RGB(241, 245, 249)
            .Borders(xlEdgeBottom).Weight = xlHairline
        End With
    Next r
    
    ' 4. Summary Target KPI Row 16
    With ws.Range(colL & "16")
        .Value = "Benchmark: " & kpiSummary
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = RGB(208, 74, 2)
        .Interior.Color = RGB(255, 247, 237)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Borders(xlEdgeTop).Color = RGB(254, 215, 170)
        .Borders(xlEdgeTop).Weight = xlThin
    End With
    
    ' 5. Action Button Row 17
    With ws.Range(colL & "17")
        .Value = "[ View Dashboard ]"
        ws.Hyperlinks.Add Anchor:=ws.Range(colL & "17"), Address:="", SubAddress:="'" & targetSheet & "'!A1", TextToDisplay:="[ View Dashboard ]"
        .Font.Name = "Segoe UI"
        .Font.Size = 9.5
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(30, 41, 59)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' 6. Card Outer Box Border
    Dim cardRng As Range
    Set cardRng = ws.Range(colL & "8:" & colL & "17")
    cardRng.Borders(xlEdgeLeft).Color = RGB(203, 213, 225)
    cardRng.Borders(xlEdgeLeft).Weight = xlThin
    cardRng.Borders(xlEdgeRight).Color = RGB(203, 213, 225)
    cardRng.Borders(xlEdgeRight).Weight = xlThin
    cardRng.Borders(xlEdgeBottom).Color = RGB(208, 74, 2)
    cardRng.Borders(xlEdgeBottom).Weight = xlMedium
End Sub

Private Sub PopulateCatalogSheet(ws As Worksheet)
    ws.Tab.Color = RGB(30, 41, 59) ' PwC Charcoal
    
    ' Activate sheet before setting Window properties (DisplayGridlines belongs to ActiveWindow)
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    
    ' Title Ribbon Banner across B2:J3
    ws.Range("B2:J3").Merge
    With ws.Range("B2")
        .Value = "PwC Switzerland | Enterprise Metadata & KPI Governance Catalog"
        .Font.Name = "Segoe UI"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(30, 41, 59)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(2).RowHeight = 22
    ws.Rows(3).RowHeight = 22
    
    ' Accent Line (PwC Tangerine) across B4:J4
    ws.Range("B4:J4").Interior.Color = RGB(208, 74, 2)
    ws.Rows(4).RowHeight = 4
    
    ' Navigation Back Link at B5
    ws.Range("B5").Value = "<-- Back to Business Domains"
    ws.Hyperlinks.Add Anchor:=ws.Range("B5"), Address:="", SubAddress:="'01_Business_Domains'!A1", TextToDisplay:="<-- Back to Business Domains"
    With ws.Range("B5")
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(208, 74, 2)
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(5).RowHeight = 22
    ws.Rows(6).RowHeight = 90 ' Spacer row to clear SaaS hero banner
    
    ' ==========================================================================
    ' SECTION 1: DATA MODEL ASSETS & METADATA INVENTORY (12 TABLES)
    ' ==========================================================================
    ws.Range("B7:J7").Merge
    With ws.Range("B7")
        .Value = "  [SECTION 1] DATA MODEL ASSETS & METADATA INVENTORY (12 TABLES)"
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(30, 41, 59)
        .Interior.Color = RGB(241, 245, 249)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeLeft).Color = RGB(208, 74, 2)
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeBottom).Color = RGB(226, 232, 240)
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    ws.Rows(7).RowHeight = 28
    
    Dim tblMetaHeaders As Variant
    tblMetaHeaders = Array("Entity_ID", "Business_Domain", "Table_Name", "Table_Type", "Source_File_Sheet", "Row_Count", "Granularity_Definition", "Primary_Key", "Business_Owner")
    
    Dim metaData As Variant
    metaData = Array( _
        Array("ENT-01", "Contact Center", "Fact_Calls", "Fact", "01 Call-Center-Dataset.xlsx", 5000, "1 Inbound Customer Call Attempt", "Call_Id", "Head of Customer Care"), _
        Array("ENT-02", "Customer Retention", "Fact_Churn", "Fact", "02 Churn-Dataset.xlsx", 7043, "1 Customer Subscription Account", "customerID", "VP Customer Success"), _
        Array("ENT-03", "Workforce Governance", "Fact_Employees", "Fact", "03 Diversity-Inclusion.xlsx", 500, "1 Corporate Employee Profile", "Employee_ID", "Chief People Officer"), _
        Array("ENT-04", "Enterprise Calendar", "DimDate", "Dimension", "M Calendar Generator", 90, "1 Calendar Day (Q1 2021)", "Date", "Enterprise BI Lead"), _
        Array("ENT-05", "Contact Center", "DimAgent", "Dimension", "Derived from Fact_Calls", 8, "1 Contact Center Agent", "Agent", "Contact Center Manager"), _
        Array("ENT-06", "Contact Center", "DimTopic", "Dimension", "Derived from Fact_Calls", 5, "1 Inquiry Topic / SLA Tier", "Topic", "QA Director"), _
        Array("ENT-07", "Customer Retention", "DimContract", "Dimension", "Derived from Fact_Churn", 3, "1 Commitment Term Tier", "Contract", "Head of Pricing"), _
        Array("ENT-08", "Workforce Governance", "DimDepartment", "Dimension", "Derived from Fact_Employees", 6, "1 Organizational Unit", "Department", "CHRO"), _
        Array("ENT-09", "Auxiliary Census", "Dim_EmployeeCensus", "Auxiliary", "Backing 1", 500, "1 Benchmark Employee Record", "Employee ID", "Macro Analytics Lead"), _
        Array("ENT-10", "Career Pathways", "Dim_CareerLadder", "Auxiliary", "Backing 2", 5, "Grade Promotion Pair", "Base_Job_Level", "Talent Management"), _
        Array("ENT-11", "Regulatory Parity", "Dim_NationalityCensus", "Auxiliary", "Backing 3", 21, "Country Representation Bucket", "Country_ID", "Corporate Compliance"), _
        Array("ENT-12", "Performance Quotas", "Dim_PRA_Equity", "Auxiliary", "Backing 4", 6, "Department Headcount Matrix", "Department", "Compensation Committee") _
    )
    
    Dim r As Long, c As Long
    For c = 0 To UBound(tblMetaHeaders)
        ws.Cells(8, c + 2).Value = tblMetaHeaders(c)
    Next c
    
    For r = 0 To UBound(metaData)
        For c = 0 To UBound(metaData(r))
            ws.Cells(9 + r, c + 2).Value = metaData(r)(c)
        Next c
    Next r
    
    Dim rngMeta As Range
    Set rngMeta = ws.Range(ws.Cells(8, 2), ws.Cells(8 + UBound(metaData) + 1, 2 + UBound(tblMetaHeaders)))
    Dim loMeta As ListObject
    Set loMeta = ws.ListObjects.Add(xlSrcRange, rngMeta, , xlYes)
    loMeta.Name = "tbl_Metadata_Catalog"
    loMeta.TableStyle = "" ' Remove default cyan style; apply bespoke PwC branding
    
    ' Style Table 1 with PwC Brand Colors
    Call ApplyPwCBrandTableStyle(loMeta, ws)
    
    ' Apply Table_Type Category Badges to Table 1
    Dim cellType As Range
    For Each cellType In loMeta.ListColumns("Table_Type").DataBodyRange
        If StrComp(cellType.Value, "Fact", vbTextCompare) = 0 Then
            cellType.Interior.Color = RGB(241, 245, 249) ' Slate wash
            cellType.Font.Color = RGB(30, 41, 59)
            cellType.Font.Bold = True
        ElseIf StrComp(cellType.Value, "Dimension", vbTextCompare) = 0 Then
            cellType.Interior.Color = RGB(255, 247, 237) ' Tangerine soft wash
            cellType.Font.Color = RGB(194, 65, 12)
            cellType.Font.Bold = True
        Else
            cellType.Interior.Color = RGB(248, 250, 252) ' Neutral tint
            cellType.Font.Color = RGB(100, 116, 139)
        End If
    Next cellType
    
    ' ==========================================================================
    ' SECTION 2: EXECUTIVE KPI GOVERNANCE & METRIC DICTIONARY (10 CORE KPIS)
    ' ==========================================================================
    Dim kpiStartRow As Long
    kpiStartRow = 9 + UBound(metaData) + 4
    
    ws.Range(ws.Cells(kpiStartRow, 2), ws.Cells(kpiStartRow, 10)).Merge
    With ws.Cells(kpiStartRow, 2)
        .Value = "  [SECTION 2] EXECUTIVE KPI GOVERNANCE & METRIC DICTIONARY (10 CORE KPIS)"
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(30, 41, 59)
        .Interior.Color = RGB(241, 245, 249)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeLeft).Color = RGB(208, 74, 2)
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeBottom).Color = RGB(226, 232, 240)
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    ws.Rows(kpiStartRow).RowHeight = 28
    
    Dim tblKPIHeaders As Variant
    tblKPIHeaders = Array("Metric_ID", "Domain", "KPI_Name", "Strategic_Objective", "Plain_Language_Definition", "Target", "Polarity", "Alert_Threshold", "Executive_Consumer")
    
    Dim kpiData As Variant
    kpiData = Array( _
        Array("KPI-CC-01", "Contact Center", "Abandonment Rate", "Queue Optimization", "% of callers who hung up before connecting", "< 15.0%", "Lower (Down)", "> 20.0%", "Head of Customer Care"), _
        Array("KPI-CC-02", "Contact Center", "Speed of Answer", "Response Time", "Average wait time in seconds for answered calls", "< 60 sec", "Lower (Down)", "> 75 sec", "Operations Director"), _
        Array("KPI-CC-03", "Contact Center", "First Contact Resolution", "Service Quality", "% of answered calls resolved on initial attempt", "> 85.0%", "Higher (Up)", "< 80.0%", "QA Director"), _
        Array("KPI-CC-04", "Contact Center", "CSAT Rating", "Customer Delight", "Average satisfaction score on 1.0 to 5.0 scale", "> 3.50", "Higher (Up)", "< 3.20", "Chief Customer Officer"), _
        Array("KPI-CH-01", "Customer Retention", "Customer Churn %", "Subscriber Protection", "Proportion of accounts terminating contract", "< 20.0%", "Lower (Down)", "> 25.0%", "VP Customer Success"), _
        Array("KPI-CH-02", "Customer Retention", "Annual Revenue at Risk", "Revenue Protection", "Annualized recurring revenue lost to churn", "< $2.0M", "Lower (Down)", "> $2.5M", "Chief Financial Officer"), _
        Array("KPI-CH-03", "Customer Retention", "Month-to-Month Churn", "Contract Health", "% of churn events on flexible month-to-month", "< 75.0%", "Lower (Down)", "> 85.0%", "Head of Pricing"), _
        Array("KPI-DI-01", "Workforce Parity", "Broken Rung Gap", "Pipeline Equity", "Drop in female representation from Mgr to Sr Mgr", "< 5.0%", "Lower (Down)", "> 15.0%", "Chief People Officer"), _
        Array("KPI-DI-02", "Workforce Parity", "Executive Parity %", "Leadership Diversity", "Female ratio at Level 1 Executive rank", "> 35.0%", "Higher (Up)", "< 25.0%", "Board of Directors"), _
        Array("KPI-DI-03", "Workforce Parity", "Promotion Velocity Lag", "Tenure Equity", "Time-in-grade gap between female and male promo", "0.0 yrs", "Lower (Down)", "> 0.3 yrs", "Compensation Committee") _
    )
    
    For c = 0 To UBound(tblKPIHeaders)
        ws.Cells(kpiStartRow + 1, c + 2).Value = tblKPIHeaders(c)
    Next c
    
    For r = 0 To UBound(kpiData)
        For c = 0 To UBound(kpiData(r))
            ws.Cells(kpiStartRow + 2 + r, c + 2).Value = kpiData(r)(c)
        Next c
    Next r
    
    Dim rngKPI As Range
    Set rngKPI = ws.Range(ws.Cells(kpiStartRow + 1, 2), ws.Cells(kpiStartRow + 1 + UBound(kpiData) + 1, 2 + UBound(tblKPIHeaders)))
    Dim loKPI As ListObject
    Set loKPI = ws.ListObjects.Add(xlSrcRange, rngKPI, , xlYes)
    loKPI.Name = "tbl_KPI_Dictionary"
    loKPI.TableStyle = "" ' Remove default cyan style; apply bespoke PwC branding
    
    ' Style Table 2 with PwC Brand Colors
    Call ApplyPwCBrandTableStyle(loKPI, ws)
    
    ' Apply Polarity Badges to Table 2
    Dim cellPolarity As Range
    For Each cellPolarity In loKPI.ListColumns("Polarity").DataBodyRange
        If InStr(1, cellPolarity.Value, "Higher", vbTextCompare) > 0 Then
            cellPolarity.Interior.Color = RGB(236, 253, 245) ' Soft Emerald
            cellPolarity.Font.Color = RGB(6, 95, 70)
            cellPolarity.Font.Bold = True
        ElseIf InStr(1, cellPolarity.Value, "Lower", vbTextCompare) > 0 Then
            cellPolarity.Interior.Color = RGB(254, 243, 199) ' Soft Amber
            cellPolarity.Font.Color = RGB(146, 64, 14)
            cellPolarity.Font.Bold = True
        End If
    Next cellPolarity
    
    ' ==========================================================================
    ' INTELLIGENT AUTO-FIT WITH PADDING & DYNAMIC MINIMUM COLUMN WIDTHS
    ' ==========================================================================
    ws.Range("B:J").Columns.AutoFit
    
    ' Add generous padding (+4.5) to every column so Table Filter Dropdowns never clip text
    Dim col As Range
    For Each col In ws.Range("B:J").Columns
        col.ColumnWidth = col.ColumnWidth + 4.5
    Next col
    
    ' Enforce minimum column width floors (never shrink if content needed more room)
    ' Col B (2): Entity_ID / Metric_ID            -> Min 14
    ' Col C (3): Business_Domain / Domain          -> Min 24
    ' Col D (4): Table_Name / KPI_Name             -> Min 26
    ' Col E (5): Table_Type / Strategic_Objective  -> Min 24
    ' Col F (6): Source / Plain_Language_Def       -> Min 52
    ' Col G (7): Row_Count / Target                -> Min 16
    ' Col H (8): Granularity_Def / Polarity        -> Min 36
    ' Col I (9): Primary_Key / Alert_Threshold     -> Min 18
    ' Col J (10): Business_Owner / Exec_Consumer   -> Min 28
    Dim minColWidths As Variant
    minColWidths = Array(14, 24, 26, 24, 52, 16, 36, 18, 28)
    
    Dim cIdx As Long
    For cIdx = 2 To 10
        If ws.Columns(cIdx).ColumnWidth < minColWidths(cIdx - 2) Then
            ws.Columns(cIdx).ColumnWidth = minColWidths(cIdx - 2)
        End If
    Next cIdx
    
    ws.Columns("A").ColumnWidth = 3   ' Left margin spacer
    
    ' Ensure natural clean vertical scrolling across both tables
    ws.Activate
End Sub

Private Sub ApplyPwCBrandTableStyle(lo As ListObject, ws As Worksheet)
    Dim cell As Range
    Dim r As Long
    
    ' 1. Header Row Formatting (PwC Charcoal + White Bold + Tangerine bottom border)
    With lo.HeaderRowRange
        .Interior.Color = RGB(30, 41, 59)
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .RowHeight = 26
        .Borders(xlEdgeBottom).Color = RGB(208, 74, 2)
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    ' 2. Data Rows Formatting (Alternating Zebra Striping + Subtle Borders)
    With lo.DataBodyRange
        .Font.Name = "Segoe UI"
        .Font.Size = 9.5
        .Font.Color = RGB(51, 65, 85)
        .VerticalAlignment = xlCenter
        .RowHeight = 20
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(226, 232, 240)
        .Borders.Weight = xlThin
    End With
    
    For r = 1 To lo.ListRows.Count
        If r Mod 2 = 0 Then
            lo.ListRows(r).Range.Interior.Color = RGB(248, 250, 252) ' Zebra tint
        Else
            lo.ListRows(r).Range.Interior.Color = RGB(255, 255, 255) ' Clean white
        End If
    Next r
    
    ' Center key and numeric columns
    If lo.Name = "tbl_Metadata_Catalog" Then
        lo.ListColumns("Entity_ID").DataBodyRange.HorizontalAlignment = xlCenter
        lo.ListColumns("Table_Type").DataBodyRange.HorizontalAlignment = xlCenter
        lo.ListColumns("Primary_Key").DataBodyRange.HorizontalAlignment = xlCenter
        lo.ListColumns("Row_Count").DataBodyRange.HorizontalAlignment = xlCenter
        lo.ListColumns("Row_Count").DataBodyRange.NumberFormat = "#,##0"
    ElseIf lo.Name = "tbl_KPI_Dictionary" Then
        lo.ListColumns("Metric_ID").DataBodyRange.HorizontalAlignment = xlCenter
        lo.ListColumns("Target").DataBodyRange.HorizontalAlignment = xlCenter
        lo.ListColumns("Polarity").DataBodyRange.HorizontalAlignment = xlCenter
        lo.ListColumns("Alert_Threshold").DataBodyRange.HorizontalAlignment = xlCenter
    End If
End Sub
