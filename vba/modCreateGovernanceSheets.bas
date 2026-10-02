Attribute VB_Name = "modCreateGovernanceSheets"
Option Explicit

' ==============================================================================
' PwC Switzerland Virtual Case Experience - Pre-Dashboard Governance Builder
' Module: modCreateGovernanceSheets.bas
' Purpose: Automatically generates and styles:
'          1. "01_Business_Domains" (Business domain briefings & context)
'          2. "02_Metadata_&_KPI_Catalog" (Metadata inventory & KPI dictionary)
' Brand Palette: PwC Charcoal (RGB 30, 41, 59), Tangerine (RGB 208, 74, 2), Slate (RGB 100, 116, 139)
' ==============================================================================

Public Sub BuildGovernanceArchitecture()
    Dim wb As Workbook
    Dim wsDomains As Worksheet
    Dim wsCatalog As Worksheet
    
    Set wb = ActiveWorkbook
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 1. Reset / recreate target sheets cleanly
    On Error Resume Next
    Set wsDomains = wb.Worksheets("01_Business_Domains")
    If Not wsDomains Is Nothing Then wsDomains.Delete
    Set wsCatalog = wb.Worksheets("02_Metadata_&_KPI_Catalog")
    If Not wsCatalog Is Nothing Then wsCatalog.Delete
    On Error GoTo 0
    
    Set wsDomains = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    wsDomains.Name = "01_Business_Domains"
    
    Set wsCatalog = wb.Worksheets.Add(After:=wsDomains)
    wsCatalog.Name = "02_Metadata_&_KPI_Catalog"
    
    ' 2. Populate Domains Sheet
    Call PopulateDomainsSheet(wsDomains)
    
    ' 3. Populate Metadata & KPI Sheet
    Call PopulateCatalogSheet(wsCatalog)
    
    wsDomains.Activate
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "PwC Governance Architecture successfully generated!" & vbCrLf & _
           "• 01_Business_Domains: Ready" & vbCrLf & _
           "• 02_Metadata_&_KPI_Catalog: 12 Tables & 10 KPIs Loaded", _
           vbInformation, "PwC Switzerland BI Architecture"
End Sub

Private Sub PopulateDomainsSheet(ws As Worksheet)
    ws.Tab.Color = RGB(208, 74, 2) ' PwC Tangerine
    ws.DisplayGridlines = False
    
    ' Title Ribbon
    ws.Range("B2:J3").Merge
    With ws.Range("B2")
        .Value = "PwC Switzerland | Client Business Domains & Strategic Context"
        .Font.Name = "Segoe UI"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(30, 41, 59)
        .VerticalAlignment = xlCenter
    End With
    
    ws.Range("B4:J4").Interior.Color = RGB(208, 74, 2)
    ws.Rows(4).RowHeight = 4
    
    ' Subtitle
    ws.Range("B5:J5").Merge
    With ws.Range("B5")
        .Value = "Strategic context, business models, operational grains, and critical risk taxonomies for the 3 client modules."
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Italic = True
        .Font.Color = RGB(100, 116, 139)
        .VerticalAlignment = xlCenter
    End With
    
    ' Card 1: Call Center
    Call DrawDomainCard(ws, "B7", "D22", "🎧 DOMAIN 1: CALL CENTER OPERATIONS", _
        "Client: Telecommunications Customer Care Hub", _
        "• Source: 01 Call-Center-Dataset.xlsx (5,000 calls)" & vbCrLf & _
        "• Granularity: 1 Inbound Customer Call Attempt" & vbCrLf & _
        "• Core Mission: Balancing operational throughput with CSAT" & vbCrLf & _
        "• Strategic Risk: 946 Abandoned Callers (Queue drop-offs)" & vbCrLf & _
        "• Peak Critical Window: 11:00 AM - 2:00 PM (62% Abandonment)" & vbCrLf & _
        "• Operational Solution: Shift 2 morning agents to midday triage.", _
        RGB(30, 41, 59))
        
    ' Card 2: Churn & Retention
    Call DrawDomainCard(ws, "E7", "G22", "🔄 DOMAIN 2: CUSTOMER CHURN & RETENTION", _
        "Client: Subscription Broadband & Telephony Operator", _
        "• Source: 02 Churn-Dataset.xlsx (7,043 subscriber accounts)" & vbCrLf & _
        "• Granularity: 1 Customer Subscription Profile" & vbCrLf & _
        "• Core Mission: Protecting Annual Recurring Revenue ($2.86M at Risk)" & vbCrLf & _
        "• High-Risk Cohort: Month-to-Month Contracts (42.7% Churn Rate)" & vbCrLf & _
        "• Friction Triggers: Electronic Check payment failure + Fiber latency" & vbCrLf & _
        "• Operational Solution: 1-Year lock-in incentive + Autopay discount.", _
        RGB(30, 41, 59))
        
    ' Card 3: Diversity & Inclusion
    Call DrawDomainCard(ws, "H7", "J22", "⚖️ DOMAIN 3: DIVERSITY & INCLUSION", _
        "Client: Pharma Group AG (Swiss Enterprise)", _
        "• Source: 03 Diversity-Inclusion-Dataset.xlsx (500 personnel)" & vbCrLf & _
        "• Granularity: 1 Employee Career Snapshot" & vbCrLf & _
        "• Core Mission: Establishing gender parity across executive ranks" & vbCrLf & _
        "• The Broken Rung: Manager (34.3% F) -> Senior Manager (14.6% F)" & vbCrLf & _
        "• Promotion Lag: Women average +5.6 months longer before promotion" & vbCrLf & _
        "• Operational Solution: Targeted executive sponsorship & PRA audits.", _
        RGB(30, 41, 59))
        
    ws.Columns("A:K").AutoFit
    ws.Columns("A").ColumnWidth = 3
End Sub

Private Sub DrawDomainCard(ws As Worksheet, topL As String, botR As String, _
                           title As String, subtitle As String, bullets As String, headerColor As Long)
    Dim rng As Range
    Set rng = ws.Range(topL & ":" & botR)
    
    ' Outer Border & Background
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Color = RGB(226, 232, 240)
    rng.Interior.Color = RGB(248, 250, 252)
    
    ' Header Box
    Dim hdrRng As Range
    Set hdrRng = ws.Range(ws.Range(topL), ws.Cells(ws.Range(topL).Row + 1, ws.Range(botR).Column))
    hdrRng.Merge
    With hdrRng
        .Value = title
        .Font.Name = "Segoe UI"
        .Font.Size = 10.5
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = headerColor
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Content Box
    Dim bodyRng As Range
    Set bodyRng = ws.Range(ws.Cells(ws.Range(topL).Row + 2, ws.Range(topL).Column), ws.Range(botR))
    bodyRng.Merge
    With bodyRng
        .Value = subtitle & vbCrLf & vbCrLf & bullets
        .Font.Name = "Segoe UI"
        .Font.Size = 9.5
        .Font.Color = RGB(30, 41, 59)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
    End With
End Sub

Private Sub PopulateCatalogSheet(ws As Worksheet)
    ws.Tab.Color = RGB(30, 41, 59) ' PwC Charcoal
    ws.DisplayGridlines = False
    
    ' Title Ribbon
    ws.Range("B2:J3").Merge
    With ws.Range("B2")
        .Value = "PwC Switzerland | Enterprise Metadata & KPI Governance Catalog"
        .Font.Name = "Segoe UI"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(30, 41, 59)
        .VerticalAlignment = xlCenter
    End With
    ws.Range("B4:J4").Interior.Color = RGB(208, 74, 2)
    ws.Rows(4).RowHeight = 4
    
    ' Section 1: Data Model Entities
    ws.Range("B6").Value = "📦 SECTION 1: DATA MODEL ASSETS & METADATA INVENTORY"
    ws.Range("B6").Font.Bold = True
    ws.Range("B6").Font.Size = 12
    ws.Range("B6").Font.Color = RGB(30, 41, 59)
    
    Dim tblMetaHeaders As Variant
    tblMetaHeaders = Array("Entity_ID", "Business_Domain", "Table_Name", "Table_Type", "Source_File_Sheet", "Row_Count", "Granularity_Definition", "Primary_Key", "Business_Owner")
    
    Dim metaData As Variant
    metaData = Array( _
        Array("ENT-01", "Contact Center", "Fact_Calls", "Fact", "01 Call-Center-Dataset.xlsx", 5000, "1 Inbound Call Attempt", "Call_Id", "Head of Customer Care"), _
        Array("ENT-02", "Customer Retention", "Fact_Churn", "Fact", "02 Churn-Dataset.xlsx", 7043, "1 Customer Subscription Account", "customerID", "VP Customer Success"), _
        Array("ENT-03", "Workforce Governance", "Fact_Employees", "Fact", "03 Diversity-Inclusion.xlsx", 500, "1 Corporate Employee Profile", "Employee_ID", "Chief People Officer"), _
        Array("ENT-04", "Enterprise Calendar", "DimDate", "Dimension", "M Calendar Generator", 90, "1 Calendar Day (Q1 2021)", "Date", "Enterprise BI Lead"), _
        Array("ENT-05", "Contact Center", "DimAgent", "Dimension", "Derived from Fact_Calls", 8, "1 Contact Center Agent", "Agent", "Contact Center Manager"), _
        Array("ENT-06", "Contact Center", "DimTopic", "Dimension", "Derived from Fact_Calls", 5, "1 Inquiry Topic / SLA Tier", "Topic", "QA Director"), _
        Array("ENT-07", "Customer Retention", "DimContract", "Dimension", "Derived from Fact_Churn", 3, "1 Commitment Term Tier", "Contract", "Head of Pricing"), _
        Array("ENT-08", "Workforce Governance", "DimDepartment", "Dimension", "Derived from Fact_Employees", 6, "1 Organizational Unit", "Department", "CHRO"), _
        Array("ENT-09", "Performance Quotas", "Dim_PRA_Equity", "Auxiliary", "Backing 4", 6, "Department Headcount Matrix", "Department", "Compensation Committee") _
    )
    
    Dim r As Long, c As Long
    For c = 0 To UBound(tblMetaHeaders)
        ws.Cells(7, c + 2).Value = tblMetaHeaders(c)
    Next c
    
    For r = 0 To UBound(metaData)
        For c = 0 To UBound(metaData(r))
            ws.Cells(8 + r, c + 2).Value = metaData(r)(c)
        Next c
    Next r
    
    Dim rngMeta As Range
    Set rngMeta = ws.Range(ws.Cells(7, 2), ws.Cells(7 + UBound(metaData) + 1, 2 + UBound(tblMetaHeaders)))
    Dim loMeta As ListObject
    Set loMeta = ws.ListObjects.Add(xlSrcRange, rngMeta, , xlYes)
    loMeta.Name = "tbl_Metadata_Catalog"
    loMeta.TableStyle = "TableStyleMedium2"
    
    ' Section 2: Executive KPI Dictionary
    Dim kpiStartRow As Long
    kpiStartRow = 8 + UBound(metaData) + 4
    
    ws.Cells(kpiStartRow, 2).Value = "🎯 SECTION 2: EXECUTIVE KPI GOVERNANCE & METRIC DICTIONARY"
    ws.Cells(kpiStartRow, 2).Font.Bold = True
    ws.Cells(kpiStartRow, 2).Font.Size = 12
    ws.Cells(kpiStartRow, 2).Font.Color = RGB(30, 41, 59)
    
    Dim tblKPIHeaders As Variant
    tblKPIHeaders = Array("Metric_ID", "Domain", "KPI_Name", "Strategic_Objective", "Plain_Language_Definition", "Target", "Polarity", "Alert_Threshold", "Executive_Consumer")
    
    Dim kpiData As Variant
    kpiData = Array( _
        Array("KPI-CC-01", "Contact Center", "Abandonment Rate", "Queue Optimization", "% of callers who hung up before connecting", "< 15.0%", "Lower", "> 20.0%", "Head of Customer Care"), _
        Array("KPI-CC-02", "Contact Center", "Speed of Answer", "Response Time", "Average wait time in seconds for answered calls", "< 60 sec", "Lower", "> 75 sec", "Operations Director"), _
        Array("KPI-CC-03", "Contact Center", "First Contact Resolution", "Service Quality", "% of answered calls resolved on initial attempt", "> 85.0%", "Higher", "< 80.0%", "QA Director"), _
        Array("KPI-CC-04", "Contact Center", "CSAT Rating", "Customer Delight", "Average satisfaction score (1 to 5 scale)", "> 3.50", "Higher", "< 3.20", "Chief Customer Officer"), _
        Array("KPI-CH-01", "Customer Retention", "Customer Churn %", "Subscriber Protection", "Proportion of accounts terminating contract", "< 20.0%", "Lower", "> 25.0%", "VP Customer Success"), _
        Array("KPI-CH-02", "Customer Retention", "Annual Revenue at Risk", "Revenue Protection", "Annualized recurring revenue lost to churn", "< $2.0M", "Lower", "> $2.5M", "Chief Financial Officer"), _
        Array("KPI-CH-03", "Customer Retention", "Month-to-Month Churn", "Contract Health", "% of churn events on flexible month-to-month", "< 75.0%", "Lower", "> 85.0%", "Head of Pricing"), _
        Array("KPI-DI-01", "Workforce Parity", "Broken Rung Gap", "Pipeline Equity", "Drop in female representation from Mgr to Sr Mgr", "< 5.0%", "Lower", "> 15.0%", "Chief People Officer"), _
        Array("KPI-DI-02", "Workforce Parity", "Executive Parity %", "Leadership Diversity", "Female ratio at Level 1 Executive rank", "> 35.0%", "Higher", "< 25.0%", "Board of Directors"), _
        Array("KPI-DI-03", "Workforce Parity", "Promotion Velocity Lag", "Tenure Equity", "Time-in-grade gap between female and male promo", "0.0 yrs", "Lower", "> 0.3 yrs", "Compensation Committee") _
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
    loKPI.TableStyle = "TableStyleMedium2"
    
    ws.Columns("A:K").AutoFit
    ws.Columns("A").ColumnWidth = 3
    
    ' Freeze Top Panes
    ws.Activate
    ws.Range("B5").Select
    ActiveWindow.FreezePanes = True
End Sub
