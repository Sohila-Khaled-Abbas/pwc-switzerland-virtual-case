# 📋 Changelog

All notable changes to the **PwC Switzerland Virtual Case Experience** platform are documented in this file following [Keep a Changelog](https://keepachangelog.com/en/1.0.0/) and [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [2.7.0] - 2026-10-03

### 🚀 Enterprise Production Release & Slicer Architecture

- **Persistent 6-Tab Global Top Navigation Suite (`vba/modNavigation.bas`)**:
  - Implemented an identical, enterprise-grade 6-tab navigation bar across all 6 worksheets (`00_Home_Portal`, `01_Business_Domains`, `02_Metadata_&_KPI_Catalog`, `03_CallCenter_Cockpit`, `04_CustomerRetention_Cockpit`, `05_DiversityInclusion_Cockpit`).
  - Active sheet dynamically highlighted in vibrant PwC Tangerine (`#D04A02`) with bold white typography; inactive tabs rendered in Slate Navy (`#1E293B`).
  - Dual-action binding: Native worksheet hyperlinks (`#'SheetName'!A1`) for instantaneous switching, backed by VBA callback routines (`modNavigation.NavigateTo...`).
  - Global header controls preserved: `Toggle Dark Mode`, `[LIVE] VERTIPAQ` health status pill, and `Export PDF` executive briefing action button.

- **9 Native VertiPaq Data Model Slicers Across All 3 Cockpits**:
  - Replaced legacy visual placeholder shapes with 9 authentic native Excel Slicers instantiated directly from the VertiPaq tabular model (`ThisWorkbookDataModel`):
    - **Call Center Cockpit (`03_CallCenter_Cockpit`)**: `Month` (`[DimDate].[Month_Name]`), `Topic` (`[DimTopic].[Topic]`), `Agent` (`[DimAgent].[Agent]`), connected across all 4 Call Center analytical PivotTables (`pt_Agent`, `PivotChartTable3`, `PivotChartTable6`, `PivotChartTable8`).
    - **Customer Retention Cockpit (`04_CustomerRetention_Cockpit`)**: `Contract` (`[DimContract].[Contract]`), `Payment Method` (`[Fact_Churn].[PaymentMethod]`), `Internet Service` (`[Fact_Churn].[InternetService]`), connected across all 4 Retention PivotTables (`pt_CH_Contract`, `pt_CH_Tenure`, `pt_CH_Payment`, `pt_CH_Service`).
    - **Diversity & Inclusion Cockpit (`05_DiversityInclusion_Cockpit`)**: `Department` (`[DimDepartment].[Department]`), `Job Level` (`[Dim_CareerLadder].[Base_Job_Level]`), `Age Group` (`[Fact_Employees].[Age_Group]`), connected across all 4 D&I PivotTables (`pt_DI_Funnel`, `pt_DI_Parity`, `pt_DI_Promo`, `pt_DI_Rating`).
  - Slicers styled in dark theme (`SlicerStyleDark2`), snapped into Left Drawer slots (`Left: 36pt`, `Width: 206pt`), with multi-pivot report connections verified.

- **8 Non-Overlapping Analytical Staging PivotTables & Native PivotCharts**:
  - Restored 8 clean, non-overlapping VertiPaq PivotTables on `Staging_Pivots`.
  - Re-established native PivotChart links on `04_CustomerRetention_Cockpit` and `05_DiversityInclusion_Cockpit`, restoring full dynamic responsiveness when slicer selections are toggled.
  - Docked into the standard 2x2 container grid (`Left: 284.0 / 776.0`, `Top: 274.0 / 558.0`, `Width: 448.0`, `Height: 210.0`) with 100% transparent decluttered styling.

- **Customer Retention Revenue Semantic Reconciliation**:
  - Reconciled the 3 distinct financial figures across the model, DAX catalog, tickers, and documentation:
    - **$139,130.85 Monthly Revenue at Risk (MRR)**: Monthly recurring revenue lost from 1,869 churned subscribers (`CALCULATE(SUM(Fact_Churn[MonthlyCharges]), Fact_Churn[Churn] = "Yes")`).
    - **$1,669,570.20 Annualized Run-Rate (ARR)**: Annualized recurring revenue leakage ($139.1K $\times$ 12 months).
    - **$2,862,926.75 Cumulative Historical Lifetime Charges**: Cumulative lifetime billing recorded from churned subscribers prior to termination (`CALCULATE(SUM(Fact_Churn[TotalCharges]), Fact_Churn[Churn] = "Yes")`).
  - Eliminated legacy ambiguity between monthly recurring run-rate loss and cumulative lifetime billing.

- **Single-Source-of-Truth Dynamic CUBEVALUE KPI Architecture**:
  - Bound all 15 BAN metric callouts to staged `CUBEVALUE("ThisWorkbookDataModel", ...)` formula cells in row 65 (`AA65:AE65`) via shape `.DrawingObject.Formula`.
  - Hidden technical helper rows (`60:75`) on all 3 cockpits to maintain an executive presentation canvas.
  - Slicer filtering propagates seamlessly without wiping card titles or SLA subtext.

- **Zero-Defect Quality Assurance & Process Lock Resolution**:
  - Diagnosed and resolved the desktop "File in Use" lock: Terminated orphaned headless background `EXCEL.EXE` processes, wiped temporary lock files, and unblocked file security attributes.
  - Standardized viewport geometry: Uniform 80% zoom level, cursor parked at `A1`, and initial view set to `00_Home_Portal`.
  - Automated audit verified 0 formula errors (`#REF!`, `#VALUE!`, `#DIV/0!`), 13 VertiPaq tables, and 81 active DAX measures.
  - Synchronized `PWC_Switzerland_Virtual_Case.xlsm` with the QA-passed build (`PWC_Switzerland_Virtual_Case_FIXED.xlsm`).

---

## [2.6.0] - 2026-10-02

### 🚀 Added & Automated
- **26pt Bold Executive BAN Typography (`vba/modDashboardUIUX.bas`)**:
  - Upgraded `Value_<cardName>` metric callouts to **26pt Bold** high-contrast typography in executive slate (`#0F172A`, `PWC_DARK_SLATE`), increasing textbox height to 40pt with zero margins and middle vertical alignment.
  - **Countered Excel Font-Reset Bug**: Programmatically re-applies 26pt bold styling across both `TextFrame2.TextRange.Font` and legacy `TextFrame.Characters.Font` immediately after `DrawingObject.Formula` assignment, preventing Excel from defaulting formula-linked shapes to 9pt regular.
  - Expanded staging columns `AA:AE` width to `18` via VBA, permanently eliminating numeric text truncation (`###`) in off-screen CUBE calculation cells.
- **Automated CUBEVALUE Engine & 3-Layer Shape Architecture (`vba/modDashboardUIUX.bas`)**:
  - Re-architected `BuildKPICard` to split every KPI card into 3 distinct, independent Excel shapes:
    - `Label_<cardName>`: Fixed micro-title label (8pt Bold, `PWC_TEXT_MUTED`).
    - `Value_<cardName>`: Dedicated dynamic metric callout (26pt Bold, `PWC_DARK_SLATE`, formula-linkable).
    - `Subtext_<cardName>`: Fixed SLA / benchmark status badge (8pt, accent color).
  - **Eliminated Excel Text-Wipe Bug**: Decoupling into 3 independent shapes ensures dynamic cell binding never wipes out the card title or SLA subtext.
  - **Automated Staging & Linking Routine (`AutomateAndLinkKPICards`)**:
    - Writes descriptive staging headers in row 64 (`AA64:AE64`).
    - Injects live `=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[...]")` formulas in row 65 (`AA65:AE65`) for all 15 KPI cards across the 3 executive cockpits.
    - Applies enterprise number formatting (`#,##0`, `0.00%`, `#,##0.0 "s"`, `$#,##0.00`) directly to staged cells.
    - Automatically binds each `Value_<cardName>` shape to `='SheetName'!$Col$65` via `shpValue.DrawingObject.Formula`.
  - **Resolved Excel Formula Reference Error**: Addressed *"This formula is missing a range reference or a defined name"* caused by lack of single quotes on numbered sheet names (`03_...`, `04_...`, `05_...`) and illegal function writing in shape formula bars by fully automating proper quoting (`='SheetName'!$Col$65`) in VBA.
  - Integrated into `BuildCallCenterCanvas`, `BuildCustomerRetentionCanvas`, and `BuildDiversityInclusionCanvas` for 1-click end-to-end canvas generation.
- **Power Pivot In-Memory VertiPaq Measure Deployment (`PWC_Switzerland_Virtual_Case.xlsm`)**:
  - Deployed 55+ missing explicit DAX measures and aliases directly into the VertiPaq tabular model (`_Measures`), bringing total active model measures to 67.
  - Completely resolved naming conflicts: added bidirectional alias bindings between `[Total Demand]` / `[Total Calls]` / `[Total Inbound Calls]`, `[At-Risk MRR]` / `[Total Revenue at Risk]`, `[Female Count]` / `[Female Employees]`, `[Total Employees]` / `[Total Headcount]`, etc.
  - Updated `dax/02_Customer_Retention_Measures.dax` and `dax/03_Diversity_Inclusion_Measures.dax` to match all measures and aliases in the data model.
- **Documentation Masterclass Overhaul**:
  - Updated `docs/00_master_project_execution_guide.md` Step 6.3 with full error diagnosis, automated engine architecture, and complete 15-KPI reference mapping table.
  - Updated `docs/07_dashboard_background_and_uiux_guide.md` with visual hierarchy diagram of the 3-layer shape architecture, CUBEVALUE staging grid rules, and 26pt bold typography guidelines.

---

## [2.5.0] - 2026-10-02

### 🔧 Fixed & Enhanced
- **Text Encoding Bug Elimination (`vba/modDashboardUIUX.bas`)**:
  - Eliminated `â€”` and `â— ` character corruption by migrating all strings and comments to 100% 7-bit pure ASCII (`"--"` placeholder, `LIVE VERTIPAQ` pill).
- **Vector SVG Action Icons (`assets/icons/`)**:
  - Authored 4 production-grade vector SVG icons themed in PwC corporate colors:
    - `icon_reset_filter.svg`: Charcoal `#1E293B` and Tangerine `#D04A02` filter reset funnel.
    - `icon_export_pdf.svg`: Crisp white `#FFFFFF` executive document download icon.
    - `icon_refresh_pipeline.svg`: Tangerine `#D04A02` circular synchronization arrows.
    - `icon_live_indicator.svg`: Emerald `#059669` pulsing live status badge.
  - Automatically embeds SVG vector icons inside the header buttons and status pill.
- **Cross-Script Ecosystem Connectivity**:
  - Wired `Reset Filters` action button directly to `modFilterController.ClearAllFilters`.
  - Wired `Export PDF` action button directly to `modExportPDF.ExportExecutiveReport` (with dynamic active cockpit naming).
  - Added `Refresh Data` action button wired directly to `modDataRefresh.RefreshPipelineSynchronously`.
  - Added master platform orchestrator: `RunCompletePwCPlatform()` to execute governance verification, canvas creation, data model refresh, and filter resets end-to-end.
- **KPI Card Input Customization**:
  - Added optional `kpiValue` argument to `BuildKPICard` and implemented `SetKPICardValue(ws, cardName, newValue, subtext)` for dynamic metric updates.
- **Explicit DAX Data Type & Format String Registry**:
  - Annotated all explicit measures across `dax/01_Call_Center_Measures.dax`, `dax/02_Customer_Retention_Measures.dax`, `dax/03_Diversity_Inclusion_Measures.dax`, `dax/04_Time_Intelligence_Measures.dax`, and `dax/05_Executive_KPI_Catalog.dax` with strict data types (`Whole Number`, `Percentage`, `Decimal Number`, `Currency`) and format strings (`#,##0`, `0.00%`, `$#,##0.00`, `#,##0.00 "s"`).
  - Updated `docs/04_dax_and_kpi_glossary.md` and `docs/00_master_project_execution_guide.md` accordingly.

---

## [2.4.0] - 2026-10-02

### 🚀 Added & Redesigned
- **Web-Application UI/UX Dashboard Architecture (`vba/modDashboardUIUX.bas`)**:
  - Redesigned executive presentation canvases into a modern Web-Application (SaaS) layout.
  - **Embedded Official PwC Logo**: Automatically detects and permanently embeds `assets/PwC_logo_rgb_colour_pos.png` (`SaveWithDocument = msoTrue`) on the top-left of the master navigation ribbon.
  - **Top SaaS Navigation Bar**: Built unified header ribbon (`1214pt × 52pt`) featuring corporate branding, interactive navigation pill tabs (`01 Domains`, `02 Catalog`, `03 CC`, `04 CH`, `05 DI`) with active tab highlighting, and live VertiPaq status pill (`[●] LIVE MODEL`).
  - **Executive Hero Header**: Added domain title, operational subtitle, and web-app action buttons (`Reset Filters`, `Export PDF`).
  - **BAN KPI Metric Cards (Zero Hardcoded Data)**: Standardized 5 floating KPI cards (`230pt × 84pt`) with top accent lines, uppercase micro-labels, clean `"—"` placeholders ready for DAX/CUBE linkage, and SLA benchmark targets (no hardcoded static counts).
  - **Left Global Filter Drawer**: Added dedicated vertical sidebar (`230pt × 554pt`) with 3 pre-styled dashed docking slots for date, categorical, and segment slicers.
  - **2x2 Visual Container Grid**: Built 4 floating cards (`476pt × 270pt`) with chart type badges and dashed drop zones watermarked `[ PIVOTCHART DOCKING ZONE ]`.
  - **Pixel-Perfect 1214pt Modular Grid**: Perfect alignment across KPI cards, slicer drawer, and chart containers.
- **Manual Execution Workflow in Master Execution Guide (`docs/00_master_project_execution_guide.md`)**:
  - Updated Phase 6 with complete manual execution guide (`Alt + F11` $\to$ `F5` or `Developer` $\to$ `Macros` $\to$ `Run`).
  - Detailed web-app layout geometry, DAX measure linking techniques (formula-linked text boxes and auxiliary CUBE cells), and chart docking instructions.
- **UI/UX Design Masterclass Updated (`docs/07_dashboard_background_and_uiux_guide.md`)**:
  - Documented web-app SaaS navigation bar, embedded logo logic, placeholder-driven KPI cards, and visual docking drop zones.

---

## [2.3.0] - 2026-10-02

### 🚀 Added
- **Automated Dashboard UI/UX Canvas Engine (`vba/modDashboardUIUX.bas`)**:
  - Implemented `BuildAllDashboardCanvases()`, `BuildCallCenterCanvas()`, `BuildCustomerRetentionCanvas()`, and `BuildDiversityInclusionCanvas()`.
  - Automatically provisions presentation-ready canvases (`03_CallCenter_Cockpit`, `04_CustomerRetention_Cockpit`, `05_DiversityInclusion_Cockpit`) before inserting visuals.
  - Floods canvas with `#F8FAFC` background, suppresses gridlines and headings, standardizes 100% zoom.
  - Constructs PwC executive header banners with `<-- Executive Hub` return breadcrumbs, top KPI scorecard ribbons (5 BAN cards per cockpit), vertical slicer panel container (`224px × 530px`), and 4 floating rounded visual container cards with diffused drop shadows (`msoShadow21`).
  - Added chart transparency decluttering procedure (`DeclutterAndFormatChart`) to strip borders and make chart backgrounds 100% transparent.
- **VBA Scaffolding Workflow in Master Guide (`docs/00_master_project_execution_guide.md`)**:
  - Restructured Phase 6 into a canvas-first UI/UX methodology detailing automated 1-click VBA canvas generation, visual docking, transparent decluttering, and slicer wiring.

### 🔧 Fixed & Optimized
- **Master Workbook Consolidation (`PWC_Switzerland_Virtual_Case.xlsm`)**:
  - Standardized on macro-enabled `.xlsm` to host both the in-memory VertiPaq tabular data model (13 tables, 9 relationships) and the complete VBA suite (`modCreateGovernanceSheets`, `modDashboardUIUX`).
  - Safely purged conflicting `.xlsx` copy and updated all documentation, scripts (`sync_pwc_docs.py`), and file watchers (`watch_autopublish.ps1`).
- **Pre-Dashboard Governance Auto-Fit & Styling Overhaul (`vba/modCreateGovernanceSheets.bas`)**:
  - Eliminated card text and title truncation: Expanded card columns B, D, F to width 54, enforced `.WrapText = True`, and increased card header height to 32pt and bullet rows to 22pt.
  - Eliminated ANSI `??` emoji corruption; adopted clean executive ASCII notation (`[01]`, `[02]`, `[03]`, `[SECTION 1]`, `[SECTION 2]`, `-->`, `<--`).
  - Prevented AutoFit blowout on column B: Merged section headers across table widths (`B7:J7` and `B24:J24`).
  - Implemented intelligent AutoFit with +4.5 padding for filter dropdown arrows and dynamic floor minimums.
  - Applied bespoke PwC corporate table styling (Charcoal `#1E293B` header, Tangerine `#D04A02` bottom border, Fact/Dimension/Auxiliary badges, soft emerald/amber polarity pills).
- **DAX Schema Reconciliations**:
  - Corrected `01_Call_Center_Measures.dax`: Replaced legacy column names with exact VertiPaq schema (`Fact_Calls[Answered]`, `Fact_Calls[Resolved]`, `Fact_Calls[Speed_Of_Answer_Sec]`, `Fact_Calls[Talk_Duration_Sec]`).
  - Corrected `03_Diversity_Inclusion_Measures.dax`: Aligned column names to `Fact_Employees[Employee_ID]`, `Fact_Employees[Promoted_FY21]`, and `Fact_Employees[FY20_Leaver]`.

---

## [2.2.1] - 2026-10-02

### 🔧 Fixed
- **VBA Compile Error Resolution (`modCreateGovernanceSheets.bas`)**:
  - Corrected `ws.DisplayGridlines = False` (which triggered `Compile error: Method or data member not found`) to `ActiveWindow.DisplayGridlines = False`.
  - Added pre-emptive cleanup for `tbl_Metadata_Catalog` and `tbl_KPI_Dictionary` to prevent global Excel table name collisions.
  - Added manual calculation suppression and robust error handling to guarantee clean UI updates.

### 🚀 Added
- **Master Project Execution Guide (`docs/00_master_project_execution_guide.md`)**: Comprehensive 8-phase step-by-step roadmap from data ingestion to VertiPaq modeling, governance sheet generation, DAX engineering, dashboard assembly, and publishing.
- **Publishing & Distribution Masterclass (`docs/09_excel_dashboard_publishing_and_distribution_guide.md`)**: Reddit-inspired distribution guide covering Browser View Options, SharePoint/OneDrive, Power BI Service, Web Embed, and Sheet Protection.

---

## [2.2.0] - 2026-10-02

### 🚀 Added
- **Pre-Dashboard Orientation Architecture**: Complete design blueprint and guide for establishing the Three-Tier Entryway in Excel workbooks (`docs/08_metadata_and_kpi_governance_guide.md`).
- **Domain-by-Domain Context Briefing**: Deep forensic business narrative covering the 3 client datasets (Call Centre Telephony, Customer Retention Economics, and Pharma Group AG Executive Parity).
- **Enterprise Metadata Catalog Table**: Structured 12-table data asset inventory detailing entity IDs, tables, grains, primary keys, source mapping, and data stewards.
- **Executive KPI Governance Dictionary**: Structured metric registry featuring 10+ core strategic KPIs, business definitions, DAX formulations, targets, polarities, alert thresholds, and executive consumers.
- **Automated Sheet Builder Macro**: VBA module `vba/modCreateGovernanceSheets.bas` to automatically construct, format in PwC corporate styling, and link `01_Business_Domains` and `02_Metadata_&_KPI_Catalog`.

---

## [2.1.1] - 2026-10-02

### 🔧 Fixed
- **ETL Data Logic & Column Mapping Correction (`Dim_PRA_Equity`)**: Fixed Power Query M transformation for auxiliary dimension `Dim_PRA_Equity` (`Backing 4`).
  - Purged blank spacer Excel column D (`Column2`).
  - Corrected ascending grade index inversion (`Column8` down to `Column3` mapping to Grades 1 to 6).
  - Resolved erroneous 0 values in `Grade_1_Executive` and eliminated false 191 director counts in Operations/Sales.
  - Reconciled departmental allocations with enterprise headcount (500 personnel) and verified 100% alignment with `broken_rung_funnel.svg` and `Fact_Employees`.

---

## [2.1.0] - 2026-10-02

### 🚀 Added
- **Dynamic Calendar Dimension**: Full M-code sequential generator with fiscal quarters and weekend flags (`power_query/04_calendar_generator.m`).
- **Time Intelligence Measures**: Added MTD, QTD, Prior Month, MoM Growth %, and 7-day rolling moving average DAX calculations.
- **Enterprise Documentation Hub**: Complete suite of 6 executive architecture guides, data dictionary, and UI/UX design blueprints.
- **Executive Dashboard Specifications**: Completed full UI/UX design blueprints and KPI specifications for all three engagement cockpits: Call Centre Operations (`dashboards/01_call_center_dashboard.md`), Customer Retention & Revenue Risk (`dashboards/02_customer_retention_dashboard.md`), and Diversity & Inclusion Leadership Scorecard (`dashboards/03_diversity_inclusion_dashboard.md`).
- **Continuous Integration Workflow**: GitHub Actions workflow for validating repository workbook assets and code syntax.

### 🔧 Fixed
- **Duplicate Connection Triage**: Eliminated ghost connections `Query - DimDate1` and `Query - DimDepartment1` via Excel COM automation.
- **Model Cleanliness**: Verified exactly 12 single-instance tables in VertiPaq memory with 100% 1-to-many relationship integrity.

---

## [2.0.0] - 2026-10-01

### 🚀 Added
- **Ralph Kimball Galaxy Schema**: Integrated all 3 client datasets (`01 Call-Center`, `02 Churn`, `03 Diversity-Inclusion`) into a single unified semantic model.
- **Auxiliary Census Dimensions**: Ingested `Backing 1` to `Backing 4` reference sheets into dedicated lookup tables.
- **VBA Macro Architecture**: Enterprise application suite including state caching, dynamic data refresh, PDF export, and filter resets.

---

## [1.0.0] - 2026-09-30

### 🚀 Initial Release
- Baseline Call Centre Operations model (5,000 telephony logs).
- Initial CSAT, Abandonment, and Speed of Answer DAX formulations.
