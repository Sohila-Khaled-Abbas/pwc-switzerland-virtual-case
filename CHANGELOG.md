# 📋 Changelog

All notable changes to the **PwC Switzerland Virtual Case Experience** platform are documented in this file following [Keep a Changelog](https://keepachangelog.com/en/1.0.0/) and [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
