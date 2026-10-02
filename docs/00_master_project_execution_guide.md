# 📘 Master Project Execution Guide: Step-by-Step Enterprise Implementation
## PwC Switzerland Virtual Case Experience | End-to-End Build Blueprint

> **Role**: Senior Analytics Consultant & BI Architect (PwC Digital Accelerator)  
> **Ecosystem**: Microsoft Excel, Power Query (M), Power Pivot (VertiPaq Tabular Engine), DAX, VBA  
> **Master Workbook**: `PWC_Switzerland_Virtual_Case.xlsm` (Macro-Enabled Enterprise Semantic Model)

---

## 🧭 The End-to-End Project Execution Pipeline

The project follows an 8-phase enterprise lifecycle. Executing the phases in this strict chronological order prevents data model refactoring, circular dependencies, and VertiPaq corruption:

```mermaid
flowchart TD
    P1["Phase 1: Project Initiation & Forensic Data Discovery"] --> P2["Phase 2: Power Query M ETL & Forensic Transformation"]
    P2 --> P3["Phase 3: Power Pivot Tabular Modeling & Relationships"]
    P3 --> P4["Phase 4: Pre-Dashboard Governance Architecture (YOU ARE HERE)"]
    P4 --> P5["Phase 5: Explicit DAX Formulations & Measure Engineering"]
    P5 --> P6["Phase 6: Interactive Dashboard Design & Canvas Assembly"]
    P6 --> P7["Phase 7: VBA Application Suite Integration"]
    P7 --> P8["Phase 8: Hardening, Security & Executive Publishing"]

    style P1 fill:#1E293B,stroke:#3B82F6,stroke-width:2px,color:#fff
    style P2 fill:#1E293B,stroke:#F59E0B,stroke-width:2px,color:#fff
    style P3 fill:#1E293B,stroke:#10B981,stroke-width:2px,color:#fff
    style P4 fill:#0F172A,stroke:#D04A02,stroke-width:3px,color:#fff
    style P5 fill:#1E293B,stroke:#8B5CF6,stroke-width:2px,color:#fff
    style P6 fill:#1E293B,stroke:#EC4899,stroke-width:2px,color:#fff
    style P7 fill:#1E293B,stroke:#06B6D4,stroke-width:2px,color:#fff
    style P8 fill:#1E293B,stroke:#10B981,stroke-width:2px,color:#fff
```

---

## 🛠️ Phase-by-Phase Detailed Implementation Blueprint

### Phase 1: Project Initiation & Forensic Data Discovery
1. **Analyze Client Engagement Briefs**:
   - Understand the three business units:
     - **Task 1: Call Center Trends** (`01 Call-Center-Dataset.xlsx`): 5,000 telephony records.
     - **Task 2: Customer Retention** (`02 Churn-Dataset.xlsx`): 7,043 subscriber profiles.
     - **Task 3: Diversity & Inclusion** (`03 Diversity-Inclusion-Dataset.xlsx`): 500 employee records with 4 auxiliary reference sheets (`Backing 1` to `Backing 4`).
2. **Identify Forensic Anomalies Early**:
   - The **946 Missing Values** in Call Center data (`Speed of answer`, `AvgTalkDuration`, `Satisfaction rating`) $\to$ Caller abandonment, not missing values.
   - The **11 Blank TotalCharges** in Churn data $\to$ Brand-new customers with `tenure = 0`.
   - The **Inverted Grade Structure** in `Backing 4` $\to$ Spacer column shift and ascending numbering.

---

### Phase 2: Power Query M ETL & Forensic Transformation
*Reference Scripts*: [`power_query/01_staging_queries.m`](../power_query/01_staging_queries.m), [`02_dimension_transformations.m`](../power_query/02_dimension_transformations.m), [`03_fact_transformations.m`](../power_query/03_fact_transformations.m), [`04_calendar_generator.m`](../power_query/04_calendar_generator.m)

1. **Step 2.1: Establish Staging Queries (Tier 1)**:
   - Create `stg_RawCalls`, `stg_RawChurn`, `stg_RawEmployees` using parameterized file paths.
   - Enforce explicit strong typing; do **NOT** impute zeros for the 946 abandoned calls.
2. **Step 2.2: Build Conformed Dimensions (Tier 2)**:
   - Generate `DimDate` dynamically in M (January 1, 2021 to March 31, 2021).
   - Extract unique lookup dimensions: `DimAgent`, `DimTopic`, `DimContract`, `DimDepartment`.
   - Ingest auxiliary reference lookups: `Dim_EmployeeCensus` (`Backing 1`), `Dim_CareerLadder` (`Backing 2`), `Dim_NationalityCensus` (`Backing 3`), and `Dim_PRA_Equity` (`Backing 4`).
   - *Crucial Check*: Ensure `Dim_PRA_Equity` maps `Column8` to `Grade_1_Executive` down to `Column3` to `Grade_6_Junior_Officer`, purging the empty `Column2`.
3. **Step 2.3: Build Fact Tables (Tier 2)**:
   - `Fact_Calls`: Calculate `Talk_Duration_Sec` from `AvgTalkDuration` (`Time.Hour * 3600 + ...`) and integer `Call_Hour`.
   - `Fact_Churn`: Impute `TotalCharges = MonthlyCharges` when `tenure = 0`; create `Tenure_Cohort` bins.
   - `Fact_Employees`: Standardize columns; calculate numeric ranks (1-6) and `Grade_Change_Delta`.
4. **Step 2.4: Load to Data Model (Tier 3)**:
   - In Power Query: **Close & Load To...** $\to$ Select **Only Create Connection** + Check **Add this data to the Data Model**.

---

### Phase 3: Power Pivot Tabular Modeling & Relationships
*Reference Architecture*: [`docs/02_galaxy_data_model.md`](02_galaxy_data_model.md)

1. **Open Diagram View**:
   - In Excel, navigate to **Power Pivot** tab $\to$ click **Manage** $\to$ click **Diagram View**.
2. **Arrange into Ralph Kimball Constellation (Galaxy Schema)**:
   - Position conformed dimensions (`DimDate`, `DimDepartment`) in the top center.
   - Position domain-specific dimensions (`DimAgent`, `DimTopic`, `DimContract`) adjacent to their facts.
   - Position the three fact tables (`Fact_Calls`, `Fact_Churn`, `Fact_Employees`) across the center row.
   - Position auxiliary lookups along the bottom row.
3. **Wire Single-Directional 1-to-Many Relationships**:
   - `DimDate[Date]` `1` $\to$ `*` `Fact_Calls[Date]`
   - `DimAgent[Agent]` `1` $\to$ `*` `Fact_Calls[Agent]`
   - `DimTopic[Topic]` `1` $\to$ `*` `Fact_Calls[Topic]`
   - `DimContract[Contract]` `1` $\to$ `*` `Fact_Churn[Contract]`
   - `DimDepartment[Department]` `1` $\to$ `*` `Fact_Employees[Department]`
4. **Enforce VertiPaq Cardinality Rules**:
   - Confirm **zero Many-to-Many (`* : *`) relationships**.
   - Mark `DimDate` as the official Date Table: Select `DimDate` $\to$ **Design** tab $\to$ **Mark as Date Table** $\to$ select `Date` column.

---

### Phase 4: Pre-Dashboard Governance Architecture (📍 Current Step)
*Reference Code*: [`vba/modCreateGovernanceSheets.bas`](../vba/modCreateGovernanceSheets.bas)  
*Reference Guide*: [`docs/08_metadata_and_kpi_governance_guide.md`](08_metadata_and_kpi_governance_guide.md)

> [!IMPORTANT]
> **Why run this step right now?**
> Before writing dozens of DAX measures and building Pivot Tables, establishing the **Metadata Inventory** and **KPI Governance Dictionary** ensures every metric has an agreed-upon business definition, target, and calculation logic.

1. **Import the Automation Module**:
   - In Excel, press `Alt + F11` to open the VBA Editor.
   - Click **File** $\to$ **Import File...** $\to$ select `vba/modCreateGovernanceSheets.bas`.
2. **Execute the Generator**:
   - Place cursor inside `Public Sub BuildGovernanceArchitecture()` and press `F5`.
3. **Verify Generated Artifacts**:
   - **`01_Business_Domains`**: Contains the 3 branded executive cards summarizing domain models, grains, volumes, and operational levers.
   - **`02_Metadata_&_KPI_Catalog`**: Contains two official Excel Tables (`tbl_Metadata_Catalog` and `tbl_KPI_Dictionary`) formatted in PwC Dark Charcoal and Tangerine with freeze panes enabled.

---

### Phase 5: Explicit DAX Formulations & Measure Engineering
*Reference DAX Scripts*: [`dax/01_Call_Center_Measures.dax`](../dax/01_Call_Center_Measures.dax), [`02_Customer_Retention_Measures.dax`](../dax/02_Customer_Retention_Measures.dax), [`03_Diversity_Inclusion_Measures.dax`](../dax/03_Diversity_Inclusion_Measures.dax), [`04_Time_Intelligence_Measures.dax`](../dax/04_Time_Intelligence_Measures.dax), [`05_Executive_KPI_Catalog.dax`](../dax/05_Executive_KPI_Catalog.dax)  
*Data Type & Formula Registry*: [`docs/04_dax_and_kpi_glossary.md`](04_dax_and_kpi_glossary.md)

1. **Create Dedicated Measure Table**:
   - Create an empty disconnected table named `_Measures` in the Data Model to house all calculations.
2. **Explicit Data Type & Format String Standards**:
   Every measure in the `dax/` library is registered with strict enterprise metadata:
   - **Whole Number (Integer)** (`#,##0`): `[Total Demand]`, `[Answered Calls]`, `[Abandoned Calls]`, `[Total Customers]`, `[Total Employees]`.
   - **Percentage (Decimal)** (`0.00%`): `[Answer Rate %]`, `[Abandonment Rate %]`, `[Churn Rate %]`, `[Female Representation %]`.
   - **Currency (Decimal)** (`$#,##0.00`): `[Total Monthly Charges]`, `[At-Risk MRR]`, `[Avg Monthly Ticket]`.
   - **Decimal Metrics** (`#,##0.00 "s"`, `0.00`): `[Average Speed of Answer (s)]`, `[Average Handle Time (s)]`, `[Average CSAT]`.
3. **Write Core Measures by Domain**:
   - **Call Center**: `[Total Demand]`, `[Answered Calls]`, `[Abandoned Calls]`, `[Answer Rate %]`, `[Abandonment Rate %]`, `[Resolved Calls]`, `[First Contact Resolution %]`, `[Average Speed of Answer (s)]`, `[Average CSAT]`.
   - **Customer Retention**: `[Total Customers]`, `[Churned Customers]`, `[Churn Rate %]`, `[Total Monthly Charges]`, `[At-Risk MRR]`, `[Avg Monthly Ticket]`, `[Avg Tech Tickets per Customer]`.
   - **Diversity & Inclusion**: `[Total Employees]`, `[Female Count]`, `[Male Count]`, `[Female Representation %]`, `[Total Promotions FY21]`, `[Female Promotion %]`, `[Turnover Rate %]`.
4. **Implement Time Intelligence**:
   - `[Calls MTD]`, `[Calls QTD]`, `[Calls Prior Month]`, `[Calls MoM Growth %]`, `[Calls 7D Moving Avg]`, `[Cumulative Churned Revenue]`.

---

### Phase 6: Automated Dashboard Canvas Generation & UI/UX Scaffolding (via VBA)
*Reference Code*: [`vba/modDashboardUIUX.bas`](../vba/modDashboardUIUX.bas)  
*Design Masterclass*: [`docs/07_dashboard_background_and_uiux_guide.md`](07_dashboard_background_and_uiux_guide.md), [`05_dashboard_design_system.md`](05_dashboard_design_system.md)  
*Connected Modules*: [`vba/modAppState.bas`](../vba/modAppState.bas), [`vba/modDataRefresh.bas`](../vba/modDataRefresh.bas), [`vba/modFilterController.bas`](../vba/modFilterController.bas), [`vba/modExportPDF.bas`](../vba/modExportPDF.bas)

> [!IMPORTANT]
> **The Golden UI/UX Rule: Build the Canvas BEFORE Inserting Visuals.**  
> Conventional Excel users create cluttered dashboards by generating PivotCharts first and awkwardly arranging them over harsh spreadsheet gridlines.  
> **Tier-1 Management Consulting Best Practice**: Treat the Excel worksheet like a **modern Web Application (SaaS) Interface**. Use VBA to programmatically generate the entire visual scaffold—including the top navigation bar, embedded PwC color logo, interactive tab pills, live status indicators, vector SVG action buttons, 5 BAN KPI card containers, global slicer drawer, and chart docking frames—**before** docking any PivotCharts or wiring DAX measures.

```mermaid
flowchart TD
    subgraph Step1["Step 6.1: Manual VBA Automation"]
        A["Open PWC_Switzerland_Virtual_Case.xlsm"] --> B["Alt + F11 (VBA Editor)"]
        B --> C["Run modDashboardUIUX.BuildAllDashboardCanvases"]
        B --> C2["Or Run modDashboardUIUX.RunCompletePwCPlatform"]
    end

    subgraph Step2["Step 6.2: Generated Web-App Layout (Zero Encoding Bugs)"]
        C --> D1["1. Top SaaS Nav Bar (Embedded PwC Logo + Nav Tabs + Live Pill)"]
        C --> D2["2. Executive Hero Header (Domain Title + 3 SVG Action Buttons)"]
        C --> D3["3. 5 BAN KPI Metric Cards (Customizable Input + '--' Fallback)"]
        C --> D4["4. Left Global Filter Drawer (3 Designated Slicer Slots)"]
        C --> D5["5. 2x2 Grid of 4 Visual Containers (Dashed Docking Zones)"]
    end

    subgraph Step3["Step 6.3: Integrated Application Ecosystem"]
        D2 --> F1["[Refresh Data] -> modDataRefresh.RefreshPipelineSynchronously"]
        D2 --> F2["[Reset Filters] -> modFilterController.ClearAllFilters"]
        D2 --> F3["[Export PDF]    -> modExportPDF.ExportExecutiveReport"]
        D3 --> E1["Link DAX Measures / Call SetKPICardValue"]
        D4 --> E2["Insert Slicers into Filter Slots & Wire Multi-Pivot Connections"]
        D5 --> E3["Dock PivotCharts into Container Frames"]
        E3 --> E4["Run DeclutterAndFormatChart (100% Transparent Overlays)"]
    end
```

---

#### Step 6.1: Manual Execution of the Canvas Generator (`modDashboardUIUX.bas`)
You execute the automated UI/UX engine **manually** inside Excel. Follow these exact steps:

1. **Open the Master Macro-Enabled Workbook**:
   - Open `PWC_Switzerland_Virtual_Case.xlsm` in Microsoft Excel.
2. **Open the Visual Basic Editor**:
   - Press `Alt + F11` (or click **Developer** tab $\to$ **Visual Basic**).
3. **Verify the Module**:
   - In the Project Explorer window (top-left), ensure `modDashboardUIUX`, `modDataRefresh`, `modFilterController`, `modExportPDF`, and `modAppState` are present under the `Modules` folder.
4. **Execute the Master Orchestrator Macro**:
   - Double-click `modDashboardUIUX` to open the code.
   - To build only the 3 dashboard canvases: Place cursor inside `Public Sub BuildAllDashboardCanvases()` and press **`F5`**.
   - To run the complete platform end-to-end (verify governance, build canvases, refresh pipeline, clear filters): Place cursor inside `Public Sub RunCompletePwCPlatform()` and press **`F5`**.
5. **Confirmation**:
   - An executive confirmation dialog will appear summarizing the generated cockpits and connected vector controls. Click **OK**.

---

#### Step 6.2: Web-App Interface Anatomy & Layout Grid
The VBA automation constructs a responsive, mathematically balanced **1214pt modular grid** (`Left = 24pt` to `Right = 1238pt`):

| UI Component | X / Left (pt) | Y / Top (pt) | Width (pt) | Height (pt) | Web App Design Characteristics |
| :--- | :---: | :---: | :---: | :---: | :--- |
| **Top SaaS Navigation Bar** | `24` | `16` | `1214` | `52` | Embeds official PwC color logo (`assets/PwC_logo_rgb_colour_pos.png`), brand title, 5 navigation pill tabs (`01 Domains`, `02 Catalog`, `03 CC`, `04 CH`, `05 DI`) with active state fill (`#D04A02`), and live status pill with embedded SVG indicator (`LIVE VERTIPAQ`). |
| **Executive Hero Header** | `24` | `76` | `1214` | `46` | Domain title (16pt bold `#1E293B`), operational subtitle (9pt `#64748B`), and 3 SaaS action buttons with embedded vector SVG icons: `Refresh Data`, `Reset Filters`, and `Export PDF`. |
| **5 BAN KPI Scorecards** | `24` + `246` step | `130` | `230` each | `84` | Top color accent line (3pt), uppercase micro-label (8pt), clean numeric value (default `"--"` or custom input), and SLA benchmark target subtext. **Zero encoding errors (pure 7-bit ASCII)**. |
| **Left Global Filter Drawer** | `24` | `226` | `230` | `554` | Dedicated sidebar containing 3 pre-styled dashed docking slots (`Slot 1: Date`, `Slot 2: Topic/Contract/Dept`, `Slot 3: Agent/Payment/JobLevel`) with helper guidance. |
| **Middle-Top Visual Card** | `270` | `226` | `476` | `270` | Card header with title, subtitle, visual badge, and interior dashed drop zone (`448pt × 210pt`) watermarked: `[ PIVOTCHART DOCKING ZONE ]`. |
| **Middle-Bottom Visual Card**| `270` | `510` | `476` | `270` | Identical 476×270 geometry; perfectly aligned with KPI Cards 2 & 3. |
| **Right-Top Visual Card** | `762` | `226` | `476` | `270` | Identical 476×270 geometry; perfectly aligned with KPI Cards 4 & 5. |
| **Right-Bottom Visual Card** | `762` | `510` | `476` | `270` | Identical 476×270 geometry; reserved for representative audit matrix or deep dive breakdown. |

---

#### Step 6.3: Connecting DAX Measures into the Pre-Formed BAN KPI Cards

> [!WARNING]
> **Common Excel Error Resolution: *"This formula is missing a range reference or a defined name"***
> If you received this error dialog when typing in the Formula Bar, it occurs due to two Excel constraints:
> 1. **Shape Formula Bar Rule**: Excel shapes and text boxes **CANNOT** evaluate functions like `=CUBEVALUE(...)`, `=SUM(...)`, or calculations directly. A shape's formula bar accepts **only a direct cell reference** (e.g., `='03_CallCenter_Cockpit'!$AA$65`) or a **Defined Range Name**. The `CUBEVALUE` formula must reside in a worksheet cell first!
> 2. **Single Quotation Rule**: When referencing a worksheet whose name starts with a number or contains special characters (like `'03_CallCenter_Cockpit'`), the sheet name **must be enclosed in single quotes**: `='03_CallCenter_Cockpit'!$AA$65`. If quotes are omitted (e.g. `=03_CallCenter_Cockpit!AA65`), Excel parses `03` as a number, fails to recognize the sheet, and throws this error.
> 3. **Independent Value Shape**: Each KPI card is built with an independent callout shape (`Value_<cardName>`). Clicking on this shape allows you to link the metric number without overwriting the title label (`Label_<cardName>`) or target subtext (`Subtext_<cardName>`).

Follow these **step-by-step manual instructions** to link your DAX measures dynamically:

---

##### Step-by-Step Manual Workflow: Auxiliary CUBE Staging Cells (Standard Enterprise Method)

1. **Step 1: Scroll to the Staging Area Below the Canvas Fold**:
   On your active dashboard sheet (e.g., `03_CallCenter_Cockpit`), click into row 65 (below the visual grid):
   - Cell `AA65` $\to$ KPI 1 Value
   - Cell `AB65` $\to$ KPI 2 Value
   - Cell `AC65` $\to$ KPI 3 Value
   - Cell `AD65` $\to$ KPI 4 Value
   - Cell `AE65` $\to$ KPI 5 Value

2. **Step 2: Enter the CUBEVALUE Formulas into the Cells**:
   Enter the following formulas directly into the regular Excel cells (NOT inside shapes):

   * **For `03_CallCenter_Cockpit`**:
     * `AA65`: `=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Total Calls]")`
     * `AB65`: `=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Answer Rate %]")`
     * `AC65`: `=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Abandonment Rate %]")`
     * `AD65`: `=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Avg Speed of Answer]")`
     * `AE65`: `=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Avg CSAT Rating]")`

   * **For `04_CustomerRetention_Cockpit`**:
     * `AA65`: `=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Total Customers]")`
     * `AB65`: `=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Churn Rate %]")`
     * `AC65`: `=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Total Revenue at Risk]")`
     * `AD65`: `=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Contract M2M Churn Rate %]")`
     * `AE65`: `=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Tech Tickets per Customer]")`

   * **For `05_DiversityInclusion_Cockpit`**:
     * `AA65`: `=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Total Headcount]")`
     * `AB65`: `=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Female Headcount Share %]")`
     * `AC65`: `=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Broken Rung Gap]")`
     * `AD65`: `=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Female Promotion Share %]")`
     * `AE65`: `=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Time in Grade Gap]")`

3. **Step 3: Format the Staging Cells**:
   Apply formatting directly to the cells (`Ctrl + 1`):
   - Whole counts: `#,##0` (e.g., `5,000` or `7,043`)
   - Percentages: `0.00%` (e.g., `81.08%` or `26.54%`)
   - Durations: `#,##0.0 "s"` (e.g., `67.5 s`)
   - Currency: `$#,##0.00` (e.g., `$2.86M` or `$2,860,000.00`)

4. **Step 4: Select the Shape Border in the KPI Card**:
   - On the KPI card, click directly on the big number placeholder (`"--"`).
   - **Crucial**: Ensure the shape's **solid outline border** is selected. Do **NOT** double-click inside the text (there must be NO blinking text cursor inside).
   - The Name Box in the top-left will show `Value_CC_TotalDemand` (or your card's value shape name).

5. **Step 5: Link the Shape in the Formula Bar**:
   - Click into the Excel **Formula Bar** at the top (`fx`).
   - Type `= `
   - Click on cell `AA65` (or type `='03_CallCenter_Cockpit'!$AA$65`).
   - Press **Enter**.

6. **Step 6: Slicer Interactivity**:
   - To make the `CUBEVALUE` formulas react to dashboard slicers dynamically, add the slicer names as arguments:
     `=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Total Calls]", Slicer_Month, Slicer_Topic)`
   - The KPI card value will now update in real-time as slicer selections change, with the title and SLA subtext remaining perfectly formatted!

---

##### Alternative: Instant Populating via VBA Macro (`SetKPICardValue`)
If you prefer to populate initial static values or test card rendering without manual cell formulas:
- In the VBA Immediate Window (`Ctrl + G`), run:
  ```vba
  modDashboardUIUX.SetKPICardValue Worksheets("03_CallCenter_Cockpit"), "CC_TotalDemand", "5,000"
  modDashboardUIUX.SetKPICardValue Worksheets("03_CallCenter_Cockpit"), "CC_Answered", "81.08%"
  ```

---

#### Step 6.4: Docking PivotCharts into the 4 Container Frames
Insert PivotCharts from your VertiPaq Data Model and snap them into the dashed docking zones:

* **Call Center Cockpit (`03_CallCenter_Cockpit`)**:
  1. *Middle-Top*: Intraday Demand Surge $\to$ Snap into `CC_HourlyVolume` (`270, 226, 476, 270`).
  2. *Middle-Bottom*: Topic SLA Breakdown $\to$ Snap into `CC_TopicBreakdown` (`270, 510, 476, 270`).
  3. *Right-Top*: Representative Efficiency Matrix $\to$ Snap into `CC_AgentQuadrant` (`762, 226, 476, 270`).
  4. *Right-Bottom*: Agent Quality & CSAT Audit Table $\to$ Snap into `CC_AgentScorecard` (`762, 510, 476, 270`).

* **Customer Retention Cockpit (`04_CustomerRetention_Cockpit`)**:
  1. *Middle-Top*: Churn Rate by Contract $\to$ Snap into `CH_ContractRisk` (`270, 226, 476, 270`).
  2. *Middle-Bottom*: Tenure Attrition Curve $\to$ Snap into `CH_TenureCohort` (`270, 510, 476, 270`).
  3. *Right-Top*: Payment Method Friction $\to$ Snap into `CH_PaymentFriction` (`762, 226, 476, 270`).
  4. *Right-Bottom*: Internet Service & Tech Add-on Matrix $\to$ Snap into `CH_ServiceMatrix` (`762, 510, 476, 270`).

* **Diversity & Inclusion Cockpit (`05_DiversityInclusion_Cockpit`)**:
  1. *Middle-Top*: Workforce Broken Rung Funnel $\to$ Snap into `DI_PipelineFunnel` (`270, 226, 476, 270`).
  2. *Middle-Bottom*: Departmental Representation Bar $\to$ Snap into `DI_DeptParity` (`270, 510, 476, 270`).
  3. *Right-Top*: Promotion Velocity by Gender $\to$ Snap into `DI_PromoVelocity` (`762, 226, 476, 270`).
  4. *Right-Bottom*: Appraisal vs Promotion Audit Matrix $\to$ Snap into `DI_PerformanceAudit` (`762, 510, 476, 270`).

---

#### Step 6.5: Applying Transparent Chart Decluttering
To eliminate spreadsheet clashing and give charts the native look of a custom SaaS web app, run the decluttering routine on each inserted PivotChart:

```vba
' Execute in the VBA Immediate Window (Ctrl + G) or via a helper macro:
Call modDashboardUIUX.DeclutterAndFormatChart(ActiveSheet.ChartObjects("YourChartName"))
```

* **Visual Transformation**:
  - Sets `.ChartArea.Format.Fill.Visible = msoFalse` (100% transparent).
  - Sets `.PlotArea.Format.Fill.Visible = msoFalse` (100% transparent).
  - Strips all exterior chart borders (`.Line.Visible = msoFalse`).
  - Softens gridlines to 0.75pt `#E2E8F0` and unifies typography to 8.5pt Segoe UI (`#64748B`).
  - The chart seamlessly floats inside the card container!

---

#### Step 6.6: Wiring Global Interactive Slicers into the Filter Drawer
1. Insert Slicers from the VertiPaq Data Model fields into the 3 designated slots in the Left Slicer Drawer:
   - **Call Center**: `DimDate[Month_Name]` (Slot 1), `DimTopic[Topic]` (Slot 2), `DimAgent[Agent]` (Slot 3).
   - **Customer Churn**: `DimContract[Contract]` (Slot 1), `Fact_Churn[PaymentMethod]` (Slot 2), `Fact_Churn[InternetService]` (Slot 3).
   - **Diversity & Inclusion**: `DimDepartment[Department]` (Slot 1), `Fact_Employees[Job_Level]` (Slot 2), `Fact_Employees[Age_Group]` (Slot 3).
2. **Connect Slicers across all Dashboard Visuals**:
   - Right-click each Slicer $\to$ **Report Connections...** $\to$ check all PivotTables on that cockpit.
3. **Apply PwC Brand Styling**:
   - Right-click Slicer $\to$ Slicer Styles $\to$ select custom style featuring PwC Tangerine (`#D04A02`) for selected items and clean off-white (`#F1F5F9`) for unselected items.

---

### Phase 7: VBA Application Suite Integration
*Reference Modules*: [`vba/modAppState.bas`](../vba/modAppState.bas), [`vba/modDashboardUIUX.bas`](../vba/modDashboardUIUX.bas), [`vba/modFilterController.bas`](../vba/modFilterController.bas), [`vba/modNavigation.bas`](../vba/modNavigation.bas), [`vba/modExportPDF.bas`](../vba/modExportPDF.bas)

1. **Import Core Modules**:
   - `modAppState.bas`: Prevents screen flicker and pauses calculation during batch refreshes.
   - `modFilterController.bas`: Provides a single-click "Clear All Filters" button per dashboard.
   - `modNavigation.bas`: Powers tab switching buttons between cockpits.
   - `modExportPDF.bas`: Automates pixel-perfect executive PDF report generation.
2. **Attach Macros to UI Buttons**:
   - Wire buttons on each sheet to navigation and reset handlers.

---

### Phase 8: Hardening, Security & Executive Publishing
*Reference Guide*: [`docs/09_excel_dashboard_publishing_and_distribution_guide.md`](09_excel_dashboard_publishing_and_distribution_guide.md)

1. **Sheet Protection Configuration**:
   - Protect sheets with `AllowUsingPivotTables = True` and `AllowFiltering = True` so users can interact with Slicers without modifying grid structures.
2. **Configure Browser View Options**:
   - Go to **File** $\to$ **Info** $\to$ **Browser View Options** $\to$ Display ONLY dashboard sheets (`01_Business_Domains`, `02_Metadata_&_KPI_Catalog`, `03_CallCenter`, `04_Retention`, `05_D&I`).
3. **Deploy to Target Channel**:
   - Publish to **SharePoint / OneDrive** for web-based interactive consumption.
   - Publish to **Power BI Service** or export executive **PDF briefings**.
