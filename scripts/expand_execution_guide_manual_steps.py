"""
Expands docs/00_master_project_execution_guide.md with exhaustive, step-by-step
manual GUI instructions for creating, formatting, and docking all charts and slicers
for:
1. Executive Homepage Portal (00_Home_Portal) & Light/Dark Theme Switcher
2. Customer Retention Cockpit (04_CustomerRetention_Cockpit)
3. Diversity & Inclusion Cockpit (05_DiversityInclusion_Cockpit)
"""

import os

GUIDE_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "docs", "00_master_project_execution_guide.md")

NEW_CONTENT = """##### Dashboard 2: Customer Retention & Revenue Risk Cockpit (`04_CustomerRetention_Cockpit`)

```
+---------------------------------------------------------------------------------------------------+
| Middle-Top: CH_ContractRisk (270, 226)         | Right-Top: CH_PaymentFriction (762, 226)         |
| Churn Rate % by Commitment Contract            | Payment Method Risk Diagnostics                  |
+------------------------------------------------+--------------------------------------------------+
| Middle-Bottom: CH_TenureCohort (270, 510)      | Right-Bottom: CH_ServiceMatrix (762, 510)        |
| Tenure Attrition Curve & Vulnerability         | Internet Service & Add-On Protection             |
+---------------------------------------------------------------------------------------------------+
```

---

### 🛠️ Complete Step-by-Step Manual Excel GUI Guide: Customer Retention Visuals & Slicers

Follow these detailed manual instructions to build all 4 visuals and wire the 3 global slicers for Customer Retention:

```mermaid
flowchart TD
    CR1["Step 1: Staging PivotTables
(Build pt_CH_Contract, pt_CH_Tenure, pt_CH_Payment, pt_CH_Service)"] --> CR2["Step 2: Build & Format Visual 2.1
(Combo Chart: Customers & Churn %)"]
    CR2 --> CR3["Step 3: Build & Format Visual 2.2
(Tenure Attrition Curve)"]
    CR3 --> CR4["Step 4: Build & Format Visual 2.3
(Horizontal Payment Friction Bar)"]
    CR4 --> CR5["Step 5: Build & Format Visual 2.4
(Internet Service Protection Matrix)"]
    CR5 --> CR6["Step 6: Insert & Wire 3 Slicers
(Report Connections across all 4 Pivots)"]

    style CR1 fill:#1E293B,stroke:#D04A02,stroke-width:2px,color:#fff
    style CR2 fill:#1E293B,stroke:#3B82F6,stroke-width:2px,color:#fff
    style CR3 fill:#1E293B,stroke:#10B981,stroke-width:2px,color:#fff
    style CR4 fill:#1E293B,stroke:#8B5CF6,stroke-width:2px,color:#fff
    style CR5 fill:#1E293B,stroke:#EC4899,stroke-width:2px,color:#fff
    style CR6 fill:#0F172A,stroke:#D04A02,stroke-width:3px,color:#fff
```

#### Step 2.1: Visual 2.1 - Churn Rate % by Commitment Contract (`CH_ContractRisk`)
* **Docking Zone Coordinates**: Left: `284 pt`, Top: `274 pt`, Width: `448 pt`, Height: `210 pt`
* **Purpose**: Proves to the CMO/CFO that 88.55% of all subscriber churn stems from Month-to-Month contracts.

##### Manual GUI Build Instructions:
1. **Create the PivotTable**:
   - Go to worksheet **`Staging_Pivots`** $\to$ click cell **`J3`**.
   - On the ribbon, click **Insert** $\to$ **PivotTable** $\to$ choose **From Data Model** (or Use this workbook's Data Model).
   - In the Destination field, confirm `'Staging_Pivots'!$J$3` $\to$ click **OK**.
   - In the **PivotTable Analyze** tab, rename the PivotTable to: **`pt_CH_Contract`**.
2. **Assign Dimension & Measures**:
   - From table `DimContract`: Drag **`Contract`** into **Rows**.
   - From table `_Measures`: Drag **`Total Customers`** into **Values**.
   - From table `_Measures`: Drag **`Churn Rate %`** into **Values**.
3. **Format Measure Number Formats**:
   - Right-click any cell under `Total Customers` in column K $\to$ **Number Format...** $\to$ **Number** $\to$ Use 1000 Separator (`,`), 0 decimal places (`#,##0`).
   - Right-click any cell under `Churn Rate %` in column L $\to$ **Number Format...** $\to$ **Percentage** $\to$ 1 decimal place (`0.0%`).
4. **Insert the Combo Chart**:
   - Click inside `pt_CH_Contract` $\to$ click ribbon tab **Insert** $\to$ **Combo Chart** $\to$ **Create Custom Combo Chart...**
   - In the dialog:
     - Set `Total Customers` to **Clustered Column** (Primary Axis $\to$ leave checkbox unselected).
     - Set `Churn Rate %` to **Line with Markers** $\to$ **CHECK** the **Secondary Axis** checkbox.
     - Click **OK**.
5. **Declutter & Style**:
   - Cut the chart (`Ctrl + X`), switch to sheet **`04_CustomerRetention_Cockpit`**, and paste (`Ctrl + V`).
   - Rename the chart in the Name Box to: **`CH_ContractRisk`**.
   - Right-click any gray field button $\to$ click **"Hide All Field Buttons on Chart"**.
   - Delete chart title and horizontal gridlines.
   - Set Chart Area and Plot Area: **Shape Fill = No Fill**, **Shape Outline = No Outline**.
   - Format Series:
     - Clustered Columns (`Total Customers`): Fill = Solid `#94A3B8` (Soft Slate Gray), Gap Width = `65%`.
     - Secondary Line (`Churn Rate %`): Line = Solid `#DC2626` (Red Alert, 2.25 pt). Markers = Circle 6 pt `#DC2626` with white 1.5 pt outline.
   - Right-click Secondary Axis $\to$ **Format Axis...** $\to$ set Maximum = `0.50` (50%), Number Format = `0.0%`.
   - Right-click the Red Line $\to$ **Add Data Labels** $\to$ Position: **Above** (Month-to-Month: **42.7%**, One Year: **11.3%**, Two Year: **2.8%**).
6. **Docking**:
   - In ribbon tab **Chart Format**, set **Height = 2.92 in (`210 pt`)**, **Width = 6.22 in (`448 pt`)**.
   - Align exactly over `DockZone_CH_ContractRisk` (Left: `284 pt`, Top: `274 pt`).

---

#### Step 2.2: Visual 2.2 - Tenure Attrition Curve & Early Risk Window (`CH_TenureCohort`)
* **Docking Zone Coordinates**: Left: `284 pt`, Top: `558 pt`, Width: `448 pt`, Height: `210 pt`
* **Purpose**: Visualizes the steep drop in churn probability as customers cross the critical 12-month tenure threshold.

##### Manual GUI Build Instructions:
1. **Create the PivotTable**:
   - Go to **`Staging_Pivots`** $\to$ click cell **`J18`**.
   - Click **Insert** $\to$ **PivotTable** $\to$ From Data Model $\to$ click **OK**.
   - Rename PivotTable to: **`pt_CH_Tenure`**.
2. **Assign Dimension & Measures**:
   - From table `Fact_Churn`: Drag **`Tenure_Cohort`** into **Rows** (0 - 12 Months, 13 - 24 Months, 25 - 48 Months, 49 - 72 Months).
   - From table `_Measures`: Drag **`Churn Rate %`** into **Values**.
3. **Format Measure Number Format**:
   - Right-click the value column $\to$ **Number Format...** $\to$ **Percentage** (`0.0%`).
4. **Insert the PivotChart**:
   - Click inside `pt_CH_Tenure` $\to$ **Insert** $\to$ **2-D Clustered Column**.
   - Cut (`Ctrl + X`), switch to **`04_CustomerRetention_Cockpit`**, and paste (`Ctrl + V`).
   - Rename to: **`CH_TenureCohort`**.
5. **Declutter & Style**:
   - Hide all field buttons, delete chart title and legend.
   - Set Chart Area and Plot Area fill and borders to None.
   - Format Series:
     - Right-click columns $\to$ **Format Data Series...** $\to$ Gap Width = `55%`.
     - Double-click the first bar (`0 - 12 Months`) $\to$ Fill: Solid `#DC2626` (Crimson Alert: **47.4%** churn).
     - Select other bars $\to$ Fill: Solid `#64748B` (Muted Slate: ~25% down to ~7%).
   - Right-click bars $\to$ **Add Data Labels** $\to$ Position: **Outside End**, Font: Segoe UI 8.5 pt Bold `#0F172A`.
6. **Docking**:
   - Set Height = `210 pt`, Width = `448 pt`. Align over `DockZone_CH_TenureCohort` (Left: `284 pt`, Top: `558 pt`).

---

#### Step 2.3: Visual 2.3 - Payment Method Risk Diagnostics (`CH_PaymentFriction`)
* **Docking Zone Coordinates**: Left: `776 pt`, Top: `274 pt`, Width: `448 pt`, Height: `210 pt`
* **Purpose**: Proves that non-automated payment friction (Electronic Check at 45.3% churn) drives disproportionate revenue loss.

##### Manual GUI Build Instructions:
1. **Create the PivotTable**:
   - Go to **`Staging_Pivots`** $\to$ click cell **`R3`**.
   - Click **Insert** $\to$ **PivotTable** $\to$ From Data Model $\to$ click **OK**.
   - Rename PivotTable to: **`pt_CH_Payment`**.
2. **Assign Dimension & Measures**:
   - From table `Fact_Churn`: Drag **`PaymentMethod`** into **Rows**.
   - From table `_Measures`: Drag **`Churn Rate %`** into **Values**.
3. **Sort Descending**:
   - Click the small arrow on the PaymentMethod column header (or right-click Electronic check) $\to$ **Sort** $\to$ **More Sort Options...** $\to$ select **Descending (Z to A) by Churn Rate %**.
   - *Result*: Electronic check (45.3%) appears at the top of the table.
4. **Insert the Horizontal Bar Chart**:
   - Click inside `pt_CH_Payment` $\to$ **Insert** $\to$ **2-D Clustered Bar** (Horizontal).
   - Cut (`Ctrl + X`), switch to **`04_CustomerRetention_Cockpit`**, and paste (`Ctrl + V`).
   - Rename to: **`CH_PaymentFriction`**.
5. **Declutter & Style**:
   - Hide field buttons, remove chart title, legend, and vertical gridlines.
   - Right-click vertical category axis $\to$ **Format Axis...** $\to$ check **"Categories in reverse order"** (so Electronic check is at the top).
   - Format Bars:
     - Double-click the Electronic Check bar $\to$ Fill: Solid `#DC2626` (Red Alert).
     - Format the remaining bars $\to$ Fill: Solid `#059669` (Emerald Green for automated Bank Transfer and Credit Card).
   - Add Data Labels: Outside End, formatted as `0.0%`.
6. **Docking**:
   - Set Height = `210 pt`, Width = `448 pt`. Align over `DockZone_CH_PaymentFriction` (Left: `776 pt`, Top: `274 pt`).

---

#### Step 2.4: Visual 2.4 - Internet Service & Add-On Protection Matrix (`CH_ServiceMatrix`)
* **Docking Zone Coordinates**: Left: `776 pt`, Top: `558 pt`, Width: `448 pt`, Height: `210 pt`
* **Purpose**: Exposes Fiber Optic's alarming 41.9% churn rate compared to DSL's 19.0%, justifying Tech Support bundling.

##### Manual GUI Build Instructions:
1. **Create the PivotTable**:
   - Go to **`Staging_Pivots`** $\to$ click cell **`R18`**.
   - Click **Insert** $\to$ **PivotTable** $\to$ From Data Model $\to$ click **OK**.
   - Rename PivotTable to: **`pt_CH_Service`**.
2. **Assign Dimension & Measures**:
   - From table `Fact_Churn`: Drag **`InternetService`** into **Rows** (DSL, Fiber optic, No).
   - From table `Fact_Churn`: Drag **`Churn`** into **Columns** (No, Yes).
   - From table `_Measures`: Drag **`Total Customers`** into **Values**.
3. **Turn Off Grand Totals**:
   - On ribbon tab **Design** $\to$ click **Grand Totals** $\to$ select **Off for Rows and Columns**.
4. **Insert 100% Stacked Column Chart**:
   - Click inside `pt_CH_Service` $\to$ **Insert** $\to$ **100% Stacked Column**.
   - Cut (`Ctrl + X`), switch to **`04_CustomerRetention_Cockpit`**, and paste (`Ctrl + V`).
   - Rename to: **`CH_ServiceMatrix`**.
5. **Declutter & Style**:
   - Hide field buttons, remove chart title.
   - Move Legend to **Top**, Font: Segoe UI 8 pt `#64748B`.
   - Format Series:
     - Series `Yes` (Churned): Fill = Solid `#DC2626` (Crimson Alert).
     - Series `No` (Retained): Fill = Solid `#059669` (Emerald Green).
     - Gap Width = `60%`.
   - Right-click series segments $\to$ **Add Data Labels** $\to$ Position: **Center** (shows % share directly inside the bars).
6. **Docking**:
   - Set Height = `210 pt`, Width = `448 pt`. Align over `DockZone_CH_ServiceMatrix` (Left: `776 pt`, Top: `558 pt`).

---

#### Step 2.5: Inserting Slicers & Wiring Report Connections (Retention Cockpit)

To provide interactive executive filtering:
1. **Insert the 3 Slicers**:
   - Click any retention PivotTable (e.g. `pt_CH_Contract`) $\to$ ribbon tab **PivotTable Analyze** $\to$ **Insert Slicer**.
   - In the dialog:
     - From `DimContract`: Check **`Contract`**.
     - From `Fact_Churn`: Check **`PaymentMethod`**.
     - From `Fact_Churn`: Check **`InternetService`**.
     - Click **OK**.
2. **Cut and Paste to Cockpit**:
   - Select the 3 slicers $\to$ Cut (`Ctrl + X`) $\to$ switch to **`04_CustomerRetention_Cockpit`** $\to$ Paste (`Ctrl + V`).
3. **Position and Snap into Slots**:
   - **Slot 1 (Contract)**: Left: `36 pt`, Top: `280 pt`, Width: `216 pt`, Height: `145 pt`. Columns = `1`.
   - **Slot 2 (PaymentMethod)**: Left: `36 pt`, Top: `436 pt`, Width: `216 pt`, Height: `145 pt`. Columns = `1`.
   - **Slot 3 (InternetService)**: Left: `36 pt`, Top: `592 pt`, Width: `216 pt`, Height: `135 pt`. Columns = `1`.
4. **Wire Report Connections (CRITICAL STEP)**:
   - Right-click Slicer 1 (`Contract`) $\to$ click **Report Connections...**
   - Check the boxes for:
     - `pt_CH_Contract`
     - `pt_CH_Tenure`
     - `pt_CH_Payment`
     - `pt_CH_Service`
   - Click **OK**.
   - **Repeat for Slicer 2 (`PaymentMethod`) and Slicer 3 (`InternetService`)**.
   - *Verification*: Click "Month-to-month" in the Contract slicer. All 4 charts and the BAN metric cards update synchronously in real time!

---

##### Dashboard 3: Diversity, Equity & Executive Parity Cockpit (`05_DiversityInclusion_Cockpit`)

```
+---------------------------------------------------------------------------------------------------+
| Middle-Top: DI_PipelineFunnel (270, 226)       | Right-Top: DI_PromoVelocity (762, 226)           |
| Workforce Hierarchy & Broken Rung Funnel       | Promotion Velocity & Time in Grade               |
+------------------------------------------------+--------------------------------------------------+
| Middle-Bottom: DI_DeptParity (270, 510)        | Right-Bottom: DI_PerformanceAudit (762, 510)     |
| Departmental Representation & Target Gaps      | Performance Appraisal & Promotion Equity         |
+---------------------------------------------------------------------------------------------------+
```

---

### 🛠️ Complete Step-by-Step Manual Excel GUI Guide: Diversity & Inclusion Visuals & Slicers

Follow these detailed manual instructions to build all 4 visuals and wire the 3 global slicers for Diversity & Inclusion:

```mermaid
flowchart TD
    DI1["Step 1: Staging PivotTables
(Build pt_DI_Funnel, pt_DI_Parity, pt_DI_Velocity, pt_DI_Audit)"] --> DI2["Step 2: Build & Format Visual 3.1
(Workforce Broken-Rung 100% Bar)"]
    DI2 --> DI3["Step 3: Build & Format Visual 3.2
(Department Gender Parity Matrix)"]
    DI3 --> DI4["Step 4: Build & Format Visual 3.3
(Promotion Velocity Comparison)"]
    DI4 --> DI5["Step 5: Build & Format Visual 3.4
(Performance vs Promotion Equity)"]
    DI5 --> DI6["Step 6: Insert & Wire 3 Slicers
(Department, Job Level, Age Group)"]

    style DI1 fill:#1E293B,stroke:#BE185D,stroke-width:2px,color:#fff
    style DI2 fill:#1E293B,stroke:#3B82F6,stroke-width:2px,color:#fff
    style DI3 fill:#1E293B,stroke:#10B981,stroke-width:2px,color:#fff
    style DI4 fill:#1E293B,stroke:#8B5CF6,stroke-width:2px,color:#fff
    style DI5 fill:#1E293B,stroke:#D04A02,stroke-width:2px,color:#fff
    style DI6 fill:#0F172A,stroke:#BE185D,stroke-width:3px,color:#fff
```

#### Step 3.1: Visual 3.1 - Workforce Hierarchy & Broken Rung Funnel (`DI_PipelineFunnel`)
* **Docking Zone Coordinates**: Left: `284 pt`, Top: `274 pt`, Width: `448 pt`, Height: `210 pt`
* **Purpose**: Primary governance visual proving the "broken rung" cliff between Manager (34.3% F) and Executive (20.0% F).

##### Manual GUI Build Instructions:
1. **Create the PivotTable**:
   - Go to **`Staging_Pivots`** $\to$ click cell **`Z3`**.
   - Click **Insert** $\to$ **PivotTable** $\to$ From Data Model $\to$ click **OK**.
   - Rename PivotTable to: **`pt_DI_Funnel`**.
2. **Assign Dimension & Measures**:
   - From table `Dim_CareerLadder`: Drag **`Base_Job_Level`** into **Rows** (Executive, Director, Senior Manager, Manager, Senior Associate, Associate).
   - From table `Fact_Employees`: Drag **`Gender`** into **Columns** (Female, Male).
   - From table `_Measures`: Drag **`Total Employees`** into **Values**.
3. **Turn Off Grand Totals**: Design $\to$ Grand Totals $\to$ Off for Rows and Columns.
4. **Insert 100% Stacked Horizontal Bar Chart**:
   - Click inside `pt_DI_Funnel` $\to$ **Insert** $\to$ **100% Stacked Bar** (Horizontal).
   - Cut (`Ctrl + X`), switch to **`05_DiversityInclusion_Cockpit`**, and paste (`Ctrl + V`).
   - Rename to: **`DI_PipelineFunnel`**.
5. **Declutter & Style**:
   - Hide field buttons, delete chart title.
   - Right-click vertical job level axis $\to$ **Format Axis...** $\to$ check **"Categories in reverse order"** (so Level 1: Executive sits at top, Level 6: Associate at bottom).
   - Move Legend to **Top**, Font: Segoe UI 8 pt `#64748B`.
   - Format Series:
     - Series `Female`: Fill = Solid `#BE185D` (PwC Plum/Rose).
     - Series `Male`: Fill = Solid `#334155` (Navy Slate).
   - Add Data Labels: Center of each segment, formatted to show % of Row Total (Level 6: **51.8% F**, Level 4: **34.3% F**, Level 1: **20.0% F**).
6. **Docking**:
   - Set Height = `210 pt`, Width = `448 pt`. Align over `DockZone_DI_PipelineFunnel` (Left: `284 pt`, Top: `274 pt`).

---

#### Step 3.2: Visual 3.2 - Departmental Representation & Target Gaps (`DI_DeptParity`)
* **Docking Zone Coordinates**: Left: `284 pt`, Top: `558 pt`, Width: `448 pt`, Height: `210 pt`
* **Purpose**: Isolates clusters with severe gender underrepresentation (Strategy at 18.2% vs HR at 70.6%).

##### Manual GUI Build Instructions:
1. **Create the PivotTable**:
   - Go to **`Staging_Pivots`** $\to$ click cell **`Z18`**.
   - Click **Insert** $\to$ **PivotTable** $\to$ From Data Model $\to$ click **OK**.
   - Rename PivotTable to: **`pt_DI_Parity`**.
2. **Assign Dimension & Measures**:
   - From table `DimDepartment`: Drag **`Department`** into **Rows**.
   - From table `_Measures`: Drag **`Female Representation %`** into **Values**.
3. **Format Measure Number Format**:
   - Right-click value column $\to$ **Number Format...** $\to$ **Percentage** (`0.0%`).
4. **Insert the Horizontal Bar Chart**:
   - Click inside `pt_DI_Parity` $\to$ **Insert** $\to$ **2-D Clustered Bar** (Horizontal).
   - Cut (`Ctrl + X`), switch to **`05_DiversityInclusion_Cockpit`**, and paste (`Ctrl + V`).
   - Rename to: **`DI_DeptParity`**.
5. **Declutter & Style**:
   - Hide field buttons, remove chart title and legend.
   - Format Series: Bar Fill = Solid `#D04A02` (PwC Tangerine), Gap Width = `60%`.
   - Reverse category order on vertical axis.
   - Horizontal Axis Scale: Minimum = `0.0`, Maximum = `1.0` (100%), Major Unit = `0.2` (20%).
   - Add Data Labels: Outside End, formatted as `0.0%` (HR: **70.6%**, Operations: **49.3%**, Strategy: **18.2%**).
6. **Docking**:
   - Set Height = `210 pt`, Width = `448 pt`. Align over `DockZone_DI_DeptParity` (Left: `284 pt`, Top: `558 pt`).

---

#### Step 3.3: Visual 3.3 - Promotion Velocity & Glass Ceiling (`DI_PromoVelocity`)
* **Docking Zone Coordinates**: Left: `776 pt`, Top: `274 pt`, Width: `448 pt`, Height: `210 pt`
* **Purpose**: Proves that while entry-level promotions are equitable, senior-tier promotions favor male candidates by 2.1x.

##### Manual GUI Build Instructions:
1. **Create the PivotTable**:
   - Go to **`Staging_Pivots`** $\to$ click cell **`AH3`**.
   - Click **Insert** $\to$ **PivotTable** $\to$ From Data Model $\to$ click **OK**.
   - Rename PivotTable to: **`pt_DI_Velocity`**.
2. **Assign Dimension & Measures**:
   - From table `Fact_Employees`: Drag **`Job_Level_Baseline`** into **Rows**.
   - From table `_Measures`: Drag **`Female Promotion Rate %`** into **Values**.
   - From table `_Measures`: Drag **`Male Promotion Rate %`** into **Values**.
3. **Format Measure Number Formats**: Set both to **Percentage** (`0.0%`).
4. **Insert 2-D Clustered Column Chart**:
   - Click inside `pt_DI_Velocity` $\to$ **Insert** $\to$ **2-D Clustered Column**.
   - Cut (`Ctrl + X`), switch to **`05_DiversityInclusion_Cockpit`**, and paste (`Ctrl + V`).
   - Rename to: **`DI_PromoVelocity`**.
5. **Declutter & Style**:
   - Hide field buttons, remove chart title.
   - Move Legend to **Top**, Font: Segoe UI 8 pt `#64748B`.
   - Format Series:
     - Series `Female Promotion Rate`: Fill = Solid `#BE185D` (Rose).
     - Series `Male Promotion Rate`: Fill = Solid `#334155` (Navy Slate).
     - Series Overlap = `0%`, Gap Width = `80%`.
   - Add Data Labels over columns showing promotion rates.
6. **Docking**:
   - Set Height = `210 pt`, Width = `448 pt`. Align over `DockZone_DI_PromoVelocity` (Left: `776 pt`, Top: `274 pt`).

---

#### Step 3.4: Visual 3.4 - Performance Appraisal vs Promotion Equity Paradox (`DI_PerformanceAudit`)
* **Docking Zone Coordinates**: Left: `776 pt`, Top: `558 pt`, Width: `448 pt`, Height: `210 pt`
* **Purpose**: Refutes the hypothesis that promotion disparities stem from rating differences (Mean Female: 2.42 vs Male: 2.41).

##### Manual GUI Build Instructions:
1. **Create the PivotTable**:
   - Go to **`Staging_Pivots`** $\to$ click cell **`AH18`**.
   - Click **Insert** $\to$ **PivotTable** $\to$ From Data Model $\to$ click **OK**.
   - Rename PivotTable to: **`pt_DI_Audit`**.
2. **Assign Dimension & Measures**:
   - From table `Fact_Employees`: Drag **`FY20_Rating`** into **Rows** (Ratings 1 to 4).
   - From table `Fact_Employees`: Drag **`Gender`** into **Columns** (Female, Male).
   - From table `_Measures`: Drag **`Promoted FY21 Count`** into **Values**.
3. **Turn Off Grand Totals**: Design $\to$ Grand Totals $\to$ Off for Rows and Columns.
4. **Insert 2-D Clustered Column Chart**:
   - Click inside `pt_DI_Audit` $\to$ **Insert** $\to$ **2-D Clustered Column**.
   - Cut (`Ctrl + X`), switch to **`05_DiversityInclusion_Cockpit`**, and paste (`Ctrl + V`).
   - Rename to: **`DI_PerformanceAudit`**.
5. **Declutter & Style**:
   - Hide field buttons, remove chart title.
   - Move Legend to **Top**, Font: Segoe UI 8 pt `#64748B`.
   - Format Series: Female in `#BE185D`, Male in `#334155`.
   - Add Data Labels: Outside End displaying promoted headcounts.
6. **Docking**:
   - Set Height = `210 pt`, Width = `448 pt`. Align over `DockZone_DI_PerformanceAudit` (Left: `776 pt`, Top: `558 pt`).

---

#### Step 3.5: Inserting Slicers & Wiring Report Connections (D&I Cockpit)

1. **Insert the 3 Slicers**:
   - Click `pt_DI_Funnel` $\to$ **PivotTable Analyze** $\to$ **Insert Slicer**.
   - Check:
     - `DimDepartment[Department]`
     - `Fact_Employees[Job_Level_Baseline]`
     - `Fact_Employees[Age_Group]`
     - Click **OK**.
2. **Cut and Paste to Cockpit**:
   - Cut the 3 slicers $\to$ switch to **`05_DiversityInclusion_Cockpit`** $\to$ Paste.
3. **Position and Snap into Slots**:
   - **Slot 1 (Department)**: Left: `36 pt`, Top: `280 pt`, Width: `216 pt`, Height: `145 pt`.
   - **Slot 2 (Job Level)**: Left: `36 pt`, Top: `436 pt`, Width: `216 pt`, Height: `145 pt`.
   - **Slot 3 (Age Group)**: Left: `36 pt`, Top: `592 pt`, Width: `216 pt`, Height: `135 pt`.
4. **Wire Report Connections (CRITICAL STEP)**:
   - Right-click Slicer 1 (`Department`) $\to$ **Report Connections...**
   - Check: `pt_DI_Funnel`, `pt_DI_Parity`, `pt_DI_Velocity`, `pt_DI_Audit` $\to$ click **OK**.
   - Repeat for Slicer 2 (`Job_Level_Baseline`) and Slicer 3 (`Age_Group`).
   - *Verification*: Select "Operations". All 4 D&I visuals filter synchronously!
"""

def update_guide():
    with open(GUIDE_PATH, "r", encoding="utf-8") as f:
        content = f.read()

    # Locate the start of Dashboard 2
    target_heading = "##### Dashboard 2: Customer Retention & Revenue Risk Cockpit (`04_CustomerRetention_Cockpit`)"
    if target_heading not in content:
        print("Target heading not found!")
        return

    idx = content.find(target_heading)
    # Locate where Phase 7 starts
    phase7_heading = "### Phase 7: VBA Application Suite Integration"
    end_idx = content.find(phase7_heading)
    if end_idx == -1:
        print("Phase 7 heading not found!")
        return

    # Replace the section between Dashboard 2 and Phase 7 with the expanded content
    updated_content = content[:idx] + NEW_CONTENT + "\n\n---\n\n" + content[end_idx:]

    with open(GUIDE_PATH, "w", encoding="utf-8") as f:
        f.write(updated_content)

    print("Successfully expanded docs/00_master_project_execution_guide.md!")

if __name__ == "__main__":
    update_guide()
