import re

with open('docs/00_master_project_execution_guide.md', 'r', encoding='utf-8') as f:
    content = f.read()

replacement_text = """###### Visual 1.4: Representative Quality & CSAT Audit (Right-Bottom)
* **Container Name**: `CC_AgentScorecard` | **Badge**: `8 Agents Active` (or `Matrix Table`)
* **Docking Zone Coordinates**: Left: `776 pt`, Top: `546 pt`, Width: `448 pt`, Height: `212 pt`

---

#### 🔍 Forensic Analysis: The 4 Visual Defects in the Legacy Scorecard & Why They Occur

If your Agent Scorecard currently looks like the legacy screenshot (misaligned margins, dashed border clashing, gray filter arrow, and `6533.1%` speed values), here is the exact forensic diagnosis:

```
+---------------------------------------------------------------------------------------------------------+
| [LEGACY DEFECT]                                     --> [MODERN SAAS WEB-APP FIX]                       |
+---------------------------------------------------------------------------------------------------------+
| 1. Avg Speed displays 6533.1%                       --> Set Field Number Format to Custom: 0.0 "s"       |
| 2. Agent header shows gray AutoFilter [v] arrow    --> Uncheck "Show field captions & filter drop downs"|
| 3. Table scrapes bottom border & has uneven sides   --> Calibrate Staging_Pivots column widths & heights |
| 4. Awkward dashed line surrounds the table          --> Convert DockZone to solid 0.75pt #E2E8F0 panel   |
+---------------------------------------------------------------------------------------------------------+
```

1. **The 6533.1% Percentage Speed Bug**: In Excel PivotTables, if a calculated measure or field is assigned the `Percentage` format instead of `Custom` or `Number`, Excel multiplies the scalar number by 100 ($65.33 \times 100 = 6533.1\%$). Average Speed is measured in **elapsed seconds**, not percentage completion!
2. **The AutoFilter Dropdown Arrow**: By default, Excel PivotTables enable row field filter dropdown arrows on `Agent`. In a modern executive dashboard, an interactive dropdown inside a docked matrix looks like an unfinished raw spreadsheet rather than a polished web app widget.
3. **The Box Fit & Aspect Ratio Mismatch**: `Container_CC_AgentScorecard` provides an interior content area of $448\\text{ pt} \\times 212\\text{ pt}$. When `C3:H11` on `Staging_Pivots` is left at default column widths, its natural aspect ratio is too narrow ($397\\text{ pt} \\times 170\\text{ pt}$), forcing Excel's Linked Picture scaler to stretch vertically, scraping the bottom border and leaving awkward blank gaps on the sides.
4. **Dashed Blueprint Border vs. Solid Modern Card**: The original canvas generated `DockZone_CC_AgentScorecard` with a dashed outline (`msoLineDash`) as a developer drop zone. Leaving a dashed line behind a finished table creates an unpolished "box-in-a-box" wireframe look. Modern SaaS UI (Stripe, Linear, Datadog) uses crisp solid borders (`#E2E8F0`) with subtle drop shadows.

---

### 🛠️ Complete Step-by-Step Manual Excel GUI Guide: Pixel-Perfect Fit & Web App Styling

Follow these exact steps manually in your Excel GUI to transform the table into a pixel-perfect modern web app component:

```mermaid
flowchart TD
    S1["Step 1: Fix Speed Number Format\n(Custom: 0.0 's')"] --> S2["Step 2: Hide AutoFilter Dropdown\n(PivotTable Options -> Display)"]
    S2 --> S3["Step 3: Calibrate Staging Cell Geometry\n(Exact Widths: ~442pt, Heights: 22.5pt)"]
    S3 --> S4["Step 4: Modernize Container Box\n(Solid 0.75pt #E2E8F0 + Soft Shadow)"]
    S4 --> S5["Step 5: Embed Vector SVG Icon\n(Insert icon_audit_matrix.svg)"]
    S5 --> S6["Step 6: Dock Live Linked Picture\n(444pt x 204pt Centered Fit)"]
    S6 --> S7["Step 7: Format Other 3 Visuals\n(Gradients, Clean Axes, Top Legends)"]

    style S1 fill:#1E293B,stroke:#D04A02,stroke-width:2px,color:#fff
    style S2 fill:#1E293B,stroke:#3B82F6,stroke-width:2px,color:#fff
    style S3 fill:#1E293B,stroke:#10B981,stroke-width:2px,color:#fff
    style S4 fill:#1E293B,stroke:#8B5CF6,stroke-width:2px,color:#fff
    style S5 fill:#1E293B,stroke:#EC4899,stroke-width:2px,color:#fff
    style S6 fill:#0F172A,stroke:#D04A02,stroke-width:3px,color:#fff
    style S7 fill:#1E293B,stroke:#06B6D4,stroke-width:2px,color:#fff
```

---

#### Step 1: Fix the `Avg Speed (s)` Number Format (Permanently in Field Settings)
1. Switch to worksheet **`Staging_Pivots`** (or locate the `pt_Agent` PivotTable).
2. Right-click any numeric cell in the **`Avg Speed (s)`** column (e.g. cell `G4`).
3. Click **Number Format...** (⚠️ *Crucial*: Do **NOT** click *Format Cells...*. Selecting *Number Format...* updates the Data Model PivotField definition permanently across all slicer updates).
4. In the **Category** list on the left, click **Custom**.
5. In the **Type:** text box at the top, clear whatever is there and type:
   ```excel
   0.0 "s"
   ```
   *(Or enter `#,##0.0 "s"` if you prefer comma separators).*
6. Click **OK**.
7. **Verification**: Confirm that Becky now reads **`65.3 s`**, Dan reads **`67.3 s`**, and Joe reads **`71.0 s`**. The `6533.1%` bug is permanently resolved!

---

#### Step 2: Remove the AutoFilter Dropdown Arrow (`Agent [v]`)
1. Right-click anywhere inside the `pt_Agent` PivotTable.
2. Click **PivotTable Options...** at the bottom of the context menu.
3. In the dialog box, click the **Display** tab.
4. Under the **Display** section, **UNCHECK** the checkbox:
   > 🔲 **Show field captions and filter drop downs**
5. Click the **Totals & Filters** tab $\to$ ensure **Show grand totals for rows** and **Show grand totals for columns** are both **UNCHECKED**.
6. Click **OK**.
7. **Result**: The gray filter arrow button `[v]` disappears completely from the `Agent` header! The table header now displays crisp, clean bold text: **`Agent`**, matching native web application tables.

---

#### Step 3: Calibrate Column Widths & Row Heights on `Staging_Pivots` (For 1:1 Box Fit)
To guarantee that the table fits the $448\\text{ pt} \\times 212\\text{ pt}$ container without distortion or scraping, calibrate the source cells on `Staging_Pivots`:

1. **Configure Exact Column Widths**:
   - Right-click column **`C`** (`Agent`) $\to$ click **Column Width...** $\to$ set to **`12.5`** (~80 pt).
   - Right-click column **`D`** (`Calls Taken`) $\to$ click **Column Width...** $\to$ set to **`10.5`** (~70 pt).
   - Right-click column **`E`** (`Answer Rate %`) $\to$ click **Column Width...** $\to$ set to **`11.5`** (~74 pt).
   - Right-click column **`F`** (`FCR Rate %`) $\to$ click **Column Width...** $\to$ set to **`10.5`** (~70 pt).
   - Right-click column **`G`** (`Avg Speed (s)`) $\to$ click **Column Width...** $\to$ set to **`12.0`** (~78 pt).
   - Right-click column **`H`** (`Avg CSAT`) $\to$ click **Column Width...** $\to$ set to **`10.5`** (~70 pt).
   *Total Table Width*: Exactly **`442 pt`** (leaves balanced 3 pt left and right padding inside the 448 pt container).

2. **Configure Exact Row Heights**:
   - Right-click row **`3`** (Header Row) $\to$ click **Row Height...** $\to$ set to **`24 pt`**.
   - Select rows **`4` through `11`** (the 8 Agent Rows) $\to$ right-click $\to$ click **Row Height...** $\to$ set to **`22.5 pt`**.
   *Total Table Height*: $24\\text{ pt} + (8 \\times 22.5\\text{ pt}) = \\mathbf{204\\text{ pt}}$ (leaves balanced 4 pt top and bottom padding inside the 212 pt container).

3. **Apply Modern Web-App Cell Styles**:
   - **Header Row (`C3:H3`)**:
     - Background Fill: Solid Dark Slate `#0F172A` (RGB `15, 23, 42`).
     - Font: **Segoe UI**, Size: **8.5 pt**, Font Style: **Bold**, Font Color: **Crisp White (`#FFFFFF`)**.
     - Alignment: Vertical **Center**; Column C **Left**, Columns D-H **Right**.
   - **Data Rows (`C4:H11`)**:
     - Font: **Segoe UI**, Size: **8.5 pt**, Color: Dark Charcoal `#1E293B`.
     - Alternating Row Fill: Odd rows `#FFFFFF` (Crisp White); Even rows `#F8FAFC` (Soft Slate).
     - Column C (`Agent` Names): Font Style: **Bold**, Color: `#0F172A`, Alignment: **Left**.
     - Horizontal Gridlines: Select `C3:H11` $\to$ Borders $\to$ **Inside Horizontal** $\to$ Line Style: **Continuous Thin**, Color: `#E2E8F0` (RGB `226, 232, 240`). Remove all vertical borders!

---

#### Step 4: Transform the Container Box (From Dashed Blueprint to Modern Solid Card)
Switch to sheet **`03_CallCenter_Cockpit`**:

1. **Outer Container Card (`Container_CC_AgentScorecard`)**:
   - Select the card outline (`Left = 762 pt, Top = 510 pt, Width = 476 pt, Height = 270 pt`).
   - In ribbon tab **Shape Format**:
     - **Shape Fill**: Solid White (`#FFFFFF`).
     - **Shape Outline**: Solid Line, Color: `#E2E8F0`, Weight: `1 pt`.
     - **Shape Effects** $\to$ **Shadow** $\to$ Presets: **Outer Offset Bottom** (`msoShadow21`):
       - *Color*: `#0F172A` | *Transparency*: **88%** | *Size*: **100%** | *Blur*: **10 pt** | *Distance*: **3.5 pt**.
       - *Result*: Produces a high-end SaaS glassmorphic elevation effect.

2. **Inner Docking Frame (`DockZone_CC_AgentScorecard`)**:
   - Click the inner box inside the card.
   - **Clear Watermark Text**: Click inside the text $\to$ press `Ctrl + A` $\to$ press `Delete` (leave shape empty).
   - In ribbon tab **Shape Format**:
     - **Shape Outline** $\to$ **Dashes**: Select **Solid** (⚠️ *Replace the legacy dashed line with a clean solid line!*).
     - **Outline Color**: Soft Gray `#E2E8F0` (RGB `226, 232, 240`), Weight: **0.75 pt**.
     - **Shape Fill**: Solid White (`#FFFFFF`).
     - **Position & Geometry**: Set Left: **`776 pt`**, Top: **`546 pt`**, Width: **`448 pt`**, Height: **`212 pt`**.
   - *Result*: The inner box becomes a subtle, modern recessed surface framing the table.

---

#### Step 5: Embed the Modern Vector SVG Icon in the Card Header
1. On sheet **`03_CallCenter_Cockpit`**, click ribbon tab: **Insert** $\to$ **Pictures** $\to$ **This Device...**
2. Navigate to: `assets/icons/icon_audit_matrix.svg`. Click **Insert**.
3. Select the inserted SVG picture:
   - In ribbon tab **Graphics Format**: Set Height = **`0.25 in`** (`18 pt`), Width = **`0.25 in`** (`18 pt`).
   - In the Name Box (top-left, above cell A1), rename the shape to: **`Icon_CC_AgentScorecard`**.
   - Position the icon at: Left: **`778 pt`**, Top: **`524 pt`**.
4. Click on the header text shape `Header_CC_AgentScorecard`:
   - Move its Left position to **`804 pt`** (so the text sits cleanly 8pt to the right of the icon).
5. Click on the top-right badge `Badge_CC_AgentScorecard`:
   - In ribbon tab **Shape Format** $\to$ Shape Fill: `#F1F5F9`, Shape Outline: `#E2E8F0`.
   - Text: Edit to read: **`8 Agents Active`** (Font: Segoe UI 7.5 pt Bold, Color: `#475569`).

---

#### Step 6: Dock the Live Table as a Pixel-Perfect Linked Picture
1. Switch to **`Staging_Pivots`**.
2. Select the calibrated range: **`C3:H11`**.
3. Press **`Ctrl + C`** (Copy).
4. Switch to **`03_CallCenter_Cockpit`**.
5. Click cell **`A1`**.
6. On ribbon tab **Home** $\to$ click the small dropdown arrow below **Paste** $\to$ select the very last icon:
   > 🔗 **Linked Picture** *(Clipboard with picture and chain link)*
7. With the newly pasted picture selected:
   - In the Name Box (top-left), rename it to: **`LiveScorecard_HTMLTable`**.
   - In ribbon tab **Picture Format**:
     - Check **Lock Aspect Ratio**.
     - Set Width: **`6.17 in`** (**`444 pt`**), Height: **`2.83 in`** (**`204 pt`**).
   - Drag the picture over `DockZone_CC_AgentScorecard`:
     - Set Position: Left: **`778 pt`**, Top: **`550 pt`**.
     - Right-click picture $\to$ **Bring to Front**.
8. **Verify the Fit**:
   - Notice that the table now sits with balanced 2 pt side padding and 4 pt top/bottom padding!
   - Stewart's row has comfortable breathing room above the bottom border.
   - The table header aligns with the card header.
   - Click any Slicer (`Topic`, `Agent`, `Month`) $\to$ the table updates instantly with live data!

---

#### Step 7: Format the Other 3 Visuals for a Cohesive Web-App Experience

To ensure all visuals match the executive web-app design system, format the remaining 3 PivotCharts on `03_CallCenter_Cockpit`:

```
+---------------------------------------------------------------------------------------------------------+
| Visual 1.1: CC_HourlyVolume                     | Visual 1.3: CC_AgentQuadrant                          |
| Clustered Column: #1E293B & #059669             | Combo: Blue Columns (#3B82F6) + Tangerine Line (#D04A02)|
+-------------------------------------------------+-------------------------------------------------------+
| Visual 1.2: CC_TopicBreakdown                   | Visual 1.4: CC_AgentScorecard                         |
| Horizontal Bar: #D04A02 | Categories Reversed   | Docked Live HTML Matrix | 0.0 "s" Speed Verified      |
+---------------------------------------------------------------------------------------------------------+
```

##### Visual 1.1: Intraday Demand Surge & Queue Triage (`CC_HourlyVolume`)
1. **Insert Header SVG Icon**: Insert `assets/icons/icon_hourly_surge.svg` at Left: `286 pt`, Top: `240 pt` (18x18 pt). Shift header text to Left: `312 pt`.
2. **Chart Type**: 2-D Clustered Column (`Call_Hour` on Axis; `Total Calls` and `Answered Calls` in Values).
3. **Format Series**:
   - Click Series 1 (`Total Calls`) $\to$ Fill: Solid Dark Slate `#1E293B`, Border: No line.
   - Click Series 2 (`Answered Calls`) $\to$ Fill: Solid Emerald `#059669`, Border: No line.
   - Right-click bars $\to$ **Format Data Series...** $\to$ set **Series Overlap = 0%**, **Gap Width = 65%**.
4. **Transparent Decluttering**:
   - Chart Area & Plot Area: Fill = **No fill**, Border = **No line**.
   - Horizontal Gridlines: Solid Line, Color: `#F1F5F9`, Width: `0.5 pt`.
   - Legend: Move to **Top**, Font: Segoe UI 8 pt `#64748B`, No fill, No border.
   - Select `DockZone_CC_HourlyVolume` $\to$ clear watermark text, set outline to solid `#E2E8F0` or hide.

##### Visual 1.2: Inquiry Topic SLA Compliance & Speed (`CC_TopicBreakdown`)
1. **Insert Header SVG Icon**: Insert `assets/icons/icon_topic_sla.svg` at Left: `286 pt`, Top: `524 pt` (18x18 pt). Shift header text to Left: `312 pt`.
2. **Chart Type**: 2-D Clustered Bar (Horizontal).
3. **Reverse Category Order**:
   - Right-click vertical topic axis $\to$ **Format Axis...** $\to$ check **"Categories in reverse order"**.
   - *Result*: Technical Support (largest volume) appears at the top.
4. **Format Series**:
   - Bar Fill: Solid Tangerine `#D04A02`, Border: No line.
   - Gap Width: **55%**.
   - Right-click bars $\to$ **Add Data Labels** $\to$ Label Position: **Outside End**, Font: Segoe UI 8 pt Bold `#0F172A`.
5. **Transparent Decluttering**: Chart Area Fill = No fill, Border = No line. Clear `DockZone_CC_TopicBreakdown`.

##### Visual 1.3: Representative Efficiency Matrix (`CC_AgentQuadrant`)
1. **Insert Header SVG Icon**: Insert `assets/icons/icon_agent_quadrant.svg` at Left: `778 pt`, Top: `240 pt` (18x18 pt). Shift header text to Left: `804 pt`.
2. **Chart Type**: **Combo Chart**:
   - Series 1 (`FCR Rate %`): **Clustered Column** on **Primary Axis** $\to$ Fill: `#3B82F6` (Executive Blue), No border.
   - Series 2 (`Avg CSAT`): **Line with Markers** on **Secondary Axis** $\to$ Line: `#D04A02` 2 pt; Marker: Circle 6 pt `#D04A02` with white 1.5 pt border.
3. **Axis Scaling**:
   - Primary Vertical Axis: Right-click $\to$ Format Axis $\to$ Minimum = `0.70` (70%), Maximum = `1.00` (100%).
   - Secondary Vertical Axis: Right-click $\to$ Format Axis $\to$ Minimum = `2.50`, Maximum = `4.00`.
4. **Data Labels**: Right-click secondary line $\to$ Add Data Labels $\to$ Above markers displaying CSAT score (Dan: 3.45, Martha: 3.47).
5. **Transparent Decluttering**: Chart Area Fill = No fill, Border = No line. Clear `DockZone_CC_AgentQuadrant`.

---

### ⚡ Fast-Track Option: 1-Click Complete Execution via VBA

If you ever wish to execute all of the above steps programmatically in 0.1 seconds, use the upgraded automation modules:

1. Press **`Alt + F11`** to open the VBA Editor.
2. In the menu bar, import:
   - [`vba/modInteractiveScorecard.bas`](../vba/modInteractiveScorecard.bas)
   - [`vba/modPivotTableFormatting.bas`](../vba/modPivotTableFormatting.bas)
   - [`vba/modDashboardUIUX.bas`](../vba/modDashboardUIUX.bas)
3. Press **`Ctrl + G`** to open the Immediate Window and run:

```vba
' 1. Rebuild & dock scorecard with pixel-perfect fit & AutoFilter hidden
Call modInteractiveScorecard.BuildAndDockInteractiveScorecard

' 2. Apply modern Web-App theme to the other 3 charts
Call modDashboardUIUX.ApplyModernChartTheme(ActiveSheet.ChartObjects("CC_HourlyVolume"), "HOURLY_VOLUME")
Call modDashboardUIUX.ApplyModernChartTheme(ActiveSheet.ChartObjects("CC_TopicBreakdown"), "TOPIC_BREAKDOWN")
Call modDashboardUIUX.ApplyModernChartTheme(ActiveSheet.ChartObjects("CC_AgentQuadrant"), "AGENT_QUADRANT")
```

---

### 🌐 Standalone Executive Web Dashboard Companion (`HTML/SVG`)

For client presentations and web demonstrations, an interactive, fully functional SaaS dashboard web application has been authored in:
[`dashboards/interactive_call_center_dashboard.html`](../dashboards/interactive_call_center_dashboard.html)

**Features Included**:
- **Pure Modern Web UI**: Responsive CSS Grid / Bento Grid with glassmorphism (`backdrop-filter: blur(16px)`), linear gradient accents, and soft multi-layer box shadows.
- **Full SVG Icon Suite**: Integrated Lucide-style vector paths for all actions, KPIs, and card headers.
- **Interactive Slicer Chips**: Click on months (`All Q1`, `January`, `February`, `March`) or topics to see simulated real-time filtering.
- **Pixel-Perfect Data Table**: Demonstrates the exact target styling: `#0F172A` dark header, alternating `#F8FAFC` zebra rows, verified `65.3 s` speed formatting, and green/amber/gold status pill badges!

---

"""

# Find start of Visual 1.4
start_marker = "###### Visual 1.4: Representative Quality & CSAT Audit (Right-Bottom)"
end_marker = "##### Dashboard 2: Customer Retention & Revenue Risk Cockpit (`04_CustomerRetention_Cockpit`)"

idx_start = content.find(start_marker)
idx_end = content.find(end_marker)

if idx_start != -1 and idx_end != -1:
    new_content = content[:idx_start] + replacement_text + content[idx_end:]
    with open('docs/00_master_project_execution_guide.md', 'w', encoding='utf-8') as f:
        f.write(new_content)
    print("00_master_project_execution_guide.md successfully updated!")
else:
    print(f"Error: Markers not found. idx_start={idx_start}, idx_end={idx_end}")
