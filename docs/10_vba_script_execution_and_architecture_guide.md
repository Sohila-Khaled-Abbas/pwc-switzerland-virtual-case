# PwC Switzerland Virtual Case Experience - VBA Execution & Multi-Module Architecture Guide

**Platform Version**: `v3.0.0 Enterprise`  
**Repository**: [pwc-switzerland-virtual-case](https://github.com/Sohila-Khaled-Abbas/pwc-switzerland-virtual-case)  
**Encoding Standard**: 100% Pure 7-bit ASCII (`ChrW(9650)` / `ChrW(9660)`)  
**Design Reference**: Executive SaaS Cockpit UI/UX with 7 BAN KPI Cards, Intraday Arrival Surge & Live Scorecard Docking  

---

## 1. Executive Architecture Overview

The PwC Switzerland Business Intelligence platform is structured into **12 specialized, interconnected Visual Basic for Applications (VBA) modules**. Each module is responsible for a single architectural boundary, ensuring clean separation of concerns, zero naming collisions, and deterministic execution.

```
                              +---------------------------------------+
                              |        modPwC_Unified_Master          |
                              |      (Master Suite Orchestrator)      |
                              +-------------------+-------------------+
                                                  |
         +--------------------+-------------------+--------------------+--------------------+
         |                    |                   |                    |                    |
+--------v-------+   +--------v-------+  +--------v-------+   +--------v-------+   +--------v-------+
|  modAppState   |   | modThemeEngine |  |  modNavigation |   |  modDataRefresh|   |modFilterContr..|
| (State Shield) |   | (Light / Dark) |  | (Fast Switch)  |   | (VertiPaq Sync)|   | (Clear Slicers)|
+----------------+   +----------------+  +----------------+   +----------------+   +----------------+
         |                    |                   |                    |                    |
+--------v-------+   +--------v-------+  +--------v-------+   +--------v-------+   +--------v-------+
|   modExportPDF |   |modCreateGov... |  |modPortalLanding|   |modDashboardUIUX|   |modPivotTable...|
| (A4 Publisher) |   | (01 & 02 Arch) |  | (00 Executive) |   | (03,04,05 UI)  |   | & Scorecards   |
+----------------+   +----------------+  +----------------+   +----------------+   +----------------+
```

### Module Directory & Responsibilities

| # | Module Name | Primary Procedures | Description |
|---|-------------|--------------------|-------------|
| **01** | `modAppState.bas` | `FreezeAppState`, `RestoreAppState` | Performance shield; deterministically freezes `ScreenUpdating`, `EnableEvents`, and calculation during macro execution. |
| **02** | `modThemeEngine.bas` | `ToggleDashboardTheme`, `ApplyTheme`, `ApplyDarkTheme`, `ApplyLightTheme` | Dynamic enterprise Light/Dark theme switcher with persistent document properties. |
| **03** | `modNavigation.bas` | `NavigateToHomePortal`, `NavigateToCallCenter`, `NavigateToRetention`, `NavigateToDiversity` | Zero-flicker viewport navigator across all 6 analytical and governance canvas sheets. |
| **04** | `modDataRefresh.bas` | `RefreshPipelineSynchronously` | Refreshes VertiPaq tabular model PivotTables and full `CUBEVALUE` metric chains. |
| **05** | `modFilterController.bas` | `ClearAllFilters` | Clears all active slicer caches and PivotTable filters across the workbook. |
| **06** | `modExportPDF.bas` | `ExportActiveDashboardPDF`, `ExportExecutiveReport` | Generates publication-grade A4 Landscape executive PDF briefings. |
| **07** | `modCreateGovernanceSheets.bas` | `BuildGovernanceArchitecture` | Generates `01_Business_Domains` briefing cards and `02_Metadata_&_KPI_Catalog` data dictionary. |
| **08** | `modPortalLanding.bas` | `BuildExecutivePortal` | Creates `00_Home_Portal` executive landing suite with branded KPI launchers and SVG icons. |
| **09** | `modDashboardUIUX.bas` | `BuildCallCenterCanvas`, `BuildCustomerRetentionCanvas`, `BuildDiversityInclusionCanvas`, `BuildAllDashboardCanvases` | Core UI/UX canvas builder, BAN cards, slicer docks, real chart generators, and staging engine. |
| **10** | `modInteractiveScorecard.bas` | `BuildAndDockInteractiveScorecard`, `CreateScorecardPivotTable` | Builds dynamic PivotTable scorecards with HTML/conditional formatting. |
| **11** | `modPivotTableFormatting.bas` | `StyleAgentScorecardPivotTable`, `DockScorecardAsLinkedPicture` | Applies executive conditional formatting, inverted speed alerts, and linked picture docking. |
| **12** | `modPwC_Unified_Master.bas` | `RunUnifiedPwCPlatform`, `RunCompletePwCPlatform`, `VerifyPlatformIntegrity` | High-level Master Orchestrator coordinating all modules in an automated sequence. |

---

## 2. Step-by-Step Manual Execution Guide

Follow these steps to run the scripts manually in Microsoft Excel without any errors:

### Step 1: Open the Macro-Enabled Workbook
1. Open `PWC_Switzerland_Virtual_Case.xlsm` in Microsoft Excel.
2. If prompted, click **"Enable Content"** or **"Enable Macros"**.
3. Verify that Power Pivot / Data Model access is trusted.

### Step 2: Open the Visual Basic for Applications Editor
1. Press `Alt + F11` to open the VBA Editor window.
2. In the **Project Explorer** (left panel), verify that the 12 modules appear under `Modules`:
   - `modAppState`
   - `modCreateGovernanceSheets`
   - `modDashboardUIUX`
   - `modDataRefresh`
   - `modExportPDF`
   - `modFilterController`
   - `modInteractiveScorecard`
   - `modNavigation`
   - `modPivotTableFormatting`
   - `modPortalLanding`
   - `modPwC_Unified_Master`
   - `modThemeEngine`

### Step 3: Compile the Project (Zero Conflict Verification)
1. In the top menu, click **Debug** -> **Compile VBAProject**.
2. **Expected Result**: The menu item becomes disabled/grayed out.  
   - There are **0 ambiguous name conflicts**.
   - There are **0 duplicate procedure or constant names**.
   - All cross-module procedure references resolve cleanly.

### Step 4: Execute the Master Orchestrator
You can run the deployment in either of two ways:

#### Option A: From Excel Macro Dialog (Recommended)
1. Switch back to Excel (`Alt + F11` or click the Excel window).
2. Press `Alt + F8` to open the **Macro** dialog.
3. Select **`RunUnifiedPwCPlatform`** (or `RunCompletePwCPlatform`).
4. Click **Run**.

#### Option B: From the VBA Editor
1. Double-click `modPwC_Unified_Master` in Project Explorer.
2. Place the cursor inside `Sub RunUnifiedPwCPlatform()`.
3. Press `F5` (or click the green **Play** button).

### Step 5: Verify Deployment Success
Upon completion, the system displays an information box:
```
PwC Switzerland BI Platform deployed successfully!

- Connected Architecture: All 12 VBA modules fully synchronized
- Interactive Visuals: Real charts, BAN cards & scorecards generated
- Zero Conflicts: 0 ambiguous names and 0 shape deletion errors
- Responsive Design: Modern rounded buttons with centered text
- Pure ASCII: 100% compatible across all Windows regional locales
```
The active view automatically lands on **`03_CallCenter_Cockpit`** with all charts, KPI BAN cards, and navigation buttons fully rendered.

---

## 3. UI/UX Interaction & Button Bindings

All buttons created on worksheets are permanently bound to their respective modules via explicit `.OnAction` properties:

| Button / Control | Location | Linked Procedure | Action Performed |
|------------------|----------|------------------|------------------|
| **"Portal Home"** | Top Nav Bar (Left) | `modNavigation.NavigateToHomePortal` | Switches to `00_Home_Portal` with scroll reset |
| **"01 Call Center"** | Top Nav Bar | `modNavigation.NavigateToCallCenter` | Switches to `03_CallCenter_Cockpit` |
| **"02 Retention"** | Top Nav Bar | `modNavigation.NavigateToRetention` | Switches to `04_CustomerRetention_Cockpit` |
| **"03 Diversity"** | Top Nav Bar | `modNavigation.NavigateToDiversity` | Switches to `05_DiversityInclusion_Cockpit` |
| **"Refresh Data"** | Top Nav Bar (Right) | `modDataRefresh.RefreshPipelineSynchronously` | Recalculates VertiPaq tabular model |
| **"Clear Slicers"** | Top Nav Bar & Slicer Panel | `modFilterController.ClearAllFilters` | Clears all manual slicer selections |
| **"Theme Mode"** | Top Nav Bar (Right) | `modThemeEngine.ToggleDashboardTheme` | Switches between Light and Dark visual theme |
| **"Export PDF"** | Top Nav Bar (Right) | `modExportPDF.ExportActiveDashboardPDF` | Publishes A4 landscape PDF in workbook folder |

---

## 4. Key Bug Fixes & Technical Resolutions

### Fix 1: Runtime Error `-2147024809` (0x80070057) Resolved
- **Root Cause**: Previous code called `.Delete` directly on named shapes or charts (e.g., `ws.Shapes("Resolution_CenterBadge").Delete` or `ws.ChartObjects("cht_TrendDaily").Delete`) after an `On Error GoTo 0` reset. On first run, because these objects do not yet exist, Excel threw runtime error `-2147024809: "The item with the specified name wasn't found."`
- **Resolution**: Implemented `SafeDeleteShape` and `SafeDeleteChart` helper procedures:
  ```vba
  Public Sub SafeDeleteShape(ByVal ws As Worksheet, ByVal shapeName As String)
      On Error Resume Next
      ws.Shapes(shapeName).Delete
      On Error GoTo 0
  End Sub
  
  Public Sub SafeDeleteChart(ByVal ws As Worksheet, ByVal chartName As String)
      On Error Resume Next
      ws.ChartObjects(chartName).Delete
      On Error GoTo 0
  End Sub
  ```
  Every chart, shape, and pill deletion is now 100% fail-safe.

### Fix 2: "Ambiguous Name Detected" Eliminated
- **Root Cause**: Having procedure declarations duplicated across `modPwC_Unified_Master` and individual specialized modules (`modDashboardUIUX`, `modThemeEngine`, `modNavigation`, etc.) caused Excel's VBE compiler to halt with ambiguous name errors.
- **Resolution**:
  - `modPwC_Unified_Master.bas` is strictly the Master Orchestrator, calling public procedures from the dedicated modules.
  - All constants inside `modDashboardUIUX.bas` and `modThemeEngine.bas` are scoped as `Private Const`.
  - `scan_vba_conflicts.py` verifies **0 duplicate procedures** and **0 duplicate constants**.

### Fix 3: Pure 7-bit ASCII Encoding Enforced
- All UTF-8 em-dashes (`—`), smart quotes (`“`, `”`), and directional triangles (`▲`, `▼`) have been replaced with pure ASCII representations (`--`, `"`, `ChrW(9650)`, `ChrW(9660)`).
- This guarantees zero character corruption on Windows-1256 (Arabic) or other international Windows code pages.

### Fix 4: Scorecard PivotTable Formatter Modal Dialog Eliminated
- **Root Cause**: `modPivotTableFormatting.bas` attempted to find `ws.PivotTables("pt_Agent")` on `03_CallCenter_Cockpit` or `Staging_Pivots`. If the user ran the formatter before provisioning, a blocking modal `MsgBox` ("No PivotTable found...") was raised, interrupting automated batch execution.
- **Resolution**: `StyleAgentScorecardPivotTable` now calls `modInteractiveScorecard.EnsureOrBuildAgentPivotTable` to automatically provision `pt_Agent` on `Staging_Pivots` on demand, and cleanly exits if uninitialized without blocking alerts.

### Fix 5: Compile Error "Variable not defined" (`wb`) Resolved
- **Root Cause**: With `Option Explicit` enabled, `RunUnifiedPwCPlatform` referenced `wb.Worksheets("Staging_Pivots")` without a local `Dim wb As Workbook` declaration.
- **Resolution**: Explicitly dimmed and initialized `Dim wb As Workbook: Set wb = ThisWorkbook: If wb Is Nothing Then Set wb = ActiveWorkbook`. Verified with 0 compile errors in Excel VBE.

### Fix 6: Multi-Tier Asset & SVG Icon Path Resolution
- **Root Cause**: Relative icon paths failed when running from varying current directories or subfolders (`assets/icons/web/` vs `assets/icons/`).
- **Resolution**: Implemented `ResolveAssetPath` in `modDashboardUIUX.bas` which resolves against workbook path, relative subfolders, parent paths, and absolute project paths, with automatic aliasing of `"logo"` to `assets/PwC_logo_rgb_colour_pos.png`.

---

## 5. Slicer & Filter Engine Architecture (`modFilterController.bas`)

The platform deploys real, interactive Excel Data Model Slicers (`SlicerCaches.Add2` & `Slicers.Add`) docked into the analytical cockpits:

| Cockpit Sheet | Slicer Slot 1 | Slicer Slot 2 | Slicer Slot 3 |
|---------------|---------------|---------------|---------------|
| `03_CallCenter_Cockpit` | `[DimDate].[Month]` (`Billing Month`) | `[DimTopic].[Topic]` (`Topic Tier`) | `[DimAgent].[Agent]` (`Representative`) |
| `04_CustomerRetention_Cockpit` | `[DimContract].[Contract]` (`Contract Type`) | `[Fact_Churn].[PaymentMethod]` (`Payment Gateway`) | `[Fact_Churn].[InternetService]` (`Internet Service`) |
| `05_DiversityInclusion_Cockpit` | `[DimDepartment].[Department]` (`Department`) | `[Fact_Employees].[JobLevel]` (`Job Level Hierarchy`) | `[Fact_Employees].[Gender]` (`Gender Demographics`) |

All slicers are created with fallback protection: if VertiPaq OLAP cubes are not yet connected, staging table columns are used seamlessly.

---

## 6. VBA MCP Server Integration

The project includes the **VBA MCP Server** located in `scripts/vba_mcp_server/`, configured in Antigravity IDE:

```json
"vba-mcp-server": {
  "command": "python",
  "args": [
    "scripts/vba_mcp_server/server.py"
  ]
}
```

### Capabilities:
- Direct programmatic inspection of VBA modules inside `.xlsm` workbooks via Windows COM (`win32com.client`).
- Automated module backup and manifest generation.
- Zero manual copy-pasting required for ongoing maintenance.
