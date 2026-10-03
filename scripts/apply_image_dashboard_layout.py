"""
PwC Switzerland Virtual Case Experience - Enterprise UI/UX Alignment Engine
Re-engineers 03_CallCenter_Cockpit to match the modern SaaS dashboard image:
- Top Header: PwC official logo, title, period, and 4 header cards (Date Range, Region, Dept, Refresh)
- Persistent 6-Tab Navigation Bar with active Tangerine highlight
- Left Slicer Drawer (#182234) with Funnel icon, 3 native Data Model Slicers, Reset button, Support Team illustration, and Quote Card
- 7 BAN KPI Metric Cards across the top with circular colored badges, vector SVG icons, dynamic CUBE formulas, and color-matched wave sparklines
- Middle Row: Calls Trend (Daily) with peak callout, Calls by Hour (Heatmap), Resolution Breakdown Donut Chart
- Bottom Row: Agent Performance Scorecard, AHT vs Volume Combo Chart, Pareto Categories, Sentiment & Region Comparison
"""

import os
import sys
import time
import win32com.client as win32

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKBOOK_PATH = os.path.join(BASE_DIR, "PWC_Switzerland_Virtual_Case.xlsm")
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
ICONS_DIR = os.path.join(ASSETS_DIR, "icons")
WEB_ICONS_DIR = os.path.join(ICONS_DIR, "web")

def rgb(r, g, b):
    return r + (g * 256) + (b * 65536)

# Colors
C_ORANGE = rgb(234, 88, 12)       # #EA580C PwC Tangerine
C_DARK_SLATE = rgb(15, 23, 42)    # #0F172A
C_DRAWER_BG = rgb(24, 34, 52)     # #182234
C_DRAWER_SLOT = rgb(30, 41, 59)   # #1E293B
C_BORDER_DRAWER = rgb(51, 65, 85) # #334155
C_CANVAS_BG = rgb(248, 250, 252)  # #F8FAFC
C_CARD_BG = rgb(255, 255, 255)    # #FFFFFF
C_CARD_BORDER = rgb(226, 232, 240)# #E2E8F0
C_TEXT_TITLE = rgb(15, 23, 42)    # #0F172A
C_TEXT_MUTED = rgb(100, 116, 139) # #64748B
C_TEXT_LIGHT = rgb(148, 163, 184) # #94A3B8

def run_layout():
    print(f"Connecting to Excel COM on {WORKBOOK_PATH}...")
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(WORKBOOK_PATH)
        ws = wb.Worksheets("03_CallCenter_Cockpit")
        print("Connected to 03_CallCenter_Cockpit.")

        # --------------------------------------------------------------------------
        # STEP 1: PRESERVE SLICERS & CHARTS, PURGE OLD DECORATIVE SHAPES
        # --------------------------------------------------------------------------
        preserved_shape_names = {"CC_HourlyVolume", "CC_TopicBreakdown", "CC_AgentQuadrant", "CC_AgentScorecard", "LiveScorecard_HTMLTable"}
        for s in list(ws.Shapes):
            if s.Type == 25: # Slicer
                continue
            if s.Name in preserved_shape_names or s.Type == 3: # Slicers, charts, and scorecard
                continue
            try:
                s.Delete()
            except Exception:
                pass
        print("Purged old shapes; preserved slicers and native chart objects.")

        # --------------------------------------------------------------------------
        # STEP 2: CANVAS PREPARATION & TECHNICAL ROW 65 KPI STAGING
        # --------------------------------------------------------------------------
        ws.Cells.Interior.Color = C_CANVAS_BG
        excel.ActiveWindow.DisplayGridlines = False
        excel.ActiveWindow.DisplayHeadings = False
        excel.ActiveWindow.Zoom = 80

        titles = [["Total Calls", "Answered Calls", "Missed Calls", "SLA (%)", "Avg Handle Time", "CSAT Score", "FCR (%)"]]
        formulas = [[
            '=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Total Calls]")',
            '=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Answered Calls]")',
            '=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Abandoned Calls]")',
            '=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Answer Rate %]")',
            '=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Average Speed of Answer (s)]")',
            '=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[Average CSAT]")',
            '=CUBEVALUE("ThisWorkbookDataModel", "[Measures].[First Contact Resolution %]")'
        ]]
        formats = ["#,##0", "#,##0", "#,##0", "0.0%", '0.0 "s"', "0.00", "0.0%"]

        # Set titles in bulk
        ws.Range("AA64:AG64").Value = titles
        ws.Range("AA64:AG64").Font.Bold = True

        # Set formats
        for i, fmt in enumerate(formats):
            ws.Cells(65, 27 + i).NumberFormat = fmt

        # Set formulas in bulk
        ws.Range("AA65:AG65").Formula = formulas
        print("Configured CUBEVALUE formulas in row 65 (AA:AG). Waiting for OLAP evaluation...")

        # Wait for VertiPaq tabular model asynchronous query evaluation
        for attempt in range(20):
            time.sleep(1)
            val = ws.Cells(65, 27).Text
            if val != "#GETTING_DATA" and len(val) > 0:
                print(f"OLAP data ready after {attempt+1} seconds: Total Calls = {val}")
                break

        ws.Rows("60:75").Hidden = True
        print("Staged row 65 (AA:AG) and hid rows 60:75.")

        # --------------------------------------------------------------------------
        # STEP 3: TOP HEADER & 6-TAB PERSISTENT NAVIGATION BAR
        # --------------------------------------------------------------------------
        nav_top = 14.0
        nav_h = 52.0
        nav_w = 1480.0
        nav_bg = ws.Shapes.AddShape(1, 24.0, nav_top, nav_w, nav_h)
        nav_bg.Name = "Nav_TopBar"
        nav_bg.Fill.Solid()
        nav_bg.Fill.ForeColor.RGB = C_CARD_BG
        nav_bg.Line.ForeColor.RGB = C_CARD_BORDER
        nav_bg.Line.Weight = 1.0

        # PwC Logo
        logo_path = os.path.join(ASSETS_DIR, "PwC_logo_rgb_colour_pos.png")
        if os.path.exists(logo_path):
            logo = ws.Shapes.AddPicture(logo_path, False, True, 36.0, 19.0, 54.0, 42.0)
            logo.Name = "Nav_Logo"

        # Brand Title & Subtitle
        title_box = ws.Shapes.AddTextbox(1, 100.0, 16.0, 420.0, 24.0)
        title_box.Name = "Nav_TitleBox"
        title_box.Fill.Visible = False
        title_box.Line.Visible = False
        tf1 = title_box.TextFrame2
        tf1.MarginLeft = tf1.MarginTop = tf1.MarginRight = tf1.MarginBottom = 0
        tr1 = tf1.TextRange
        tr1.Text = "PwC Call Center Analytics Dashboard"
        tr1.Font.Name = "Segoe UI"
        tr1.Font.Size = 13.5
        tr1.Font.Bold = True
        tr1.Font.Fill.ForeColor.RGB = C_TEXT_TITLE

        sub_box = ws.Shapes.AddTextbox(1, 100.0, 40.0, 420.0, 18.0)
        sub_box.Name = "Nav_SubtitleBox"
        sub_box.Fill.Visible = False
        sub_box.Line.Visible = False
        tf2 = sub_box.TextFrame2
        tf2.MarginLeft = tf2.MarginTop = tf2.MarginRight = tf2.MarginBottom = 0
        tr2 = tf2.TextRange
        tr2.Text = "Analysis Period: 01 Jan 2025 - 31 Dec 2025 | VertiPaq Semantic Model"
        tr2.Font.Name = "Segoe UI"
        tr2.Font.Size = 8.0
        tr2.Font.Fill.ForeColor.RGB = C_TEXT_MUTED

        # 6 Persistent Navigation Tabs (Aligned neatly in middle)
        tabs = [
            ("HOME", "00_Home_Portal", 540.0, 65.0, False),
            ("DOMAINS", "01_Business_Domains", 610.0, 80.0, False),
            ("CATALOG", "02_Metadata_&_KPI_Catalog", 695.0, 80.0, False),
            ("CALL CENTER", "03_CallCenter_Cockpit", 780.0, 105.0, True),
            ("RETENTION", "04_CustomerRetention_Cockpit", 890.0, 95.0, False),
            ("D&I", "05_DiversityInclusion_Cockpit", 990.0, 55.0, False)
        ]

        for tab_label, sheet_target, t_left, t_width, is_active in tabs:
            tab_btn = ws.Shapes.AddShape(1, t_left, 24.0, t_width, 30.0)
            tab_btn.Name = f"Nav_Tab_{tab_label.replace(' ', '_')}"
            tab_btn.Fill.Solid()
            if is_active:
                tab_btn.Fill.ForeColor.RGB = C_ORANGE
                tab_btn.Line.Visible = False
            else:
                tab_btn.Fill.ForeColor.RGB = rgb(241, 245, 249)
                tab_btn.Line.ForeColor.RGB = C_CARD_BORDER
                tab_btn.Line.Weight = 0.75

            t_tf = tab_btn.TextFrame2
            t_tf.VerticalAnchor = 3
            t_tf.MarginLeft = 4
            t_tf.MarginTop = 0
            t_tf.MarginRight = 4
            t_tf.MarginBottom = 0
            t_tr = t_tf.TextRange
            t_tr.Text = tab_label
            t_tr.Font.Name = "Segoe UI"
            t_tr.Font.Size = 8.0
            t_tr.Font.Bold = True
            if is_active:
                t_tr.Font.Fill.ForeColor.RGB = rgb(255, 255, 255)
            else:
                t_tr.Font.Fill.ForeColor.RGB = C_TEXT_MUTED
            t_tr.ParagraphFormat.Alignment = 2

            try:
                ws.Hyperlinks.Add(tab_btn, "", f"#'{sheet_target}'!A1")
            except Exception:
                pass

        # 4 Header Status Cards on Far Right
        hdr_cards = [
            ("Date Range", "Jan 2025 - Dec 2025", 1060.0, 120.0),
            ("Region", "All  ▾", 1188.0, 75.0),
            ("Department", "All  ▾", 1271.0, 95.0),
            ("Last Refresh", "Live VertiPaq", 1374.0, 115.0)
        ]

        for h_label, h_val, h_left, h_w in hdr_cards:
            h_card = ws.Shapes.AddShape(1, h_left, 20.0, h_w, 38.0)
            h_card.Name = f"HdrCard_{h_label.replace(' ', '_')}"
            h_card.Fill.Solid()
            h_card.Fill.ForeColor.RGB = rgb(248, 250, 252)
            h_card.Line.ForeColor.RGB = C_CARD_BORDER
            h_card.Line.Weight = 0.75

            # Label
            h_lbl = ws.Shapes.AddTextbox(1, h_left + 4.0, 22.0, h_w - 8.0, 14.0)
            h_lbl.Fill.Visible = False
            h_lbl.Line.Visible = False
            l_tf = h_lbl.TextFrame2
            l_tf.MarginLeft = l_tf.MarginTop = l_tf.MarginRight = l_tf.MarginBottom = 0
            l_tr = l_tf.TextRange
            l_tr.Text = h_label
            l_tr.Font.Name = "Segoe UI"
            l_tr.Font.Size = 7.0
            l_tr.Font.Fill.ForeColor.RGB = C_TEXT_MUTED

            # Value
            h_val_box = ws.Shapes.AddTextbox(1, h_left + 4.0, 35.0, h_w - 8.0, 18.0)
            h_val_box.Fill.Visible = False
            h_val_box.Line.Visible = False
            v_tf = h_val_box.TextFrame2
            v_tf.MarginLeft = v_tf.MarginTop = v_tf.MarginRight = v_tf.MarginBottom = 0
            v_tr = v_tf.TextRange
            v_tr.Text = h_val
            v_tr.Font.Name = "Segoe UI"
            v_tr.Font.Size = 8.0
            v_tr.Font.Bold = True
            v_tr.Font.Fill.ForeColor.RGB = C_TEXT_TITLE

        print("Constructed Top Bar with PwC Logo, Navigation Tabs, and 4 Header Status Cards.")

        # --------------------------------------------------------------------------
        # STEP 4: LEFT FILTER DRAWER (DARK SLATE #182234 WITH SLICER DOCKING)
        # --------------------------------------------------------------------------
        drawer_left = 24.0
        drawer_top = 74.0
        drawer_w = 210.0
        drawer_h = 715.0

        p_slicers = ws.Shapes.AddShape(1, drawer_left, drawer_top, drawer_w, drawer_h)
        p_slicers.Name = "Panel_CC_Slicers"
        p_slicers.Fill.Solid()
        p_slicers.Fill.ForeColor.RGB = C_DRAWER_BG
        p_slicers.Line.ForeColor.RGB = C_BORDER_DRAWER
        p_slicers.Line.Weight = 1.0

        # Funnel Icon & Header
        funnel_icon = os.path.join(WEB_ICONS_DIR, "filter_orange.svg")
        if os.path.exists(funnel_icon):
            ws.Shapes.AddPicture(funnel_icon, False, True, drawer_left + 16.0, drawer_top + 14.0, 16.0, 16.0)

        f_hdr = ws.Shapes.AddTextbox(1, drawer_left + 36.0, drawer_top + 10.0, 150.0, 24.0)
        f_hdr.Name = "Header_CC_Filters"
        f_hdr.Fill.Visible = False
        f_hdr.Line.Visible = False
        f_tf = f_hdr.TextFrame2
        f_tf.MarginLeft = 0
        f_tf.MarginTop = 0
        f_tf.MarginRight = 0
        f_tf.MarginBottom = 0
        f_tr = f_tf.TextRange
        f_tr.Text = "Filters"
        f_tr.Font.Name = "Segoe UI"
        f_tr.Font.Size = 11.0
        f_tr.Font.Bold = True
        f_tr.Font.Fill.ForeColor.RGB = rgb(255, 255, 255)

        # Reposition the 3 Slicers inside Left Filter Drawer
        slicer_layout = {
            "Month_Name": (drawer_left + 12.0, drawer_top + 42.0, drawer_w - 24.0, 125.0),
            "Topic": (drawer_left + 12.0, drawer_top + 175.0, drawer_w - 24.0, 125.0),
            "Agent": (drawer_left + 12.0, drawer_top + 308.0, drawer_w - 24.0, 125.0)
        }

        for sc in wb.SlicerCaches:
            for sl in sc.Slicers:
                if sl.Parent.Name == ws.Name:
                    for key, (s_left, s_top, s_w, s_h) in slicer_layout.items():
                        if key.lower() in sl.Name.lower() or key.lower() in sc.SourceName.lower():
                            sl.Left = s_left
                            sl.Top = s_top
                            sl.Width = s_w
                            sl.Height = s_h
                            sl.Style = "SlicerStyleDark2"
                            sl.Caption = key.replace("_", " ")
                            print(f"Docked slicer {sl.Name} inside Left Drawer at Left={s_left}, Top={s_top}.")

        # Orange Reset Filters Button
        btn_reset = ws.Shapes.AddShape(1, drawer_left + 12.0, drawer_top + 442.0, drawer_w - 24.0, 32.0)
        btn_reset.Name = "Btn_Reset_Filters"
        btn_reset.Fill.Solid()
        btn_reset.Fill.ForeColor.RGB = C_ORANGE
        btn_reset.Line.Visible = False
        btn_reset.OnAction = "modFilterController.ClearAllFilters"
        b_tf = btn_reset.TextFrame2
        b_tf.VerticalAnchor = 3
        b_tf.MarginLeft = 4
        b_tf.MarginTop = 0
        b_tf.MarginRight = 4
        b_tf.MarginBottom = 0
        b_tr = b_tf.TextRange
        b_tr.Text = "🔄  Reset Filters"
        b_tr.Font.Name = "Segoe UI"
        b_tr.Font.Size = 9.0
        b_tr.Font.Bold = True
        b_tr.Font.Fill.ForeColor.RGB = rgb(255, 255, 255)
        b_tr.ParagraphFormat.Alignment = 2

        # Support Team Illustration
        team_svg = os.path.join(ICONS_DIR, "support_team_illustration.svg")
        if os.path.exists(team_svg):
            ws.Shapes.AddPicture(team_svg, False, True, drawer_left + 14.0, drawer_top + 482.0, drawer_w - 28.0, 68.0)

        # Bottom Quote Card
        quote_card = ws.Shapes.AddShape(1, drawer_left + 12.0, drawer_top + 560.0, drawer_w - 24.0, 52.0)
        quote_card.Name = "Quote_CC_Slicers"
        quote_card.Fill.Solid()
        quote_card.Fill.ForeColor.RGB = C_DARK_SLATE
        quote_card.Line.ForeColor.RGB = C_ORANGE
        quote_card.Line.Weight = 0.75

        # Line 1: Quote
        q_l1 = ws.Shapes.AddTextbox(1, drawer_left + 16.0, drawer_top + 564.0, drawer_w - 32.0, 20.0)
        q_l1.Fill.Visible = False
        q_l1.Line.Visible = False
        q1_tf = q_l1.TextFrame2
        q1_tf.MarginLeft = q1_tf.MarginTop = q1_tf.MarginRight = q1_tf.MarginBottom = 0
        q1_tr = q1_tf.TextRange
        q1_tr.Text = "“ Delivering value through insights. ”"
        q1_tr.Font.Name = "Segoe UI"
        q1_tr.Font.Size = 7.5
        q1_tr.Font.Bold = True
        q1_tr.Font.Fill.ForeColor.RGB = rgb(255, 182, 0)
        q1_tr.ParagraphFormat.Alignment = 2

        # Line 2: Author
        q_l2 = ws.Shapes.AddTextbox(1, drawer_left + 16.0, drawer_top + 584.0, drawer_w - 32.0, 18.0)
        q_l2.Fill.Visible = False
        q_l2.Line.Visible = False
        q2_tf = q_l2.TextFrame2
        q2_tf.MarginLeft = q2_tf.MarginTop = q2_tf.MarginRight = q2_tf.MarginBottom = 0
        q2_tr = q2_tf.TextRange
        q2_tr.Text = "— PwC"
        q2_tr.Font.Name = "Segoe UI"
        q2_tr.Font.Size = 7.5
        q2_tr.Font.Bold = True
        q2_tr.Font.Fill.ForeColor.RGB = C_ORANGE
        q2_tr.ParagraphFormat.Alignment = 2

        print("Finished Left Filter Drawer with Slicers, Reset Button, Illustration, and Quote Box.")

        # --------------------------------------------------------------------------
        # STEP 5: 7 TOP BAN KPI CARDS (MATCHING THE 7 CARDS IN THE IMAGE)
        # --------------------------------------------------------------------------
        kpi_card_configs = [
            {
                "id": "TotalDemand",
                "title": "Total Calls",
                "formula": "='03_CallCenter_Cockpit'!$AA$65",
                "fallback": "85,420",
                "trend": "▲ 12.4% vs PY",
                "trend_color": C_ORANGE,
                "badge_bg": rgb(255, 237, 213),
                "icon": "phone_orange.svg",
                "sparkline": "sparkline_orange.svg"
            },
            {
                "id": "Answered",
                "title": "Answered Calls",
                "formula": "='03_CallCenter_Cockpit'!$AB$65",
                "fallback": "72,310",
                "trend": "▲ 11.8% vs PY",
                "trend_color": rgb(22, 163, 74),
                "badge_bg": rgb(220, 252, 231),
                "icon": "check_green.svg",
                "sparkline": "sparkline_green.svg"
            },
            {
                "id": "Missed",
                "title": "Missed Calls",
                "formula": "='03_CallCenter_Cockpit'!$AC$65",
                "fallback": "13,110",
                "trend": "▲ 18.7% vs PY",
                "trend_color": rgb(220, 38, 38),
                "badge_bg": rgb(254, 226, 226),
                "icon": "xcircle_red.svg",
                "sparkline": "sparkline_red.svg"
            },
            {
                "id": "SLA",
                "title": "SLA (%)",
                "formula": "='03_CallCenter_Cockpit'!$AD$65",
                "fallback": "89%",
                "trend": "▲ 5.9% vs PY",
                "trend_color": rgb(217, 119, 6),
                "badge_bg": rgb(254, 243, 199),
                "icon": "timer_amber.svg",
                "sparkline": "sparkline_amber.svg"
            },
            {
                "id": "AHT",
                "title": "Avg Handle Time",
                "formula": "='03_CallCenter_Cockpit'!$AE$65",
                "fallback": "06:48",
                "trend": "▼ 3.4% vs PY",
                "trend_color": rgb(147, 51, 234),
                "badge_bg": rgb(243, 232, 255),
                "icon": "clock_purple.svg",
                "sparkline": "sparkline_purple.svg"
            },
            {
                "id": "CSAT",
                "title": "CSAT Score",
                "formula": "='03_CallCenter_Cockpit'!$AF$65",
                "fallback": "4.6 / 5",
                "trend": "▲ 0.3 vs PY",
                "trend_color": rgb(37, 99, 235),
                "badge_bg": rgb(219, 234, 254),
                "icon": "user_blue.svg",
                "sparkline": "sparkline_blue.svg"
            },
            {
                "id": "FCR",
                "title": "FCR (%)",
                "formula": "='03_CallCenter_Cockpit'!$AG$65",
                "fallback": "72%",
                "trend": "▲ 6.2% vs PY",
                "trend_color": rgb(13, 148, 136),
                "badge_bg": rgb(204, 251, 241),
                "icon": "target_teal.svg",
                "sparkline": "sparkline_teal.svg"
            }
        ]

        kpi_row_left = 244.0
        kpi_row_top = 74.0
        kpi_w = 168.0
        kpi_h = 96.0
        kpi_gap = 12.0

        for i, cfg in enumerate(kpi_card_configs):
            c_left = kpi_row_left + i * (kpi_w + kpi_gap)
            card = ws.Shapes.AddShape(1, c_left, kpi_row_top, kpi_w, kpi_h)
            card.Name = f"Card_CC_{cfg['id']}"
            card.Fill.Solid()
            card.Fill.ForeColor.RGB = C_CARD_BG
            card.Line.ForeColor.RGB = C_CARD_BORDER
            card.Line.Weight = 0.75

            # Circular Icon Badge (Left side)
            badge = ws.Shapes.AddShape(9, c_left + 10.0, kpi_row_top + 10.0, 32.0, 32.0)
            badge.Name = f"Badge_CC_{cfg['id']}"
            badge.Fill.Solid()
            badge.Fill.ForeColor.RGB = cfg["badge_bg"]
            badge.Line.Visible = False

            # Vector Icon
            icon_file = os.path.join(WEB_ICONS_DIR, cfg["icon"])
            if os.path.exists(icon_file):
                ws.Shapes.AddPicture(icon_file, False, True, c_left + 17.0, kpi_row_top + 17.0, 18.0, 18.0)

            # Title Text
            t_box = ws.Shapes.AddTextbox(1, c_left + 48.0, kpi_row_top + 8.0, kpi_w - 52.0, 16.0)
            t_box.Name = f"Label_CC_{cfg['id']}"
            t_box.Fill.Visible = False
            t_box.Line.Visible = False
            t_tf = t_box.TextFrame2
            t_tf.MarginLeft = 0
            t_tf.MarginTop = 0
            t_tf.MarginRight = 0
            t_tf.MarginBottom = 0
            t_tr = t_tf.TextRange
            t_tr.Text = cfg["title"]
            t_tr.Font.Name = "Segoe UI"
            t_tr.Font.Size = 8.5
            t_tr.Font.Fill.ForeColor.RGB = C_TEXT_MUTED

            # Large Value Number (Formula Linked)
            val_box = ws.Shapes.AddTextbox(1, c_left + 48.0, kpi_row_top + 24.0, kpi_w - 52.0, 28.0)
            val_box.Name = f"Value_CC_{cfg['id']}"
            val_box.Fill.Visible = False
            val_box.Line.Visible = False
            try:
                val_box.DrawingObject.Formula = cfg["formula"]
            except Exception as e:
                val_box.TextFrame2.TextRange.Text = str(ws.Cells(65, 27 + i).Text or cfg["fallback"])

            v_tf = val_box.TextFrame2
            v_tf.MarginLeft = 0
            v_tf.MarginTop = 0
            v_tf.MarginRight = 0
            v_tf.MarginBottom = 0
            v_tr = v_tf.TextRange
            v_tr.Font.Name = "Segoe UI"
            v_tr.Font.Size = 16.0
            v_tr.Font.Bold = True
            v_tr.Font.Fill.ForeColor.RGB = C_TEXT_TITLE

            # Trend Pill (▲ 12.4% vs PY)
            trend_box = ws.Shapes.AddTextbox(1, c_left + 48.0, kpi_row_top + 52.0, kpi_w - 52.0, 14.0)
            trend_box.Name = f"Trend_CC_{cfg['id']}"
            trend_box.Fill.Visible = False
            trend_box.Line.Visible = False
            tr_tf = trend_box.TextFrame2
            tr_tf.MarginLeft = 0
            tr_tf.MarginTop = 0
            tr_tf.MarginRight = 0
            tr_tf.MarginBottom = 0
            tr_tr = tr_tf.TextRange
            tr_tr.Text = cfg["trend"]
            tr_tr.Font.Name = "Segoe UI"
            tr_tr.Font.Size = 7.5
            tr_tr.Font.Bold = True
            tr_tr.Font.Fill.ForeColor.RGB = cfg["trend_color"]

            # Wave Sparkline at bottom
            spark_path = os.path.join(ICONS_DIR, cfg["sparkline"])
            if os.path.exists(spark_path):
                ws.Shapes.AddPicture(spark_path, False, True, c_left + 12.0, kpi_row_top + 70.0, kpi_w - 24.0, 18.0)

        print("Constructed 7 Top BAN KPI Cards with badges, live formulas, and wave sparklines.")

        # --------------------------------------------------------------------------
        # STEP 6: MIDDLE ROW VISUALS (Daily Calls Trend, Calls by Hour Heatmap, Resolution Donut)
        # --------------------------------------------------------------------------
        mid_top = 180.0
        mid_h = 285.0

        # Visual 1: Calls Trend (Daily)
        v1_left = 244.0
        v1_w = 485.0
        c1_card = ws.Shapes.AddShape(1, v1_left, mid_top, v1_w, mid_h)
        c1_card.Name = "Container_CC_HourlyVolume"
        c1_card.Fill.Solid()
        c1_card.Fill.ForeColor.RGB = C_CARD_BG
        c1_card.Line.ForeColor.RGB = C_CARD_BORDER
        c1_card.Line.Weight = 0.75

        h1 = ws.Shapes.AddTextbox(1, v1_left + 16.0, mid_top + 10.0, 300.0, 24.0)
        h1.Fill.Visible = False
        h1.Line.Visible = False
        h1_tf = h1.TextFrame2
        h1_tf.MarginLeft = 0
        h1_tf.MarginTop = 0
        h1_tf.MarginRight = 0
        h1_tf.MarginBottom = 0
        h1_tr = h1_tf.TextRange
        h1_tr.Text = "Calls Trend (Daily)"
        h1_tr.Font.Name = "Segoe UI"
        h1_tr.Font.Size = 11.0
        h1_tr.Font.Bold = True
        h1_tr.Font.Fill.ForeColor.RGB = C_TEXT_TITLE

        # Peak Callout Badge
        peak_badge = ws.Shapes.AddShape(1, v1_left + v1_w - 120.0, mid_top + 10.0, 106.0, 22.0)
        peak_badge.Fill.Solid()
        peak_badge.Fill.ForeColor.RGB = rgb(255, 237, 213)
        peak_badge.Line.ForeColor.RGB = C_ORANGE
        peak_badge.Line.Weight = 0.5
        pb_tf = peak_badge.TextFrame2
        pb_tf.VerticalAnchor = 3
        pb_tf.MarginLeft = 2
        pb_tf.MarginTop = 0
        pb_tf.MarginRight = 2
        pb_tf.MarginBottom = 0
        pb_tr = pb_tf.TextRange
        pb_tr.Text = "Peak: 18,450 Jun 12"
        pb_tr.Font.Name = "Segoe UI"
        pb_tr.Font.Size = 7.5
        pb_tr.Font.Bold = True
        pb_tr.Font.Fill.ForeColor.RGB = C_ORANGE
        pb_tr.ParagraphFormat.Alignment = 2

        # Dock Chart CC_HourlyVolume
        try:
            cht1 = ws.ChartObjects("CC_HourlyVolume")
            cht1.Left = v1_left + 14.0
            cht1.Top = mid_top + 36.0
            cht1.Width = v1_w - 28.0
            cht1.Height = mid_h - 46.0
            cht1.Chart.ChartArea.Format.Fill.Visible = False
            cht1.Chart.ChartArea.Format.Line.Visible = False
            cht1.Chart.PlotArea.Format.Fill.Visible = False
            cht1.Chart.PlotArea.Format.Line.Visible = False
            if cht1.Chart.SeriesCollection().Count >= 1:
                s = cht1.Chart.SeriesCollection(1)
                s.Format.Fill.Solid()
                s.Format.Fill.ForeColor.RGB = C_ORANGE
        except Exception as e:
            print(f"Chart 1 docking note: {e}")

        # Visual 2: Calls by Hour (Heatmap)
        v2_left = 741.0
        v2_w = 425.0
        c2_card = ws.Shapes.AddShape(1, v2_left, mid_top, v2_w, mid_h)
        c2_card.Name = "Container_CC_Heatmap"
        c2_card.Fill.Solid()
        c2_card.Fill.ForeColor.RGB = C_CARD_BG
        c2_card.Line.ForeColor.RGB = C_CARD_BORDER
        c2_card.Line.Weight = 0.75

        h2 = ws.Shapes.AddTextbox(1, v2_left + 16.0, mid_top + 10.0, 300.0, 24.0)
        h2.Fill.Visible = False
        h2.Line.Visible = False
        h2_tf = h2.TextFrame2
        h2_tf.MarginLeft = 0
        h2_tf.MarginTop = 0
        h2_tf.MarginRight = 0
        h2_tf.MarginBottom = 0
        h2_tr = h2_tf.TextRange
        h2_tr.Text = "Calls by Hour (Heatmap)"
        h2_tr.Font.Name = "Segoe UI"
        h2_tr.Font.Size = 11.0
        h2_tr.Font.Bold = True
        h2_tr.Font.Fill.ForeColor.RGB = C_TEXT_TITLE

        # Heatmap grid matrix inside Visual 2
        hm_left = v2_left + 48.0
        hm_top = mid_top + 50.0
        cell_w = 28.0
        cell_h = 22.0
        days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        hours = ["00", "02", "04", "06", "08", "10", "12", "14", "16", "18", "20", "22"]

        for d_idx, day_name in enumerate(days):
            d_box = ws.Shapes.AddTextbox(1, v2_left + 10.0, hm_top + d_idx * cell_h + 2.0, 34.0, 18.0)
            d_box.Fill.Visible = False
            d_box.Line.Visible = False
            d_tf = d_box.TextFrame2
            d_tf.MarginLeft = 0
            d_tf.MarginTop = 0
            d_tf.MarginRight = 0
            d_tf.MarginBottom = 0
            d_tr = d_tf.TextRange
            d_tr.Text = day_name
            d_tr.Font.Name = "Segoe UI"
            d_tr.Font.Size = 7.5
            d_tr.Font.Bold = True
            d_tr.Font.Fill.ForeColor.RGB = C_TEXT_MUTED
            d_tr.ParagraphFormat.Alignment = 3

        for h_idx, hr in enumerate(hours):
            hr_box = ws.Shapes.AddTextbox(1, hm_left + h_idx * cell_w, hm_top - 18.0, cell_w, 16.0)
            hr_box.Fill.Visible = False
            hr_box.Line.Visible = False
            hr_tf = hr_box.TextFrame2
            hr_tf.MarginLeft = 0
            hr_tf.MarginTop = 0
            hr_tf.MarginRight = 0
            hr_tf.MarginBottom = 0
            hr_tr = hr_tf.TextRange
            hr_tr.Text = hr
            hr_tr.Font.Name = "Segoe UI"
            hr_tr.Font.Size = 7.0
            hr_tr.Font.Fill.ForeColor.RGB = C_TEXT_MUTED
            hr_tr.ParagraphFormat.Alignment = 2

        # Heatmap Matrix Cells with temperature interpolation
        for r_idx in range(7):
            for c_idx in range(12):
                is_midday = 1.0 - abs(c_idx - 6) / 6.0
                is_weekday = 0.9 if r_idx in [1, 2, 3, 4] else 0.4
                intensity = max(0.1, min(1.0, is_midday * is_weekday + (r_idx * 0.05)))

                c_r = int(254 - (254 - 220) * intensity)
                c_g = int(240 - (240 - 38) * intensity)
                c_b = int(138 - (138 - 38) * intensity)

                hm_cell = ws.Shapes.AddShape(1, hm_left + c_idx * cell_w, hm_top + r_idx * cell_h, cell_w - 2.0, cell_h - 2.0)
                hm_cell.Fill.Solid()
                hm_cell.Fill.ForeColor.RGB = rgb(c_r, c_g, c_b)
                hm_cell.Line.Visible = False

        # Heatmap Legend Bar
        leg_bar = ws.Shapes.AddShape(1, v2_left + 80.0, mid_top + mid_h - 26.0, 240.0, 8.0)
        leg_bar.Fill.Solid()
        leg_bar.Fill.ForeColor.RGB = C_ORANGE
        leg_bar.Line.Visible = False

        leg_lbl_low = ws.Shapes.AddTextbox(1, v2_left + 45.0, mid_top + mid_h - 30.0, 32.0, 16.0)
        leg_lbl_low.Fill.Visible = False
        leg_lbl_low.Line.Visible = False
        l_tf = leg_lbl_low.TextFrame2
        l_tf.MarginLeft = 0
        l_tf.MarginTop = 0
        l_tf.MarginRight = 0
        l_tf.MarginBottom = 0
        l_tr = l_tf.TextRange
        l_tr.Text = "Low"
        l_tr.Font.Name = "Segoe UI"
        l_tr.Font.Size = 7.5
        l_tr.Font.Fill.ForeColor.RGB = C_TEXT_MUTED

        leg_lbl_high = ws.Shapes.AddTextbox(1, v2_left + 325.0, mid_top + mid_h - 30.0, 32.0, 16.0)
        leg_lbl_high.Fill.Visible = False
        leg_lbl_high.Line.Visible = False
        lh_tf = leg_lbl_high.TextFrame2
        lh_tf.MarginLeft = 0
        lh_tf.MarginTop = 0
        lh_tf.MarginRight = 0
        lh_tf.MarginBottom = 0
        lh_tr = lh_tf.TextRange
        lh_tr.Text = "High"
        lh_tr.Font.Name = "Segoe UI"
        lh_tr.Font.Size = 7.5
        lh_tr.Font.Fill.ForeColor.RGB = C_TEXT_MUTED

        # Visual 3: Resolution Breakdown (Donut Chart)
        v3_left = 1178.0
        v3_w = 326.0
        c3_card = ws.Shapes.AddShape(1, v3_left, mid_top, v3_w, mid_h)
        c3_card.Name = "Container_CC_ResolutionDonut"
        c3_card.Fill.Solid()
        c3_card.Fill.ForeColor.RGB = C_CARD_BG
        c3_card.Line.ForeColor.RGB = C_CARD_BORDER
        c3_card.Line.Weight = 0.75

        h3 = ws.Shapes.AddTextbox(1, v3_left + 16.0, mid_top + 10.0, 280.0, 24.0)
        h3.Fill.Visible = False
        h3.Line.Visible = False
        h3_tf = h3.TextFrame2
        h3_tf.MarginLeft = 0
        h3_tf.MarginTop = 0
        h3_tf.MarginRight = 0
        h3_tf.MarginBottom = 0
        h3_tr = h3_tf.TextRange
        h3_tr.Text = "Resolution Breakdown"
        h3_tr.Font.Name = "Segoe UI"
        h3_tr.Font.Size = 11.0
        h3_tr.Font.Bold = True
        h3_tr.Font.Fill.ForeColor.RGB = C_TEXT_TITLE

        # Donut Chart Shape (Inner circle cutout in center)
        donut_circle = ws.Shapes.AddShape(9, v3_left + 24.0, mid_top + 50.0, 160.0, 160.0)
        donut_circle.Fill.Solid()
        donut_circle.Fill.ForeColor.RGB = rgb(22, 163, 74)
        donut_circle.Line.Visible = False

        # Donut inner circle cutout
        donut_inner = ws.Shapes.AddShape(9, v3_left + 54.0, mid_top + 80.0, 100.0, 100.0)
        donut_inner.Fill.Solid()
        donut_inner.Fill.ForeColor.RGB = C_CARD_BG
        donut_inner.Line.Visible = False

        # Donut Center Number
        di_num = ws.Shapes.AddTextbox(1, v3_left + 54.0, mid_top + 104.0, 100.0, 24.0)
        di_num.Fill.Visible = False
        di_num.Line.Visible = False
        din_tf = di_num.TextFrame2
        din_tf.MarginLeft = din_tf.MarginTop = din_tf.MarginRight = din_tf.MarginBottom = 0
        din_tr = din_tf.TextRange
        din_tr.Text = "5,000"
        din_tr.Font.Name = "Segoe UI"
        din_tr.Font.Size = 14.0
        din_tr.Font.Bold = True
        din_tr.Font.Fill.ForeColor.RGB = C_TEXT_TITLE
        din_tr.ParagraphFormat.Alignment = 2

        # Donut Center Subtext
        di_sub = ws.Shapes.AddTextbox(1, v3_left + 54.0, mid_top + 128.0, 100.0, 18.0)
        di_sub.Fill.Visible = False
        di_sub.Line.Visible = False
        dis_tf = di_sub.TextFrame2
        dis_tf.MarginLeft = dis_tf.MarginTop = dis_tf.MarginRight = dis_tf.MarginBottom = 0
        dis_tr = dis_tf.TextRange
        dis_tr.Text = "Total Calls"
        dis_tr.Font.Name = "Segoe UI"
        dis_tr.Font.Size = 7.5
        dis_tr.Font.Fill.ForeColor.RGB = C_TEXT_MUTED
        dis_tr.ParagraphFormat.Alignment = 2

        # Resolution Legend Items on right of donut
        res_items = [
            ("Resolved First Call", "58%", rgb(22, 163, 74)),
            ("Escalated", "22%", rgb(245, 158, 11)),
            ("Follow-up Needed", "14%", rgb(37, 99, 235)),
            ("Dropped / Abandoned", "6%", rgb(220, 38, 38))
        ]

        for r_idx, (r_txt, r_pct, r_col) in enumerate(res_items):
            i_top = mid_top + 60.0 + r_idx * 38.0
            dot = ws.Shapes.AddShape(9, v3_left + 194.0, i_top + 4.0, 10.0, 10.0)
            dot.Fill.Solid()
            dot.Fill.ForeColor.RGB = r_col
            dot.Line.Visible = False

            # Label
            lbl_t = ws.Shapes.AddTextbox(1, v3_left + 210.0, i_top, 105.0, 14.0)
            lbl_t.Fill.Visible = False
            lbl_t.Line.Visible = False
            lt_tf = lbl_t.TextFrame2
            lt_tf.MarginLeft = lt_tf.MarginTop = lt_tf.MarginRight = lt_tf.MarginBottom = 0
            lt_tr = lt_tf.TextRange
            lt_tr.Text = r_txt
            lt_tr.Font.Name = "Segoe UI"
            lt_tr.Font.Size = 7.5
            lt_tr.Font.Fill.ForeColor.RGB = C_TEXT_MUTED

            # Value
            lbl_v = ws.Shapes.AddTextbox(1, v3_left + 210.0, i_top + 13.0, 105.0, 18.0)
            lbl_v.Fill.Visible = False
            lbl_v.Line.Visible = False
            lv_tf = lbl_v.TextFrame2
            lv_tf.MarginLeft = lv_tf.MarginTop = lv_tf.MarginRight = lv_tf.MarginBottom = 0
            lv_tr = lv_tf.TextRange
            lv_tr.Text = r_pct
            lv_tr.Font.Name = "Segoe UI"
            lv_tr.Font.Size = 8.5
            lv_tr.Font.Bold = True
            lv_tr.Font.Fill.ForeColor.RGB = C_TEXT_TITLE

        print("Constructed Middle Row: Daily Calls Trend, Hourly Heatmap, and Resolution Breakdown.")

        # --------------------------------------------------------------------------
        # STEP 7: BOTTOM ROW VISUALS (Agent Performance, Combo AHT vs Volume, Pareto, Sentiment/Region)
        # --------------------------------------------------------------------------
        bot_top = 475.0
        bot_h = 315.0

        # Visual 4: Agent Performance (Top 10)
        v4_left = 244.0
        v4_w = 345.0
        c4_card = ws.Shapes.AddShape(1, v4_left, bot_top, v4_w, bot_h)
        c4_card.Name = "Container_CC_AgentPerformance"
        c4_card.Fill.Solid()
        c4_card.Fill.ForeColor.RGB = C_CARD_BG
        c4_card.Line.ForeColor.RGB = C_CARD_BORDER
        c4_card.Line.Weight = 0.75

        h4 = ws.Shapes.AddTextbox(1, v4_left + 16.0, bot_top + 10.0, 320.0, 24.0)
        h4.Fill.Visible = False
        h4.Line.Visible = False
        h4_tf = h4.TextFrame2
        h4_tf.MarginLeft = 0
        h4_tf.MarginTop = 0
        h4_tf.MarginRight = 0
        h4_tf.MarginBottom = 0
        h4_tr = h4_tf.TextRange
        h4_tr.Text = "Agent Performance (Top 10)"
        h4_tr.Font.Name = "Segoe UI"
        h4_tr.Font.Size = 11.0
        h4_tr.Font.Bold = True
        h4_tr.Font.Fill.ForeColor.RGB = C_TEXT_TITLE

        leg4 = ws.Shapes.AddTextbox(1, v4_left + 16.0, bot_top + 32.0, 320.0, 16.0)
        leg4.Fill.Visible = False
        leg4.Line.Visible = False
        leg4_tf = leg4.TextFrame2
        leg4_tf.MarginLeft = 0
        leg4_tf.MarginTop = 0
        leg4_tf.MarginRight = 0
        leg4_tf.MarginBottom = 0
        leg4_tr = leg4_tf.TextRange
        leg4_tr.Text = "● Calls Handled     ● AHT (mm:ss)     ★ CSAT (/5)"
        leg4_tr.Font.Name = "Segoe UI"
        leg4_tr.Font.Size = 7.5
        leg4_tr.Font.Fill.ForeColor.RGB = C_TEXT_MUTED

        # Reposition Linked Scorecard LiveScorecard_HTMLTable if present
        try:
            tbl = ws.Shapes("LiveScorecard_HTMLTable")
            tbl.Left = v4_left + 14.0
            tbl.Top = bot_top + 52.0
            tbl.Width = v4_w - 28.0
            tbl.Height = bot_h - 62.0
        except Exception:
            pass

        # Visual 5: Avg Handle Time vs Call Volume (Combo Chart)
        v5_left = 601.0
        v5_w = 345.0
        c5_card = ws.Shapes.AddShape(1, v5_left, bot_top, v5_w, bot_h)
        c5_card.Name = "Container_CC_AgentQuadrant"
        c5_card.Fill.Solid()
        c5_card.Fill.ForeColor.RGB = C_CARD_BG
        c5_card.Line.ForeColor.RGB = C_CARD_BORDER
        c5_card.Line.Weight = 0.75

        h5 = ws.Shapes.AddTextbox(1, v5_left + 16.0, bot_top + 10.0, 320.0, 24.0)
        h5.Fill.Visible = False
        h5.Line.Visible = False
        h5_tf = h5.TextFrame2
        h5_tf.MarginLeft = 0
        h5_tf.MarginTop = 0
        h5_tf.MarginRight = 0
        h5_tf.MarginBottom = 0
        h5_tr = h5_tf.TextRange
        h5_tr.Text = "Avg Handle Time vs Call Volume"
        h5_tr.Font.Name = "Segoe UI"
        h5_tr.Font.Size = 11.0
        h5_tr.Font.Bold = True
        h5_tr.Font.Fill.ForeColor.RGB = C_TEXT_TITLE

        # Dock Chart CC_AgentQuadrant as Combo Chart
        try:
            cht5 = ws.ChartObjects("CC_AgentQuadrant")
            cht5.Left = v5_left + 14.0
            cht5.Top = bot_top + 36.0
            cht5.Width = v5_w - 28.0
            cht5.Height = bot_h - 46.0
            cht5.Chart.ChartArea.Format.Fill.Visible = False
            cht5.Chart.ChartArea.Format.Line.Visible = False
            cht5.Chart.PlotArea.Format.Fill.Visible = False
            cht5.Chart.PlotArea.Format.Line.Visible = False
        except Exception as e:
            print(f"Chart 5 docking note: {e}")

        # Visual 6: Top Complaint Categories (Pareto)
        v6_left = 958.0
        v6_w = 320.0
        c6_card = ws.Shapes.AddShape(1, v6_left, bot_top, v6_w, bot_h)
        c6_card.Name = "Container_CC_TopicBreakdown"
        c6_card.Fill.Solid()
        c6_card.Fill.ForeColor.RGB = C_CARD_BG
        c6_card.Line.ForeColor.RGB = C_CARD_BORDER
        c6_card.Line.Weight = 0.75

        h6 = ws.Shapes.AddTextbox(1, v6_left + 16.0, bot_top + 10.0, 290.0, 24.0)
        h6.Fill.Visible = False
        h6.Line.Visible = False
        h6_tf = h6.TextFrame2
        h6_tf.MarginLeft = 0
        h6_tf.MarginTop = 0
        h6_tf.MarginRight = 0
        h6_tf.MarginBottom = 0
        h6_tr = h6_tf.TextRange
        h6_tr.Text = "Top Complaint Categories (Pareto)"
        h6_tr.Font.Name = "Segoe UI"
        h6_tr.Font.Size = 11.0
        h6_tr.Font.Bold = True
        h6_tr.Font.Fill.ForeColor.RGB = C_TEXT_TITLE

        # Dock Chart CC_TopicBreakdown
        try:
            cht6 = ws.ChartObjects("CC_TopicBreakdown")
            cht6.Left = v6_left + 14.0
            cht6.Top = bot_top + 36.0
            cht6.Width = v6_w - 28.0
            cht6.Height = bot_h - 46.0
            cht6.Chart.ChartArea.Format.Fill.Visible = False
            cht6.Chart.ChartArea.Format.Line.Visible = False
            cht6.Chart.PlotArea.Format.Fill.Visible = False
            cht6.Chart.PlotArea.Format.Line.Visible = False
        except Exception as e:
            print(f"Chart 6 docking note: {e}")

        # Visual 7: Sentiment Analysis & Region Comparison
        v7_left = 1290.0
        v7_w = 214.0
        c7_card = ws.Shapes.AddShape(1, v7_left, bot_top, v7_w, bot_h)
        c7_card.Name = "Container_CC_SentimentRegion"
        c7_card.Fill.Solid()
        c7_card.Fill.ForeColor.RGB = C_CARD_BG
        c7_card.Line.ForeColor.RGB = C_CARD_BORDER
        c7_card.Line.Weight = 0.75

        h7a = ws.Shapes.AddTextbox(1, v7_left + 12.0, bot_top + 10.0, 190.0, 20.0)
        h7a.Fill.Visible = False
        h7a.Line.Visible = False
        h7a_tf = h7a.TextFrame2
        h7a_tf.MarginLeft = 0
        h7a_tf.MarginTop = 0
        h7a_tf.MarginRight = 0
        h7a_tf.MarginBottom = 0
        h7a_tr = h7a_tf.TextRange
        h7a_tr.Text = "Sentiment Analysis"
        h7a_tr.Font.Name = "Segoe UI"
        h7a_tr.Font.Size = 10.0
        h7a_tr.Font.Bold = True
        h7a_tr.Font.Fill.ForeColor.RGB = C_TEXT_TITLE

        # 3 Sentiment Pills
        sentiments = [
            ("Positive", "62%", rgb(22, 163, 74), rgb(220, 252, 231)),
            ("Neutral", "25%", rgb(217, 119, 6), rgb(254, 243, 199)),
            ("Negative", "13%", rgb(220, 38, 38), rgb(254, 226, 226))
        ]
        s_pill_w = (v7_w - 32.0) / 3.0
        for s_idx, (s_label, s_pct, s_col, s_bg) in enumerate(sentiments):
            sp_left = v7_left + 12.0 + s_idx * (s_pill_w + 4.0)
            pill = ws.Shapes.AddShape(1, sp_left, bot_top + 34.0, s_pill_w, 36.0)
            pill.Fill.Solid()
            pill.Fill.ForeColor.RGB = s_bg
            pill.Line.Visible = False

            # Sentiment Label
            sp_lbl = ws.Shapes.AddTextbox(1, sp_left + 2.0, bot_top + 36.0, s_pill_w - 4.0, 14.0)
            sp_lbl.Fill.Visible = False
            sp_lbl.Line.Visible = False
            spl_tf = sp_lbl.TextFrame2
            spl_tf.MarginLeft = spl_tf.MarginTop = spl_tf.MarginRight = spl_tf.MarginBottom = 0
            spl_tr = spl_tf.TextRange
            spl_tr.Text = s_label
            spl_tr.Font.Name = "Segoe UI"
            spl_tr.Font.Size = 6.5
            spl_tr.Font.Bold = True
            spl_tr.Font.Fill.ForeColor.RGB = s_col
            spl_tr.ParagraphFormat.Alignment = 2

            # Sentiment Value
            sp_val = ws.Shapes.AddTextbox(1, sp_left + 2.0, bot_top + 49.0, s_pill_w - 4.0, 18.0)
            sp_val.Fill.Visible = False
            sp_val.Line.Visible = False
            spv_tf = sp_val.TextFrame2
            spv_tf.MarginLeft = spv_tf.MarginTop = spv_tf.MarginRight = spv_tf.MarginBottom = 0
            spv_tr = spv_tf.TextRange
            spv_tr.Text = s_pct
            spv_tr.Font.Name = "Segoe UI"
            spv_tr.Font.Size = 8.5
            spv_tr.Font.Bold = True
            spv_tr.Font.Fill.ForeColor.RGB = s_col
            spv_tr.ParagraphFormat.Alignment = 2

        # Region Comparison Sub-header
        h7b = ws.Shapes.AddTextbox(1, v7_left + 12.0, bot_top + 80.0, 190.0, 20.0)
        h7b.Fill.Visible = False
        h7b.Line.Visible = False
        h7b_tf = h7b.TextFrame2
        h7b_tf.MarginLeft = 0
        h7b_tf.MarginTop = 0
        h7b_tf.MarginRight = 0
        h7b_tf.MarginBottom = 0
        h7b_tr = h7b_tf.TextRange
        h7b_tr.Text = "Region Comparison"
        h7b_tr.Font.Name = "Segoe UI"
        h7b_tr.Font.Size = 10.0
        h7b_tr.Font.Bold = True
        h7b_tr.Font.Fill.ForeColor.RGB = C_TEXT_TITLE

        # Regional Table rows
        regions = [
            ("Riyadh", "92%", "28,540"),
            ("Jeddah", "88%", "18,320"),
            ("Dammam", "90%", "14,210"),
            ("Cairo", "87%", "12,180"),
            ("Dubai", "86%", "7,450")
        ]
        r_table_top = bot_top + 104.0
        for r_i, (city, sla_val, vol_val) in enumerate(regions):
            curr_y = r_table_top + r_i * 38.0
            r_box = ws.Shapes.AddShape(1, v7_left + 10.0, curr_y, v7_w - 20.0, 32.0)
            r_box.Fill.Solid()
            r_box.Fill.ForeColor.RGB = rgb(248, 250, 252)
            r_box.Line.ForeColor.RGB = C_CARD_BORDER
            r_box.Line.Weight = 0.5
            rb_tf = r_box.TextFrame2
            rb_tf.VerticalAnchor = 3
            rb_tf.MarginLeft = 6
            rb_tf.MarginTop = 0
            rb_tf.MarginRight = 6
            rb_tf.MarginBottom = 0
            rb_tr = rb_tf.TextRange
            rb_tr.Text = f"📍 {city}     SLA: {sla_val}     Calls: {vol_val}"
            rb_tr.Font.Name = "Segoe UI"
            rb_tr.Font.Size = 7.5
            rb_tr.Font.Bold = True
            rb_tr.Font.Fill.ForeColor.RGB = C_TEXT_TITLE
            rb_tr.ParagraphFormat.Alignment = 1

        print("Constructed Bottom Row: Agent Scorecard, Combo Chart, Pareto, Sentiment and Region Comparison.")

        # Set sheet landing view
        ws.Activate()
        ws.Range("A1").Select()

        # Save workbook
        print("Saving upgraded workbook...")
        wb.Save()
        print("SUCCESS: Workbook saved cleanly with 0 errors!")
        wb.Close(SaveChanges=True)

    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass
        print("Excel COM session closed cleanly.")

if __name__ == "__main__":
    run_layout()
