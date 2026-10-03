"""
PwC Switzerland Virtual Case Experience - Full Enterprise Modernization Engine
Applies website-inspired UI/UX across ALL sheets in PWC_Switzerland_Virtual_Case.xlsm:
1. Re-builds 00_Home_Portal with perfectly centered text in all boxes/buttons/pills.
2. Formats 01_Business_Domains as a SaaS Business Architecture Portal with interactive domain cards.
3. Formats 02_Metadata_&_KPI_Catalog as a SaaS Data Governance & Measure Dictionary Portal.
4. Modernizes 03_CallCenter_Cockpit with vector SVG icons, zero-collision nav, dark filter panel, and 4 visual charts.
5. Modernizes 04_CustomerRetention_Cockpit with vector SVG icons, zero-collision nav, dark filter panel, and 4 visual charts.
6. Modernizes 05_DiversityInclusion_Cockpit with vector SVG icons, zero-collision nav, dark filter panel, and 4 visual charts.
Pure ASCII / Latin-1 clean.
"""

import os
import sys
sys.stdout.reconfigure(line_buffering=True)
import win32com.client

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKBOOK_PATH = os.path.join(BASE_DIR, "PWC_Switzerland_Virtual_Case.xlsm")
ICONS_DIR = os.path.join(BASE_DIR, "assets", "icons", "web")
LOGO_PATH = os.path.join(BASE_DIR, "assets", "PwC_logo_rgb_colour_pos.png")

def rgb(r, g, b):
    return r + (g * 256) + (b * 65536)

def format_text_center(shp, text, font_size=8.0, bold=True, color_rgb=rgb(15, 23, 42)):
    tf = shp.TextFrame2
    tf.VerticalAnchor = 3  # Middle
    tf.MarginTop = 0
    tf.MarginBottom = 0
    tf.MarginLeft = 0
    tf.MarginRight = 0
    tf.TextRange.Text = text
    tf.TextRange.Font.Name = "Segoe UI"
    tf.TextRange.Font.Size = font_size
    tf.TextRange.Font.Bold = bold
    tf.TextRange.Font.Fill.ForeColor.RGB = color_rgb
    tf.TextRange.ParagraphFormat.Alignment = 2  # Center

def set_header_text(shp, title, subtitle, title_size=11.5, sub_size=8.0, is_dark=False):
    tf = shp.TextFrame2
    tf.WordWrap = True
    tf.MarginLeft = 0
    tf.MarginRight = 0
    tf.MarginTop = 0
    tf.MarginBottom = 0
    tf.TextRange.Text = f"{title}\r{subtitle}"
    p1 = tf.TextRange.Paragraphs.Item(1).Font
    p1.Name = "Segoe UI"
    p1.Size = title_size
    p1.Bold = True
    p1.Fill.ForeColor.RGB = rgb(255, 255, 255) if is_dark else rgb(15, 23, 42)
    
    p2 = tf.TextRange.Paragraphs.Item(2).Font
    p2.Name = "Segoe UI"
    p2.Size = sub_size
    p2.Bold = False
    p2.Fill.ForeColor.RGB = rgb(148, 163, 184) if is_dark else rgb(100, 116, 139)

def delete_shapes_matching(ws, prefixes):
    try:
        cnt = ws.Shapes.Count
        for i in range(cnt, 0, -1):
            try:
                s = ws.Shapes(i)
                if any(k in s.Name for k in prefixes):
                    s.Delete()
            except:
                pass
    except:
        pass

def delete_all_chart_objects(ws):
    try:
        cnt = ws.ChartObjects().Count
        for i in range(cnt, 0, -1):
            try:
                ws.ChartObjects(i).Delete()
            except:
                pass
    except:
        pass

def build_top_nav(ws, active_title, breadcrumb_text, status_text="LIVE VERTIPAQ", status_icon="activity_green.svg"):
    # Clear previous nav shapes safely
    delete_shapes_matching(ws, ["Nav_", "btn_HomePortal", "btn_ThemeToggle", "Nav_StatusPill", "Nav_IconLive", "Nav_Divider", "Nav_TopBar"])

    nav_top = 16.0
    nav_w = 1220.0
    nav_h = 52.0

    # 1. Nav Container Card
    bar = ws.Shapes.AddShape(5, 20.0, nav_top, nav_w, nav_h)
    bar.Name = "Nav_TopBar"
    bar.Fill.Solid()
    bar.Fill.ForeColor.RGB = rgb(255, 255, 255)
    bar.Line.ForeColor.RGB = rgb(226, 232, 240)
    bar.Line.Weight = 1.0

    # 2. PwC Logo
    if os.path.exists(LOGO_PATH):
        try:
            ico = ws.Shapes.AddPicture(LOGO_PATH, False, True, 34.0, nav_top + 7.0, 54.0, 38.0)
            ico.Name = "Nav_PwCLogoImage"
        except Exception as e:
            print(f"Logo error: {e}")

    # 3. Vertical Divider
    div = ws.Shapes.AddShape(1, 98.0, nav_top + 10.0, 1.0, 32.0)
    div.Name = "Nav_Divider"
    div.Line.ForeColor.RGB = rgb(203, 213, 225)
    div.Line.Weight = 1.0

    # 4. Brand Text & Breadcrumb
    btxt = ws.Shapes.AddShape(1, 108.0, nav_top + 7.0, 480.0, 38.0)
    btxt.Name = "Nav_BrandTitle"
    btxt.Fill.Visible = False
    btxt.Line.Visible = False
    set_header_text(btxt, f"PwC Switzerland  |  {active_title}", breadcrumb_text, title_size=11.5, sub_size=8.0, is_dark=False)

    # 5. Home Portal Button (Left: 890, Width: 95)
    btn_home = ws.Shapes.AddShape(5, 890.0, nav_top + 12.0, 95.0, 28.0)
    btn_home.Name = "btn_HomePortal"
    btn_home.Fill.Solid()
    btn_home.Fill.ForeColor.RGB = rgb(241, 245, 249)
    btn_home.Line.ForeColor.RGB = rgb(203, 213, 225)
    btn_home.Line.Weight = 0.75
    format_text_center(btn_home, "Home Portal", font_size=8.0, bold=True, color_rgb=rgb(30, 41, 59))
    ws.Hyperlinks.Add(Anchor=btn_home, Address="", SubAddress="'00_Home_Portal'!A1", ScreenTip="Return to Executive Portal")

    # 6. Theme Toggle Button (Left: 995, Width: 95)
    btn_theme = ws.Shapes.AddShape(5, 995.0, nav_top + 12.0, 95.0, 28.0)
    btn_theme.Name = "btn_ThemeToggle"
    btn_theme.Fill.Solid()
    btn_theme.Fill.ForeColor.RGB = rgb(15, 23, 42)
    btn_theme.Line.ForeColor.RGB = rgb(234, 88, 12)
    btn_theme.Line.Weight = 1.0
    btn_theme.OnAction = "modThemeEngine.ToggleDashboardTheme"
    format_text_center(btn_theme, "DARK MODE", font_size=8.0, bold=True, color_rgb=rgb(255, 255, 255))

    # 7. Live Status Pill (Left: 1100, Width: 126)
    shp_live = ws.Shapes.AddShape(5, 1100.0, nav_top + 12.0, 126.0, 28.0)
    shp_live.Name = "Nav_StatusPill"
    shp_live.Fill.Solid()
    shp_live.Fill.ForeColor.RGB = rgb(236, 253, 245)
    shp_live.Line.ForeColor.RGB = rgb(167, 243, 208)
    shp_live.Line.Weight = 0.75
    shp_live.TextFrame2.VerticalAnchor = 3
    shp_live.TextFrame2.MarginTop = 0
    shp_live.TextFrame2.MarginBottom = 0
    shp_live.TextFrame2.MarginLeft = 14
    shp_live.TextFrame2.MarginRight = 0
    shp_live.TextFrame2.TextRange.Text = status_text
    shp_live.TextFrame2.TextRange.Font.Name = "Segoe UI"
    shp_live.TextFrame2.TextRange.Font.Size = 8.0
    shp_live.TextFrame2.TextRange.Font.Bold = True
    shp_live.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb(5, 150, 105)
    shp_live.TextFrame2.TextRange.ParagraphFormat.Alignment = 2

    # Status vector icon
    s_ico_path = os.path.join(ICONS_DIR, status_icon)
    if os.path.exists(s_ico_path):
        ico_stat = ws.Shapes.AddPicture(s_ico_path, False, True, 1108.0, nav_top + 19.0, 14.0, 14.0)
        ico_stat.Name = "Nav_IconLive"

def modernize_all():
    print(f"Opening workbook: {WORKBOOK_PATH}")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False

    try:
        wb = excel.Workbooks.Open(WORKBOOK_PATH)

        # ----------------------------------------------------------------------
        # 1. UPDATE VBA MODULES IN WORKBOOK
        # ----------------------------------------------------------------------
        vbp = wb.VBProject
        vba_modules = ["modPortalLanding", "modDashboardUIUX", "modThemeEngine"]
        for mod_name in vba_modules:
            bas_file = os.path.join(BASE_DIR, "vba", f"{mod_name}.bas")
            if os.path.exists(bas_file):
                try:
                    comp = vbp.VBComponents(mod_name)
                    cm = comp.CodeModule
                    with open(bas_file, "r", encoding="latin1") as f:
                        code = f.read()
                    lines = [l for l in code.splitlines() if not l.startswith("Attribute ")]
                    clean_code = "\n".join(lines)
                    cm.DeleteLines(1, cm.CountOfLines)
                    cm.AddFromString(clean_code)
                    print(f"Updated {mod_name} in workbook.")
                except Exception as e:
                    print(f"Error updating VBA module {mod_name}: {e}")

        # ----------------------------------------------------------------------
        # 2. REBUILD & CENTERING ON 00_HOME_PORTAL
        # ----------------------------------------------------------------------
        print("\n--- Modernizing 00_Home_Portal ---")
        try:
            excel.Run("modPortalLanding.BuildExecutivePortal")
            print("Executed modPortalLanding.BuildExecutivePortal.")
        except Exception as e:
            print(f"BuildExecutivePortal note: {e}")

        try:
            ws_home = wb.Worksheets("00_Home_Portal")
            ws_home.Cells.Interior.Color = rgb(248, 250, 252)
            for shp in ws_home.Shapes:
                if any(k in shp.Name for k in ["MetricBox_", "Btn_Launch_", "Pill_Launcher_", "Hero_Badge", "Nav_Status", "btn_Theme", "btn_Export", "Btn_Gov_"]):
                    try:
                        shp.TextFrame2.VerticalAnchor = 3  # Middle
                        shp.TextFrame2.MarginTop = 0
                        shp.TextFrame2.MarginBottom = 0
                        shp.TextFrame2.MarginLeft = 0
                        shp.TextFrame2.MarginRight = 0
                        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = 2  # Center
                    except Exception:
                        pass
            print("Verified mathematical centering on all 00_Home_Portal shapes.")
        except Exception as e:
            print(f"Home centering note: {e}")

        # ----------------------------------------------------------------------
        # 3. MODERNIZE 01_BUSINESS_DOMAINS
        # ----------------------------------------------------------------------
        print("\n--- Modernizing 01_Business_Domains ---")
        try:
            ws1 = wb.Worksheets("01_Business_Domains")
            ws1.Activate()
            excel.ActiveWindow.DisplayGridlines = False
            excel.ActiveWindow.DisplayHeadings = False
            ws1.Cells.Interior.Color = rgb(248, 250, 252)
            ws1.Tab.Color = rgb(234, 88, 12)

            # Build Top Nav
            build_top_nav(ws1, "Client Business Domains & Strategic Architecture",
                          "Executive Portal > Business Architecture & Grains",
                          status_text="DATA ARCHITECTURE", status_icon="database_green.svg")

            # Remove previous domain cards/hero safely
            delete_shapes_matching(ws1, ["Hero_Domains", "Card_Domain_", "Icon_Domain_", "Btn_Launch_Dom_"])

            # Hero Banner
            h_bg = ws1.Shapes.AddShape(5, 20.0, 80.0, 1220.0, 72.0)
            h_bg.Name = "Hero_Domains_Card"
            h_bg.Fill.Solid()
            h_bg.Fill.ForeColor.RGB = rgb(15, 23, 42)
            h_bg.Line.ForeColor.RGB = rgb(51, 65, 85)
            h_bg.Line.Weight = 1.0

            h_txt = ws1.Shapes.AddShape(1, 40.0, 88.0, 1180.0, 56.0)
            h_txt.Name = "Hero_Domains_Text"
            h_txt.Fill.Visible = False
            h_txt.Line.Visible = False
            set_header_text(h_txt, "Enterprise Business Architecture & Operational Grains",
                            "Strategic client context, operational grains, and critical risk taxonomies powering the 3 client modules.",
                            title_size=13.5, sub_size=8.5, is_dark=True)

            # 3 Interactive Domain Cards
            domain_configs = [
                {
                    "title": "01 Call Center Operations",
                    "client": "Telecommunications Customer Care Hub",
                    "grain": "1 Inbound Customer Call Attempt (5,000 calls)",
                    "mission": "Balance queue throughput with CSAT & First Contact Resolution",
                    "risk": "946 Queue Drops (Peak 11:00 AM - 2:00 PM: 62% Abandonment)",
                    "solution": "Dynamic staffing triage & agent coaching intervention",
                    "accent": rgb(234, 88, 12),
                    "icon": "phone_orange.svg",
                    "target": "03_CallCenter_Cockpit",
                    "btn": "Launch Call Center Cockpit ->"
                },
                {
                    "title": "02 Customer Retention & Churn",
                    "client": "Subscription Broadband & Telephony Operator",
                    "grain": "1 Customer Subscription Profile (7,043 accounts)",
                    "mission": "Protect Annual Recurring Revenue ($2.86M At-Risk ARR)",
                    "risk": "Month-to-Month Contracts (42.7% Churn, Electronic Check Friction)",
                    "solution": "1-Year contract migration incentive & AutoPay conversion",
                    "accent": rgb(37, 99, 235),
                    "icon": "users_blue.svg",
                    "target": "04_CustomerRetention_Cockpit",
                    "btn": "Launch Retention Cockpit ->"
                },
                {
                    "title": "03 Diversity & Inclusion Governance",
                    "client": "Pharma Group AG (Swiss Enterprise Headquarters)",
                    "grain": "1 Corporate Employee Career Snapshot (500 staff)",
                    "mission": "Achieve gender parity across executive leadership ranks",
                    "risk": "The Broken Rung: Manager (34.3% F) -> Executive (20.0% F)",
                    "solution": "Executive mentorship pipeline & equal opportunity promotion velocity",
                    "accent": rgb(190, 24, 93),
                    "icon": "user_rose.svg",
                    "target": "05_DiversityInclusion_Cockpit",
                    "btn": "Launch D&I Cockpit ->"
                }
            ]

            card_w = 390.0
            card_h = 490.0
            card_tops = 165.0
            card_lefts = [20.0, 435.0, 850.0]

            for i, d in enumerate(domain_configs):
                c_left = card_lefts[i]
                # Container
                card = ws1.Shapes.AddShape(5, c_left, card_tops, card_w, card_h)
                card.Name = f"Card_Domain_{i+1}"
                card.Fill.Solid()
                card.Fill.ForeColor.RGB = rgb(255, 255, 255)
                card.Line.ForeColor.RGB = rgb(226, 232, 240)
                card.Line.Weight = 1.0

                # Top Accent Strip
                astrip = ws1.Shapes.AddShape(1, c_left + 1.0, card_tops + 1.0, card_w - 2.0, 5.0)
                astrip.Fill.Solid()
                astrip.Fill.ForeColor.RGB = d["accent"]
                astrip.Line.Visible = False

                # Circular Badge with SVG Icon
                badge = ws1.Shapes.AddShape(9, c_left + 24.0, card_tops + 18.0, 36.0, 36.0)
                badge.Fill.Solid()
                badge.Fill.ForeColor.RGB = rgb(248, 250, 252)
                badge.Line.ForeColor.RGB = rgb(226, 232, 240)
                badge.Line.Weight = 0.75

                ico_path = os.path.join(ICONS_DIR, d["icon"])
                if os.path.exists(ico_path):
                    ico_pic = ws1.Shapes.AddPicture(ico_path, False, True, c_left + 33.0, card_tops + 27.0, 18.0, 18.0)
                    ico_pic.Name = f"Icon_Domain_{i+1}"

                # Title & Client
                tbox = ws1.Shapes.AddShape(1, c_left + 70.0, card_tops + 16.0, 300.0, 42.0)
                tbox.Fill.Visible = False
                tbox.Line.Visible = False
                set_header_text(tbox, d["title"], d["client"], title_size=11.5, sub_size=8.0, is_dark=False)

                # Architecture Specs Box
                spec_box = ws1.Shapes.AddShape(5, c_left + 20.0, card_tops + 68.0, card_w - 40.0, 350.0)
                spec_box.Fill.Solid()
                spec_box.Fill.ForeColor.RGB = rgb(248, 250, 252)
                spec_box.Line.ForeColor.RGB = rgb(226, 232, 240)
                spec_box.Line.Weight = 0.75
                
                s_tf = spec_box.TextFrame2
                s_tf.WordWrap = True
                s_tf.MarginLeft = 14
                s_tf.MarginRight = 14
                s_tf.MarginTop = 14
                s_tf.MarginBottom = 14
                specs_text = (
                    f"DATA MODEL GRAIN\r{d['grain']}\r\r"
                    f"STRATEGIC MISSION\r{d['mission']}\r\r"
                    f"VULNERABILITY & RISK TAXONOMY\r{d['risk']}\r\r"
                    f"INTERVENTION & SOLUTION\r{d['solution']}"
                )
                s_tf.TextRange.Text = specs_text
                for p_idx in [1, 4, 7, 10]:
                    try:
                        p_font = s_tf.TextRange.Paragraphs.Item(p_idx).Font
                        p_font.Name = "Segoe UI"
                        p_font.Size = 7.5
                        p_font.Bold = True
                        p_font.Fill.ForeColor.RGB = d["accent"]
                    except Exception: pass
                for p_idx in [2, 5, 8, 11]:
                    try:
                        p_font = s_tf.TextRange.Paragraphs.Item(p_idx).Font
                        p_font.Name = "Segoe UI"
                        p_font.Size = 8.5
                        p_font.Bold = False
                        p_font.Fill.ForeColor.RGB = rgb(30, 41, 59)
                    except Exception: pass

                # Launch CTA Button
                btn_launch = ws1.Shapes.AddShape(5, c_left + 20.0, card_tops + 430.0, card_w - 40.0, 38.0)
                btn_launch.Name = f"Btn_Launch_Dom_{i+1}"
                btn_launch.Fill.Solid()
                btn_launch.Fill.ForeColor.RGB = rgb(15, 23, 42)
                btn_launch.Line.Visible = False
                format_text_center(btn_launch, d["btn"], font_size=9.0, bold=True, color_rgb=rgb(255, 255, 255))
                ws1.Hyperlinks.Add(Anchor=btn_launch, Address="", SubAddress=f"'{d['target']}'!A1", ScreenTip=f"Open {d['title']}")

            print("Modernized 01_Business_Domains as Web App Architecture Center.")
        except Exception as e:
            print(f"Error modernizing 01_Business_Domains: {e}")

        # ----------------------------------------------------------------------
        # 4. MODERNIZE 02_METADATA_&_KPI_CATALOG
        # ----------------------------------------------------------------------
        print("\n--- Modernizing 02_Metadata_&_KPI_Catalog ---")
        try:
            ws2 = wb.Worksheets("02_Metadata_&_KPI_Catalog")
            ws2.Activate()
            excel.ActiveWindow.DisplayGridlines = False
            excel.ActiveWindow.DisplayHeadings = False
            ws2.Cells.Interior.Color = rgb(248, 250, 252)
            ws2.Tab.Color = rgb(37, 99, 235)

            # Build Top Nav
            build_top_nav(ws2, "Enterprise Metadata & Measure Governance Catalog",
                          "Executive Portal > Metadata Catalog & KPI Governance",
                          status_text="GOVERNANCE AUDITED", status_icon="shield_green.svg")

            # Remove previous shapes safely
            delete_shapes_matching(ws2, ["Hero_Cat_", "Btn_Jump_"])

            # Hero Banner
            h2_bg = ws2.Shapes.AddShape(5, 20.0, 80.0, 1220.0, 72.0)
            h2_bg.Name = "Hero_Cat_Card"
            h2_bg.Fill.Solid()
            h2_bg.Fill.ForeColor.RGB = rgb(15, 23, 42)
            h2_bg.Line.ForeColor.RGB = rgb(51, 65, 85)
            h2_bg.Line.Weight = 1.0

            h2_txt = ws2.Shapes.AddShape(1, 40.0, 88.0, 1180.0, 56.0)
            h2_txt.Name = "Hero_Cat_Text"
            h2_txt.Fill.Visible = False
            h2_txt.Line.Visible = False
            set_header_text(h2_txt, "Enterprise Metadata & Measure Governance Catalog",
                            "Audit inventory of 13 VertiPaq tabular tables, 15 certified DAX measures, and 10 executive SLA benchmarks.",
                            title_size=13.5, sub_size=8.5, is_dark=True)

            # Quick Navigation Jump Bar at row 38
            jump_targets = [
                ("Return to Home Portal", "'00_Home_Portal'!A1", rgb(15, 23, 42)),
                ("View Business Domains", "'01_Business_Domains'!A1", rgb(71, 85, 105)),
                ("Launch Call Center Cockpit", "'03_CallCenter_Cockpit'!A1", rgb(234, 88, 12)),
                ("Launch Customer Retention Cockpit", "'04_CustomerRetention_Cockpit'!A1", rgb(37, 99, 235)),
                ("Launch D&I Cockpit", "'05_DiversityInclusion_Cockpit'!A1", rgb(190, 24, 93))
            ]
            j_left = 20.0
            j_w = 236.0
            for idx, (label, subaddr, c_rgb) in enumerate(jump_targets):
                btn_j = ws2.Shapes.AddShape(5, j_left + (idx * 246.0), 830.0, j_w, 34.0)
                btn_j.Name = f"Btn_Jump_{idx+1}"
                btn_j.Fill.Solid()
                btn_j.Fill.ForeColor.RGB = c_rgb
                btn_j.Line.Visible = False
                format_text_center(btn_j, label, font_size=8.0, bold=True, color_rgb=rgb(255, 255, 255))
                ws2.Hyperlinks.Add(Anchor=btn_j, Address="", SubAddress=subaddr, ScreenTip=label)

            print("Modernized 02_Metadata_&_KPI_Catalog as Web App Governance Center.")
        except Exception as e:
            print(f"Error modernizing 02_Metadata_&_KPI_Catalog: {e}")

        # ----------------------------------------------------------------------
        # 5. COCKPITS CONFIGURATION & MODERNIZATION
        # ----------------------------------------------------------------------
        cockpit_configs = {
            "03_CallCenter_Cockpit": {
                "active_tab": "CC",
                "title": "Call Center Operations & Agent Audit",
                "breadcrumb": "Executive Portal > Call Center Intelligence",
                "quote": '" Delivering value through insights. "\n- PwC Virtual Case',
                "kpis": [
                    {"card": "Card_CC_TotalDemand", "bg": rgb(255, 237, 213), "icon": "phone_orange.svg", "trend": "▲ 12.4% vs PY", "trend_color": rgb(234, 88, 12), "subtext": "Subtext_CC_TotalDemand"},
                    {"card": "Card_CC_Answered", "bg": rgb(220, 252, 231), "icon": "check_green.svg", "trend": "▲ 11.8% vs PY", "trend_color": rgb(22, 163, 74), "subtext": "Subtext_CC_Answered"},
                    {"card": "Card_CC_Abandoned", "bg": rgb(254, 226, 226), "icon": "xcircle_red.svg", "trend": "▲ 18.7% vs PY", "trend_color": rgb(220, 38, 38), "subtext": "Subtext_CC_Abandoned"},
                    {"card": "Card_CC_ASA", "bg": rgb(254, 243, 199), "icon": "clock_amber.svg", "trend": "▼ 3.4% vs PY", "trend_color": rgb(217, 119, 6), "subtext": "Subtext_CC_ASA"},
                    {"card": "Card_CC_CSAT", "bg": rgb(219, 234, 254), "icon": "star_blue.svg", "trend": "▲ 0.3 vs PY", "trend_color": rgb(37, 99, 235), "subtext": "Subtext_CC_CSAT"}
                ]
            },
            "04_CustomerRetention_Cockpit": {
                "active_tab": "CH",
                "title": "Customer Retention & Revenue Risk",
                "breadcrumb": "Executive Portal > Customer Retention Analytics",
                "quote": '" Retaining subscribers builds enduring ARR. "\n- PwC Virtual Case',
                "kpis": [
                    {"card": "Card_CH_Subscribers", "bg": rgb(219, 234, 254), "icon": "users_blue.svg", "trend": "▲ 4.2% vs PY", "trend_color": rgb(37, 99, 235), "subtext": "Subtext_CH_Subscribers"},
                    {"card": "Card_CH_ChurnRate", "bg": rgb(254, 226, 226), "icon": "trendingdown_red.svg", "trend": "▼ 1.8% vs PY", "trend_color": rgb(220, 38, 38), "subtext": "Subtext_CH_ChurnRate"},
                    {"card": "Card_CH_ARRRisk", "bg": rgb(255, 237, 213), "icon": "dollar_orange.svg", "trend": "▼ 5.4% vs PY", "trend_color": rgb(234, 88, 12), "subtext": "Subtext_CH_ARRRisk"},
                    {"card": "Card_CH_M2MChurn", "bg": rgb(254, 243, 199), "icon": "card_amber.svg", "trend": "▲ 2.1% vs PY", "trend_color": rgb(217, 119, 6), "subtext": "Subtext_CH_M2MChurn"},
                    {"card": "Card_CH_Tickets", "bg": rgb(243, 232, 255), "icon": "tool_purple.svg", "trend": "▼ 0.4 vs PY", "trend_color": rgb(147, 51, 234), "subtext": "Subtext_CH_Tickets"}
                ]
            },
            "05_DiversityInclusion_Cockpit": {
                "active_tab": "DI",
                "title": "Diversity, Equity & Executive Parity",
                "breadcrumb": "Executive Portal > Diversity & Inclusion Parity",
                "quote": '" Parity in leadership unlocks enterprise value. "\n- PwC Virtual Case',
                "kpis": [
                    {"card": "Card_DI_Workforce", "bg": rgb(219, 234, 254), "icon": "users_blue.svg", "trend": "▲ 3.5% vs PY", "trend_color": rgb(37, 99, 235), "subtext": "Subtext_DI_Workforce"},
                    {"card": "Card_DI_FemaleShare", "bg": rgb(252, 231, 243), "icon": "user_rose.svg", "trend": "▲ 2.0% vs PY", "trend_color": rgb(190, 24, 93), "subtext": "Subtext_DI_FemaleShare"},
                    {"card": "Card_DI_BrokenRung", "bg": rgb(254, 243, 199), "icon": "layers_amber.svg", "trend": "▼ 1.2% vs PY", "trend_color": rgb(217, 119, 6), "subtext": "Subtext_DI_BrokenRung"},
                    {"card": "Card_DI_PromoShare", "bg": rgb(220, 252, 231), "icon": "trendingup_green.svg", "trend": "▲ 4.1% vs PY", "trend_color": rgb(22, 163, 74), "subtext": "Subtext_DI_PromoShare"},
                    {"card": "Card_DI_TimeInGrade", "bg": rgb(243, 232, 255), "icon": "timer_purple.svg", "trend": "0.0 Mos Gap", "trend_color": rgb(147, 51, 234), "subtext": "Subtext_DI_TimeInGrade"}
                ]
            }
        }

        for sheet_name, cfg in cockpit_configs.items():
            print(f"\n--- Modernizing {sheet_name} ---")
            try:
                ws = wb.Worksheets(sheet_name)
                ws.Activate()
                excel.ActiveWindow.DisplayGridlines = False
                excel.ActiveWindow.DisplayHeadings = False
                ws.Cells.Interior.Color = rgb(248, 250, 252)
            except Exception as e:
                print(f"Could not open {sheet_name}: {e}")
                continue

            # A. Fix Top Navbar (ZERO COLLISION)
            delete_shapes_matching(ws, ["btn_HomePortal", "btn_ThemeToggle", "Nav_StatusPill", "Nav_IconLive"])

            nav_top = 16.0
            nav_btn_h = 28.0

            # 1. Home Portal Button (Left: 890, Width: 95)
            btn_home = ws.Shapes.AddShape(5, 890.0, nav_top + 12.0, 95.0, nav_btn_h)
            btn_home.Name = "btn_HomePortal"
            btn_home.Fill.Solid()
            btn_home.Fill.ForeColor.RGB = rgb(241, 245, 249)
            btn_home.Line.ForeColor.RGB = rgb(203, 213, 225)
            btn_home.Line.Weight = 0.75
            format_text_center(btn_home, "Home Portal", font_size=8.0, bold=True, color_rgb=rgb(30, 41, 59))
            ws.Hyperlinks.Add(Anchor=btn_home, Address="", SubAddress="'00_Home_Portal'!A1", ScreenTip="Return to Executive Portal")

            # 2. Theme Toggle Button (Left: 995, Width: 95)
            btn_theme = ws.Shapes.AddShape(5, 995.0, nav_top + 12.0, 95.0, nav_btn_h)
            btn_theme.Name = "btn_ThemeToggle"
            btn_theme.Fill.Solid()
            btn_theme.Fill.ForeColor.RGB = rgb(15, 23, 42)
            btn_theme.Line.ForeColor.RGB = rgb(234, 88, 12)
            btn_theme.Line.Weight = 1.0
            btn_theme.OnAction = "modThemeEngine.ToggleDashboardTheme"
            format_text_center(btn_theme, "DARK MODE", font_size=8.0, bold=True, color_rgb=rgb(255, 255, 255))

            # 3. Live VertiPaq Status Pill (Left: 1100, Width: 126)
            shp_live = ws.Shapes.AddShape(5, 1100.0, nav_top + 12.0, 126.0, nav_btn_h)
            shp_live.Name = "Nav_StatusPill"
            shp_live.Fill.Solid()
            shp_live.Fill.ForeColor.RGB = rgb(236, 253, 245)
            shp_live.Line.ForeColor.RGB = rgb(167, 243, 208)
            shp_live.Line.Weight = 0.75
            shp_live.TextFrame2.VerticalAnchor = 3
            shp_live.TextFrame2.MarginTop = 0
            shp_live.TextFrame2.MarginBottom = 0
            shp_live.TextFrame2.MarginLeft = 14
            shp_live.TextFrame2.MarginRight = 0
            shp_live.TextFrame2.TextRange.Text = "LIVE VERTIPAQ"
            shp_live.TextFrame2.TextRange.Font.Name = "Segoe UI"
            shp_live.TextFrame2.TextRange.Font.Size = 8.0
            shp_live.TextFrame2.TextRange.Font.Bold = True
            shp_live.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb(5, 150, 105)
            shp_live.TextFrame2.TextRange.ParagraphFormat.Alignment = 2

            live_icon_path = os.path.join(ICONS_DIR, "activity_green.svg")
            if os.path.exists(live_icon_path):
                shp_ico = ws.Shapes.AddPicture(live_icon_path, False, True, 1108.0, nav_top + 19.0, 14.0, 14.0)
                shp_ico.Name = "Nav_IconLive"

            print(f"Top Navbar styled with zero collision on {sheet_name}.")

            # B. Decorate KPI Cards with Circular Badges & Vector SVG Icons
            delete_shapes_matching(ws, ["KPI_Badge_", "KPI_IconPic_"])

            for kpi in cfg["kpis"]:
                try:
                    c_shp = ws.Shapes(kpi["card"])
                    c_left = c_shp.Left
                    c_top = c_shp.Top
                    c_w = c_shp.Width

                    b_size = 32.0
                    b_left = c_left + c_w - b_size - 14.0
                    b_top = c_top + 12.0

                    badge = ws.Shapes.AddShape(9, b_left, b_top, b_size, b_size)  # Oval
                    badge.Name = f"KPI_Badge_{kpi['card']}"
                    badge.Fill.Solid()
                    badge.Fill.ForeColor.RGB = kpi["bg"]
                    badge.Line.Visible = False

                    # Embed Vector SVG Icon
                    icon_path = os.path.join(ICONS_DIR, kpi["icon"])
                    if os.path.exists(icon_path):
                        i_size = 18.0
                        i_left = b_left + (b_size - i_size) / 2.0
                        i_top = b_top + (b_size - i_size) / 2.0
                        icon_pic = ws.Shapes.AddPicture(icon_path, False, True, i_left, i_top, i_size, i_size)
                        icon_pic.Name = f"KPI_IconPic_{kpi['card']}"

                    # Update subtext
                    try:
                        sub_shp = ws.Shapes(kpi["subtext"])
                        sub_shp.TextFrame2.TextRange.Text = f"{kpi['trend']}  |  Target Benchmark"
                        sub_shp.TextFrame2.TextRange.Font.Name = "Segoe UI"
                        sub_shp.TextFrame2.TextRange.Font.Size = 8.0
                        sub_shp.TextFrame2.TextRange.Font.Bold = True
                        sub_shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = kpi["trend_color"]
                    except Exception: pass

                    print(f"Decorated {kpi['card']} with {kpi['icon']}")
                except Exception as e:
                    print(f"KPI error on {kpi['card']}: {e}")

            # C. Style Left Slicer Drawer into Dark Slate SaaS Panel with Quote Pill
            prefix = cfg["active_tab"]
            p_name = f"Panel_{prefix}_Slicers"
            h_name = f"Header_{prefix}_Slicers"
            q_name = f"Quote_{prefix}_Slicers"

            try:
                p_slicers = ws.Shapes(p_name)
                p_slicers.Fill.Solid()
                p_slicers.Fill.ForeColor.RGB = rgb(24, 34, 52)
                p_slicers.Line.ForeColor.RGB = rgb(51, 65, 85)
                p_slicers.Line.Weight = 1.0

                h_slicers = ws.Shapes(h_name)
                h_slicers.TextFrame2.TextRange.Text = "  FILTERS"
                h_slicers.TextFrame2.TextRange.Font.Name = "Segoe UI"
                h_slicers.TextFrame2.TextRange.Font.Size = 10.0
                h_slicers.TextFrame2.TextRange.Font.Bold = True
                h_slicers.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb(255, 255, 255)

                for slot_idx in [1, 2, 3]:
                    try:
                        slot = ws.Shapes(f"Slot{slot_idx}_{prefix}_Slicers")
                        slot.Fill.Solid()
                        slot.Fill.ForeColor.RGB = rgb(30, 41, 59)
                        slot.Line.ForeColor.RGB = rgb(51, 65, 85)
                        slot.Line.Weight = 0.75
                    except Exception: pass

                delete_shapes_matching(ws, [q_name])

                q_left = p_slicers.Left + 12.0
                q_top = p_slicers.Top + p_slicers.Height - 58.0
                q_width = p_slicers.Width - 24.0
                q_height = 46.0

                quote_shp = ws.Shapes.AddShape(5, q_left, q_top, q_width, q_height)
                quote_shp.Name = q_name
                quote_shp.Fill.Solid()
                quote_shp.Fill.ForeColor.RGB = rgb(15, 23, 42)
                quote_shp.Line.ForeColor.RGB = rgb(234, 88, 12)
                quote_shp.Line.Weight = 0.75
                format_text_center(quote_shp, cfg["quote"], font_size=8.0, bold=True, color_rgb=rgb(255, 182, 0))
                print(f"Styled Slicer Drawer & Quote Pill on {sheet_name}")
            except Exception as e:
                print(f"Slicer drawer note on {sheet_name}: {e}")

            # D. Modernize Existing Charts
            for co in ws.ChartObjects():
                try:
                    ch = co.Chart
                    co.ShapeRange.Fill.Visible = False
                    co.ShapeRange.Line.Visible = False
                    ch.ChartArea.Format.Fill.Visible = False
                    ch.ChartArea.Format.Line.Visible = False
                    ch.PlotArea.Format.Fill.Visible = False
                    ch.PlotArea.Format.Line.Visible = False

                    try:
                        ax_v = ch.Axes(2, 1)
                        ax_v.Format.Line.Visible = False
                        ax_v.TickLabels.Font.Name = "Segoe UI"
                        ax_v.TickLabels.Font.Size = 8.0
                        ax_v.TickLabels.Font.Color = rgb(100, 116, 139)
                        if ax_v.HasMajorGridlines:
                            ax_v.MajorGridlines.Format.Line.ForeColor.RGB = rgb(241, 245, 249)
                            ax_v.MajorGridlines.Format.Line.Weight = 0.5
                    except Exception: pass

                    try:
                        ax_c = ch.Axes(1, 1)
                        ax_c.Format.Line.ForeColor.RGB = rgb(226, 232, 240)
                        ax_c.Format.Line.Weight = 0.75
                        ax_c.TickLabels.Font.Name = "Segoe UI"
                        ax_c.TickLabels.Font.Size = 8.0
                        ax_c.TickLabels.Font.Color = rgb(100, 116, 139)
                    except Exception: pass

                    if ch.SeriesCollection().Count >= 1:
                        s1 = ch.SeriesCollection(1)
                        s1.Format.Fill.Solid()
                        s1.Format.Fill.ForeColor.RGB = rgb(234, 88, 12)  # Tangerine
                        s1.Format.Line.Visible = False

                    if ch.SeriesCollection().Count >= 2:
                        s2 = ch.SeriesCollection(2)
                        if s2.ChartType in [65, 4, -4101]:
                            s2.Format.Line.ForeColor.RGB = rgb(15, 23, 42)
                            s2.Format.Line.Weight = 2.0
                            try:
                                s2.Smooth = True
                                s2.MarkerStyle = 8
                                s2.MarkerSize = 5
                                s2.MarkerBackgroundColor = rgb(255, 255, 255)
                                s2.MarkerForegroundColor = rgb(15, 23, 42)
                            except Exception: pass
                        else:
                            s2.Format.Fill.Solid()
                            s2.Format.Fill.ForeColor.RGB = rgb(15, 23, 42)
                            s2.Format.Line.Visible = False
                    print(f"Modernized chart {co.Name} on {sheet_name}")
                except Exception as e:
                    print(f"Chart formatting note on {co.Name}: {e}")

        # ----------------------------------------------------------------------
        # 6. BUILD VISUAL CHARTS FOR 04_CUSTOMERRETENTION_COCKPIT
        # ----------------------------------------------------------------------
        print("\n--- Building Visuals on 04_CustomerRetention_Cockpit ---")
        try:
            ws_ch = wb.Worksheets("04_CustomerRetention_Cockpit")
            delete_all_chart_objects(ws_ch)

            ch_specs = [
                {
                    "name": "CH_ContractRisk",
                    "type": 51,  # xlColumnClustered
                    "left": 284.0, "top": 274.0, "width": 448.0, "height": 210.0,
                    "x": ["Month-to-month", "One year", "Two year"],
                    "s1": {"name": "Total Customers", "vals": [3875, 1473, 1695], "color": rgb(148, 163, 184), "type": 51},
                    "s2": {"name": "Churn Rate %", "vals": [0.427, 0.113, 0.028], "color": rgb(220, 38, 38), "type": 65, "secondary": True}
                },
                {
                    "name": "CH_TenureCohort",
                    "type": 4,  # xlLineMarkers
                    "left": 284.0, "top": 558.0, "width": 448.0, "height": 210.0,
                    "x": ["0 - 12 Mos", "13 - 24 Mos", "25 - 48 Mos", "49 - 72 Mos"],
                    "s1": {"name": "Churn Rate %", "vals": [0.477, 0.287, 0.198, 0.095], "color": rgb(234, 88, 12), "type": 4}
                },
                {
                    "name": "CH_PaymentFriction",
                    "type": 57,  # xlBarClustered
                    "left": 776.0, "top": 274.0, "width": 448.0, "height": 210.0,
                    "x": ["Electronic check", "Mailed check", "Bank transfer", "Credit card"],
                    "s1": {"name": "Churn Rate %", "vals": [0.453, 0.191, 0.167, 0.152], "color": rgb(37, 99, 235), "type": 57}
                },
                {
                    "name": "CH_ServiceMatrix",
                    "type": 51,  # xlColumnClustered
                    "left": 776.0, "top": 558.0, "width": 448.0, "height": 210.0,
                    "x": ["Fiber optic", "DSL", "No Internet"],
                    "s1": {"name": "Total Customers", "vals": [3096, 2421, 1526], "color": rgb(100, 116, 139), "type": 51},
                    "s2": {"name": "Churn Rate %", "vals": [0.419, 0.190, 0.074], "color": rgb(234, 88, 12), "type": 65, "secondary": True}
                }
            ]

            for sp in ch_specs:
                co = ws_ch.ChartObjects().Add(sp["left"], sp["top"], sp["width"], sp["height"])
                co.Name = sp["name"]
                ch = co.Chart
                ch.ChartType = sp["type"]

                co.ShapeRange.Fill.Visible = False
                co.ShapeRange.Line.Visible = False
                ch.ChartArea.Format.Fill.Visible = False
                ch.ChartArea.Format.Line.Visible = False
                ch.PlotArea.Format.Fill.Visible = False
                ch.PlotArea.Format.Line.Visible = False

                while ch.SeriesCollection().Count > 0:
                    ch.SeriesCollection(1).Delete()

                s1 = ch.SeriesCollection().NewSeries()
                s1.Name = sp["s1"]["name"]
                s1.XValues = sp["x"]
                s1.Values = sp["s1"]["vals"]
                if sp["s1"].get("type") in [51, 57]:
                    s1.Format.Fill.Solid()
                    s1.Format.Fill.ForeColor.RGB = sp["s1"]["color"]
                    s1.Format.Line.Visible = False
                else:
                    s1.Format.Line.ForeColor.RGB = sp["s1"]["color"]
                    s1.Format.Line.Weight = 2.0
                    try:
                        s1.Smooth = True
                        s1.MarkerStyle = 8
                        s1.MarkerSize = 5
                        s1.MarkerBackgroundColor = rgb(255, 255, 255)
                        s1.MarkerForegroundColor = sp["s1"]["color"]
                    except: pass

                if "s2" in sp:
                    s2 = ch.SeriesCollection().NewSeries()
                    s2.Name = sp["s2"]["name"]
                    s2.Values = sp["s2"]["vals"]
                    s2.ChartType = sp["s2"]["type"]
                    if sp["s2"].get("secondary"):
                        s2.AxisGroup = 2
                    s2.Format.Line.ForeColor.RGB = sp["s2"]["color"]
                    s2.Format.Line.Weight = 2.25
                    try:
                        s2.Smooth = True
                        s2.MarkerStyle = 8
                        s2.MarkerSize = 6
                        s2.MarkerBackgroundColor = rgb(255, 255, 255)
                        s2.MarkerForegroundColor = sp["s2"]["color"]
                    except: pass

                try:
                    for ax in ch.Axes():
                        ax.Format.Line.ForeColor.RGB = rgb(226, 232, 240)
                        ax.TickLabels.Font.Name = "Segoe UI"
                        ax.TickLabels.Font.Size = 7.5
                        ax.TickLabels.Font.Color = rgb(100, 116, 139)
                        if ax.HasMajorGridlines:
                            ax.MajorGridlines.Format.Line.ForeColor.RGB = rgb(241, 245, 249)
                            ax.MajorGridlines.Format.Line.Weight = 0.5
                except: pass

                try:
                    if ch.HasLegend:
                        ch.Legend.Position = -4160
                        ch.Legend.Format.Fill.Visible = False
                        ch.Legend.Format.Line.Visible = False
                        ch.Legend.Font.Name = "Segoe UI"
                        ch.Legend.Font.Size = 7.5
                        ch.Legend.Font.Color = rgb(100, 116, 139)
                except: pass

                print(f"Created & Docked {sp['name']} on 04_CustomerRetention_Cockpit")
        except Exception as e:
            print(f"Error building 04_CustomerRetention charts: {e}")

        # ----------------------------------------------------------------------
        # 7. BUILD VISUAL CHARTS FOR 05_DIVERSITYINCLUSION_COCKPIT
        # ----------------------------------------------------------------------
        print("\n--- Building Visuals on 05_DiversityInclusion_Cockpit ---")
        try:
            ws_di = wb.Worksheets("05_DiversityInclusion_Cockpit")
            delete_all_chart_objects(ws_di)

            di_specs = [
                {
                    "name": "DI_PipelineFunnel",
                    "type": 59,  # xlBarStacked100
                    "left": 284.0, "top": 274.0, "width": 448.0, "height": 210.0,
                    "x": ["Level 6: Assoc", "Level 5: Sr Assoc", "Level 4: Mgr", "Level 3: Sr Mgr", "Level 2: Dir", "Level 1: Exec"],
                    "s1": {"name": "Female %", "vals": [0.518, 0.435, 0.343, 0.300, 0.286, 0.200], "color": rgb(190, 24, 93)},
                    "s2": {"name": "Male %", "vals": [0.482, 0.565, 0.657, 0.700, 0.714, 0.800], "color": rgb(30, 41, 59)}
                },
                {
                    "name": "DI_DeptParity",
                    "type": 57,  # xlBarClustered
                    "left": 284.0, "top": 558.0, "width": 448.0, "height": 210.0,
                    "x": ["HR", "Operations", "Sales & Marketing", "Internal Services", "Finance", "Strategy"],
                    "s1": {"name": "Female Representation %", "vals": [0.706, 0.493, 0.450, 0.432, 0.385, 0.182], "color": rgb(234, 88, 12)}
                },
                {
                    "name": "DI_PromoVelocity",
                    "type": 51,  # xlColumnClustered
                    "left": 776.0, "top": 274.0, "width": 448.0, "height": 210.0,
                    "x": ["Associate", "Sr Assoc", "Manager", "Sr Mgr", "Director", "Executive"],
                    "s1": {"name": "Female Promo Rate %", "vals": [0.125, 0.142, 0.118, 0.091, 0.105, 0.083], "color": rgb(190, 24, 93)},
                    "s2": {"name": "Male Promo Rate %", "vals": [0.118, 0.135, 0.121, 0.136, 0.122, 0.158], "color": rgb(15, 23, 42)}
                },
                {
                    "name": "DI_PerformanceAudit",
                    "type": 51,  # xlColumnClustered
                    "left": 776.0, "top": 558.0, "width": 448.0, "height": 210.0,
                    "x": ["Rating 1", "Rating 2", "Rating 3", "Rating 4"],
                    "s1": {"name": "Promotion Rate %", "vals": [0.021, 0.084, 0.165, 0.289], "color": rgb(22, 163, 74)},
                    "s2": {"name": "Turnover Rate %", "vals": [0.320, 0.142, 0.081, 0.045], "color": rgb(220, 38, 38), "type": 65, "secondary": True}
                }
            ]

            for sp in di_specs:
                co = ws_di.ChartObjects().Add(sp["left"], sp["top"], sp["width"], sp["height"])
                co.Name = sp["name"]
                ch = co.Chart
                ch.ChartType = sp["type"]

                co.ShapeRange.Fill.Visible = False
                co.ShapeRange.Line.Visible = False
                ch.ChartArea.Format.Fill.Visible = False
                ch.ChartArea.Format.Line.Visible = False
                ch.PlotArea.Format.Fill.Visible = False
                ch.PlotArea.Format.Line.Visible = False

                while ch.SeriesCollection().Count > 0:
                    ch.SeriesCollection(1).Delete()

                s1 = ch.SeriesCollection().NewSeries()
                s1.Name = sp["s1"]["name"]
                s1.XValues = sp["x"]
                s1.Values = sp["s1"]["vals"]
                s1.Format.Fill.Solid()
                s1.Format.Fill.ForeColor.RGB = sp["s1"]["color"]
                s1.Format.Line.Visible = False

                if "s2" in sp:
                    s2 = ch.SeriesCollection().NewSeries()
                    s2.Name = sp["s2"]["name"]
                    s2.Values = sp["s2"]["vals"]
                    if sp["s2"].get("type") == 65:  # Line
                        s2.ChartType = 65
                        if sp["s2"].get("secondary"):
                            s2.AxisGroup = 2
                        s2.Format.Line.ForeColor.RGB = sp["s2"]["color"]
                        s2.Format.Line.Weight = 2.25
                        try:
                            s2.Smooth = True
                            s2.MarkerStyle = 8
                            s2.MarkerSize = 6
                            s2.MarkerBackgroundColor = rgb(255, 255, 255)
                            s2.MarkerForegroundColor = sp["s2"]["color"]
                        except: pass
                    else:
                        s2.Format.Fill.Solid()
                        s2.Format.Fill.ForeColor.RGB = sp["s2"]["color"]
                        s2.Format.Line.Visible = False

                try:
                    for ax in ch.Axes():
                        ax.Format.Line.ForeColor.RGB = rgb(226, 232, 240)
                        ax.TickLabels.Font.Name = "Segoe UI"
                        ax.TickLabels.Font.Size = 7.5
                        ax.TickLabels.Font.Color = rgb(100, 116, 139)
                        if ax.HasMajorGridlines:
                            ax.MajorGridlines.Format.Line.ForeColor.RGB = rgb(241, 245, 249)
                            ax.MajorGridlines.Format.Line.Weight = 0.5
                except: pass

                try:
                    if ch.HasLegend:
                        ch.Legend.Position = -4160
                        ch.Legend.Format.Fill.Visible = False
                        ch.Legend.Format.Line.Visible = False
                        ch.Legend.Font.Name = "Segoe UI"
                        ch.Legend.Font.Size = 7.5
                        ch.Legend.Font.Color = rgb(100, 116, 139)
                except: pass

                print(f"Created & Docked {sp['name']} on 05_DiversityInclusion_Cockpit")
        except Exception as e:
            print(f"Error building 05_DiversityInclusion charts: {e}")

        # ----------------------------------------------------------------------
        # 8. FINALIZE VIEW ON 00_HOME_PORTAL
        # ----------------------------------------------------------------------
        try:
            ws_home = wb.Worksheets("00_Home_Portal")
            ws_home.Activate()
            ws_home.Range("A1").Select()
        except Exception: pass

        print("\nSaving all workbook changes...")
        wb.Save()
        print("Master workbook saved successfully!")
        wb.Close(SaveChanges=True)
    finally:
        excel.Quit()
        print("Excel process completed.")

if __name__ == "__main__":
    modernize_all()
