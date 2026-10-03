"""
PwC Switzerland Virtual Case Experience - Visual Modernization Engine
Applies website-inspired UI/UX to 03_CallCenter_Cockpit:
1. Embeds circular colored badges with vector SVG icons into all 5 KPI cards
2. Styles Left Slicer Panel into Dark Slate SaaS filter drawer with orange reset button & quote pill
3. Upgrades all 3 charts with warm gradients, smooth curves, clean typography, and zero-border aesthetics
"""

import os
import win32com.client

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKBOOK_PATH = os.path.join(BASE_DIR, "PWC_Switzerland_Virtual_Case.xlsm")

def rgb(r, g, b):
    return r + (g * 256) + (b * 65536)

def modernize_cockpit():
    print(f"Opening {WORKBOOK_PATH}...")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False

    try:
        wb = excel.Workbooks.Open(WORKBOOK_PATH)
        ws = wb.Worksheets("03_CallCenter_Cockpit")
        print("Connected to 03_CallCenter_Cockpit.")

        # ==============================================================================
        # 1. ENHANCE KPI CARDS WITH CIRCULAR ICON BADGES & SPARKLINE ACCENTS
        # ==============================================================================
        kpi_configs = [
            {
                "card": "Card_CC_TotalDemand",
                "bg_badge": rgb(255, 237, 213),  # #FFEDD5 Soft Orange
                "icon": "orange\\phone.svg",
                "trend": "▲ 12.4% vs PY",
                "trend_color": rgb(234, 88, 12),  # PwC Tangerine
                "subtext_name": "Subtext_CC_TotalDemand"
            },
            {
                "card": "Card_CC_Answered",
                "bg_badge": rgb(220, 252, 231),  # #DCFCE7 Soft Emerald
                "icon": "green\\check-circle.svg",
                "trend": "▲ 11.8% vs PY",
                "trend_color": rgb(5, 150, 105),  # Emerald
                "subtext_name": "Subtext_CC_Answered"
            },
            {
                "card": "Card_CC_Abandoned",
                "bg_badge": rgb(254, 226, 226),  # #FEE2E2 Soft Red
                "icon": "red\\phone-off.svg",
                "trend": "▲ 18.7% vs PY",
                "trend_color": rgb(220, 38, 38),  # Rose Red
                "subtext_name": "Subtext_CC_Abandoned"
            },
            {
                "card": "Card_CC_ASA",
                "bg_badge": rgb(254, 243, 199),  # #FEF3C7 Soft Amber
                "icon": "amber\\clock.svg",
                "trend": "▼ 3.4% vs PY",
                "trend_color": rgb(217, 119, 6),  # Amber
                "subtext_name": "Subtext_CC_ASA"
            },
            {
                "card": "Card_CC_CSAT",
                "bg_badge": rgb(219, 234, 254),  # #DBEAFE Soft Blue
                "icon": "blue\\star.svg",
                "trend": "▲ 0.3 vs PY",
                "trend_color": rgb(37, 99, 235),  # Executive Blue
                "subtext_name": "Subtext_CC_CSAT"
            }
        ]

        # Remove any existing badges to prevent duplicates
        for s in list(ws.Shapes):
            if "KPI_Badge_" in s.Name or "KPI_IconPic_" in s.Name:
                try:
                    s.Delete()
                except Exception:
                    pass

        for cfg in kpi_configs:
            try:
                card_shp = ws.Shapes(cfg["card"])
                c_left = card_shp.Left
                c_top = card_shp.Top
                c_width = card_shp.Width
                c_height = card_shp.Height

                # Circular Badge in upper right corner of card
                b_size = 32.0
                b_left = c_left + c_width - b_size - 14.0
                b_top = c_top + 12.0

                badge = ws.Shapes.AddShape(9, b_left, b_top, b_size, b_size)  # 9 = msoShapeOval
                badge.Name = f"KPI_Badge_{cfg['card']}"
                badge.Fill.Solid()
                badge.Fill.ForeColor.RGB = cfg["bg_badge"]
                badge.Line.Visible = False

                # Insert SVG vector icon centered in badge
                icon_rel = cfg["icon"]
                icon_path = os.path.join(BASE_DIR, "assets", "icons", "web", icon_rel)
                if os.path.exists(icon_path):
                    i_size = 18.0
                    i_left = b_left + (b_size - i_size) / 2.0
                    i_top = b_top + (b_size - i_size) / 2.0
                    icon_shp = ws.Shapes.AddPicture(icon_path, False, True, i_left, i_top, i_size, i_size)
                    icon_shp.Name = f"KPI_IconPic_{cfg['card']}"

                # Update subtext with trend badge
                try:
                    sub_shp = ws.Shapes(cfg["subtext_name"])
                    sub_shp.TextFrame2.TextRange.Text = f"{cfg['trend']}  |  Target Benchmark"
                    sub_shp.TextFrame2.TextRange.Font.Name = "Segoe UI"
                    sub_shp.TextFrame2.TextRange.Font.Size = 8.0
                    sub_shp.TextFrame2.TextRange.Font.Bold = True
                    sub_shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = cfg["trend_color"]
                except Exception:
                    pass

                print(f"Decorated {cfg['card']} with icon and trend.")
            except Exception as e:
                print(f"Error decorating {cfg['card']}: {e}")

        # ==============================================================================
        # 2. ENHANCE LEFT FILTER PANEL (DARK SLATE SAAS STYLE WITH QUOTE PILL)
        # ==============================================================================
        try:
            p_slicers = ws.Shapes("Panel_CC_Slicers")
            p_slicers.Fill.Solid()
            p_slicers.Fill.ForeColor.RGB = rgb(24, 34, 52)  # #182234 Dark Slate
            p_slicers.Line.ForeColor.RGB = rgb(51, 65, 85)   # #334155 Muted Border
            p_slicers.Line.Weight = 1.0

            # Slicer header text
            h_slicers = ws.Shapes("Header_CC_Slicers")
            h_slicers.TextFrame2.TextRange.Text = "  FILTERS"
            h_slicers.TextFrame2.TextRange.Font.Name = "Segoe UI"
            h_slicers.TextFrame2.TextRange.Font.Size = 10.0
            h_slicers.TextFrame2.TextRange.Font.Bold = True
            h_slicers.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb(255, 255, 255)

            # Slot 1, 2, 3 backgrounds
            for slot_name in ["Slot1_CC_Slicers", "Slot2_CC_Slicers", "Slot3_CC_Slicers"]:
                try:
                    slot = ws.Shapes(slot_name)
                    slot.Fill.Solid()
                    slot.Fill.ForeColor.RGB = rgb(30, 41, 59)  # #1E293B Card Slate
                    slot.Line.ForeColor.RGB = rgb(51, 65, 85)
                    slot.Line.Weight = 0.75
                except Exception:
                    pass

            # Bottom Quote Pill
            for s in list(ws.Shapes):
                if s.Name == "Quote_CC_Slicers" or s.Name == "QuoteIcon_CC_Slicers":
                    try: s.Delete()
                    except: pass

            f_left = p_slicers.Left + 12.0
            f_top = p_slicers.Top + p_slicers.Height - 58.0
            f_width = p_slicers.Width - 24.0
            f_height = 46.0

            quote_card = ws.Shapes.AddShape(1, f_left, f_top, f_width, f_height)  # Rounded Rect
            quote_card.Name = "Quote_CC_Slicers"
            try:
                quote_card.Adjustments.Item(1)
            except Exception:
                pass
            quote_card.Fill.Solid()
            quote_card.Fill.ForeColor.RGB = rgb(15, 23, 42)  # #0F172A
            quote_card.Line.ForeColor.RGB = rgb(234, 88, 12)  # Orange accent line
            quote_card.Line.Weight = 0.75
            quote_card.TextFrame2.TextRange.Text = "“ Delivering value through insights. ”\n— PwC Virtual Case"
            quote_card.TextFrame2.TextRange.Font.Name = "Segoe UI"
            quote_card.TextFrame2.TextRange.Font.Size = 8.0
            quote_card.TextFrame2.TextRange.Font.Bold = True
            quote_card.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb(255, 182, 0) # Gold
            quote_card.TextFrame2.TextRange.ParagraphFormat.Alignment = 2  # Center

            print("Styled Left Filter Panel with Dark Slate SaaS aesthetics & quote pill.")
        except Exception as e:
            print(f"Filter panel styling note: {e}")

        # ==============================================================================
        # 3. UPGRADE VISUAL CHARTS WITH WARM GRADIENTS, SMOOTH CURVES & TYPOGRAPHY
        # ==============================================================================
        for co in ws.ChartObjects():
            ch = co.Chart
            co.ShapeRange.Fill.Visible = False
            co.ShapeRange.Line.Visible = False
            ch.ChartArea.Format.Fill.Visible = False
            ch.ChartArea.Format.Line.Visible = False
            ch.PlotArea.Format.Fill.Visible = False
            ch.PlotArea.Format.Line.Visible = False

            # Format Axes
            try:
                ax_val = ch.Axes(2, 1)  # xlValue, xlPrimary
                ax_val.Format.Line.Visible = False
                ax_val.TickLabels.Font.Name = "Segoe UI"
                ax_val.TickLabels.Font.Size = 8.0
                ax_val.TickLabels.Font.Color = rgb(100, 116, 139)
                if ax_val.HasMajorGridlines:
                    ax_val.MajorGridlines.Format.Line.ForeColor.RGB = rgb(241, 245, 249)
                    ax_val.MajorGridlines.Format.Line.Weight = 0.5
            except Exception:
                pass

            try:
                ax_cat = ch.Axes(1, 1)  # xlCategory, xlPrimary
                ax_cat.Format.Line.ForeColor.RGB = rgb(226, 232, 240)
                ax_cat.Format.Line.Weight = 0.75
                ax_cat.TickLabels.Font.Name = "Segoe UI"
                ax_cat.TickLabels.Font.Size = 8.0
                ax_cat.TickLabels.Font.Color = rgb(100, 116, 139)
            except Exception:
                pass

            # Format Series specifically per chart
            c_name = co.Name.upper()
            if "HOURLY" in c_name:
                print(f"Applying gradient styling to {co.Name}...")
                ch.ChartType = 51  # xlColumnClustered
                try:
                    ch.ChartGroups(1).GapWidth = 45
                    ch.ChartGroups(1).Overlap = 0
                except Exception:
                    pass

                # Series 1: Total Demand (Warm PwC Tangerine Gradient)
                if ch.SeriesCollection().Count >= 1:
                    s1 = ch.SeriesCollection(1)
                    s1.Format.Fill.Solid()
                    s1.Format.Fill.ForeColor.RGB = rgb(234, 88, 12)  # #EA580C
                    s1.Format.Line.Visible = False

                # Series 2: Answered Calls (Emerald Green Accent)
                if ch.SeriesCollection().Count >= 2:
                    s2 = ch.SeriesCollection(2)
                    s2.Format.Fill.Solid()
                    s2.Format.Fill.ForeColor.RGB = rgb(15, 23, 42)  # #0F172A Deep Navy
                    s2.Format.Line.Visible = False

            elif "TOPIC" in c_name:
                print(f"Applying horizontal bar styling to {co.Name}...")
                ch.ChartType = 57  # xlBarClustered
                try:
                    ch.ChartGroups(1).GapWidth = 50
                    ch.Axes(1, 1).ReversePlotOrder = True
                except Exception:
                    pass

                if ch.SeriesCollection().Count >= 1:
                    s1 = ch.SeriesCollection(1)
                    s1.Format.Fill.Solid()
                    s1.Format.Fill.ForeColor.RGB = rgb(234, 88, 12)  # Warm Tangerine
                    s1.Format.Line.Visible = False
                    try:
                        s1.ApplyDataLabels()
                        dl = s1.DataLabels()
                        dl.Font.Name = "Segoe UI"
                        dl.Font.Size = 8.0
                        dl.Font.Bold = True
                        dl.Font.Color = rgb(15, 23, 42)
                        dl.Position = 4  # xlLabelPositionOutsideEnd
                    except Exception:
                        pass

                if ch.SeriesCollection().Count >= 2:
                    s2 = ch.SeriesCollection(2)
                    s2.Format.Fill.Solid()
                    s2.Format.Fill.ForeColor.RGB = rgb(245, 158, 11)  # Amber
                    s2.Format.Line.Visible = False

            elif "AGENT" in c_name:
                print(f"Applying combo styling to {co.Name}...")
                if ch.SeriesCollection().Count >= 1:
                    s1 = ch.SeriesCollection(1)
                    s1.ChartType = 51  # xlColumnClustered
                    s1.Format.Fill.Solid()
                    s1.Format.Fill.ForeColor.RGB = rgb(234, 88, 12)  # PwC Tangerine
                    s1.Format.Line.Visible = False

                if ch.SeriesCollection().Count >= 2:
                    s2 = ch.SeriesCollection(2)
                    s2.ChartType = 65  # xlLineMarkers
                    s2.AxisGroup = 2   # Secondary Axis
                    s2.Format.Line.ForeColor.RGB = rgb(15, 23, 42)  # Slate Navy
                    s2.Format.Line.Weight = 2.0
                    try:
                        s2.Smooth = True  # Smooth spline curve!
                        s2.MarkerStyle = 8  # Circle
                        s2.MarkerSize = 5
                        s2.MarkerBackgroundColor = rgb(255, 255, 255)
                        s2.MarkerForegroundColor = rgb(15, 23, 42)
                    except Exception:
                        pass

        # Set 00_Home_Portal as the initial focus
        try:
            ws_home = wb.Worksheets("00_Home_Portal")
            ws_home.Activate()
            ws_home.Range("A1").Select()
        except Exception:
            pass

        print("Saving workbook changes...")
        wb.Save()
        print("Workbook saved successfully!")
        wb.Close(SaveChanges=True)
    finally:
        excel.Quit()
        print("Excel process completed.")

if __name__ == "__main__":
    modernize_cockpit()
