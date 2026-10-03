import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ICONS_DIR = os.path.join(BASE_DIR, "assets", "icons")
os.makedirs(ICONS_DIR, exist_ok=True)

# 1. Sparklines for the 7 KPI cards
sparklines = {
    "sparkline_orange.svg": "#EA580C",
    "sparkline_green.svg": "#16A34A",
    "sparkline_red.svg": "#DC2626",
    "sparkline_amber.svg": "#D97706",
    "sparkline_purple.svg": "#9333EA",
    "sparkline_blue.svg": "#2563EB",
    "sparkline_teal.svg": "#0D9488"
}

for fname, color in sparklines.items():
    svg_content = f'''<svg width="120" height="24" viewBox="0 0 120 24" fill="none" xmlns="http://www.w3.org/2000/svg">
  <path d="M2 18C16 18 20 8 34 13C48 18 52 4 66 9C80 14 86 2 98 7C106 10 112 14 118 15" stroke="{color}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>
</svg>'''
    with open(os.path.join(ICONS_DIR, fname), "w", encoding="utf-8") as f:
        f.write(svg_content)
    print(f"Created {fname}")

# 2. Support team illustration for Left Filter Drawer
support_team_svg = '''<svg width="180" height="70" viewBox="0 0 180 70" fill="none" xmlns="http://www.w3.org/2000/svg">
  <rect width="180" height="70" rx="8" fill="#151D2C"/>
  <!-- Speech Bubble -->
  <rect x="75" y="8" width="30" height="18" rx="4" fill="#EA580C"/>
  <path d="M85 26L89 30L91 26H85Z" fill="#EA580C"/>
  <circle cx="83" cy="17" r="2" fill="#FFFFFF"/>
  <circle cx="90" cy="17" r="2" fill="#FFFFFF"/>
  <circle cx="97" cy="17" r="2" fill="#FFFFFF"/>
  
  <!-- Left Agent -->
  <circle cx="45" cy="28" r="10" fill="#CBD5E1"/>
  <!-- Headset -->
  <path d="M35 28C35 22 40 18 45 18C50 18 55 22 55 28" stroke="#EA580C" stroke-width="2.5" stroke-linecap="round"/>
  <rect x="33" y="26" width="3" height="6" rx="1.5" fill="#EA580C"/>
  <path d="M35 30C35 34 40 36 43 35" stroke="#EA580C" stroke-width="1.5" stroke-linecap="round"/>
  <!-- Body -->
  <path d="M30 52C30 42 36 39 45 39C54 39 60 42 60 52V56H30V52Z" fill="#334155"/>
  <path d="M41 39L45 47L49 39H41Z" fill="#EA580C"/>

  <!-- Right Agent -->
  <circle cx="135" cy="28" r="10" fill="#CBD5E1"/>
  <!-- Headset -->
  <path d="M125 28C125 22 130 18 135 18C140 18 145 22 145 28" stroke="#EA580C" stroke-width="2.5" stroke-linecap="round"/>
  <rect x="144" y="26" width="3" height="6" rx="1.5" fill="#EA580C"/>
  <path d="M145 30C145 34 140 36 137 35" stroke="#EA580C" stroke-width="1.5" stroke-linecap="round"/>
  <!-- Body -->
  <path d="M120 52C120 42 126 39 135 39C144 39 150 42 150 52V56H120V52Z" fill="#334155"/>
  <path d="M131 39L135 47L139 39H131Z" fill="#EA580C"/>

  <!-- Laptop in middle -->
  <rect x="72" y="44" width="36" height="12" rx="2" fill="#475569"/>
  <rect x="68" y="55" width="44" height="3" rx="1" fill="#64748B"/>
</svg>'''

with open(os.path.join(ICONS_DIR, "support_team_illustration.svg"), "w", encoding="utf-8") as f:
    f.write(support_team_svg)
print("Created support_team_illustration.svg")
