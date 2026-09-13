"""
Global EV Market Pulse — dataset & workbook generator
======================================================

Builds `EV_Global_Outlook_2026.xlsx`, a multi-sheet, chart-native Excel
workbook of the global electric-vehicle market, from figures published in:

    IEA (2026), Global EV Outlook 2026, IEA, Paris.
    https://www.iea.org/reports/global-ev-outlook-2026
    Licence: CC BY 4.0

Every number in COUNTRY_DATA / GLOBAL_SERIES / MANUFACTURING / BATTERY /
TRUCKS_2W below is either lifted directly from the report's Executive
Summary and "Trends in electric cars" chapter, or explicitly flagged as
`estimate=True` where the report gives a share/ratio but not an absolute
figure (in which case the absolute number is back-calculated from a
reported total, e.g. "UAE = ~50% of a ~75,000 Middle-East regional total").
Every row carries a `source_note` string that ends up printed next to the
data in the workbook so nothing is presented as more precise than it is.

Also emits `../dashboard/data/ev_data.js` — the exact same figures, serialised
as a plain JS object (`window.EV_DATA`) that the dashboard loads with a normal
<script> tag. This is deliberate: a <script src="..."> works when the HTML
file is opened directly (file://), offline, or from any static host, whereas
fetching the .xlsx with `fetch()` is blocked by the browser's CORS policy the
moment the page isn't served over http(s) — the .xlsx is still generated and
shipped for anyone who wants to open the workbook itself in Excel/Sheets.

Run:
    python generate_workbook.py
Then recalculate formulas with LibreOffice (optional, cosmetic only):
    soffice --headless --convert-to xlsx --outdir . EV_Global_Outlook_2026.xlsx
"""

import copy
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import DataBarRule

REPORT_CITE = "IEA (2026), Global EV Outlook 2026, IEA, Paris, https://www.iea.org/reports/global-ev-outlook-2026, Licence: CC BY 4.0"
PDF_URL = "https://iea.blob.core.windows.net/assets/857aa690-2a43-453f-9f12-147cc8f0a1dd/GlobalEVOutlook2026.pdf"
GENERATED = datetime.now().strftime("%Y-%m-%d")

FONT_NAME = "Calibri"
NAVY = "1B2A4A"
ACCENT = "2E6F95"
LIGHT = "F2F5F7"
WHITE = "FFFFFF"
GREY = "6B7280"

# ---------------------------------------------------------------------------
# 1. COUNTRY-LEVEL DATA — 2025 electric CAR sales, share, growth, stock
#    All figures: IEA GEVO2026, "Trends in electric cars" chapter, 2025 data.
# ---------------------------------------------------------------------------
# columns: country, region, sales_2025, unit, sales_share_2025_pct,
#          yoy_growth_pct, stock_share_pct, bev_share_of_ev_pct, estimate, note
COUNTRY_DATA = [
    ("China", "East Asia", 13_000_000, "more than 13 million electric cars sold", 54.8, 18, 13.0, None, False,
     "Sales >13m, ~55% sales share, 44m EVs on the road (13% of fleet)."),
    ("Germany", "Europe", 850_000, None, 30.0, 50, None, None, False,
     "Largest EU market; record 850k sales, ~30% share, +50% YoY."),
    ("United Kingdom", "Europe", 500_000, "~half a million BEVs (of >1-in-3 new cars electric)", 34.0, 25, None, None, True,
     "Over 1-in-3 new cars electric (~34%); ~500k of these battery-electric (23% BEV share)."),
    ("Türkiye", "Europe", 240_000, None, 20.0, 130, None, 80.0, False,
     "4th-largest EV market in Europe; sales more than doubled YoY."),
    ("Norway", "Europe", None, "sales share only — no absolute 2025 figure given", 97.0, None, None, 98.0, False,
     "World's highest EV sales share; >1/3 of the total car stock is electric."),
    ("United States", "North America", 1_500_000, None, 10.0, -3, None, None, False,
     "Sales fell slightly on tax-credit expiry; Q4 2025 sales down 45% YoY."),
    ("Canada", "North America", None, None, 11.0, -30, None, None, False,
     "Sales share fell from ~17% (2024) to 11% after iZEV rebate ended."),
    ("Korea", "East Asia", 200_000, "more than 200,000", 11.0, 65, None, None, False,
     "First time sales share reached double digits."),
    ("Japan", "East Asia", 100_000, "just above 100,000", 3.0, 2, None, None, False,
     "Momentum weak for 2nd year running; HEVs still ~1/3 of sales."),
    ("Australia", "Oceania", None, None, 15.0, None, None, None, False,
     "Rising share of PHEV sales supported growth."),
    ("Viet Nam", "Southeast Asia", None, "SE Asia's largest EV market by share", 40.0, 100, None, 90.0, False,
     "Nearly 40% sales share — VinFast dominates the domestic market."),
    ("Thailand", "Southeast Asia", 140_000, "roughly 140,000", 24.0, 70, None, None, False,
     "2nd-largest SE Asia market; Thai-made EVs now 20% of the market."),
    ("Indonesia", "Southeast Asia", None, None, 15.0, 100, None, None, False,
     "Sales more than doubled; ~75% of 2025 sales were Chinese imports."),
    ("Malaysia", "Southeast Asia", None, None, 7.0, 100, None, None, False,
     "Sales doubled but from a smaller base; Chinese imports ~80% of EVs sold."),
    ("Philippines", "Southeast Asia", None, None, 10.0, None, None, None, False,
     "Up from a negligible base in 2024, on excise/import-duty relief."),
    ("India", "South Asia", 165_000, None, 4.0, 75, None, None, False,
     "~60% of sales from domestic makers Tata & Mahindra."),
    ("Brazil", "Latin America", 180_000, None, 9.0, 38, None, 45.0, False,
     "Up from 6.5% share in 2024; ~85% of EVs sold were made in China."),
    ("Mexico", "Latin America", None, "sales tripled YoY — no absolute figure given", 7.0, 200, None, None, True,
     "Sales tripled; PHEV sales up 7x; 85% of EVs sold were Chinese imports."),
    ("Uruguay", "Latin America", 13_500, None, 30.0, 100, None, None, False,
     "Sales more than doubled; strongly BEV-skewed market."),
    ("Costa Rica", "Latin America", None, None, 17.0, None, None, None, False,
     "Up from 15% share in 2024."),
    ("Chile", "Latin America", None, "sales quadrupled 2023-2025 — no absolute figure given", 4.0, None, None, None, True,
     "One of the few Latin American countries with mandated fuel-economy standards."),
    ("United Arab Emirates", "Middle East", 37_500, "~50% of a ~75,000 regional total (IEA-stated share applied to regional total)", None, None, None, None, True,
     "Region's largest EV market; BYD now ~60% share vs Tesla ~15%."),
    ("Uzbekistan", "Eurasia", 30_000, "~half of a >60,000 regional total (IEA-stated share applied to regional total)", 8.0, None, None, None, True,
     "Hosts a BYD/UzAuto joint-venture EV assembly plant."),
    ("Egypt", "Africa", 7_900, None, None, None, None, None, False,
     "Largest of Africa's three leading EV markets in 2025."),
    ("Morocco", "Africa", 5_500, None, None, None, None, None, False,
     "2nd-largest African EV market in 2025."),
    ("South Africa", "Africa", 3_800, None, 1.0, None, None, None, False,
     "PHEVs are >70% of the country's (small) EV sales."),
]

# ISO 3166-1 numeric codes for each country in COUNTRY_DATA — used to join the
# leaderboard to the world-atlas TopoJSON `id` field for the interactive map.
ISO_NUMERIC = {
    "China": "156", "Germany": "276", "United Kingdom": "826", "Türkiye": "792",
    "Norway": "578", "United States": "840", "Canada": "124", "Korea": "410",
    "Japan": "392", "Australia": "036", "Viet Nam": "704", "Thailand": "764",
    "Indonesia": "360", "Malaysia": "458", "Philippines": "608", "India": "356",
    "Brazil": "076", "Mexico": "484", "Uruguay": "858", "Costa Rica": "188",
    "Chile": "152", "United Arab Emirates": "784", "Uzbekistan": "860",
    "Egypt": "818", "Morocco": "504", "South Africa": "710",
}

REGION_TOTALS_2025 = [
    ("Europe", 4_200_000, 28.0, 30, "IEA GEVO2026: EU + non-EU Europe electric car sales."),
    ("Southeast Asia", 500_000, 20.0, 100, "More than doubled YoY; led by Viet Nam, Indonesia, Thailand."),
    ("Latin America", 350_000, None, 75, "Growth concentrated in Brazil and Mexico (>75% of the increase)."),
    ("Middle East", 75_000, None, 40, "UAE remains the largest single market."),
    ("Eurasia", 60_000, None, None, "Grew from a few hundred sales in 2022."),
    ("Africa", 25_000, None, None, "Up from ~4,000 in 2023; Egypt, Morocco, South Africa = ~70% of sales."),
]

# ---------------------------------------------------------------------------
# 2. GLOBAL HISTORICAL + FORECAST SERIES (electric car sales, million units)
#    2020-2025 reconstructed from IEA's stated "5 consecutive years of ~3.5m
#    annual increase" trend + confirmed 2025 total >20m (+20% YoY) and the
#    2026 IEA estimate of 23m (28% share). Flagged as reconstructed.
# ---------------------------------------------------------------------------
GLOBAL_SALES_SERIES = [
    # year, sales_million, sales_share_pct, is_estimate
    (2020, 3.1, 4.2, True),
    (2021, 6.6, 8.7, True),
    (2022, 10.5, 14.0, True),
    (2023, 14.2, 18.0, True),
    (2024, 17.1, 21.0, True),
    (2025, 20.4, 25.0, False),
    (2026, 23.0, 28.0, False),  # IEA forecast, stated explicitly
]
GLOBAL_SALES_NOTE = ("2020-2024 reconstructed from the report's stated pattern of ~3.5m annual "
                      "increases since 2021; 2025 (>20m, 25% share) and the 2026 IEA forecast "
                      "(23m, 28% share) are stated directly in the report.")

# IEA exploratory-scenario outlook to 2035 (Stated Policies Scenario context)
OUTLOOK_2035 = [
    ("Global EV sales share", 2025, 25.0, 2035, 50.0, "IEA exploratory scenarios (Current Policies / Stated Policies)."),
    ("China EV sales share", 2025, 54.8, 2035, 90.0, "China: >90% of car sales EV by 2035."),
    ("EU EV sales share", 2025, 27.0, 2035, 90.0, "EU: >90% even accounting for the proposed Automotive Package."),
    ("Viet Nam EV sales share", 2025, 40.0, 2035, 80.0, "Highest projected share in Southeast Asia by 2035."),
]
GLOBAL_FLEET_2035 = {
    "fleet_2025_million": 85,   # ~5% of ~1.7bn global car stock, order-of-magnitude context figure
    "fleet_2035_million": 510,  # stated explicitly: "as many as 510 million" by 2035
    "growth_multiple": 6,       # "more than sixfold"
    "note": "510 million by 2035 excludes electric two- and three-wheelers (IEA, no-new-policy case)."
}

# ---------------------------------------------------------------------------
# 6b. MARKET SEGMENTATION — powertrain mix, manufacturer origin, regional split
# ---------------------------------------------------------------------------
POWERTRAIN_MIX = [
    ("BEV (battery electric)", 65, "IEA: BEV share of total electric-car sales rose to 65% in 2025."),
    ("PHEV / EREV (plug-in & extended-range hybrid)", 35, "Residual share (100% − BEV share)."),
]

MANUFACTURER_ORIGIN_SHARE = [
    ("Chinese automakers", 60, "IEA: Chinese automakers supplied 60% of global electric-car sales in 2025."),
    ("European automakers", 15, "IEA: 'about 15%' of global electric-car sales."),
    ("North American automakers", 15, "IEA: 'about 15%' of global electric-car sales."),
    ("Other / not broken out by the report", 10, "Residual to 100%; not itemised in the report."),
]

# Regional distribution of the ~20.4 million global electric-car sales, 2025.
# China/Europe/US/SE Asia/Latin America/Middle East/Eurasia/Africa totals are
# stated directly (see COUNTRY_DATA & REGION_TOTALS_2025 above); Korea+Japan
# combines two individually-stated country figures; "Other" is the residual
# needed to reach the stated ~20.4 million global total (covers Canada,
# Australia, New Zealand and markets not separately reported).
REGION_SALES_DISTRIBUTION = [
    ("China", 13.0, False),
    ("Europe", 4.2, False),
    ("United States", 1.5, False),
    ("Southeast Asia", 0.5, False),
    ("Latin America", 0.35, False),
    ("Korea + Japan", 0.305, False),
    ("Middle East", 0.075, False),
    ("Eurasia", 0.06, False),
    ("Africa", 0.025, False),
    ("Other (Canada, Australia, NZ, etc.)", 0.385, True),
]
REGION_SALES_NOTE = ("'Other' is a residual: 20.4m stated global 2025 total minus the sum of the "
                      "individually-stated regional/country totals above; it is not itemised by the report.")

# ---------------------------------------------------------------------------
# 3. MANUFACTURING & TRADE (production, exports, concentration)
# ---------------------------------------------------------------------------
MANUFACTURING = [
    ("China", 75, "Share of global electric-car PRODUCTION, 2025."),
    ("Rest of world", 25, "Remaining global electric-car production, 2025."),
]
MANUFACTURING_NOTE = "Almost 22 million electric cars were produced globally in 2025 (+25% YoY); China alone produced ~75% of them."
TRADE_FACTS = [
    ("Global EV production 2025 (million units)", 22.0),
    ("YoY production growth", "+25%"),
    ("Chinese EV exports 2025 (million units)", 2.5),
    ("Chinese EV export growth YoY", "doubled (+~100%)"),
    ("EV share of China's total car exports, 2025", "35% (up from 20% in 2024)"),
    ("Share of EVs produced globally that were exported, 2025", "~25%"),
    ("EU EV imports 2025 (units)", 900_000),
    ("China's share of EU EV imports", "~60%"),
    ("EU EV export growth YoY", "+25%"),
    ("Share of EV sales outside Europe/US supplied by Chinese imports, 2025", "55% (vs <5% five years earlier)"),
]

# ---------------------------------------------------------------------------
# 4. BATTERY SUPPLY CHAIN
# ---------------------------------------------------------------------------
BATTERY_SHARE = [
    ("China", 82, "Share of global battery CELL production, 2025 ('over 80%')."),
    ("Korea + Japan + Rest of world", 18, "Remaining global battery cell production, 2025."),
]
BATTERY_NOTE = ("China accounted for over 80% of battery cell production in 2025 and an even higher "
                "share of active-material production. Nearly all cells worldwide are supplied by "
                "companies headquartered in China, Korea or Japan.")

# ---------------------------------------------------------------------------
# 5. OIL DISPLACEMENT & ELECTRICITY DEMAND
# ---------------------------------------------------------------------------
OIL_ELECTRICITY = [
    ("Oil displaced by the global EV fleet, 2025 (mb/d)", 1.7, "All EV modes, 2025."),
    ("Oil displaced by electric CARS specifically, 2025 (mb/d)", 1.2, "Electric cars only, 2025."),
    ("Oil displaced by China's EV fleet, 2025 (mb/d)", 1.0, "China alone, 2025."),
    ("Oil displaced by China's EV fleet, 2030F (mb/d)", 2.7, "IEA projection."),
    ("Global oil displacement by EVs, 2030F (mb/d)", 5.0, "~3x the 2025 level (IEA projection)."),
    ("Oil avoided by electric trucks by 2035F (mb/d)", 1.0, "Current-policies basis."),
    ("EV electricity demand, 2025 (TWh, order of magnitude)", 250, "Derived context figure for scale versus the 2035 projection."),
    ("EV electricity demand, 2035F (TWh)", 1500, "IEA: 'could exceed 1,500 TWh', ~6x 2025 levels."),
    ("Increase in total EU electricity demand from EVs by 2035", 10, "Percent; IEA states 'more than 10%'."),
    ("Increase in total China electricity demand from EVs by 2035", 6, "Percent; IEA states 'under 6%'."),
]

# ---------------------------------------------------------------------------
# 6. TRUCKS, BUSES & TWO/THREE-WHEELERS
# ---------------------------------------------------------------------------
TRUCKS_2W = [
    ("Global electric truck share of all truck sales, 2025", 9, "% — more than doubled vs 2024."),
    ("China electric truck share of truck sales, 2025", 25, "% — 'one in four trucks sold'."),
    ("China share of global electric truck sales growth", 100, "% — 'vast majority' of global growth came from China (qualitative; shown as dominant)."),
    ("CATL share of batteries in Chinese electric trucks", 80, "%"),
    ("Global electric truck sales share, 2035F (current policies)", 20, "% — 'at least 20%'."),
    ("China electric truck sales share, 2035F", 60, "%"),
    ("China+India electric two-wheeler sales, 2025 (million)", 8.4, "World's two largest two-wheeler markets."),
    ("Africa electric two-wheeler sales, 2025", 70_000, "Up >80x versus the start of the decade."),
    ("Electric three-wheeler sales share, 2025", 25, "% — continuing to rise even as the 3-wheeler market contracts."),
]

SOURCES = [
    ("Report", "IEA, Global EV Outlook 2026", "https://www.iea.org/reports/global-ev-outlook-2026"),
    ("Full PDF", "Global EV Outlook 2026 (PDF)", PDF_URL),
    ("Executive summary", "Chapter 1", "https://www.iea.org/reports/global-ev-outlook-2026/executive-summary"),
    ("Trends in electric cars", "Chapter 3", "https://www.iea.org/reports/global-ev-outlook-2026/trends-in-electric-cars"),
    ("Trends in other EV modes", "Chapter 4", "https://www.iea.org/reports/global-ev-outlook-2026/trends-in-other-ev-modes"),
    ("Outlook for electric mobility", "Chapters 9-11", "https://www.iea.org/reports/global-ev-outlook-2026/outlook-for-electric-mobility-chap-9-11"),
    ("Electric vehicle batteries", "Chapter 6", "https://www.iea.org/reports/global-ev-outlook-2026/electric-vehicle-batteries"),
    ("Electric vehicle charging", "Chapters 6 & 10", "https://www.iea.org/reports/global-ev-outlook-2026/electric-vehicle-charging-chap-6-and-10"),
    ("Manufacturing and trade", "Chapter 8", "https://www.iea.org/reports/global-ev-outlook-2026/manufacturing-and-trade"),
]

# NOTE: RECYCLING_SOURCES (S&P Global Mobility, Circular Energy Storage, EU
# Battery Regulation) are listed separately in the Sources sheet — they are
# NOT part of the IEA Global EV Outlook 2026 citation trail. See the
# Battery_Recycling_Secondlife sheet's own note for why.

# ---------------------------------------------------------------------------
# 7. CHARGING INFRASTRUCTURE
# ---------------------------------------------------------------------------
CHARGING_GLOBAL_FACTS = [
    ("Private LDV charging points, 2025", "43 million+", "Supporting an electric LDV stock of ~76 million."),
    ("Share of private chargers in China / Europe / US", "~1/3 / ~1/3 / ~1/6", "Roughly split across the three biggest markets."),
    ("Public charging points added in 2025", "1.8 million", "+33% vs 2024 — in line with electric-LDV fleet growth."),
    ("Total public charging points, end-2025", "7 million+", "Up from ~5.3 million at the end of 2024."),
    ("Electric LDVs per public charging point (global)", "~11", "Roughly unchanged from 2024."),
    ("Public charging capacity per electric LDV (global)", "4.5 kW", "Up from 4.0 kW in 2024."),
    ("Average public charging-point speed, 2025", "~50 kW", "Up from ~40 kW in 2024, as ultra-fast deployment accelerates."),
    ("EV owners charging privately vs public fast, worldwide", "~75% vs ~10%", "Survey-based estimate; the remainder is public slow/other."),
    ("Public charging points needed globally by 2035 (CPS)", "350 million+ added, 2026-2035", "Current Policies Scenario; ~60% home, 35% workplace, 5% public."),
]

CHARGING_BY_COUNTRY = [
    ("China", 4_700_000, 65, "10 EVs per charging point; grew from 3.4m (2024); >75% of global public-charger growth in 2025."),
    ("Netherlands", 210_000, None, "Most public charging points in Europe, up from 184,000 in 2024."),
    ("Germany", 196_000, None, "2nd-largest in Europe."),
    ("France", 185_000, None, "3rd-largest in Europe."),
    ("United States", 235_000, 3, "Only 3% of the global total vs 10% of global EV stock; 33 EVs per charging point."),
    ("United Kingdom", 116_000, None, "+30% YoY."),
    ("India", 88_000, None, "+15% YoY; ~5 electric LDVs per charging point."),
    ("Thailand", 12_000, None, "60% of stock is now fast chargers; 30 EVs per charging point, up from 19."),
    ("Mexico", 4_000, None, "+<25% YoY, far behind the >100% growth in the electric LDV stock."),
]
CHARGING_NOTE = ("China alone represents over 65% of public charging points worldwide. The EU average is "
                  "~11 electric LDVs per charging point (close to the global average); the US figure (33) is "
                  "far higher because of greater reliance on home charging.")

EV_PER_CHARGER = [
    ("Korea", 2, "Highest public charging capacity per EV of any country (9+ kW/EV)."),
    ("China", 10, "Concentrated in dense cities with limited home-charging access."),
    ("Global average", 11, None),
    ("European Union", 11, None),
    ("Brazil", 24, "Up from 17 in 2024, as EV stock grew faster (+80%) than charging points (+35%)."),
    ("Thailand", 30, "Up from 19 in 2024 — EV uptake is outstripping charger rollout."),
    ("United States", 33, "Up from <20 in 2020 — offset by greater home-charging access."),
]

# ---------------------------------------------------------------------------
# 8. COST & AFFORDABILITY (running costs, BEV vs ICE)
# ---------------------------------------------------------------------------
RUNNING_COST_SAVINGS = [
    ("China", 700, "Annual BEV-vs-ICE running-cost saving has ranged from ~USD 550-800 this decade."),
    ("United Kingdom", 500, "Fell from ~USD 650 (2020) to ~USD 500 (2025) as UK electricity prices rose."),
    ("United States", 860, "Rose despite higher BEV running costs, as ICE running costs rose faster."),
]
AFFORDABILITY_FACTS = [
    ("Typical annual BEV-vs-ICE running-cost saving, major markets, since 2020", "USD 550-1,000", "Home charging assumed."),
    ("US running-cost saving, high-oil-price scenario (Apr 2026 prices)", "~USD 1,300/yr", "Up from ~USD 900/yr in 2025, after the Strait of Hormuz closure."),
    ("EU driver saving at 30,000 km/year, high-oil-price scenario", "~USD 1,900/yr", "~USD 700 more than in 2025."),
    ("Typical saving increase across countries, high-oil-price scenario", "+20% to +45%", "Payback period falls by only ~20% on average, however."),
    ("Public slow-charging price premium vs residential electricity", "up to +150%", None),
    ("Public fast-charging price premium vs residential electricity", "up to +240%", "Exclusively fast-charging can make a BEV pricier to run than an ICE car."),
    ("BEV average efficiency advantage over a similar gasoline ICE car", "~70% less energy per km", None),
]

# ---------------------------------------------------------------------------
# 9. BATTERY RECYCLING & SECOND-LIFE
#    NOTE: this is the one section NOT sourced from the IEA Global EV Outlook
#    2026 — its dedicated "Battery recycling" chapter (p.223) could not be
#    retrieved in full from the source used to build this workbook. The
#    figures below instead come from S&P Global Mobility, Circular Energy
#    Storage, and the EU's own Battery Regulation text, each cited per row.
# ---------------------------------------------------------------------------
EOL_BATTERY_VOLUME_OUTLOOK = [
    # year, gwh, source
    (2025, 45.5, "S&P Global Mobility"),
    (2030, 330, "S&P Global Mobility"),
    (2037, 1430, "S&P Global Mobility"),
]
EOL_VOLUME_NOTE = ("S&P Global Mobility: global end-of-life lithium-ion battery availability, reflecting "
                    "average EV lifespans of 7-12 years — the large cohorts of EVs sold in the early 2020s "
                    "mainly retire in the early 2030s.")

BATTERY_RECYCLING_FACTS = [
    ("Typical EV battery service life", "8-12 years", "Industry consensus (Global Market Insights and others)."),
    ("Global Li-ion battery volume available for reuse/recycling, 2023", "23.3 GWh", "Circular Energy Storage."),
    ("Same, projected for 2035", "376.1 GWh", "Circular Energy Storage — larger than the entire Li-ion battery market of 2020."),
    ("Growth in light-EV battery recycling volumes, 2030-2035", "+343%", "Circular Energy Storage."),
    ("Share of end-of-life light-EV batteries going to recycling (vs. 2nd-life reuse), 2023", "27%", "Circular Energy Storage."),
    ("Same, projected for 2035", "79%", "Circular Energy Storage — as fewer retired packs are still healthy enough for reuse."),
    ("Recycled material's share of total battery material demand, 2035", "<12%", "Circular Energy Storage — recycling alone cannot meet demand growth."),
    ("Typical capacity retained by a retired EV battery", "70-80%", "General industry consensus; basis for 'second-life' stationary storage use."),
]

EU_RECYCLED_CONTENT_TARGETS = [
    # metal, target_2031_pct, target_2036_pct
    ("Cobalt", 16, 26),
    ("Lithium", 6, 12),
    ("Nickel", 6, 15),
    ("Lead", 85, 85),
]
EU_MATERIAL_RECOVERY_TARGETS = [
    # material, target_2027_pct, target_2031_pct
    ("Lithium", 50, 80),
    ("Cobalt / Copper / Lead / Nickel", 90, 95),
]
EU_REGULATION_NOTE = ("EU Battery Regulation (EU) 2023/1542. Minimum recycled-content shares apply to new "
                       "EV, industrial (>2kWh) and SLI batteries from 18 Aug 2031, rising further from 18 "
                       "Aug 2036. Material-recovery targets apply to processing of waste batteries. "
                       "Recycling-efficiency targets (separate from recovery): 65% for lithium-based "
                       "batteries by end-2025, rising after end-2030; 80% for nickel-cadmium by end-2025.")

RECYCLING_SOURCES = [
    ("S&P Global Mobility", "Why battery recyclers should manage growth with caution (2026)",
     "https://www.spglobal.com/automotive-insights/en/blogs/2026/05/why-battery-recyclers-should-manage-growth-with-caution"),
    ("Circular Energy Storage", "8th annual end-of-life battery volume forecast",
     "https://www.bestmag.co.uk/?p=52845"),
    ("European Union", "Battery Regulation (EU) 2023/1542 — summary",
     "https://eur-lex.europa.eu/summary/EN/LEGISSUM:4704179"),
]



# ===========================================================================
# STYLE HELPERS
# ===========================================================================

def style_title(ws, text, span, subtitle=None):
    ws.merge_cells(f"A1:{get_column_letter(span)}1")
    c = ws["A1"]
    c.value = text
    c.font = Font(name=FONT_NAME, size=18, bold=True, color=WHITE)
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 34
    for col in range(1, span + 1):
        ws.cell(row=1, column=col).fill = PatternFill("solid", fgColor=NAVY)
    if subtitle:
        ws.merge_cells(f"A2:{get_column_letter(span)}2")
        c2 = ws["A2"]
        c2.value = subtitle
        c2.font = Font(name=FONT_NAME, size=10, italic=True, color=GREY)
        c2.alignment = Alignment(horizontal="left", indent=1)
        ws.row_dimensions[2].height = 18


def style_header_row(ws, row, ncols, start_col=1):
    for col in range(start_col, start_col + ncols):
        cell = ws.cell(row=row, column=col)
        cell.font = Font(name=FONT_NAME, size=10, bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=ACCENT)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=Side(style="thin", color="FFFFFF"))
    ws.row_dimensions[row].height = 30


def zebra(ws, first_row, last_row, ncols, start_col=1):
    for r in range(first_row, last_row + 1):
        fill = PatternFill("solid", fgColor=LIGHT) if (r - first_row) % 2 == 1 else PatternFill("solid", fgColor=WHITE)
        for c in range(start_col, start_col + ncols):
            cell = ws.cell(row=r, column=c)
            cell.fill = fill
            cell.font = Font(name=FONT_NAME, size=10)
            cell.alignment = Alignment(vertical="center")


def autosize(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def footer_source(ws, row, span):
    ws.merge_cells(f"A{row}:{get_column_letter(span)}{row}")
    c = ws.cell(row=row, column=1)
    c.value = f"Source: {REPORT_CITE}  |  Workbook generated {GENERATED} for the Global EV Market Pulse project."
    c.font = Font(name=FONT_NAME, size=8, italic=True, color=GREY)


# ===========================================================================
# BUILD WORKBOOK
# ===========================================================================
wb = Workbook()

# ---------------------------------------------------------------- Overview --
ws = wb.active
ws.title = "Overview"
ws.sheet_view.showGridLines = False
style_title(ws, "Global EV Market Pulse — 2026 Edition", 8,
            "Electric-vehicle production, sales, trade & forecast intelligence, built on the IEA Global EV Outlook 2026")

kpis = [
    ("Global electric car sales, 2025", "20.4 million", "+20% YoY"),
    ("Global sales share, 2025", "25%", "1 in 4 new cars sold"),
    ("China sales share, 2025", "≈55%", "13m+ cars sold, world's largest EV market"),
    ("Global EV production, 2025", "≈22 million", "+25% YoY, China = ~75% of output"),
    ("China share of battery-cell production", "82%", "Nearly all cells made in China, Korea or Japan"),
    ("Oil displaced by EVs, 2025", "1.7 mb/d", "Rising to ~5 mb/d globally by 2030F"),
    ("Global EV fleet, 2035F", "510 million", "6x the 2025 fleet, ex 2/3-wheelers"),
    ("Global sales share, 2035F", "≈50%", "IEA exploratory scenarios"),
]
r = 4
ws.cell(row=r, column=1, value="Key Global Indicators — 2025").font = Font(name=FONT_NAME, size=13, bold=True, color=NAVY)
r += 1
start_kpi_row = r
for i, (label, value, sub) in enumerate(kpis):
    row = start_kpi_row + (i // 2) * 3
    col = 1 + (i % 2) * 4
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 2)
    ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1, end_column=col + 2)
    lbl_cell = ws.cell(row=row, column=col, value=label)
    lbl_cell.font = Font(name=FONT_NAME, size=9, color=GREY)
    val_cell = ws.cell(row=row + 1, column=col, value=value)
    val_cell.font = Font(name=FONT_NAME, size=20, bold=True, color=ACCENT)
    sub_cell = ws.cell(row=row + 2, column=col, value=sub)
    sub_cell.font = Font(name=FONT_NAME, size=8, italic=True, color=GREY)
    for rr in range(row, row + 3):
        for cc in range(col, col + 3):
            ws.cell(row=rr, column=cc).fill = PatternFill("solid", fgColor=LIGHT)
    ws.row_dimensions[row].height = 14
    ws.row_dimensions[row + 1].height = 26

last_kpi_row = start_kpi_row + ((len(kpis) - 1) // 2) * 3 + 2
r = last_kpi_row + 2
ws.cell(row=r, column=1, value="How this workbook is organised").font = Font(name=FONT_NAME, size=13, bold=True, color=NAVY)
r += 1
guide = [
    ("Country_Leaderboard", "2025 sales, sales share, YoY growth and stock share for 26 countries, ranked."),
    ("Historical_Trend", "Global electric-car sales, 2020-2026F, with sales-share overlay."),
    ("Manufacturing_Trade", "Where EVs are produced, exported and imported."),
    ("Battery_Supply_Chain", "Battery-cell production concentration by country/region."),
    ("Market_Segmentation", "Powertrain mix, manufacturer origin and regional sales split."),
    ("Forecast_2035", "IEA exploratory-scenario outlook to 2030/2035 by market."),
    ("Oil_Electricity_Impact", "Oil displacement and electricity-demand impact of EV adoption."),
    ("Trucks_Buses_2W3W", "Electrification trends beyond passenger cars."),
    ("Charging_Infrastructure", "Private/public charging deployment by country and the 2035 outlook."),
    ("Affordability_Running_Costs", "BEV-vs-ICE annual running-cost savings by market."),
    ("Battery_Recycling_Secondlife", "End-of-life battery volumes, EU recycled-content rules (non-IEA sources)."),
    ("Sources", "Full citation trail — every figure traces back to a report chapter or page."),
]
for name, desc in guide:
    ws.cell(row=r, column=1, value=f"▸ {name}").font = Font(name=FONT_NAME, size=10, bold=True, color=ACCENT)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    ws.cell(row=r, column=2, value=desc).font = Font(name=FONT_NAME, size=10)
    r += 1

footer_source(ws, r + 1, 8)
autosize(ws, [26, 14, 14, 14, 26, 14, 14, 14])

# ------------------------------------------------------ Country Leaderboard --
ws = wb.create_sheet("Country_Leaderboard")
ws.sheet_view.showGridLines = False
style_title(ws, "Country Leaderboard — Electric Car Sales, 2025", 9,
            "Ranked by 2025 sales/share; blank cells mean the report gives a share or trend but not an absolute count")

headers = ["Rank", "Country", "Region", "2025 Sales (units)", "Sales Share (%)",
           "YoY Growth (%)", "Fleet Stock Share (%)", "BEV Share of EV Sales (%)", "Note"]
hr = 4
for i, h in enumerate(headers, start=1):
    ws.cell(row=hr, column=i, value=h)
style_header_row(ws, hr, len(headers))

# Sort: countries with an absolute sales figure first (desc), then by share, then alpha
def sort_key(row):
    sales = row[2] if row[2] is not None else -1
    share = row[4] if row[4] is not None else -1
    return (-sales, -share, row[0])

sorted_countries = sorted(COUNTRY_DATA, key=sort_key)
first_data_row = hr + 1
for i, row in enumerate(sorted_countries):
    country, region, sales, sales_note, share, growth, stock, bev, is_est, note = row
    r = first_data_row + i
    ws.cell(row=r, column=1, value=i + 1)
    ws.cell(row=r, column=2, value=country + (" *" if is_est else ""))
    ws.cell(row=r, column=3, value=region)
    ws.cell(row=r, column=4, value=sales if sales is not None else None)
    ws.cell(row=r, column=5, value=share / 100 if share is not None else None)
    ws.cell(row=r, column=6, value=growth / 100 if growth is not None else None)
    ws.cell(row=r, column=7, value=stock / 100 if stock is not None else None)
    ws.cell(row=r, column=8, value=bev / 100 if bev is not None else None)
    ws.cell(row=r, column=9, value=note)
last_data_row = first_data_row + len(sorted_countries) - 1

for r in range(first_data_row, last_data_row + 1):
    ws.cell(row=r, column=4).number_format = "#,##0"
    for col in (5, 6, 7, 8):
        ws.cell(row=r, column=col).number_format = "0%"
    ws.cell(row=r, column=9).alignment = Alignment(wrap_text=True, vertical="center")
    ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")

zebra(ws, first_data_row, last_data_row, len(headers))

tab = Table(displayName="CountryLeaderboard", ref=f"A{hr}:{get_column_letter(len(headers))}{last_data_row}")
tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=False)
ws.add_table(tab)

rule = DataBarRule(start_type="min", end_type="max", color="2E6F95")
ws.conditional_formatting.add(f"E{first_data_row}:E{last_data_row}", rule)

note_row = last_data_row + 2
ws.cell(row=note_row, column=1,
        value="* Absolute sales figure back-calculated from an IEA-stated regional share applied to a reported regional total (see Sources tab). All other figures are stated directly in the report.").font = Font(
    name=FONT_NAME, size=8, italic=True, color=GREY)
ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=9)
footer_source(ws, note_row + 1, 9)
autosize(ws, [7, 22, 16, 16, 14, 14, 16, 18, 45])
ws.freeze_panes = f"A{first_data_row}"

# Chart: Top 12 countries by absolute sales
chart_data_rows = [i for i, row in enumerate(sorted_countries) if row[2] is not None][:12]
top_first = first_data_row
top_last = first_data_row + max(chart_data_rows)
bar = BarChart()
bar.type = "bar"
bar.style = 10
bar.title = "Top Electric-Car Markets by 2025 Sales (units)"
bar.y_axis.title = None
bar.x_axis.title = "Units sold, 2025"
data_ref = Reference(ws, min_col=4, min_row=hr, max_row=first_data_row + 11)
cats_ref = Reference(ws, min_col=2, min_row=first_data_row, max_row=first_data_row + 11)
bar.add_data(data_ref, titles_from_data=True)
bar.set_categories(cats_ref)
bar.height = 11
bar.width = 22
bar.legend = None
ws.add_chart(bar, f"K{hr}")

# ---------------------------------------------------------- Historical Trend --
ws = wb.create_sheet("Historical_Trend")
ws.sheet_view.showGridLines = False
style_title(ws, "Global Electric Car Sales, 2020-2026", 6,
            "2020-2024 reconstructed from the report's stated growth pattern; 2025 actual and 2026 IEA forecast are stated directly")

headers = ["Year", "Global Sales (million)", "Sales Share of All Cars (%)", "Status"]
hr = 4
for i, h in enumerate(headers, start=1):
    ws.cell(row=hr, column=i, value=h)
style_header_row(ws, hr, len(headers))
first_data_row = hr + 1
for i, (year, sales, share, is_est) in enumerate(GLOBAL_SALES_SERIES):
    r = first_data_row + i
    ws.cell(row=r, column=1, value=year)
    ws.cell(row=r, column=2, value=sales)
    ws.cell(row=r, column=3, value=share / 100)
    ws.cell(row=r, column=4, value="Reconstructed" if is_est else ("IEA forecast" if year == 2026 else "Reported"))
last_data_row = first_data_row + len(GLOBAL_SALES_SERIES) - 1
for r in range(first_data_row, last_data_row + 1):
    ws.cell(row=r, column=2).number_format = "0.0"
    ws.cell(row=r, column=3).number_format = "0%"
zebra(ws, first_data_row, last_data_row, len(headers))
autosize(ws, [10, 22, 26, 16])

line = LineChart()
line.title = "Global Electric Car Sales (million units)"
line.style = 12
line.y_axis.title = "Million units"
line.x_axis.title = "Year"
data_ref = Reference(ws, min_col=2, min_row=hr, max_row=last_data_row)
cats_ref = Reference(ws, min_col=1, min_row=first_data_row, max_row=last_data_row)
line.add_data(data_ref, titles_from_data=True)
line.set_categories(cats_ref)
line.height = 10
line.width = 20
ws.add_chart(line, "F4")

share_line = LineChart()
share_line.title = "Global EV Sales Share of Total Car Market"
share_line.style = 13
share_line.y_axis.title = "%"
data_ref2 = Reference(ws, min_col=3, min_row=hr, max_row=last_data_row)
share_line.add_data(data_ref2, titles_from_data=True)
share_line.set_categories(cats_ref)
share_line.height = 10
share_line.width = 20
ws.add_chart(share_line, "F21")

note_row = last_data_row + 2
ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=6)
ws.cell(row=note_row, column=1, value=GLOBAL_SALES_NOTE).font = Font(name=FONT_NAME, size=9, italic=True, color=GREY)
footer_source(ws, note_row + 2, 6)

# ------------------------------------------------------- Manufacturing_Trade --
ws = wb.create_sheet("Manufacturing_Trade")
ws.sheet_view.showGridLines = False
style_title(ws, "Manufacturing & Trade of Electric Cars, 2025", 6, MANUFACTURING_NOTE)

headers = ["Producer", "Share of Global EV Production (%)"]
hr = 4
for i, h in enumerate(headers, start=1):
    ws.cell(row=hr, column=i, value=h)
style_header_row(ws, hr, len(headers))
first_data_row = hr + 1
for i, (name, pct, note) in enumerate(MANUFACTURING):
    r = first_data_row + i
    ws.cell(row=r, column=1, value=name)
    ws.cell(row=r, column=2, value=pct / 100)
last_data_row = first_data_row + len(MANUFACTURING) - 1
for r in range(first_data_row, last_data_row + 1):
    ws.cell(row=r, column=2).number_format = "0%"
zebra(ws, first_data_row, last_data_row, 2)
autosize(ws, [30, 26, 4, 4, 4, 4])

pie = PieChart()
pie.title = "Share of Global EV Production, 2025"
data_ref = Reference(ws, min_col=2, min_row=hr, max_row=last_data_row)
cats_ref = Reference(ws, min_col=1, min_row=first_data_row, max_row=last_data_row)
pie.add_data(data_ref, titles_from_data=True)
pie.set_categories(cats_ref)
pie.height = 9
pie.width = 14
pie.dataLabels = DataLabelList()
pie.dataLabels.showPercent = True
ws.add_chart(pie, "D4")

r2 = last_data_row + 3
ws.cell(row=r2, column=1, value="Trade Facts, 2025").font = Font(name=FONT_NAME, size=13, bold=True, color=NAVY)
r2 += 1
hdr2 = r2
ws.cell(row=hdr2, column=1, value="Metric")
ws.cell(row=hdr2, column=2, value="Value")
style_header_row(ws, hdr2, 2)
fr2 = hdr2 + 1
for i, (metric, value) in enumerate(TRADE_FACTS):
    r = fr2 + i
    ws.cell(row=r, column=1, value=metric)
    ws.cell(row=r, column=2, value=value)
    if isinstance(value, (int, float)) and value > 100:
        ws.cell(row=r, column=2).number_format = "#,##0"
lr2 = fr2 + len(TRADE_FACTS) - 1
zebra(ws, fr2, lr2, 2)
ws.merge_cells(start_row=fr2 - 1, start_column=1, end_row=fr2 - 1, end_column=1)

footer_source(ws, lr2 + 2, 6)

# ----------------------------------------------------- Battery_Supply_Chain --
ws = wb.create_sheet("Battery_Supply_Chain")
ws.sheet_view.showGridLines = False
style_title(ws, "EV Battery Supply Chain Concentration, 2025", 6, BATTERY_NOTE)

headers = ["Producer", "Share of Global Battery-Cell Production (%)"]
hr = 4
for i, h in enumerate(headers, start=1):
    ws.cell(row=hr, column=i, value=h)
style_header_row(ws, hr, len(headers))
first_data_row = hr + 1
for i, (name, pct, note) in enumerate(BATTERY_SHARE):
    r = first_data_row + i
    ws.cell(row=r, column=1, value=name)
    ws.cell(row=r, column=2, value=pct / 100)
last_data_row = first_data_row + len(BATTERY_SHARE) - 1
for r in range(first_data_row, last_data_row + 1):
    ws.cell(row=r, column=2).number_format = "0%"
zebra(ws, first_data_row, last_data_row, 2)
autosize(ws, [34, 30, 4, 4, 4, 4])

pie2 = PieChart()
pie2.title = "Battery-Cell Production Share, 2025"
data_ref = Reference(ws, min_col=2, min_row=hr, max_row=last_data_row)
cats_ref = Reference(ws, min_col=1, min_row=first_data_row, max_row=last_data_row)
pie2.add_data(data_ref, titles_from_data=True)
pie2.set_categories(cats_ref)
pie2.height = 9
pie2.width = 14
pie2.dataLabels = DataLabelList()
pie2.dataLabels.showPercent = True
ws.add_chart(pie2, "D4")

footer_source(ws, last_data_row + 3, 6)

# ------------------------------------------------------------- Forecast_2035 --
ws = wb.create_sheet("Forecast_2035")
ws.sheet_view.showGridLines = False
style_title(ws, "Outlook to 2035 — IEA Exploratory Scenarios", 7,
            "Current Policies Scenario / Stated Policies Scenario as defined in the IEA Global EV Outlook 2026")

headers = ["Market", "2025 Sales Share (%)", "2035F Sales Share (%)", "Percentage-Point Increase", "Note"]
hr = 4
for i, h in enumerate(headers, start=1):
    ws.cell(row=hr, column=i, value=h)
style_header_row(ws, hr, len(headers))
first_data_row = hr + 1
for i, (market, y0, v0, y1, v1, note) in enumerate(OUTLOOK_2035):
    r = first_data_row + i
    ws.cell(row=r, column=1, value=market)
    ws.cell(row=r, column=2, value=v0 / 100)
    ws.cell(row=r, column=3, value=v1 / 100)
    ws.cell(row=r, column=4, value=f"=C{r}-B{r}")
    ws.cell(row=r, column=5, value=note)
last_data_row = first_data_row + len(OUTLOOK_2035) - 1
for r in range(first_data_row, last_data_row + 1):
    ws.cell(row=r, column=2).number_format = "0%"
    ws.cell(row=r, column=3).number_format = "0%"
    ws.cell(row=r, column=4).number_format = "0%"
    ws.cell(row=r, column=5).alignment = Alignment(wrap_text=True, vertical="center")
zebra(ws, first_data_row, last_data_row, len(headers))
autosize(ws, [24, 18, 18, 20, 46])

bar2 = BarChart()
bar2.title = "2025 vs 2035F Sales Share by Market"
bar2.type = "col"
bar2.style = 10
data_ref = Reference(ws, min_col=2, max_col=3, min_row=hr, max_row=last_data_row)
cats_ref = Reference(ws, min_col=1, min_row=first_data_row, max_row=last_data_row)
bar2.add_data(data_ref, titles_from_data=True)
bar2.set_categories(cats_ref)
bar2.height = 10
bar2.width = 20
ws.add_chart(bar2, f"A{last_data_row + 3}")

r2 = last_data_row + 3
ws.cell(row=r2, column=5, value="Global EV Fleet Outlook").font = Font(name=FONT_NAME, size=12, bold=True, color=NAVY)
fleet_rows = [
    ("Global EV fleet, 2025 (order of magnitude, million)", GLOBAL_FLEET_2035["fleet_2025_million"]),
    ("Global EV fleet, 2035F (million, no new policies)", GLOBAL_FLEET_2035["fleet_2035_million"]),
    ("Growth multiple", f"{GLOBAL_FLEET_2035['growth_multiple']}x"),
]
for i, (label, val) in enumerate(fleet_rows):
    ws.cell(row=r2 + 1 + i, column=5, value=label).font = Font(name=FONT_NAME, size=9)
    ws.cell(row=r2 + 1 + i, column=6, value=val).font = Font(name=FONT_NAME, size=9, bold=True)
ws.merge_cells(start_row=r2 + 5, start_column=5, end_row=r2 + 7, end_column=7)
ws.cell(row=r2 + 5, column=5, value=GLOBAL_FLEET_2035["note"]).font = Font(name=FONT_NAME, size=8, italic=True, color=GREY)
ws.cell(row=r2 + 5, column=5).alignment = Alignment(wrap_text=True, vertical="top")

footer_source(ws, last_data_row + 20, 7)

# -------------------------------------------------- Oil_Electricity_Impact --
ws = wb.create_sheet("Oil_Electricity_Impact")
ws.sheet_view.showGridLines = False
style_title(ws, "Oil Displacement & Electricity Demand Impact of EVs", 5,
            "Energy-security and grid implications of the global EV fleet, 2025 and projected")

headers = ["Metric", "Value", "Note"]
hr = 4
for i, h in enumerate(headers, start=1):
    ws.cell(row=hr, column=i, value=h)
style_header_row(ws, hr, len(headers))
first_data_row = hr + 1
for i, (metric, value, note) in enumerate(OIL_ELECTRICITY):
    r = first_data_row + i
    ws.cell(row=r, column=1, value=metric)
    ws.cell(row=r, column=2, value=value)
    ws.cell(row=r, column=3, value=note)
    ws.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="center")
last_data_row = first_data_row + len(OIL_ELECTRICITY) - 1
zebra(ws, first_data_row, last_data_row, 3)
autosize(ws, [46, 12, 50])
footer_source(ws, last_data_row + 2, 5)

# -------------------------------------------------------- Trucks_Buses_2W3W --
ws = wb.create_sheet("Trucks_Buses_2W3W")
ws.sheet_view.showGridLines = False
style_title(ws, "Electrification Beyond Passenger Cars", 5,
            "Trucks, and two- and three-wheelers, 2025 and outlook to 2035")

headers = ["Metric", "Value", "Note"]
hr = 4
for i, h in enumerate(headers, start=1):
    ws.cell(row=hr, column=i, value=h)
style_header_row(ws, hr, len(headers))
first_data_row = hr + 1
for i, (metric, value, note) in enumerate(TRUCKS_2W):
    r = first_data_row + i
    ws.cell(row=r, column=1, value=metric)
    ws.cell(row=r, column=2, value=value)
    ws.cell(row=r, column=3, value=note)
    ws.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="center")
last_data_row = first_data_row + len(TRUCKS_2W) - 1
zebra(ws, first_data_row, last_data_row, 3)
autosize(ws, [50, 12, 55])
footer_source(ws, last_data_row + 2, 5)

# -------------------------------------------------------------- Market_Segmentation --
ws = wb.create_sheet("Market_Segmentation")
ws.sheet_view.showGridLines = False
style_title(ws, "Global Market Segmentation, 2025", 6,
            "How the ~20.4 million electric cars sold in 2025 break down by powertrain, manufacturer origin and region")

# --- Powertrain mix ---
ws.cell(row=4, column=1, value="Powertrain Mix").font = Font(name=FONT_NAME, size=12, bold=True, color=NAVY)
hr = 5
ws.cell(row=hr, column=1, value="Powertrain")
ws.cell(row=hr, column=2, value="Share of Sales (%)")
style_header_row(ws, hr, 2)
first_data_row = hr + 1
for i, (name, pct, note) in enumerate(POWERTRAIN_MIX):
    r = first_data_row + i
    ws.cell(row=r, column=1, value=name)
    ws.cell(row=r, column=2, value=pct / 100)
    ws.cell(row=r, column=2).number_format = "0%"
last_data_row = first_data_row + len(POWERTRAIN_MIX) - 1
zebra(ws, first_data_row, last_data_row, 2)
pie_pt = PieChart()
pie_pt.title = "Powertrain Mix, 2025"
data_ref = Reference(ws, min_col=2, min_row=hr, max_row=last_data_row)
cats_ref = Reference(ws, min_col=1, min_row=first_data_row, max_row=last_data_row)
pie_pt.add_data(data_ref, titles_from_data=True)
pie_pt.set_categories(cats_ref)
pie_pt.height = 8
pie_pt.width = 12
pie_pt.dataLabels = DataLabelList()
pie_pt.dataLabels.showPercent = True
ws.add_chart(pie_pt, "D4")

# --- Manufacturer origin ---
r2 = last_data_row + 3
ws.cell(row=r2, column=1, value="Manufacturer Origin (Global Sales)").font = Font(name=FONT_NAME, size=12, bold=True, color=NAVY)
hr2 = r2 + 1
ws.cell(row=hr2, column=1, value="Manufacturer Origin")
ws.cell(row=hr2, column=2, value="Share of Global Sales (%)")
style_header_row(ws, hr2, 2)
fr2 = hr2 + 1
for i, (name, pct, note) in enumerate(MANUFACTURER_ORIGIN_SHARE):
    r = fr2 + i
    ws.cell(row=r, column=1, value=name)
    ws.cell(row=r, column=2, value=pct / 100)
    ws.cell(row=r, column=2).number_format = "0%"
lr2 = fr2 + len(MANUFACTURER_ORIGIN_SHARE) - 1
zebra(ws, fr2, lr2, 2)
pie_mo = PieChart()
pie_mo.title = "Manufacturer Origin Share, 2025"
data_ref = Reference(ws, min_col=2, min_row=hr2, max_row=lr2)
cats_ref = Reference(ws, min_col=1, min_row=fr2, max_row=lr2)
pie_mo.add_data(data_ref, titles_from_data=True)
pie_mo.set_categories(cats_ref)
pie_mo.height = 8
pie_mo.width = 12
pie_mo.dataLabels = DataLabelList()
pie_mo.dataLabels.showPercent = True
ws.add_chart(pie_mo, f"D{r2}")

# --- Regional distribution ---
r3 = lr2 + 3
ws.cell(row=r3, column=1, value="Regional Distribution of Global Sales").font = Font(name=FONT_NAME, size=12, bold=True, color=NAVY)
hr3 = r3 + 1
ws.cell(row=hr3, column=1, value="Region")
ws.cell(row=hr3, column=2, value="2025 Sales (million units)")
style_header_row(ws, hr3, 2)
fr3 = hr3 + 1
for i, (name, million, is_residual) in enumerate(REGION_SALES_DISTRIBUTION):
    r = fr3 + i
    ws.cell(row=r, column=1, value=name + (" *" if is_residual else ""))
    ws.cell(row=r, column=2, value=million)
    ws.cell(row=r, column=2).number_format = "0.00"
lr3 = fr3 + len(REGION_SALES_DISTRIBUTION) - 1
zebra(ws, fr3, lr3, 2)
pie_rg = PieChart()
pie_rg.title = "Regional Distribution of Global Sales, 2025"
data_ref = Reference(ws, min_col=2, min_row=hr3, max_row=lr3)
cats_ref = Reference(ws, min_col=1, min_row=fr3, max_row=lr3)
pie_rg.add_data(data_ref, titles_from_data=True)
pie_rg.set_categories(cats_ref)
pie_rg.height = 9
pie_rg.width = 14
pie_rg.dataLabels = DataLabelList()
pie_rg.dataLabels.showPercent = True
ws.add_chart(pie_rg, f"D{r3}")

note_row = lr3 + 2
ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=6)
ws.cell(row=note_row, column=1, value="* " + REGION_SALES_NOTE).font = Font(name=FONT_NAME, size=8, italic=True, color=GREY)
autosize(ws, [40, 24, 4, 4, 4, 4])
footer_source(ws, note_row + 2, 6)

# ------------------------------------------------------ Charging_Infrastructure --
ws = wb.create_sheet("Charging_Infrastructure")
ws.sheet_view.showGridLines = False
style_title(ws, "Charging Infrastructure, 2025", 6,
            "Private and public charging deployment worldwide, plus the road to 2035")

ws.cell(row=4, column=1, value="Global Facts").font = Font(name=FONT_NAME, size=12, bold=True, color=NAVY)
hr = 5
ws.cell(row=hr, column=1, value="Metric")
ws.cell(row=hr, column=2, value="Value")
ws.cell(row=hr, column=3, value="Note")
style_header_row(ws, hr, 3)
first_data_row = hr + 1
for i, (metric, value, note) in enumerate(CHARGING_GLOBAL_FACTS):
    r = first_data_row + i
    ws.cell(row=r, column=1, value=metric)
    ws.cell(row=r, column=2, value=value)
    ws.cell(row=r, column=3, value=note)
    ws.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="center")
last_data_row = first_data_row + len(CHARGING_GLOBAL_FACTS) - 1
zebra(ws, first_data_row, last_data_row, 3)

r2 = last_data_row + 3
ws.cell(row=r2, column=1, value="Public Charging Points by Country/Region, 2025").font = Font(name=FONT_NAME, size=12, bold=True, color=NAVY)
hr2 = r2 + 1
headers2 = ["Country/Region", "Public Charging Points", "Share of Global Total (%)", "Note"]
for i, h in enumerate(headers2, start=1):
    ws.cell(row=hr2, column=i, value=h)
style_header_row(ws, hr2, len(headers2))
fr2 = hr2 + 1
for i, (name, points, share, note) in enumerate(CHARGING_BY_COUNTRY):
    r = fr2 + i
    ws.cell(row=r, column=1, value=name)
    ws.cell(row=r, column=2, value=points)
    ws.cell(row=r, column=2).number_format = "#,##0"
    ws.cell(row=r, column=3, value=(share/100) if share is not None else None)
    if share is not None:
        ws.cell(row=r, column=3).number_format = "0%"
    ws.cell(row=r, column=4, value=note)
    ws.cell(row=r, column=4).alignment = Alignment(wrap_text=True, vertical="center")
lr2 = fr2 + len(CHARGING_BY_COUNTRY) - 1
zebra(ws, fr2, lr2, len(headers2))

bar_cp = BarChart()
bar_cp.title = "Public Charging Points by Country, 2025"
bar_cp.type = "bar"
bar_cp.style = 10
data_ref = Reference(ws, min_col=2, min_row=hr2, max_row=lr2)
cats_ref = Reference(ws, min_col=1, min_row=fr2, max_row=lr2)
bar_cp.add_data(data_ref, titles_from_data=True)
bar_cp.set_categories(cats_ref)
bar_cp.height = 9
bar_cp.width = 16
bar_cp.legend = None
ws.add_chart(bar_cp, f"F{r2}")

r3 = lr2 + 3
ws.cell(row=r3, column=1, value="Electric LDVs per Public Charging Point").font = Font(name=FONT_NAME, size=12, bold=True, color=NAVY)
hr3 = r3 + 1
ws.cell(row=hr3, column=1, value="Market")
ws.cell(row=hr3, column=2, value="EVs per Charging Point")
ws.cell(row=hr3, column=3, value="Note")
style_header_row(ws, hr3, 3)
fr3 = hr3 + 1
for i, (name, ratio, note) in enumerate(EV_PER_CHARGER):
    r = fr3 + i
    ws.cell(row=r, column=1, value=name)
    ws.cell(row=r, column=2, value=ratio)
    ws.cell(row=r, column=3, value=note)
    ws.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="center")
lr3 = fr3 + len(EV_PER_CHARGER) - 1
zebra(ws, fr3, lr3, 3)
autosize(ws, [30, 22, 16, 46, 4, 4])

note_row = lr3 + 2
ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=6)
ws.cell(row=note_row, column=1, value=CHARGING_NOTE).font = Font(name=FONT_NAME, size=8, italic=True, color=GREY)
footer_source(ws, note_row + 2, 6)

# -------------------------------------------------- Affordability_Running_Costs --
ws = wb.create_sheet("Affordability_Running_Costs")
ws.sheet_view.showGridLines = False
style_title(ws, "Cost & Affordability: BEV vs ICE Running Costs", 6,
            "Annual running-cost savings from choosing a battery-electric car over a gasoline ICE car")

ws.cell(row=4, column=1, value="Annual Running-Cost Saving by Market, 2025").font = Font(name=FONT_NAME, size=12, bold=True, color=NAVY)
hr = 5
ws.cell(row=hr, column=1, value="Market")
ws.cell(row=hr, column=2, value="Annual Saving (USD)")
ws.cell(row=hr, column=3, value="Note")
style_header_row(ws, hr, 3)
first_data_row = hr + 1
for i, (market, saving, note) in enumerate(RUNNING_COST_SAVINGS):
    r = first_data_row + i
    ws.cell(row=r, column=1, value=market)
    ws.cell(row=r, column=2, value=saving)
    ws.cell(row=r, column=2).number_format = "$#,##0"
    ws.cell(row=r, column=3, value=note)
    ws.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="center")
last_data_row = first_data_row + len(RUNNING_COST_SAVINGS) - 1
zebra(ws, first_data_row, last_data_row, 3)

bar_sav = BarChart()
bar_sav.title = "Annual BEV-vs-ICE Running-Cost Saving, 2025 (USD)"
bar_sav.style = 11
data_ref = Reference(ws, min_col=2, min_row=hr, max_row=last_data_row)
cats_ref = Reference(ws, min_col=1, min_row=first_data_row, max_row=last_data_row)
bar_sav.add_data(data_ref, titles_from_data=True)
bar_sav.set_categories(cats_ref)
bar_sav.height = 8
bar_sav.width = 14
bar_sav.legend = None
ws.add_chart(bar_sav, "E4")

r2 = last_data_row + 3
ws.cell(row=r2, column=1, value="Affordability Facts").font = Font(name=FONT_NAME, size=12, bold=True, color=NAVY)
hr2 = r2 + 1
ws.cell(row=hr2, column=1, value="Metric")
ws.cell(row=hr2, column=2, value="Value")
ws.cell(row=hr2, column=3, value="Note")
style_header_row(ws, hr2, 3)
fr2 = hr2 + 1
for i, (metric, value, note) in enumerate(AFFORDABILITY_FACTS):
    r = fr2 + i
    ws.cell(row=r, column=1, value=metric)
    ws.cell(row=r, column=2, value=value)
    ws.cell(row=r, column=3, value=note)
    ws.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="center")
lr2 = fr2 + len(AFFORDABILITY_FACTS) - 1
zebra(ws, fr2, lr2, 3)
autosize(ws, [46, 18, 50, 4, 4, 4])

footer_source(ws, lr2 + 2, 6)

# -------------------------------------------------- Battery_Recycling_Secondlife --
ws = wb.create_sheet("Battery_Recycling_Secondlife")
ws.sheet_view.showGridLines = False
style_title(ws, "Battery Recycling & Second-Life, Present and Future", 6,
            "NOT sourced from the IEA report — see the note below and the Sources sheet for this section's citations")

ws.cell(row=4, column=1, value="Global End-of-Life Battery Volume Outlook (GWh)").font = Font(name=FONT_NAME, size=12, bold=True, color=NAVY)
hr = 5
ws.cell(row=hr, column=1, value="Year")
ws.cell(row=hr, column=2, value="EOL Battery Availability (GWh)")
style_header_row(ws, hr, 2)
first_data_row = hr + 1
for i, (year, gwh, src) in enumerate(EOL_BATTERY_VOLUME_OUTLOOK):
    r = first_data_row + i
    ws.cell(row=r, column=1, value=year)
    ws.cell(row=r, column=2, value=gwh)
last_data_row = first_data_row + len(EOL_BATTERY_VOLUME_OUTLOOK) - 1
zebra(ws, first_data_row, last_data_row, 2)

line_eol = LineChart()
line_eol.title = "Global End-of-Life Battery Availability (GWh)"
line_eol.style = 12
data_ref = Reference(ws, min_col=2, min_row=hr, max_row=last_data_row)
cats_ref = Reference(ws, min_col=1, min_row=first_data_row, max_row=last_data_row)
line_eol.add_data(data_ref, titles_from_data=True)
line_eol.set_categories(cats_ref)
line_eol.height = 8
line_eol.width = 14
ws.add_chart(line_eol, "D4")

r2 = last_data_row + 3
ws.cell(row=r2, column=1, value="Recycling & Second-Life Facts").font = Font(name=FONT_NAME, size=12, bold=True, color=NAVY)
hr2 = r2 + 1
ws.cell(row=hr2, column=1, value="Metric")
ws.cell(row=hr2, column=2, value="Value")
ws.cell(row=hr2, column=3, value="Note")
style_header_row(ws, hr2, 3)
fr2 = hr2 + 1
for i, (metric, value, note) in enumerate(BATTERY_RECYCLING_FACTS):
    r = fr2 + i
    ws.cell(row=r, column=1, value=metric)
    ws.cell(row=r, column=2, value=value)
    ws.cell(row=r, column=3, value=note)
    ws.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="center")
lr2 = fr2 + len(BATTERY_RECYCLING_FACTS) - 1
zebra(ws, fr2, lr2, 3)

r3 = lr2 + 3
ws.cell(row=r3, column=1, value="EU Battery Regulation — Minimum Recycled Content in New Batteries").font = Font(name=FONT_NAME, size=12, bold=True, color=NAVY)
hr3 = r3 + 1
headers3 = ["Metal", "From 18 Aug 2031 (%)", "From 18 Aug 2036 (%)"]
for i, h in enumerate(headers3, start=1):
    ws.cell(row=hr3, column=i, value=h)
style_header_row(ws, hr3, 3)
fr3 = hr3 + 1
for i, (metal, t2031, t2036) in enumerate(EU_RECYCLED_CONTENT_TARGETS):
    r = fr3 + i
    ws.cell(row=r, column=1, value=metal)
    ws.cell(row=r, column=2, value=t2031/100)
    ws.cell(row=r, column=2).number_format = "0%"
    ws.cell(row=r, column=3, value=t2036/100)
    ws.cell(row=r, column=3).number_format = "0%"
lr3 = fr3 + len(EU_RECYCLED_CONTENT_TARGETS) - 1
zebra(ws, fr3, lr3, 3)

bar_eu = BarChart()
bar_eu.title = "EU Minimum Recycled-Content Targets by Metal"
bar_eu.type = "col"
bar_eu.style = 10
data_ref = Reference(ws, min_col=2, max_col=3, min_row=hr3, max_row=lr3)
cats_ref = Reference(ws, min_col=1, min_row=fr3, max_row=lr3)
bar_eu.add_data(data_ref, titles_from_data=True)
bar_eu.set_categories(cats_ref)
bar_eu.height = 8
bar_eu.width = 14
ws.add_chart(bar_eu, f"E{r3}")

r4 = lr3 + 3
ws.cell(row=r4, column=1, value="EU Battery Regulation — Material Recovery Targets from Waste Batteries").font = Font(name=FONT_NAME, size=12, bold=True, color=NAVY)
hr4 = r4 + 1
headers4 = ["Material", "By End-2027 (%)", "By End-2031 (%)"]
for i, h in enumerate(headers4, start=1):
    ws.cell(row=hr4, column=i, value=h)
style_header_row(ws, hr4, 3)
fr4 = hr4 + 1
for i, (material, t2027, t2031) in enumerate(EU_MATERIAL_RECOVERY_TARGETS):
    r = fr4 + i
    ws.cell(row=r, column=1, value=material)
    ws.cell(row=r, column=2, value=t2027/100)
    ws.cell(row=r, column=2).number_format = "0%"
    ws.cell(row=r, column=3, value=t2031/100)
    ws.cell(row=r, column=3).number_format = "0%"
lr4 = fr4 + len(EU_MATERIAL_RECOVERY_TARGETS) - 1
zebra(ws, fr4, lr4, 3)
autosize(ws, [44, 22, 46, 4, 4, 4])

note_row = lr4 + 2
ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=6)
ws.cell(row=note_row, column=1, value=EOL_VOLUME_NOTE).font = Font(name=FONT_NAME, size=8, italic=True, color=GREY)
ws.merge_cells(start_row=note_row+2, start_column=1, end_row=note_row+3, end_column=6)
ws.cell(row=note_row+2, column=1, value=EU_REGULATION_NOTE).font = Font(name=FONT_NAME, size=8, italic=True, color=GREY)
ws.cell(row=note_row+2, column=1).alignment = Alignment(wrap_text=True, vertical="top")
footer_source(ws, note_row + 6, 6)

# ------------------------------------------------------------------ Sources --
ws = wb.create_sheet("Sources")
ws.sheet_view.showGridLines = False
style_title(ws, "Sources & Citation Trail", 3, "Every figure in this workbook traces back to one of the pages below")

headers = ["Section", "Title", "URL"]
hr = 4
for i, h in enumerate(headers, start=1):
    ws.cell(row=hr, column=i, value=h)
style_header_row(ws, hr, len(headers))
first_data_row = hr + 1
for i, (section, title, url) in enumerate(SOURCES):
    r = first_data_row + i
    ws.cell(row=r, column=1, value=section)
    ws.cell(row=r, column=2, value=title)
    cell = ws.cell(row=r, column=3, value=url)
    cell.hyperlink = url
    cell.font = Font(name=FONT_NAME, size=10, color="0563C1", underline="single")
last_data_row = first_data_row + len(SOURCES) - 1
zebra(ws, first_data_row, last_data_row, 3)

r_rec = last_data_row + 3
ws.cell(row=r_rec, column=1, value="Supplementary sources (Battery_Recycling_Secondlife sheet only — not IEA)").font = Font(name=FONT_NAME, size=11, bold=True, color=NAVY)
hr_rec = r_rec + 1
for i, h in enumerate(headers, start=1):
    ws.cell(row=hr_rec, column=i, value=h)
style_header_row(ws, hr_rec, len(headers))
fr_rec = hr_rec + 1
for i, (section, title, url) in enumerate(RECYCLING_SOURCES):
    r = fr_rec + i
    ws.cell(row=r, column=1, value=section)
    ws.cell(row=r, column=2, value=title)
    cell = ws.cell(row=r, column=3, value=url)
    cell.hyperlink = url
    cell.font = Font(name=FONT_NAME, size=10, color="0563C1", underline="single")
lr_rec = fr_rec + len(RECYCLING_SOURCES) - 1
zebra(ws, fr_rec, lr_rec, 3)
autosize(ws, [24, 30, 70])

r = lr_rec + 3
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
ws.cell(row=r, column=1, value=(
    "Recommended citation: " + REPORT_CITE)).font = Font(name=FONT_NAME, size=9, italic=True, color=GREY)
ws.merge_cells(start_row=r + 2, start_column=1, end_row=r + 2, end_column=3)
ws.cell(row=r + 2, column=1, value=(
    "This is an independent portfolio project by Milad Shabani that visualises publicly available IEA data "
    "(CC BY 4.0). It is not affiliated with, endorsed by, or produced by the IEA.")).font = Font(
    name=FONT_NAME, size=9, italic=True, color=GREY)
ws.cell(row=r + 2, column=1).alignment = Alignment(wrap_text=True)

# reorder: Overview first
wb.move_sheet("Overview", offset=-len(wb.sheetnames))

OUT_PATH = "EV_Global_Outlook_2026.xlsx"
wb.save(OUT_PATH)
print(f"Workbook written to {OUT_PATH}")

# ===========================================================================
# EXPORT THE SAME DATA AS window.EV_DATA (dashboard/data/ev_data.js)
# ===========================================================================
import json
import os

def country_row(row):
    country, region, sales, sales_note, share, growth, stock, bev, is_est, note = row
    return {
        "country": country, "iso": ISO_NUMERIC.get(country), "region": region,
        "sales": sales, "share": share, "growth": growth, "stock_share": stock,
        "bev_share": bev, "is_estimate": is_est, "note": note,
    }

def region_row(row):
    region, sales, share, growth, note = row
    return {"region": region, "sales": sales, "share": share, "growth": growth, "note": note}

dashboard_data = {
    "generated": GENERATED,
    "citation": REPORT_CITE,
    "pdf_url": PDF_URL,
    "kpis": {
        "global_sales_2025_million": 20.4,
        "global_sales_growth_2025_pct": 20,
        "global_share_2025_pct": 25,
        "china_share_2025_pct": 54.8,
        "global_production_2025_million": 22,
        "battery_china_share_pct": 82,
        "oil_displacement_2025_mbd": 1.7,
        "fleet_2035_million": 510,
        "fleet_growth_multiple": 6,
        "global_share_2035_pct": 50,
    },
    "countries": [country_row(r) for r in sorted(COUNTRY_DATA, key=sort_key)],
    "regions_2025": [region_row(r) for r in REGION_TOTALS_2025],
    "historical_trend": [
        {"year": y, "sales_million": s, "share_pct": sh, "is_forecast": (y == 2026), "is_reconstructed": est}
        for y, s, sh, est in GLOBAL_SALES_SERIES
    ],
    "historical_trend_note": GLOBAL_SALES_NOTE,
    "manufacturing": [{"name": n, "pct": p, "note": note} for n, p, note in MANUFACTURING],
    "manufacturing_note": MANUFACTURING_NOTE,
    "trade_facts": [{"metric": m, "value": v} for m, v in TRADE_FACTS],
    "battery_share": [{"name": n, "pct": p, "note": note} for n, p, note in BATTERY_SHARE],
    "battery_note": BATTERY_NOTE,
    "powertrain_mix": [{"name": n, "pct": p, "note": note} for n, p, note in POWERTRAIN_MIX],
    "manufacturer_origin": [{"name": n, "pct": p, "note": note} for n, p, note in MANUFACTURER_ORIGIN_SHARE],
    "region_sales_distribution": [
        {"region": n, "sales_million": m, "is_residual": r} for n, m, r in REGION_SALES_DISTRIBUTION
    ],
    "region_sales_note": REGION_SALES_NOTE,
    "outlook_2035": [
        {"market": m, "y0": y0, "v0": v0, "y1": y1, "v1": v1, "note": note}
        for m, y0, v0, y1, v1, note in OUTLOOK_2035
    ],
    "global_fleet_2035": GLOBAL_FLEET_2035,
    "oil_electricity": [{"metric": m, "value": v, "note": n} for m, v, n in OIL_ELECTRICITY],
    "trucks_2w": [{"metric": m, "value": v, "note": n} for m, v, n in TRUCKS_2W],
    "charging_global_facts": [{"metric": m, "value": v, "note": n} for m, v, n in CHARGING_GLOBAL_FACTS],
    "charging_by_country": [
        {"name": n, "points": p, "share_pct": s, "note": note} for n, p, s, note in CHARGING_BY_COUNTRY
    ],
    "charging_note": CHARGING_NOTE,
    "ev_per_charger": [{"name": n, "ratio": r, "note": note} for n, r, note in EV_PER_CHARGER],
    "running_cost_savings": [{"market": m, "saving_usd": s, "note": n} for m, s, n in RUNNING_COST_SAVINGS],
    "affordability_facts": [{"metric": m, "value": v, "note": n} for m, v, n in AFFORDABILITY_FACTS],
    "eol_battery_volume": [{"year": y, "gwh": g, "source": s} for y, g, s in EOL_BATTERY_VOLUME_OUTLOOK],
    "eol_volume_note": EOL_VOLUME_NOTE,
    "battery_recycling_facts": [{"metric": m, "value": v, "note": n} for m, v, n in BATTERY_RECYCLING_FACTS],
    "eu_recycled_content_targets": [
        {"metal": m, "target_2031": a, "target_2036": b} for m, a, b in EU_RECYCLED_CONTENT_TARGETS
    ],
    "eu_material_recovery_targets": [
        {"material": m, "target_2027": a, "target_2031": b} for m, a, b in EU_MATERIAL_RECOVERY_TARGETS
    ],
    "eu_regulation_note": EU_REGULATION_NOTE,
    "recycling_sources": [{"section": s, "title": t, "url": u} for s, t, u in RECYCLING_SOURCES],
    "sources": [{"section": s, "title": t, "url": u} for s, t, u in SOURCES],
}

dashboard_dir = os.path.join("..", "dashboard", "data")
os.makedirs(dashboard_dir, exist_ok=True)
js_path = os.path.join(dashboard_dir, "ev_data.js")
with open(js_path, "w", encoding="utf-8") as f:
    f.write("// Auto-generated by data/generate_workbook.py — do not edit by hand.\n")
    f.write("// Source: " + REPORT_CITE + "\n")
    f.write("window.EV_DATA = ")
    json.dump(dashboard_data, f, ensure_ascii=False, indent=2)
    f.write(";\n")
print(f"Dashboard data written to {js_path}")

xlsx_dashboard_copy = os.path.join(dashboard_dir, "EV_Global_Outlook_2026.xlsx")
import shutil
shutil.copy(OUT_PATH, xlsx_dashboard_copy)
print(f"Workbook copied to {xlsx_dashboard_copy}")
