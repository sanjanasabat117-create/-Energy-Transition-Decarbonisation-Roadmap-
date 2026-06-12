# Energy Transition & Decarbonisation Roadmap

![Python](https://img.shields.io/badge/Python-3.10+-blue?logo=python&logoColor=white)
![Excel](https://img.shields.io/badge/Excel-6_Sheets-217346?logo=microsoft-excel&logoColor=white)
![Word](https://img.shields.io/badge/Report-9_Sections-2B579A?logo=microsoft-word&logoColor=white)
![Domain](https://img.shields.io/badge/Domain-ESG%20%7C%20Energy%20Transition-green)
![Data](https://img.shields.io/badge/Data-Real%20%7C%20Official%20SRs-brightgreen)
![Status](https://img.shields.io/badge/Status-Complete-success)

## Overview

A cross-sector energy transition and decarbonisation analysis comparing **4 companies** — Reliance Industries, Tata Motors, ITC Limited, and Nestlé Global — across Oil & Gas, Automotive, and FMCG sectors.

All data is sourced directly from official sustainability reports and BRSR filings (FY2024-25). The project includes a professional 9-section Word report, a 6-sheet Excel workbook, and 6 publication-quality charts.

> **100% Real Data — No synthetic or estimated figures (except where explicitly noted)**
> Sources: Reliance BRSR 2024-25 | Tata Motors BRSR 2024-25 | ITC SR 2024-25 | Nestlé SR 2024

---

## The 4 Companies

| Company | Sector | Country | Scope 1 (FY25) | RE % (FY25) | Net Zero |
|---------|--------|---------|----------------|-------------|---------|
| **Reliance Industries** | Oil & Gas | India | 36.46 MtCO₂e | 1.11% | 2035 |
| **Tata Motors** | Automotive | India | 0.071 MtCO₂e | 45% | 2040/2045 |
| **ITC Limited** | FMCG | India | 1.105 MtCO₂e | 52% | 2050 |
| **Nestlé Global** | FMCG | Switzerland | 2.82 MtCO₂e | 40.6% | 2050 |

---

## Why These 4 Companies?

Each company was selected to represent a different stage of the energy transition journey and a different sector challenge:

- **Reliance** — India's largest private company; massive legacy emissions + boldest investment commitment ($10B+)
- **Tata Motors** — Fastest RE transition in the cohort (+15pp in FY25); split Net Zero across EV and CV business
- **ITC** — India's only net carbon positive FMCG company; 17 years of sustained carbon sequestration
- **Nestlé** — Most advanced governance; only company with SBTi-validated 1.5°C targets in this cohort

---

## Project Structure

```
Energy-Transition-Decarbonisation-Roadmap/
│
├── report/
│   └── EnergyTransition_Decarbonisation_Report.docx    # 9-section Word report (~1MB)
│
├── data/
│   └── EnergyTransition_Decarbonisation_Data.xlsx      # 6-sheet Excel workbook
│
├── charts/
│   ├── chart1_scope1_2_comparison.png                   # Scope 1+2 cross-sector (log scale)
│   ├── chart2_renewable_energy_progress.png             # RE progress FY23→FY25→target
│   ├── chart3_net_zero_timeline.png                     # Net Zero Gantt timeline
│   ├── chart4_scope3_analysis.png                       # Scope 3 breakdown
│   ├── chart5_scorecard_radar.png                       # 8-dimension scorecard
│   └── chart6_intensity_investment.png                  # GHG intensity & investment
│
├── scripts/
│   ├── 01_build_excel.py                                # Generates the Excel workbook
│   ├── 02_generate_charts.py                            # Generates all 6 charts
│   └── 03_build_report.js                               # Generates Word report (Node.js)
│
├── requirements.txt
└── README.md
```

---

## Excel Workbook — 6 Sheets

| Sheet | Contents |
|-------|----------|
| 1. Company Profiles | Sector, reporting standards, employees, assurance provider |
| 2. GHG Emissions | Scope 1, 2, 3 real data; carbon sequestration; net emissions |
| 3. Energy & Renewables | Total energy, RE %, onsite solar, ENCON savings, RE targets |
| 4. Net Zero Roadmap | Milestone comparison table; SBTi status; water/waste targets |
| 5. Quantitative Summary | Chart-ready data tables for all 4 companies |
| 6. Water, Waste & Circular | Water withdrawal, recycling rates, ZWtL certifications |

---

## Word Report — 9 Sections

| Section | Contents |
|---------|----------|
| 1. Executive Summary | Key findings table, strategic highlights |
| 2. Company Profiles | Detailed profile for each of the 4 companies |
| 3. GHG Emissions Analysis | Scope 1, 2, 3 deep dive with real data tables |
| 4. Energy & Renewables | RE progress, ENCON projects, onsite solar |
| 5. Net Zero Roadmap | Gantt timeline, milestone comparison, SBTi status |
| 6. Decarbonisation Scorecard | 8-dimension analyst scorecard with colour coding |
| 7. Water, Waste & Circular Economy | Water Positive plants, ZWtL, Re.Wi.Re., EPR |
| 8. GHG Intensity & Investment | Intensity index, investment scale comparison |
| 9. Conclusions & Data Sources | Key takeaways, cross-cutting themes, source table |

---

## Charts Preview

### Chart 1 — Scope 1 & 2 Comparison (Log Scale)
Compares absolute Scope 1 and Scope 2 emissions across all 4 companies on a log scale to handle the orders-of-magnitude difference between Reliance (36.46 Mt) and Tata Motors (0.071 Mt).

### Chart 2 — Renewable Energy Progress
Tracks RE % from FY2023 to FY2025 (actual) and projects to 2030 target for all 4 companies. Highlights Tata Motors' standout +15pp improvement in a single year.

### Chart 3 — Net Zero Timeline (Gantt)
Visual timeline of each company's key milestones from 2024 to 2050 — investments, intermediate targets, and final Net Zero commitments.

### Chart 4 — Scope 3 Breakdown
Pie chart of Tata Motors' Scope 3 by category (use of sold products = 94%) + cross-company S1+2 vs S3 comparison.

### Chart 5 — Decarbonisation Scorecard (Radar)
8-dimension analyst scorecard: RE Progress, Net Zero Ambition, Scope 3 Strategy, Disclosure Quality, Carbon Sequestration, Water Stewardship, Circular Economy, SBTi Alignment.

### Chart 6 — GHG Intensity & Investment
Relative GHG intensity index (Reliance = 100 base) and approximate energy transition investment scale comparison.

---

## Key Findings

**Reliance Industries** is at an energy transition inflection point — 1.11% RE today but $10B+ committed to green hydrogen, solar, and battery manufacturing at the Jamnagar Giga Complex. Net Zero 2035 is the most ambitious target in this cohort.

**Tata Motors** demonstrated the fastest renewable energy transition: 30% → 45% in a single year (FY25). RE100 signatory. 3 plants certified Water Positive and Zero Waste to Landfill. ENCON saved 18,423 tCO₂e in FY25 alone.

**ITC Limited** is the sustainability benchmark — net carbon positive for 17 consecutive years, sequestering 6.46 MtCO₂e annually through forestry programs. GHG intensity reduced 46% vs FY2018-19 baseline.

**Nestlé Global** leads on governance — the only company with SBTi-validated 1.5°C targets. Manufacturing electricity is 98.6% renewable. Distinct FLAG (agricultural) accounting with specific science-based targets.

---

## Technical Skills Demonstrated

- **GHG Accounting** — Scope 1, 2, 3 per GHG Protocol; market-based vs location-based Scope 2
- **ESG Reporting Frameworks** — BRSR (India), GRI, SBTi, RE100, CDP, TCFD
- **Data Analysis** — Python (Pandas, Matplotlib, Seaborn, openpyxl), Excel
- **Report Writing** — Professional 9-section Word report with embedded charts (python-docx)
- **Cross-Sector Analysis** — Oil & Gas vs Automotive vs FMCG decarbonisation pathways
- **Nature-Based Solutions** — Carbon sequestration accounting, FLAG emissions

---

## Data Sources

| Company | Source | Assurance |
|---------|--------|-----------|
| Reliance Industries | BRSR 2024-25 (SEBI Framework) | Deloitte |
| Tata Motors | BRSR 2024-25 (TML+TMPVL+TPEML) | KPMG |
| ITC Limited | Sustainability Report 2024-25 (GRI/BRSR) | Third-party verified |
| Nestlé Global | Sustainability Report 2024, Appendix 4 | SGS |

---



---

## Author

**Sanjana Sabat**
ESG & Sustainability Analyst | Bengaluru, India
📧 sanjanasbat117@gmail.com

---

*All data sourced from official, publicly available sustainability reports. All analysis and interpretations are the author's own.*
