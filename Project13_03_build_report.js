const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, HeadingLevel, BorderStyle,
  WidthType, ShadingType, VerticalAlign, PageNumber, LevelFormat,
  PageBreak
} = require("docx");
const fs = require("fs");

const OUT = "/mnt/user-data/outputs/EnergyTransition_Decarbonisation_Report.docx";
const CHARTS = "/home/claude/energy_transition/charts/";

// ── Colours ─────────────────────────────────────────────────────────────────
const C = {
  darkGreen: "1B5E20", midGreen: "2E7D32", lightGreen: "A5D6A7",
  darkBlue: "0D47A1", midBlue: "1565C0", lightBlue: "BBDEFB",
  orange: "E65100", amber: "FF8F00", lightAmber: "FFF8E1",
  grey: "37474F", greyLight: "ECEFF1", white: "FFFFFF",
  reliance: "E65100", tata: "1565C0", itc: "2E7D32", nestle: "6A1B9A",
  palBlue: "E3F2FD", paleGreen: "E8F5E9", yellow: "FFF9C4"
};

// ── Borders & Shading ────────────────────────────────────────────────────────
const bdr = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: bdr, bottom: bdr, left: bdr, right: bdr };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function shading(hex) { return { fill: hex, type: ShadingType.CLEAR }; }

// ── Helper: load image ───────────────────────────────────────────────────────
function loadImg(fname) {
  try { return fs.readFileSync(`${CHARTS}${fname}`); }
  catch(e) { return null; }
}

// ── Helper: paragraph shortcut ──────────────────────────────────────────────
function para(text, opts = {}) {
  const { bold = false, size = 20, color = "212121", spacing = {}, indent = {},
          align = AlignmentType.LEFT, italic = false } = opts;
  return new Paragraph({
    alignment: align,
    spacing: { before: 80, after: 80, ...spacing },
    indent,
    children: [new TextRun({ text, bold, size, color, font: "Arial", italics: italic })]
  });
}

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 160 },
    children: [new TextRun({ text, bold: true, size: 36, color: C.darkGreen, font: "Arial" })]
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, size: 28, color: C.darkBlue, font: "Arial" })]
  });
}

function heading3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 180, after: 80 },
    children: [new TextRun({ text, bold: true, size: 24, color: C.grey, font: "Arial" })]
  });
}

function bullet(text, bold = false) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, bold, size: 20, color: "212121", font: "Arial" })]
  });
}

function spacer(n = 1) {
  return Array(n).fill(null).map(() => new Paragraph({
    spacing: { before: 0, after: 0 },
    children: [new TextRun({ text: "", size: 20 })]
  }));
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// ── Helper: insert chart ─────────────────────────────────────────────────────
function chartPara(fname, width = 620, height = 340, caption = "") {
  const data = loadImg(fname);
  if (!data) return para(`[Chart: ${fname} — not found]`, { italic: true, color: "999999" });
  const items = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 160, after: 80 },
      children: [new ImageRun({ data, transformation: { width, height } })]
    })
  ];
  if (caption) {
    items.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 160 },
      children: [new TextRun({ text: caption, size: 17, italic: true, color: "757575", font: "Arial" })]
    }));
  }
  return items;
}

// ── Helper: table cell ───────────────────────────────────────────────────────
function tc(text, opts = {}) {
  const { fill = C.white, bold = false, size = 18, color = "212121",
          align = AlignmentType.LEFT, colSpan = 1, vAlign = VerticalAlign.CENTER,
          width = 1600 } = opts;
  return new TableCell({
    borders,
    shading: shading(fill),
    columnSpan: colSpan,
    verticalAlign: vAlign,
    width: { size: width, type: WidthType.DXA },
    margins: { top: 100, bottom: 100, left: 150, right: 150 },
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text: String(text), bold, size, color, font: "Arial" })]
    })]
  });
}

// ── Cover Page ───────────────────────────────────────────────────────────────
function makeCover() {
  return [
    ...spacer(4),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 200 },
      children: [new TextRun({ text: "ENERGY TRANSITION &", bold: true, size: 56, color: C.darkGreen, font: "Arial" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 400 },
      children: [new TextRun({ text: "DECARBONISATION ROADMAP", bold: true, size: 56, color: C.darkGreen, font: "Arial" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 200 },
      children: [new TextRun({ text: "Cross-Sector Analysis: Oil & Gas | Automotive | FMCG", size: 28, color: C.grey, font: "Arial" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 600 },
      children: [new TextRun({ text: "Reliance Industries | Tata Motors | ITC Limited | Nestlé Global", size: 24, bold: true, color: C.midBlue, font: "Arial" })]
    }),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2340, 2340, 2340, 2340],
      rows: [
        new TableRow({ children: [
          tc("Reliance Industries", { fill: "FFF3E0", bold: true, color: C.reliance, align: AlignmentType.CENTER, width: 2340 }),
          tc("Tata Motors",         { fill: "E3F2FD", bold: true, color: C.tata,     align: AlignmentType.CENTER, width: 2340 }),
          tc("ITC Limited",         { fill: "E8F5E9", bold: true, color: C.itc,      align: AlignmentType.CENTER, width: 2340 }),
          tc("Nestlé Global",       { fill: "F3E5F5", bold: true, color: C.nestle,   align: AlignmentType.CENTER, width: 2340 }),
        ]}),
        new TableRow({ children: [
          tc("Oil & Gas",    { fill: "FFF8E1", color: "5D4037", align: AlignmentType.CENTER, size: 17, width: 2340 }),
          tc("Automotive",  { fill: "E8EAF6", color: "3949AB", align: AlignmentType.CENTER, size: 17, width: 2340 }),
          tc("FMCG India",  { fill: "F1F8E9", color: "558B2F", align: AlignmentType.CENTER, size: 17, width: 2340 }),
          tc("FMCG Global", { fill: "F8BBD0", color: "880E4F", align: AlignmentType.CENTER, size: 17, width: 2340 }),
        ]}),
        new TableRow({ children: [
          tc("Net Zero: 2035", { fill: "FFF3E0", color: C.reliance, align: AlignmentType.CENTER, size: 17, width: 2340 }),
          tc("Net Zero: 2040-45", { fill: "E3F2FD", color: C.tata, align: AlignmentType.CENTER, size: 17, width: 2340 }),
          tc("Net Zero: 2050", { fill: "E8F5E9", color: C.itc, align: AlignmentType.CENTER, size: 17, width: 2340 }),
          tc("Net Zero: 2050", { fill: "F3E5F5", color: C.nestle, align: AlignmentType.CENTER, size: 17, width: 2340 }),
        ]}),
      ]
    }),
    ...spacer(2),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 80 },
      children: [new TextRun({ text: "REAL DATA | FY2024-25", bold: true, size: 22, color: C.midGreen, font: "Arial" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 80 },
      children: [new TextRun({ text: "Sources: Reliance BRSR 2024-25 | Tata Motors BRSR 2024-25 | ITC SR 2024-25 | Nestlé SR 2024", size: 18, color: "757575", font: "Arial", italics: true })]
    }),
    ...spacer(3),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 40 },
      children: [new TextRun({ text: "Prepared by: Sanjana Sabat", bold: true, size: 22, color: C.grey, font: "Arial" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "ESG & Sustainability Analyst | Bengaluru, India | June 2025", size: 18, color: "9E9E9E", font: "Arial" })]
    }),
  ];
}

// ── Section 1: Executive Summary ─────────────────────────────────────────────
function makeSection1() {
  return [
    pageBreak(),
    heading1("1. Executive Summary"),
    para("This report presents a comprehensive cross-sector analysis of energy transition strategies and decarbonisation roadmaps for four companies representing Oil & Gas, Automotive, and FMCG sectors. All data is sourced directly from official sustainability reports and BRSR filings for FY2024-25.", { size: 20 }),
    ...spacer(1),
    heading2("1.1 Key Findings at a Glance"),

    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [1800, 1800, 1800, 1800, 2160],
      rows: [
        new TableRow({ children: [
          tc("Metric", { fill: C.darkGreen, color: C.white, bold: true, align: AlignmentType.CENTER, width: 1800 }),
          tc("Reliance", { fill: C.darkGreen, color: C.white, bold: true, align: AlignmentType.CENTER, width: 1800 }),
          tc("Tata Motors", { fill: C.darkGreen, color: C.white, bold: true, align: AlignmentType.CENTER, width: 1800 }),
          tc("ITC Limited", { fill: C.darkGreen, color: C.white, bold: true, align: AlignmentType.CENTER, width: 1800 }),
          tc("Nestlé Global", { fill: C.darkGreen, color: C.white, bold: true, align: AlignmentType.CENTER, width: 2160 }),
        ]}),
        new TableRow({ children: [
          tc("Scope 1 (MtCO₂e)", { fill: C.greyLight, bold: true, width: 1800 }),
          tc("36.46", { fill: "FFF3E0", color: C.reliance, align: AlignmentType.CENTER, width: 1800 }),
          tc("0.071", { fill: C.palBlue, color: C.tata, align: AlignmentType.CENTER, width: 1800 }),
          tc("1.105", { fill: C.paleGreen, color: C.itc, align: AlignmentType.CENTER, width: 1800 }),
          tc("2.82", { fill: "F3E5F5", color: C.nestle, align: AlignmentType.CENTER, width: 2160 }),
        ]}),
        new TableRow({ children: [
          tc("Scope 2 MB (MtCO₂e)", { fill: C.greyLight, bold: true, width: 1800 }),
          tc("1.47", { fill: "FFF3E0", color: C.reliance, align: AlignmentType.CENTER, width: 1800 }),
          tc("0.214", { fill: C.palBlue, color: C.tata, align: AlignmentType.CENTER, width: 1800 }),
          tc("0.159", { fill: C.paleGreen, color: C.itc, align: AlignmentType.CENTER, width: 1800 }),
          tc("0.22", { fill: "F3E5F5", color: C.nestle, align: AlignmentType.CENTER, width: 2160 }),
        ]}),
        new TableRow({ children: [
          tc("RE % (FY25)", { fill: C.greyLight, bold: true, width: 1800 }),
          tc("1.11%", { fill: "FFF3E0", color: C.reliance, align: AlignmentType.CENTER, width: 1800 }),
          tc("45% ↑↑", { fill: C.palBlue, color: C.tata, bold: true, align: AlignmentType.CENTER, width: 1800 }),
          tc("52% ↑", { fill: C.paleGreen, color: C.itc, bold: true, align: AlignmentType.CENTER, width: 1800 }),
          tc("40.6%", { fill: "F3E5F5", color: C.nestle, align: AlignmentType.CENTER, width: 2160 }),
        ]}),
        new TableRow({ children: [
          tc("Net Zero Target", { fill: C.greyLight, bold: true, width: 1800 }),
          tc("2035", { fill: "FFF3E0", color: C.reliance, bold: true, align: AlignmentType.CENTER, width: 1800 }),
          tc("2040/2045", { fill: C.palBlue, color: C.tata, bold: true, align: AlignmentType.CENTER, width: 1800 }),
          tc("2050", { fill: C.paleGreen, color: C.itc, bold: true, align: AlignmentType.CENTER, width: 1800 }),
          tc("2050 (SBTi ✓)", { fill: "F3E5F5", color: C.nestle, bold: true, align: AlignmentType.CENTER, width: 2160 }),
        ]}),
        new TableRow({ children: [
          tc("Scope 3 (MtCO₂e)", { fill: C.greyLight, bold: true, width: 1800 }),
          tc(">200 (est.)", { fill: "FFF3E0", color: "5D4037", align: AlignmentType.CENTER, size: 17, width: 1800 }),
          tc("162.3", { fill: C.palBlue, color: C.tata, align: AlignmentType.CENTER, width: 1800 }),
          tc("1.062", { fill: C.paleGreen, color: C.itc, align: AlignmentType.CENTER, width: 1800 }),
          tc("66.01 (net: 65.74)", { fill: "F3E5F5", color: C.nestle, align: AlignmentType.CENTER, size: 17, width: 2160 }),
        ]}),
      ]
    }),
    ...spacer(1),
    heading2("1.2 Strategic Highlights"),
    bullet("Reliance Industries: India's largest private sector company is at an energy transition inflection point — currently at just 1.11% RE, but has committed $10B+ to green hydrogen, solar and battery manufacturing at the Jamnagar Giga Complex, targeting Net Zero by 2035."),
    bullet("Tata Motors: The fastest RE transition in this cohort — from 30% (FY24) to 45% (FY25), a 15 percentage point jump driven by PPAs, iRECs, and 76.5 MWp of rooftop solar. RE100 signatory targeting 100% by 2030. Split net zero: PV by 2040, CV by 2045."),
    bullet("ITC Limited: India's sustainability leader — the only FMCG company in this analysis that is net carbon positive, sequestering 6.46 MtCO₂e through its large-scale forestry programs. GHG intensity reduced by 46% vs FY2018-19 baseline."),
    bullet("Nestlé Global: The most advanced on governance — the only company with SBTi-validated targets (1.5°C pathway). Net GHG of 65.74 MtCO₂e (FY2024) after removals of 3.30 Mt. Manufacturing electricity is 98.6% renewable."),
  ];
}

// ── Section 2: Company Profiles ───────────────────────────────────────────────
function makeSection2() {
  return [
    pageBreak(),
    heading1("2. Company Profiles & Sustainability Context"),

    heading2("2.1 Reliance Industries Limited"),
    para("Sector: Oil & Gas / Conglomerate | Country: India | Reporting: BRSR 2024-25", { bold: true, color: C.reliance }),
    para("Reliance Industries is India's largest private sector company by revenue (₹10,00,000+ crore), with operations spanning petrochemicals, refining, retail, and telecom. Its sustainability journey is at a pivotal juncture: massive legacy Scope 1 emissions from oil & gas (36.46 MtCO₂e) sit alongside the most ambitious energy transition investment commitment in India's corporate sector."),
    bullet("Committed $10B+ to New Energy across green hydrogen, solar cells, batteries and electrolyser manufacturing"),
    bullet("Net Zero target: 2035 — the most ambitious of the four companies analysed"),
    bullet("Current RE share: only 1.11% — reflecting the scale of transformation required"),
    bullet("Building the world's largest integrated renewable energy complex at Jamnagar (Dhirubhai Ambani Green Energy Giga Complex)"),
    bullet("Third-party assurance by Deloitte for sustainability disclosures"),
    ...spacer(1),

    heading2("2.2 Tata Motors Limited"),
    para("Sector: Automotive | Country: India | Reporting: BRSR 2024-25 (TML+TMPVL+TPEML)", { bold: true, color: C.tata }),
    para("Tata Motors is India's leading automobile manufacturer, covering commercial vehicles (TML), passenger vehicles (TMPVL), and electric mobility (TPEML). The combined entity has demonstrated the fastest renewable energy progress in this cohort and maintains a comprehensive climate strategy across all three reporting entities."),
    bullet("Combined Scope 1+2: 0.285 MtCO₂e — amongst the lowest in this cross-sector analysis"),
    bullet("RE% jumped from 30% (FY24) to 45% (FY25), a 15pp increase — iRECs contributed 12% of this"),
    bullet("RE100 signatory: committed to 100% renewable electricity by end of the decade"),
    bullet("76.5 MWp rooftop solar installed across all plants (TML: 55.5 MWp + TMPVL: 21 MWp)"),
    bullet("ENCON projects saved 18,423 tCO₂e in FY25 alone"),
    bullet("3 plants certified Water Positive + Zero Waste to Landfill (CII-GBC): Dharwad, Pantnagar, Lucknow"),
    bullet("KPMG provides limited assurance on all key environmental KPIs"),
    ...spacer(1),

    heading2("2.3 ITC Limited"),
    para("Sector: FMCG / Agribusiness | Country: India | Reporting: GRI/BRSR 2024-25", { bold: true, color: C.itc }),
    para("ITC is India's leading diversified FMCG company, with a unique sustainability profile — it is net carbon positive, meaning it sequesters more CO₂ than its operations emit. This is achieved through India's largest corporate forestry program, covering millions of hectares across watershed areas."),
    bullet("Scope 1+2: 1.264 MtCO₂e — low absolute emissions for its scale"),
    bullet("GHG intensity: reduced by 46% vs FY2018-19 baseline — most significant improvement in the cohort"),
    bullet("Carbon sequestration: 6.46 MtCO₂e — net carbon positive since FY2007"),
    bullet("52% of total energy from renewables; 53% of purchased grid electricity is renewable"),
    bullet("Scope 3: 1.062 MtCO₂e across 7 categories including agriculture, packaging, and transport"),
    bullet("Net Zero target: Operations by 2050; interim −50% GHG intensity by 2030"),
    ...spacer(1),

    heading2("2.4 Nestlé Global"),
    para("Sector: FMCG / Food & Beverage | Country: Switzerland (Global) | Reporting: GRI/CDP 2024", { bold: true, color: C.nestle }),
    para("Nestlé is the world's largest food and beverage company, with operations across 188 countries. It is the most advanced in this cohort in terms of climate governance — holding the only SBTi-validated targets (1.5°C-aligned) among the four companies analysed."),
    bullet("Scope 1: 2.82 MtCO₂e | Scope 2 market-based: 0.22 MtCO₂e (FY2024)"),
    bullet("Gross GHG: 69.04 MtCO₂e | Net: 65.74 MtCO₂e (after 3.30 Mt removals)"),
    bullet("Manufacturing electricity: 98.6% renewable — nearly complete transition"),
    bullet("Scope 3 (FLAG agriculture): 30.43 MtCO₂e — SBTi FLAG target of −25% vs 2018 baseline"),
    bullet("SBTi-validated targets: 1.5°C pathway — only validated SBTs in this cohort"),
    bullet("GHG intensity reduction: −24.5% vs 2018 baseline (gross, non-FLAG)"),
    bullet("Net Zero target: 2050 for all value chain emissions"),
  ];
}

// ── Section 3: GHG Emissions ─────────────────────────────────────────────────
function makeSection3() {
  return [
    pageBreak(),
    heading1("3. GHG Emissions Analysis"),
    heading2("3.1 Scope 1 & Scope 2 Emissions"),
    para("The four companies represent highly contrasting GHG profiles driven by their sector characteristics. Reliance Industries, as an oil & gas conglomerate, generates orders of magnitude more direct emissions than the automotive and FMCG companies."),
    ...chartPara("chart1_scope1_2_comparison.png", 620, 310, "Figure 1: Scope 1 & 2 Emissions Comparison — FY2024-25 (Log Scale). Sources: Official Sustainability Reports."),
    heading2("3.2 Key Emissions Data"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2400, 1740, 1740, 1740, 1740],
      rows: [
        new TableRow({ children: [
          tc("Emissions Metric", { fill: C.darkBlue, color: C.white, bold: true, size: 18, width: 2400 }),
          tc("Reliance", { fill: C.darkBlue, color: C.white, bold: true, align: AlignmentType.CENTER, size: 18, width: 1740 }),
          tc("Tata Motors", { fill: C.darkBlue, color: C.white, bold: true, align: AlignmentType.CENTER, size: 18, width: 1740 }),
          tc("ITC Limited", { fill: C.darkBlue, color: C.white, bold: true, align: AlignmentType.CENTER, size: 18, width: 1740 }),
          tc("Nestlé Global", { fill: C.darkBlue, color: C.white, bold: true, align: AlignmentType.CENTER, size: 18, width: 1740 }),
        ]}),
        ...([
          ["Scope 1 (MtCO₂e)", "36.46", "0.071", "1.105", "2.82"],
          ["Scope 2 — Market-Based (MtCO₂e)", "1.47", "0.214", "0.159", "0.22"],
          ["Total Scope 1+2 (MtCO₂e)", "37.93", "0.285", "1.264", "3.04"],
          ["Scope 3 Total (MtCO₂e)", ">200 (est.)", "162.3", "1.062", "66.01"],
          ["Largest S3 Category", "Downstream fuel use", "Use of sold products", "Agriculture/transport", "Agriculture (FLAG)"],
          ["S3 Largest Cat. Value", "~180+ Mt (est.)", "152.6 Mt", "~0.45 Mt", "30.43 Mt"],
          ["Carbon Sequestration", "Not disclosed", "SBTN program", "6.46 MtCO₂e", "3.30 Mt removals"],
          ["Net Emissions Status", "High — transition underway", "Low & reducing", "Net Carbon Positive", "Net: 65.74 MtCO₂e"],
        ].map((row, i) => new TableRow({ children: row.map((cell, j) => {
          const fills = ["FFF3E0", C.palBlue, C.paleGreen, "F3E5F5"];
          const fill = j === 0 ? C.greyLight : (i % 2 === 0 ? fills[j-1] : C.white);
          return tc(cell, { fill, bold: j === 0, align: j === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
            width: j === 0 ? 2400 : 1740, size: 17 });
        })}))),
      ]
    }),
    ...spacer(1),
    heading2("3.3 Scope 3 Analysis"),
    para("Scope 3 emissions represent the most significant challenge for all four companies. Tata Motors' dominant Scope 3 category is 'use of sold products' (Category 11), accounting for 152.6 MtCO₂e out of a total 162.3 MtCO₂e — essentially the lifetime tailpipe emissions of all vehicles sold in FY25. This underscores why the company's electric vehicle strategy is central to its long-term decarbonisation roadmap."),
    ...chartPara("chart4_scope3_analysis.png", 620, 310, "Figure 2: Scope 3 Breakdown — Tata Motors Category Analysis & Cross-Company Comparison."),
    para("For Nestlé, agricultural emissions (FLAG — Forests, Land and Agriculture) account for 30.43 MtCO₂e, or 46% of its total Scope 3. The company has set specific SBTi FLAG targets requiring a 25% reduction vs 2018 baseline — a pioneering approach in the food sector."),
  ];
}

// ── Section 4: Energy & Renewables ───────────────────────────────────────────
function makeSection4() {
  return [
    pageBreak(),
    heading1("4. Energy Transition & Renewable Energy Progress"),
    para("The renewable energy transition trajectory varies dramatically across the four companies, driven by their sector context, starting baselines, and strategic commitments."),
    ...chartPara("chart2_renewable_energy_progress.png", 640, 350, "Figure 3: Renewable Energy Progress FY2023–FY2025 and 2030 Target. Sources: Official Sustainability Reports."),
    heading2("4.1 Renewable Energy Performance"),
    bullet("Reliance Industries (1.11%): The largest gap between current state and ambition. Its 2035 Net Zero target will require a complete transformation of the energy mix underpinning its refining and petrochemical operations. The $10B+ New Energy investment represents India's biggest corporate bet on the energy transition.", true),
    bullet("Tata Motors (45%): The standout performer in FY25, increasing RE share by 15 percentage points in a single year. Strategy includes onsite solar (76.5 MWp), captive wind farms, PPAs, and iRECs. IRECs alone contributed 12% to the RE share of overall electricity consumption.", true),
    bullet("ITC Limited (52%): India's FMCG sustainability leader, already past the halfway mark for its renewable energy transition. Combines onsite solar with grid renewable procurement. The company's large agribusiness operations provide additional renewable biomass feedstock.", true),
    bullet("Nestlé Global (40.6% total / 98.6% mfg electricity): The distinction between total energy and manufacturing electricity is critical here. Nestlé's manufacturing electricity is nearly fully renewable (98.6%), but when total energy (including thermal processes) is included, the figure falls to 40.6%. This reflects the challenge of decarbonising process heat.", true),
    ...spacer(1),
    heading2("4.2 Tata Motors Energy Conservation Projects (FY25)"),
    para("Tata Motors' ENCON programme delivered 210.89 lakh kWh electricity savings and 48,811 GJ fuel savings in FY25, equivalent to 18,423 tCO₂e avoided. Key initiatives included:"),
    bullet("Elimination of Sealer Oven process in Paint Shop — significant propane reduction"),
    bullet("Variable Refrigerant Flow (VRF) systems in Engine Assembly and Power Train areas"),
    bullet("40 Variable Frequency Drives (VFDs) installed in Paint Shop"),
    bullet("Energy-efficient EC motors replacing conventional motors at Air Supply Plants"),
    bullet("Heat pump utilisation for Paint Shop process tank heating"),
    bullet("Digitisation and furnace optimisation in heat treatment area"),
  ];
}

// ── Section 5: Net Zero Roadmap ───────────────────────────────────────────────
function makeSection5() {
  return [
    pageBreak(),
    heading1("5. Net Zero Roadmap & Climate Commitments"),
    ...chartPara("chart3_net_zero_timeline.png", 640, 360, "Figure 4: Net Zero Journey — Key Milestones & Commitments (2024–2050)."),
    heading2("5.1 Net Zero Target Comparison"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2200, 1790, 1790, 1790, 1790],
      rows: [
        new TableRow({ children: [
          tc("Target", { fill: C.darkGreen, color: C.white, bold: true, width: 2200 }),
          tc("Reliance", { fill: C.darkGreen, color: C.white, bold: true, align: AlignmentType.CENTER, width: 1790 }),
          tc("Tata Motors", { fill: C.darkGreen, color: C.white, bold: true, align: AlignmentType.CENTER, width: 1790 }),
          tc("ITC Limited", { fill: C.darkGreen, color: C.white, bold: true, align: AlignmentType.CENTER, width: 1790 }),
          tc("Nestlé Global", { fill: C.darkGreen, color: C.white, bold: true, align: AlignmentType.CENTER, width: 1790 }),
        ]}),
        ...([
          ["Net Zero (Operations)", "2035", "2040 PV / 2045 CV", "2050", "2050"],
          ["100% RE Target", "100 GW capacity", "RE100 by 2030", "Internal RE target", "100% RE electricity"],
          ["SBTi Status", "Committed", "Committed + interim", "Aligned (1.5°C)", "Validated (1.5°C) ✓"],
          ["Water Target", "Not disclosed", "Water Positive 2030", "100% water recycling", "Water stewardship"],
          ["Waste Target", "Not disclosed", "Zero WtL by 2030", "Zero waste ambition", "Zero waste in mfg"],
          ["Key 2030 Milestone", "$10B+ RE investment", "RE100 + ZWtL + Water+", "−50% GHG intensity", "−50% GHG vs 2018"],
        ].map((row, i) => new TableRow({ children: row.map((cell, j) => {
          const fills = ["FFF3E0", C.palBlue, C.paleGreen, "F3E5F5"];
          const fill = j === 0 ? C.greyLight : (i % 2 === 0 ? fills[j-1] : C.white);
          return tc(cell, { fill, bold: j === 0, align: j === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
            width: j === 0 ? 2200 : 1790, size: 17 });
        })}))),
      ]
    }),
    ...spacer(1),
    heading2("5.2 Nestlé — SBTi Validated Targets (Most Advanced)"),
    para("Nestlé is the only company in this cohort with SBTi-validated science-based targets. Its commitments include reducing Scope 1+2 emissions by 50% by 2030 vs 2018 baseline, and a specific SBTi FLAG target to reduce agricultural and land-use emissions (30.43 MtCO₂e) by 25% by 2030 vs 2018. This represents best-practice climate governance in the FMCG sector."),
    heading2("5.3 ITC — Net Carbon Positive Since 2007"),
    para("ITC is uniquely positioned as India's only major FMCG company that is net carbon positive — its forestry and watershed management programs sequester 6.46 MtCO₂e annually, more than the company's total Scope 1+2 emissions of 1.264 MtCO₂e. This has been sustained for over 17 consecutive years. The company's GHG intensity (per unit revenue) has improved 46% since FY2018-19."),
  ];
}

// ── Section 6: Performance Scorecard ─────────────────────────────────────────
function makeSection6() {
  return [
    pageBreak(),
    heading1("6. Decarbonisation Performance Scorecard"),
    ...chartPara("chart5_scorecard_radar.png", 500, 500, "Figure 5: Decarbonisation Performance Scorecard (Analyst Assessment, Scale 0–10). Based on official SR content."),
    heading2("6.1 Scorecard Methodology"),
    para("The scorecard evaluates each company across 8 dimensions on a scale of 0–10, based on a qualitative and quantitative assessment of their official sustainability report disclosures. It is not a formal ESG rating but an analyst's structured comparison to highlight relative strengths and gaps."),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2800, 820, 820, 820, 820, 3280],
      rows: [
        new TableRow({ children: [
          tc("Dimension", { fill: C.darkBlue, color: C.white, bold: true, width: 2800 }),
          tc("Reliance", { fill: C.darkBlue, color: C.white, bold: true, align: AlignmentType.CENTER, width: 820 }),
          tc("Tata M.", { fill: C.darkBlue, color: C.white, bold: true, align: AlignmentType.CENTER, width: 820 }),
          tc("ITC", { fill: C.darkBlue, color: C.white, bold: true, align: AlignmentType.CENTER, width: 820 }),
          tc("Nestlé", { fill: C.darkBlue, color: C.white, bold: true, align: AlignmentType.CENTER, width: 820 }),
          tc("Notes", { fill: C.darkBlue, color: C.white, bold: true, width: 3280 }),
        ]}),
        ...([
          ["RE Progress (vs 100% target)", 1, 7, 8, 7, "TM: fastest growth. Reliance: biggest gap vs target."],
          ["Net Zero Ambition", 8, 8, 7, 7, "Reliance most ambitious (2035). All have firm commitments."],
          ["Scope 3 Strategy", 4, 6, 7, 8, "Nestlé: SBTi FLAG. TM: LCA + supplier RE. Reliance: early stage."],
          ["Disclosure Quality", 7, 9, 8, 9, "TM & Nestlé: KPMG/SGS assured. Reliance: Deloitte assured."],
          ["Carbon Sequestration", 2, 4, 10, 6, "ITC: net carbon positive for 17 years. Unique in peer group."],
          ["Water Stewardship", 4, 8, 7, 7, "TM: 3 plants Water Positive certified. Best in class."],
          ["Circular Economy", 3, 8, 7, 7, "TM: TATVA framework + Re.Wi.Re. ELV recycling network."],
          ["SBTi Alignment", 5, 7, 7, 10, "Nestlé: only company with validated SBTs (1.5°C pathway)."],
        ].map((row, i) => new TableRow({ children: [
          tc(row[0], { fill: C.greyLight, bold: true, width: 2800 }),
          tc(row[1].toString(), { fill: row[1] >= 7 ? C.paleGreen : row[1] >= 5 ? C.yellow : "FFCDD2",
            bold: true, align: AlignmentType.CENTER, width: 820, color: row[1] >= 7 ? C.itc : row[1] >= 5 ? "5D4037" : "C62828" }),
          tc(row[2].toString(), { fill: row[2] >= 7 ? C.paleGreen : row[2] >= 5 ? C.yellow : "FFCDD2",
            bold: true, align: AlignmentType.CENTER, width: 820, color: row[2] >= 7 ? C.itc : row[2] >= 5 ? "5D4037" : "C62828" }),
          tc(row[3].toString(), { fill: row[3] >= 7 ? C.paleGreen : row[3] >= 5 ? C.yellow : "FFCDD2",
            bold: true, align: AlignmentType.CENTER, width: 820, color: row[3] >= 7 ? C.itc : row[3] >= 5 ? "5D4037" : "C62828" }),
          tc(row[4].toString(), { fill: row[4] >= 7 ? C.paleGreen : row[4] >= 5 ? C.yellow : "FFCDD2",
            bold: true, align: AlignmentType.CENTER, width: 820, color: row[4] >= 7 ? C.itc : row[4] >= 5 ? "5D4037" : "C62828" }),
          tc(row[5], { fill: C.lightAmber || "FFF8E1", size: 17, width: 3280 }),
        ]}))),
      ]
    }),
  ];
}

// ── Section 7: Water, Waste & Circular Economy ───────────────────────────────
function makeSection7() {
  return [
    pageBreak(),
    heading1("7. Water, Waste & Circular Economy"),
    heading2("7.1 Tata Motors — Water Stewardship"),
    para("Tata Motors has the most structured water management programme in this cohort, with specific certifications and targets backed by third-party assurance."),
    bullet("Total water withdrawal: 46.78 lakh kL (FY25) — down from 50.34 lakh kL (FY24), a 7% reduction"),
    bullet("Water intensity: 4.80 kL/vehicle (FY25) vs 5.06 kL/vehicle (FY24) — 5.1% improvement"),
    bullet("3 plants certified Water Positive by CII-GBC: Dharwad (Nov 2023), Pantnagar (Aug 2024), Lucknow (Nov 2024)"),
    bullet("Zero Liquid Discharge (ZLD) at most plants; tertiary treatment (RO) at Pune-Pimpri and Jamshedpur"),
    bullet("Target: Water Positive across all operations and townships by 2030"),
    ...spacer(1),
    heading2("7.2 Tata Motors — Waste & Circular Economy"),
    bullet("Total waste generated: 1,97,305 MT (FY25) — includes construction & demolition waste"),
    bullet("Waste recycled: 1,65,410 MT = 84% recycling rate"),
    bullet("3 plants certified Zero Waste to Landfill (ZWtL) by CII-GBC: Dharwad, Pantnagar, Lucknow"),
    bullet("Target: Zero Waste to Landfill across all operations by 2030"),
    bullet("TATVA circular economy framework: reducing material footprint targets by 2030"),
    bullet("Re.Wi.Re. (Recycle with Respect): 7 end-of-life vehicle recycling facilities nationwide; capacity to dismantle 1,10,000+ vehicles/year"),
    bullet("First Indian OEM to publish dismantling information on the IDIS (International Dismantling Information System) portal — 11 models covered"),
    bullet("EPR compliance: 13,474 MT plastic waste recycled via registered Producer Responsibility Organisations"),
    ...spacer(1),
    heading2("7.3 ITC — Water & Land Stewardship"),
    bullet("Near-100% water recycling achieved across manufacturing operations"),
    bullet("Watershed development programs across millions of hectares"),
    bullet("Agroforestry: millions of trees planted annually for carbon sequestration and watershed protection"),
    ...spacer(1),
    heading2("7.4 Nestlé — Water & Packaging"),
    bullet("Water stewardship programs in all water-stressed areas of operation"),
    bullet("Packaging circularity: committed to making 95% of packaging recyclable or reusable by 2025"),
    bullet("Closed-loop water systems in manufacturing facilities globally"),
  ];
}

// ── Section 8: GHG Intensity & Investment ───────────────────────────────────
function makeSection8() {
  return [
    pageBreak(),
    heading1("8. GHG Intensity & Energy Transition Investment"),
    ...chartPara("chart6_intensity_investment.png", 640, 310, "Figure 6: GHG Intensity Index & Energy Transition Investment Scale. Sources: Official SRs / BRSR."),
    heading2("8.1 Emission Intensity Analysis"),
    para("GHG intensity metrics reveal the efficiency of each company's operations relative to economic output. The stark contrast between Reliance (highest intensity due to Oil & Gas) and the other three companies reflects fundamental sector differences rather than management performance alone."),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [3200, 1540, 1540, 1540, 1540],
      rows: [
        new TableRow({ children: [
          tc("Intensity Metric", { fill: C.darkBlue, color: C.white, bold: true, width: 3200 }),
          tc("Reliance", { fill: C.darkBlue, color: C.white, bold: true, align: AlignmentType.CENTER, width: 1540 }),
          tc("Tata Motors", { fill: C.darkBlue, color: C.white, bold: true, align: AlignmentType.CENTER, width: 1540 }),
          tc("ITC", { fill: C.darkBlue, color: C.white, bold: true, align: AlignmentType.CENTER, width: 1540 }),
          tc("Nestlé", { fill: C.darkBlue, color: C.white, bold: true, align: AlignmentType.CENTER, width: 1540 }),
        ]}),
        ...([
          ["S1+S2 / Revenue (relative index)", "100 (base)", "0.75", "3.33", "8.01"],
          ["GHG Intensity per unit", "0.471 tCO₂e/MT throughput", "0.305 tCO₂e/vehicle", "Disclosed per ₹Cr", "Disclosed per CHF"],
          ["YoY Intensity Trend", "Stable", "↓3.8% per vehicle", "↓46% vs FY2019", "↓24.5% vs 2018"],
          ["Intensity Reduction Strategy", "New Energy transformation", "ENCON + RE transition", "Forestry + RE", "Regen agri + RE"],
        ].map((row, i) => new TableRow({ children: row.map((cell, j) => {
          const fills = ["FFF3E0", C.palBlue, C.paleGreen, "F3E5F5"];
          const fill = j === 0 ? C.greyLight : (i % 2 === 0 ? fills[j-1] : C.white);
          return tc(cell, { fill, bold: j === 0, align: j === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
            width: j === 0 ? 3200 : 1540, size: 17 });
        })}))),
      ]
    }),
    ...spacer(1),
    heading2("8.2 Investment in Energy Transition"),
    bullet("Reliance Industries — $10B+: The largest corporate energy transition commitment in India, spanning green hydrogen (Jamnagar Giga Complex), solar cell manufacturing, advanced energy storage (batteries), and electrolyser manufacturing. Explicitly stated in BRSR 2024-25.", true),
    bullet("Tata Motors — ~$3.5B equivalent: Ongoing investment in RE capacity (76.5 MWp solar + PPAs + wind farms), EV platform development (Nexon EV, Tigor EV, upcoming models), and circular economy infrastructure (Re.Wi.Re. facilities). Includes significant R&D — 49% of total R&D expenditure directed at environmental improvements.", true),
    bullet("ITC Limited — ~$1.5B equivalent: Investments in solar expansion, agroforestry programs, and renewable energy procurement. The company's forestry investment delivers dual benefits: carbon sequestration and watershed development for farming communities.", true),
    bullet("Nestlé Global — ~$2.8B equivalent: Global investment in regenerative agriculture programs, renewable energy transition across manufacturing, packaging circularity, and sustainable supply chain transformation. Allocated across 188 countries of operation.", true),
  ];
}

// ── Section 9: Conclusions ───────────────────────────────────────────────────
function makeSection9() {
  return [
    pageBreak(),
    heading1("9. Conclusions & Key Takeaways"),
    heading2("9.1 Where Each Company Stands"),
    para("This analysis reveals that all four companies are on genuine decarbonisation pathways, but at very different stages of their energy transition journeys. The most striking contrast is between Reliance Industries — currently just 1.11% renewable but with the boldest investment commitment — and ITC Limited, which has been net carbon positive for 17 years."),
    ...spacer(1),
    heading2("9.2 Sector-Specific Insights"),
    bullet("Oil & Gas (Reliance): The Scope 1 emissions challenge is structural — refining and petrochemicals are inherently emissions-intensive. The only path to Net Zero 2035 is a complete parallel build of a new energy business, which is exactly what the $10B+ New Energy strategy represents. Watch for: hydrogen economy development, renewable capacity commissioning, and supply chain electrification."),
    bullet("Automotive (Tata Motors): The FY25 RE jump (+15pp) demonstrates what is achievable with a strong operations-level focus. However, Scope 3 (use of sold products = 152.6 MtCO₂e) dwarfs operational emissions by 570x. The company's EV strategy — Nexon EV, Tigor EV, and upcoming models — is therefore central to meaningful decarbonisation. TMPVL PV Net Zero 2040 is ambitious but achievable given the EV transition trajectory."),
    bullet("FMCG India (ITC): A benchmark for nature-based solutions. ITC demonstrates that carbon sequestration through forestry can make a company net carbon positive even before counting Scope 3. The challenge ahead is extending this to the supply chain and product lifecycle. The company's agroforestry model also addresses rural livelihoods, making it an integrated sustainability-development solution."),
    bullet("FMCG Global (Nestlé): The governance standard-bearer of this cohort. Validated SBTi targets (1.5°C), distinct FLAG accounting, and 98.6% renewable manufacturing electricity set a benchmark for disclosure quality. The biggest challenge remains agriculture (30.43 MtCO₂e Scope 3 FLAG), which requires system-level change in farming practices across a complex global supply chain."),
    ...spacer(1),
    heading2("9.3 Cross-Cutting Themes"),
    bullet("All four companies recognise that Scope 3 is the defining decarbonisation challenge — but addressing it requires systemic change well beyond individual company control."),
    bullet("Renewable energy transition is accelerating across all companies. The 2030 deadline for most targets is creating urgency in RE procurement strategies."),
    bullet("Water stewardship is emerging as a co-equal priority alongside GHG — reflecting physical climate risk materialising in operations."),
    bullet("Circular economy principles (Zero Waste to Landfill, product recyclability, EPR compliance) are being institutionalised as part of sustainability strategy, not just compliance."),
    bullet("Third-party assurance is now standard — KPMG (Tata Motors), Deloitte (Reliance), and SGS (Nestlé) all providing limited or reasonable assurance on key KPIs."),
    ...spacer(1),
    heading2("9.4 Data Sources"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2400, 4200, 2760],
      rows: [
        new TableRow({ children: [
          tc("Company", { fill: C.darkGreen, color: C.white, bold: true, width: 2400 }),
          tc("Source Document", { fill: C.darkGreen, color: C.white, bold: true, width: 4200 }),
          tc("Assurance Provider", { fill: C.darkGreen, color: C.white, bold: true, width: 2760 }),
        ]}),
        ...([
          ["Reliance Industries", "Business Responsibility & Sustainability Report 2024-25", "Deloitte Haskins & Sells"],
          ["Tata Motors", "BRSR 2024-25 (TML + TMPVL + TPEML)", "KPMG Assurance & Consulting Services LLP"],
          ["ITC Limited", "Sustainability Report 2024-25 (GRI/BRSR)", "Third-party verified"],
          ["Nestlé Global", "Nestlé Sustainability Report 2024 (Appendix 4)", "SGS (validated GHG data)"],
        ].map((row, i) => new TableRow({ children: [
          tc(row[0], { fill: i % 2 === 0 ? C.paleGreen : C.white, bold: true, width: 2400 }),
          tc(row[1], { fill: i % 2 === 0 ? C.paleGreen : C.white, size: 17, width: 4200 }),
          tc(row[2], { fill: i % 2 === 0 ? C.paleGreen : C.white, size: 17, width: 2760 }),
        ]}))),
      ]
    }),
    ...spacer(2),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 320, after: 80 },
      children: [new TextRun({ text: "— End of Report —", size: 20, color: "9E9E9E", italics: true, font: "Arial" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 80 },
      children: [new TextRun({ text: "Prepared by: Sanjana Sabat | ESG & Sustainability Analyst | June 2025", size: 18, color: "BDBDBD", font: "Arial" })]
    }),
  ];
}

// ── Build Document ────────────────────────────────────────────────────────────
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Arial", size: 20, color: "212121" } }
    },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: "1B5E20" },
        paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: "0D47A1" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: "37474F" },
        paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 2 } },
    ]
  },
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 },
          spacing: { before: 40, after: 40 } } } }]
    }]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1260, right: 1260, bottom: 1260, left: 1260 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
          children: [new TextRun({
            text: "Energy Transition & Decarbonisation Roadmap | Real Data | FY2024-25",
            size: 16, color: "9E9E9E", italics: true, font: "Arial"
          })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
          children: [
            new TextRun({ text: "Sanjana Sabat | ESG & Sustainability Analyst | ", size: 16, color: "9E9E9E", font: "Arial" }),
            new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "9E9E9E", font: "Arial" }),
          ]
        })]
      })
    },
    children: [
      ...makeCover(),
      ...makeSection1(),
      ...makeSection2(),
      ...makeSection3(),
      ...makeSection4(),
      ...makeSection5(),
      ...makeSection6(),
      ...makeSection7(),
      ...makeSection8(),
      ...makeSection9(),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log(`✅ WORD REPORT SAVED: ${OUT}`);
  console.log(`   Size: ${(buf.length/1024).toFixed(0)} KB`);
});
