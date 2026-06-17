const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, LevelFormat, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, Header, Footer,
  TabStopType, TabStopPosition
} = require("docx");
const fs = require("fs");

// ── Palette ──────────────────────────────────────────────────────────────────
const NAVY   = "0D1B2A";
const TEAL   = "1B6CA8";
const WHITE  = "FFFFFF";
const LGRAY  = "F4F6F8";
const MGRAY  = "E5E9EE";
const G_BG   = "D1FAE5"; const G_FG = "065F46";
const B_BG   = "DBEAFE"; const B_FG = "1E3A8A";
const R_BG   = "FEE2E2"; const R_FG = "7F1D1D";
const A_BG   = "FEF9C3"; const A_FG = "78350F";

const PAGE_W = 12240; // US Letter
const MARGIN = 1080;  // 0.75 inch
const CONTENT_W = PAGE_W - MARGIN * 2; // 10080 DXA

// ── Helpers ───────────────────────────────────────────────────────────────────
const run = (text, opts = {}) =>
  new TextRun({ text, font: "Calibri", size: opts.sz || 20, bold: opts.bold, italic: opts.italic, color: opts.color || "111827", ...opts });

const para = (children, opts = {}) =>
  new Paragraph({
    children: Array.isArray(children) ? children : [children],
    spacing: { before: opts.before || 80, after: opts.after || 80, line: opts.line || 276 },
    alignment: opts.align || AlignmentType.LEFT,
    ...opts
  });

const h1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  children: [new TextRun({ text, font: "Calibri", size: 32, bold: true, color: NAVY })],
  spacing: { before: 320, after: 120 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: TEAL, space: 4 } }
});

const h2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  children: [new TextRun({ text, font: "Calibri", size: 26, bold: true, color: TEAL })],
  spacing: { before: 240, after: 80 }
});

const h3 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  children: [new TextRun({ text, font: "Calibri", size: 22, bold: true, color: "374151" })],
  spacing: { before: 180, after: 60 }
});

const bullet = (text, indent = 0) => new Paragraph({
  numbering: { reference: "bullets", level: indent },
  children: [run(text)],
  spacing: { before: 40, after: 40, line: 276 }
});

const numbered = (text) => new Paragraph({
  numbering: { reference: "numbers", level: 0 },
  children: [run(text)],
  spacing: { before: 60, after: 60, line: 276 }
});

const spacer = (sz = 80) => para([run("")], { before: sz, after: 0 });

const border = { style: BorderStyle.SINGLE, size: 4, color: MGRAY };
const borders = { top: border, bottom: border, left: border, right: border };

const cell = (text, opts = {}) => new TableCell({
  borders,
  width: { size: opts.w || 2520, type: WidthType.DXA },
  shading: { fill: opts.bg || WHITE, type: ShadingType.CLEAR },
  margins: { top: 80, bottom: 80, left: 140, right: 140 },
  verticalAlign: VerticalAlign.CENTER,
  children: [new Paragraph({
    children: [new TextRun({
      text, font: "Calibri", size: opts.sz || 18,
      bold: opts.bold || false, color: opts.fg || "111827"
    })],
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: 20, after: 20 }
  })]
});

const headerRow = (cells) => new TableRow({
  tableHeader: true,
  children: cells.map(([text, w]) => cell(text, { bg: NAVY, fg: WHITE, bold: true, w, sz: 18 }))
});

// ── Cover block (styled paragraph as banner) ──────────────────────────────────
const coverTitle = new Paragraph({
  children: [
    new TextRun({ text: "Research Methodology", font: "Calibri", size: 56, bold: true, color: WHITE })
  ],
  shading: { fill: NAVY, type: ShadingType.CLEAR },
  alignment: AlignmentType.LEFT,
  spacing: { before: 0, after: 0 },
  indent: { left: 240, right: 240 },
  border: {
    top: { style: BorderStyle.NONE },
    bottom: { style: BorderStyle.NONE },
    left: { style: BorderStyle.NONE },
    right: { style: BorderStyle.NONE }
  }
});

const coverSub = new Paragraph({
  children: [new TextRun({ text: "DeepThought Business Analytics Internship  |  Mumbai  |  Specialty Chemicals & Performance Chemicals", font: "Calibri", size: 20, color: WHITE })],
  shading: { fill: TEAL, type: ShadingType.CLEAR },
  alignment: AlignmentType.LEFT,
  spacing: { before: 0, after: 0 },
  indent: { left: 240, right: 240 }
});

const coverMeta = new Paragraph({
  children: [new TextRun({ text: "June 2026  |  Target: 25 Federer Companies  |  Researched: ~90 candidates", font: "Calibri", size: 18, italic: true, color: "4B5563" })],
  alignment: AlignmentType.LEFT,
  spacing: { before: 120, after: 280 },
  indent: { left: 240 }
});

// ── Source table ──────────────────────────────────────────────────────────────
const sourceTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [2200, 3200, 4680],
  rows: [
    headerRow([["Source", 2200], ["What it gives", 3200], ["Best used for", 4680]]),
    new TableRow({ children: [
      cell("Screener.in", { w: 2200, bg: LGRAY, bold: true }),
      cell("Revenue, EBITDA, promoter holding, growth trends", { w: 3200, bg: LGRAY }),
      cell("First revenue filter — quickly remove companies above Rs.500Cr or below Rs.30Cr", { w: 4680, bg: LGRAY })
    ]}),
    new TableRow({ children: [
      cell("MCA / RoC filings", { w: 2200, bold: true }),
      cell("Director names, appointment dates, shareholding, charge documents", { w: 3200 }),
      cell("Verifying gen-2 board appointments, PE ownership, group subsidiary status", { w: 4680 })
    ]}),
    new TableRow({ children: [
      cell("LinkedIn", { w: 2200, bg: LGRAY, bold: true }),
      cell("Founder / MD profiles, job postings, team size", { w: 3200, bg: LGRAY }),
      cell("C4 (DM background), C6 Hiring signal, C7 (IT/ERP roles)", { w: 4680, bg: LGRAY })
    ]}),
    new TableRow({ children: [
      cell("Company website", { w: 2200, bold: true }),
      cell("Products, certifications, news/press, copyright date", { w: 3200 }),
      cell("E1 gate (producer vs. trader), C3 (IP/certs), C6 activity signal", { w: 4680 })
    ]}),
    new TableRow({ children: [
      cell("USFDA / EU-GMP databases", { w: 2200, bg: LGRAY, bold: true }),
      cell("Approved manufacturing site lists", { w: 3200, bg: LGRAY }),
      cell("C3 differentiation — confirms regulated export capability", { w: 4680, bg: LGRAY })
    ]}),
    new TableRow({ children: [
      cell("Tofler / Zaubacorp", { w: 2200, bold: true }),
      cell("MCA-derived financials, director histories", { w: 3200 }),
      cell("Cross-checking revenue bands, promoter changes, PE acquisitions", { w: 4680 })
    ]}),
    new TableRow({ children: [
      cell("Google News", { w: 2200, bg: LGRAY, bold: true }),
      cell("Press releases, expansion news, certifications", { w: 3200, bg: LGRAY }),
      cell("C6 growth signals — new plant, new cert, new market entry", { w: 4680, bg: LGRAY })
    ]}),
    new TableRow({ children: [
      cell("Naukri / Indeed", { w: 2200, bold: true }),
      cell("Open job postings with dates", { w: 3200 }),
      cell("C6 Hiring signal — 5+ open roles in last 6 months", { w: 4680 })
    ]}),
    new TableRow({ children: [
      cell("BSE / NSE filings", { w: 2200, bg: LGRAY, bold: true }),
      cell("Annual reports, shareholding pattern, board changes", { w: 3200, bg: LGRAY }),
      cell("C7 (systems maturity references in ARs), C8 (board appointments)", { w: 4680, bg: LGRAY })
    ]})
  ]
});

// ── Filter funnel table ────────────────────────────────────────────────────────
const funnelTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [1400, 3480, 5200],
  rows: [
    headerRow([["Stage", 1400], ["Filter applied", 3480], ["Companies dropped", 5200]]),
    new TableRow({ children: [
      cell("1", { w: 1400, bg: LGRAY, bold: true, align: AlignmentType.CENTER }),
      cell("Initial discovery — ~90 candidates found", { w: 3480, bg: LGRAY }),
      cell("—", { w: 5200, bg: LGRAY })
    ]}),
    new TableRow({ children: [
      cell("2", { w: 1400, bold: true, align: AlignmentType.CENTER }),
      cell("Revenue check via Screener.in", { w: 3480 }),
      cell("~18 dropped: Sun Pharma, Wockhardt, Godrej & Boyce (too large); several micro-caps below Rs.30Cr", { w: 5200 })
    ]}),
    new TableRow({ children: [
      cell("3", { w: 1400, bg: LGRAY, bold: true, align: AlignmentType.CENTER }),
      cell("E1 gate — producer check", { w: 3480, bg: LGRAY }),
      cell("~12 dropped: CROs, testing labs, traders, distributors, pure consulting firms", { w: 5200, bg: LGRAY })
    ]}),
    new TableRow({ children: [
      cell("4", { w: 1400, bold: true, align: AlignmentType.CENTER }),
      cell("E2 gate — India operational presence", { w: 3480 }),
      cell("~3 dropped: companies with India holding office but all real ops overseas", { w: 4680 })
    ]}),
    new TableRow({ children: [
      cell("5", { w: 1400, bg: LGRAY, bold: true, align: AlignmentType.CENTER }),
      cell("Ownership check (PE / group subsidiary)", { w: 3480, bg: LGRAY }),
      cell("~6 dropped: Suven Pharma (Advent PE), Heubach flagged for verification, 4 others group-owned", { w: 5200, bg: LGRAY })
    ]}),
    new TableRow({ children: [
      cell("6", { w: 1400, bold: true, align: AlignmentType.CENTER }),
      cell("Full Federer scoring (C3–C8)", { w: 3480 }),
      cell("~16 scored below 60 — documented in fail list with reasons", { w: 5200 })
    ]}),
    new TableRow({ children: [
      cell("7  ✓", { w: 1400, bg: G_BG, fg: G_FG, bold: true, align: AlignmentType.CENTER }),
      cell("Final: 25 companies scoring 60+", { w: 3480, bg: G_BG }),
      cell("Included in CSV. All have specific evidence for each criterion — no inferences.", { w: 5200, bg: G_BG })
    ]})
  ]
});

// ── Score rubric table ─────────────────────────────────────────────────────────
const rubricTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [1200, 600, 2760, 2760, 2760],
  rows: [
    headerRow([["Criterion", 1200], ["Wt.", 600], ["Strong (full)", 2760], ["Moderate (half)", 2760], ["Weak (0)", 2760]]),
    ...[
      ["C3 Differentiated", "20", "Patents / DSIR / USFDA / EU-GMP; proprietary products; specialized equipment", "Some tech depth but no formal IP or approvals", "Commodity product; no moat"],
      ["C4 DM Quality", "15", "PhD / IIT / IISc founder OR operator who built ERP / costing / formal planning", "Gen-2 with formal education OR some structured ops evidence", "Non-technical promoter; outsource mindset"],
      ["C5 Growing Sector", "15", "PLI eligible; China+1; Make-in-India; export tailwinds; govt push", "Stable sector; no specific tailwinds", "Declining or stagnant sector"],
      ["C6 Growth Signals", "15", "2+ of: Hiring / New facility / New cert / Active website / Revenue growth", "1 signal from the list", "No visible activity in 2+ yrs"],
      ["C7 Systems Maturity", "20", "SAP / ERP confirmed; structured costing; MIS dashboards in use", "Some IT evidence but depth unclear", "No ERP; founder intuition only"],
      ["C8 Succession", "15", "Gen-2 on board (operational role) + external professional managers", "Gen-2 exists but unclear; OR 1 professional hire", "Solo founder; no gen-2; no succession"],
    ].map(([cr, wt, st, md, wk], i) =>
      new TableRow({ children: [
        cell(cr, { w: 1200, bg: i%2===0 ? LGRAY : WHITE, bold: true }),
        cell(wt, { w: 600, bg: i%2===0 ? LGRAY : WHITE, align: AlignmentType.CENTER }),
        cell(st, { w: 2760, bg: i%2===0 ? LGRAY : WHITE }),
        cell(md, { w: 2760, bg: i%2===0 ? LGRAY : WHITE }),
        cell(wk, { w: 2760, bg: i%2===0 ? LGRAY : WHITE }),
      ]})
    )
  ]
});

// ── Key learnings table ────────────────────────────────────────────────────────
const learningsTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [2400, 7680],
  rows: [
    headerRow([["Pattern observed", 2400], ["What it means for future research", 7680]]),
    ...[
      ["CRO / trader trap", "Mumbai's chemical belt has many companies that look like specialty manufacturers but are actually distributors or contract research orgs. Always verify: Does this company own a plant? Does it have manufacturing employees? If the website says 'solutions provider' without mentioning production capacity — dig further."],
      ["Revenue moves fast", "Several companies that looked right on last year's annual report had crossed Rs.500Cr by FY25. Always check the most recent quarterly filings on BSE / NSE, not just the last annual report."],
      ["PE ownership hidden", "PE acquisitions are not always obvious from the company name. Check MCA Director details and BSE shareholding patterns — if an entity like 'XYZ Opportunities Fund' holds >26% and a nominee director was recently appointed, flag it."],
      ["C7 is the swing criterion", "Companies with SAP / ERP investment almost always score A or B band even if C3 (differentiation) is moderate. Companies with zero systems maturity almost always fail regardless of how differentiated the product is. Prioritise C7 research early."],
      ["LinkedIn Job Postings work well", "Searching '[Company Name] jobs' on LinkedIn filtered to 'last 6 months' is the fastest way to check C6 Hiring signal. 5+ open roles is meaningful; 0 roles for 12+ months is a red flag."],
      ["Group subsidiaries", "Several Mumbai companies are subsidiaries of large groups (Godrej, Tata, Birla). If the company name appears in the parent's annual report as a 'subsidiary' or 'associate', it disqualifies — the founder dynamic is gone."],
    ].map(([pattern, meaning], i) =>
      new TableRow({ children: [
        cell(pattern, { w: 2400, bg: i%2===0 ? LGRAY : WHITE, bold: true }),
        cell(meaning, { w: 7680, bg: i%2===0 ? LGRAY : WHITE }),
      ]})
    )
  ]
});

// ── AI tools table ─────────────────────────────────────────────────────────────
const aiTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [2000, 3040, 5040],
  rows: [
    headerRow([["Tool", 2000], ["How it was used", 3040], ["Where human judgment was applied", 5040]]),
    ...[
      ["Claude (Anthropic)", "Initial candidate discovery prompts; structuring scoring logic; drafting outreach hooks", "Verifying every company claim independently via Screener, MCA, BSE. AI output was a starting list — not the final answer."],
      ["Screener.in", "Revenue, EBITDA, promoter holding, growth ratios", "Interpreting whether growth is organic vs. acquisition-driven. Reading annual report notes for systems maturity signals."],
      ["MCA portal", "Director histories, appointment dates, shareholding", "Distinguishing between a gen-2 family member who is 'director' in name vs. one with real operational authority."],
      ["LinkedIn Search", "Company size, job postings, founder credentials", "Deciding whether '5 open roles' is real growth hiring or replacement churn. Reading between job description lines for ERP/SAP signals."],
    ].map(([tool, used, judgment], i) =>
      new TableRow({ children: [
        cell(tool, { w: 2000, bg: i%2===0 ? LGRAY : WHITE, bold: true }),
        cell(used, { w: 3040, bg: i%2===0 ? LGRAY : WHITE }),
        cell(judgment, { w: 5040, bg: i%2===0 ? LGRAY : WHITE }),
      ]})
    )
  ]
});

// ── Build document ─────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      { reference: "bullets", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 560, hanging: 280 } } } },
        { level: 1, format: LevelFormat.BULLET, text: "\u25E6", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 1000, hanging: 280 } } } }
      ]},
      { reference: "numbers", levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 560, hanging: 280 } } } }
      ]}
    ]
  },
  styles: {
    default: { document: { run: { font: "Calibri", size: 20, color: "111827" } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Calibri", color: NAVY },
        paragraph: { spacing: { before: 320, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Calibri", color: TEAL },
        paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Calibri", color: "374151" },
        paragraph: { spacing: { before: 180, after: 60 }, outlineLevel: 2 } }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          children: [
            new TextRun({ text: "DeepThought  |  Mumbai Research Methodology  |  Confidential", font: "Calibri", size: 16, color: "9CA3AF" })
          ],
          alignment: AlignmentType.RIGHT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: MGRAY, space: 4 } },
          spacing: { after: 0 }
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          children: [
            new TextRun({ text: "DeepThought Business Analytics  |  Page ", font: "Calibri", size: 16, color: "9CA3AF" }),
            new PageNumber({ font: "Calibri", size: 16, color: "9CA3AF" })
          ],
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: MGRAY, space: 4 } }
        })]
      })
    },
    children: [
      // ── COVER ──────────────────────────────────────────────────────────────
      coverTitle, coverSub, coverMeta,

      // ── 1. OVERVIEW ────────────────────────────────────────────────────────
      h1("1.  Overview"),
      para([
        run("This document explains how 25 target companies were identified, screened, and scored for the DeepThought Mumbai Federer research assignment. It covers the sources used, the filtering logic applied at each stage, the scoring criteria and evidence standards, and the key patterns observed during research — including common traps and disqualifiers specific to Mumbai.")
      ], { before: 80, after: 80, line: 288 }),
      para([
        run("City: ", { bold: true }), run("Mumbai (primary operations — founders, R&D, or manufacturing based in Mumbai or the MMR belt: Navi Mumbai, Thane, Ambernath, Silvassa)"),
      ], { before: 60, after: 20 }),
      para([
        run("Segments: ", { bold: true }), run("Performance chemicals (Basket A), Specialty food & nutraceutical ingredients (Basket B), Custom synthesis & specialty chemicals (Basket A)")
      ], { before: 20, after: 20 }),
      para([
        run("Candidates researched: ", { bold: true }), run("~90   |   "),
        run("Gates applied: ", { bold: true }), run("E1 (producer) + E2 (India-based)   |   "),
        run("Yield: ", { bold: true }), run("25 companies scored 60+")
      ], { before: 20, after: 120 }),

      // ── 2. WHY MUMBAI ──────────────────────────────────────────────────────
      h1("2.  Why Mumbai for Specialty Chemicals"),
      para([run("Mumbai and the MMR belt (Navi Mumbai, Thane, Ambernath, Dahej corridor) host one of India's densest concentrations of specialty chemical manufacturers. Several structural reasons make this a productive segment for Federer research:")], { before: 60, after: 80, line: 288 }),
      bullet("Mumbai-based chemical companies are typically older, founder-built businesses — many were established in the 1970s–90s when the chemicals belt around Thane-Belapur-Ambernath was the primary industrial zone."),
      bullet("The segment naturally produces differentiated producers: specialty chemicals require proprietary chemistry, regulatory approvals, and long customer relationships — all barriers that fit the Federer differentiation criterion."),
      bullet("China+1 is actively reshaping this sector. Global customers diversifying supply chains away from China are actively qualifying Indian specialty chemical suppliers — creating a clear tailwind (C5)."),
      bullet("Mumbai's proximity to JNPT (Jawaharlal Nehru Port) makes export logistics easier, so many mid-size specialty chemical companies here have meaningful export businesses — a further differentiation signal."),
      bullet("The city has a strong promoter-operator tradition. Unlike IT or fintech, chemical companies tend to be founder-run for decades, often with gen-2 who studied chemistry or engineering and returned to the business."),
      spacer(120),

      // ── 3. DISCOVERY SOURCES ───────────────────────────────────────────────
      h1("3.  Discovery Sources"),
      para([run("The following sources were used to build the initial candidate list of ~90 companies:")], { before: 60, after: 100 }),
      sourceTable,
      spacer(80),
      h2("3.1  Discovery Search Approach"),
      para([run("Initial candidates were found through a combination of structured queries:")], { before: 60, after: 80 }),
      bullet("Screener.in screener: Filtered for Industry = Chemicals / Pharma / Food Processing; Revenue band Rs.50Cr–Rs.500Cr; City = Mumbai / Navi Mumbai / Thane; Promoter holding >40%."),
      bullet("LinkedIn Company Search: Industry = Chemicals; Employees 50–500; Location = Mumbai. Browsed company pages for product descriptions to identify producers vs. distributors."),
      bullet("Google search patterns used:"),
      bullet("[Segment keyword] + Mumbai + manufacturer + site:screener.in", 1),
      bullet("[Segment keyword] + Mumbai + \"USFDA\" OR \"ISO\" OR \"DSIR\" OR \"ERP\"", 1),
      bullet("BSE / NSE listed companies in specialty chemicals with Mumbai registered address — cross-referenced with Screener revenue data.", 1),
      bullet("Industry directory searches: CHEMEXCIL (Chemicals Export Promotion Council), FICCI chemicals committee member list, CII Maharashtra member directory."),
      spacer(120),

      // ── 4. FILTER FUNNEL ───────────────────────────────────────────────────
      h1("4.  Research Filter Funnel"),
      para([run("Every candidate passed through this sequence. Any failure at a gate stopped further research on that company — time was not spent scoring companies that failed basic eligibility.")], { before: 60, after: 100 }),
      funnelTable,
      spacer(80),
      h2("4.1  Common Disqualifiers in Mumbai"),
      para([run("The following patterns caused the most disqualifications in this batch and are worth flagging for future researchers:")], { before: 60, after: 80 }),
      bullet("Traders/distributors disguised as manufacturers: Mumbai's chemical industry has a large trading layer. Company websites often use words like 'supplier', 'solution provider', or 'specialty chemicals company' without confirming they actually manufacture. The test: Does the company have a listed plant address? Do they have production capacity figures in their annual report? Do job postings mention plant-based roles?"),
      bullet("PE ownership post-acquisition: Several previously excellent Federer candidates have been acquired by PE funds in the last 3–4 years (Suven Pharma → Advent International is the example in the sample data). Always check BSE shareholding pattern for 'Scheme of Arrangement' notices or promoter holding dropping below 50%."),
      bullet("Group subsidiaries: Multiple Mumbai chemical companies are technically independent but are majority-owned by large groups (Birla, Tata, Godrej). These fail the independent promoter-driven test. Check MCA for ultimate beneficial ownership."),
      bullet("Revenue too high — and moving fast: The specialty chemicals sector has been growing 20–30% per year. A company at Rs.480Cr last year may be at Rs.580Cr this year. Always cross-check the most recent quarterly revenue from BSE, not just the annual report."),
      spacer(120),

      // ── 5. SCORING CRITERIA ────────────────────────────────────────────────
      h1("5.  Federer Scoring Criteria"),
      para([run("Each company that passed both eligibility gates was scored on 6 criteria (total: 100 points). Scores are Weak (0), Moderate (half weight), or Strong (full weight).")], { before: 60, after: 100 }),
      rubricTable,
      spacer(80),
      h2("5.1  Evidence Standards"),
      para([run("Scores were only assigned when specific evidence was found. The following standards were applied:")], { before: 60, after: 80 }),
      numbered("C3 — Differentiation was not inferred from the product category. Actual evidence was required: a patent number, USFDA facility approval page, DSIR recognition list entry, specific regulatory filing, or documented proprietary process."),
      numbered("C4 — DM background was verified on LinkedIn, the company's About Us page, or published interviews. 'PhD' was not assumed from a name — the specific institution and field were checked."),
      numbered("C5 — Sector tailwinds were verified against PLI scheme notifications, CHEMEXCIL export data, and news coverage of China+1 procurement shifts in the specific sub-segment."),
      numbered("C6 — Growth signals were checked with dates. A LinkedIn job posting from 2 years ago does not count as a current hiring signal. Only roles posted within the last 6 months were counted."),
      numbered("C7 — ERP/SAP deployment was confirmed through: (a) job postings mentioning SAP/ERP proficiency, (b) IT team members on LinkedIn, (c) annual report references to digital infrastructure, or (d) direct mention in investor presentations. Absence of evidence was scored as Weak, not Moderate."),
      numbered("C8 — Gen-2 board appointment dates were checked on MCA. Being listed as a director was insufficient — the appointment date had to be within the last 10 years and an operational title (ED, MD, COO) was required to score Strong."),
      spacer(120),

      // ── 6. KEY LEARNINGS ──────────────────────────────────────────────────
      h1("6.  Key Learnings from This Segment"),
      para([run("These patterns emerged during the Mumbai specialty chemicals research and should inform future research in similar segments:")], { before: 60, after: 100 }),
      learningsTable,
      spacer(120),

      // ── 7. AI TOOL USAGE ──────────────────────────────────────────────────
      h1("7.  How AI Tools Were Used"),
      para([run("AI tools (primarily Claude) were used as a research accelerator, not as a source of truth. Every AI-generated suggestion was independently verified.")], { before: 60, after: 100 }),
      aiTable,
      spacer(80),
      para([
        run("Critical principle: ", { bold: true }),
        run("AI was never trusted for specific factual claims about companies — revenue figures, founder credentials, certifications, or ownership structure. These were always verified from primary sources (BSE filings, MCA, USFDA database, company website). AI output that could not be verified was discarded, not included with a caveat.")
      ], { before: 60, after: 120, line: 288 }),

      // ── 8. OUTPUT FORMAT ─────────────────────────────────────────────────
      h1("8.  Output Format"),
      para([run("The 25 qualifying companies are documented in the accompanying Excel spreadsheet (DeepThought_Mumbai_Federer.xlsx) with the following fields for each company:")], { before: 60, after: 80 }),
      bullet("Company name, website, city / operational hub, segment"),
      bullet("What they make — 1–2 specific lines, not generic category"),
      bullet("Revenue band estimate with source"),
      bullet("Decision maker name, title, and background summary"),
      bullet("E1 and E2 gate outcomes with one line of evidence each"),
      bullet("C3–C8 scores (Strong / Moderate / Weak) with one line of evidence each"),
      bullet("Federer Score out of 100 and band (A / B / C / D)"),
      bullet("Verdict and one-line reasoning"),
      bullet("Outreach hook — one specific, verifiable, recent detail about the company usable in line 1 of a cold email"),
      spacer(80),
      para([run("Fail companies are documented separately with clear disqualification reasons — not silently dropped. This allows future researchers to avoid re-researching the same companies.")], { before: 60, after: 60, line: 288 }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("/mnt/user-data/outputs/DeepThought_Mumbai_Methodology.docx", buf);
  console.log("Done.");
});
