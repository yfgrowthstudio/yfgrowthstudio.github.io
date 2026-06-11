const path = require('path');
const GLOBAL = 'C:\\Users\\1\\AppData\\Roaming\\npm\\node_modules';
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
        LevelFormat, Footer, PageNumber } = require(path.join(GLOBAL, 'docx'));
const fs = require('fs');

const BRAND = "1452F0";
const INK = "0E1726";
const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

function cell(text, w, opts = {}) {
  return new TableCell({
    borders,
    width: { size: w, type: WidthType.DXA },
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text, bold: !!opts.bold, color: opts.color || INK, size: 20 })] })]
  });
}

function kvTable(rows) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2600, 6760],
    rows: rows.map(([k, v]) => new TableRow({
      children: [
        cell(k, 2600, { bold: true, fill: "EEF3FE" }),
        cell(v, 6760),
      ]
    }))
  });
}

function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] });
}
function bullet(text) {
  return new Paragraph({ numbering: { reference: "bullets", level: 0 },
    children: [new TextRun({ text, size: 20 })] });
}
function body(text) {
  return new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text, size: 20 })] });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, font: "Arial", color: INK },
        paragraph: { spacing: { before: 120, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: BRAND },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 1 } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 460, hanging: 260 } } } }] },
    ]
  },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 },
      margin: { top: 1080, right: 1440, bottom: 1080, left: 1440 } } },
    footers: { default: new Footer({ children: [ new Paragraph({ alignment: AlignmentType.CENTER,
      children: [ new TextRun({ text: "YF Growth Studio — Google Ads API Tool Design  |  Page ", size: 16, color: "888888" }),
                  new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "888888" }) ] }) ] }) },
    children: [
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Google Ads API Tool — Design Document")] }),
      new Paragraph({ spacing: { after: 160 }, children: [
        new TextRun({ text: "YF Growth Studio · Internal Ads Monitoring & Optimization Tool · v1.0", color: "555555", size: 18 }) ] }),

      kvTable([
        ["Applicant", "YF Growth Studio (independent agency)"],
        ["Manager account (MCC)", "875-893-6695"],
        ["Contact", "jiangyufan718@gmail.com"],
        ["Website", "https://yfgrowthstudio.github.io/"],
        ["Access requested", "Basic Access"],
        ["Distribution", "Internal use only — not distributed to third parties"],
      ]),

      h2("1. Purpose"),
      body("This tool is an internal monitoring and optimization system used by YF Growth Studio to manage the Google Ads accounts of our e-commerce clients, all linked under our own manager account (MCC). It is not a public product and is not resold or distributed. Its goal is to automate the repetitive, high-volume parts of day-to-day account management while keeping every spend-affecting decision under explicit human review."),

      h2("2. Architecture Overview"),
      body("The tool follows a propose-then-approve model: automated jobs read account data and generate recommendations; a human operator reviews and approves any change before it is written back to Google Ads."),
      bullet("Data layer — Scheduled jobs call the Google Ads API to pull campaign, ad group, keyword, search-term and budget data into an internal datastore."),
      bullet("Analysis layer — A real-time dashboard computes KPIs and runs anomaly detection across four buckets: low ROAS, low CTR, zero-conversion spend, and abnormal entity status."),
      bullet("Recommendation layer — The system proposes negative keywords (from search-term reports) and budget/status adjustments, surfaced as a review queue."),
      bullet("Approval gate — No mutate call is sent automatically. An operator approves each recommendation; only then is the change written via the API."),
      bullet("Alerting — Budget pacing and overspend conditions raise notifications to the operator."),

      h2("3. Google Ads API Usage"),
      kvTable([
        ["Primary read calls", "GoogleAdsService.SearchStream / Search (GAQL) for campaigns, ad groups, keywords, search terms, budgets and metrics"],
        ["Mutate calls (gated)", "CampaignBudgetService, CampaignService, AdGroupCriterionService (negative keywords), status changes — all behind human approval"],
        ["Frequency", "Scheduled reporting pulls plus on-demand refresh; well within Basic Access daily limits"],
        ["Scale", "A small number of accounts under one MCC; low overall call volume"],
      ]),

      h2("4. Human Oversight & Safety"),
      bullet("Every money-impacting change (budget, bid, status, negative keyword) requires explicit operator approval before the API mutate is executed."),
      bullet("Automated optimization runs on a weekly cadence to respect Smart Bidding learning periods; daily automation is limited to read-only monitoring and loss-prevention alerts."),
      bullet("Recommendations are logged with before/after values for auditability."),

      h2("5. Access, Security & Compliance"),
      bullet("OAuth 2.0 with least-privilege scopes; refresh tokens and the developer token are stored as secrets, never in client-side code or shared channels."),
      bullet("Client accounts are linked to the MCC by invitation; the tool never stores client login credentials."),
      bullet("Use complies with the Google Ads API Terms & Conditions and Google Ads policies; the API contact email is kept current to receive service notices."),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(path.join(__dirname, 'YF_Growth_Studio_API_Tool_Design.docx'), buf);
  console.log('written');
});
