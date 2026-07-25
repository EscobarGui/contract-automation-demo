/**
 * Contract Automation Demo
 * -------------------------
 * Simulates a real workflow: a company receives contract requests (normally
 * typed manually into a Word template, one at a time, ~20 minutes each).
 * This script reads the requests from a CSV, validates them, and generates
 * a ready-to-send .docx contract for every valid request automatically.
 *
 * Run:  node generate-contracts.js
 */

const fs = require("fs");
const path = require("path");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  AlignmentType,
} = require("docx");

const INPUT_FILE = path.join(__dirname, "contract_requests.csv");
const OUTPUT_DIR = path.join(__dirname, "generated_contracts");
const LOG_FILE = path.join(__dirname, "processing_log.csv");

// Reference point for the "before" comparison: manual drafting of one
// contract (find template, fill in fields, proofread, save, rename).
const MANUAL_MINUTES_PER_CONTRACT = 20;

// ---------- 1. Read + parse the CSV ----------
function parseCSV(filePath) {
  const raw = fs.readFileSync(filePath, "utf-8").trim();
  const [headerLine, ...lines] = raw.split("\n");
  const headers = headerLine.split(",");
  return lines.map((line) => {
    const cells = line.split(",");
    const row = {};
    headers.forEach((h, i) => (row[h.trim()] = (cells[i] || "").trim()));
    return row;
  });
}

// ---------- 2. Validate a single request ----------
function validateRequest(row) {
  const errors = [];

  if (!row.client_name) errors.push("missing client_name");
  if (!row.contract_type) errors.push("missing contract_type");

  if (!row.value || isNaN(Number(row.value))) {
    errors.push("missing or invalid value");
  }

  if (!/^\d{4}-\d{2}-\d{2}$/.test(row.start_date)) {
    errors.push("invalid start_date (expected YYYY-MM-DD)");
  }

  if (!row.email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(row.email)) {
    errors.push("missing or invalid email");
  }

  return errors;
}

// ---------- 3. Build the .docx contract for a valid request ----------
function formatCurrency(value) {
  return Number(value).toLocaleString("en-US", {
    style: "currency",
    currency: "USD",
    minimumFractionDigits: 2,
  });
}

function formatDate(isoDate) {
  const d = new Date(isoDate + "T00:00:00");
  return d.toLocaleDateString("en-US", {
    year: "numeric",
    month: "long",
    day: "numeric",
  });
}

function buildContractDoc(row, contractId) {
  const generatedOn = new Date().toLocaleDateString("en-US", {
    year: "numeric",
    month: "long",
    day: "numeric",
  });

  const infoRow = (label, value) =>
    new TableRow({
      children: [
        new TableCell({
          width: { size: 3000, type: WidthType.DXA },
          borders: noBorders(),
          children: [
            new Paragraph({
              children: [new TextRun({ text: label, bold: true })],
            }),
          ],
        }),
        new TableCell({
          width: { size: 6000, type: WidthType.DXA },
          borders: noBorders(),
          children: [new Paragraph({ text: value })],
        }),
      ],
    });

  function noBorders() {
    const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
    return { top: none, bottom: none, left: none, right: none };
  }

  const doc = new Document({
    sections: [
      {
        properties: {
          page: { size: { width: 12240, height: 15840 } }, // US Letter
        },
        children: [
          new Paragraph({
            text: row.contract_type.toUpperCase(),
            heading: HeadingLevel.TITLE,
            alignment: AlignmentType.CENTER,
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: `Contract ID: ${contractId}  |  Generated automatically on ${generatedOn}`,
                italics: true,
                size: 18,
                color: "666666",
              }),
            ],
          }),
          new Paragraph({ text: "", spacing: { after: 200 } }),

          new Paragraph({
            text: "Parties",
            heading: HeadingLevel.HEADING_2,
          }),
          new Table({
            width: { size: 9000, type: WidthType.DXA },
            columnWidths: [3000, 6000],
            rows: [
              infoRow("Client:", row.client_name),
              infoRow("Contact email:", row.email),
              infoRow("Provider:", "[Your Company Name]"),
            ],
          }),

          new Paragraph({
            text: "Commercial Terms",
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 300 },
          }),
          new Table({
            width: { size: 9000, type: WidthType.DXA },
            columnWidths: [3000, 6000],
            rows: [
              infoRow("Contract type:", row.contract_type),
              infoRow("Contract value:", formatCurrency(row.value)),
              infoRow("Start date:", formatDate(row.start_date)),
              infoRow("Payment terms:", row.payment_terms || "Net 30"),
            ],
          }),

          new Paragraph({
            text: "Terms & Conditions",
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 300 },
          }),
          new Paragraph({
            spacing: { after: 150 },
            children: [
              new TextRun(
                "1. Confidentiality. Both parties agree to keep the terms of this agreement, and any non-public information exchanged during its execution, confidential."
              ),
            ],
          }),
          new Paragraph({
            spacing: { after: 150 },
            children: [
              new TextRun(
                "2. Term & Termination. This agreement is effective from the start date above and may be terminated by either party with 30 days' written notice."
              ),
            ],
          }),
          new Paragraph({
            spacing: { after: 150 },
            children: [
              new TextRun(
                "3. Payment. Payment is due according to the payment terms specified above, invoiced from the start date."
              ),
            ],
          }),

          new Paragraph({
            spacing: { before: 400 },
            children: [
              new TextRun({
                text: "Signature: ____________________________          Date: ______________",
              }),
            ],
          }),
        ],
      },
    ],
  });

  return doc;
}

// ---------- 4. Run the pipeline ----------
async function run() {
  if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

  const requests = parseCSV(INPUT_FILE);
  const logRows = ["id,client_name,status,reason,generated_file,processing_time_ms"];

  let validCount = 0;
  let invalidCount = 0;
  const pipelineStart = process.hrtime.bigint();

  for (const row of requests) {
    const rowStart = process.hrtime.bigint();
    const errors = validateRequest(row);

    if (errors.length > 0) {
      invalidCount++;
      const reason = errors.join(" | ");
      console.log(`  [SKIPPED] Request #${row.id} (${row.client_name || "unnamed"}) -> ${reason}`);
      logRows.push(`${row.id},"${row.client_name}",error,"${reason}",,0`);
      continue;
    }

    const contractId = `CT-2026-${String(row.id).padStart(4, "0")}`;
    const doc = buildContractDoc(row, contractId);
    const fileName = `${contractId}_${row.client_name.replace(/\s+/g, "_")}.docx`;
    const filePath = path.join(OUTPUT_DIR, fileName);

    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(filePath, buffer);

    const rowEnd = process.hrtime.bigint();
    const elapsedMs = Number(rowEnd - rowStart) / 1e6;

    validCount++;
    console.log(`  [GENERATED] ${contractId} -> ${fileName}  (${elapsedMs.toFixed(1)} ms)`);
    logRows.push(`${row.id},"${row.client_name}",generated,,${fileName},${elapsedMs.toFixed(1)}`);
  }

  const pipelineEnd = process.hrtime.bigint();
  const totalAutomatedSeconds = Number(pipelineEnd - pipelineStart) / 1e9;
  const manualMinutes = validCount * MANUAL_MINUTES_PER_CONTRACT;
  const manualHours = (manualMinutes / 60).toFixed(1);
  const reductionPct = (
    ((manualMinutes * 60 - totalAutomatedSeconds) / (manualMinutes * 60)) *
    100
  ).toFixed(1);

  fs.writeFileSync(LOG_FILE, logRows.join("\n"));

  console.log("\n================ SUMMARY ================");
  console.log(`Requests received:     ${requests.length}`);
  console.log(`Contracts generated:   ${validCount}`);
  console.log(`Rejected (validation): ${invalidCount}`);
  console.log(`Automated time:        ${totalAutomatedSeconds.toFixed(2)} seconds`);
  console.log(`Manual equivalent:     ~${manualMinutes} minutes (~${manualHours} hours)`);
  console.log(`Estimated reduction:   ${reductionPct}%`);
  console.log("===========================================");
  console.log(`\nGenerated contracts: ${OUTPUT_DIR}`);
  console.log(`Processing log:      ${LOG_FILE}`);
}

run();
