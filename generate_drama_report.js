/**
 * Drama Report Generator for Bombo Radyo Philippines
 * Parses Raduga .LOG files and generates a Drama Report .DOCX
 *
 * Usage:
 *   node generate_drama_report.js [--folder <path>] [--out <output.docx>]
 *
 * --folder  Path to the folder containing Raduga .LOG files (default: current directory)
 * --out     Output filename (default: Drama_Report_<date>.docx)
 *
 * Automatically reads ALL .LOG files in the folder (Raduga keeps the last 4 days).
 * Picks the Monday–Thursday window; week range is derived from the log filenames.
 */

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, ImageRun, Header,
  HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom,
  HorizontalPositionAlign, VerticalPositionAlign,
  TextWrappingType,
} = require("docx");

// ─── CLI argument parsing ────────────────────────────────────────────────────
const args = process.argv.slice(2);
let logFolder = "\\\\192.168.86.134\\Raduga Log";
let outputFile = null; // auto-named below once we know the dates

for (let i = 0; i < args.length; i++) {
  if (args[i] === "--folder" && args[i + 1]) {
    logFolder = args[++i];
  } else if (args[i] === "--out" && args[i + 1]) {
    outputFile = args[++i];
  }
}

// ─── Auto-discover all .LOG files in folder ──────────────────────────────────
function discoverLogFiles(folder) {
  let files;
  try {
    files = fs.readdirSync(folder);
  } catch (e) {
    console.error("Cannot read folder: " + folder + " (" + e.message + ")");
    process.exit(1);
  }

  // Keep only files matching YYYY-MM-DD.LOG (Raduga naming convention)
  const logFiles = files
    .filter(f => /^\d{4}-\d{2}-\d{2}\.LOG$/i.test(f))
    .map(f => path.join(folder, f))
    .sort(); // ascending date order

  if (logFiles.length === 0) {
    console.error("No Raduga .LOG files found in: " + folder);
    console.error("Expected filenames like 2026-04-28.LOG");
    process.exit(1);
  }

  console.log("Found " + logFiles.length + " log file(s):");
  logFiles.forEach(f => console.log("   " + path.basename(f)));
  return logFiles;
}

const logFiles = discoverLogFiles(logFolder);

// ─── Log parser ──────────────────────────────────────────────────────────────
/**
 * Extracts drama title, season, and chapters from a Raduga log filename entry.
 *
 * Raduga log drama lines look like:
 *   HH:MM:SS  start  SALAMIN NG BUHAY-SEASON 2-CHAPTER 142 A
 *   HH:MM:SS  start  RECUERDOS DELA VIDA-SEASON 2-CHAPTER 103
 *
 * Returns: { title, season, chapter } or null
 */
function parseDramaEntry(logLine) {
  // Raduga log drama lines:
  //   HH:MM:SS  start  SALAMIN NG BUHAY-SEASON 2-CHAPTER 142 A
  //   HH:MM:SS  start  HIMIG NG PAG IBIG-SEASON 02-CHAPTER  141 A   (double space before chapter#)
  //   HH:MM:SS  start  MISTERIOSO ESTRANGHERO-SEASON 1 CHAPTER 17 A (no hyphen before CHAPTER)
  const match = logLine.match(
    /start\t(.+?)-SEASON\s+(\d+)-?\s*CHAPTER\s+(\d+)/i
  );
  if (!match) return null;

  const title = match[1].trim().toUpperCase();
  const season = parseInt(match[2], 10);
  const chapter = parseInt(match[3], 10);
  return { title, season, chapter };
}

/**
 * Parse a single log file and return array of drama entries.
 */
function parseLogFile(filePath) {
  const content = fs.readFileSync(filePath, "utf-8");
  const lines = content.split(/\r?\n/);
  const entries = [];

  for (const line of lines) {
    const entry = parseDramaEntry(line);
    if (entry) entries.push(entry);
  }
  return entries;
}

// ─── Fixed drama roster (always shown in this order) ────────────────────────
// logKey must match the uppercased title AS WRITTEN in the Raduga log.
// displayName is what appears in the report table.
const DRAMA_ROSTER = [
  { logKey: "SALAMIN NG BUHAY",        displayName: "SALAMIN NG BUHAY" },
  { logKey: "HIMIG NG PAG IBIG",       displayName: "HIMIG NG PAG-IBIG" },       // log has no hyphen
  { logKey: "DOMING PWEDE LAHAT",      displayName: "DOMING PWEDE LAHAT" },
  { logKey: "RECUERDOS",               displayName: "RECUERDOS DELA VIDA" },      // log shortens to RECUERDOS
  { logKey: "SUPER PEPITA",            displayName: "SUPER PEPITA" },
  { logKey: "MISTERIOSO ESTRANGHERO",  displayName: "MISTERIOSO ESTRANGHERO" },
];

/**
 * Aggregate all log files.
 * Always returns all 6 dramas in fixed order.
 * Dramas not found in the logs show "–" for season and chapters.
 */
function aggregateDramas(logFiles) {
  // Map: logKey => { season, minChapter, maxChapter }
  const found = new Map();

  for (const file of logFiles) {
    let entries;
    try {
      entries = parseLogFile(file);
    } catch (e) {
      console.warn(`Warning: Could not read ${file}: ${e.message}`);
      continue;
    }

    for (const { title, season, chapter } of entries) {
      if (!found.has(title)) {
        found.set(title, { season, minChapter: chapter, maxChapter: chapter });
      } else {
        const d = found.get(title);
        // If season changes mid-week, keep the latest one seen
        d.season = season;
        if (chapter < d.minChapter) d.minChapter = chapter;
        if (chapter > d.maxChapter) d.maxChapter = chapter;
      }
    }
  }

  // Build result in fixed roster order
  return DRAMA_ROSTER.map(({ logKey, displayName }) => {
    const data = found.get(logKey);
    return {
      title: displayName,
      season: data ? data.season : null,
      minChapter: data ? data.minChapter : null,
      maxChapter: data ? data.maxChapter : null,
      found: !!data,
    };
  });
}

// ─── Date helpers ────────────────────────────────────────────────────────────
function getWeekRange(logFiles) {
  const dates = [];
  for (const f of logFiles) {
    const base = path.basename(f, path.extname(f));
    // Parse as UTC noon to avoid timezone shifting the date
    const d = new Date(base + "T12:00:00Z");
    if (!isNaN(d)) dates.push(d);
  }
  if (dates.length === 0) return { start: new Date(), end: new Date() };
  dates.sort((a, b) => a - b);
  return { start: dates[0], end: dates[dates.length - 1] };
}

function formatDate(date) {
  return date.toLocaleDateString("en-US", {
    month: "long", day: "2-digit", year: "numeric", timeZone: "UTC"
  });
}

function isoDate(date) {
  return date.toISOString().slice(0, 10);
}

// ─── DOCX helpers ────────────────────────────────────────────────────────────
const BORDER = { style: BorderStyle.SINGLE, size: 8, color: "000000" };
const ALL_BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };
const NO_BORDERS = {
  top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};

function cell(text, opts = {}) {
  const {
    bold = false,
    align = AlignmentType.CENTER,
    shade = null,
    width = 3120,
    fontSize = 22,
    verticalAlign = VerticalAlign.CENTER,
    borders = ALL_BORDERS,
    italics = false,
  } = opts;

  const shadingProp = shade
    ? { fill: shade, type: ShadingType.CLEAR }
    : undefined;

  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: shadingProp,
    verticalAlign,
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
    children: [
      new Paragraph({
        alignment: align,
        children: [
          new TextRun({ text: String(text), bold, font: "Aptos", size: fontSize, italics }),
        ],
      }),
    ],
  });
}

function para(runs, opts = {}) {
  const { align = AlignmentType.LEFT, spacing = {} } = opts;
  return new Paragraph({
    alignment: align,
    spacing: { before: 0, after: 0, ...spacing },
    children: Array.isArray(runs) ? runs : [runs],
  });
}

function textRun(text, opts = {}) {
  return new TextRun({ text, font: "Aptos", size: 22, ...opts });
}

// ─── Main document builder ───────────────────────────────────────────────────
async function buildReport(logFiles, outputPath) {
  const dramas = aggregateDramas(logFiles);
  const { start, end } = getWeekRange(logFiles);

  // Auto-name output if not specified and save to specific directory
  if (!outputPath) {
    const year = end.toLocaleString("en-US", { year: "numeric", timeZone: "UTC" });
    const month = end.toLocaleString("en-US", { month: "long", timeZone: "UTC" });
    const dirPath = path.join("D:\\Documents\\Reports\\Drama", year, month);
    
    if (!fs.existsSync(dirPath)) {
      fs.mkdirSync(dirPath, { recursive: true });
    }
    
    outputPath = path.join(dirPath, "Drama Report - " + formatDate(end) + ".docx");
  }

  const reportDate = formatDate(end);
  const weekRange = start.getTime() === end.getTime()
    ? formatDate(start)
    : `${formatDate(start)} – ${formatDate(end)}`;

  // ── Drama table rows ────────────────────────────────────────────────────
  const COL_WIDTHS = [3120, 1560, 4680]; // Cebuano | Season | Chapters
  const TABLE_WIDTH = COL_WIDTHS.reduce((a, b) => a + b, 0); // 9360

  const headerRow = new TableRow({
    children: [
      new TableCell({
        columnSpan: 3,
        borders: ALL_BORDERS,
        shading: { fill: "FFFFFF", type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        verticalAlign: VerticalAlign.CENTER,
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: "DRAMA REPORT", bold: true, font: "Aptos", size: 24 }),
            ],
          }),
        ],
      }),
    ],
  });

  const subHeaderRow = new TableRow({
    children: [
      cell("Cebuano",  { bold: false }),
      cell("Season",   { bold: false }),
      cell("Chapters", { bold: false }),
    ],
  });

  const dataRows = dramas.map(({ title, season, minChapter, maxChapter, found }) => {
    const seasonText  = found ? String(season) : "\u2013";
    const chapterText = found
      ? (minChapter === maxChapter ? `${minChapter}` : `${minChapter} \u2013 ${maxChapter}`)
      : "\u2013";
    return new TableRow({
      children: [
        cell(title,       { bold: false }),
        cell(seasonText,  { bold: false }),
        cell(chapterText, { bold: false }),
      ],
    });
  });

  const dramaTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [headerRow, subHeaderRow, ...dataRows],
  });

  // ── Document sections ───────────────────────────────────────────────────
  const headerChildren = [
    // ── Letterhead area (with image) ──────────────────────────
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 600 },
      children: [
        textRun(" "), // Invisible text run to prevent Word XML schema error for image-only paragraphs
        new ImageRun({
          data: fs.readFileSync("C:\\Users\\User\\Pictures\\Logo\\Basta Radyo... BOMBO!\\Basta Radyo... BOMBO!.png"),
          transformation: {
            width: 94,
            height: 120,
          },
          floating: {
            horizontalPosition: {
              relative: HorizontalPositionRelativeFrom.MARGIN,
              align: HorizontalPositionAlign.LEFT,
            },
            verticalPosition: {
              relative: VerticalPositionRelativeFrom.PARAGRAPH,
              align: VerticalPositionAlign.TOP,
            },
            wrap: {
              type: TextWrappingType.NONE,
            },
            behindDocument: false,
          },
        }),
        new ImageRun({
          data: fs.readFileSync("C:\\Users\\User\\Pictures\\Logo\\Bombo Radyo Philippines - No.1 Radio Network in the Country.png"),
          transformation: {
            width: 329,
            height: 48,
          },
          floating: {
            horizontalPosition: {
              relative: HorizontalPositionRelativeFrom.MARGIN,
              align: HorizontalPositionAlign.CENTER,
            },
            verticalPosition: {
              relative: VerticalPositionRelativeFrom.PARAGRAPH,
              align: VerticalPositionAlign.TOP,
            },
            wrap: {
              type: TextWrappingType.NONE,
            },
            behindDocument: false,
          },
        }),
        new ImageRun({
          data: fs.readFileSync("C:\\Users\\User\\Pictures\\Logo\\Star FM - It's All For You\\Star FM - It's All For You.png"),
          transformation: {
            width: 94,
            height: 120,
          },
          floating: {
            horizontalPosition: {
              relative: HorizontalPositionRelativeFrom.MARGIN,
              align: HorizontalPositionAlign.RIGHT,
            },
            verticalPosition: {
              relative: VerticalPositionRelativeFrom.PARAGRAPH,
              align: VerticalPositionAlign.TOP,
            },
            wrap: {
              type: TextWrappingType.NONE,
            },
            behindDocument: false,
          },
        }),
      ],
    }),
    para(
      [textRun("Peoples Broadcasting Services, Inc.", { bold: false, size: 20 })],
      { align: AlignmentType.CENTER, spacing: { before: 0, after: 40 } }
    ),
    para(
      [textRun("DXCB 864 – Malaybalay City, Bukidnon", { bold: false, size: 20 })],
      { align: AlignmentType.CENTER, spacing: { before: 0, after: 0 } }
    ),
  ];

  const children = [
    // 2 Paragraph Breaks before the Date
    para([]),
    para([]),

    // ── Date ────────────────────────────────────────────────────────────
    para([textRun(reportDate)], { spacing: { before: 0, after: 0 } }),
    
    // 2 Paragraph Breaks after the Date
    para([]),
    para([]),

    // ── Addressee ───────────────────────────────────────────────────────
    para([textRun("Mr. Ricky E. Collado", { bold: true })]),
    para([textRun("Area Manager AM Division")]),
    para([textRun("Mindanao Area")]),
    para([textRun("Davao City")], { spacing: { before: 0, after: 0 } }),

    // 2 Paragraph Breaks after the Addressee block
    para([]),
    para([]),

    // ── Body text ───────────────────────────────────────────────────────
    new Paragraph({
      alignment: AlignmentType.BOTH,
      spacing: { before: 0, after: 0 },
      indent: { firstLine: 720 },
      children: [
        textRun("This is to formally report that during the Facebook live of the drama titled "),
        textRun("Recuerdos Dela Vida – Season # – Chapter #", { bold: true }),
        textRun(" which was streamed via Facebook Live."),
      ],
    }),
    para([]),

    new Paragraph({
      alignment: AlignmentType.BOTH,
      spacing: { before: 0, after: 0 },
      indent: { firstLine: 720 },
      children: [
        textRun("The video was subjected to copyright-related actions automatically enforced by Facebook's content recognition system. The detected claim was reportedly triggered by a song titled "),
        textRun('"Kiss Broken Hearts".', { bold: true, italics: true }),
      ],
    }),
    para([]),

    new Paragraph({
      alignment: AlignmentType.BOTH,
      spacing: { before: 0, after: 0 },
      indent: { firstLine: 720 },
      children: [
        textRun("To prevent further violations and reduce the risk to our Facebook Page, "),
        textRun("we have already deleted the live video. ", { bold: true }),
        textRun("Should you require further details or clarification, please do not hesitate to contact me."),
      ],
    }),
    para([]),

    // ── Drama Table ─────────────────────────────────────────────────────
    dramaTable,

    // ── Spacer ──────────────────────────────────────────────────────────
    para([], { spacing: { before: 360, after: 120 } }),

    // ── Signature block ─────────────────────────────────────────────────
    new Paragraph({
      spacing: { before: 0, after: 40 },
      children: [
        textRun("Prepared by:", { bold: true }),
        new TextRun({ text: "\t", font: "Aptos", size: 22 }),
        textRun("Noted by:", { bold: true }),
      ],
      tabStops: [{ type: "left", position: 4680 }],
    }),

    // Signature line space
    para([], { spacing: { before: 720, after: 40 } }),

    new Paragraph({
      spacing: { before: 0, after: 40 },
      children: [
        textRun("Dashiel Dinopol", { underline: { type: "single" } }),
        new TextRun({ text: "\t", font: "Aptos", size: 22 }),
        textRun("Michael Licuanan", { underline: { type: "single" } }),
      ],
      tabStops: [{ type: "left", position: 4680 }],
    }),

    new Paragraph({
      spacing: { before: 0, after: 0 },
      children: [
        textRun("Malaybalay City, Bukidnon \u2013 Technician"),
        new TextRun({ text: "\t", font: "Aptos", size: 22 }),
        textRun("Station Manager"),
      ],
      tabStops: [{ type: "left", position: 4680 }],
    }),
  ];

  // ── Build document ──────────────────────────────────────────────────────
  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: "Aptos", size: 22 } },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 12240, height: 15840 }, // Letter: 8.5" x 11" (in twentieths of a point)
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }, // 1" margins
          },
        },
        headers: {
          default: new Header({
            children: headerChildren,
          }),
        },
        children,
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
  const foundCount = dramas.filter(d => d.found).length;
  console.log("Drama Report saved to: " + outputPath);
  console.log("   " + foundCount + " of " + dramas.length + " dramas logged this week:");
  for (const d of dramas) {
    if (d.found) {
      const range = d.minChapter === d.maxChapter
        ? "Ch. " + d.minChapter
        : "Ch. " + d.minChapter + "-" + d.maxChapter;
      console.log("   [found] " + d.title + " | Season " + d.season + " | " + range);
    } else {
      console.log("   [--] " + d.title + " | not found in logs");
    }
  }
}

// ─── Run ─────────────────────────────────────────────────────────────────────
buildReport(logFiles, outputFile).catch((err) => {
  console.error("Error generating report:", err);
  process.exit(1);
});
