// excelGenerators.js
import ExcelJS from "exceljs";

// ── Color palette ─────────────────────────────────────────────────────────────
const C = {
  DARK_BLUE: "FF0D1B2A",
  MID_BLUE: "FF1A3C5E",
  ACCENT: "FF1E90FF",
  GREEN: "FF22C55E",
  YELLOW: "FFF59E0B",
  RED: "FFEF4444",
  GREY: "FF64748B",
  LIGHT_BG: "FFF0F4FB",
  WHITE: "FFFFFFFF",
  ROW_EVEN: "FFEEF2FA",
  ROW_ODD: "FFFFFFFF",
  PASS_BG: "FFDCFCE7",
  PARTIAL_BG: "FFFEF3C7",
  FAIL_BG: "FFFEE2E2",
  NA_BG: "FFF1F5F9",
};

function scoreColor(score) {
  if (score >= 95) return "FF22C55E";
  if (score >= 85) return "FF84CC16";
  if (score >= 70) return "FFF59E0B";
  return "FFEF4444";
}

function statusFill(status) {
  const m = {
    Pass: C.PASS_BG,
    Partial: C.PARTIAL_BG,
    Fail: C.FAIL_BG,
    "N.A.": C.NA_BG,
  };
  return m[status] ?? C.NA_BG;
}

function statusFontColor(status) {
  const m = {
    Pass: "FF16A34A",
    Partial: "FFD97706",
    Fail: "FFDC2626",
    "N.A.": C.GREY,
  };
  return m[status] ?? C.GREY;
}

// ── Style helpers ─────────────────────────────────────────────────────────────
function fill(hex) {
  return { type: "pattern", pattern: "solid", fgColor: { argb: hex } };
}

function font(opts = {}) {
  return {
    name: "Arial",
    size: opts.size ?? 10,
    bold: opts.bold ?? false,
    italic: opts.italic ?? false,
    color: { argb: opts.color ?? "FF000000" },
  };
}

function thinBorder(color = "FFD0D8E8") {
  const s = { style: "thin", color: { argb: color } };
  return { top: s, left: s, bottom: s, right: s };
}

function medBorder(color = "FF1A3C5E") {
  const s = { style: "medium", color: { argb: color } };
  return { top: s, left: s, bottom: s, right: s };
}

function alignCenter() {
  return { horizontal: "center", vertical: "middle", wrapText: true };
}

function alignLeft() {
  return { horizontal: "left", vertical: "middle", wrapText: true, indent: 1 };
}

function applyStyle(cell, opts = {}) {
  if (opts.fill) cell.fill = opts.fill;
  if (opts.font) cell.font = opts.font;
  if (opts.alignment) cell.alignment = opts.alignment;
  if (opts.border) cell.border = opts.border;
}

// ── Title row: full-width merged heading ──────────────────────────────────────
function titleRow(
  ws,
  row,
  colStart,
  colEnd,
  text,
  bg,
  fg,
  fontSize = 13,
  bold = true,
  height = 28,
) {
  // merge first so the cell exists before styling
  if (colEnd > colStart) ws.mergeCells(row, colStart, row, colEnd);
  const cell = ws.getCell(row, colStart);
  cell.value = text;
  applyStyle(cell, {
    fill: fill(bg),
    font: font({ bold, size: fontSize, color: fg }),
    alignment: alignCenter(),
    border: thinBorder(bg === C.DARK_BLUE ? C.MID_BLUE : "FFD0D8E8"),
  });
  ws.getRow(row).height = height;
}

// ── Header row ────────────────────────────────────────────────────────────────
function headerRow(ws, row, headers, widths, bg = C.MID_BLUE) {
  ws.getRow(row).height = 26;
  headers.forEach((h, i) => {
    const c = ws.getCell(row, i + 1);
    c.value = h;
    applyStyle(c, {
      fill: fill(bg),
      font: font({ bold: true, size: 10, color: C.WHITE }),
      alignment: alignCenter(),
      border: thinBorder("FF0D1B2A"),
    });
    if (widths[i]) ws.getColumn(i + 1).width = widths[i];
  });
}

// ── Generic data cell ─────────────────────────────────────────────────────────
function dataCell(ws, row, col, value, opts = {}) {
  const c = ws.getCell(row, col);
  c.value = value;
  applyStyle(c, {
    fill: fill(opts.bg ?? (row % 2 === 0 ? C.ROW_EVEN : C.ROW_ODD)),
    font: font({
      size: opts.size ?? 10,
      bold: opts.bold ?? false,
      color: opts.color ?? "FF000000",
      italic: opts.italic ?? false,
    }),
    alignment: opts.align === "center" ? alignCenter() : alignLeft(),
    border: thinBorder(),
  });
  return c;
}

// ── KPI block: label above value ──────────────────────────────────────────────
function kpiBlock(ws, labelRow, valueRow, col, label, value, valueColor) {
  const lc = ws.getCell(labelRow, col);
  lc.value = label;
  applyStyle(lc, {
    fill: fill(C.ROW_EVEN),
    font: font({ bold: true, size: 9, color: C.MID_BLUE }),
    alignment: alignCenter(),
    border: thinBorder(),
  });
  const vc = ws.getCell(valueRow, col);
  vc.value = value;
  applyStyle(vc, {
    fill: fill(C.WHITE),
    font: font({ bold: true, size: 14, color: valueColor }),
    alignment: alignCenter(),
    border: thinBorder(),
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
//  INDIVIDUAL EXCEL
// ═══════════════════════════════════════════════════════════════════════════════

export async function generateIndividualExcel(sub) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Yakka Puka SLA Audit";
  wb.created = new Date();

  buildSummarySheet(wb, sub);
  buildSectionsSheet(wb, sub);
  buildChecklistSheet(wb, sub);
  buildAnalysisSheet(wb, sub);

  const buf = await wb.xlsx.writeBuffer();
  return buf;
}

// ── Sheet 1: Summary ──────────────────────────────────────────────────────────
function buildSummarySheet(wb, sub) {
  const ws = wb.addWorksheet("📋 Summary", {
    views: [{ showGridLines: false }],
  });

  const cf = sub.customerFeedback || {};
  const sc = scoreColor(sub.overallScore ?? 0);
  const COLS = 10;

  // Set column widths upfront so merges render correctly
  const colWidths = [6, 6, 14, 14, 14, 14, 14, 14, 14, 14];
  colWidths.forEach((w, i) => {
    ws.getColumn(i + 1).width = w;
  });

  let r = 1;

  // ── Main Title ──
  titleRow(
    ws,
    r,
    1,
    COLS,
    "🏢  YAKKA PUKA — SLA AUDIT REPORT",
    C.DARK_BLUE,
    C.WHITE,
    16,
    true,
    42,
  );
  r++;
  titleRow(
    ws,
    r,
    1,
    COLS,
    `Generated: ${new Date().toLocaleString("en-IN", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" })}`,
    C.MID_BLUE,
    "FFA0C4E8",
    9,
    false,
    20,
  );

  // ── Score & Rating ──
  r += 2;
  titleRow(ws, r, 1, 5, "OVERALL SCORE", C.ACCENT, C.WHITE, 11, true, 24);
  titleRow(ws, r, 6, COLS, "RATING", C.ACCENT, C.WHITE, 11, true, 24);

  r++;
  ws.mergeCells(r, 1, r + 1, 5);
  const scoreCell = ws.getCell(r, 1);
  scoreCell.value = `${sub.overallScore ?? 0}%`;
  applyStyle(scoreCell, {
    fill: fill(C.LIGHT_BG),
    font: font({ bold: true, size: 40, color: sc }),
    alignment: alignCenter(),
    border: thinBorder(),
  });

  ws.mergeCells(r, 6, r, COLS);
  const ratingCell = ws.getCell(r, 6);
  ratingCell.value = sub.ratingLabel ?? "";
  applyStyle(ratingCell, {
    fill: fill(C.LIGHT_BG),
    font: font({ bold: true, size: 22, color: sc }),
    alignment: alignCenter(),
    border: thinBorder(),
  });
  ws.getRow(r).height = 46;

  r++;
  ws.mergeCells(r, 6, r, COLS);
  const actionCell = ws.getCell(r, 6);
  actionCell.value = sub.ratingAction ?? "";
  applyStyle(actionCell, {
    fill: fill(C.LIGHT_BG),
    font: font({ size: 10, italic: true, color: C.GREY }),
    alignment: alignCenter(),
    border: thinBorder(),
  });
  ws.getRow(r).height = 24;

  // ── Info Section ──
  r += 2;
  titleRow(
    ws,
    r,
    1,
    COLS,
    "👤  AUDITOR & CUSTOMER INFORMATION",
    C.MID_BLUE,
    C.WHITE,
    11,
    true,
    24,
  );
  r++;

  const infoPairs = [
    ["Auditor Name", sub.auditorName ?? "—"],
    ["Auditor Email", sub.auditorEmail ?? "—"],
    [
      "Submitted At",
      sub.submittedAt ? new Date(sub.submittedAt).toLocaleString("en-IN") : "—",
    ],
    ["Customer Name", cf.name ?? "—"],
    ["Mobile", cf.mobile ?? "—"],
    ["Company", cf.company ?? "—"],
    ["Comments", cf.otherComments ?? "—"],
  ];

  for (const [label, value] of infoPairs) {
    ws.mergeCells(r, 1, r, 3);
    const lc = ws.getCell(r, 1);
    lc.value = label;
    applyStyle(lc, {
      fill: fill(C.ROW_EVEN),
      font: font({ bold: true, size: 10, color: C.MID_BLUE }),
      alignment: alignLeft(),
      border: thinBorder(),
    });
    ws.mergeCells(r, 4, r, COLS);
    const vc = ws.getCell(r, 4);
    vc.value = value;
    applyStyle(vc, {
      fill: fill(C.ROW_ODD),
      font: font({ size: 10 }),
      alignment: alignLeft(),
      border: thinBorder(),
    });
    ws.getRow(r).height = 22;
    r++;
  }

  // ── Statistics ──
  r++;
  titleRow(
    ws,
    r,
    1,
    COLS,
    "📊  AUDIT STATISTICS",
    C.MID_BLUE,
    C.WHITE,
    11,
    true,
    24,
  );
  r++;

  const stats = [
    { l: "Total Items", v: sub.totalItems ?? 0, c: C.GREY },
    { l: "Total Pass", v: sub.totalPass ?? 0, c: C.GREEN },
    { l: "Partial", v: sub.totalPartial ?? 0, c: C.YELLOW },
    { l: "Total Fail", v: sub.totalFail ?? 0, c: C.RED },
    { l: "N.A.", v: sub.totalNA ?? 0, c: C.GREY },
    { l: "Earned Marks", v: sub.earnedMarks ?? 0, c: C.ACCENT },
    { l: "Max Marks", v: sub.maxMarks ?? 0, c: C.MID_BLUE },
    { l: "Pass Rate", v: `${sub.passRate ?? 0}%`, c: sc },
  ];

  stats.forEach((s, i) => kpiBlock(ws, r, r + 1, i + 1, s.l, s.v, s.c));
  ws.getRow(r).height = 20;
  ws.getRow(r + 1).height = 32;

  // ── Progress bar ──
  r += 3;
  titleRow(
    ws,
    r,
    1,
    COLS,
    `SCORE PROGRESS: ${sub.overallScore ?? 0}%`,
    C.LIGHT_BG,
    C.MID_BLUE,
    10,
    true,
    22,
  );
  r++;
  const filled = Math.max(
    1,
    Math.round((sub.overallScore ?? 0) / (100 / COLS)),
  );
  for (let i = 0; i < COLS; i++) {
    const c = ws.getCell(r, i + 1);
    c.fill = fill(i < filled ? sc : "FFE2E8F0");
    c.border = thinBorder();
    c.alignment = alignCenter();
  }
  ws.getRow(r).height = 18;
}

// ── Sheet 2: Section Scores ───────────────────────────────────────────────────
function buildSectionsSheet(wb, sub) {
  const ws = wb.addWorksheet("📈 Section Scores", {
    views: [{ showGridLines: false }],
  });

  const headers = [
    "Section",
    "Icon",
    "Score %",
    "Earned Pts",
    "Max Pts",
    "Pass",
    "Partial",
    "Fail",
    "N.A.",
  ];
  const widths = [38, 8, 13, 13, 11, 9, 10, 9, 9];

  let r = 1;
  titleRow(
    ws,
    r,
    1,
    headers.length,
    "📈  SECTION-WISE SCORES",
    C.DARK_BLUE,
    C.WHITE,
    13,
    true,
    32,
  );
  r++;
  headerRow(ws, r, headers, widths);
  r++;

  for (const [idx, sec] of (sub.sections ?? []).entries()) {
    const sc = scoreColor(sec.sectionScore ?? 0);
    const bg = idx % 2 === 0 ? C.ROW_EVEN : C.ROW_ODD;
    const vals = [
      sec.sectionTitle ?? "",
      sec.sectionIcon ?? "",
      `${sec.sectionScore ?? 0}%`,
      sec.earnedPoints ?? 0,
      sec.maxPoints ?? 0,
      sec.passCount ?? 0,
      sec.partialCount ?? 0,
      sec.failCount ?? 0,
      sec.naCount ?? 0,
    ];

    vals.forEach((val, ci) => {
      const c = ws.getCell(r, ci + 1);
      c.value = val;
      applyStyle(c, {
        fill: fill(bg),
        border: thinBorder(),
        alignment: ci === 0 ? alignLeft() : alignCenter(),
      });
      if (ci === 0) c.font = font({ bold: true, size: 10, color: C.DARK_BLUE });
      else if (ci === 2) c.font = font({ bold: true, size: 11, color: sc });
      else if (ci === 5)
        c.font = font({ bold: true, size: 10, color: "FF16A34A" });
      else if (ci === 7)
        c.font = font({
          bold: true,
          size: 10,
          color: (sec.failCount ?? 0) > 0 ? "FFDC2626" : C.GREY,
        });
      else c.font = font({ size: 10 });
    });
    ws.getRow(r).height = 24;
    r++;

    // Mini progress bar
    const lbl = ws.getCell(r, 1);
    lbl.value = "Progress:";
    lbl.font = font({ size: 8, italic: true, color: C.GREY });
    lbl.alignment = alignLeft();

    const barFill = Math.max(1, Math.round((sec.sectionScore ?? 0) / 10));
    for (let bi = 0; bi < 10; bi++) {
      const bc = ws.getCell(r, bi + 2);
      bc.fill = fill(bi < barFill ? sc : "FFE2E8F0");
      bc.border = thinBorder();
      bc.alignment = alignCenter();
    }
    ws.getRow(r).height = 10;
    r++;
  }

  // Totals row
  const totals = [
    "TOTAL / OVERALL",
    "",
    `${sub.overallScore ?? 0}%`,
    sub.earnedMarks ?? 0,
    sub.maxMarks ?? 0,
    sub.totalPass ?? 0,
    sub.totalPartial ?? 0,
    sub.totalFail ?? 0,
    sub.totalNA ?? 0,
  ];
  totals.forEach((val, ci) => {
    const c = ws.getCell(r, ci + 1);
    c.value = val;
    applyStyle(c, {
      fill: fill(C.MID_BLUE),
      font: font({ bold: true, size: 10, color: C.WHITE }),
      alignment: ci === 0 ? alignLeft() : alignCenter(),
      border: medBorder(),
    });
  });
  ws.getRow(r).height = 24;
}

// ── Sheet 3: Checklist ────────────────────────────────────────────────────────
function buildChecklistSheet(wb, sub) {
  const ws = wb.addWorksheet("✅ Checklist", {
    views: [{ showGridLines: false }],
  });

  const headers = [
    "#",
    "Section",
    "Checklist Item",
    "Frequency",
    "Status",
    "Points",
    "Remark",
  ];
  const widths = [6, 24, 52, 16, 13, 9, 36];

  let r = 1;
  titleRow(
    ws,
    r,
    1,
    headers.length,
    "✅  DETAILED CHECKLIST",
    C.DARK_BLUE,
    C.WHITE,
    13,
    true,
    32,
  );
  r++;
  headerRow(ws, r, headers, widths);
  r++;

  let rowIdx = 0;
  for (const sec of sub.sections ?? []) {
    // Section header band
    ws.mergeCells(r, 1, r, 5);
    const sc = ws.getCell(r, 1);
    sc.value = `${sec.sectionIcon ?? ""}  ${sec.sectionTitle ?? ""}`;
    applyStyle(sc, {
      fill: fill(C.MID_BLUE),
      font: font({ bold: true, size: 11, color: C.WHITE }),
      alignment: alignLeft(),
      border: thinBorder("FF0D1B2A"),
    });

    ws.mergeCells(r, 6, r, 7);
    const sp = ws.getCell(r, 6);
    sp.value = `Score: ${sec.sectionScore ?? 0}%`;
    applyStyle(sp, {
      fill: fill(C.MID_BLUE),
      font: font({ bold: true, size: 11, color: C.WHITE }),
      alignment: alignCenter(),
      border: thinBorder("FF0D1B2A"),
    });
    ws.getRow(r).height = 26;
    r++;

    for (const row of sec.rows ?? []) {
      const status = row.status ?? "N.A.";
      const pts = row.points ?? 0;
      const bg = rowIdx % 2 === 0 ? C.ROW_EVEN : C.ROW_ODD;

      dataCell(ws, r, 1, (row.itemIndex ?? 0) + 1, { bg, align: "center" });
      dataCell(ws, r, 2, sec.sectionTitle ?? "", { bg });
      dataCell(ws, r, 3, row.itemText ?? "", { bg });
      dataCell(ws, r, 4, row.frequency ?? "", { bg, align: "center" });

      // Status cell
      const statusCell = ws.getCell(r, 5);
      statusCell.value = status;
      applyStyle(statusCell, {
        fill: fill(statusFill(status)),
        font: font({ bold: true, size: 10, color: statusFontColor(status) }),
        alignment: alignCenter(),
        border: thinBorder(),
      });

      // Points cell
      const ptsCell = ws.getCell(r, 6);
      ptsCell.value = pts;
      applyStyle(ptsCell, {
        fill: fill(bg),
        font: font({
          bold: true,
          size: 10,
          color:
            pts === 100 ? "FF16A34A" : pts === 50 ? "FFD97706" : "FFDC2626",
        }),
        alignment: alignCenter(),
        border: thinBorder(),
      });

      // Remark cell
      const remCell = ws.getCell(r, 7);
      remCell.value = row.remark || "—";
      applyStyle(remCell, {
        fill: fill(bg),
        font: font({ size: 9, italic: !row.remark, color: C.GREY }),
        alignment: alignLeft(),
        border: thinBorder(),
      });

      ws.getRow(r).height = 24;
      rowIdx++;
      r++;
    }
  }
}

// ── Sheet 4: Analysis ─────────────────────────────────────────────────────────
function buildAnalysisSheet(wb, sub) {
  const ws = wb.addWorksheet("📊 Analysis", {
    views: [{ showGridLines: false }],
  });

  // Column widths
  [36, 13, 11, 11, 9, 9].forEach((w, i) => {
    ws.getColumn(i + 1).width = w;
  });

  const sections = sub.sections ?? [];

  titleRow(
    ws,
    1,
    1,
    6,
    "📊  VISUAL ANALYSIS — SECTION SCORES & STATUS BREAKDOWN",
    C.DARK_BLUE,
    C.WHITE,
    13,
    true,
    32,
  );

  // Section Scores table
  titleRow(ws, 3, 1, 6, "SECTION SCORES", C.MID_BLUE, C.WHITE, 10, true, 22);
  headerRow(
    ws,
    4,
    ["Section", "Score %", "Earned", "Max", "Pass", "Fail"],
    [36, 13, 11, 11, 9, 9],
    C.MID_BLUE,
  );

  sections.forEach((sec, i) => {
    const sc = scoreColor(sec.sectionScore ?? 0);
    const bg = i % 2 === 0 ? C.ROW_EVEN : C.ROW_ODD;
    const row = 5 + i;

    dataCell(ws, row, 1, sec.sectionTitle ?? "", { bg });

    const scoreC = ws.getCell(row, 2);
    scoreC.value = `${sec.sectionScore ?? 0}%`;
    applyStyle(scoreC, {
      fill: fill(bg),
      font: font({ bold: true, size: 11, color: sc }),
      alignment: alignCenter(),
      border: thinBorder(),
    });

    dataCell(ws, row, 3, sec.earnedPoints ?? 0, { bg, align: "center" });
    dataCell(ws, row, 4, sec.maxPoints ?? 0, { bg, align: "center" });

    const passC = ws.getCell(row, 5);
    passC.value = sec.passCount ?? 0;
    applyStyle(passC, {
      fill: fill(bg),
      font: font({ bold: true, size: 10, color: "FF16A34A" }),
      alignment: alignCenter(),
      border: thinBorder(),
    });

    const failC = ws.getCell(row, 6);
    failC.value = sec.failCount ?? 0;
    applyStyle(failC, {
      fill: fill(bg),
      font: font({
        bold: true,
        size: 10,
        color: (sec.failCount ?? 0) > 0 ? "FFDC2626" : C.GREY,
      }),
      alignment: alignCenter(),
      border: thinBorder(),
    });

    ws.getRow(row).height = 22;
  });

  // Status Breakdown table
  const ssRow = 5 + sections.length + 2;
  titleRow(
    ws,
    ssRow,
    1,
    6,
    "STATUS BREAKDOWN",
    C.MID_BLUE,
    C.WHITE,
    10,
    true,
    22,
  );
  headerRow(
    ws,
    ssRow + 1,
    ["Status", "Count", "% of Total", "", "", ""],
    [18, 13, 17, 10, 10, 10],
    C.MID_BLUE,
  );

  const totalItems = sub.totalItems || 1;
  const statusData = [
    {
      label: "✅ Pass",
      count: sub.totalPass ?? 0,
      color: "FF16A34A",
      bg: C.PASS_BG,
    },
    {
      label: "⚡ Partial",
      count: sub.totalPartial ?? 0,
      color: "FFD97706",
      bg: C.PARTIAL_BG,
    },
    {
      label: "❌ Fail",
      count: sub.totalFail ?? 0,
      color: "FFDC2626",
      bg: C.FAIL_BG,
    },
    { label: "➖ N.A.", count: sub.totalNA ?? 0, color: C.GREY, bg: C.NA_BG },
  ];

  statusData.forEach((s, i) => {
    const row = ssRow + 2 + i;
    const pct = `${((s.count / totalItems) * 100).toFixed(1)}%`;

    const lc = ws.getCell(row, 1);
    lc.value = s.label;
    applyStyle(lc, {
      fill: fill(s.bg),
      font: font({ bold: true, size: 10, color: s.color }),
      alignment: alignLeft(),
      border: thinBorder(),
    });

    const cc = ws.getCell(row, 2);
    cc.value = s.count;
    applyStyle(cc, {
      fill: fill(s.bg),
      font: font({ bold: true, size: 12, color: s.color }),
      alignment: alignCenter(),
      border: thinBorder(),
    });

    const pc = ws.getCell(row, 3);
    pc.value = pct;
    applyStyle(pc, {
      fill: fill(s.bg),
      font: font({ size: 10, color: s.color }),
      alignment: alignCenter(),
      border: thinBorder(),
    });

    ws.getRow(row).height = 26;
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
//  OVERALL EXCEL
// ═══════════════════════════════════════════════════════════════════════════════

export async function generateOverallExcel(submissions) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Yakka Puka SLA Audit";
  wb.created = new Date();

  buildOverallSummarySheet(wb, submissions);
  buildAllSubmissionsSheet(wb, submissions);
  buildRatingBreakdownSheet(wb, submissions);
  buildScoreDistributionSheet(wb, submissions);

  const buf = await wb.xlsx.writeBuffer();
  return buf;
}

// ── Overall Sheet 1: KPI Dashboard ───────────────────────────────────────────
function buildOverallSummarySheet(wb, submissions) {
  const ws = wb.addWorksheet("🏆 Overall Summary", {
    views: [{ showGridLines: false }],
  });

  const total = submissions.length;
  const avg = total
    ? Math.round(
        submissions.reduce((a, s) => a + (s.overallScore ?? 0), 0) / total,
      )
    : 0;
  const excellent = submissions.filter(
    (s) => (s.overallScore ?? 0) >= 95,
  ).length;
  const good = submissions.filter(
    (s) => (s.overallScore ?? 0) >= 85 && (s.overallScore ?? 0) < 95,
  ).length;
  const average = submissions.filter(
    (s) => (s.overallScore ?? 0) >= 70 && (s.overallScore ?? 0) < 85,
  ).length;
  const critical = submissions.filter((s) => (s.overallScore ?? 0) < 70).length;
  const totalPass = submissions.reduce((a, s) => a + (s.totalPass ?? 0), 0);
  const totalPartial = submissions.reduce(
    (a, s) => a + (s.totalPartial ?? 0),
    0,
  );
  const totalFail = submissions.reduce((a, s) => a + (s.totalFail ?? 0), 0);
  const totalNA = submissions.reduce((a, s) => a + (s.totalNA ?? 0), 0);
  const sc = scoreColor(avg);

  // Column widths: 10 equal columns
  for (let i = 1; i <= 10; i++) ws.getColumn(i).width = 17;

  let r = 1;
  titleRow(
    ws,
    r,
    1,
    10,
    "🏢  YAKKA PUKA — SLA AUDIT OVERALL DASHBOARD",
    C.DARK_BLUE,
    C.WHITE,
    16,
    true,
    44,
  );
  r++;
  titleRow(
    ws,
    r,
    1,
    10,
    `Report generated: ${new Date().toLocaleString("en-IN")}  |  Total Records: ${total}`,
    C.MID_BLUE,
    "FFA0C4E8",
    9,
    false,
    20,
  );

  // KPIs
  r += 2;
  titleRow(
    ws,
    r,
    1,
    10,
    "📌  KEY PERFORMANCE INDICATORS",
    C.MID_BLUE,
    C.WHITE,
    11,
    true,
    24,
  );
  r++;

  const kpis = [
    { l: "Total Audits", v: total, c: C.ACCENT },
    { l: "Average Score", v: `${avg}%`, c: sc },
    { l: "⭐ Excellent", v: excellent, c: C.GREEN },
    { l: "✅ Good", v: good, c: "FF84CC16" },
    { l: "📊 Average", v: average, c: C.YELLOW },
    { l: "⚠️ Critical", v: critical, c: C.RED },
    { l: "Total Pass", v: totalPass, c: C.GREEN },
    { l: "Total Partial", v: totalPartial, c: C.YELLOW },
    { l: "Total Fail", v: totalFail, c: C.RED },
    { l: "Total N.A.", v: totalNA, c: C.GREY },
  ];
  kpis.forEach((k, i) => kpiBlock(ws, r, r + 1, i + 1, k.l, k.v, k.c));
  ws.getRow(r).height = 22;
  ws.getRow(r + 1).height = 34;

  // Score Distribution
  r += 3;
  titleRow(
    ws,
    r,
    1,
    10,
    "📊  SCORE DISTRIBUTION",
    C.MID_BLUE,
    C.WHITE,
    11,
    true,
    24,
  );
  r++;

  const dist = [
    { l: "Excellent\n(≥ 95%)", v: excellent, c: C.GREEN, bg: C.PASS_BG },
    { l: "Good\n(85–94%)", v: good, c: "FF84CC16", bg: "FFE8FAD0" },
    { l: "Average\n(70–84%)", v: average, c: C.YELLOW, bg: C.PARTIAL_BG },
    { l: "Critical\n(< 70%)", v: critical, c: C.RED, bg: C.FAIL_BG },
  ];

  dist.forEach((d, i) => {
    const col = i * 2 + 1;
    ws.mergeCells(r, col, r, col + 1);
    ws.mergeCells(r + 1, col, r + 1, col + 1);
    ws.mergeCells(r + 2, col, r + 2, col + 1);

    const lc = ws.getCell(r, col);
    lc.value = d.l;
    applyStyle(lc, {
      fill: fill(d.bg),
      font: font({ bold: true, size: 11, color: d.c }),
      alignment: alignCenter(),
      border: thinBorder(),
    });

    const vc = ws.getCell(r + 1, col);
    vc.value = d.v;
    applyStyle(vc, {
      fill: fill(C.WHITE),
      font: font({ bold: true, size: 30, color: d.c }),
      alignment: alignCenter(),
      border: thinBorder(),
    });

    const pc = ws.getCell(r + 2, col);
    pc.value = total ? `${Math.round((d.v / total) * 100)}%` : "0%";
    applyStyle(pc, {
      fill: fill(d.bg),
      font: font({ size: 11, color: d.c }),
      alignment: alignCenter(),
      border: thinBorder(),
    });
  });

  ws.getRow(r).height = 30;
  ws.getRow(r + 1).height = 46;
  ws.getRow(r + 2).height = 24;
}

// ── Overall Sheet 2: All Submissions ─────────────────────────────────────────
function buildAllSubmissionsSheet(wb, submissions) {
  const ws = wb.addWorksheet("📋 All Submissions", {
    views: [{ showGridLines: false }],
  });

  const headers = [
    "#",
    "Auditor Name",
    "Auditor Email",
    "Customer Name",
    "Mobile",
    "Company",
    "Submitted At",
    "Score %",
    "Rating",
    "Pass",
    "Partial",
    "Fail",
    "N.A.",
    "Earned",
    "Max",
    "Pass Rate",
  ];
  const widths = [5, 20, 26, 20, 16, 20, 22, 10, 18, 8, 9, 8, 8, 9, 8, 10];

  let r = 1;
  titleRow(
    ws,
    r,
    1,
    headers.length,
    "📋  ALL AUDIT SUBMISSIONS",
    C.DARK_BLUE,
    C.WHITE,
    13,
    true,
    32,
  );
  r++;
  headerRow(ws, r, headers, widths);
  r++;

  for (const [idx, sub] of submissions.entries()) {
    const cf = sub.customerFeedback || {};
    const score = sub.overallScore ?? 0;
    const sc = scoreColor(score);
    const bg = idx % 2 === 0 ? C.ROW_EVEN : C.ROW_ODD;

    const rowData = [
      idx + 1,
      sub.auditorName ?? "",
      sub.auditorEmail ?? "",
      cf.name ?? "",
      cf.mobile ?? "",
      cf.company ?? "",
      sub.submittedAt ? new Date(sub.submittedAt).toLocaleString("en-IN") : "",
      `${score}%`,
      sub.ratingLabel ?? "",
      sub.totalPass ?? 0,
      sub.totalPartial ?? 0,
      sub.totalFail ?? 0,
      sub.totalNA ?? 0,
      sub.earnedMarks ?? 0,
      sub.maxMarks ?? 0,
      `${sub.passRate ?? 0}%`,
    ];

    rowData.forEach((val, ci) => {
      const c = ws.getCell(r, ci + 1);
      c.value = val;
      let fnt = font({ size: 10 });
      if (ci === 1) fnt = font({ bold: true, size: 10, color: C.DARK_BLUE });
      if (ci === 7) fnt = font({ bold: true, size: 10, color: sc });
      if (ci === 8) fnt = font({ bold: true, size: 9, color: sc });
      if (ci === 9) fnt = font({ bold: true, size: 10, color: "FF16A34A" });
      if (ci === 11)
        fnt = font({
          bold: true,
          size: 10,
          color: (sub.totalFail ?? 0) > 0 ? "FFDC2626" : C.GREY,
        });
      applyStyle(c, {
        fill: fill(bg),
        font: fnt,
        alignment: ci === 1 ? alignLeft() : alignCenter(),
        border: thinBorder(),
      });
    });
    ws.getRow(r).height = 22;
    r++;
  }

  // Totals
  const totalCount = submissions.length;
  const avgScore = totalCount
    ? Math.round(
        submissions.reduce((a, s) => a + (s.overallScore ?? 0), 0) / totalCount,
      )
    : 0;
  const totalsRow = [
    "TOTAL",
    "",
    "",
    "",
    "",
    "",
    `${totalCount} records`,
    `${avgScore}%`,
    "",
    submissions.reduce((a, s) => a + (s.totalPass ?? 0), 0),
    submissions.reduce((a, s) => a + (s.totalPartial ?? 0), 0),
    submissions.reduce((a, s) => a + (s.totalFail ?? 0), 0),
    submissions.reduce((a, s) => a + (s.totalNA ?? 0), 0),
    submissions.reduce((a, s) => a + (s.earnedMarks ?? 0), 0),
    submissions.reduce((a, s) => a + (s.maxMarks ?? 0), 0),
    "",
  ];
  totalsRow.forEach((val, ci) => {
    const c = ws.getCell(r, ci + 1);
    c.value = val;
    applyStyle(c, {
      fill: fill(C.MID_BLUE),
      font: font({ bold: true, size: 10, color: C.WHITE }),
      alignment: alignCenter(),
      border: medBorder(),
    });
  });
  ws.getRow(r).height = 26;
}

// ── Overall Sheet 3: Rating Breakdown ────────────────────────────────────────
function buildRatingBreakdownSheet(wb, submissions) {
  const ws = wb.addWorksheet("🏅 Rating Breakdown", {
    views: [{ showGridLines: false }],
  });

  titleRow(
    ws,
    1,
    1,
    5,
    "🏅  RATING BREAKDOWN",
    C.DARK_BLUE,
    C.WHITE,
    13,
    true,
    32,
  );
  headerRow(
    ws,
    2,
    ["Rating", "Count", "% of Total", "Avg Score", "Score Bar"],
    [26, 12, 16, 14, 36],
  );

  const ratingMap = {};
  for (const s of submissions) {
    const label = s.ratingLabel ?? "Unknown";
    if (!ratingMap[label]) ratingMap[label] = [];
    ratingMap[label].push(s);
  }

  const total = submissions.length;
  let r = 3;
  Object.entries(ratingMap)
    .sort((a, b) => b[1].length - a[1].length)
    .forEach(([rating, subs], idx) => {
      const avg = Math.round(
        subs.reduce((a, s) => a + (s.overallScore ?? 0), 0) / subs.length,
      );
      const sc = scoreColor(avg);
      const pct = total ? `${((subs.length / total) * 100).toFixed(1)}%` : "0%";
      const bg = idx % 2 === 0 ? C.ROW_EVEN : C.ROW_ODD;
      const bar = "█".repeat(Math.max(1, Math.round(avg / 5)));

      dataCell(ws, r, 1, rating, { bg, align: "center" });
      dataCell(ws, r, 2, subs.length, { bg, align: "center" });
      dataCell(ws, r, 3, pct, { bg, align: "center" });

      const avgCell = ws.getCell(r, 4);
      avgCell.value = `${avg}%`;
      applyStyle(avgCell, {
        fill: fill(bg),
        font: font({ bold: true, size: 10, color: sc }),
        alignment: alignCenter(),
        border: thinBorder(),
      });

      const barCell = ws.getCell(r, 5);
      barCell.value = bar;
      applyStyle(barCell, {
        fill: fill(bg),
        font: font({ size: 9, color: sc }),
        alignment: alignLeft(),
        border: thinBorder(),
      });

      ws.getRow(r).height = 24;
      r++;
    });
}

// ── Overall Sheet 4: Score Distribution ──────────────────────────────────────
function buildScoreDistributionSheet(wb, submissions) {
  const ws = wb.addWorksheet("📊 Score Distribution", {
    views: [{ showGridLines: false }],
  });

  [20, 13, 17, 36, 10, 10].forEach((w, i) => {
    ws.getColumn(i + 1).width = w;
  });

  titleRow(
    ws,
    1,
    1,
    6,
    "📊  SCORE DISTRIBUTION & STATUS ANALYSIS",
    C.DARK_BLUE,
    C.WHITE,
    13,
    true,
    32,
  );

  // Score Range
  titleRow(
    ws,
    3,
    1,
    6,
    "SCORE RANGE BREAKDOWN",
    C.MID_BLUE,
    C.WHITE,
    10,
    true,
    22,
  );
  headerRow(
    ws,
    4,
    ["Score Range", "Count", "% of Total", "Visual Bar", "", ""],
    [20, 13, 17, 36, 10, 10],
  );

  const total = submissions.length || 1;
  const ranges = [
    {
      label: "⭐ 95–100%",
      count: submissions.filter((s) => (s.overallScore ?? 0) >= 95).length,
      color: C.GREEN,
      bg: C.PASS_BG,
    },
    {
      label: "✅ 85–94%",
      count: submissions.filter(
        (s) => (s.overallScore ?? 0) >= 85 && (s.overallScore ?? 0) < 95,
      ).length,
      color: "FF84CC16",
      bg: "FFE8FAD0",
    },
    {
      label: "📊 70–84%",
      count: submissions.filter(
        (s) => (s.overallScore ?? 0) >= 70 && (s.overallScore ?? 0) < 85,
      ).length,
      color: C.YELLOW,
      bg: C.PARTIAL_BG,
    },
    {
      label: "⚠️ 50–69%",
      count: submissions.filter(
        (s) => (s.overallScore ?? 0) >= 50 && (s.overallScore ?? 0) < 70,
      ).length,
      color: "FFEA580C",
      bg: "FFFFF0E0",
    },
    {
      label: "❌ < 50%",
      count: submissions.filter((s) => (s.overallScore ?? 0) < 50).length,
      color: C.RED,
      bg: C.FAIL_BG,
    },
  ];

  ranges.forEach((range, i) => {
    const row = 5 + i;
    const pct = `${((range.count / total) * 100).toFixed(1)}%`;
    const bar =
      "█".repeat(Math.max(0, Math.round((range.count / total) * 20))) || "—";

    const lc = ws.getCell(row, 1);
    lc.value = range.label;
    applyStyle(lc, {
      fill: fill(range.bg),
      font: font({ bold: true, size: 10, color: range.color }),
      alignment: alignLeft(),
      border: thinBorder(),
    });

    const cc = ws.getCell(row, 2);
    cc.value = range.count;
    applyStyle(cc, {
      fill: fill(range.bg),
      font: font({ bold: true, size: 12, color: range.color }),
      alignment: alignCenter(),
      border: thinBorder(),
    });

    const pc = ws.getCell(row, 3);
    pc.value = pct;
    applyStyle(pc, {
      fill: fill(range.bg),
      font: font({ size: 10, color: range.color }),
      alignment: alignCenter(),
      border: thinBorder(),
    });

    const bc = ws.getCell(row, 4);
    bc.value = bar;
    applyStyle(bc, {
      fill: fill(range.bg),
      font: font({ size: 9, color: range.color }),
      alignment: alignLeft(),
      border: thinBorder(),
    });

    ws.getRow(row).height = 26;
  });

  // Overall Status Totals
  const ssRow = 12;
  titleRow(
    ws,
    ssRow,
    1,
    6,
    "OVERALL STATUS TOTALS",
    C.MID_BLUE,
    C.WHITE,
    10,
    true,
    22,
  );
  headerRow(
    ws,
    ssRow + 1,
    ["Status", "Total Count", "% of All Items", "Visual Bar", "", ""],
    [20, 16, 18, 36, 10, 10],
  );

  const totalItems =
    submissions.reduce((a, s) => a + (s.totalItems ?? 0), 0) || 1;
  const statusData = [
    {
      label: "✅ Pass",
      count: submissions.reduce((a, s) => a + (s.totalPass ?? 0), 0),
      color: "FF16A34A",
      bg: C.PASS_BG,
    },
    {
      label: "⚡ Partial",
      count: submissions.reduce((a, s) => a + (s.totalPartial ?? 0), 0),
      color: "FFD97706",
      bg: C.PARTIAL_BG,
    },
    {
      label: "❌ Fail",
      count: submissions.reduce((a, s) => a + (s.totalFail ?? 0), 0),
      color: "FFDC2626",
      bg: C.FAIL_BG,
    },
    {
      label: "➖ N.A.",
      count: submissions.reduce((a, s) => a + (s.totalNA ?? 0), 0),
      color: C.GREY,
      bg: C.NA_BG,
    },
  ];

  statusData.forEach((s, i) => {
    const row = ssRow + 2 + i;
    const pct = `${((s.count / totalItems) * 100).toFixed(1)}%`;
    const bar =
      "█".repeat(Math.max(0, Math.round((s.count / totalItems) * 20))) || "—";

    const lc = ws.getCell(row, 1);
    lc.value = s.label;
    applyStyle(lc, {
      fill: fill(s.bg),
      font: font({ bold: true, size: 10, color: s.color }),
      alignment: alignLeft(),
      border: thinBorder(),
    });

    const cc = ws.getCell(row, 2);
    cc.value = s.count;
    applyStyle(cc, {
      fill: fill(s.bg),
      font: font({ bold: true, size: 12, color: s.color }),
      alignment: alignCenter(),
      border: thinBorder(),
    });

    const pc = ws.getCell(row, 3);
    pc.value = pct;
    applyStyle(pc, {
      fill: fill(s.bg),
      font: font({ size: 10, color: s.color }),
      alignment: alignCenter(),
      border: thinBorder(),
    });

    const bc = ws.getCell(row, 4);
    bc.value = bar;
    applyStyle(bc, {
      fill: fill(s.bg),
      font: font({ size: 9, color: s.color }),
      alignment: alignLeft(),
      border: thinBorder(),
    });

    ws.getRow(row).height = 26;
  });
}
