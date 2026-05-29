// controllers/PptxExportController.js
// POST /audit/export/dashboard-pptx
// 10-slide comprehensive PPTX — mirrors exact dashboard layout

import PptxGenJS from "pptxgenjs";
import AuditSubmission from "../models/Auditsubmissionmodel.js";

const TOTAL = 10; // total slides

// ─── color helpers ────────────────────────────────────────────────────────────
function gradeColor(score) {
  if (score >= 95)
    return {
      hex: "16a34a",
      light: "dcfce7",
      label: "Excellent",
      ring: "22c55e",
    };
  if (score >= 85)
    return { hex: "2563eb", light: "dbeafe", label: "Good", ring: "3b82f6" };
  if (score >= 75)
    return {
      hex: "d97706",
      light: "fef3c7",
      label: "Needs Improvement",
      ring: "f59e0b",
    };
  if (score >= 60)
    return { hex: "ea580c", light: "ffedd5", label: "At Risk", ring: "f97316" };
  if (score >= 40)
    return {
      hex: "dc2626",
      light: "fee2e2",
      label: "Critical",
      ring: "ef4444",
    };
  return { hex: "7f1d1d", light: "fee2e2", label: "Severe", ring: "991b1b" };
}

function failRateColor(rate) {
  if (rate === 0) return { hex: "16a34a", label: "No Failures" };
  if (rate <= 5) return { hex: "2563eb", label: "Low" };
  if (rate <= 15) return { hex: "d97706", label: "Moderate" };
  if (rate <= 30) return { hex: "ea580c", label: "High" };
  if (rate <= 50) return { hex: "dc2626", label: "Critical" };
  return { hex: "7f1d1d", label: "Severe" };
}

// ─── shared layout helpers ────────────────────────────────────────────────────
function addFooter(slide, pres, num) {
  const now = new Date().toLocaleDateString("en-IN", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0,
    y: 5.36,
    w: 10,
    h: 0.02,
    fill: { color: "E2E8F0" },
  });
  slide.addText(`Yakka Pukka · SLA Dashboard · ${now}`, {
    x: 0.3,
    y: 5.38,
    w: 8.5,
    h: 0.22,
    fontSize: 8,
    color: "94A3B8",
    fontFace: "Calibri",
    margin: 0,
  });
  slide.addText(`${num} / ${TOTAL}`, {
    x: 9.0,
    y: 5.38,
    w: 0.8,
    h: 0.22,
    fontSize: 8,
    color: "94A3B8",
    align: "right",
    fontFace: "Calibri",
    margin: 0,
  });
}

function addSlideHeader(slide, pres, accentColor, title, subtitle) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0,
    y: 0,
    w: 10,
    h: 0.08,
    fill: { color: accentColor },
  });
  slide.addText(title, {
    x: 0.4,
    y: 0.16,
    w: 9.2,
    h: 0.5,
    fontSize: 20,
    bold: true,
    color: "0F172A",
    fontFace: "Calibri",
    margin: 0,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.4,
      y: 0.68,
      w: 9.2,
      h: 0.26,
      fontSize: 10,
      color: "94A3B8",
      fontFace: "Calibri",
      margin: 0,
    });
  }
}

function kpiTile(slide, pres, x, y, w, h, val, label, color) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x,
    y,
    w,
    h,
    fill: { color: color + "1A" },
    line: { color: color, width: 1 },
  });
  slide.addText(val, {
    x,
    y: y + 0.05,
    w,
    h: h * 0.56,
    fontSize: 18,
    bold: true,
    color,
    align: "center",
    fontFace: "Calibri",
    margin: 0,
  });
  slide.addText(label, {
    x,
    y: y + h * 0.62,
    w,
    h: h * 0.36,
    fontSize: 9,
    color: "64748B",
    align: "center",
    fontFace: "Calibri",
    margin: 0,
  });
}

// ─── data helpers ─────────────────────────────────────────────────────────────
const MONTHS_ARR = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
];

function bucketByMonth(audits, field) {
  const map = {};
  audits.forEach((a) => {
    const d = new Date(a.submittedAt);
    const k = `${MONTHS_ARR[d.getMonth()]} ${d.getFullYear()}`;
    if (!map[k]) map[k] = [];
    const v = field === "totalFail" ? a.totalFail || 0 : a[field] || 0;
    map[k].push(v);
  });
  const sorted = Object.entries(map).sort(([ka], [kb]) => {
    const [ma, ya] = ka.split(" ");
    const [mb, yb] = kb.split(" ");
    return (
      new Date(+ya, MONTHS_ARR.indexOf(ma)) -
      new Date(+yb, MONTHS_ARR.indexOf(mb))
    );
  });
  return {
    labels: sorted.map(([k]) => k),
    values: sorted.map(([, v]) =>
      Math.round(v.reduce((s, x) => s + x, 0) / v.length),
    ),
  };
}

function topFailItems(audits, n) {
  const map = {};
  audits.forEach((a) => {
    (a.sections || []).forEach((sec) => {
      (sec.rows || []).forEach((row) => {
        if (row.status === "Fail") {
          const k = (row.itemText || "").trim() || "Unknown";
          if (!map[k]) map[k] = { count: 0, section: sec.sectionTitle || "" };
          map[k].count++;
        }
      });
    });
  });
  return Object.entries(map)
    .map(([text, d]) => ({ text, count: d.count, section: d.section }))
    .sort((a, b) => b.count - a.count)
    .slice(0, n);
}

function sectionStats(audits) {
  const map = {};
  audits.forEach((a) => {
    (a.sections || []).forEach((sec) => {
      if (!map[sec.sectionTitle])
        map[sec.sectionTitle] = {
          scores: [],
          pass: 0,
          fail: 0,
          partial: 0,
          na: 0,
        };
      const sm = map[sec.sectionTitle];
      sm.scores.push(sec.sectionScore || 0);
      sm.pass += sec.passCount || 0;
      sm.fail += sec.failCount || 0;
      sm.partial += sec.partialCount || 0;
      sm.na += sec.naCount || 0;
    });
  });
  return Object.entries(map)
    .map(([title, d]) => ({
      title,
      avg: Math.round(d.scores.reduce((s, x) => s + x, 0) / d.scores.length),
      min: Math.round(Math.min(...d.scores)),
      max: Math.round(Math.max(...d.scores)),
      pass: d.pass,
      fail: d.fail,
      partial: d.partial,
    }))
    .sort((a, b) => b.avg - a.avg);
}

function ratingDist(audits) {
  return [
    { label: "Excellent (95-100%)", min: 95, max: 100, color: "22c55e" },
    { label: "Good (85-94%)", min: 85, max: 94, color: "3b82f6" },
    { label: "Needs Impr (75-84%)", min: 75, max: 84, color: "f59e0b" },
    { label: "At Risk (60-74%)", min: 60, max: 74, color: "ea580c" },
    { label: "Critical (<60%)", min: 0, max: 59, color: "dc2626" },
  ].map((r) => ({
    ...r,
    count: audits.filter(
      (a) => (a.overallScore || 0) >= r.min && (a.overallScore || 0) <= r.max,
    ).length,
  }));
}

function complianceStat(audits) {
  const totalPass = audits.reduce((s, a) => s + (a.totalPass || 0), 0);
  const totalFail = audits.reduce((s, a) => s + (a.totalFail || 0), 0);
  const totalPartial = audits.reduce((s, a) => s + (a.totalPartial || 0), 0);
  const totalNA = audits.reduce((s, a) => s + (a.totalNA || 0), 0);
  const totalItems = audits.reduce((s, a) => s + (a.totalItems || 0), 0);
  const scored = totalItems - totalNA;
  return {
    totalPass,
    totalFail,
    totalPartial,
    totalNA,
    totalItems,
    pct: scored > 0 ? Math.round((totalPass / scored) * 100) : 0,
  };
}

// ─── SLIDE 1 — Cover ──────────────────────────────────────────────────────────
function makeSlide1(pres, audits, avgScore) {
  const g = gradeColor(avgScore);
  const cs = complianceStat(audits);
  const slide = pres.addSlide();
  slide.background = { color: "0F172A" };

  // left accent
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0,
    y: 0,
    w: 0.22,
    h: 5.625,
    fill: { color: g.hex },
  });

  slide.addText("YAKKA PUKKA", {
    x: 0.5,
    y: 0.45,
    w: 9,
    h: 0.38,
    fontSize: 11,
    bold: true,
    color: "94A3B8",
    charSpacing: 7,
    fontFace: "Calibri",
    margin: 0,
  });
  slide.addText("Admin Dashboard Report", {
    x: 0.5,
    y: 0.9,
    w: 9,
    h: 0.85,
    fontSize: 38,
    bold: true,
    color: "FFFFFF",
    fontFace: "Calibri",
    margin: 0,
  });
  slide.addText(
    "Complete SLA Audit Performance — Charts, Analytics & Critical Insights",
    {
      x: 0.5,
      y: 1.82,
      w: 9,
      h: 0.38,
      fontSize: 13,
      color: "94A3B8",
      fontFace: "Calibri",
      margin: 0,
    },
  );
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5,
    y: 2.32,
    w: 3.8,
    h: 0.04,
    fill: { color: g.hex },
  });

  // 4 KPI boxes
  const kpis = [
    { v: String(audits.length), l: "Total Audits", c: "3b82f6" },
    { v: `${avgScore}%`, l: "Avg Score", c: g.hex },
    { v: String(cs.totalPass), l: "Total Pass", c: "22c55e" },
    { v: String(cs.totalFail), l: "Total Fail", c: "ef4444" },
  ];
  kpis.forEach((k, i) => {
    const x = 0.5 + i * 2.35;
    slide.addShape(pres.shapes.RECTANGLE, {
      x,
      y: 2.65,
      w: 2.18,
      h: 1.1,
      fill: { color: "1E293B" },
      line: { color: k.c + "60", width: 1 },
    });
    slide.addText(k.v, {
      x,
      y: 2.7,
      w: 2.18,
      h: 0.62,
      fontSize: 26,
      bold: true,
      color: k.c,
      align: "center",
      fontFace: "Calibri",
      margin: 0,
    });
    slide.addText(k.l, {
      x,
      y: 3.32,
      w: 2.18,
      h: 0.3,
      fontSize: 9,
      color: "64748B",
      align: "center",
      fontFace: "Calibri",
      margin: 0,
    });
  });

  // status badge
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5,
    y: 4.05,
    w: 2.3,
    h: 0.45,
    fill: { color: g.hex + "28" },
    line: { color: g.hex, width: 1 },
  });
  slide.addText(g.label, {
    x: 0.5,
    y: 4.07,
    w: 2.3,
    h: 0.41,
    fontSize: 13,
    bold: true,
    color: g.hex,
    align: "center",
    fontFace: "Calibri",
    margin: 0,
  });
  slide.addText(`Partial: ${cs.totalPartial} · Compliance: ${cs.pct}%`, {
    x: 3.0,
    y: 4.1,
    w: 4,
    h: 0.35,
    fontSize: 11,
    color: "d97706",
    fontFace: "Calibri",
    margin: 0,
  });

  const now = new Date().toLocaleDateString("en-IN", {
    day: "2-digit",
    month: "long",
    year: "numeric",
  });
  slide.addText(`Generated on ${now}`, {
    x: 0.5,
    y: 5.0,
    w: 9,
    h: 0.28,
    fontSize: 9.5,
    color: "475569",
    fontFace: "Calibri",
    margin: 0,
  });
  addFooter(slide, pres, 1);
}

// ─── SLIDE 2 — Overall Avg Score ──────────────────────────────────────────────
function makeSlide2(pres, audits) {
  const { labels, values } = bucketByMonth(audits, "overallScore");
  const avg = values.length
    ? Math.round(values.reduce((s, x) => s + x, 0) / values.length)
    : 0;
  const best = values.length ? Math.max(...values) : 0;
  const worst = values.length ? Math.min(...values) : 0;
  const g = gradeColor(avg);

  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };
  addSlideHeader(
    slide,
    pres,
    g.hex,
    "📊  Overall Average Score",
    "Monthly performance trend · all audit submissions",
  );

  // 3 KPI tiles top right
  [
    { v: `${best}%`, l: "Best", c: gradeColor(best).hex },
    { v: `${avg}%`, l: "Avg", c: g.hex },
    { v: `${worst}%`, l: "Lowest", c: gradeColor(worst).hex },
  ].forEach((k, i) =>
    kpiTile(slide, pres, 7.1 + i * 0.9, 0.1, 0.82, 0.9, k.v, k.l, k.c),
  );

  if (values.length > 0) {
    // Line chart left
    slide.addChart(pres.charts.LINE, [{ name: "Avg Score", labels, values }], {
      x: 0.3,
      y: 1.07,
      w: 6.1,
      h: 3.25,
      chartColors: [g.hex],
      lineSize: 3,
      lineSmooth: true,
      chartArea: { fill: { color: "FFFFFF" }, roundedCorners: true },
      catAxisLabelColor: "94A3B8",
      valAxisLabelColor: "94A3B8",
      valGridLine: { color: "E2E8F0", size: 0.5 },
      catGridLine: { style: "none" },
      valAxisMinVal: 0,
      valAxisMaxVal: 100,
      showValue: true,
      dataLabelColor: g.hex,
      showLegend: false,
    });
    // Bar chart right
    slide.addChart(pres.charts.BAR, [{ name: "Score", labels, values }], {
      x: 6.6,
      y: 1.07,
      w: 3.1,
      h: 3.25,
      barDir: "col",
      chartColors: values.map((v) => gradeColor(v).hex),
      chartArea: { fill: { color: "F8FAFC" }, roundedCorners: true },
      catAxisLabelColor: "94A3B8",
      valAxisLabelColor: "94A3B8",
      valGridLine: { color: "E2E8F0", size: 0.5 },
      catGridLine: { style: "none" },
      valAxisMinVal: 0,
      valAxisMaxVal: 100,
      showValue: true,
      dataLabelColor: "334155",
      showLegend: false,
    });
  } else {
    slide.addText("No data available yet", {
      x: 2,
      y: 2.5,
      w: 6,
      h: 1,
      fontSize: 13,
      color: "94A3B8",
      align: "center",
    });
  }

  // Info strip
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.3,
    y: 4.45,
    w: 9.4,
    h: 0.62,
    fill: { color: g.hex + "0F" },
  });
  slide.addText(
    `Status: ${g.label}  ·  ${audits.length} audits  ·  Best: ${best}%  ·  Target: 95%  ·  Threshold: 75%`,
    {
      x: 0.5,
      y: 4.5,
      w: 9,
      h: 0.52,
      fontSize: 10,
      color: g.hex,
      fontFace: "Calibri",
      margin: 0,
    },
  );

  addFooter(slide, pres, 2);
}

// ─── SLIDE 3 — Critical Fail Items (trend + top bar) ─────────────────────────
function makeSlide3(pres, audits) {
  const totalFail = audits.reduce((s, a) => s + (a.totalFail || 0), 0);
  const totalItems = audits.reduce((s, a) => s + (a.totalItems || 0), 0);
  const failRate =
    totalItems > 0
      ? Math.min(Math.round((totalFail / totalItems) * 100), 100)
      : 0;
  const avgFails = audits.length ? +(totalFail / audits.length).toFixed(1) : 0;
  const g = failRateColor(failRate);
  const items = topFailItems(audits, 10);
  const { labels: mL, values: mV } = bucketByMonth(audits, "totalFail");

  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };
  addSlideHeader(
    slide,
    pres,
    g.hex,
    "⚠️  Critical Fail Items",
    `Failure trend · ${totalFail} total fails · ${failRate}% fail rate · avg ${avgFails}/audit`,
  );

  // 4 KPI tiles
  [
    { v: String(totalFail), l: "Total Fails", c: g.hex },
    { v: String(avgFails), l: "Avg / Audit", c: "ea580c" },
    { v: `${failRate}%`, l: "Fail Rate", c: g.hex },
    { v: String(items.length), l: "Unique Issues", c: "7f1d1d" },
  ].forEach((k, i) =>
    kpiTile(slide, pres, 0.3 + i * 2.38, 1.02, 2.16, 0.82, k.v, k.l, k.c),
  );

  // Monthly trend line (left)
  if (mL.length > 0) {
    slide.addChart(
      pres.charts.LINE,
      [{ name: "Avg Fails", labels: mL, values: mV }],
      {
        x: 0.3,
        y: 2.0,
        w: 4.6,
        h: 2.7,
        chartColors: [g.hex],
        lineSize: 2.5,
        lineSmooth: true,
        chartArea: { fill: { color: "FFFFFF" }, roundedCorners: true },
        catAxisLabelColor: "94A3B8",
        valAxisLabelColor: "94A3B8",
        valGridLine: { color: "E2E8F0", size: 0.5 },
        catGridLine: { style: "none" },
        showValue: true,
        dataLabelColor: g.hex,
        showLegend: false,
      },
    );
  }

  // Top items horizontal bar (right)
  if (items.length > 0) {
    slide.addChart(
      pres.charts.BAR,
      [
        {
          name: "Fail Count",
          labels: items.map((it) =>
            it.text.length > 28 ? it.text.slice(0, 28) + "…" : it.text,
          ),
          values: items.map((it) => it.count),
        },
      ],
      {
        x: 5.1,
        y: 2.0,
        w: 4.6,
        h: 2.7,
        barDir: "bar",
        chartColors: items.map((_, i) => {
          const p = [
            "7f1d1d",
            "991b1b",
            "b91c1c",
            "dc2626",
            "ef4444",
            "f87171",
            "ea580c",
            "f97316",
            "f59e0b",
            "d97706",
          ];
          return p[i] || "dc2626";
        }),
        chartArea: { fill: { color: "FFFFFF" }, roundedCorners: true },
        catAxisLabelColor: "64748B",
        valAxisLabelColor: "64748B",
        valGridLine: { color: "E2E8F0", size: 0.5 },
        catGridLine: { style: "none" },
        showValue: true,
        dataLabelColor: "1E293B",
        showLegend: false,
      },
    );
  }

  addFooter(slide, pres, 3);
}

// ─── SLIDE 4 — Top Critical Issues Table ─────────────────────────────────────
function makeSlide4(pres, audits) {
  const items = topFailItems(audits, 15);
  const maxCnt = items.length ? items[0].count : 1;

  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };
  addSlideHeader(
    slide,
    pres,
    "dc2626",
    "🔴  Top Critical Issues — Complete List",
    `${items.length} unique failure items ranked by frequency`,
  );

  if (!items.length) {
    slide.addText("🎉 No critical failures! Outstanding performance.", {
      x: 1.5,
      y: 2.5,
      w: 7,
      h: 1,
      fontSize: 16,
      color: "22c55e",
      align: "center",
    });
  } else {
    const header = [
      [
        {
          text: "#",
          options: {
            bold: true,
            fill: { color: "0F172A" },
            color: "FFFFFF",
            fontSize: 9,
          },
        },
        {
          text: "Item",
          options: {
            bold: true,
            fill: { color: "0F172A" },
            color: "FFFFFF",
            fontSize: 9,
          },
        },
        {
          text: "Section",
          options: {
            bold: true,
            fill: { color: "0F172A" },
            color: "FFFFFF",
            fontSize: 9,
          },
        },
        {
          text: "Count",
          options: {
            bold: true,
            fill: { color: "0F172A" },
            color: "FFFFFF",
            fontSize: 9,
          },
        },
        {
          text: "% Share",
          options: {
            bold: true,
            fill: { color: "0F172A" },
            color: "FFFFFF",
            fontSize: 9,
          },
        },
      ],
    ];
    const FAIL_PALETTE = [
      "7f1d1d",
      "991b1b",
      "b91c1c",
      "dc2626",
      "ef4444",
      "f87171",
      "ea580c",
      "f97316",
      "f59e0b",
      "d97706",
      "dc2626",
      "ef4444",
      "f97316",
      "f59e0b",
      "ea580c",
    ];
    const rows = items.map((item, i) => {
      const c = FAIL_PALETTE[i] || "dc2626";
      const pct = Math.round((item.count / maxCnt) * 100);
      return [
        {
          text: String(i + 1),
          options: { fontSize: 9, color: "64748B", align: "center" },
        },
        {
          text:
            item.text.length > 52 ? item.text.slice(0, 52) + "…" : item.text,
          options: { fontSize: 9, color: "1E293B" },
        },
        {
          text: (item.section || "—").slice(0, 26),
          options: { fontSize: 8.5, color: "64748B" },
        },
        {
          text: String(item.count) + "×",
          options: { fontSize: 9, bold: true, color: c, align: "center" },
        },
        {
          text: `${pct}%`,
          options: { fontSize: 9, color: c, align: "center" },
        },
      ];
    });
    slide.addTable([...header, ...rows], {
      x: 0.3,
      y: 1.05,
      w: 9.4,
      colW: [0.4, 4.6, 2.5, 0.75, 0.75],
      fontSize: 9,
      fontFace: "Calibri",
      border: { pt: 0.5, color: "E2E8F0" },
      rowH: 0.268,
    });
  }

  addFooter(slide, pres, 4);
}

// ─── SLIDE 5 — Monthly Target ─────────────────────────────────────────────────
function makeSlide5(pres, audits) {
  const { labels, values } = bucketByMonth(audits, "overallScore");
  const avg = values.length
    ? Math.round(values.reduce((s, x) => s + x, 0) / values.length)
    : 0;
  const best = values.length ? Math.max(...values) : 0;
  const g = gradeColor(avg);

  const now = new Date();
  const thisM = audits.filter((a) => {
    const d = new Date(a.submittedAt);
    return (
      d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear()
    );
  });
  const lMon = now.getMonth() === 0 ? 11 : now.getMonth() - 1;
  const lYr = now.getMonth() === 0 ? now.getFullYear() - 1 : now.getFullYear();
  const lastM = audits.filter((a) => {
    const d = new Date(a.submittedAt);
    return d.getMonth() === lMon && d.getFullYear() === lYr;
  });
  const thisAvg = thisM.length
    ? Math.round(thisM.reduce((s, a) => s + a.overallScore, 0) / thisM.length)
    : 0;
  const lastAvg = lastM.length
    ? Math.round(lastM.reduce((s, a) => s + a.overallScore, 0) / lastM.length)
    : 0;
  const diff = thisAvg - lastAvg;

  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };
  addSlideHeader(
    slide,
    pres,
    g.hex,
    "🎯  Monthly Target",
    "Overall score vs 95% excellence target",
  );

  // Big score box (left)
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.3,
    y: 1.05,
    w: 3.9,
    h: 3.05,
    fill: { color: g.hex + "0D" },
    line: { color: g.hex + "40", width: 1 },
  });
  slide.addText(`${avg}%`, {
    x: 0.3,
    y: 1.1,
    w: 3.9,
    h: 1.55,
    fontSize: 80,
    bold: true,
    color: g.hex,
    align: "center",
    fontFace: "Calibri",
    margin: 0,
  });
  slide.addText(g.label, {
    x: 0.3,
    y: 2.7,
    w: 3.9,
    h: 0.42,
    fontSize: 14,
    bold: true,
    color: g.hex,
    align: "center",
    fontFace: "Calibri",
    margin: 0,
  });

  // vs last month badge
  const tc = diff >= 0 ? "16a34a" : "dc2626";
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.65,
    y: 3.25,
    w: 3.2,
    h: 0.42,
    fill: { color: tc + "18" },
    line: { color: tc + "50", width: 1 },
  });
  slide.addText(`${diff >= 0 ? "▲ +" : "▼ "}${diff}% vs last month`, {
    x: 0.65,
    y: 3.27,
    w: 3.2,
    h: 0.38,
    fontSize: 11,
    bold: true,
    color: tc,
    align: "center",
    fontFace: "Calibri",
    margin: 0,
  });

  const msgs = {
    Excellent: "Outstanding performance! Keep it up.",
    Good: "Good performance. Push harder for Excellent!",
    "Needs Improvement": "Below target. Focus on failing checklist items.",
    "At Risk": "At risk! Immediate attention needed on failures.",
    Critical: "Critical! Immediate remediation required.",
    Severe: "Severe failure! Emergency action required.",
  };
  slide.addText(msgs[g.label] || `Best month: ${MONTHS_ARR[0]}`, {
    x: 0.3,
    y: 3.78,
    w: 3.9,
    h: 0.3,
    fontSize: 9,
    color: g.hex,
    align: "center",
    italic: true,
    fontFace: "Calibri",
    margin: 0,
  });

  // Line chart (right)
  if (values.length > 0) {
    slide.addChart(
      pres.charts.LINE,
      [{ name: "Monthly Score", labels, values }],
      {
        x: 4.45,
        y: 1.05,
        w: 5.2,
        h: 3.05,
        chartColors: [g.hex],
        lineSize: 2.5,
        lineSmooth: true,
        chartArea: { fill: { color: "FFFFFF" }, roundedCorners: true },
        catAxisLabelColor: "94A3B8",
        valAxisLabelColor: "94A3B8",
        valGridLine: { color: "E2E8F0", size: 0.5 },
        catGridLine: { style: "none" },
        valAxisMinVal: 0,
        valAxisMaxVal: 100,
        showValue: true,
        dataLabelColor: g.hex,
        showLegend: false,
      },
    );
  }

  // Stats row bottom
  const cs = complianceStat(audits);
  [
    { v: String(cs.totalPass), l: "Pass", c: "22c55e" },
    { v: String(cs.totalFail), l: "Fail", c: "ef4444" },
    { v: String(cs.totalPartial), l: "Partial", c: "d97706" },
    { v: `${best}%`, l: "Best", c: gradeColor(best).hex },
  ].forEach((k, i) => {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 4.45 + i * 1.3,
      y: 4.28,
      w: 1.18,
      h: 0.72,
      fill: { color: k.c + "14" },
      line: { color: k.c + "40", width: 1 },
    });
    slide.addText(k.v, {
      x: 4.45 + i * 1.3,
      y: 4.3,
      w: 1.18,
      h: 0.36,
      fontSize: 14,
      bold: true,
      color: k.c,
      align: "center",
      fontFace: "Calibri",
      margin: 0,
    });
    slide.addText(k.l, {
      x: 4.45 + i * 1.3,
      y: 4.66,
      w: 1.18,
      h: 0.28,
      fontSize: 8.5,
      color: "64748B",
      align: "center",
      fontFace: "Calibri",
      margin: 0,
    });
  });

  addFooter(slide, pres, 5);
}

// ─── SLIDE 6 — Overall Compliance ────────────────────────────────────────────
function makeSlide6(pres, audits) {
  const cs = complianceStat(audits);
  const g = gradeColor(cs.pct);
  const gap = cs.pct - 95;
  const { labels, values } = bucketByMonth(audits, "overallScore");

  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };
  addSlideHeader(
    slide,
    pres,
    g.hex,
    "✅  Overall Compliance",
    `Target: 95% · Current: ${cs.pct}% · Gap: ${gap >= 0 ? "+" : ""}${gap}%`,
  );

  // Big score (left)
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.3,
    y: 1.05,
    w: 3.0,
    h: 2.65,
    fill: { color: g.hex + "0D" },
    line: { color: g.hex + "40", width: 1 },
  });
  slide.addText(`${cs.pct}%`, {
    x: 0.3,
    y: 1.1,
    w: 3.0,
    h: 1.35,
    fontSize: 70,
    bold: true,
    color: g.hex,
    align: "center",
    fontFace: "Calibri",
    margin: 0,
  });
  slide.addText(g.label, {
    x: 0.3,
    y: 2.48,
    w: 3.0,
    h: 0.38,
    fontSize: 13,
    bold: true,
    color: g.hex,
    align: "center",
    fontFace: "Calibri",
    margin: 0,
  });
  const gc = gap >= 0 ? "16a34a" : "dc2626";
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.55,
    y: 2.95,
    w: 2.5,
    h: 0.38,
    fill: { color: gc + "18" },
    line: { color: gc + "50", width: 1 },
  });
  slide.addText(`${gap >= 0 ? "+" : ""}${gap}% vs target`, {
    x: 0.55,
    y: 2.97,
    w: 2.5,
    h: 0.34,
    fontSize: 11,
    bold: true,
    color: gc,
    align: "center",
    fontFace: "Calibri",
    margin: 0,
  });

  // Donut (center)
  slide.addChart(
    pres.charts.DOUGHNUT,
    [
      {
        name: "Compliance",
        labels: ["Pass", "Fail", "Partial", "N.A."],
        values: [
          cs.totalPass || 0.001,
          cs.totalFail || 0.001,
          cs.totalPartial || 0.001,
          cs.totalNA || 0.001,
        ],
      },
    ],
    {
      x: 3.6,
      y: 1.05,
      w: 3.5,
      h: 3.0,
      chartColors: ["22c55e", g.hex, "d97706", "94A3B8"],
      holeSize: 58,
      showPercent: true,
      dataLabelColor: "FFFFFF",
      showLegend: true,
      legendPos: "b",
      chartArea: { fill: { color: "FFFFFF" } },
    },
  );

  // Stats table (right)
  const header = [
    [
      {
        text: "Status",
        options: {
          bold: true,
          fill: { color: "0F172A" },
          color: "FFFFFF",
          fontSize: 9.5,
        },
      },
      {
        text: "Count",
        options: {
          bold: true,
          fill: { color: "0F172A" },
          color: "FFFFFF",
          fontSize: 9.5,
        },
      },
      {
        text: "% Total",
        options: {
          bold: true,
          fill: { color: "0F172A" },
          color: "FFFFFF",
          fontSize: 9.5,
        },
      },
    ],
  ];
  const statData = [
    ["✓ Pass", cs.totalPass, "22c55e"],
    ["✗ Fail", cs.totalFail, "ef4444"],
    ["~ Partial", cs.totalPartial, "d97706"],
    ["– N.A.", cs.totalNA, "94A3B8"],
  ];
  const rows = statData.map(([lbl, val, c]) => [
    { text: String(lbl), options: { fontSize: 10, color: c } },
    {
      text: val.toLocaleString(),
      options: { fontSize: 10, bold: true, color: c, align: "right" },
    },
    {
      text:
        cs.totalItems > 0
          ? `${Math.round((val / cs.totalItems) * 100)}%`
          : "0%",
      options: { fontSize: 10, color: c, align: "right" },
    },
  ]);
  slide.addTable([...header, ...rows], {
    x: 7.35,
    y: 1.12,
    w: 2.35,
    fontSize: 10,
    fontFace: "Calibri",
    border: { pt: 0.5, color: "E2E8F0" },
    rowH: 0.42,
  });
  slide.addText(
    `Total: ${cs.totalItems.toLocaleString()} items\n${audits.length} audits`,
    {
      x: 7.35,
      y: 2.92,
      w: 2.35,
      h: 0.5,
      fontSize: 8.5,
      color: "94A3B8",
      fontFace: "Calibri",
      align: "center",
      margin: 0,
    },
  );

  addFooter(slide, pres, 6);
}

// ─── SLIDE 7 — Score by Section ───────────────────────────────────────────────
function makeSlide7(pres, audits) {
  const sections = sectionStats(audits).slice(0, 10);

  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };
  addSlideHeader(
    slide,
    pres,
    "3b82f6",
    "📋  Score by Section",
    "Per-section average scores · click Pass / Fail / Partial to see item breakdown",
  );

  if (!sections.length) {
    slide.addText("No section data available", {
      x: 2,
      y: 2.5,
      w: 6,
      h: 1,
      fontSize: 13,
      color: "94A3B8",
      align: "center",
    });
    addFooter(slide, pres, 7);
    return;
  }

  // Horizontal bar chart (left)
  slide.addChart(
    pres.charts.BAR,
    [
      {
        name: "Avg Score",
        labels: sections.map((s) =>
          s.title.length > 24 ? s.title.slice(0, 24) + "…" : s.title,
        ),
        values: sections.map((s) => s.avg),
      },
    ],
    {
      x: 0.3,
      y: 1.05,
      w: 5.5,
      h: 4.08,
      barDir: "bar",
      chartColors: sections.map((s) => gradeColor(s.avg).hex),
      chartArea: { fill: { color: "FFFFFF" }, roundedCorners: true },
      catAxisLabelColor: "64748B",
      valAxisLabelColor: "64748B",
      valGridLine: { color: "E2E8F0", size: 0.5 },
      catGridLine: { style: "none" },
      valAxisMinVal: 0,
      valAxisMaxVal: 100,
      showValue: true,
      dataLabelColor: "1E293B",
      showLegend: false,
    },
  );

  // Table right
  const header = [
    [
      {
        text: "Section",
        options: {
          bold: true,
          fill: { color: "0F172A" },
          color: "FFFFFF",
          fontSize: 8,
        },
      },
      {
        text: "Avg",
        options: {
          bold: true,
          fill: { color: "0F172A" },
          color: "FFFFFF",
          fontSize: 8,
        },
      },
      {
        text: "✓",
        options: {
          bold: true,
          fill: { color: "0F172A" },
          color: "22c55e",
          fontSize: 8,
        },
      },
      {
        text: "✗",
        options: {
          bold: true,
          fill: { color: "0F172A" },
          color: "ef4444",
          fontSize: 8,
        },
      },
      {
        text: "~",
        options: {
          bold: true,
          fill: { color: "0F172A" },
          color: "f59e0b",
          fontSize: 8,
        },
      },
    ],
  ];
  const rows = sections.map((s) => [
    {
      text: s.title.length > 18 ? s.title.slice(0, 18) + "…" : s.title,
      options: { fontSize: 7.5, color: "334155" },
    },
    {
      text: `${s.avg}%`,
      options: { fontSize: 8, bold: true, color: gradeColor(s.avg).hex },
    },
    { text: String(s.pass), options: { fontSize: 8, color: "16a34a" } },
    { text: String(s.fail), options: { fontSize: 8, color: "dc2626" } },
    { text: String(s.partial), options: { fontSize: 8, color: "d97706" } },
  ]);
  slide.addTable([...header, ...rows], {
    x: 6.05,
    y: 1.05,
    w: 3.65,
    colW: [1.55, 0.65, 0.45, 0.45, 0.45],
    fontSize: 8,
    fontFace: "Calibri",
    border: { pt: 0.5, color: "E2E8F0" },
    rowH: 0.36,
  });

  addFooter(slide, pres, 7);
}

// ─── SLIDE 8 — Rating Distribution ───────────────────────────────────────────
function makeSlide8(pres, audits) {
  const dist = ratingDist(audits);

  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };
  addSlideHeader(
    slide,
    pres,
    "8b5cf6",
    "📊  Rating Distribution",
    "Audit count across performance rating categories",
  );

  // Horizontal bar (left)
  slide.addChart(
    pres.charts.BAR,
    [
      {
        name: "Audits",
        labels: dist.map((d) => d.label),
        values: dist.map((d) => d.count),
      },
    ],
    {
      x: 0.3,
      y: 1.05,
      w: 5.8,
      h: 3.65,
      barDir: "bar",
      chartColors: dist.map((d) => d.color),
      chartArea: { fill: { color: "FFFFFF" }, roundedCorners: true },
      catAxisLabelColor: "64748B",
      valAxisLabelColor: "64748B",
      valGridLine: { color: "E2E8F0", size: 0.5 },
      catGridLine: { style: "none" },
      showValue: true,
      dataLabelColor: "1E293B",
      showLegend: false,
    },
  );

  // Donut (right)
  slide.addChart(
    pres.charts.DOUGHNUT,
    [
      {
        name: "Distribution",
        labels: dist.map((d) => d.label.split(" ")[0]),
        values: dist.map((d) => d.count || 0.001),
      },
    ],
    {
      x: 6.3,
      y: 1.05,
      w: 3.4,
      h: 2.95,
      chartColors: dist.map((d) => d.color),
      holeSize: 55,
      showPercent: true,
      dataLabelColor: "FFFFFF",
      showLegend: false,
      chartArea: { fill: { color: "FFFFFF" } },
    },
  );

  // Legend
  dist.forEach((d, i) => {
    const y = 4.1 + i * 0.22;
    slide.addShape(pres.shapes.OVAL, {
      x: 6.38,
      y: y + 0.04,
      w: 0.12,
      h: 0.12,
      fill: { color: d.color },
    });
    slide.addText(`${d.label}: ${d.count} audits`, {
      x: 6.58,
      y,
      w: 3.1,
      h: 0.22,
      fontSize: 8.5,
      color: "475569",
      fontFace: "Calibri",
      margin: 0,
    });
  });

  addFooter(slide, pres, 8);
}

// ─── SLIDE 9 — SLA Summary Report ────────────────────────────────────────────
function makeSlide9(pres, audits) {
  const cs = complianceStat(audits);
  const avg = audits.length
    ? Math.round(
        audits.reduce((s, a) => s + (a.overallScore || 0), 0) / audits.length,
      )
    : 0;
  const g = gradeColor(avg);
  const dist = ratingDist(audits);
  const excellent = dist.find((d) => d.label.includes("Excellent"))?.count || 0;
  const critical = dist.find((d) => d.label.includes("Critical"))?.count || 0;

  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };
  addSlideHeader(
    slide,
    pres,
    "0F172A",
    "📑  SLA Scoring Report — Summary",
    `Based on ${audits.length} audit submissions · complete performance overview`,
  );

  // 6 KPI tiles (2 rows × 3)
  const kpis = [
    { v: `${avg}%`, l: "Overall Avg Score", c: g.hex },
    { v: String(cs.totalItems), l: "Total Items Audited", c: "3b82f6" },
    { v: String(audits.length), l: "Total Audits", c: "8b5cf6" },
    { v: `${cs.pct}%`, l: "Overall Compliance", c: gradeColor(cs.pct).hex },
    { v: String(cs.totalFail), l: "Critical Fail Items", c: "dc2626" },
    { v: `${excellent}/${critical}`, l: "Excellent / Critical", c: "22c55e" },
  ];
  kpis.forEach((k, i) => {
    const col = i % 3,
      row = Math.floor(i / 3);
    kpiTile(
      slide,
      pres,
      0.3 + col * 3.15,
      1.08 + row * 1.32,
      2.92,
      1.1,
      k.v,
      k.l,
      k.c,
    );
  });

  // Score Range reference table
  slide.addText("Score Range Reference", {
    x: 0.3,
    y: 3.65,
    w: 9.4,
    h: 0.3,
    fontSize: 12,
    bold: true,
    color: "1E293B",
    fontFace: "Calibri",
    margin: 0,
  });
  const RANGES = [
    {
      range: "95–100%",
      label: "Excellent",
      action: "Maintain",
      color: "22c55e",
    },
    { range: "85–94%", label: "Good", action: "Monitor", color: "3b82f6" },
    {
      range: "75–84%",
      label: "Needs Improvement",
      action: "Root cause analysis",
      color: "f59e0b",
    },
    {
      range: "60–74%",
      label: "At Risk",
      action: "Immediate attention",
      color: "ea580c",
    },
    {
      range: "<60%",
      label: "Critical",
      action: "Immediate remediation",
      color: "dc2626",
    },
  ];
  const rangeHeader = [
    [
      {
        text: "Range",
        options: {
          bold: true,
          fill: { color: "0F172A" },
          color: "FFFFFF",
          fontSize: 9.5,
        },
      },
      {
        text: "Label",
        options: {
          bold: true,
          fill: { color: "0F172A" },
          color: "FFFFFF",
          fontSize: 9.5,
        },
      },
      {
        text: "Action",
        options: {
          bold: true,
          fill: { color: "0F172A" },
          color: "FFFFFF",
          fontSize: 9.5,
        },
      },
      {
        text: "Status",
        options: {
          bold: true,
          fill: { color: "0F172A" },
          color: "FFFFFF",
          fontSize: 9.5,
        },
      },
    ],
  ];
  const rangeRows = RANGES.map((r) => [
    { text: r.range, options: { fontSize: 9, color: "334155" } },
    { text: r.label, options: { fontSize: 9, bold: true, color: r.color } },
    { text: r.action, options: { fontSize: 9, color: "64748B" } },
    {
      text: g.label === r.label ? "◀ Current" : "",
      options: { fontSize: 9, bold: true, color: r.color },
    },
  ]);
  slide.addTable([...rangeHeader, ...rangeRows], {
    x: 0.3,
    y: 4.0,
    w: 9.4,
    colW: [1.4, 2.2, 4.3, 1.4],
    fontSize: 9,
    fontFace: "Calibri",
    border: { pt: 0.5, color: "E2E8F0" },
    rowH: 0.24,
  });

  addFooter(slide, pres, 9);
}

// ─── SLIDE 10 — Recent Audit Submissions Table ────────────────────────────────
function makeSlide10(pres, audits) {
  const recent = [...audits]
    .sort(
      (a, b) =>
        new Date(b.submittedAt).getTime() - new Date(a.submittedAt).getTime(),
    )
    .slice(0, 15);

  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };
  addSlideHeader(
    slide,
    pres,
    "0F172A",
    "📋  Recent Audit Submissions",
    "Latest 15 records with full score details",
  );

  const header = [
    [
      {
        text: "#",
        options: {
          bold: true,
          fill: { color: "0F172A" },
          color: "FFFFFF",
          fontSize: 9,
        },
      },
      {
        text: "Auditor",
        options: {
          bold: true,
          fill: { color: "0F172A" },
          color: "FFFFFF",
          fontSize: 9,
        },
      },
      {
        text: "Date",
        options: {
          bold: true,
          fill: { color: "0F172A" },
          color: "FFFFFF",
          fontSize: 9,
        },
      },
      {
        text: "Score",
        options: {
          bold: true,
          fill: { color: "0F172A" },
          color: "FFFFFF",
          fontSize: 9,
        },
      },
      {
        text: "✓ Pass",
        options: {
          bold: true,
          fill: { color: "0F172A" },
          color: "22c55e",
          fontSize: 9,
        },
      },
      {
        text: "✗ Fail",
        options: {
          bold: true,
          fill: { color: "0F172A" },
          color: "ef4444",
          fontSize: 9,
        },
      },
      {
        text: "~ Part",
        options: {
          bold: true,
          fill: { color: "0F172A" },
          color: "f59e0b",
          fontSize: 9,
        },
      },
      {
        text: "Rating",
        options: {
          bold: true,
          fill: { color: "0F172A" },
          color: "FFFFFF",
          fontSize: 9,
        },
      },
    ],
  ];
  const rows = recent.map((a, i) => {
    const g = gradeColor(a.overallScore || 0);
    const dt = new Date(a.submittedAt).toLocaleDateString("en-IN", {
      day: "2-digit",
      month: "short",
      year: "2-digit",
    });
    return [
      {
        text: String(i + 1),
        options: { fontSize: 8.5, color: "94A3B8", align: "center" },
      },
      {
        text: (a.auditorName || "").slice(0, 16),
        options: { fontSize: 8.5, color: "1E293B" },
      },
      { text: dt, options: { fontSize: 8.5, color: "64748B" } },
      {
        text: `${a.overallScore || 0}%`,
        options: { fontSize: 8.5, bold: true, color: g.hex },
      },
      {
        text: String(a.totalPass || 0),
        options: { fontSize: 8.5, color: "16a34a" },
      },
      {
        text: String(a.totalFail || 0),
        options: { fontSize: 8.5, color: "dc2626" },
      },
      {
        text: String(a.totalPartial || 0),
        options: { fontSize: 8.5, color: "d97706" },
      },
      {
        text: (a.ratingLabel || g.label).slice(0, 14),
        options: { fontSize: 8, color: g.hex },
      },
    ];
  });
  slide.addTable([...header, ...rows], {
    x: 0.2,
    y: 1.05,
    w: 9.6,
    colW: [0.35, 1.7, 0.9, 0.65, 0.65, 0.6, 0.6, 1.35],
    fontSize: 8.5,
    fontFace: "Calibri",
    border: { pt: 0.4, color: "E2E8F0" },
    rowH: 0.27,
  });

  addFooter(slide, pres, 10);
}

// ─── MAIN EXPORT ──────────────────────────────────────────────────────────────
export async function exportDashboardPptx(req, res) {
  try {
    const audits = await AuditSubmission.find({})
      .sort({ submittedAt: -1 })
      .limit(1000)
      .lean();

    if (!audits.length) {
      return res.status(404).json({ error: "No audit data found" });
    }

    const avgScore = Math.round(
      audits.reduce((s, a) => s + (a.overallScore || 0), 0) / audits.length,
    );

    const pres = new PptxGenJS();
    pres.layout = "LAYOUT_16x9";
    pres.author = "Yakka Pukka";
    pres.title = "Admin Dashboard Report";
    pres.subject = "SLA Audit Performance";

    makeSlide1(pres, audits, avgScore); // Cover
    makeSlide2(pres, audits); // Avg Score (line + bar)
    makeSlide3(pres, audits); // Critical Fails (trend + bar)
    makeSlide4(pres, audits); // Top Issues Table
    makeSlide5(pres, audits); // Monthly Target (gauge style)
    makeSlide6(pres, audits); // Compliance (donut + table)
    makeSlide7(pres, audits); // Section Scores (bar + table)
    makeSlide8(pres, audits); // Rating Distribution (bar + donut)
    makeSlide9(pres, audits); // SLA Summary KPIs + range table
    makeSlide10(pres, audits); // Recent Audits Table

    const buffer = await pres.write({ outputType: "nodebuffer" });
    const date = new Date().toISOString().slice(0, 10);

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="YakkaPukka_Dashboard_${date}.pptx"`,
    );
    res.setHeader("Content-Length", buffer.length);
    res.send(buffer);
  } catch (err) {
    console.error("[PptxExportController]", err);
    res
      .status(500)
      .json({ error: "Failed to generate PowerPoint", details: err.message });
  }
}
