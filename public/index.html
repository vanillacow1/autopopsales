import { useState, useCallback, useRef } from "react";

// ── Exact style indices from styles.xml ──────────────────────────────────────
// s=17 : week header (blue FF2E5B9A bg, white bold)
// s=14 : empty filler cells in header row
// s=2  : column header (light blue FFD6E4F7, dark bold, center)
// s=5  : day name white bg, left
// s=8  : day name grey FFF5F7FA bg, left
// s=10 : sales white bg, \$#,##0.00, right
// s=12 : sales grey bg, \$#,##0.00, right
// s=11 : daily % white bg, 0.0%, center
// s=30 : dashboard label
// s=31 : dashboard SUM value ($format)
// s=32 : dashboard WoW diff (green bold)

const DAY_SS = {
  Sunday: 6,
  Monday: 15,
  Tuesday: 17,
  Wednesday: 19,
  Thursday: 21,
  Friday: 5,
  Saturday: 24,
};
const COL_SS = { Day: 9, Sales: 34, Pct: 35 };
const DAYS = [
  "Sunday",
  "Monday",
  "Tuesday",
  "Wednesday",
  "Thursday",
  "Friday",
  "Saturday",
];
const DAY_S = [5, 8]; // alternating day name style
const SALES_S = [10, 12]; // alternating sales style

const COLORS = {
  bg: "#0F1923",
  surface: "#1A2535",
  border: "#2A3F58",
  accent: "#2E5B9A",
  accentLight: "#4A7CC4",
  accentGlow: "rgba(46,91,154,0.25)",
  highlight: "#16A085",
  highlightGlow: "rgba(22,160,133,0.2)",
  danger: "#E74C3C",
  text: "#E8EFF8",
  textMuted: "#7A93B0",
  textDim: "#4A6280",
  headerBg: "#1B2A4A",
  warning: "#F39C12",
};

const initSales = () => DAYS.reduce((a, d) => ({ ...a, [d]: "" }), {});

function esc(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// Build the 9 row XML block for one week in sheet2
// headerRow = the merged title row (e.g. 173)
// dataStart = first data row (e.g. 175 = headerRow+2)
function buildWeekXml(headerRow, weekLabel, sales) {
  const dataStart = headerRow + 2;
  const total = DAYS.reduce((s, d) => s + (parseFloat(sales[d]) || 0), 0);

  let x = "";

  // Title row (merged B:E)
  x += `<row r="${headerRow}" spans="2:5" ht="24" customHeight="1" x14ac:dyDescent="0.25">`;
  x += `<c r="B${headerRow}" s="17" t="str">`;
  x += `<f>"${esc(
    weekLabel
  )} | Weekly Total: $" &amp; TEXT(SUM(C${dataStart}:C${
    dataStart + 6
  }),"#,##0.00")</f>`;
  x += `<v>${esc(weekLabel)} | Weekly Total: $${total.toFixed(2)}</v>`;
  x += `</c><c r="C${headerRow}" s="14"/><c r="D${headerRow}" s="14"/><c r="E${headerRow}" s="14"/></row>`;

  // Column header row
  const ch = headerRow + 1;
  x += `<row r="${ch}" spans="2:5" ht="18" customHeight="1" x14ac:dyDescent="0.25">`;
  x += `<c r="B${ch}" s="2" t="s"><v>${COL_SS.Day}</v></c>`;
  x += `<c r="C${ch}" s="2" t="s"><v>${COL_SS.Sales}</v></c>`;
  x += `<c r="D${ch}" s="2" t="s"><v>${COL_SS.Pct}</v></c>`;
  x += `</row>`;

  // 7 data rows
  DAYS.forEach((day, i) => {
    const r = dataStart + i;
    const p = i % 2;
    const val = parseFloat(sales[day]) || 0;
    const pct = total > 0 ? val / total : 0;
    x += `<row r="${r}" spans="2:5" ht="18" customHeight="1" x14ac:dyDescent="0.25">`;
    x += `<c r="B${r}" s="${DAY_S[p]}" t="s"><v>${DAY_SS[day]}</v></c>`;
    if (val) {
      x += `<c r="C${r}" s="${SALES_S[p]}"><v>${val}</v></c>`;
    } else {
      x += `<c r="C${r}" s="${SALES_S[p]}"/>`;
    }
    // Full explicit formula on every row (no shared formula — avoids si conflicts)
    x += `<c r="D${r}" s="11"><f>C${r}/SUM($C$${dataStart}:$C$${
      dataStart + 6
    })</f><v>${pct}</v></c>`;
    x += `</row>`;
  });

  return { xml: x, dataStart, total };
}

// Build one dashboard row for sheet1
function buildDashXml(
  dashRow,
  weekIdx,
  weekLabel,
  dataStart,
  total,
  labelSsIdx,
  isFirst
) {
  // F-label style alternates s=3 (grey bg) / s=7 (white bg) by weekIdx parity
  const fStyle = weekIdx % 2 === 1 ? 7 : 3;
  // s=31 = G total cell: numFmt 8 ("$"#,##0.00;[Red]-"$"#,##0.00), grey fill, border
  // s=32 = H WoW cell:   numFmt 8, white fill, border
  let x = `<row r="${dashRow}" spans="6:8" ht="18" customHeight="1" x14ac:dyDescent="0.25">`;
  x += `<c r="F${dashRow}" s="${fStyle}" t="s"><v>${labelSsIdx}</v></c>`;
  x += `<c r="G${dashRow}" s="10"><f>SUM('📅 Weekly Data'!C${dataStart}:C${
    dataStart + 6
  })</f><v>${total}</v></c>`;
  if (!isFirst) {
    x += `<c r="H${dashRow}" s="5"><f>G${dashRow}-G${
      dashRow - 1
    }</f><v>0</v></c>`;
  }
  x += `</row>`;
  return x;
}

async function loadJSZip() {
  if (window.JSZip) return window.JSZip;
  await new Promise((res, rej) => {
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";
    s.onload = res;
    s.onerror = rej;
    document.head.appendChild(s);
  });
  return window.JSZip;
}

async function processXlsx(file, weekLabel, sales) {
  const JSZip = await loadJSZip();
  const zip = await JSZip.loadAsync(await file.arrayBuffer());

  let s1 = await zip.file("xl/worksheets/sheet1.xml").async("string"); // dashboard
  let s2 = await zip.file("xl/worksheets/sheet2.xml").async("string"); // weekly data
  let ss = await zip.file("xl/sharedStrings.xml").async("string");

  // ── 1. Detect where to write ────────────────────────────────────────────
  // Each week block: header at row 3+n*10, col-header at 4+n*10, data at 5+n*10 to 11+n*10
  // A week slot "exists" if the header row is present in XML.
  // A week slot is "empty" (needs filling) if Sunday's C cell has no value.
  // We find the LAST header row present, then check if it's empty.

  let weekIdx = 0;
  while (s2.includes(`r="B${3 + weekIdx * 10}"`)) weekIdx++;
  // weekIdx is now the count of existing header rows

  // Check if the last existing week's Sunday data row (dataStart) has a value
  const lastWeekIdx = weekIdx - 1;
  const lastDataStart = 5 + lastWeekIdx * 10;
  const lastSundayHasValue = new RegExp(`r="C${lastDataStart}"[^>]*><v>`).test(
    s2
  );

  let targetHeaderRow, targetDataStart, isNewBlock;

  if (!lastSundayHasValue) {
    // Fill in the existing empty block
    targetHeaderRow = 3 + lastWeekIdx * 10;
    targetDataStart = lastDataStart;
    isNewBlock = false;
  } else {
    // Append a completely new block after the last one
    targetHeaderRow = 3 + weekIdx * 10;
    targetDataStart = 5 + weekIdx * 10;
    isNewBlock = true;
    weekIdx = weekIdx; // this is the new week's index
  }

  // ── 2. Add week label to sharedStrings ─────────────────────────────────
  const siCount = (ss.match(/<si>/g) || []).length;
  const newSsIdx = siCount;
  ss = ss.replace("</sst>", `<si><t>${esc(weekLabel)}</t></si></sst>`);
  ss = ss.replace(
    /(<sst[^>]* count=")(\d+)(")/,
    (_, a, n, b) => `${a}${+n + 1}${b}`
  );
  ss = ss.replace(
    /(<sst[^>]* uniqueCount=")(\d+)(")/,
    (_, a, n, b) => `${a}${+n + 1}${b}`
  );

  // ── 3. Build week XML ───────────────────────────────────────────────────
  const {
    xml: weekXml,
    dataStart,
    total,
  } = buildWeekXml(targetHeaderRow, weekLabel, sales);

  if (isNewBlock) {
    // Append rows before </sheetData>
    s2 = s2.replace("</sheetData>", weekXml + "</sheetData>");
    // Add merge cell
    s2 = s2.replace(
      /(<mergeCells count=")(\d+)(")/,
      (_, a, n, b) => `${a}${+n + 1}${b}`
    );
    s2 = s2.replace(
      "</mergeCells>",
      `<mergeCell ref="B${targetHeaderRow}:E${targetHeaderRow}"/></mergeCells>`
    );
  } else {
    // Replace the existing rows for this week block (title + col header + 7 data rows = 9 rows)
    // We replace the existing header row through the last data row
    const endRow = targetDataStart + 6;
    // Build a regex that matches all rows from headerRow to endRow
    // Strategy: find the block by locating header row and replacing up to and including row endRow
    const headerPattern = new RegExp(
      `<row r="${targetHeaderRow}"[\\s\\S]*?</row>` +
        `[\\s\\S]*?` +
        `<row r="${endRow}"[^>]*>[\\s\\S]*?</row>`
    );
    if (headerPattern.test(s2)) {
      s2 = s2.replace(headerPattern, weekXml.trimEnd());
    } else {
      // Fallback: just insert before </sheetData>
      s2 = s2.replace("</sheetData>", weekXml + "</sheetData>");
    }
    // Update the header row's formula value (it exists but had wrong total)
    // Already handled by full replacement above
  }

  // ── 4. Dashboard row ────────────────────────────────────────────────────
  // Determine which dashboard row this corresponds to
  // Dashboard row 8 = week index 0, row 9 = week index 1, etc.
  const targetWeekIdx = isNewBlock ? weekIdx : lastWeekIdx;
  const dashRow = 8 + targetWeekIdx;
  const isFirstWeek = targetWeekIdx === 0;

  const dashXml = buildDashXml(
    dashRow,
    targetWeekIdx,
    weekLabel,
    dataStart,
    total,
    newSsIdx,
    isFirstWeek
  );

  if (!isNewBlock) {
    // Replace existing dashboard row
    const existingDashRow = new RegExp(
      `<row r="${dashRow}"[^>]*>[\\s\\S]*?</row>`
    );
    if (existingDashRow.test(s1)) {
      s1 = s1.replace(existingDashRow, dashXml);
    } else {
      s1 = s1.replace("</sheetData>", dashXml + "</sheetData>");
    }
  } else {
    // Append new dashboard row
    s1 = s1.replace("</sheetData>", dashXml + "</sheetData>");
  }

  // ── 5. Remove calcChain so Excel recalculates ───────────────────────────
  zip.remove("xl/calcChain.xml");
  let ct = await zip.file("[Content_Types].xml").async("string");
  ct = ct.replace(/<Override PartName="\/xl\/calcChain\.xml"[^/]*\/>/g, "");
  zip.file("[Content_Types].xml", ct);

  // ── 6. Write back ───────────────────────────────────────────────────────
  zip.file("xl/worksheets/sheet1.xml", s1);
  zip.file("xl/worksheets/sheet2.xml", s2);
  zip.file("xl/sharedStrings.xml", ss);

  return zip.generateAsync({
    type: "blob",
    mimeType:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    compression: "DEFLATE",
  });
}

// ── UI ───────────────────────────────────────────────────────────────────────
export default function App() {
  const [weekLabel, setWeekLabel] = useState("");
  const [sales, setSales] = useState(initSales());
  const [file, setFile] = useState(null);
  const [fileName, setFileName] = useState("");
  const [status, setStatus] = useState(null);
  const [busy, setBusy] = useState(false);
  const [preview, setPreview] = useState(null);
  const fileRef = useRef();

  const total = DAYS.reduce((s, d) => s + (parseFloat(sales[d]) || 0), 0);

  const pickFile = (e) => {
    const f = e.target.files[0];
    if (!f) return;
    setFile(f);
    setFileName(f.name);
    setStatus({ type: "info", msg: `Loaded: ${f.name}` });
  };
  const onDrop = useCallback((e) => {
    e.preventDefault();
    const f = e.dataTransfer.files[0];
    if (f?.name.endsWith(".xlsx")) {
      setFile(f);
      setFileName(f.name);
      setStatus({ type: "info", msg: `Loaded: ${f.name}` });
    }
  }, []);

  const doPreview = () => {
    if (!file)
      return setStatus({
        type: "error",
        msg: "Upload your spreadsheet first.",
      });
    if (!weekLabel.trim())
      return setStatus({ type: "error", msg: "Enter a week label." });
    if (!DAYS.some((d) => sales[d] !== ""))
      return setStatus({
        type: "error",
        msg: "Enter at least one day's sales.",
      });
    setPreview({
      label: weekLabel.trim(),
      rows: DAYS.map((d) => ({ day: d, val: parseFloat(sales[d]) || 0 })),
      total,
    });
    setStatus({ type: "info", msg: "Confirm the preview below, then submit." });
  };

  const doSubmit = async () => {
    if (!preview) return;
    setBusy(true);
    setStatus({ type: "info", msg: "Processing — don't close this tab…" });
    try {
      const blob = await processXlsx(file, preview.label, sales);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = fileName.replace(".xlsx", "_updated.xlsx");
      a.click();
      URL.revokeObjectURL(url);
      setStatus({
        type: "success",
        msg: `✓ "${preview.label}" saved. File downloaded!`,
      });
      setPreview(null);
      setWeekLabel("");
      setSales(initSales());
    } catch (err) {
      console.error(err);
      setStatus({ type: "error", msg: `Error: ${err.message}` });
    }
    setBusy(false);
  };

  const doClear = () => {
    setSales(initSales());
    setWeekLabel("");
    setPreview(null);
    setStatus(null);
  };

  const barMax = Math.max(...DAYS.map((d) => parseFloat(sales[d]) || 0), 1);

  return (
    <div
      style={{
        minHeight: "100vh",
        background: COLORS.bg,
        fontFamily: "'DM Sans','Segoe UI',sans-serif",
        color: COLORS.text,
      }}
    >
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        input[type=number]::-webkit-inner-spin-button{-webkit-appearance:none}
        input[type=number]{-moz-appearance:textfield}
        .drop{border:2px dashed ${COLORS.border};border-radius:12px;padding:32px;text-align:center;cursor:pointer;transition:all .2s;background:${COLORS.surface}}
        .drop:hover{border-color:${COLORS.accentLight};background:#1e2d40}
        .di{width:100%;background:${COLORS.bg};border:1px solid ${COLORS.border};border-radius:8px;color:${COLORS.text};font-family:'DM Mono',monospace;font-size:15px;padding:10px 14px 10px 26px;outline:none;transition:all .15s}
        .di:focus{border-color:${COLORS.accentLight};box-shadow:0 0 0 3px ${COLORS.accentGlow}}
        .wi{width:100%;background:${COLORS.bg};border:1px solid ${COLORS.border};border-radius:8px;color:${COLORS.text};font-size:15px;padding:11px 14px;outline:none;transition:all .15s;font-family:'DM Sans',sans-serif}
        .wi:focus{border-color:${COLORS.accentLight};box-shadow:0 0 0 3px ${COLORS.accentGlow}}
        .btn{border:none;border-radius:10px;cursor:pointer;font-family:'DM Sans',sans-serif;font-weight:600;font-size:14px;padding:12px 28px;transition:all .18s}
        .pa{background:${COLORS.accent};color:#fff}.pa:hover{background:${COLORS.accentLight};transform:translateY(-1px);box-shadow:0 4px 16px ${COLORS.accentGlow}}
        .su{background:${COLORS.highlight};color:#fff}.su:hover{background:#1abc9c;transform:translateY(-1px)}
        .gh{background:transparent;color:${COLORS.textMuted};border:1px solid ${COLORS.border}}.gh:hover{background:${COLORS.surface};color:${COLORS.text}}
        .btn:disabled{opacity:.4;cursor:not-allowed;transform:none!important}
        .card{background:${COLORS.surface};border:1px solid ${COLORS.border};border-radius:14px;padding:24px}
        .bar{height:3px;border-radius:2px;background:${COLORS.accent};transition:width .4s}
        .lbl{font-size:11px;font-weight:600;letter-spacing:1px;text-transform:uppercase;color:${COLORS.textMuted};margin-bottom:6px}
      `}</style>

      <div
        style={{
          borderBottom: `1px solid ${COLORS.border}`,
          background: COLORS.headerBg,
          padding: "0 32px",
        }}
      >
        <div
          style={{
            maxWidth: 860,
            margin: "0 auto",
            height: 64,
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
          }}
        >
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div
              style={{
                width: 32,
                height: 32,
                borderRadius: 8,
                background: COLORS.accent,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontSize: 16,
              }}
            >
              📊
            </div>
            <div>
              <div style={{ fontWeight: 700, fontSize: 16 }}>
                Traffic Tracker
              </div>
              <div
                style={{
                  fontSize: 11,
                  color: COLORS.textMuted,
                  letterSpacing: "0.5px",
                }}
              >
                WEEKLY DATA ENTRY
              </div>
            </div>
          </div>
          {total > 0 && (
            <div
              style={{
                background: COLORS.accentGlow,
                border: `1px solid ${COLORS.accent}`,
                borderRadius: 8,
                padding: "6px 14px",
                fontFamily: "'DM Mono',monospace",
                fontSize: 13,
              }}
            >
              Total:{" "}
              <span style={{ color: COLORS.accentLight, fontWeight: 600 }}>
                $
                {total.toLocaleString("en-US", {
                  minimumFractionDigits: 2,
                  maximumFractionDigits: 2,
                })}
              </span>
            </div>
          )}
        </div>
      </div>

      <div
        style={{
          maxWidth: 860,
          margin: "0 auto",
          padding: "32px 24px",
          display: "flex",
          flexDirection: "column",
          gap: 20,
        }}
      >
        {/* Step 1 */}
        <div className="card">
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 10,
              marginBottom: 16,
            }}
          >
            <div
              style={{
                width: 24,
                height: 24,
                borderRadius: "50%",
                background: file ? COLORS.highlight : COLORS.accent,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontSize: 12,
                fontWeight: 700,
                color: "#fff",
              }}
            >
              {file ? "✓" : "1"}
            </div>
            <div style={{ fontWeight: 600, fontSize: 15 }}>
              Upload Spreadsheet
            </div>
          </div>
          <div
            className="drop"
            onDragOver={(e) => e.preventDefault()}
            onDrop={onDrop}
            onClick={() => fileRef.current.click()}
          >
            <input
              ref={fileRef}
              type="file"
              accept=".xlsx"
              style={{ display: "none" }}
              onChange={pickFile}
            />
            {file ? (
              <div>
                <div style={{ fontSize: 28, marginBottom: 8 }}>📁</div>
                <div
                  style={{
                    color: COLORS.highlight,
                    fontWeight: 600,
                    fontSize: 14,
                  }}
                >
                  {fileName}
                </div>
                <div
                  style={{
                    color: COLORS.textMuted,
                    fontSize: 12,
                    marginTop: 4,
                  }}
                >
                  Click to replace
                </div>
              </div>
            ) : (
              <div>
                <div style={{ fontSize: 32, marginBottom: 8 }}>⬆️</div>
                <div style={{ fontWeight: 600, marginBottom: 4 }}>
                  Drop your Traffic_Tracker.xlsx here
                </div>
                <div style={{ color: COLORS.textMuted, fontSize: 13 }}>
                  or click to browse
                </div>
              </div>
            )}
          </div>
        </div>

        {/* Step 2 */}
        <div className="card">
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 10,
              marginBottom: 16,
            }}
          >
            <div
              style={{
                width: 24,
                height: 24,
                borderRadius: "50%",
                background: weekLabel ? COLORS.highlight : COLORS.accent,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontSize: 12,
                fontWeight: 700,
                color: "#fff",
              }}
            >
              {weekLabel ? "✓" : "2"}
            </div>
            <div style={{ fontWeight: 600, fontSize: 15 }}>Week Label</div>
          </div>
          <div className="lbl">Date range</div>
          <input
            className="wi"
            type="text"
            placeholder="e.g. April 20-26"
            value={weekLabel}
            onChange={(e) => setWeekLabel(e.target.value)}
          />
          <div style={{ color: COLORS.textMuted, fontSize: 12, marginTop: 8 }}>
            Appears in the Weekly Sales Trend on the dashboard.
          </div>
        </div>

        {/* Step 3 */}
        <div className="card">
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 10,
              marginBottom: 20,
            }}
          >
            <div
              style={{
                width: 24,
                height: 24,
                borderRadius: "50%",
                background: COLORS.accent,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontSize: 12,
                fontWeight: 700,
                color: "#fff",
              }}
            >
              3
            </div>
            <div style={{ fontWeight: 600, fontSize: 15 }}>Daily Sales</div>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {DAYS.map((day, i) => {
              const val = parseFloat(sales[day]) || 0;
              const bw = barMax > 0 ? (val / barMax) * 100 : 0;
              const wk = i === 0 || i === 6;
              return (
                <div
                  key={day}
                  style={{
                    display: "grid",
                    gridTemplateColumns: "110px 1fr 80px",
                    gap: 12,
                    alignItems: "center",
                  }}
                >
                  <div
                    style={{ display: "flex", alignItems: "center", gap: 8 }}
                  >
                    <div
                      style={{
                        width: 6,
                        height: 6,
                        borderRadius: "50%",
                        background: wk ? COLORS.warning : COLORS.accentLight,
                        flexShrink: 0,
                      }}
                    />
                    <span
                      style={{
                        fontSize: 13,
                        fontWeight: 500,
                        color: wk ? COLORS.text : COLORS.textMuted,
                      }}
                    >
                      {day}
                    </span>
                  </div>
                  <div>
                    <div style={{ position: "relative" }}>
                      <span
                        style={{
                          position: "absolute",
                          left: 10,
                          top: "50%",
                          transform: "translateY(-50%)",
                          color: COLORS.textDim,
                          fontFamily: "'DM Mono',monospace",
                          fontSize: 14,
                          pointerEvents: "none",
                        }}
                      >
                        $
                      </span>
                      <input
                        className="di"
                        type="number"
                        min="0"
                        step="0.01"
                        placeholder="0.00"
                        value={sales[day]}
                        onChange={(e) =>
                          setSales((p) => ({ ...p, [day]: e.target.value }))
                        }
                      />
                    </div>
                    {val > 0 && (
                      <div style={{ marginTop: 4 }}>
                        <div className="bar" style={{ width: `${bw}%` }} />
                      </div>
                    )}
                  </div>
                  <div
                    style={{
                      textAlign: "right",
                      fontFamily: "'DM Mono',monospace",
                      fontSize: 13,
                      color: val > 0 ? COLORS.accentLight : COLORS.textDim,
                    }}
                  >
                    {total > 0 && val > 0
                      ? ((val / total) * 100).toFixed(1) + "%"
                      : "—"}
                  </div>
                </div>
              );
            })}
          </div>
          {total > 0 && (
            <div
              style={{
                marginTop: 16,
                paddingTop: 16,
                borderTop: `1px solid ${COLORS.border}`,
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
              }}
            >
              <span
                style={{
                  fontSize: 11,
                  fontWeight: 600,
                  textTransform: "uppercase",
                  letterSpacing: "0.8px",
                  color: COLORS.textMuted,
                }}
              >
                Weekly Total
              </span>
              <span
                style={{
                  fontFamily: "'DM Mono',monospace",
                  fontWeight: 700,
                  fontSize: 18,
                }}
              >
                $
                {total.toLocaleString("en-US", {
                  minimumFractionDigits: 2,
                  maximumFractionDigits: 2,
                })}
              </span>
            </div>
          )}
        </div>

        {status && (
          <div
            style={{
              padding: "12px 16px",
              borderRadius: 10,
              fontSize: 13,
              fontWeight: 500,
              background:
                status.type === "success"
                  ? COLORS.highlightGlow
                  : status.type === "error"
                  ? "rgba(231,76,60,0.15)"
                  : COLORS.accentGlow,
              border: `1px solid ${
                status.type === "success"
                  ? COLORS.highlight
                  : status.type === "error"
                  ? COLORS.danger
                  : COLORS.accent
              }`,
              color:
                status.type === "success"
                  ? "#1abc9c"
                  : status.type === "error"
                  ? COLORS.danger
                  : COLORS.accentLight,
            }}
          >
            {status.msg}
          </div>
        )}

        {preview && (
          <div className="card" style={{ borderColor: COLORS.accent }}>
            <div
              style={{
                fontWeight: 600,
                fontSize: 14,
                marginBottom: 14,
                color: COLORS.accentLight,
              }}
            >
              Preview — {preview.label}
            </div>
            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(7,1fr)",
                gap: 8,
              }}
            >
              {preview.rows.map(({ day, val }) => (
                <div
                  key={day}
                  style={{
                    textAlign: "center",
                    background: COLORS.bg,
                    borderRadius: 8,
                    padding: "10px 6px",
                  }}
                >
                  <div
                    style={{
                      fontSize: 11,
                      color: COLORS.textMuted,
                      marginBottom: 4,
                      fontWeight: 600,
                    }}
                  >
                    {day.slice(0, 3).toUpperCase()}
                  </div>
                  <div
                    style={{
                      fontFamily: "'DM Mono',monospace",
                      fontSize: 13,
                      fontWeight: 600,
                    }}
                  >
                    ${val.toFixed(0)}
                  </div>
                  <div
                    style={{
                      fontSize: 11,
                      color: COLORS.textDim,
                      marginTop: 2,
                    }}
                  >
                    {preview.total > 0
                      ? ((val / preview.total) * 100).toFixed(1) + "%"
                      : "—"}
                  </div>
                </div>
              ))}
            </div>
            <div
              style={{
                marginTop: 12,
                paddingTop: 12,
                borderTop: `1px solid ${COLORS.border}`,
                display: "flex",
                justifyContent: "space-between",
                fontSize: 13,
                color: COLORS.textMuted,
              }}
            >
              <span>Total</span>
              <span
                style={{
                  fontFamily: "'DM Mono',monospace",
                  color: COLORS.text,
                  fontWeight: 700,
                }}
              >
                $
                {preview.total.toLocaleString("en-US", {
                  minimumFractionDigits: 2,
                  maximumFractionDigits: 2,
                })}
              </span>
            </div>
          </div>
        )}

        <div style={{ display: "flex", gap: 12, justifyContent: "flex-end" }}>
          <button className="btn gh" onClick={doClear}>
            Clear
          </button>
          {!preview ? (
            <button className="btn pa" onClick={doPreview}>
              Preview Week →
            </button>
          ) : (
            <>
              <button className="btn gh" onClick={() => setPreview(null)}>
                Edit
              </button>
              <button className="btn su" onClick={doSubmit} disabled={busy}>
                {busy ? "Processing…" : "✓ Submit to Excel"}
              </button>
            </>
          )}
        </div>

        <div
          style={{
            textAlign: "center",
            color: COLORS.textDim,
            fontSize: 12,
            paddingBottom: 8,
          }}
        >
          Writes directly to xlsx XML · All original formatting preserved
        </div>
      </div>
    </div>
  );
}
