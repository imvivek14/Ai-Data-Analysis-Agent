import { useState, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import {
  LineChart, Line, BarChart, Bar, ScatterChart, Scatter,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  Cell, AreaChart, Area
} from "recharts";

const COLORS = ["#378ADD", "#1D9E75", "#D85A30", "#D4537E", "#639922", "#BA7517", "#7F77DD", "#888780"];

function StatCard({ label, value, sub }) {
  return (
    <div style={{
      background: "#111827",
      border: "1px solid #1f2937",
      borderRadius: "20px",
      padding: "1.25rem",
      minWidth: 0,
      boxShadow: "0 10px 25px rgba(0,0,0,0.25)",
      transition: "all 0.3s ease"
    }}>
      <p style={{ fontSize: 12, color: "var(--color-text-secondary)", margin: "0 0 4px", textTransform: "uppercase", letterSpacing: "0.06em" }}>{label}</p>
      <p style={{ fontSize: 22, fontWeight: 500, margin: 0, color: "#f8fafc" }}>{value}</p>
      {sub && <p style={{ fontSize: 12, color: "#94a3b8", margin: "4px 0 0" }}>{sub}</p>}
    </div>
  );
}

function ChartBlock({ title, children }) {
  return (
    <div
      style={{
        background: "#111827",
        border: "1px solid #1f2937",
        borderRadius: "24px",
        padding: "1.5rem",
        marginBottom: "1.5rem",
        boxShadow: "0 10px 25px rgba(0,0,0,0.25)",
        overflow: "hidden"
      }}
    >
      <p
        style={{
          fontWeight: "600",
          margin: "0 0 1rem",
          fontSize: "0.9rem",
          color: "#38bdf8",
          textTransform: "uppercase",
          letterSpacing: "0.08em"
        }}
      >
        {title}
      </p>

      {children}
    </div>
  );
}

function buildHistogram(values, bins = 12) {
  const clean = values.filter(v => v != null && !isNaN(v));
  if (!clean.length) return [];
  const min = Math.min(...clean), max = Math.max(...clean);
  const step = (max - min) / bins || 1;
  const buckets = Array.from({ length: bins }, (_, i) => ({
    range: `${(min + i * step).toFixed(1)}–${(min + (i + 1) * step).toFixed(1)}`,
    count: 0
  }));
  clean.forEach(v => {
    const idx = Math.min(Math.floor((v - min) / step), bins - 1);
    buckets[idx].count++;
  });
  return buckets;
}

function computeStats(values) {
  const clean = values.filter(v => v != null && !isNaN(Number(v))).map(Number);
  if (!clean.length) return null;
  clean.sort((a, b) => a - b);
  const n = clean.length;
  const mean = clean.reduce((s, x) => s + x, 0) / n;
  const median = n % 2 === 0 ? (clean[n / 2 - 1] + clean[n / 2]) / 2 : clean[Math.floor(n / 2)];
  const variance = clean.reduce((s, x) => s + (x - mean) ** 2, 0) / n;
  const std = Math.sqrt(variance);
  const q1 = clean[Math.floor(n * 0.25)];
  const q3 = clean[Math.floor(n * 0.75)];
  const iqr = q3 - q1;
  const outliers = clean.filter(x => x < q1 - 1.5 * iqr || x > q3 + 1.5 * iqr);
  const freq = {};
  clean.forEach(x => { freq[x] = (freq[x] || 0) + 1; });
  const mode = Object.entries(freq).sort((a, b) => b[1] - a[1])[0]?.[0];
  return { mean, median, std, variance, q1, q3, iqr, min: clean[0], max: clean[n - 1], n, outliers, mode: Number(mode) };
}

function pearsonCorr(a, b) {
  const n = Math.min(a.length, b.length);
  const pairs = Array.from({ length: n }, (_, i) => [a[i], b[i]]).filter(([x, y]) => x != null && y != null && !isNaN(x) && !isNaN(y));
  if (pairs.length < 2) return null;
  const meanA = pairs.reduce((s, [x]) => s + x, 0) / pairs.length;
  const meanB = pairs.reduce((s, [, y]) => s + y, 0) / pairs.length;
  const num = pairs.reduce((s, [x, y]) => s + (x - meanA) * (y - meanB), 0);
  const den = Math.sqrt(pairs.reduce((s, [x]) => s + (x - meanA) ** 2, 0) * pairs.reduce((s, [, y]) => s + (y - meanB) ** 2, 0));
  return den === 0 ? 0 : num / den;
}

function isNumericCol(values) {
  const clean = values.filter(v => v != null && v !== "");
  return clean.length > 0 && clean.filter(v => !isNaN(Number(v))).length / clean.length > 0.8;
}

export default function DataAnalysisAgent() {
  const [phase, setPhase] = useState("idle"); // idle | loading | analyzing | done | error
  const [rawData, setRawData] = useState(null);
  const [columns, setColumns] = useState([]);
  const [numericCols, setNumericCols] = useState([]);
  const [categoricalCols, setCategoricalCols] = useState([]);
  const [stats, setStats] = useState({});
  const [correlations, setCorrelations] = useState([]);
  const [insights, setInsights] = useState("");
  const [executiveSummary, setExecutiveSummary] = useState("");
  const [dataIssues, setDataIssues] = useState([]);
  const [streamBuffer, setStreamBuffer] = useState("");
  const [fileName, setFileName] = useState("");
  const fileRef = useRef();

  const processFile = useCallback(async (file) => {
    setPhase("loading");
    setFileName(file.name);
    let rows = [];
    try {
      if (file.name.endsWith(".csv")) {
        await new Promise((res, rej) => Papa.parse(file, {
          header: true, skipEmptyLines: true,
          complete: r => { rows = r.data; res(); },
          error: rej
        }));
      } else {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf);
        rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: null });
      }
      if (!rows.length) throw new Error("No data found");
      setRawData(rows);
      runAnalysis(rows);
    } catch (e) {
      setPhase("error");
      setInsights("Failed to load file: " + e.message);
    }
  }, []);

  const runAnalysis = useCallback(async (rows) => {
    setPhase("analyzing");
    const cols = Object.keys(rows[0]);
    setColumns(cols);
    const issues = [];
    const colValues = {};
    cols.forEach(c => { colValues[c] = rows.map(r => r[c]); });

    const nullCounts = {};
    cols.forEach(c => {
      nullCounts[c] = colValues[c].filter(v => v == null || v === "").length;
      if (nullCounts[c] > 0) issues.push(`Column "${c}" has ${nullCounts[c]} missing value(s) (${((nullCounts[c] / rows.length) * 100).toFixed(1)}%)`);
    });
    setDataIssues(issues);

    const numCols = cols.filter(c => isNumericCol(colValues[c]));
    const catCols = cols.filter(c => !isNumericCol(colValues[c]));
    setNumericCols(numCols);
    setCategoricalCols(catCols);

    const statsMap = {};
    numCols.forEach(c => {
      statsMap[c] = computeStats(colValues[c].map(Number));
    });
    setStats(statsMap);

    const corrList = [];
    for (let i = 0; i < numCols.length; i++) {
      for (let j = i + 1; j < numCols.length; j++) {
        const r = pearsonCorr(colValues[numCols[i]].map(Number), colValues[numCols[j]].map(Number));
        if (r !== null) corrList.push({ a: numCols[i], b: numCols[j], r: +r.toFixed(3) });
      }
    }
    corrList.sort((a, b) => Math.abs(b.r) - Math.abs(a.r));
    setCorrelations(corrList);

    const topStats = numCols.slice(0, 6).map(c => {
      const s = statsMap[c];
      return s ? `- ${c}: mean=${s.mean.toFixed(2)}, std=${s.std.toFixed(2)}, min=${s.min.toFixed(2)}, max=${s.max.toFixed(2)}, outliers=${s.outliers.length}` : "";
    }).join("\n");
    const topCorr = corrList.slice(0, 5).map(c => `- ${c.a} vs ${c.b}: r=${c.r}`).join("\n");
    const catSummary = catCols.slice(0, 4).map(c => {
      const uniq = [...new Set(colValues[c].filter(Boolean))];
      return `- ${c}: ${uniq.length} unique values (${uniq.slice(0, 5).join(", ")}${uniq.length > 5 ? "..." : ""})`;
    }).join("\n");

    const prompt = `You are a senior data analyst. A dataset with ${rows.length} rows and ${cols.length} columns has been loaded: ${fileName || "uploaded file"}.

Columns: ${cols.join(", ")}
Numeric columns: ${numCols.join(", ")}
Categorical columns: ${catCols.join(", ")}

Descriptive statistics:
${topStats || "None"}

Top correlations:
${topCorr || "None"}

Categorical summaries:
${catSummary || "None"}

Data quality issues: ${issues.length ? issues.join("; ") : "None"}

Write a comprehensive but concise analysis report covering:
1. **Dataset overview** — shape, types, completeness
2. **Key statistical findings** — notable distributions, extremes, variability
3. **Correlation insights** — notable relationships
4. **Outliers & anomalies** — flag them
5. **Actionable insights** — what patterns or decisions does this data suggest?
6. **Executive summary** (clearly marked with "## Executive Summary") — 3-5 bullet points for a non-technical audience

Be specific, reference actual column names and numbers. Use clear section headers.`;

    try {
      setInsights(`
## Dataset Overview

Rows: ${rows.length}
Columns: ${cols.length}

Numeric Columns:
${numCols.join(", ") || "None"}

Categorical Columns:
${catCols.join(", ") || "None"}

## Data Quality

${issues.length ? issues.join("\n") : "No major data quality issues detected."}

## Correlation Insights

${topCorr || "No significant correlations found."}

## Statistical Summary

${topStats || "No numeric statistics available."}
`);

      setExecutiveSummary(`
## Executive Summary

- Dataset contains ${rows.length} rows and ${cols.length} columns.
- ${numCols.length} numeric columns detected.
- ${catCols.length} categorical columns detected.
- ${issues.length} data quality issue(s) identified.
- ${corrList.length} correlation pair(s) analyzed.
`);

      setPhase("done");
    } catch (e) {
      setInsights("Analysis failed: " + e.message);
      setPhase("error");
    }
  }, [fileName]);

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) processFile(file);
  }, [processFile]);

  const fmt = v => v == null ? "—" : typeof v === "number" ? v.toFixed(2) : v;

  const renderMarkdown = (text) => {
    if (!text) return null;
    return text.split("\n").map((line, i) => {
      if (line.startsWith("## ")) return <h2 key={i} style={{ fontSize: 16, fontWeight: 500, margin: "1.5rem 0 0.5rem", color: "var(--color-text-primary)" }}>{line.slice(3)}</h2>;
      if (line.startsWith("### ")) return <h3 key={i} style={{ fontSize: 14, fontWeight: 500, margin: "1rem 0 0.25rem", color: "var(--color-text-primary)" }}>{line.slice(4)}</h3>;
      if (line.startsWith("**") && line.endsWith("**")) return <p key={i} style={{ fontWeight: 500, margin: "0.25rem 0", fontSize: 14 }}>{line.slice(2, -2)}</p>;
      if (line.startsWith("- ") || line.startsWith("• ")) return <li key={i} style={{ fontSize: 14, color: "var(--color-text-primary)", margin: "0.2rem 0", lineHeight: 1.6 }}>{line.slice(2)}</li>;
      if (line.trim() === "") return <div key={i} style={{ height: "0.5rem" }} />;
      const bolded = line.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
      return <p key={i} style={{ fontSize: 14, lineHeight: 1.7, margin: "0.2rem 0", color: "var(--color-text-primary)" }} dangerouslySetInnerHTML={{ __html: bolded }} />;
    });
  };

  const topNumCols = numericCols.slice(0, 4);
  const histData = topNumCols.map(c => ({ col: c, bins: buildHistogram((rawData || []).map(r => Number(r[c]))) }));
  const scatterPairs = correlations.slice(0, 2).map(({ a, b }) => ({
    a, b, data: (rawData || []).slice(0, 200).map(r => ({ x: Number(r[a]), y: Number(r[b]) })).filter(p => !isNaN(p.x) && !isNaN(p.y))
  }));
  const catBarData = categoricalCols.slice(0, 2).map(c => {
    const freq = {};
    (rawData || []).forEach(r => { if (r[c]) freq[r[c]] = (freq[r[c]] || 0) + 1; });
    return { col: c, data: Object.entries(freq).sort((a, b) => b[1] - a[1]).slice(0, 10).map(([k, v]) => ({ name: k, count: v })) };
  });

  return (
    <div style={{ maxWidth: 1200, margin: "0 auto", padding: "1.5rem" }}>

      <div
        style={{
          textAlign: "center",
          padding: "4rem 1rem",
          marginBottom: "2rem",
          background:
            "radial-gradient(circle at center, rgba(56,189,248,0.12) 0%, rgba(15,23,42,0) 70%)"
        }}
      >
        <h1
          style={{
            fontSize: "3.5rem",
            fontWeight: "800",
            margin: 0,
            color: "#f8fafc",
            letterSpacing: "-0.04em"
          }}
        >
          AI Data Analysis Agent
        </h1>

        <p
          style={{
            color: "#94a3b8",
            fontSize: "1.15rem",
            maxWidth: "750px",
            margin: "1.25rem auto",
            lineHeight: 1.8
          }}
        >
          Transform raw spreadsheets into actionable insights, statistics,
          correlations and visualizations in seconds.
        </p>

        <div
          style={{
            display: "flex",
            justifyContent: "center",
            gap: "1rem",
            flexWrap: "wrap",
            marginTop: "1.5rem"
          }}
        >
          <span style={{ color: "#38bdf8" }}>✓ CSV & Excel</span>
          <span style={{ color: "#38bdf8" }}>✓ Automated EDA</span>
          <span style={{ color: "#38bdf8" }}>✓ Correlation Analysis</span>
          <span style={{ color: "#38bdf8" }}>✓ Interactive Charts</span>
        </div>
      </div>
      <h2 style={{ sr: "only", position: "absolute", opacity: 0, pointerEvents: "none" }}>AI data analysis agent</h2>

      {/* Drop zone */}
      {phase === "idle" && (
        <div
          onDrop={handleDrop}
          onDragOver={e => e.preventDefault()}
          onClick={() => fileRef.current.click()}
          style={{
            background: "#111827",
            border: "2px dashed #334155",
            borderRadius: "24px",
            padding: "4rem 2rem",
            textAlign: "center",
            cursor: "pointer",
            maxWidth: "700px",
            margin: "2rem auto",
            boxShadow: "0 20px 40px rgba(0,0,0,0.3)",
            transition: "all 0.3s ease"
          }}
        >
          <i className="ti ti-upload" style={{ fontSize: 28, color: "var(--color-text-tertiary)", display: "block", marginBottom: 12 }} aria-hidden="true" />
          <p style={{ fontWeight: 500, margin: "0 0 4px", fontSize: 15 }}>Upload Your dataset here</p>
          <p style={{ fontSize: 13, color: "var(--color-text-secondary)", margin: 0 }}>CSV • Excel • Automated Analysis • Instant Insights</p>
          <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls" style={{ display: "none" }} onChange={e => e.target.files[0] && processFile(e.target.files[0])} />
        </div>
      )}

      {/* Loading */}
      {(phase === "loading" || phase === "analyzing") && (
        <div style={{ textAlign: "center", padding: "3rem" }}>
          <div style={{
            width: 40, height: 40, border: "2px solid var(--color-border-secondary)",
            borderTopColor: "var(--color-text-info)", borderRadius: "50%",
            animation: "spin 0.8s linear infinite", margin: "0 auto 1rem"
          }} />
          <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
          <p style={{ fontWeight: 500, margin: "0 0 4px" }}>{phase === "loading" ? "Loading dataset…" : "Analysing with AI…"}</p>
          <p style={{ fontSize: 13, color: "var(--color-text-secondary)", margin: 0 }}>{fileName}</p>
          {phase === "analyzing" && streamBuffer && (
            <div style={{ textAlign: "left", marginTop: "1.5rem", background: "var(--color-background-secondary)", borderRadius: "var(--border-radius-md)", padding: "1rem", fontSize: 13, color: "var(--color-text-secondary)", maxHeight: 120, overflow: "hidden" }}>
              {streamBuffer.slice(-400)}
            </div>
          )}
        </div>
      )}

      {/* Error */}
      {phase === "error" && (
        <div style={{ background: "var(--color-background-danger)", borderRadius: "var(--border-radius-md)", padding: "1rem", fontSize: 14, color: "var(--color-text-danger)" }}>
          <i className="ti ti-alert-circle" style={{ marginRight: 8 }} aria-hidden="true" />{insights}
        </div>
      )}

      {/* Done */}
      {phase === "done" && rawData && (
        <div>
          {/* Overview cards */}
          <div
            style={{
              background: "#111827",
              border: "1px solid #1f2937",
              borderRadius: "24px",
              padding: "1.5rem",
              marginBottom: "1.5rem",
              boxShadow: "0 10px 25px rgba(0,0,0,0.25)"
            }}
          >
            <div
              style={{
                fontSize: "0.85rem",
                color: "#38bdf8",
                fontWeight: "600",
                marginBottom: "0.75rem"
              }}
            >
              AI INSIGHT
            </div>

            <div
              style={{
                fontSize: "1.1rem",
                color: "#f8fafc",
                lineHeight: 1.8
              }}
            >
              Dataset contains <strong>{rawData.length}</strong> rows,
              <strong> {columns.length}</strong> columns and
              <strong> {numericCols.length}</strong> numeric variables ready for analysis.
            </div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(130px, 1fr))", gap: 12, marginBottom: "1.5rem" }}>
            <StatCard label="Rows" value={rawData.length.toLocaleString()} />
            <StatCard label="Columns" value={columns.length} />
            <StatCard label="Numeric" value={numericCols.length} />
            <StatCard label="Categorical" value={categoricalCols.length} />
            <StatCard label="Data issues" value={dataIssues.length} sub={dataIssues.length ? "see report" : "none found"} />
          </div>

          {/* Data issues */}
          {dataIssues.length > 0 && (
            <div style={{ background: "var(--color-background-warning)", borderRadius: "var(--border-radius-md)", padding: "0.75rem 1rem", marginBottom: "1rem" }}>
              <p style={{ fontWeight: 500, fontSize: 13, color: "var(--color-text-warning)", margin: "0 0 4px" }}>
                <i className="ti ti-alert-triangle" style={{ marginRight: 6 }} aria-hidden="true" />Data quality notes
              </p>
              {dataIssues.map((d, i) => <p key={i} style={{ fontSize: 12, color: "var(--color-text-warning)", margin: "2px 0" }}>• {d}</p>)}
            </div>
          )}

          {/* Charts */}
          {histData.map(({ col, bins }, ci) => (
            <ChartBlock key={col} title={`Distribution — ${col}`}>
              <ResponsiveContainer width="100%" height={250}>
                <BarChart data={bins} margin={{ top: 0, right: 8, bottom: 0, left: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border-tertiary)" />
                  <XAxis
                    dataKey="range"
                    tick={{ fontSize: 11, fill: "#cbd5e1" }}
                    interval={0}
                  />
                  <YAxis
                    tick={{ fontSize: 11, fill: "#cbd5e1" }}
                  />
                  <Tooltip contentStyle={{ fontSize: 12 }} />
                  <Bar dataKey="count" fill={COLORS[ci % COLORS.length]} radius={[2, 2, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </ChartBlock>
          ))}

          {scatterPairs.map(({ a, b, data }, ci) => (
            <ChartBlock key={`${a}-${b}`} title={`Scatter — ${a} vs ${b} (r=${correlations[ci]?.r})`}>
              <ResponsiveContainer width="100%" height={300}>
                <ScatterChart margin={{ top: 10, right: 20, bottom: 40, left: 20 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border-tertiary)" />
                  <XAxis
                    dataKey="x"
                    name={a}
                    type="number"
                    tick={{
                      fontSize: 13,
                      fill: "#e2e8f0"
                    }}
                    axisLine={{ stroke: "#64748b" }}
                    tickLine={{ stroke: "#64748b" }}
                    domain={["dataMin", "dataMax"]}
                    allowDecimals={false}
                    tickCount={5}
                  />
                  <YAxis
                    dataKey="y"
                    name={b}
                    type="number"
                    tick={{
                      fontSize: 13,
                      fill: "#e2e8f0"
                    }}
                    axisLine={{ stroke: "#64748b" }}
                    tickLine={{ stroke: "#64748b" }}
                    domain={["dataMin", "dataMax"]}
                    tickCount={5}
                  />
                  <Tooltip cursor={{ strokeDasharray: "3 3" }} contentStyle={{ fontSize: 12 }} />
                  <Scatter
                    data={data}
                    fill={COLORS[ci + 2]}
                    fillOpacity={0.9}
                  />
                </ScatterChart>
              </ResponsiveContainer>
            </ChartBlock>
          ))}

          {catBarData.map(({ col, data }, ci) => (
            <ChartBlock key={col} title={`Category breakdown — ${col}`}>
              <ResponsiveContainer width="100%" height={180}>
                <BarChart data={data} layout="vertical" margin={{ top: 0, right: 8, bottom: 0, left: 60 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border-tertiary)" />
                  <XAxis type="number" tick={{ fontSize: 10, fill: "var(--color-text-tertiary)" }} />
                  <YAxis
                    dataKey="name"
                    type="category"
                    tick={{ fontSize: 12, fill: "#cbd5e1" }}
                    width={180}
                  />
                  <Tooltip contentStyle={{ fontSize: 12 }} />
                  <Bar dataKey="count" fill={COLORS[(ci + 4) % COLORS.length]} radius={[0, 2, 2, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </ChartBlock>
          ))}

          {/* Correlations table */}
          {correlations.length > 0 && (
            <ChartBlock title="Correlation matrix (top pairs)">
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr auto", gap: "4px 12px", fontSize: 13 }}>
                <span style={{ color: "var(--color-text-tertiary)", fontSize: 11, textTransform: "uppercase" }}>Column A</span>
                <span style={{ color: "var(--color-text-tertiary)", fontSize: 11, textTransform: "uppercase" }}>Column B</span>
                <span style={{ color: "var(--color-text-tertiary)", fontSize: 11, textTransform: "uppercase" }}>r</span>
                {correlations.slice(0, 8).map(({ a, b, r }, i) => {
                  const abs = Math.abs(r);
                  const color = abs > 0.7 ? "var(--color-text-danger)" : abs > 0.4 ? "var(--color-text-warning)" : "var(--color-text-secondary)";
                  return [
                    <span key={`a${i}`} style={{ borderTop: "0.5px solid var(--color-border-tertiary)", padding: "4px 0" }}>{a}</span>,
                    <span key={`b${i}`} style={{ borderTop: "0.5px solid var(--color-border-tertiary)", padding: "4px 0" }}>{b}</span>,
                    <span key={`r${i}`} style={{ borderTop: "0.5px solid var(--color-border-tertiary)", padding: "4px 0", color, fontWeight: 500 }}>{r}</span>
                  ];
                })}
              </div>
            </ChartBlock>
          )}

          {/* Descriptive stats */}
          {numericCols.length > 0 && (
            <ChartBlock title="Descriptive statistics">
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", fontSize: 12, borderCollapse: "collapse", tableLayout: "fixed" }}>
                  <thead>
                    <tr>
                      {["Column", "Mean", "Median", "Std", "Min", "Max", "Q1", "Q3", "Outliers"].map(h => (
                        <th key={h} style={{ textAlign: "left", padding: "4px 8px", color: "var(--color-text-tertiary)", fontWeight: 400, borderBottom: "0.5px solid var(--color-border-tertiary)", whiteSpace: "nowrap" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {numericCols.map(c => {
                      const s = stats[c];
                      if (!s) return null;
                      return (
                        <tr key={c}>
                          {[c, s.mean, s.median, s.std, s.min, s.max, s.q1, s.q3].map((v, i) => (
                            <td key={i} style={{ padding: "4px 8px", borderBottom: "0.5px solid var(--color-border-tertiary)", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                              {i === 0 ? v : fmt(v)}
                            </td>
                          ))}
                          <td style={{ padding: "4px 8px", borderBottom: "0.5px solid var(--color-border-tertiary)" }}>
                            <span style={{ background: s.outliers.length ? "var(--color-background-danger)" : "var(--color-background-success)", color: s.outliers.length ? "var(--color-text-danger)" : "var(--color-text-success)", borderRadius: 4, padding: "1px 6px", fontSize: 11 }}>{s.outliers.length}</span>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </ChartBlock>
          )}

          {/* AI Analysis Report */}
          {insights && (
            <div style={{ border: "0.5px solid var(--color-border-tertiary)", borderRadius: "var(--border-radius-lg)", padding: "1.25rem", marginBottom: "1rem" }}>
              <p style={{ fontWeight: 500, margin: "0 0 1rem", fontSize: 14, display: "flex", alignItems: "center", gap: 8, color: "var(--color-text-secondary)" }}>
                <i className="ti ti-sparkles" aria-hidden="true" /> AI analysis report
              </p>
              <div>{renderMarkdown(insights)}</div>
            </div>
          )}

          {/* Executive Summary */}
          {executiveSummary && (
            <div style={{ background: "var(--color-background-info)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: "var(--border-radius-lg)", padding: "1.25rem" }}>
              <p style={{ fontWeight: 500, margin: "0 0 0.75rem", fontSize: 14, color: "var(--color-text-info)", display: "flex", alignItems: "center", gap: 8 }}>
                <i className="ti ti-file-text" aria-hidden="true" /> Executive summary
              </p>
              <div>{renderMarkdown(executiveSummary)}</div>
            </div>
          )}

          {/* Reset */}
          <button onClick={() => { setPhase("idle"); setRawData(null); setInsights(""); setExecutiveSummary(""); setStreamBuffer(""); }} style={{ marginTop: "1.5rem", fontSize: 13 }}>
            <i className="ti ti-refresh" style={{ marginRight: 6 }} aria-hidden="true" />Analyse another file
          </button>
        </div>
      )}
    </div>
  );
}
