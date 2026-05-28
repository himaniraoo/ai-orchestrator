import { useState, useRef, useEffect } from "react";
import { marked } from "marked";

const API = "http://localhost:8000";

const SPECIALTIES = [
  { value: "", label: "All Specialties" },
  { value: "Medical Oncology", label: "Medical Oncology" },
  { value: "Pulmonology", label: "Pulmonology" },
  { value: "Radiation Oncology", label: "Radiation Oncology" },
  { value: "Thoracic Surgery", label: "Thoracic Surgery" },
];

const STATES = [
  "CA",
  "CO",
  "CT",
  "GA",
  "IL",
  "MA",
  "MD",
  "MI",
  "MN",
  "NJ",
  "NY",
  "OH",
  "PA",
  "TN",
  "TX",
  "WA",
];

const ICD10_OPTIONS = [
  { value: "C341", label: "C341 — Upper Lobe NSCLC" },
  { value: "C342", label: "C342 — Middle Lobe NSCLC" },
];

const VOLUME_TIERS = [
  { value: "", label: "All Volumes" },
  { value: "low", label: "Low" },
  { value: "high", label: "High" },
  { value: "very_high", label: "Very High" },
];

const TOOL_LABELS = {
  get_physician_data: "Fetching physician data",
  call_ppt_agent: "Generating PowerPoint deck",
  call_excel_agent: "Building Excel workbook",
  call_report_agent: "Writing market access report",
  call_sandbox_agent: "Running Python analysis",
};

const TOOL_ICONS = {
  get_physician_data: "⬡",
  call_ppt_agent: "◈",
  call_excel_agent: "◉",
  call_report_agent: "◎",
  call_sandbox_agent: "◊",
};

export default function App() {
  const [query, setQuery] = useState("");
  const [specialty, setSpecialty] = useState("");
  const [selectedStates, setSelectedStates] = useState([]);
  const [icd10Codes, setIcd10Codes] = useState([]);
  const [volumeTier, setVolumeTier] = useState("");
  const [isStreaming, setIsStreaming] = useState(false);
  const [traceSteps, setTraceSteps] = useState([]);
  const [artifacts, setArtifacts] = useState([]);
  const [reportMarkdown, setReportMarkdown] = useState(null);
  const [sandboxResult, setSandboxResult] = useState(null);
  const [summary, setSummary] = useState(null);
  const [error, setError] = useState(null);
  const [hasRun, setHasRun] = useState(false);
  const esRef = useRef(null);
  const traceEndRef = useRef(null);

  useEffect(() => {
    if (traceSteps.length > 0) {
      traceEndRef.current?.scrollIntoView({ behavior: "smooth" });
    }
  }, [traceSteps]);

  useEffect(() => {
    return () => {
      esRef.current?.close();
    };
  }, []);

  const toggleState = (s) => {
    setSelectedStates((prev) =>
      prev.includes(s) ? prev.filter((x) => x !== s) : [...prev, s]
    );
  };

  const toggleIcd10 = (code) => {
    setIcd10Codes((prev) =>
      prev.includes(code) ? prev.filter((x) => x !== code) : [...prev, code]
    );
  };

  const reset = () => {
    setTraceSteps([]);
    setArtifacts([]);
    setReportMarkdown(null);
    setSandboxResult(null);
    setSummary(null);
    setError(null);
  };

  const handleSubmit = () => {
    if (!query.trim() || isStreaming) return;
    reset();
    setHasRun(true);
    setIsStreaming(true);

    // Build query string
    const params = new URLSearchParams();
    params.set("query", query.trim());
    if (specialty) params.set("specialty", specialty);
    selectedStates.forEach((s) => params.append("states", s));
    icd10Codes.forEach((c) => params.append("icd10_codes", c));
    if (volumeTier) params.set("volume_threshold", volumeTier);

    const url = `${API}/stream?${params.toString()}`;
    const es = new EventSource(url);
    esRef.current = es;

    es.onmessage = (e) => {
      try {
        const { event, data } = JSON.parse(e.data);

        if (event === "trace") {
          setTraceSteps((prev) => {
            const key = `${data.step}-${data.tool}`;
            const existing = prev.findIndex((s) => s.key === key);
            const entry = { ...data, key };
            if (existing >= 0) {
              const updated = [...prev];
              updated[existing] = entry;
              return updated;
            }
            return [...prev, entry];
          });
        }

        if (event === "artifact") {
          setArtifacts((prev) => {
            const exists = prev.some(
              (a) => a.artifact_id === data.artifact_id
            );
          
            return exists ? prev : [...prev, data];
          });
        }

        if (event === "report") {
          setReportMarkdown(data.markdown);
        }

        if (event === "sandbox") {
          setSandboxResult(data);
        }

        if (event === "summary") {
          setSummary(data.text);
        }

        if (event === "error") {
          setError(data.message);
          setIsStreaming(false);
          es.close();
        }

        if (event === "done") {
          setIsStreaming(false);
          es.close();
        }
      } catch {
        // ignore parse errors
      }
    };

    es.onerror = () => {
      setError("Connection to backend lost. Is the server running?");
      setIsStreaming(false);
      es.close();
    };
  };

  const handleKeyDown = (e) => {
    if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) {
      handleSubmit();
    }
  };

  const artifactLabel = (type) => {
    const map = {
      ppt: "PowerPoint Deck",
      excel: "Excel Workbook",
      report: "Word Report",
    };
    return map[type] || type.toUpperCase();
  };

  const artifactExt = (id = "") =>
    id.includes(".") ? id.split(".").pop().toUpperCase() : "";

  return (
    <div style={styles.root}>
      {/* Background grid */}
      <div style={styles.gridBg} />

      {/* Header */}
      <header style={styles.header}>
        <div style={styles.headerInner}>
          <div style={styles.logo}>
            <span style={styles.logoMark}>◈</span>
            <span style={styles.logoText}>DocNexus</span>
            <span style={styles.logoBadge}>AI</span>
          </div>
          <p style={styles.tagline}>
            Physician intelligence, orchestrated.
          </p>
        </div>
      </header>

      <main style={styles.main}>

        {/* ── QUERY INPUT ── */}
        <section style={styles.section}>
          <label style={styles.sectionLabel}>Query</label>
          <div style={styles.queryBox}>
            <textarea
              style={styles.textarea}
              placeholder={`Ask anything about the physician landscape...\n\nExamples:\n"Give me a slide deck and Excel breakdown of high-volume NSCLC oncologists in CA and NY"\n"Write a market access report on NSCLC physician density in the Northeast"\n"Plot which states have the highest concentration of high-volume NSCLC prescribers"`}
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              onKeyDown={handleKeyDown}
              rows={5}
              disabled={isStreaming}
            />
            <div style={styles.queryFooter}>
              <span style={styles.hint}>⌘ + Enter to run</span>
              <button
                style={{
                  ...styles.submitBtn,
                  ...(isStreaming ? styles.submitBtnDisabled : {}),
                }}
                onClick={handleSubmit}
                disabled={isStreaming || !query.trim()}
              >
                {isStreaming ? (
                  <span style={styles.btnInner}>
                    <span style={styles.spinner} />
                    Running...
                  </span>
                ) : (
                  <span style={styles.btnInner}>
                    <span>Run Query</span>
                    <span style={styles.btnArrow}>→</span>
                  </span>
                )}
              </button>
            </div>
          </div>
        </section>

        {/* ── PREFERENCES ── */}
        <section style={styles.section}>
          <label style={styles.sectionLabel}>
            Preferences
            <span style={styles.optional}>optional filters</span>
          </label>
          <div style={styles.prefsGrid}>

            {/* Specialty */}
            <div style={styles.prefGroup}>
              <span style={styles.prefLabel}>Specialty</span>
              <select
                style={styles.select}
                value={specialty}
                onChange={(e) => setSpecialty(e.target.value)}
                disabled={isStreaming}
              >
                {SPECIALTIES.map((s) => (
                  <option key={s.value} value={s.value}>{s.label}</option>
                ))}
              </select>
            </div>

            {/* Volume Tier */}
            <div style={styles.prefGroup}>
              <span style={styles.prefLabel}>Volume Threshold</span>
              <select
                style={styles.select}
                value={volumeTier}
                onChange={(e) => setVolumeTier(e.target.value)}
                disabled={isStreaming}
              >
                {VOLUME_TIERS.map((t) => (
                  <option key={t.value} value={t.value}>{t.label}</option>
                ))}
              </select>
            </div>

            {/* ICD-10 */}
            <div style={styles.prefGroup}>
              <span style={styles.prefLabel}>ICD-10 Codes</span>
              <div style={styles.checkboxGroup}>
                {ICD10_OPTIONS.map((opt) => (
                  <label key={opt.value} style={styles.checkboxLabel}>
                    <input
                      type="checkbox"
                      checked={icd10Codes.includes(opt.value)}
                      onChange={() => toggleIcd10(opt.value)}
                      disabled={isStreaming}
                      style={styles.checkbox}
                    />
                    <span>{opt.label}</span>
                  </label>
                ))}
              </div>
            </div>

            {/* States */}
            <div style={{ ...styles.prefGroup, gridColumn: "1 / -1" }}>
              <span style={styles.prefLabel}>States</span>
              <div style={styles.stateGrid}>
                {STATES.map((s) => (
                  <button
                    key={s}
                    style={{
                      ...styles.stateChip,
                      ...(selectedStates.includes(s) ? styles.stateChipActive : {}),
                    }}
                    onClick={() => toggleState(s)}
                    disabled={isStreaming}
                  >
                    {s}
                  </button>
                ))}
              </div>
            </div>

          </div>
        </section>

        {/* ── LIVE AGENT TRACE ── */}
        {hasRun && (
          <section style={styles.section}>
            <label style={styles.sectionLabel}>
              Agent Trace
              {isStreaming && <span style={styles.liveBadge}>● LIVE</span>}
            </label>
            <div style={styles.traceBox}>
              {traceSteps.length === 0 && isStreaming && (
                <div style={styles.traceWaiting}>
                  <span style={styles.spinner} />
                  <span style={{ color: "#6B7280", marginLeft: 10 }}>
                    Connecting to orchestrator...
                  </span>
                </div>
              )}
              {traceSteps.map((step) => (
                <div key={step.key} style={styles.traceStep}>
                  <div style={styles.traceLeft}>
                    <span style={styles.traceIcon}>
                      {TOOL_ICONS[step.tool] || "◆"}
                    </span>
                    <div style={styles.traceInfo}>
                      <span style={styles.traceLabel}>
                        {TOOL_LABELS[step.tool] || step.tool}
                      </span>
                      {step.args_summary && Object.keys(step.args_summary).length > 0 && (
                        <span style={styles.traceArgs}>
                          {Object.entries(step.args_summary)
                            .map(([k, v]) =>
                              `${k}: ${Array.isArray(v) ? v.join(", ") : v}`
                            )
                            .join(" · ")}
                        </span>
                      )}
                    </div>
                  </div>
                  <div style={styles.traceRight}>
                    {step.status === "running" ? (
                      <span style={styles.traceRunning}>
                        <span style={styles.spinnerSmall} /> running
                      </span>
                    ) : (
                      <span style={styles.traceDone}>
                        ✓{" "}
                        {step.elapsed_ms
                          ? `${(step.elapsed_ms / 1000).toFixed(1)}s`
                          : "done"}
                      </span>
                    )}
                  </div>
                </div>
              ))}
              <div ref={traceEndRef} />
            </div>
          </section>
        )}

        {/* ── RESULTS ── */}
        {(artifacts.length > 0 || reportMarkdown || sandboxResult || summary || error) && (
          <section style={styles.section}>
            <label style={styles.sectionLabel}>Results</label>

            {/* Error */}
            {error && (
              <div style={styles.errorBox}>
                <span style={styles.errorIcon}>⚠</span>
                <span>{error}</span>
              </div>
            )}

            {/* Summary */}
            {summary && (
              <div style={styles.summaryBox}>
                <span style={styles.summaryIcon}>◎</span>
                <p style={styles.summaryText}>{summary}</p>
              </div>
            )}

            {/* Artifacts */}
            {artifacts.length > 0 && (
              <div style={styles.artifactsRow}>
                {artifacts.map((a, i) => (
                  <a
                    key={i}
                    href={`${API}${a.download_url}`}
                    download
                    style={styles.artifactCard}
                  >
                    <span style={styles.artifactIcon}>
                      {a.type === "ppt" ? "◈" : a.type === "excel" ? "◉" : "◎"}
                    </span>
                    <div style={styles.artifactInfo}>
                      <span style={styles.artifactName}>
                        {artifactLabel(a.type)}
                      </span>
                      <span style={styles.artifactExt}>
                        .{artifactExt(a.artifact_id)} · Click to download
                      </span>
                    </div>
                    <span style={styles.artifactArrow}>↓</span>
                  </a>
                ))}
              </div>
            )}

            {/* Sandbox Result */}
            {sandboxResult && (
              <div style={styles.sandboxBox}>
                <div style={styles.sandboxHeader}>
                  <span style={styles.sandboxTitle}>◊ Analysis Output</span>
                  <span style={styles.sandboxMeta}>
                    {sandboxResult.attempts > 1
                      ? `✓ self-corrected after ${sandboxResult.attempts} attempts`
                      : "✓ executed successfully"}
                  </span>
                </div>

                {/* Generated code */}
                {sandboxResult.code && (
                  <div style={styles.codeBlock}>
                    <div style={styles.codeHeader}>Generated Python</div>
                    <pre style={styles.codePre}>
                      <code>{sandboxResult.code}</code>
                    </pre>
                  </div>
                )}

                {/* Text output */}
                {sandboxResult.output && (
                  <div style={styles.outputBlock}>
                    <div style={styles.codeHeader}>Output</div>
                    <pre style={styles.outputPre}>{sandboxResult.output}</pre>
                  </div>
                )}

                {/* Chart */}
                {sandboxResult.chart_base64 && (
                  <div style={styles.chartBlock}>
                    <div style={styles.codeHeader}>Chart</div>
                    <img
                      src={`data:image/png;base64,${sandboxResult.chart_base64}`}
                      alt="Analysis chart"
                      style={styles.chartImg}
                    />
                  </div>
                )}

                {/* Failed state */}
                {sandboxResult.status === "failed" && (
                  <div style={styles.sandboxError}>
                    <span>⚠ Analysis failed after 2 attempts</span>
                    {sandboxResult.error && (
                      <pre style={styles.outputPre}>{sandboxResult.error}</pre>
                    )}
                  </div>
                )}
              </div>
            )}

            {/* Report Markdown */}
            {reportMarkdown && (
              <div style={styles.reportBox}>
                <div style={styles.reportHeader}>
                  <span style={styles.sandboxTitle}>◎ Market Access Report</span>
                  {artifacts.find((a) => a.type === "report") && (
                    <a
                      href={`${API}${artifacts.find((a) => a.type === "report").download_url}`}
                      download
                      style={styles.reportDownload}
                    >
                      ↓ Download .docx
                    </a>
                  )}
                </div>
                <div
                  style={styles.reportContent}
                  dangerouslySetInnerHTML={{
                    __html: marked.parse(reportMarkdown),
                  }}
                />
              </div>
            )}

          </section>
        )}

      </main>

      <footer style={styles.footer}>
        DocNexus · AI Engineer Intern Assignment · Powered by Gemini 2.5 Flash
      </footer>
    </div>
  );
}

// ─────────────────────────────────────────────
// STYLES
// ─────────────────────────────────────────────

const styles = {
  root: {
    minHeight: "100vh",
    backgroundColor: "#080E1A",
    color: "#E8EDF5",
    fontFamily: "'DM Mono', 'Fira Code', 'Courier New', monospace",
    position: "relative",
    overflowX: "hidden",
  },
  gridBg: {
    position: "fixed",
    inset: 0,
    backgroundImage: `
      linear-gradient(rgba(0,168,232,0.04) 1px, transparent 1px),
      linear-gradient(90deg, rgba(0,168,232,0.04) 1px, transparent 1px)
    `,
    backgroundSize: "40px 40px",
    pointerEvents: "none",
    zIndex: 0,
  },
  header: {
    borderBottom: "1px solid rgba(0,168,232,0.15)",
    padding: "24px 0",
    position: "relative",
    zIndex: 1,
  },
  headerInner: {
    maxWidth: 860,
    margin: "0 auto",
    padding: "0 24px",
    display: "flex",
    alignItems: "center",
    gap: 24,
  },
  logo: {
    display: "flex",
    alignItems: "center",
    gap: 10,
  },
  logoMark: {
    fontSize: 28,
    color: "#00A8E8",
    lineHeight: 1,
  },
  logoText: {
    fontSize: 22,
    fontWeight: 700,
    letterSpacing: "-0.02em",
    color: "#E8EDF5",
  },
  logoBadge: {
    fontSize: 10,
    fontWeight: 700,
    letterSpacing: "0.1em",
    color: "#00A8E8",
    border: "1px solid rgba(0,168,232,0.4)",
    padding: "2px 6px",
    borderRadius: 3,
  },
  tagline: {
    fontSize: 13,
    color: "#4B5563",
    margin: 0,
    letterSpacing: "0.02em",
  },
  main: {
    maxWidth: 860,
    margin: "0 auto",
    padding: "40px 24px 80px",
    position: "relative",
    zIndex: 1,
    display: "flex",
    flexDirection: "column",
    gap: 32,
  },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: 12,
  },
  sectionLabel: {
    fontSize: 11,
    fontWeight: 700,
    letterSpacing: "0.12em",
    textTransform: "uppercase",
    color: "#4B5563",
    display: "flex",
    alignItems: "center",
    gap: 10,
  },
  optional: {
    fontSize: 10,
    fontWeight: 400,
    letterSpacing: "0.06em",
    color: "#374151",
    textTransform: "none",
  },
  queryBox: {
    border: "1px solid rgba(0,168,232,0.2)",
    borderRadius: 8,
    overflow: "hidden",
    backgroundColor: "rgba(0,168,232,0.03)",
    transition: "border-color 0.2s",
  },
  textarea: {
    width: "100%",
    backgroundColor: "transparent",
    border: "none",
    outline: "none",
    color: "#E8EDF5",
    fontFamily: "'DM Mono', 'Fira Code', monospace",
    fontSize: 14,
    lineHeight: 1.7,
    padding: "20px 20px 12px",
    resize: "none",
    boxSizing: "border-box",
  },
  queryFooter: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "10px 16px",
    borderTop: "1px solid rgba(0,168,232,0.1)",
  },
  hint: {
    fontSize: 11,
    color: "#374151",
    letterSpacing: "0.04em",
  },
  submitBtn: {
    backgroundColor: "#00A8E8",
    color: "#080E1A",
    border: "none",
    borderRadius: 6,
    padding: "10px 20px",
    fontSize: 13,
    fontWeight: 700,
    fontFamily: "'DM Mono', monospace",
    letterSpacing: "0.04em",
    cursor: "pointer",
    transition: "opacity 0.2s, transform 0.1s",
  },
  submitBtnDisabled: {
    opacity: 0.6,
    cursor: "not-allowed",
  },
  btnInner: {
    display: "flex",
    alignItems: "center",
    gap: 8,
  },
  btnArrow: {
    fontSize: 16,
  },
  spinner: {
    display: "inline-block",
    width: 12,
    height: 12,
    border: "2px solid rgba(8,14,26,0.3)",
    borderTop: "2px solid #080E1A",
    borderRadius: "50%",
    animation: "spin 0.8s linear infinite",
  },
  spinnerSmall: {
    display: "inline-block",
    width: 8,
    height: 8,
    border: "1.5px solid rgba(0,168,232,0.3)",
    borderTop: "1.5px solid #00A8E8",
    borderRadius: "50%",
    animation: "spin 0.8s linear infinite",
    marginRight: 4,
  },
  prefsGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: 16,
    padding: 20,
    border: "1px solid rgba(255,255,255,0.06)",
    borderRadius: 8,
    backgroundColor: "rgba(255,255,255,0.02)",
  },
  prefGroup: {
    display: "flex",
    flexDirection: "column",
    gap: 8,
  },
  prefLabel: {
    fontSize: 11,
    color: "#6B7280",
    letterSpacing: "0.06em",
    textTransform: "uppercase",
  },
  select: {
    backgroundColor: "rgba(0,168,232,0.05)",
    border: "1px solid rgba(0,168,232,0.15)",
    borderRadius: 6,
    color: "#E8EDF5",
    fontFamily: "'DM Mono', monospace",
    fontSize: 13,
    padding: "8px 12px",
    outline: "none",
    cursor: "pointer",
  },
  checkboxGroup: {
    display: "flex",
    flexDirection: "column",
    gap: 8,
  },
  checkboxLabel: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    fontSize: 13,
    color: "#9CA3AF",
    cursor: "pointer",
  },
  checkbox: {
    accentColor: "#00A8E8",
    width: 14,
    height: 14,
    cursor: "pointer",
  },
  stateGrid: {
    display: "flex",
    flexWrap: "wrap",
    gap: 8,
  },
  stateChip: {
    backgroundColor: "transparent",
    border: "1px solid rgba(255,255,255,0.1)",
    borderRadius: 4,
    color: "#6B7280",
    fontFamily: "'DM Mono', monospace",
    fontSize: 12,
    fontWeight: 600,
    padding: "5px 10px",
    cursor: "pointer",
    transition: "all 0.15s",
    letterSpacing: "0.04em",
  },
  stateChipActive: {
    backgroundColor: "rgba(0,168,232,0.15)",
    border: "1px solid rgba(0,168,232,0.5)",
    color: "#00A8E8",
  },
  traceBox: {
    border: "1px solid rgba(255,255,255,0.06)",
    borderRadius: 8,
    overflow: "hidden",
    backgroundColor: "rgba(0,0,0,0.2)",
  },
  traceWaiting: {
    display: "flex",
    alignItems: "center",
    padding: "20px 24px",
  },
  traceStep: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "14px 20px",
    borderBottom: "1px solid rgba(255,255,255,0.04)",
    transition: "background 0.2s",
  },
  traceLeft: {
    display: "flex",
    alignItems: "center",
    gap: 14,
    flex: 1,
  },
  traceIcon: {
    fontSize: 18,
    color: "#00A8E8",
    width: 24,
    textAlign: "center",
    flexShrink: 0,
  },
  traceInfo: {
    display: "flex",
    flexDirection: "column",
    gap: 2,
  },
  traceLabel: {
    fontSize: 13,
    color: "#E8EDF5",
    fontWeight: 500,
  },
  traceArgs: {
    fontSize: 11,
    color: "#4B5563",
    letterSpacing: "0.02em",
  },
  traceRight: {
    flexShrink: 0,
    marginLeft: 16,
  },
  traceRunning: {
    display: "flex",
    alignItems: "center",
    fontSize: 12,
    color: "#00A8E8",
    letterSpacing: "0.04em",
  },
  traceDone: {
    fontSize: 12,
    color: "#10B981",
    letterSpacing: "0.04em",
  },
  liveBadge: {
    fontSize: 10,
    color: "#EF4444",
    letterSpacing: "0.1em",
    animation: "pulse 1.5s ease-in-out infinite",
    marginLeft: 8,
  },
  errorBox: {
    display: "flex",
    alignItems: "center",
    gap: 12,
    padding: "16px 20px",
    backgroundColor: "rgba(239,68,68,0.08)",
    border: "1px solid rgba(239,68,68,0.2)",
    borderRadius: 8,
    fontSize: 13,
    color: "#FCA5A5",
  },
  errorIcon: {
    fontSize: 18,
    color: "#EF4444",
    flexShrink: 0,
  },
  summaryBox: {
    display: "flex",
    gap: 14,
    padding: "16px 20px",
    backgroundColor: "rgba(0,168,232,0.05)",
    border: "1px solid rgba(0,168,232,0.15)",
    borderRadius: 8,
  },
  summaryIcon: {
    fontSize: 18,
    color: "#00A8E8",
    flexShrink: 0,
    marginTop: 1,
  },
  summaryText: {
    fontSize: 13,
    color: "#9CA3AF",
    margin: 0,
    lineHeight: 1.7,
  },
  artifactsRow: {
    display: "flex",
    flexDirection: "column",
    gap: 10,
  },
  artifactCard: {
    display: "flex",
    alignItems: "center",
    gap: 16,
    padding: "16px 20px",
    backgroundColor: "rgba(0,168,232,0.05)",
    border: "1px solid rgba(0,168,232,0.2)",
    borderRadius: 8,
    textDecoration: "none",
    color: "#E8EDF5",
    transition: "background 0.2s, border-color 0.2s",
    cursor: "pointer",
  },
  artifactIcon: {
    fontSize: 22,
    color: "#00A8E8",
    flexShrink: 0,
  },
  artifactInfo: {
    display: "flex",
    flexDirection: "column",
    gap: 2,
    flex: 1,
  },
  artifactName: {
    fontSize: 14,
    fontWeight: 600,
    color: "#E8EDF5",
  },
  artifactExt: {
    fontSize: 11,
    color: "#4B5563",
    letterSpacing: "0.04em",
  },
  artifactArrow: {
    fontSize: 18,
    color: "#00A8E8",
    flexShrink: 0,
  },
  sandboxBox: {
    border: "1px solid rgba(255,255,255,0.08)",
    borderRadius: 8,
    overflow: "hidden",
  },
  sandboxHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "14px 20px",
    backgroundColor: "rgba(0,0,0,0.3)",
    borderBottom: "1px solid rgba(255,255,255,0.06)",
  },
  sandboxTitle: {
    fontSize: 13,
    fontWeight: 600,
    color: "#E8EDF5",
    letterSpacing: "0.04em",
  },
  sandboxMeta: {
    fontSize: 11,
    color: "#10B981",
    letterSpacing: "0.04em",
  },
  sandboxError: {
    padding: "16px 20px",
    color: "#FCA5A5",
    fontSize: 13,
  },
  codeBlock: {
    borderBottom: "1px solid rgba(255,255,255,0.06)",
  },
  outputBlock: {
    borderBottom: "1px solid rgba(255,255,255,0.06)",
  },
  chartBlock: {
    padding: "0 0 0 0",
  },
  codeHeader: {
    fontSize: 10,
    fontWeight: 700,
    letterSpacing: "0.1em",
    textTransform: "uppercase",
    color: "#4B5563",
    padding: "10px 20px 6px",
    backgroundColor: "rgba(0,0,0,0.2)",
  },
  codePre: {
    margin: 0,
    padding: "16px 20px",
    overflowX: "auto",
    fontSize: 12,
    lineHeight: 1.6,
    color: "#A5F3FC",
    backgroundColor: "rgba(0,0,0,0.15)",
  },
  outputPre: {
    margin: 0,
    padding: "16px 20px",
    overflowX: "auto",
    fontSize: 12,
    lineHeight: 1.6,
    color: "#D1FAE5",
    backgroundColor: "rgba(0,0,0,0.1)",
  },
  chartImg: {
    width: "100%",
    display: "block",
    borderTop: "1px solid rgba(255,255,255,0.06)",
  },
  reportBox: {
    border: "1px solid rgba(255,255,255,0.08)",
    borderRadius: 8,
    overflow: "hidden",
  },
  reportHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "14px 20px",
    backgroundColor: "rgba(0,0,0,0.3)",
    borderBottom: "1px solid rgba(255,255,255,0.06)",
  },
  reportDownload: {
    fontSize: 12,
    color: "#00A8E8",
    textDecoration: "none",
    letterSpacing: "0.04em",
    border: "1px solid rgba(0,168,232,0.3)",
    padding: "4px 10px",
    borderRadius: 4,
  },
  reportContent: {
    padding: "24px 28px",
    fontSize: 14,
    lineHeight: 1.8,
    color: "#9CA3AF",
    maxHeight: 600,
    overflowY: "auto",
  },
  footer: {
    textAlign: "center",
    padding: "24px",
    fontSize: 11,
    color: "#1F2937",
    letterSpacing: "0.06em",
    borderTop: "1px solid rgba(255,255,255,0.04)",
    position: "relative",
    zIndex: 1,
  },
};
