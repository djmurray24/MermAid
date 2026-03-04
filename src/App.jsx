import { useState, useEffect, useRef, useCallback } from "react";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { loginRequest } from "./authConfig";
import { listDiagrams, loadDiagram, saveDiagram } from "./sharepointService";

const MERMAID_CDN = "https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js";

// ─── Sequence Diagram Parser & Renderer ──────────────────────────────────────
function parseSequence(src) {
  const lines = src.split("\n").map(l => l.trim()).filter(Boolean);
  const participants = [], participantSet = new Set(), steps = [];
  let title = "";

  const ensure = (name, isActor = false) => {
    if (!participantSet.has(name)) { participantSet.add(name); participants.push({ name, isActor }); }
  };

  for (const line of lines) {
    if (/^title:/i.test(line)) { title = line.replace(/^title:\s*/i, ""); continue; }
    if (/^participant\s+/i.test(line)) { ensure(line.replace(/^participant\s+/i, "")); continue; }
    if (/^actor\s+/i.test(line)) { ensure(line.replace(/^actor\s+/i, ""), true); continue; }
    if (/^activate\s+/i.test(line)) { steps.push({ type: "activate", name: line.replace(/^activate\s+/i, "") }); continue; }
    if (/^deactivate\s+/i.test(line)) { steps.push({ type: "deactivate", name: line.replace(/^deactivate\s+/i, "") }); continue; }
    if (/^note\s+(over|left of|right of)\s+/i.test(line)) {
      const m = line.match(/^note\s+(over|left of|right of)\s+([^:]+):\s*(.+)/i);
      if (m) steps.push({ type: "note", pos: m[1].toLowerCase(), target: m[2].trim(), text: m[3].trim() });
      continue;
    }
    const m = line.match(/^(.+?)(-->>|-->|->>|->)([+-]?)(.+?):\s*(.*)$/);
    if (m) {
      const [, from, arrow, activation, to, msg] = m;
      ensure(from.trim()); ensure(to.trim());
      steps.push({ type: "message", from: from.trim(), to: to.trim(), arrow, msg, activation });
    }
  }
  return { title, participants, steps };
}

function SequenceDiagram({ src, theme }) {
  const PAD = 40, BOX_W = 110, BOX_H = 36, GAP = 80, LIFE_TOP = PAD + BOX_H, ACT_W = 12;
  const { title, participants, steps } = parseSequence(src);
  if (!participants.length) return <text fill="#aaa" x="20" y="40" fontSize="13">Type a diagram…</text>;

  const cols = {};
  participants.forEach((p, i) => { cols[p.name] = PAD + i * (BOX_W + GAP) + BOX_W / 2; });
  const totalW = PAD * 2 + participants.length * (BOX_W + GAP) - GAP;

  const activations = {}, actStart = {}, actRects = [];
  participants.forEach(p => { activations[p.name] = 0; actStart[p.name] = []; });

  let y = LIFE_TOP + 20;
  const rows = [];

  for (const step of steps) {
    if (step.type === "activate") {
      const d = activations[step.name] || 0;
      actStart[step.name].push({ y, depth: d });
      activations[step.name] = d + 1;
    } else if (step.type === "deactivate") {
      const stack = actStart[step.name];
      if (stack?.length) {
        const { y: sy, depth: d } = stack.pop();
        actRects.push({ name: step.name, y1: sy, y2: y, depth: d });
        activations[step.name] = Math.max(0, (activations[step.name] || 1) - 1);
      }
    } else if (step.type === "note") {
      rows.push({ ...step, y }); y += 44;
    } else if (step.type === "message") {
      if (step.activation === "+") {
        const d = activations[step.to] || 0;
        actStart[step.to].push({ y: y + 10, depth: d });
        activations[step.to] = d + 1;
      }
      rows.push({ ...step, y }); y += 44;
      if (step.activation === "-") {
        const stack = actStart[step.to];
        if (stack?.length) {
          const { y: sy, depth: d } = stack.pop();
          actRects.push({ name: step.to, y1: sy, y2: y - 10, depth: d });
          activations[step.to] = Math.max(0, (activations[step.to] || 1) - 1);
        }
      }
    }
  }

  const totalH = y + BOX_H + PAD;
  const isHand = theme === "hand";
  const ff = isHand ? "'Segoe Print','Comic Sans MS',cursive" : "'Segoe UI',system-ui,sans-serif";
  const boxFill = isHand ? "#fffde7" : "#e8eaf6";
  const boxStroke = isHand ? "#5d4037" : "#7986cb";
  const lineColor = isHand ? "#5d4037" : "#555";
  const arrowColor = isHand ? "#5d4037" : "#333";
  const activeFill = isHand ? "#fff9c4" : "#c5cae9";

  const arrowHead = (x, y, dir, open) => {
    const s = 9;
    const pts = dir === "right"
      ? [[x, y], [x - s, y - s / 2], [x - s, y + s / 2]]
      : [[x, y], [x + s, y - s / 2], [x + s, y + s / 2]];
    const d = `M${pts[0]}L${pts[1]}L${pts[2]}${open ? "" : "Z"}`;
    return <path d={d} fill={open ? "white" : arrowColor} stroke={arrowColor} strokeWidth="1.5" />;
  };

  const renderMsg = (row) => {
    const fx = cols[row.from], tx = cols[row.to];
    const isSelf = row.from === row.to;
    const isAsync = row.arrow === "->>" || row.arrow === "-->>";
    const isDashed = row.arrow === "-->" || row.arrow === "-->>";
    const goRight = tx > fx;
    const dash = isDashed ? "6,4" : "none";
    let linePath, arrowX, arrowDir;
    if (isSelf) {
      const ox = fx + ACT_W;
      linePath = `M${ox},${row.y} Q${ox + 50},${row.y} ${ox + 50},${row.y + 18} Q${ox + 50},${row.y + 36} ${ox},${row.y + 36}`;
      arrowX = ox; arrowDir = "left";
    } else {
      const sx = goRight ? fx + ACT_W / 2 : fx - ACT_W / 2;
      const ex = goRight ? tx - ACT_W / 2 : tx + ACT_W / 2;
      linePath = `M${sx},${row.y} L${ex},${row.y}`;
      arrowX = ex; arrowDir = goRight ? "right" : "left";
    }
    const midX = isSelf ? fx + 28 : (fx + tx) / 2;
    const midY = isSelf ? row.y + 18 : row.y;
    return (
      <g key={`msg-${row.y}`}>
        <path d={linePath} fill="none" stroke={lineColor} strokeWidth="1.5" strokeDasharray={dash} />
        {arrowHead(arrowX, isSelf ? row.y + 36 : row.y, arrowDir, isAsync)}
        <text x={midX} y={midY - 7} textAnchor="middle" fontSize="12" fontFamily={ff} fill="#333">{row.msg}</text>
      </g>
    );
  };

  return (
    <svg width={totalW} height={totalH} xmlns="http://www.w3.org/2000/svg" fontFamily={ff}>
      {title && <text x={totalW / 2} y={20} textAnchor="middle" fontSize="15" fontWeight="bold" fill="#333">{title}</text>}
      {participants.map(p => {
        const cx = cols[p.name];
        return (
          <g key={`top-${p.name}`}>
            {p.isActor ? (
              <g>
                <circle cx={cx} cy={PAD + 10} r={10} fill={boxFill} stroke={boxStroke} strokeWidth="1.5" />
                <line x1={cx} y1={PAD + 20} x2={cx} y2={PAD + 34} stroke={boxStroke} strokeWidth="1.5" />
                <line x1={cx - 10} y1={PAD + 26} x2={cx + 10} y2={PAD + 26} stroke={boxStroke} strokeWidth="1.5" />
                <line x1={cx} y1={PAD + 34} x2={cx - 8} y2={PAD + 44} stroke={boxStroke} strokeWidth="1.5" />
                <line x1={cx} y1={PAD + 34} x2={cx + 8} y2={PAD + 44} stroke={boxStroke} strokeWidth="1.5" />
                <text x={cx} y={PAD + 60} textAnchor="middle" fontSize="13" fontWeight="600" fill="#333">{p.name}</text>
              </g>
            ) : (
              <g>
                <rect x={cx - BOX_W / 2} y={PAD} width={BOX_W} height={BOX_H} rx={4} fill={boxFill} stroke={boxStroke} strokeWidth="1.5" />
                <text x={cx} y={PAD + BOX_H / 2 + 5} textAnchor="middle" fontSize="13" fontWeight="600" fill="#333">{p.name}</text>
              </g>
            )}
          </g>
        );
      })}
      {participants.map(p => (
        <line key={`life-${p.name}`} x1={cols[p.name]} y1={LIFE_TOP + (p.isActor ? 30 : BOX_H)} x2={cols[p.name]} y2={totalH - BOX_H - PAD} stroke={lineColor} strokeWidth="1" strokeDasharray="4,4" opacity="0.5" />
      ))}
      {actRects.map((r, i) => (
        <rect key={`act-${i}`} x={cols[r.name] - ACT_W / 2 + r.depth * 4} y={r.y1} width={ACT_W} height={r.y2 - r.y1} fill={activeFill} stroke={boxStroke} strokeWidth="1.2" />
      ))}
      {rows.map(row => {
        if (row.type === "note") {
          const tx = cols[row.target];
          return (
            <g key={`note-${row.y}`}>
              <rect x={tx - 55} y={row.y - 12} width={110} height={28} rx={3} fill="#fff9c2" stroke="#f0c000" strokeWidth="1.2" />
              <text x={tx} y={row.y + 6} textAnchor="middle" fontSize="11" fontFamily={ff} fill="#555">{row.text}</text>
            </g>
          );
        }
        return renderMsg(row);
      })}
      {participants.map(p => {
        const cx = cols[p.name], by = totalH - BOX_H - PAD;
        return p.isActor ? null : (
          <g key={`bot-${p.name}`}>
            <rect x={cx - BOX_W / 2} y={by} width={BOX_W} height={BOX_H} rx={4} fill={boxFill} stroke={boxStroke} strokeWidth="1.5" />
            <text x={cx} y={by + BOX_H / 2 + 5} textAnchor="middle" fontSize="13" fontWeight="600" fill="#333">{p.name}</text>
          </g>
        );
      })}
    </svg>
  );
}

// ─── Mermaid Loader ───────────────────────────────────────────────────────────
let mermaidReady = false, mermaidLoading = false, mermaidCbs = [];
function loadMermaid(cb) {
  if (mermaidReady) return cb();
  mermaidCbs.push(cb);
  if (mermaidLoading) return;
  mermaidLoading = true;
  const s = document.createElement("script");
  s.src = MERMAID_CDN;
  s.onload = () => {
    window.mermaid.initialize({ startOnLoad: false, theme: "default", securityLevel: "loose" });
    mermaidReady = true;
    mermaidCbs.forEach(f => f()); mermaidCbs = [];
  };
  document.head.appendChild(s);
}
let mermaidCounter = 0;

// ─── Examples ─────────────────────────────────────────────────────────────────
const JS_EXAMPLES = {
  "Request Flow": `Title: Request Flow\nactor User\nparticipant Controller\nparticipant Service\nparticipant Repo\n\nUser->+Controller: POST /api/resource\nController->+Service: Create(request)\nService->+Repo: Insert(entity)\nRepo-->-Service: entity\nService-->-Controller: result\nController-->-User: 201 Created`,
  "Async Messaging": `Title: Service Bus Flow\nparticipant Client\nparticipant Queue\nparticipant Consumer\nparticipant DB\n\nClient->>Queue: Publish message\nnote over Queue: Persisted async\nQueue-->>Consumer: Deliver\nConsumer->DB: Insert batch\nDB-->Consumer: OK\nConsumer-->>Queue: Acknowledge`,
  "Error Flow": `Title: Dead Letter Queue\nparticipant Processor\nparticipant ServiceBus\nparticipant DB\n\nProcessor->ServiceBus: Fetch message\nactivate Processor\nServiceBus-->Processor: Message\nProcessor->DB: Insert\nDB-->Processor: Deadlock error\nProcessor->ServiceBus: Abandon\ndeactivate Processor`,
};

const MERMAID_EXAMPLES = {
  "Sequence Diagram": `sequenceDiagram\n    actor User\n    participant Controller\n    participant Service\n    participant Repo\n\n    User->>+Controller: POST /api/resource\n    Controller->>+Service: Create(request)\n    Service->>+Repo: Insert(entity)\n    Repo-->>-Service: entity\n    Service-->>-Controller: result\n    Controller-->>-User: 201 Created`,
  "Flowchart": `flowchart TD\n    A[Start] --> B{Redis Lock?}\n    B -->|Yes| C[Acquire Lock]\n    B -->|No| D[Wait 30s]\n    D --> B\n    C --> E[Fetch 250 rows]\n    E --> F[Send to Service Bus]\n    F --> G[BMS Processes]\n    G --> H{More rows?}\n    H -->|Yes| E\n    H -->|No| I[Release Lock]`,
  "ER Diagram": `erDiagram\n    DIAGRAM {\n      string id PK\n      string title\n      string sharepoint_path\n      datetime created_at\n    }\n    USER ||--o{ DIAGRAM : creates`,
};

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function App() {
  const { instance } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  const [renderer, setRenderer] = useState("jssd");
  const [code, setCode] = useState(JS_EXAMPLES["Request Flow"]);
  const [title, setTitle] = useState("Untitled Diagram");
  const [editingTitle, setEditingTitle] = useState(false);
  const [mSvg, setMSvg] = useState("");
  const [error, setError] = useState("");
  const [saveStatus, setSaveStatus] = useState("unsaved");
  const [savedFileId, setSavedFileId] = useState(null);
  const [autoSave, setAutoSave] = useState(true);
  const [darkMode, setDarkMode] = useState(false);
  const [zoom, setZoom] = useState(1);
  const [pan, setPan] = useState({ x: 0, y: 0 });
  const [dragging, setDragging] = useState(false);
  const dragStart = useRef(null);
  const previewRef = useRef(null);
  const [splitPct, setSplitPct] = useState(30);
  const splitDragging = useRef(false);
  const splitContainerRef = useRef(null);
  const [seqTheme, setSeqTheme] = useState("simple");
  const [mReady, setMReady] = useState(mermaidReady);
  const [diagrams, setDiagrams] = useState([]);
  const [showBrowser, setShowBrowser] = useState(false);
  const [browserLoading, setBrowserLoading] = useState(false);
  const autoSaveTimer = useRef(null);

  useEffect(() => { loadMermaid(() => setMReady(true)); }, []);

  // Mermaid render
  useEffect(() => {
    if (renderer !== "mermaid" || !mReady) return;
    (async () => {
      try {
        const id = `mermaid-${++mermaidCounter}`;
        const { svg } = await window.mermaid.render(id, code);
        setMSvg(svg); setError("");
      } catch (e) { setError(e.message || "Invalid syntax"); }
    })();
  }, [renderer, code, mReady]);

  // Auto-save
  useEffect(() => {
    if (!autoSave || !isAuthenticated) return;
    setSaveStatus("unsaved");
    clearTimeout(autoSaveTimer.current);
    autoSaveTimer.current = setTimeout(handleSave, 3000);
    return () => clearTimeout(autoSaveTimer.current);
  }, [code, title, autoSave, isAuthenticated]);

  async function handleSave() {
    if (!isAuthenticated) return;
    try {
      setSaveStatus("saving");
      const result = await saveDiagram(instance, title, code);
      setSavedFileId(result.id);
      setSaveStatus("saved");
    } catch (e) {
      console.error(e);
      setSaveStatus("error");
    }
  }

  async function handleOpenBrowser() {
    setShowBrowser(true);
    setBrowserLoading(true);
    try {
      const list = await listDiagrams(instance);
      setDiagrams(list);
    } catch (e) {
      console.error(e);
    } finally {
      setBrowserLoading(false);
    }
  }

  async function handleLoadDiagram(fileId, name) {
    try {
      const content = await loadDiagram(instance, fileId);
      setCode(content);
      setTitle(name);
      setSavedFileId(fileId);
      setSaveStatus("saved");
      setShowBrowser(false);
      // Auto-detect renderer
      setRenderer(content.trim().startsWith("sequenceDiagram") ? "mermaid" : "jssd");
    } catch (e) {
      console.error(e);
    }
  }

  function switchRenderer(r) {
    setRenderer(r); setError("");
    setCode(r === "jssd" ? JS_EXAMPLES["Request Flow"] : MERMAID_EXAMPLES["Flowchart"]);
    setSaveStatus("unsaved");
    setPan({ x: 0, y: 0 }); setZoom(1);
  }

  function handleSplitMouseDown() { splitDragging.current = true; }
  function handleSplitMouseMove(e) {
    if (!splitDragging.current || !splitContainerRef.current) return;
    const rect = splitContainerRef.current.getBoundingClientRect();
    const pct = ((e.clientX - rect.left) / rect.width) * 100;
    setSplitPct(Math.min(60, Math.max(15, pct)));
  }
  function handleSplitMouseUp() { splitDragging.current = false; }

  function handleMouseDown(e) {
    setDragging(true);
    dragStart.current = { x: e.clientX - pan.x, y: e.clientY - pan.y };
  }
  function handleMouseMove(e) {
    if (!dragging || !dragStart.current) return;
    setPan({ x: e.clientX - dragStart.current.x, y: e.clientY - dragStart.current.y });
  }
  function handleMouseUp() { setDragging(false); }

  function handleWheel(e) {
    e.preventDefault();
    const delta = e.deltaY > 0 ? -0.1 : 0.1;
    setZoom(z => Math.min(2, Math.max(0.2, z + delta)));
  }

  function fitToWidth() {
    if (!previewRef.current) return;
    const svgEl = previewRef.current.querySelector("svg");
    if (!svgEl) return;
    const containerW = previewRef.current.clientWidth - 48;
    const containerH = previewRef.current.clientHeight - 48;
    const svgW = svgEl.getBoundingClientRect().width / zoom;
    const svgH = svgEl.getBoundingClientRect().height / zoom;
    if (svgW && svgH) {
      const scaleW = containerW / svgW;
      const scaleH = containerH / svgH;
      const newZoom = Math.min(scaleW, scaleH, 1); // never scale up beyond 100%
      setZoom(newZoom);
      setPan({ x: 0, y: 0 });
    }
  }

  // Auto fit when diagram changes
  useEffect(() => {
    const t = setTimeout(fitToWidth, 100);
    return () => clearTimeout(t);
  }, [mSvg, code, mReady]);

  function handleLogin() { instance.loginPopup(loginRequest); }
  function handleLogout() { instance.logoutPopup(); }

  function handleExport() {
    const svgEl = document.querySelector("#preview-area svg");
    if (!svgEl) return;
    const blob = new Blob([svgEl.outerHTML], { type: "image/svg+xml" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `${title.replace(/\s+/g, "_")}.svg`;
    a.click();
  }

  const examples = renderer === "jssd" ? JS_EXAMPLES : MERMAID_EXAMPLES;
  const bg = darkMode ? "#1a1a2e" : "#f0f2f5";
  const panel = darkMode ? "#16213e" : "#ffffff";
  const border = darkMode ? "#2a2a4a" : "#e2e8f0";
  const text = darkMode ? "#e2e8f0" : "#1a202c";
  const subtext = darkMode ? "#94a3b8" : "#64748b";
  const accent = "#6366f1";
  const accentLight = darkMode ? "#1e1b4b" : "#eef2ff";
  const editorBg = darkMode ? "#0f172a" : "#fafbfc";
  const headerBg = darkMode ? "#0f172a" : "#f8fafc";
  const saveColors = { unsaved: "#f59e0b", saving: "#6366f1", saved: "#10b981", error: "#ef4444" };
  const saveLabels = { unsaved: "● Unsaved", saving: "↺ Saving…", saved: "✓ Saved", error: "✕ Failed" };
  const btn = { padding: "4px 11px", borderRadius: 6, cursor: "pointer", fontSize: 12, border: `1px solid ${border}`, background: "transparent", color: subtext };

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100vh", background: bg, color: text, fontFamily: "'Inter',system-ui,sans-serif", fontSize: 14 }}>

      {/* Header */}
      <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 16px", background: panel, borderBottom: `1px solid ${border}`, boxShadow: "0 1px 3px rgba(0,0,0,0.08)", flexWrap: "wrap" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <div style={{ width: 28, height: 28, borderRadius: 8, background: accent, display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontWeight: 700, fontSize: 15 }}>M</div>
          <span style={{ fontWeight: 700, fontSize: 16, color: accent, letterSpacing: -0.5 }}>MurMade</span>
        </div>
        <div style={{ width: 1, height: 20, background: border }} />

        {/* Renderer toggle */}
        <div style={{ display: "flex", borderRadius: 8, border: `1px solid ${border}`, overflow: "hidden", fontSize: 12 }}>
          {[["jssd", "⇒ Sequence"], ["mermaid", "⬡ Mermaid"]].map(([r, lbl]) => (
            <button key={r} onClick={() => switchRenderer(r)} style={{ padding: "4px 12px", border: "none", cursor: "pointer", fontWeight: renderer === r ? 700 : 400, background: renderer === r ? accent : "transparent", color: renderer === r ? "#fff" : subtext, transition: "all 0.15s" }}>{lbl}</button>
          ))}
        </div>

        {renderer === "jssd" && (
          <div style={{ display: "flex", borderRadius: 8, border: `1px solid ${border}`, overflow: "hidden", fontSize: 12 }}>
            {[["simple", "Clean"], ["hand", "Sketchy"]].map(([t, lbl]) => (
              <button key={t} onClick={() => setSeqTheme(t)} style={{ padding: "4px 10px", border: "none", cursor: "pointer", fontWeight: seqTheme === t ? 700 : 400, background: seqTheme === t ? "#0ea5e9" : "transparent", color: seqTheme === t ? "#fff" : subtext, transition: "all 0.15s" }}>{lbl}</button>
            ))}
          </div>
        )}

        <div style={{ width: 1, height: 20, background: border }} />

        {/* Title */}
        {editingTitle
          ? <input value={title} onChange={e => setTitle(e.target.value)} onBlur={() => setEditingTitle(false)} onKeyDown={e => e.key === "Enter" && setEditingTitle(false)} autoFocus
              style={{ border: `1px solid ${accent}`, borderRadius: 6, padding: "3px 8px", fontSize: 14, fontWeight: 600, color: text, background: panel, outline: "none", minWidth: 160 }} />
          : <span onClick={() => setEditingTitle(true)} style={{ fontWeight: 600, fontSize: 14, cursor: "pointer", padding: "3px 8px", borderRadius: 6, border: "1px solid transparent" }}
              onMouseEnter={e => e.target.style.borderColor = border} onMouseLeave={e => e.target.style.borderColor = "transparent"}>{title} ✎</span>
        }
        {isAuthenticated && <span style={{ fontSize: 12, color: saveColors[saveStatus], fontWeight: 500 }}>{saveLabels[saveStatus]}</span>}

        <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 8 }}>
          {isAuthenticated ? (
            <>
              <label style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer", fontSize: 12, color: subtext }}>
                <div onClick={() => setAutoSave(a => !a)} style={{ width: 32, height: 18, borderRadius: 9, background: autoSave ? accent : border, position: "relative", cursor: "pointer", transition: "background 0.2s" }}>
                  <div style={{ position: "absolute", top: 2, left: autoSave ? 16 : 2, width: 14, height: 14, borderRadius: 7, background: "#fff", transition: "left 0.2s" }} />
                </div>
                Auto-save
              </label>
              <button onClick={handleSave} style={{ ...btn, border: `1px solid ${accent}`, color: accent, fontWeight: 600 }}>Save</button>
              <button onClick={handleOpenBrowser} style={btn}>📂 Open</button>
              <button onClick={handleExport} style={btn}>↓ SVG</button>
              <button onClick={handleLogout} style={btn}>Sign out</button>
            </>
          ) : (
            <button onClick={handleLogin} style={{ ...btn, border: `1px solid ${accent}`, color: accent, fontWeight: 600 }}>🔐 Sign in with Microsoft</button>
          )}
          <button onClick={() => setDarkMode(d => !d)} style={{ ...btn, padding: "4px 9px" }}>{darkMode ? "☀️" : "🌙"}</button>
        </div>
      </div>

      {/* Examples bar */}
      <div style={{ display: "flex", gap: 6, padding: "7px 16px", background: panel, borderBottom: `1px solid ${border}`, overflowX: "auto", alignItems: "center" }}>
        <span style={{ fontSize: 11, color: subtext, marginRight: 4, whiteSpace: "nowrap" }}>Examples:</span>
        {Object.entries(examples).map(([k, v]) => (
          <button key={k} onClick={() => { setCode(v); setTitle(k); setSaveStatus("unsaved"); }}
            style={{ padding: "3px 10px", borderRadius: 20, border: `1px solid ${border}`, background: "transparent", color: subtext, cursor: "pointer", fontSize: 11, whiteSpace: "nowrap" }}>{k}</button>
        ))}
        {renderer === "jssd" && (
          <div style={{ marginLeft: "auto", fontSize: 11, color: subtext, whiteSpace: "nowrap" }}>
            {[["->", "sync"], ["->>" , "async"], ["-->", "return"], ["-->>", "async ret"]].map(([a, l]) => (
              <span key={a} style={{ marginRight: 8 }}><code style={{ background: accentLight, color: accent, padding: "1px 5px", borderRadius: 3 }}>{a}</code> {l}</span>
            ))}
          </div>
        )}
      </div>

      {/* Diagram Browser Modal */}
      {showBrowser && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div style={{ background: panel, borderRadius: 12, padding: 24, width: 480, maxHeight: "70vh", display: "flex", flexDirection: "column", boxShadow: "0 20px 60px rgba(0,0,0,0.3)" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
              <h2 style={{ margin: 0, fontSize: 16, fontWeight: 700 }}>Open Diagram</h2>
              <button onClick={() => setShowBrowser(false)} style={{ border: "none", background: "transparent", fontSize: 20, cursor: "pointer", color: subtext }}>✕</button>
            </div>
            {browserLoading
              ? <p style={{ color: subtext, textAlign: "center" }}>Loading diagrams…</p>
              : diagrams.length === 0
                ? <p style={{ color: subtext, textAlign: "center" }}>No saved diagrams found in SharePoint folder.</p>
                : <div style={{ overflowY: "auto" }}>
                    {diagrams.map(d => (
                      <div key={d.id} onClick={() => handleLoadDiagram(d.id, d.name)}
                        style={{ padding: "10px 14px", borderRadius: 8, cursor: "pointer", marginBottom: 6, border: `1px solid ${border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}
                        onMouseEnter={e => e.currentTarget.style.background = accentLight}
                        onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                        <span style={{ fontWeight: 600, fontSize: 13 }}>📄 {d.name}</span>
                        <span style={{ fontSize: 11, color: subtext }}>{d.lastModified}</span>
                      </div>
                    ))}
                  </div>
            }
          </div>
        </div>
      )}

      {/* Split pane */}
      <div ref={splitContainerRef} style={{ display: "flex", flex: 1, overflow: "hidden" }} onMouseMove={handleSplitMouseMove} onMouseUp={handleSplitMouseUp} onMouseLeave={handleSplitMouseUp}>
        <div style={{ width: `${splitPct}%`, display: "flex", flexDirection: "column", borderRight: `1px solid ${border}` }}>
          <div style={{ padding: "6px 12px", fontSize: 11, fontWeight: 600, color: subtext, background: headerBg, borderBottom: `1px solid ${border}`, letterSpacing: 0.5, textTransform: "uppercase" }}>
            {renderer === "jssd" ? "Sequence Source" : "Mermaid Source"}
          </div>
          <textarea value={code} onChange={e => { setCode(e.target.value); setSaveStatus("unsaved"); }} spellCheck={false}
            style={{ flex: 1, padding: 16, fontFamily: "'JetBrains Mono','Fira Code','Consolas',monospace", fontSize: 13, lineHeight: 1.6, border: "none", outline: "none", resize: "none", background: editorBg, color: darkMode ? "#e2e8f0" : "#1a202c", tabSize: 2 }} />
          {error && <div style={{ padding: "8px 12px", background: "#fef2f2", borderTop: "1px solid #fecaca", color: "#dc2626", fontSize: 12, fontFamily: "monospace" }}>⚠ {error}</div>}
        </div>

        {/* Draggable divider */}
        <div onMouseDown={handleSplitMouseDown} style={{ width: 5, cursor: "col-resize", background: border, flexShrink: 0, transition: "background 0.15s" }}
          onMouseEnter={e => e.target.style.background = accent}
          onMouseLeave={e => e.target.style.background = border} />

        <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>
          <div style={{ padding: "6px 12px", fontSize: 11, fontWeight: 600, color: subtext, background: headerBg, borderBottom: `1px solid ${border}`, letterSpacing: 0.5, textTransform: "uppercase", display: "flex", alignItems: "center", gap: 8 }}>
            Preview
            <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 6 }}>
              {[["−", -0.1], ["+", 0.1]].map(([l, d]) => (
                <button key={l} onClick={() => setZoom(z => Math.min(2, Math.max(0.2, z + d)))}
                  style={{ width: 22, height: 22, borderRadius: 4, border: `1px solid ${border}`, background: "transparent", cursor: "pointer", color: subtext, fontSize: 14, display: "flex", alignItems: "center", justifyContent: "center" }}>{l}</button>
              ))}
              <span style={{ fontSize: 11, color: subtext, minWidth: 36, textAlign: "center" }}>{Math.round(zoom * 100)}%</span>
              <button onClick={() => { setZoom(1); setPan({ x: 0, y: 0 }); }} style={{ fontSize: 11, padding: "2px 6px", borderRadius: 4, border: `1px solid ${border}`, background: "transparent", cursor: "pointer", color: subtext }}>Reset</button>
              <button onClick={fitToWidth} style={{ fontSize: 11, padding: "2px 6px", borderRadius: 4, border: `1px solid ${border}`, background: "transparent", cursor: "pointer", color: subtext }}>Fit</button>
            </div>
          </div>
          <div ref={previewRef} onMouseDown={handleMouseDown} onMouseMove={handleMouseMove} onMouseUp={handleMouseUp} onMouseLeave={handleMouseUp} onWheel={handleWheel}
            style={{ flex: 1, overflow: "hidden", padding: 24, background: "#fff", display: "flex", alignItems: "flex-start", justifyContent: "center", cursor: dragging ? "grabbing" : "grab", userSelect: "none" }}>
            <div id="preview-area" style={{ transform: `translate(${pan.x}px, ${pan.y}px) scale(${zoom})`, transformOrigin: "top center", transition: dragging ? "none" : "transform 0.15s" }}>
              {renderer === "jssd"
                ? <SequenceDiagram src={code} theme={seqTheme} />
                : <div dangerouslySetInnerHTML={{ __html: mSvg || `<p style='color:#94a3b8;font-size:13px'>${mReady ? "Type a diagram…" : "Loading Mermaid…"}</p>` }} />
              }
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}