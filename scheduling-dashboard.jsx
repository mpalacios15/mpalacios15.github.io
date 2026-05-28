import { useState, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";

const SCHEDULED = "Scheduled";
const ASSIGNING = "Assigning Technician";

const AGENT_COLORS = [
  "#7F77DD", "#1D9E75", "#D85A30", "#378ADD", "#BA7517", "#D4537E"
];

function dateLabel(serial) {
  if (!serial) return "—";
  const d = new Date((serial - 25569) * 86400 * 1000);
  return d.toLocaleDateString("en-US", { weekday: "short", month: "short", day: "numeric", timeZone: "UTC" });
}

function UploadScreen({ onFile }) {
  const [dragging, setDragging] = useState(false);
  return (
    <div
      onDragOver={e => { e.preventDefault(); setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={e => { e.preventDefault(); setDragging(false); onFile(e.dataTransfer.files[0]); }}
      style={{ padding: "3rem 1rem", display: "flex", flexDirection: "column", alignItems: "center", gap: "1.5rem" }}
    >
      <div style={{ textAlign: "center" }}>
        <h2 style={{ fontSize: 22, fontWeight: 500, margin: "0 0 8px" }}>Scheduling ops dashboard</h2>
        <p style={{ fontSize: 14, color: "var(--color-text-secondary)", margin: 0 }}>
          Load your daily assignment file to monitor agent progress in real time
        </p>
      </div>
      <label htmlFor="file-upload" style={{
        border: `2px dashed ${dragging ? "var(--color-border-info)" : "var(--color-border-secondary)"}`,
        borderRadius: "var(--border-radius-lg)",
        padding: "2.5rem 4rem",
        display: "flex", flexDirection: "column", alignItems: "center", gap: 12,
        cursor: "pointer",
        background: dragging ? "var(--color-background-info)" : "var(--color-background-secondary)",
        transition: "all 0.15s", maxWidth: 420, width: "100%",
      }}>
        <i className="ti ti-file-spreadsheet" style={{ fontSize: 36, color: "var(--color-text-secondary)" }} aria-hidden="true" />
        <span style={{ fontSize: 15, fontWeight: 500 }}>Drop .xlsx or click to browse</span>
        <span style={{ fontSize: 12, color: "var(--color-text-tertiary)" }}>Daily scheduling export · all dates loaded automatically</span>
      </label>
      <input id="file-upload" type="file" accept=".xlsx,.xls" style={{ display: "none" }}
        onChange={e => onFile(e.target.files[0])} />
    </div>
  );
}

function ProgressBar({ pct }) {
  return (
    <div style={{ height: 3, background: "var(--color-background-secondary)", margin: "0 14px" }}>
      <div style={{
        height: "100%", width: `${pct}%`,
        background: pct === 100 ? "var(--color-text-success)" : "var(--color-border-info)",
        borderRadius: 2, transition: "width 0.4s ease",
      }} />
    </div>
  );
}

function SiteRow({ site }) {
  const isSched = site.Status === SCHEDULED;
  return (
    <div style={{
      padding: "9px 14px",
      borderBottom: "0.5px solid var(--color-border-tertiary)",
      display: "flex", gap: 10, alignItems: "flex-start",
    }}>
      <div style={{
        width: 6, height: 6, borderRadius: "50%", flexShrink: 0, marginTop: 5,
        background: isSched ? "var(--color-text-success)" : "var(--color-text-warning)",
      }} />
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 3 }}>
          <span style={{ fontSize: 11, color: "var(--color-text-tertiary)", fontFamily: "var(--font-mono)" }}>
            #{site["Incident #"]}
          </span>
          <span style={{
            fontSize: 10, padding: "1px 6px", borderRadius: 4,
            background: isSched ? "var(--color-background-success)" : "var(--color-background-warning)",
            color: isSched ? "var(--color-text-success)" : "var(--color-text-warning)",
          }}>
            {isSched ? "Scheduled" : "Assigning tech"}
          </span>
          <span style={{ fontSize: 10, color: "var(--color-text-tertiary)", marginLeft: "auto", flexShrink: 0 }}>
            {site["Ss or Exe"]}
          </span>
        </div>
        <div style={{ fontSize: 13, fontWeight: 500, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
          {site["SiteAddress"] || "No address"}
        </div>
        <div style={{ fontSize: 12, color: "var(--color-text-secondary)" }}>
          {[site["SiteCity"], site["State"], site["SiteZip"]].filter(Boolean).join(", ")}
          {site["Time zone"] ? <span style={{ marginLeft: 8, fontSize: 11, color: "var(--color-text-tertiary)" }}>{site["Time zone"]}</span> : null}
        </div>
        {site["VendorName"] && site["VendorName"].trim() !== "" && (
          <div style={{ fontSize: 11, color: "var(--color-text-tertiary)", marginTop: 3, display: "flex", gap: 6, flexWrap: "wrap" }}>
            <span><i className="ti ti-user" style={{ fontSize: 11, marginRight: 3 }} aria-hidden="true" />{site["VendorName"]}</span>
            {site["VendorTechCell"] && site["VendorTechCell"].trim() !== "" && (
              <span><i className="ti ti-phone" style={{ fontSize: 11, marginRight: 3 }} aria-hidden="true" />{site["VendorTechCell"]}</span>
            )}
            {site["RATES"] && site["RATES"].trim() !== "" && (
              <span><i className="ti ti-coin" style={{ fontSize: 11, marginRight: 3 }} aria-hidden="true" />{site["RATES"]}</span>
            )}
          </div>
        )}
        {(!site["VendorName"] || site["VendorName"].trim() === "") && site["RATES"] && (
          <div style={{ fontSize: 11, color: "var(--color-text-tertiary)", marginTop: 3 }}>
            <i className="ti ti-coin" style={{ fontSize: 11, marginRight: 3 }} aria-hidden="true" />{site["RATES"]}
          </div>
        )}
      </div>
    </div>
  );
}

function AgentCard({ agent, sites, color, scheduled, assigning, isOpen, onToggle, filterStatus }) {
  const total = sites.length;
  const pct = total > 0 ? Math.round((scheduled / total) * 100) : 0;
  const visible = filterStatus === "all" ? sites :
    sites.filter(s => s.Status === (filterStatus === "scheduled" ? SCHEDULED : ASSIGNING));

  return (
    <div style={{
      background: "var(--color-background-primary)",
      border: `0.5px solid ${isOpen ? "var(--color-border-secondary)" : "var(--color-border-tertiary)"}`,
      borderRadius: "var(--border-radius-lg)", overflow: "hidden",
    }}>
      <div onClick={onToggle} style={{
        padding: "12px 14px", cursor: "pointer",
        display: "flex", alignItems: "center", gap: 10, userSelect: "none",
      }}>
        <div style={{ width: 8, height: 8, borderRadius: "50%", background: color, flexShrink: 0 }} />
        <span style={{ flex: 1, fontSize: 14, fontWeight: 500 }}>{agent}</span>
        <span style={{
          fontSize: 11, padding: "2px 8px", borderRadius: 4,
          background: "var(--color-background-success)", color: "var(--color-text-success)",
        }}>
          <i className="ti ti-check" style={{ fontSize: 11, marginRight: 3 }} aria-hidden="true" />{scheduled}
        </span>
        {assigning > 0 && (
          <span style={{
            fontSize: 11, padding: "2px 8px", borderRadius: 4,
            background: "var(--color-background-warning)", color: "var(--color-text-warning)",
          }}>
            <i className="ti ti-clock" style={{ fontSize: 11, marginRight: 3 }} aria-hidden="true" />{assigning}
          </span>
        )}
        <i className={`ti ti-chevron-${isOpen ? "up" : "down"}`}
          style={{ fontSize: 14, color: "var(--color-text-tertiary)" }} aria-hidden="true" />
      </div>

      <ProgressBar pct={pct} />
      <div style={{ fontSize: 11, color: "var(--color-text-tertiary)", padding: "3px 14px 8px", textAlign: "right" }}>
        {pct}% · {total} site{total !== 1 ? "s" : ""}
      </div>

      {isOpen && (
        <div style={{ borderTop: "0.5px solid var(--color-border-tertiary)" }}>
          {visible.length === 0
            ? <div style={{ padding: "14px", fontSize: 13, color: "var(--color-text-tertiary)", textAlign: "center" }}>No sites match filter</div>
            : visible.map((site, i) => <SiteRow key={i} site={site} />)
          }
        </div>
      )}
    </div>
  );
}

export default function SchedulingDashboard() {
  const [records, setRecords] = useState(null);
  const [selectedDate, setSelectedDate] = useState(null);
  const [expanded, setExpanded] = useState(new Set());
  const [filterStatus, setFilterStatus] = useState("all");
  const [search, setSearch] = useState("");

  const handleFile = useCallback((file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      const [headers, ...data] = rows;
      const recs = data
        .filter(r => r.some(c => c !== ""))
        .map(row => Object.fromEntries(headers.map((h, i) => [String(h ?? "").trim(), row[i]])));
      setRecords(recs);
      const dates = [...new Set(recs.map(r => r.Date).filter(Boolean))].sort((a, b) => a - b);
      setSelectedDate(dates[0] ?? null);
      setExpanded(new Set());
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const { dates, agentGroups, summary } = useMemo(() => {
    if (!records) return { dates: [], agentGroups: [], summary: {} };
    const allDates = [...new Set(records.map(r => r.Date).filter(Boolean))].sort((a, b) => a - b);
    const dayRecs = records.filter(r => r.Date === selectedDate);

    const sq = search.trim().toLowerCase();
    const searched = sq
      ? dayRecs.filter(r =>
          (r["ASSIGNED BY"] || "").toLowerCase().includes(sq) ||
          (r["SiteAddress"] || "").toLowerCase().includes(sq) ||
          (r["SiteCity"] || "").toLowerCase().includes(sq) ||
          String(r["Incident #"]).includes(sq) ||
          (r["VendorName"] || "").toLowerCase().includes(sq)
        )
      : dayRecs;

    const agentMap = new Map();
    searched.forEach(r => {
      const a = r["ASSIGNED BY"] || "Unassigned";
      if (!agentMap.has(a)) agentMap.set(a, []);
      agentMap.get(a).push(r);
    });

    const groups = [...agentMap.entries()].map(([agent, sites], i) => ({
      agent, sites, color: AGENT_COLORS[i % AGENT_COLORS.length],
      scheduled: sites.filter(s => s.Status === SCHEDULED).length,
      assigning: sites.filter(s => s.Status === ASSIGNING).length,
    })).sort((a, b) => b.sites.length - a.sites.length);

    const totalSched = searched.filter(r => r.Status === SCHEDULED).length;
    const totalAssign = searched.filter(r => r.Status === ASSIGNING).length;
    return {
      dates: allDates,
      agentGroups: groups,
      summary: { total: searched.length, scheduled: totalSched, assigning: totalAssign, agents: groups.length },
    };
  }, [records, selectedDate, search]);

  const toggleAgent = (agent) => setExpanded(prev => {
    const next = new Set(prev);
    next.has(agent) ? next.delete(agent) : next.add(agent);
    return next;
  });

  const allExpanded = expanded.size === agentGroups.length && agentGroups.length > 0;

  if (!records) return <UploadScreen onFile={handleFile} />;

  return (
    <div style={{ fontFamily: "var(--font-sans)", padding: "0 0 2rem" }}>
      <h2 className="sr-only">Scheduling ops dashboard — agent site assignment tracker</h2>

      {/* Top bar */}
      <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 14, flexWrap: "wrap" }}>
        <span style={{ fontSize: 12, color: "var(--color-text-tertiary)" }}>Date</span>
        {dates.map(d => (
          <button key={d} onClick={() => setSelectedDate(d)} style={{
            fontSize: 13, padding: "5px 12px",
            borderRadius: "var(--border-radius-md)",
            border: selectedDate === d ? "2px solid var(--color-border-info)" : "0.5px solid var(--color-border-tertiary)",
            background: selectedDate === d ? "var(--color-background-info)" : "transparent",
            color: selectedDate === d ? "var(--color-text-info)" : "var(--color-text-secondary)",
            cursor: "pointer", fontWeight: selectedDate === d ? 500 : 400,
          }}>
            {dateLabel(d)}
          </button>
        ))}
        <button onClick={() => { setRecords(null); setSelectedDate(null); setSearch(""); }} style={{
          marginLeft: "auto", fontSize: 12, padding: "5px 10px",
          borderRadius: "var(--border-radius-md)",
          border: "0.5px solid var(--color-border-tertiary)",
          background: "transparent", color: "var(--color-text-tertiary)", cursor: "pointer",
        }}>
          <i className="ti ti-upload" style={{ fontSize: 12, marginRight: 4 }} aria-hidden="true" />
          New file
        </button>
      </div>

      {/* Summary stats */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 8, marginBottom: 16 }}>
        {[
          { label: "Total sites", value: summary.total, color: "var(--color-text-primary)" },
          { label: "Scheduled", value: summary.scheduled, color: "var(--color-text-success)" },
          { label: "Assigning tech", value: summary.assigning, color: "var(--color-text-warning)" },
          { label: "Agents", value: summary.agents, color: "var(--color-text-secondary)" },
        ].map(({ label, value, color }) => (
          <div key={label} style={{
            background: "var(--color-background-secondary)",
            borderRadius: "var(--border-radius-md)", padding: "12px 14px",
          }}>
            <div style={{ fontSize: 12, color: "var(--color-text-secondary)", marginBottom: 4 }}>{label}</div>
            <div style={{ fontSize: 24, fontWeight: 500, color }}>{value}</div>
          </div>
        ))}
      </div>

      {/* Filters row */}
      <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 12, flexWrap: "wrap" }}>
        <div style={{ position: "relative", flex: "1 1 180px", maxWidth: 260 }}>
          <i className="ti ti-search" style={{
            position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)",
            fontSize: 14, color: "var(--color-text-tertiary)", pointerEvents: "none",
          }} aria-hidden="true" />
          <input
            type="text"
            placeholder="Search agent, address, incident…"
            value={search}
            onChange={e => setSearch(e.target.value)}
            style={{ width: "100%", paddingLeft: 32, fontSize: 13, height: 34, boxSizing: "border-box" }}
          />
        </div>

        {[
          { key: "all", label: "All" },
          { key: "scheduled", label: "Scheduled" },
          { key: "assigning", label: "Assigning" },
        ].map(({ key, label }) => (
          <button key={key} onClick={() => setFilterStatus(key)} style={{
            fontSize: 12, padding: "5px 12px",
            borderRadius: "var(--border-radius-md)",
            border: filterStatus === key ? "0.5px solid var(--color-border-secondary)" : "0.5px solid var(--color-border-tertiary)",
            background: filterStatus === key ? "var(--color-background-secondary)" : "transparent",
            color: filterStatus === key ? "var(--color-text-primary)" : "var(--color-text-tertiary)",
            cursor: "pointer",
          }}>
            {label}
          </button>
        ))}

        <button onClick={() => setExpanded(allExpanded ? new Set() : new Set(agentGroups.map(g => g.agent)))} style={{
          marginLeft: "auto", fontSize: 12, padding: "5px 10px",
          borderRadius: "var(--border-radius-md)",
          border: "0.5px solid var(--color-border-tertiary)",
          background: "transparent", color: "var(--color-text-tertiary)", cursor: "pointer",
        }}>
          <i className={`ti ti-${allExpanded ? "fold" : "unfold"}`} style={{ fontSize: 12, marginRight: 4 }} aria-hidden="true" />
          {allExpanded ? "Collapse all" : "Expand all"}
        </button>
      </div>

      {/* Agent grid */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: 10 }}>
        {agentGroups.map(({ agent, sites, color, scheduled, assigning }) => (
          <AgentCard
            key={agent}
            agent={agent}
            sites={sites}
            color={color}
            scheduled={scheduled}
            assigning={assigning}
            isOpen={expanded.has(agent)}
            onToggle={() => toggleAgent(agent)}
            filterStatus={filterStatus}
          />
        ))}
        {agentGroups.length === 0 && (
          <div style={{ gridColumn: "1 / -1", padding: "2rem", textAlign: "center", color: "var(--color-text-tertiary)", fontSize: 14 }}>
            No results for this date{search ? " and search" : ""}.
          </div>
        )}
      </div>
    </div>
  );
}
