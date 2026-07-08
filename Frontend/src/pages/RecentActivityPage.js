import React, { useState, useEffect, useCallback, useMemo } from "react";
import { useNavigate } from "react-router-dom";
import { FiClock, FiRefreshCw, FiSearch, FiUsers, FiDollarSign, FiPhone } from "react-icons/fi";
import apiClient from "../api/client";

// ── helpers ───────────────────────────────────────────────────────────────────
function timeAgo(ts) {
  if (!ts) return "";
  const diff = Math.floor((Date.now() - new Date(ts)) / 1000);
  if (diff < 60) return `${diff}s ago`;
  if (diff < 3600) return `${Math.floor(diff / 60)}m ago`;
  if (diff < 86400) return `${Math.floor(diff / 3600)}h ago`;
  return `${Math.floor(diff / 86400)}d ago`;
}
function fullTime(ts) {
  if (!ts) return "";
  try { return new Date(ts).toLocaleString("en-IN", { timeZone: "Asia/Kolkata" }); }
  catch { return String(ts); }
}

const SOURCE = {
  calllog: { color: "bg-teal-500", light: "bg-teal-50", text: "text-teal-600", label: "Call Log", plural: "Call Logs", Icon: FiPhone },
  contact: { color: "bg-sky-500", light: "bg-sky-50", text: "text-sky-600", label: "Contact Note", plural: "Contact Notes", Icon: FiUsers },
  deal: { color: "bg-violet-500", light: "bg-violet-50", text: "text-violet-600", label: "Deal Note", plural: "Deal Notes", Icon: FiDollarSign },
};

const RANGES = [
  { key: "1", label: "24 hours", days: 1 },
  { key: "7", label: "7 days", days: 7 },
  { key: "30", label: "30 days", days: 30 },
  { key: "all", label: "All time", days: null },
];

export default function RecentActivityPage() {
  const navigate = useNavigate();

  const [items, setItems] = useState([]);
  const [loading, setLoading] = useState(false);
  const [range, setRange] = useState("7"); // default: last 7 days (the header shows only 24h)
  const [source, setSource] = useState("all"); // all | calllog | contact | deal
  const [search, setSearch] = useState("");

  const fetchRecents = useCallback(async () => {
    setLoading(true);
    const list = [];

    // Both endpoints hit concurrently (was sequential) so the page loads faster.
    const [notesRes, callRes] = await Promise.allSettled([
      apiClient.get("/Notes"),
      apiClient.get("/CallLog"),
    ]);

    // Notes (contact + deal notes)
    if (notesRes.status === "fulfilled") {
      const arr = Array.isArray(notesRes.value?.data) ? notesRes.value.data : [];
      arr.forEach((n) => {
        const ts = n.CreatedAt || n.createdAt;
        const desc = n.Description || n.description || "";
        if (!ts || !desc.trim()) return;
        const hasDeal = !!(n.DealId || n.dealId);
        const hasContact = !!(n.ContactId || n.contactId);
        list.push({
          id: `note-${n.Id || n.id}`,
          entityId: hasContact ? (n.ContactId || n.contactId) : (n.DealId || n.dealId),
          source: hasDeal ? "deal" : "contact",
          title: desc.length > 90 ? desc.slice(0, 90) + "…" : desc,
          preview: desc.length > 90 ? desc : null,
          time: ts,
          extra: hasDeal ? "Deal Note" : "Contact Note",
        });
      });
    }

    // Call logs
    if (callRes.status === "fulfilled") {
      const arr = Array.isArray(callRes.value?.data) ? callRes.value.data : [];
      arr.forEach((l) => {
        const ts = l.CreatedAt || l.createdAt;
        if (!ts) return;
        const dir = l.CallDirection || l.callDirection || "";
        const duration = l.CallDuration || l.callDuration || "";
        const assoc = l.AssociatedWithCall || l.associatedWithCall || "";
        const owner = l.CallOwner || l.callOwner || "";
        const notes = l.Notes || l.notes || "";
        list.push({
          id: `calllog-${l.CallLogId || l.callLogId}`,
          entityId: l.ContactId || l.contactId,
          source: "calllog",
          title: assoc || owner || "Call Log",
          preview: notes || null,
          time: ts,
          extra: [dir, duration ? String(duration).replace("00:", "").replace(/^0/, "") : ""].filter(Boolean).join(" · ") || "Call",
        });
      });
    }

    list.sort((a, b) => new Date(b.time) - new Date(a.time));
    setItems(list);
    setLoading(false);
  }, []);

  useEffect(() => { fetchRecents(); }, [fetchRecents]);

  const withinRange = useCallback((ts) => {
    const days = RANGES.find((r) => r.key === range)?.days;
    if (days == null) return true;
    return new Date(ts) >= new Date(Date.now() - days * 24 * 60 * 60 * 1000);
  }, [range]);

  // Range applies to the counts as well, so the tab badges reflect the selected window.
  const inRange = useMemo(() => items.filter((i) => withinRange(i.time)), [items, withinRange]);

  const counts = {
    all: inRange.length,
    calllog: inRange.filter((i) => i.source === "calllog").length,
    contact: inRange.filter((i) => i.source === "contact").length,
    deal: inRange.filter((i) => i.source === "deal").length,
  };

  const visible = useMemo(() => {
    const q = search.trim().toLowerCase();
    return inRange.filter((i) => {
      if (source !== "all" && i.source !== source) return false;
      if (!q) return true;
      return [i.title, i.preview, i.extra].some((v) => (v || "").toLowerCase().includes(q));
    });
  }, [inRange, source, search]);

  const onItemClick = (item) => {
    if (item.source === "contact" && item.entityId) {
      window.dispatchEvent(new CustomEvent("showContact", { detail: { contactId: item.entityId } }));
    } else if (item.source === "deal") {
      navigate("/dashboard/Deals");
    } else if (item.source === "calllog") {
      if (item.entityId) {
        window.dispatchEvent(new CustomEvent("showContact", { detail: { contactId: item.entityId } }));
      } else {
        navigate("/dashboard/contacts");
      }
    }
  };

  const TABS = [
    { key: "all", label: "All" },
    { key: "calllog", label: "Call Logs" },
    { key: "contact", label: "Contact Notes" },
    { key: "deal", label: "Deal Notes" },
  ];

  return (
    <div className="p-4 sm:p-6 h-full overflow-hidden bg-gray-50">
      <div className="w-full h-full flex flex-col min-h-0">
        {/* Back */}
        <button
          onClick={() => navigate(-1)}
          className="inline-flex items-center gap-1.5 text-sm text-gray-600 hover:text-gray-900 font-medium mb-3"
        >
          <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 19l-7-7 7-7" />
          </svg>
          Back
        </button>

        {/* Title row */}
        <div className="flex items-center justify-between flex-wrap gap-3 mb-4">
          <div className="flex items-center gap-3">
            <span className="w-10 h-10 rounded-xl bg-teal-500 flex items-center justify-center text-white flex-shrink-0">
              <FiClock size={20} />
            </span>
            <div>
              <h1 className="text-xl font-bold text-gray-900">Recent Activity</h1>
              <p className="text-sm text-gray-500">Notes and call logs across the CRM, newest first.</p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <select value={range} onChange={(e) => setRange(e.target.value)}
              className="px-3 py-2 border border-gray-300 rounded-lg text-sm bg-white focus:outline-none focus:ring-2 focus:ring-teal-500">
              {RANGES.map((r) => <option key={r.key} value={r.key}>Last {r.label}</option>)}
            </select>
            <button onClick={fetchRecents}
              className="inline-flex items-center gap-1.5 text-sm text-gray-700 font-medium px-3 py-2 rounded-lg border border-gray-300 bg-white hover:bg-gray-50 transition-colors">
              <FiRefreshCw size={15} className={loading ? "animate-spin" : ""} /> Refresh
            </button>
          </div>
        </div>

        {/* Filter tabs + search */}
        <div className="flex flex-wrap items-center gap-2 mb-4">
          {TABS.map((t) => (
            <button key={t.key} onClick={() => setSource(t.key)}
              className={`inline-flex items-center gap-2 px-3.5 py-1.5 rounded-lg text-sm font-medium border transition ${
                source === t.key ? "bg-teal-600 text-white border-teal-600 shadow-sm" : "bg-white text-gray-700 border-gray-300 hover:bg-gray-50"
              }`}>
              {t.label}
              <span className={`text-[11px] font-bold px-1.5 rounded-full ${source === t.key ? "bg-white/25 text-white" : "bg-gray-100 text-gray-500"}`}>{counts[t.key]}</span>
            </button>
          ))}
          <div className="relative flex-1 min-w-[180px]">
            <FiSearch size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
            <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search activity…"
              className="w-full pl-9 pr-3 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-teal-500" />
          </div>
        </div>

        {/* List */}
        <div className="flex-1 min-h-0 overflow-auto bg-white border border-gray-200 rounded-xl">
          {loading && items.length === 0 ? (
            <div className="p-4 space-y-3">
              {[...Array(6)].map((_, i) => (
                <div key={i} className="flex gap-3 items-start animate-pulse">
                  <div className="w-10 h-10 rounded-xl bg-gray-200 flex-shrink-0" />
                  <div className="flex-1 space-y-2 pt-1">
                    <div className="h-3 w-2/3 bg-gray-200 rounded" />
                    <div className="h-2.5 w-1/3 bg-gray-100 rounded" />
                  </div>
                </div>
              ))}
            </div>
          ) : visible.length === 0 ? (
            <div className="flex flex-col items-center justify-center py-20 px-4 text-center">
              <div className="w-16 h-16 rounded-full bg-gray-100 flex items-center justify-center mb-3">
                <FiClock size={28} className="text-gray-300" />
              </div>
              <p className="text-sm font-medium text-gray-500">No recent activity</p>
              <p className="text-xs text-gray-400 mt-1">Notes and call logs in the selected window will appear here.</p>
            </div>
          ) : (
            <ul className="divide-y divide-gray-100">
              {visible.map((item) => {
                const cfg = SOURCE[item.source] || SOURCE.contact;
                const Icon = cfg.Icon;
                return (
                  <li key={item.id}>
                    <button onClick={() => onItemClick(item)}
                      className="w-full text-left px-4 py-3.5 flex gap-3 hover:bg-gray-50 transition-colors group">
                      <div className={`w-10 h-10 rounded-xl ${cfg.light} flex items-center justify-center flex-shrink-0`}>
                        <Icon size={17} className={cfg.text} />
                      </div>
                      <div className="flex-1 min-w-0">
                        <div className="flex items-start justify-between gap-2">
                          <p className="text-sm font-semibold text-gray-900 truncate group-hover:text-teal-600 transition-colors">{item.title}</p>
                          <span className="text-[11px] text-gray-400 whitespace-nowrap flex-shrink-0" title={fullTime(item.time)}>{timeAgo(item.time)}</span>
                        </div>
                        <div className="flex items-center gap-2 mt-0.5">
                          <span className={`text-[10px] font-semibold px-1.5 py-0.5 rounded-full ${cfg.light} ${cfg.text}`}>{cfg.label}</span>
                          {item.extra && <span className={`text-[11px] font-medium ${cfg.text}`}>{item.extra}</span>}
                        </div>
                        {item.preview && <p className="text-xs text-gray-500 mt-1 line-clamp-2 leading-tight">{item.preview}</p>}
                      </div>
                      <svg className="w-4 h-4 text-gray-300 group-hover:text-teal-500 self-center flex-shrink-0 transition-colors" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5l7 7-7 7" />
                      </svg>
                    </button>
                  </li>
                );
              })}
            </ul>
          )}
        </div>

        {!loading && (
          <p className="text-xs text-gray-400 text-center mt-3">
            {visible.length} item{visible.length !== 1 ? "s" : ""} shown
          </p>
        )}
      </div>
    </div>
  );
}
