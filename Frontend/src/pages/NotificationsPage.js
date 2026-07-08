import React, { useState, useEffect, useCallback, useMemo, useRef } from "react";
import { useNavigate } from "react-router-dom";
import { useMsal } from "@azure/msal-react";
import { FiBell, FiMail, FiMessageSquare, FiRefreshCw, FiCheckCircle, FiExternalLink, FiSearch } from "react-icons/fi";

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
  email: { color: "bg-blue-500", light: "bg-blue-50", text: "text-blue-600", ring: "ring-blue-100", label: "Email", Icon: FiMail },
  teams: { color: "bg-purple-500", light: "bg-purple-50", text: "text-purple-600", ring: "ring-purple-100", label: "Teams", Icon: FiMessageSquare },
};

const READ_KEY = "_crm_notif_read";

export default function NotificationsPage() {
  const { instance, accounts } = useMsal();
  const navigate = useNavigate();

  const [items, setItems] = useState([]);
  const [loading, setLoading] = useState(false);
  const [filter, setFilter] = useState("all"); // all | email | teams
  const [search, setSearch] = useState("");
  const [readIds, setReadIds] = useState(() => {
    try { return new Set(JSON.parse(localStorage.getItem(READ_KEY) || "[]")); } catch { return new Set(); }
  });

  const msalRef = useRef({ accounts, instance });
  useEffect(() => { msalRef.current = { accounts, instance }; }, [accounts, instance]);

  const fetchNotifications = useCallback(async () => {
    const { accounts: accs, instance: inst } = msalRef.current;
    if (accs.length === 0) { setItems([]); setLoading(false); return; }
    setLoading(true);
    const account = accs[0];
    const byTime = (arr) => [...arr].sort((a, b) => new Date(b.time) - new Date(a.time));

    // Emails and Teams run concurrently; whichever finishes first is rendered.
    const emailsPromise = (async () => {
      try {
        const tok = await inst.acquireTokenSilent({
          scopes: ["https://graph.microsoft.com/User.Read", "https://graph.microsoft.com/Mail.Read"],
          account,
        });
        const res = await window.fetch(
          "https://graph.microsoft.com/v1.0/me/messages?$filter=isRead eq false&$top=50&$select=id,subject,from,receivedDateTime,bodyPreview,isRead,webLink&$orderby=receivedDateTime desc",
          { headers: { Authorization: `Bearer ${tok.accessToken}` } }
        );
        const data = await res.json();
        return (data.value || []).map((m) => ({
          id: m.id,
          source: "email",
          title: m.subject || "(No subject)",
          from: m.from?.emailAddress?.name || m.from?.emailAddress?.address || "Unknown",
          preview: m.bodyPreview || "",
          time: m.receivedDateTime,
          link: m.webLink || null,
        }));
      } catch { return []; }
    })();

    const teamsPromise = (async () => {
      try {
        const tok = await inst.acquireTokenSilent({
          scopes: ["https://graph.microsoft.com/User.Read", "https://graph.microsoft.com/Chat.Read"],
          account,
        });
        const authHeader = { headers: { Authorization: `Bearer ${tok.accessToken}` } };
        const chatsRes = await window.fetch("https://graph.microsoft.com/v1.0/me/chats?$top=10", authHeader);
        const chats = ((await chatsRes.json()).value || []).slice(0, 10);
        // All chats' messages fetched in parallel (was a sequential loop → the slow part).
        const groups = await Promise.all(
          chats.map(async (chat) => {
            try {
              const r = await window.fetch(`https://graph.microsoft.com/v1.0/me/chats/${chat.id}/messages?$top=3`, authHeader);
              const d = await r.json();
              return (d.value || [])
                .map((msg) => {
                  const body = msg.body?.content?.replace(/<[^>]*>/g, "").trim() || "";
                  if (!body || msg.messageType === "unknownFutureValue") return null;
                  return {
                    id: `teams-${msg.id}`,
                    source: "teams",
                    title: chat.topic || "Teams Chat",
                    from: msg.from?.user?.displayName || "Someone",
                    preview: body.slice(0, 200),
                    time: msg.createdDateTime,
                    link: null,
                  };
                })
                .filter(Boolean);
            } catch { return []; }
          })
        );
        return groups.flat();
      } catch { return []; }
    })();

    // Show emails the moment they land so the page isn't stuck behind Teams.
    const emails = await emailsPromise;
    setItems(byTime(emails));
    const teams = await teamsPromise;
    setItems(byTime([...emails, ...teams]));
    setLoading(false);
  }, []);

  useEffect(() => { fetchNotifications(); }, [fetchNotifications]);

  const markOneRead = (id) => {
    setReadIds((prev) => {
      const next = new Set(prev); next.add(id);
      try { localStorage.setItem(READ_KEY, JSON.stringify([...next])); } catch {}
      return next;
    });
  };
  const markAllRead = () => {
    const all = new Set(items.map((n) => n.id));
    setReadIds(all);
    try { localStorage.setItem(READ_KEY, JSON.stringify([...all])); } catch {}
  };

  const unreadCount = items.filter((n) => !readIds.has(n.id)).length;
  const counts = {
    all: items.length,
    email: items.filter((n) => n.source === "email").length,
    teams: items.filter((n) => n.source === "teams").length,
  };

  const visible = useMemo(() => {
    const q = search.trim().toLowerCase();
    return items.filter((n) => {
      if (filter !== "all" && n.source !== filter) return false;
      if (!q) return true;
      return [n.from, n.title, n.preview].some((v) => (v || "").toLowerCase().includes(q));
    });
  }, [items, filter, search]);

  const onItemClick = (n) => {
    markOneRead(n.id);
    if (n.link) window.open(n.link, "_blank", "noopener");
  };

  const TABS = [
    { key: "all", label: "All" },
    { key: "email", label: "Email" },
    { key: "teams", label: "Teams" },
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
            <span className="w-10 h-10 rounded-xl bg-blue-500 flex items-center justify-center text-white flex-shrink-0">
              <FiBell size={20} />
            </span>
            <div>
              <h1 className="text-xl font-bold text-gray-900 flex items-center gap-2">
                Notifications
                {unreadCount > 0 && (
                  <span className="px-2 py-0.5 text-[11px] font-bold bg-red-100 text-red-600 rounded-full">{unreadCount} new</span>
                )}
              </h1>
              <p className="text-sm text-gray-500">Unread emails and recent Teams messages from your Microsoft account.</p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            {unreadCount > 0 && (
              <button onClick={markAllRead}
                className="inline-flex items-center gap-1.5 text-sm text-blue-600 hover:text-blue-800 font-medium px-3 py-2 rounded-lg hover:bg-blue-50 transition-colors">
                <FiCheckCircle size={15} /> Mark all read
              </button>
            )}
            <button onClick={fetchNotifications}
              className="inline-flex items-center gap-1.5 text-sm text-gray-700 font-medium px-3 py-2 rounded-lg border border-gray-300 bg-white hover:bg-gray-50 transition-colors">
              <FiRefreshCw size={15} className={loading ? "animate-spin" : ""} /> Refresh
            </button>
          </div>
        </div>

        {/* Filter tabs + search */}
        <div className="flex flex-wrap items-center gap-2 mb-4">
          {TABS.map((t) => (
            <button key={t.key} onClick={() => setFilter(t.key)}
              className={`inline-flex items-center gap-2 px-3.5 py-1.5 rounded-lg text-sm font-medium border transition ${
                filter === t.key ? "bg-blue-500 text-white border-blue-600 shadow-sm" : "bg-white text-gray-700 border-gray-300 hover:bg-gray-50"
              }`}>
              {t.label}
              <span className={`text-[11px] font-bold px-1.5 rounded-full ${filter === t.key ? "bg-white/25 text-white" : "bg-gray-100 text-gray-500"}`}>{counts[t.key]}</span>
            </button>
          ))}
          <div className="relative flex-1 min-w-[180px]">
            <FiSearch size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
            <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search notifications…"
              className="w-full pl-9 pr-3 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
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
                    <div className="h-3 w-1/3 bg-gray-200 rounded" />
                    <div className="h-2.5 w-3/4 bg-gray-100 rounded" />
                  </div>
                </div>
              ))}
            </div>
          ) : visible.length === 0 ? (
            <div className="flex flex-col items-center justify-center py-20 px-4 text-center">
              <div className="w-16 h-16 rounded-full bg-gray-100 flex items-center justify-center mb-3">
                <FiBell size={28} className="text-gray-300" />
              </div>
              <p className="text-sm font-medium text-gray-500">{items.length === 0 ? "All caught up!" : "No matches"}</p>
              <p className="text-xs text-gray-400 mt-1">
                {accounts.length === 0
                  ? "Sign in with Microsoft to see emails & Teams messages."
                  : items.length === 0 ? "No new notifications right now." : "Try a different filter or search."}
              </p>
            </div>
          ) : (
            <ul className="divide-y divide-gray-100">
              {visible.map((n) => {
                const cfg = SOURCE[n.source] || SOURCE.email;
                const Icon = cfg.Icon;
                const isRead = readIds.has(n.id);
                return (
                  <li key={n.id}>
                    <button onClick={() => onItemClick(n)}
                      className="w-full text-left px-4 py-3.5 flex gap-3 hover:bg-gray-50 transition-colors group">
                      <div className={`w-10 h-10 rounded-xl ${cfg.color} flex items-center justify-center flex-shrink-0 text-white font-bold`}>
                        {n.from?.[0]?.toUpperCase() || <Icon size={16} />}
                      </div>
                      <div className="flex-1 min-w-0">
                        <div className="flex items-start justify-between gap-2">
                          <div className="flex items-center gap-2 min-w-0">
                            <p className={`text-sm font-semibold truncate ${isRead ? "text-gray-500" : "text-gray-900"}`}>{n.from}</p>
                            <span className={`text-[10px] font-semibold px-1.5 py-0.5 rounded-full ${cfg.light} ${cfg.text} flex-shrink-0`}>{cfg.label}</span>
                          </div>
                          <span className="text-[11px] text-gray-400 whitespace-nowrap flex-shrink-0" title={fullTime(n.time)}>{timeAgo(n.time)}</span>
                        </div>
                        <p className={`text-sm truncate mt-0.5 ${isRead ? "text-gray-400" : "text-gray-800"}`}>{n.title}</p>
                        {n.preview && <p className="text-xs text-gray-400 truncate mt-0.5">{n.preview}</p>}
                      </div>
                      <div className="flex flex-col items-end justify-between flex-shrink-0">
                        {!isRead && <div className="w-2.5 h-2.5 rounded-full bg-blue-500 mt-1.5" />}
                        {n.link && <FiExternalLink size={14} className="text-gray-300 group-hover:text-blue-500 transition-colors mt-auto" />}
                      </div>
                    </button>
                  </li>
                );
              })}
            </ul>
          )}
        </div>

        {!loading && items.length > 0 && (
          <p className="text-xs text-gray-400 text-center mt-3">
            Showing {visible.length} of {items.length} notification{items.length !== 1 ? "s" : ""}
          </p>
        )}
      </div>
    </div>
  );
}
