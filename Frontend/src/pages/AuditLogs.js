import React, { useState, useEffect, useCallback, useMemo } from "react";
import { useNavigate } from "react-router-dom";
import apiClient from "../api/client";

const ACTION_STYLES = {
  Insert: "bg-green-100 text-green-700 border-green-200",
  Create: "bg-green-100 text-green-700 border-green-200",
  Update: "bg-amber-100 text-amber-700 border-amber-200",
  Delete: "bg-red-100 text-red-700 border-red-200",
  Login: "bg-sky-100 text-sky-700 border-sky-200",
  SignIn: "bg-sky-100 text-sky-700 border-sky-200",
};

function TabButton({ active, onClick, label, accent }) {
  const activeClasses = accent === "sky"
    ? "bg-sky-600 text-white border-sky-600 shadow-sm"
    : "bg-indigo-600 text-white border-indigo-600 shadow-sm";
  const idleClasses = "bg-white text-gray-700 border-gray-300 hover:bg-gray-50";

  return (
    <button
      type="button"
      onClick={onClick}
      className={`inline-flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium border transition ${active ? activeClasses : idleClasses}`}
    >
      <span>{label}</span>
      <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 7l5 5m0 0l-5 5m5-5H4" />
      </svg>
    </button>
  );
}

function fmt(dt) {
  if (!dt) return "-";
  try { return new Date(dt).toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }); } catch { return String(dt); }
}

function getLogCategory(row) {
  const action = String(row?.action ?? row?.Action ?? "").toLowerCase();
  const entity = String(row?.entityName ?? row?.EntityName ?? "").toLowerCase();
  const changes = String(row?.changes ?? row?.Changes ?? "").toLowerCase();

  if (["login", "signin", "sign-in", "sign_in", "auth", "loginattempt", "login-attempt"].includes(action)
    || ["login", "signin", "sign-in", "sign_in", "auth"].includes(entity)
    || changes.includes("login")
    || changes.includes("signin")) {
    return "login";
  }
  return "crud";
}

// Pretty-print the Changes JSON for a row, with a friendly per-action message.
function ChangesCell({ changes, action }) {
  const [open, setOpen] = useState(false);

  let parsed = null;
  if (changes) {
    try { parsed = JSON.parse(changes); } catch { return <span className="text-xs text-gray-600 break-all">{changes}</span>; }
  }
  const entries = parsed ? Object.entries(parsed) : [];

  if (entries.length === 0) {
    if (action === "Insert" || action === "Create") return <span className="text-green-600 font-medium">Created successfully</span>;
    if (action === "Delete") return <span className="text-red-600 font-medium">Deleted successfully</span>;
    return <span className="text-gray-400">—</span>;
  }

  const verb = action === "Insert" || action === "Create" ? "Created" : action === "Delete" ? "Deleted" : "Updated";
  const summary = entries.slice(0, 2).map(([k]) => k).join(", ") + (entries.length > 2 ? ` +${entries.length - 2}` : "");

  return (
    <div className="text-xs">
      <button type="button" onClick={() => setOpen((o) => !o)}
        className="text-indigo-600 hover:text-indigo-800 font-medium">
        {open ? "Hide" : `${verb} · ${entries.length} field${entries.length > 1 ? "s" : ""} (${summary})`}
      </button>
      {open && (
        <div className="mt-2 space-y-1 max-w-md">
          {entries.map(([field, val]) => {
            const isDiff = val && typeof val === "object" && ("old" in val || "new" in val);
            return (
              <div key={field} className="bg-gray-50 border border-gray-200 rounded px-2 py-1">
                <span className="font-semibold text-gray-700">{field}</span>
                {isDiff ? (
                  <span className="ml-1">
                    <span className="line-through text-red-500 break-all">{String(val.old ?? "∅")}</span>
                    <span className="mx-1 text-gray-400">→</span>
                    <span className="text-green-600 break-all">{String(val.new ?? "∅")}</span>
                  </span>
                ) : (
                  <span className="ml-1 text-gray-600 break-all">{String(val ?? "")}</span>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

export default function AuditLogs() {
  const navigate = useNavigate();
  const [rows, setRows] = useState([]);
  const [total, setTotal] = useState(0);
  const [page, setPage] = useState(1);
  const [pageSize] = useState(50);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);

  const [search, setSearch] = useState("");
  const [entity, setEntity] = useState("");
  const [action, setAction] = useState("");
  const [entities, setEntities] = useState([]);
  const [activeTab, setActiveTab] = useState("login");

  useEffect(() => {
    apiClient.get("/AuditLog/entities")
      .then((r) => setEntities(Array.isArray(r.data) ? r.data : []))
      .catch(() => setEntities([]));
  }, []);

  const fetchLogs = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await apiClient.get("/AuditLog", {
        params: { search: search || undefined, entity: entity || undefined, action: action || undefined, page, pageSize },
      });
      const data = res.data || {};
      setRows(Array.isArray(data.items) ? data.items : []);
      setTotal(Number(data.totalCount) || 0);
    } catch (err) {
      console.error("Audit log fetch failed", err);
      setError(err?.response?.data?.message || err?.message || "Failed to load audit logs (Admins only).");
      setRows([]);
      setTotal(0);
    } finally {
      setLoading(false);
    }
  }, [search, entity, action, page, pageSize]);

  useEffect(() => { fetchLogs(); }, [fetchLogs]);

  const totalPages = Math.max(1, Math.ceil(total / pageSize));

  const loginRows = useMemo(() => rows.filter((row) => getLogCategory(row) === "login"), [rows]);
  const crudRows = useMemo(() => rows.filter((row) => getLogCategory(row) === "crud"), [rows]);
  const activeRows = activeTab === "login" ? loginRows : crudRows;
  const activeCount = activeRows.length;

  return (
    <div className="p-4 sm:p-6 h-full overflow-hidden bg-gray-50">
      <div className="max-w-7xl mx-auto h-full flex flex-col min-h-0">
        <div className="sticky top-0 z-10 bg-gray-50 pb-3">
          <button
            onClick={() => navigate(-1)}
            className="inline-flex items-center gap-1.5 text-sm text-gray-600 hover:text-gray-900 font-medium mb-3"
          >
            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 19l-7-7 7-7" />
            </svg>
            Back
          </button>
          </div>
        <div className="flex items-center justify-between flex-wrap gap-3 mb-4">
          <div>
            <h1 className="text-xl font-bold text-gray-900">Audit Logs</h1>
            <p className="text-sm text-gray-500">Login history is shown separately from CRUD operations so it is easier to review.</p>
          </div>
          <span className="text-sm text-gray-500">{activeCount} {activeTab === "login" ? "login" : "CRUD"} record{activeCount !== 1 ? "s" : ""}</span>
        </div>

        <div className="flex flex-wrap gap-2 mb-4">
          <TabButton active={activeTab === "crud"} onClick={() => setActiveTab("crud")} label="Activity" accent="indigo" />
          <TabButton active={activeTab === "login"} onClick={() => setActiveTab("login")} label="Login History" accent="sky" />
        </div>

        {/* Filters */}
        <div className="flex flex-wrap items-center gap-2 mb-4">
          <input
            value={search}
            onChange={(e) => { setSearch(e.target.value); setPage(1); }}
            placeholder="Search user, entity, record, change…"
            className="flex-1 min-w-[200px] px-3 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
          />
          <select value={entity} onChange={(e) => { setEntity(e.target.value); setPage(1); }}
            className="px-3 py-2 border border-gray-300 rounded-lg text-sm bg-white focus:outline-none focus:ring-2 focus:ring-indigo-500">
            <option value="">All entities</option>
            {entities.map((en) => <option key={en} value={en}>{en}</option>)}
          </select>
          <select value={action} onChange={(e) => { setAction(e.target.value); setPage(1); }}
            className="px-3 py-2 border border-gray-300 rounded-lg text-sm bg-white focus:outline-none focus:ring-2 focus:ring-indigo-500">
            <option value="">All actions</option>
            {["Insert", "Create", "Update", "Delete", "Login"].map((a) => <option key={a} value={a}>{a}</option>)}
          </select>
          <button onClick={() => fetchLogs()}
            className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white text-sm rounded-lg">Refresh</button>
        </div>

        {error && (
          <div className="mb-4 px-4 py-3 rounded-lg bg-red-50 border border-red-200 text-red-700 text-sm">{error}</div>
        )}

        {/* Table */}
        <div className="flex-1 min-h-0 overflow-auto bg-white border border-gray-200 rounded-xl">
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead className="bg-gray-50 text-gray-500 text-xs uppercase tracking-wide sticky top-0 z-10">
                <tr>
                  <th className="text-left font-semibold px-3 py-2.5 whitespace-nowrap w-[220px]">When</th>
                  <th className="text-left font-semibold px-3 py-2.5 whitespace-nowrap w-[190px]">User</th>
                  <th className="text-left font-semibold px-3 py-2.5 whitespace-nowrap w-[110px]">Role</th>
                  <th className="text-left font-semibold px-3 py-2.5 whitespace-nowrap w-[110px]">Action</th>
                  <th className="text-left font-semibold px-3 py-2.5 whitespace-nowrap w-[130px]">Entity</th>
                  <th className="text-left font-semibold px-3 py-2.5 whitespace-nowrap w-[120px]">Record</th>
                  <th className="text-left font-semibold px-3 py-2.5">Changes</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100">
                {loading ? (
                  <tr><td colSpan={7} className="px-4 py-10 text-center text-gray-500">Loading…</td></tr>
                ) : activeRows.length === 0 ? (
                  <tr><td colSpan={7} className="px-4 py-10 text-center text-gray-400">No {activeTab === "login" ? "login history" : "Operation"} entries.</td></tr>
                ) : activeRows.map((r) => (
                  <tr key={r.id ?? r.Id} className={`hover:bg-gray-50 align-top border-l-4 ${activeTab === "login" ? "border-sky-400" : "border-indigo-400"}`}>
                    <td className="px-3 py-2.5 whitespace-nowrap text-gray-700 w-[220px]">
                      <div className="flex items-center gap-2">
                        <div className={`flex h-8 w-8 items-center justify-center rounded-full ${activeTab === "login" ? "bg-sky-100 text-sky-600" : "bg-indigo-100 text-indigo-600"}`}>
                          <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 7l5 5m0 0l-5 5m5-5H4" />
                          </svg>
                        </div>
                        <div>
                          <div className="font-medium text-gray-800">{fmt(r.changedAt ?? r.ChangedAt)}</div>
                          <div className="text-xs text-gray-500">{activeTab === "login" ? "Login activity" : "Record activity"}</div>
                        </div>
                      </div>
                    </td>
                    <td className="px-3 py-2.5 whitespace-nowrap w-[190px]">
                      <div className="font-medium text-gray-800">{r.changedByName ?? r.ChangedByName ?? "—"}</div>
                      <div className="text-xs text-gray-500">{r.changedBy ?? r.ChangedBy}</div>
                    </td>
                    <td className="px-3 py-2.5 whitespace-nowrap text-gray-600 w-[110px]">{r.changedByRole ?? r.ChangedByRole ?? "—"}</td>
                    <td className="px-3 py-2.5 whitespace-nowrap w-[110px]">
                      <span className={`px-2 py-0.5 rounded-full text-xs font-medium border ${ACTION_STYLES[r.action ?? r.Action] || "bg-gray-100 text-gray-700 border-gray-200"}`}>
                        {r.action ?? r.Action}
                      </span>
                    </td>
                    <td className="px-3 py-2.5 whitespace-nowrap text-gray-700 w-[130px]">{r.entityName ?? r.EntityName ?? "—"}</td>
                    <td className="px-3 py-2.5 whitespace-nowrap text-gray-500 w-[120px]">{r.entityId ?? r.EntityId ?? "—"}</td>
                    <td className="px-3 py-2.5"><ChangesCell changes={r.changes ?? r.Changes} action={r.action ?? r.Action} /></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Pagination */}
        <div className="flex items-center justify-between mt-4 text-sm">
          <span className="text-gray-500">Page {page} of {totalPages}</span>
          <div className="flex gap-2">
            <button disabled={page <= 1} onClick={() => setPage((p) => Math.max(1, p - 1))}
              className="px-3 py-1.5 rounded-lg border border-gray-300 bg-white disabled:opacity-50">Prev</button>
            <button disabled={page >= totalPages} onClick={() => setPage((p) => Math.min(totalPages, p + 1))}
              className="px-3 py-1.5 rounded-lg border border-gray-300 bg-white disabled:opacity-50">Next</button>
          </div>
        </div>
      </div>
    </div>
  );
}
