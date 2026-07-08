import React, { useEffect, useState, useRef, useMemo, useCallback } from "react";
import { useNavigate } from "react-router-dom";
import { RotateCcw, Trash2 } from "lucide-react";
import apiClient from "../api/client";

const entityLabel = (type) => {
  switch (type) {
    case "Account": return "Account";
    case "Contact": return "Contact";
    case "Deal": return "Deal";
    case "Product": return "Product";
    case "Note": return "Note";
    default: return type || "Item";
  }
};

function fmt(dt) {
  if (!dt) return "—";
  try {
    // DeletedAt is stored already in IST (the server converts UTC→IST before saving)
    // and carries no timezone marker. Render its wall-clock parts directly so the
    // value isn't shifted a second time by the viewer's browser timezone.
    const m = String(dt).match(/(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})(?::(\d{2}))?/);
    if (m) {
      const [, y, mo, d, h, mi, s] = m;
      const local = new Date(
        Number(y), Number(mo) - 1, Number(d),
        Number(h), Number(mi), Number(s || 0)
      );
      return local.toLocaleString("en-IN");
    }
    return new Date(dt).toLocaleString("en-IN", { timeZone: "Asia/Kolkata" });
  } catch { return String(dt); }
}

// Builds a compact page list with ellipses, e.g. [1, 2, 3, '…', 8, 9, 10]
function getPageNumbers(current, total) {
  if (total <= 7) return Array.from({ length: total }, (_, i) => i + 1);
  const pages = new Set([1, 2, total - 1, total, current - 1, current, current + 1]);
  const sorted = Array.from(pages).filter((p) => p >= 1 && p <= total).sort((a, b) => a - b);
  const result = [];
  let prev = null;
  for (const p of sorted) {
    if (prev !== null && p - prev > 1) result.push("…");
    result.push(p);
    prev = p;
  }
  return result;
}

export default function RecycleBin() {
  const navigate = useNavigate();
  const [items, setItems] = useState([]);
  const [loading, setLoading] = useState(true);
  const [actionLoading, setActionLoading] = useState(null);
  const [message, setMessage] = useState("");
  const messageTimerRef = useRef(null);

  const [search, setSearch] = useState("");
  const [entity, setEntity] = useState("");
  const [page, setPage] = useState(1);
  const [pageSize] = useState(10);

  // Restore confirmation modal state
  const [showRestoreModal, setShowRestoreModal] = useState(false);
  const [restoreTarget, setRestoreTarget] = useState(null);

  // Permanent delete confirmation modal state
  const [showDeleteModal, setShowDeleteModal] = useState(false);
  const [deleteTarget, setDeleteTarget] = useState(null);

  const showMessage = (text) => {
    if (messageTimerRef.current) clearTimeout(messageTimerRef.current);
    setMessage(text);
    messageTimerRef.current = setTimeout(() => setMessage(""), 3000);
  };

  // Fetches data only — doesn't touch loading skeleton, page, filters, or scroll on refresh.
  const fetchItems = useCallback(async ({ silent } = {}) => {
    try {
      if (!silent) setLoading(true);
      const response = await apiClient.get("/RecycleBin");
      setItems(response.data || []);
    } catch (error) {
      console.error("Failed to load recycle bin items", error);
      showMessage("Unable to load recycle bin items.");
    } finally {
      if (!silent) setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchItems();
    return () => {
      if (messageTimerRef.current) clearTimeout(messageTimerRef.current);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const handleRefresh = () => {
    fetchItems({ silent: true });
  };

  const entityTypes = useMemo(() => {
    const set = new Set(items.map((i) => i.entityType).filter(Boolean));
    return Array.from(set);
  }, [items]);

  const filteredItems = useMemo(() => {
    const q = search.trim().toLowerCase();
    return items.filter((item) => {
      if (entity && item.entityType !== entity) return false;
      if (!q) return true;
      const haystack = [item.displayName, entityLabel(item.entityType), item.deletedBy, item.details]
        .filter(Boolean)
        .join(" ")
        .toLowerCase();
      return haystack.includes(q);
    });
  }, [items, search, entity]);

  const totalPages = Math.max(1, Math.ceil(filteredItems.length / pageSize));
  const pagedItems = useMemo(() => {
    const start = (page - 1) * pageSize;
    return filteredItems.slice(start, start + pageSize);
  }, [filteredItems, page, pageSize]);

  // Keep page in range whenever filters shrink the result set.
  useEffect(() => {
    if (page > totalPages) setPage(totalPages);
  }, [totalPages, page]);

  const handleRestoreClick = (item) => {
    setRestoreTarget(item);
    setShowRestoreModal(true);
  };

  const confirmRestore = async () => {
    const item = restoreTarget;
    setShowRestoreModal(false);
    if (!item) return;
    try {
      setActionLoading(`${item.entityType}-${item.entityId}`);
      await apiClient.post(`/RecycleBin/${encodeURIComponent(item.entityType)}/${encodeURIComponent(item.entityId)}/restore`);
      showMessage(`${entityLabel(item.entityType)} restored.`);
      await fetchItems({ silent: true });
    } catch (error) {
      console.error("Restore failed", error);
      showMessage("Restore failed.");
    } finally {
      setActionLoading(null);
      setRestoreTarget(null);
    }
  };

  const handleDeleteClick = (item) => {
    setDeleteTarget(item);
    setShowDeleteModal(true);
  };

  const confirmPermanentDelete = async () => {
    const item = deleteTarget;
    setShowDeleteModal(false);
    if (!item) return;
    try {
      setActionLoading(`${item.entityType}-${item.entityId}-delete`);
      await apiClient.delete(`/RecycleBin/${encodeURIComponent(item.entityType)}/${encodeURIComponent(item.entityId)}`);
      showMessage(`${entityLabel(item.entityType)} permanently deleted.`);
      await fetchItems({ silent: true });
    } catch (error) {
      console.error("Permanent delete failed", error);
      showMessage("Permanent delete failed.");
    } finally {
      setActionLoading(null);
      setDeleteTarget(null);
    }
  };

  return (
    <div className="p-4 sm:p-6 h-full flex flex-col overflow-hidden bg-gray-50">
      <div className="max-w-7xl mx-auto w-full flex flex-col flex-1 min-h-0">
        {/* Non-scrolling header area: back button, message, filters */}
        <div className="flex-shrink-0">
          <button
            onClick={() => navigate(-1)}
            className="inline-flex items-center gap-1.5 text-sm text-gray-600 hover:text-gray-900 font-medium mb-3"
          >
            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 19l-7-7 7-7" />
            </svg>
            Back
          </button>
          <div className="flex items-center justify-between flex-wrap gap-3 mb-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-lg bg-gray-100 flex items-center justify-center flex-shrink-0">
                <Trash2 className="w-5 h-5 text-gray-500" />
              </div>
              <div>
                <h1 className="text-xl font-semibold text-gray-900">Recycle Bin</h1>
                <p className="text-sm text-gray-500">Recover or permanently remove deleted CRM records.</p>
              </div>
            </div>
            <span className="text-sm text-gray-500">{filteredItems.length} record{filteredItems.length !== 1 ? "s" : ""}</span>
          </div>

          {message && (
            <div className="mb-4 px-4 py-3 rounded-lg bg-emerald-50 border border-emerald-200 text-emerald-700 text-sm">{message}</div>
          )}

          {/* Filters */}
          <div className="flex flex-wrap items-center gap-2 mb-4">
            <input
              value={search}
              onChange={(e) => { setSearch(e.target.value); setPage(1); }}
              placeholder="Search item, deleted by, details…"
              className="flex-1 min-w-[200px] px-3 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
            <select value={entity} onChange={(e) => { setEntity(e.target.value); setPage(1); }}
              className="px-3 py-2 border border-gray-300 rounded-lg text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-500">
              <option value="">All entities</option>
              {entityTypes.map((en) => <option key={en} value={en}>{entityLabel(en)}</option>)}
            </select>
            <button onClick={handleRefresh}
              className="px-4 py-2 bg-blue-500 hover:bg-blue-600 text-white text-sm rounded-lg">Refresh</button>
          </div>
        </div>

        {/* Table: only this area scrolls; header row stays pinned via sticky */}
        <div className="bg-white border border-gray-200 rounded-xl flex-1 min-h-0 flex flex-col overflow-hidden">
          <div className="overflow-auto flex-1 min-h-0">
            <table className="w-full text-sm">
              <thead className="bg-gray-50 text-gray-500 text-xs uppercase tracking-wide sticky top-0 z-10">
                <tr>
                  <th className="text-left font-semibold px-4 py-3 whitespace-nowrap bg-gray-50">Item</th>
                  <th className="text-left font-semibold px-4 py-3 whitespace-nowrap bg-gray-50">Deleted By</th>
                  <th className="text-left font-semibold px-4 py-3 whitespace-nowrap bg-gray-50">Deleted At</th>
                  <th className="text-left font-semibold px-4 py-3 bg-gray-50">Details</th>
                  <th className="text-left font-semibold px-4 py-3 whitespace-nowrap bg-gray-50">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100">
                {loading ? (
                  <tr><td colSpan={5} className="px-4 py-16 text-center text-sm text-gray-500">Loading…</td></tr>
                ) : pagedItems.length === 0 ? (
                  <tr><td colSpan={5} className="px-4 py-16">
                    <div className="flex flex-col items-center text-center">
                      <div className="w-14 h-14 rounded-full bg-gray-100 flex items-center justify-center mb-3">
                        <Trash2 className="w-6 h-6 text-gray-400" />
                      </div>
                      <p className="text-sm font-medium text-gray-700">
                        {items.length === 0 ? "Recycle bin is empty" : "No matching items"}
                      </p>
                      <p className="text-sm text-gray-500 mt-1">
                        {items.length === 0 ? "Deleted records will appear here for recovery." : "Try adjusting your search or entity filter."}
                      </p>
                    </div>
                  </td></tr>
                ) : pagedItems.map((item) => (
                  <tr key={`${item.entityType}-${item.entityId}`} className="hover:bg-gray-50 align-top">
                    <td className="px-4 py-3 whitespace-nowrap">
                      <div className="font-medium text-gray-800">{item.displayName || entityLabel(item.entityType)}</div>
                      <div className="text-xs text-gray-500">{entityLabel(item.entityType)}</div>
                    </td>
                    <td className="px-4 py-3 whitespace-nowrap text-gray-600">{item.deletedBy || "System"}</td>
                    <td className="px-4 py-3 whitespace-nowrap text-gray-700">{fmt(item.deletedAt)}</td>
                    <td className="px-4 py-3 text-gray-500">{item.details || "—"}</td>
                    <td className="px-4 py-3 whitespace-nowrap">
                      <div className="flex items-center gap-2">
                        <button
                          onClick={() => handleRestoreClick(item)}
                          disabled={actionLoading === `${item.entityType}-${item.entityId}`}
                          className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-medium border bg-green-100 text-green-700 border-green-200 hover:bg-green-200 disabled:opacity-50"
                        >
                          <RotateCcw size={12} /> Restore
                        </button>
                        <button
                          onClick={() => handleDeleteClick(item)}
                          disabled={actionLoading === `${item.entityType}-${item.entityId}-delete`}
                          className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-medium border bg-red-100 text-red-700 border-red-200 hover:bg-red-200 disabled:opacity-50"
                        >
                          <Trash2 size={12} /> Delete
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Pagination: stays pinned below the scrollable table */}
        <div className="flex items-center justify-between mt-4 text-sm flex-shrink-0">
          <span className="text-gray-500">Page {page} of {totalPages}</span>
          <div className="flex items-center gap-1">
            <button
              disabled={page <= 1}
              onClick={() => setPage((p) => Math.max(1, p - 1))}
              className="px-3 py-1.5 rounded-lg border border-gray-300 bg-white disabled:opacity-50 hover:bg-gray-50"
            >
              Prev
            </button>
            {/* {getPageNumbers(page, totalPages).map((p, idx) =>
              p === "…" ? (
                <span key={`ellipsis-${idx}`} className="px-2 text-gray-400">…</span>
              ) : (
                <button
                  key={p}
                  onClick={() => setPage(p)}
                  className={`px-3 py-1.5 rounded-lg border text-sm ${
                    p === page
                      ? "bg-blue-500 border-slate-800 text-white"
                      : "border-gray-300 bg-white text-gray-700 hover:bg-gray-50"
                  }`}
                >
                  {p}
                </button>
              )
            )} */}
            <button
              disabled={page >= totalPages}
              onClick={() => setPage((p) => Math.min(totalPages, p + 1))}
              className="px-3 py-1.5 rounded-lg border border-gray-300 bg-white disabled:opacity-50 hover:bg-gray-50"
            >
              Next
            </button>
          </div>
        </div>
      </div>

      {/* Restore Confirmation Modal */}
      {showRestoreModal && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-[9999] p-4 animate-in fade-in duration-200" role="dialog" aria-modal="true">
          <div className="bg-white w-full sm:w-96 rounded-xl border border-gray-200 shadow-xl p-6 transform animate-in zoom-in-95 duration-200">
            <div className="mb-6 text-center">
              <div className="w-14 h-14 bg-emerald-100 rounded-full flex items-center justify-center mx-auto mb-4">
                <RotateCcw className="w-7 h-7 text-emerald-600" />
              </div>
              <h4 className="text-lg sm:text-xl font-semibold text-gray-900 mb-2">Restore Item</h4>
              <p className="text-sm sm:text-base text-gray-600">
                Are you sure you want to restore{" "}
                <span className="font-semibold text-gray-900">
                  {restoreTarget?.displayName || entityLabel(restoreTarget?.entityType)}
                </span>
                ?
              </p>
            </div>
            <div className="flex gap-3 justify-center">
              <button
                type="button"
                onClick={() => { setShowRestoreModal(false); setRestoreTarget(null); }}
                className="px-6 py-2.5 border border-gray-300 bg-white hover:bg-gray-50 text-gray-700 rounded-lg font-medium text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={confirmRestore}
                className="px-6 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg font-medium text-sm focus:ring-2 focus:ring-emerald-500 focus:outline-none"
              >
                Restore
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Permanent Delete Confirmation Modal */}
      {showDeleteModal && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-[9999] p-4 animate-in fade-in duration-200" role="dialog" aria-modal="true">
          <div className="bg-white w-full sm:w-96 rounded-xl border border-gray-200 shadow-xl p-6 transform animate-in zoom-in-95 duration-200">
            <div className="mb-6 text-center">
              <div className="w-14 h-14 bg-red-100 rounded-full flex items-center justify-center mx-auto mb-4">
                <svg className="w-7 h-7 text-red-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
                </svg>
              </div>
              <h4 className="text-lg sm:text-xl font-semibold text-gray-900 mb-2">Confirm Delete</h4>
              <p className="text-sm sm:text-base text-gray-600">
                Are you sure you want to permanently delete{" "}
                <span className="font-semibold text-gray-900">
                  {deleteTarget?.displayName || entityLabel(deleteTarget?.entityType)}
                </span>
                ? This action cannot be undone.
              </p>
            </div>
            <div className="flex gap-3 justify-center">
              <button
                type="button"
                onClick={() => { setShowDeleteModal(false); setDeleteTarget(null); }}
                className="px-6 py-2.5 border border-gray-300 bg-white hover:bg-gray-50 text-gray-700 rounded-lg font-medium text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={confirmPermanentDelete}
                className="px-6 py-2.5 bg-red-600 hover:bg-red-700 text-white rounded-lg font-medium text-sm focus:ring-2 focus:ring-red-500 focus:outline-none"
              >
                Delete
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}