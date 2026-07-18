import React, { useState, useEffect, useMemo, useRef } from "react";
import { createPortal } from "react-dom";

// Sent one message per person, this many at a time, so a big list doesn't fire thousands
// of Graph calls back to back.
export const SEND_BATCH_SIZE = 150;

// One page of the list is exactly one send batch, so "page 2 of 4" and "batch 2 of 4" are the
// same 150 people. Keeping these equal is the point — don't split them.
export const PAGE_SIZE = SEND_BATCH_SIZE;

export const PREFIX_MR = "Mr.";
export const PREFIX_MISS = "Miss.";

// The greeting each recipient's copy opens with. Kept here so the preview in this modal and
// the line actually sent can never drift apart — both call this.
export function greetingFor(recipient) {
  const name = (recipient.name || "").trim();
  const prefix = recipient.prefix || "";
  if (!prefix) return "";
  return name ? `Dear ${prefix} ${name},` : `Dear ${prefix},`;
}

// Turns raw contact rows into the shape this modal works with, defaulting everyone to Mr.
// contactId is carried through untouched so the send can tie each address back to its contact.
// `sent` marks people a previous batch already covered — they come back unticked and labelled
// "Sent" (instead of "Will not be emailed"), so a resumed run can't email anyone twice.
export function toPersonalizedRecipients(rows) {
  return (rows || []).map((r) => ({
    name: (r.name || "").trim(),
    email: (r.email || "").trim(),
    contactId: r.contactId ?? null,
    prefix: r.prefix === undefined ? PREFIX_MR : r.prefix,
    sent: !!r.sent,
  }));
}

/**
 * Confirms who gets the email and how each is addressed, before anything is sent.
 * Every row starts at "Mr."; picking Miss swaps it, and clearing both drops that
 * person from the send (the row greys out and the count updates).
 *
 * ONE PAGE AT A TIME: onConfirm receives only the included recipients on the CURRENT page —
 * at most PAGE_SIZE (150) — never the whole list. A long list is worked through a page per
 * send: confirm this page, send it, come back and pick the next. This keeps one send to one
 * page of people, which is why page size and batch size are the same number.
 *
 * onConfirm receives: (pageRecipients, allRows) — the included recipients on the current page
 * [{ name, email, contactId, prefix }], plus EVERY row with its current prefix/exclusion state.
 * Callers keep allRows so the popup can be reopened after a send with nothing lost.
 */
function PersonalizeRecipientsModal({
  isOpen,
  recipients,
  onConfirm,
  onCancel,
  isDark = false,
  title = "Personalize recipients",
  confirmLabel = "Continue",
  // Page to land on when the popup opens — used after a mid-run send so the user arrives at
  // the first page that still has someone left to send, instead of back at page 1.
  initialPage = 0,
}) {
  const [rows, setRows] = useState([]);
  const [query, setQuery] = useState("");
  const [page, setPage] = useState(0);
  const wasOpen = useRef(false);
  // How far the dialog has been dragged from its centered position. Grabbing the header with
  // the cursor moves the whole page anywhere on screen; reset to center on every open.
  const [drag, setDrag] = useState({ x: 0, y: 0 });
  const startDrag = (e) => {
    if (e.button !== undefined && e.button !== 0) return; // left button / touch only
    const start = { px: e.clientX, py: e.clientY, ox: drag.x, oy: drag.y };
    const move = (ev) => setDrag({ x: start.ox + ev.clientX - start.px, y: start.oy + ev.clientY - start.py });
    const up = () => {
      document.removeEventListener("pointermove", move);
      document.removeEventListener("pointerup", up);
    };
    document.addEventListener("pointermove", move);
    document.addEventListener("pointerup", up);
  };

  // Load on OPEN only — deliberately not whenever `recipients` changes identity. A caller that
  // passes a freshly built array on each render would otherwise wipe the user's ticks on every
  // keystroke in the search box.
  useEffect(() => {
    if (isOpen && !wasOpen.current) {
      setRows(toPersonalizedRecipients(recipients));
      setQuery("");
      setPage(initialPage || 0);
      setDrag({ x: 0, y: 0 });
    }
    wasOpen.current = isOpen;
  }, [isOpen, recipients, initialPage]);

  // Search first, then page. One page holds exactly one send batch, so a page IS a shot.
  const matching = useMemo(() => {
    const q = query.trim().toLowerCase();
    const withIndex = rows.map((r, i) => ({ ...r, i }));
    if (!q) return withIndex;
    return withIndex.filter(
      (r) => r.name.toLowerCase().includes(q) || r.email.toLowerCase().includes(q)
    );
  }, [rows, query]);

  const pageCount = Math.max(1, Math.ceil(matching.length / PAGE_SIZE));
  // Clamped rather than stored blind: a search that shrinks the list must not strand the user
  // on a page that no longer exists.
  const currentPage = Math.min(page, pageCount - 1);
  const pageStart = currentPage * PAGE_SIZE;
  const visible = useMemo(
    () => matching.slice(pageStart, pageStart + PAGE_SIZE),
    [matching, pageStart]
  );

  if (!isOpen) return null;

  // Mr and Miss behave as one choice: ticking one clears the other, and unticking the
  // ticked one leaves the contact with no prefix, which means "don't email this person".
  const setPrefix = (index, prefix) =>
    setRows((prev) => prev.map((r, i) => (i === index ? { ...r, prefix } : r)));

  // Bulk actions apply to THIS page only — 150 rows is the unit the user is looking at and
  // the unit that goes out in one shot. Rows already SENT are skipped, matching their disabled
  // checkboxes: someone a previous batch covered can never be re-included.
  const applyToPage = (prefix) => {
    const targets = new Set(visible.filter((r) => !r.sent).map((r) => r.i));
    setRows((prev) => prev.map((r, i) => (targets.has(i) ? { ...r, prefix } : r)));
  };

  const selectableOnPage = visible.filter((r) => !r.sent).length;
  const includedOnPage = visible.filter((r) => r.prefix).length;
  const allOnPageIncluded = selectableOnPage > 0 && includedOnPage === selectableOnPage;
  const someOnPageIncluded = includedOnPage > 0 && !allOnPageIncluded;
  const surface = isDark ? "bg-[#252423] text-gray-100" : "bg-white text-gray-900";
  const border = isDark ? "border-[#3d3b39]" : "border-gray-200";
  const muted = isDark ? "text-gray-400" : "text-gray-500";
  const headRow = isDark ? "bg-[#1b1a19]" : "bg-gray-50";

  // Rendered into document.body, not wherever it was mounted. The Contacts page mounts this near
  // the page root while the composer mounts it deep inside a panel; portalling guarantees both
  // get the same full-screen dialog rather than one being boxed in by an ancestor.
  return createPortal(
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-[1500] p-3 sm:p-6">
      {/* Sized to work: 150 rows need room, so this takes most of the viewport rather than a
          narrow dialog you'd scroll forever in. */}
      <div
        className={`${surface} rounded-xl shadow-2xl w-full max-w-6xl flex flex-col h-[92vh]`}
        style={{ transform: `translate(${drag.x}px, ${drag.y}px)` }}
      >
        {/* Header — grab it with the cursor to move the whole page anywhere on screen. */}
        <div className={`px-6 py-3 border-b ${border} shrink-0 cursor-move select-none`} onPointerDown={startDrag}>
          <h2 className="text-base font-semibold">{title}</h2>
          <p className={`text-xs ${muted} mt-0.5`}>
            Everyone starts as <strong>Mr.</strong> — switch to <strong>Miss.</strong>, or clear both
            to leave someone out. Each gets their own email opening <em>Dear Mr. &lt;name&gt;,</em>
            {pageCount > 1 && (
              <> · Only this page is sent ({PAGE_SIZE} max) — send it, then use <strong>Next</strong> for the rest.</>
            )}
          </p>
        </div>

        {/* Toolbar */}
        <div className={`px-6 py-3 border-b ${border} flex items-center gap-3 flex-wrap shrink-0`}>
          <input
            type="text"
            value={query}
            onChange={(e) => { setQuery(e.target.value); setPage(0); }}
            placeholder="Search name or email…"
            className={`flex-1 min-w-[180px] px-3 py-1.5 border rounded-lg text-sm outline-none focus:ring-2 focus:ring-blue-500 ${
              isDark ? "bg-[#1b1a19] border-[#3d3b39]" : "bg-white border-gray-300"
            }`}
          />
          <span className={`text-xs ${muted}`}>Set this page:</span>
          <button
            type="button"
            onClick={() => applyToPage(PREFIX_MR)}
            className={`px-2.5 py-1 text-xs rounded-lg border ${border} hover:bg-blue-50 hover:text-blue-700`}
          >
            Mr.
          </button>
          <button
            type="button"
            onClick={() => applyToPage(PREFIX_MISS)}
            className={`px-2.5 py-1 text-xs rounded-lg border ${border} hover:bg-blue-50 hover:text-blue-700`}
          >
            Miss.
          </button>
          <button
            type="button"
            onClick={() => applyToPage("")}
            className={`px-2.5 py-1 text-xs rounded-lg border ${border} hover:bg-red-50 hover:text-red-600`}
          >
            None
          </button>
        </div>

        {/* Rows */}
        <div className="flex-1 overflow-y-auto">
          <table className="w-full text-sm border-collapse">
            <thead className={`${headRow} sticky top-0 z-10`}>
              <tr className={`border-b ${border}`}>
                <th className={`text-right font-semibold pl-6 pr-2 py-2 ${muted} text-xs uppercase tracking-wide w-14`}>#</th>
                <th className={`text-left font-semibold px-3 py-2 ${muted} text-xs uppercase tracking-wide`}>
                  {/* Select all on THIS page — never beyond it, so a tick can only ever include the
                      150 in front of you, which is also exactly one send batch. */}
                  <label className="inline-flex items-center gap-2 cursor-pointer select-none normal-case">
                    <input
                      type="checkbox"
                      aria-label="Select all on this page"
                      checked={allOnPageIncluded}
                      ref={(el) => { if (el) el.indeterminate = someOnPageIncluded; }}
                      onChange={() => applyToPage(allOnPageIncluded ? "" : PREFIX_MR)}
                      className="w-4 h-4 accent-blue-600 cursor-pointer"
                    />
                    <span className="text-xs uppercase tracking-wide font-semibold">
                      Contact {visible.length > 0 && <span className="normal-case font-normal">({includedOnPage}/{visible.length})</span>}
                    </span>
                  </label>
                </th>
                <th className={`text-center font-semibold px-3 py-2 ${muted} text-xs uppercase tracking-wide w-20`}>Mr.</th>
                <th className={`text-center font-semibold px-3 py-2 ${muted} text-xs uppercase tracking-wide w-20`}>Miss.</th>
                <th className={`text-left font-semibold px-4 py-2 ${muted} text-xs uppercase tracking-wide`}>Email opens with</th>
              </tr>
            </thead>
            <tbody>
              {visible.length === 0 && (
                <tr>
                  <td colSpan={5} className={`px-6 py-8 text-center ${muted}`}>
                    No contacts match “{query}”.
                  </td>
                </tr>
              )}
              {visible.map((r, idx) => {
                const excluded = !r.prefix;
                return (
                  // A SENT row is NOT faded like an excluded one — its "Sent ✓" badge must stay
                  // clearly readable; the muted name + locked checkboxes already say it's done.
                  <tr key={`${r.email}-${r.i}`} className={`border-b ${border} ${excluded && !r.sent ? "opacity-40" : ""}`}>
                    {/* Numbering runs across pages, not per page: page 2 starts at 151, so the
                        count reads continuously through the whole list. */}
                    <td className={`pl-6 pr-2 py-2.5 text-right text-xs tabular-nums ${muted}`}>
                      {pageStart + idx + 1}
                    </td>
                    <td className="px-3 py-2.5">
                      <div className={`font-medium truncate ${r.sent ? muted : ""}`}>{r.name || <span className={muted}>— no name —</span>}</div>
                      <div className={`text-xs ${muted} truncate`}>{r.email}</div>
                    </td>
                    {/* A SENT row is locked — its email already went out, so it can't be ticked
                        back in and emailed twice. */}
                    <td className="px-3 py-2.5 text-center">
                      <input
                        type="checkbox"
                        aria-label={`Address ${r.email} as Mr.`}
                        checked={r.prefix === PREFIX_MR}
                        disabled={r.sent}
                        onChange={(e) => setPrefix(r.i, e.target.checked ? PREFIX_MR : "")}
                        className={`w-4 h-4 accent-blue-600 ${r.sent ? "cursor-not-allowed" : "cursor-pointer"}`}
                      />
                    </td>
                    <td className="px-3 py-2.5 text-center">
                      <input
                        type="checkbox"
                        aria-label={`Address ${r.email} as Miss.`}
                        checked={r.prefix === PREFIX_MISS}
                        disabled={r.sent}
                        onChange={(e) => setPrefix(r.i, e.target.checked ? PREFIX_MISS : "")}
                        className={`w-4 h-4 accent-blue-600 ${r.sent ? "cursor-not-allowed" : "cursor-pointer"}`}
                      />
                    </td>
                    <td className={`px-4 py-2.5 text-xs ${excluded ? "" : "italic"}`}>
                      {excluded
                        ? (r.sent
                            ? <span className={`inline-flex items-center px-2 py-0.5 rounded-full font-semibold ${
                                isDark ? "bg-emerald-900/40 text-emerald-300" : "bg-emerald-100 text-emerald-700"
                              }`}>Sent ✓</span>
                            : <span className={muted}>Will not be emailed</span>)
                        : greetingFor(r)}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {/* Pager — one page is one send batch of PAGE_SIZE */}
        {/* Footer — counts, paging and the actions in ONE bar. Kept as a single short row so it
            always fits: as separate stacked bars with long prose it was getting pushed out of
            view on shorter screens. */}
        <div className={`px-6 py-3 border-t ${border} flex items-center justify-between gap-x-4 gap-y-2 flex-wrap shrink-0`}>
          <span className={`text-xs ${muted} tabular-nums`}>
            <strong className={isDark ? "text-gray-100" : "text-gray-900"}>
              {pageStart + 1}–{pageStart + visible.length}
            </strong>{" "}
            of {matching.length}
            {query.trim() ? " matching" : ""} · {includedOnPage} selected here
          </span>

          {pageCount > 1 && (
            <div className="flex items-center gap-2">
              <button
                type="button"
                disabled={currentPage === 0}
                onClick={() => setPage(currentPage - 1)}
                className={`px-3 py-1 text-xs rounded-lg border ${border} disabled:opacity-40 disabled:cursor-not-allowed hover:bg-blue-50 hover:text-blue-700`}
              >
                ‹ Prev
              </button>
              <span className={`text-xs ${muted} tabular-nums whitespace-nowrap`}>
                Page {currentPage + 1} of {pageCount}
              </span>
              <button
                type="button"
                disabled={currentPage >= pageCount - 1}
                onClick={() => setPage(currentPage + 1)}
                className={`px-3 py-1 text-xs rounded-lg border ${border} disabled:opacity-40 disabled:cursor-not-allowed hover:bg-blue-50 hover:text-blue-700`}
              >
                Next ›
              </button>
            </div>
          )}

          <div className="flex gap-3 items-center">
            <button
              type="button"
              onClick={onCancel}
              className={`px-4 py-2 border ${border} rounded-lg text-sm hover:bg-gray-50 ${isDark ? "hover:bg-[#1b1a19]" : ""}`}
            >
              Cancel
            </button>
            <button
              type="button"
              disabled={includedOnPage === 0}
              // THIS PAGE ONLY — `visible`, not `rows`. One send covers one page of people.
              // `rows` rides along so the caller can reopen this popup mid-run (page 2, 3, …)
              // with every tick exactly as the user left it.
              onClick={() => onConfirm(visible.filter((r) => r.prefix).map(({ i, ...r }) => r), rows)}
              className="px-5 py-2 bg-blue-600 hover:bg-blue-700 disabled:opacity-40 disabled:cursor-not-allowed text-white rounded-lg text-sm font-medium whitespace-nowrap"
            >
              {confirmLabel}{includedOnPage > 0 && ` (${includedOnPage})`}
            </button>
          </div>
        </div>
      </div>
    </div>,
    document.body
  );
}

export default PersonalizeRecipientsModal;
