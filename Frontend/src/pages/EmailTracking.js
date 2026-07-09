import React, { useState, useEffect, useCallback, useRef } from "react";
import { useMsal } from "@azure/msal-react";
import apiClient from "../api/client";
import {
  Send, MailOpen, MousePointerClick, UserMinus, AlertTriangle, Megaphone,
  RefreshCw, ArrowLeft, ChevronRight, ChevronDown, Inbox, Reply, Filter, Gauge, MailX,
} from "lucide-react";

// ── Freshworks "Crayons" palette (their real product colours) ─────────────────
const INK = "#12344d";       // elephant-900  — headings / big numbers
const BODY = "#475867";      // smoke-700     — body text
const MUTED = "#8a98a5";     // smoke-300ish  — muted labels
const BORDER = "#e6ebef";    // smoke-50/100  — card borders
const SURFACE = "#f5f7f9";   // smoke-25      — page background
const PRIMARY = "#2c5cc5";   // azure-800     — primary action / links

// One brand hue per metric (drawn from the Crayons families): c = solid, l = soft/light, tint = pale.
const M = {
  campaigns:    { c: "#5b6b7b", l: "#b3bcc4", tint: "#eef1f4", icon: Megaphone },
  sent:         { c: "#2c5cc5", l: "#a6c1ee", tint: "#e8f0fd", icon: Send },            // azure
  opened:       { c: "#00a886", l: "#8ed6c5", tint: "#e0f5f1", icon: MailOpen },        // jungle teal
  clicked:      { c: "#6c44b4", l: "#c0aae0", tint: "#f1ecfa", icon: MousePointerClick },// freshworks purple
  bounced:      { c: "#c2410c", l: "#f0b090", tint: "#fbeae0", icon: MailX },            // burnt-orange — undeliverable
  replied:      { c: "#1288c9", l: "#93c7e6", tint: "#e4f2fb", icon: Reply },           // turquoise family
  unsubscribed: { c: "#e86f25", l: "#f4bd97", tint: "#fdf0e3", icon: UserMinus },       // casablanca orange
  failed:       { c: "#64748b", l: "#b7c0ca", tint: "#eef2f6", icon: AlertTriangle },   // neutral slate (no red)
};

// Get (or create) the Outlook folder campaign mail is filed into, matching the name the
// user picked in the composer (persisted in localStorage). Returns the folder id, or null.
async function ensureCampaignFolderId(accessToken) {
  let name = "CRM Campaigns";
  try { name = (localStorage.getItem("_crm_campaign_folder") || "CRM Campaigns").trim() || "CRM Campaigns"; } catch { /* default */ }
  const H = { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" };
  try {
    const listRes = await fetch("https://graph.microsoft.com/v1.0/me/mailFolders?$top=100&$select=id,displayName", { headers: H });
    const list = await listRes.json();
    const found = (list.value || []).find((f) => (f.displayName || "").toLowerCase() === name.toLowerCase());
    if (found) return found.id;
    const createRes = await fetch("https://graph.microsoft.com/v1.0/me/mailFolders", {
      method: "POST", headers: H, body: JSON.stringify({ displayName: name }),
    });
    const created = await createRes.json();
    return created.id || null;
  } catch { return null; }
}

// Given a snapshot of inbox messages and one campaign's recipients, classify each campaign-related
// message (read receipt / delivery receipt / genuine reply), record the receipts to the backend,
// and file ALL of them into the campaign folder so the inbox stays clean. Reused by the per-campaign
// "Check replies" and the overview "File all campaign mail" sweep. Returns per-type counts.
async function scanAndFileCampaign(accessToken, messages, recipients, campaignId, campaignSubject, canMove, folderId, processedIds) {
  const recEmails = (recipients || []).map((r) => (typeof r === "string" ? r : r.email || "").toLowerCase()).filter(Boolean);
  const recSet = new Set(recEmails);
  // Shared across campaigns in a sweep so a message that matches several campaigns (same recipients)
  // is only handled once — a message's id changes after /move, so re-touching it would 404.
  const seen = processedIds || new Set();
  const opens = new Set(), delivered = new Set(), replies = new Set(), bounces = new Set();
  const moveIds = [];
  const replyCandidates = [];
  // Strict gate: ONLY this campaign's replies and read/delivery receipts get filed — never other
  // inbox mail. A message qualifies only if its subject references THIS campaign's subject
  // (replies are "RE: <subject>", read receipts "Read: <subject>", etc.).
  const strip = (s) => (s || "").toLowerCase().replace(/^((re|fw|fwd|aw|sv|wg|read|delivered|undeliverable|gelesen|zugestellt|automatic reply|out of office)\s*:\s*)+/i, "").trim();
  const campNorm = strip(campaignSubject);
  const matchesCampaign = (subj) => {
    if (!campNorm) return false;               // no subject to match → move nothing (don't touch the inbox)
    const s = strip(subj);
    return !!s && (s === campNorm || s.includes(campNorm) || campNorm.includes(s));
  };
  const isReplyLike = (subj) => /^(re|fw|fwd|aw|sv|wg|回复|答复|antwort)\s*:/i.test((subj || "").trim());
  for (const mm of (messages || [])) {
    if (!mm.id || seen.has(mm.id)) continue;
    const from = (mm.from?.emailAddress?.address || "").toLowerCase();
    const subj = (mm.subject || "").trim();
    const isRead = /^(read:|read receipt|gelesen:|lu:|leído:|已读)/i.test(subj);
    // Bounce-back / NDR: a message from the mail system (mailer-daemon / postmaster / Exchange) OR a
    // classic failure subject. The failed recipient's address appears in the report body. These do
    // NOT carry the campaign subject reliably (e.g. Gmail's "Delivery Status Notification (Failure)")
    // so they are matched by the recipient email inside the report, NOT by matchesCampaign.
    const isNdrSender = /mailer-daemon|postmaster|microsoftexchange|maildeliverysystem|mail delivery sub/i.test(from);
    const isBounceSubj = /^(undeliverable:|undelivered mail|returned mail|mail delivery failed|failure notice|delivery has failed|delivery incomplete)/i.test(subj)
      || /delivery status notification\s*\(fail/i.test(subj)
      || (/delivery status notification/i.test(subj) && /fail|undeliver|not.*deliver|rejected/i.test((mm.bodyPreview || "")));
    if (isNdrSender || isBounceSubj) {
      const hay = (subj + " " + (mm.bodyPreview || "")).toLowerCase();
      let hit = false;
      for (const e of recEmails) if (hay.includes(e)) { bounces.add(e); hit = true; }   // this address bounced
      if (hit) { moveIds.push(mm.id); seen.add(mm.id); }
      continue;
    }
    // A genuine delivery RECEIPT (not an NDR) confirms delivery.
    const isDelivered = /^(delivered:|delivery receipt|zugestellt:|remis:)/i.test(subj);
    if (!matchesCampaign(subj)) continue;      // not related to this campaign → leave it in the inbox
    if (isRead) {
      if (recSet.has(from)) { opens.add(from); moveIds.push(mm.id); seen.add(mm.id); }
    } else if (isDelivered) {
      const hay = (subj + " " + (mm.bodyPreview || "")).toLowerCase();
      let hit = false;
      for (const e of recEmails) if (hay.includes(e)) { delivered.add(e); hit = true; }
      if (hit) { moveIds.push(mm.id); seen.add(mm.id); }
    } else if (recSet.has(from) && isReplyLike(subj)) {
      replyCandidates.push({ id: mm.id, from }); seen.add(mm.id);   // a genuine reply to this campaign
    }
  }
  // Disambiguate reply candidates by message class so receipts in any language count correctly.
  for (const c of replyCandidates) {
    let msgClass = "";
    try {
      const cr = await fetch(
        `https://graph.microsoft.com/v1.0/me/messages/${c.id}?$select=id&$expand=singleValueExtendedProperties($filter=id eq 'String 0x001A')`,
        { headers: { Authorization: `Bearer ${accessToken}` } }
      );
      if (cr.ok) { const cj = await cr.json(); msgClass = (cj.singleValueExtendedProperties?.[0]?.value || "").toUpperCase(); }
    } catch { /* class unknown → treat as reply */ }
    if (msgClass.startsWith("REPORT")) {
      if (msgClass.includes("IPNRN") || msgClass.includes("IPNNRN")) opens.add(c.from);
      else delivered.add(c.from);
    } else {
      replies.add(c.from);
    }
    moveIds.push(c.id);
  }
  const oArr = [...opens], dArr = [...delivered], rArr = [...replies], bArr = [...bounces];
  let movedCount = 0;
  if (canMove && moveIds.length && folderId) {
    for (const id of moveIds) {
      try {
        const mv = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${id}/move`, {
          method: "POST",
          headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
          body: JSON.stringify({ destinationId: folderId }),
        });
        if (mv.ok) movedCount++;
      } catch { /* skip */ }
    }
  }
  if (oArr.length || dArr.length || rArr.length || bArr.length) {
    try { await apiClient.post(`/EmailCampaign/${campaignId}/receipts`, { opens: oArr, delivered: dArr, replies: rArr, bounces: bArr }); } catch { /* best-effort */ }
  }
  return { oArr, dArr, rArr, bArr, movedCount };
}

const fmtNum = (n) => (n == null ? "0" : Number(n).toLocaleString());
const fmtDate = (v) => {
  if (!v) return "—";
  // Backend timestamps are UTC. A bare string with no timezone marker is treated as UTC (append Z),
  // then always displayed in India Standard Time regardless of the viewer's browser timezone.
  let s = v;
  if (typeof s === "string" && !/[zZ]|[+-]\d{2}:?\d{2}$/.test(s)) s = s + "Z";
  const d = new Date(s);
  return isNaN(d.getTime()) ? "—" : d.toLocaleString("en-IN", { dateStyle: "medium", timeStyle: "short", timeZone: "Asia/Kolkata" });
};
const pctv = (part, whole) => (whole > 0 ? Math.round((part / whole) * 1000) / 10 : 0);

const StatusPill = ({ status }) => {
  const s = String(status || "").toLowerCase();
  const map = {
    completed: M.opened, sent: M.opened, sending: M.sent, queued: M.unsubscribed,
    pending: M.campaigns, failed: M.failed,
  };
  const m = map[s] || M.campaigns;
  return <span className="inline-flex items-center px-2 py-0.5 rounded-full text-[11px] font-semibold" style={{ background: m.tint, color: m.c }}>{status || "—"}</span>;
};

// Freshsales-style engagement chip.
const Chip = ({ label, metric }) => {
  const m = M[metric];
  return <span className="inline-flex items-center px-2 py-0.5 rounded text-[11px] font-semibold" style={{ background: m.tint, color: m.c }}>{label}</span>;
};

const RecipientChips = ({ r }) => {
  const chips = [];
  if (r.status === "Bounced") chips.push(<Chip key="b" label="Bounced" metric="bounced" />);
  else if (r.status === "Failed") chips.push(<Chip key="f" label="Failed" metric="failed" />);
  else if (r.openCount === 0 && !r.replied) chips.push(<Chip key="s" label="Sent" metric="campaigns" />);
  if (r.openCount > 0) chips.push(<Chip key="o" label={`Opened${r.openCount > 1 ? ` ${r.openCount}×` : ""}`} metric="opened" />);
  if (r.replied) chips.push(<Chip key="r" label="Replied" metric="replied" />);
  if (r.unsubscribed) chips.push(<Chip key="u" label="Unsubscribed" metric="unsubscribed" />);
  return <div className="flex flex-wrap gap-1">{chips}</div>;
};

// KPI card — the whole card is a light tint of its metric colour, navy number (Freshsales look).
const StatTile = ({ metric, label, value, sub }) => {
  const m = M[metric] || M.campaigns;
  const Icon = m.icon;
  return (
    <div className="rounded-lg border p-4" style={{ background: m.tint, borderColor: m.l }}>
      <div className="flex items-center justify-between">
        <span className="text-[11px] font-semibold uppercase tracking-wide" style={{ color: m.c }}>{label}</span>
        <span className="w-7 h-7 rounded-full flex items-center justify-center flex-shrink-0" style={{ background: "#fff", color: m.c }}>
          <Icon className="w-[15px] h-[15px]" strokeWidth={2.2} />
        </span>
      </div>
      <p className="text-[26px] font-bold mt-1.5 tabular-nums leading-none" style={{ color: INK }}>{value}</p>
      {sub != null && <p className="text-xs mt-1.5" style={{ color: m.c }}>{sub}</p>}
    </div>
  );
};


// ── Per-recipient event timeline ──────────────────────────────────────────────
const EVENT_META = {
  Sent: { metric: "sent", label: "Sent" },
  Delivered: { metric: "sent", label: "Delivered (confirmed)" },
  Open: { metric: "opened", label: "Opened" },
  Click: { metric: "clicked", label: "Clicked a link" },
  Bounce: { metric: "bounced", label: "Bounced (invalid address)" },
  Reply: { metric: "replied", label: "Replied" },
  Unsubscribe: { metric: "unsubscribed", label: "Unsubscribed" },
  Subscribe: { metric: "opened", label: "Re-subscribed" },
  Failed: { metric: "failed", label: "Send failed" },
};

function Timeline({ events }) {
  if (!events) return <p className="text-xs py-2" style={{ color: MUTED }}>Loading activity…</p>;
  if (events.length === 0) return <p className="text-xs py-2" style={{ color: MUTED }}>No activity recorded yet.</p>;
  return (
    <ol className="relative ml-1.5 space-y-3 py-1" style={{ borderLeft: `1px solid ${BORDER}` }}>
      {events.map((e, i) => {
        const m = EVENT_META[e.type] || { metric: "campaigns", label: e.type };
        const col = M[m.metric].c;
        return (
          <li key={i} className="ml-4">
            <span className="absolute -left-[5px] w-2.5 h-2.5 rounded-full" style={{ background: col }} />
            <div className="flex flex-wrap items-baseline gap-x-2">
              <span className="text-sm font-medium" style={{ color: INK }}>{m.label}</span>
              <span className="text-xs" style={{ color: MUTED }}>{fmtDate(e.occurredAt)}</span>
              {e.type === "Open" && e.proxy && <span className="text-[11px] font-medium" style={{ color: M.unsubscribed.c }}>automated (image proxy)</span>}
            </div>
            {(e.client || e.device || e.ipAddress) && (
              <div className="text-[11px] mt-0.5" style={{ color: BODY }}>{[e.client, e.device].filter(Boolean).join(" · ")}{e.ipAddress ? ` · ${e.ipAddress}` : ""}</div>
            )}
            {e.url && <div className="text-[11px] truncate max-w-lg mt-0.5" style={{ color: M.clicked.c }} title={e.url}>{e.url}</div>}
          </li>
        );
      })}
    </ol>
  );
}

const FILTERS = [
  { key: "", label: "All" }, { key: "opened", label: "Opened" }, { key: "bounced", label: "Bounced" },
  { key: "replied", label: "Replied" }, { key: "unopened", label: "Not opened" },
  { key: "unsubscribed", label: "Unsubscribed" }, { key: "failed", label: "Failed" },
];

function Card({ children, className = "" }) {
  return <div className={`bg-white rounded-lg border ${className}`} style={{ borderColor: BORDER }}>{children}</div>;
}

function CampaignDetail({ campaignId, onBack }) {
  const { instance, accounts } = useMsal();
  const [detail, setDetail] = useState(null);
  const [recipients, setRecipients] = useState([]);
  const [filter, setFilter] = useState("");
  const [loading, setLoading] = useState(true);
  const [expandedId, setExpandedId] = useState(null);
  const [events, setEvents] = useState({});
  const [checkingReplies, setCheckingReplies] = useState(false);
  const [replyMsg, setReplyMsg] = useState("");

  const loadDetail = useCallback(() => {
    apiClient.get(`/EmailCampaign/${campaignId}`).then((r) => setDetail(r.data)).catch(() => setDetail(null));
  }, [campaignId]);
  const loadRecipients = useCallback(() => {
    setLoading(true);
    apiClient.get(`/EmailCampaign/${campaignId}/recipients`, { params: { filter, page: 1, pageSize: 500 } })
      .then((r) => setRecipients(Array.isArray(r.data?.items) ? r.data.items : []))
      .catch(() => setRecipients([])).finally(() => setLoading(false));
  }, [campaignId, filter]);

  useEffect(() => { loadDetail(); }, [loadDetail]);
  useEffect(() => { loadRecipients(); }, [loadRecipients]);

  const toggleExpand = async (r) => {
    if (expandedId === r.id) { setExpandedId(null); return; }
    setExpandedId(r.id);
    if (!events[r.id]) {
      try {
        const res = await apiClient.get(`/EmailCampaign/recipient/${r.id}/events`);
        setEvents((prev) => ({ ...prev, [r.id]: Array.isArray(res.data?.events) ? res.data.events : [] }));
      } catch { setEvents((prev) => ({ ...prev, [r.id]: [] })); }
    }
  };

  // Scans the signed-in mailbox once and derives three signals per recipient:
  //   • read receipts (subject "Read: …", sent by the recipient) → confirmed OPEN (works even
  //     when the tracking pixel was blocked by images-off clients),
  //   • delivery receipts (subject "Delivered: …", sent by the mail system) → DELIVERED,
  //   • any other message from a recipient → genuine REPLY.
  const checkReplies = async (silent = false) => {
    if (!accounts || accounts.length === 0) { if (!silent) setReplyMsg("Sign in to Outlook to check replies."); return; }
    setCheckingReplies(true); if (!silent) setReplyMsg("");
    try {
      // Prefer Mail.ReadWrite so matched messages can also be moved into the campaign folder;
      // fall back to read-only (still counts replies/receipts, just can't file them) if not consented.
      let tok, canMove = true;
      try {
        tok = await instance.acquireTokenSilent({ account: accounts[0], scopes: ["https://graph.microsoft.com/Mail.ReadWrite"] });
      } catch {
        canMove = false;
        tok = await instance.acquireTokenSilent({ account: accounts[0], scopes: ["https://graph.microsoft.com/Mail.Read"] });
      }
      const res = await fetch("https://graph.microsoft.com/v1.0/me/messages?$top=250&$select=id,from,subject,bodyPreview,receivedDateTime&$orderby=receivedDateTime desc", { headers: { Authorization: `Bearer ${tok.accessToken}` } });
      const data = await res.json();
      const folderId = canMove ? await ensureCampaignFolderId(tok.accessToken) : null;
      const { oArr, dArr, rArr, bArr, movedCount } = await scanAndFileCampaign(tok.accessToken, data.value || [], recipients, campaignId, detail?.subject, canMove, folderId);
      if (oArr.length || dArr.length || rArr.length || bArr.length) {
        const parts = [];
        if (rArr.length) parts.push(`${rArr.length} repl${rArr.length === 1 ? "y" : "ies"}`);
        if (bArr.length) parts.push(`${bArr.length} bounced`);
        if (oArr.length) parts.push(`${oArr.length} read receipt${oArr.length === 1 ? "" : "s"}`);
        if (dArr.length) parts.push(`${dArr.length} delivered`);
        setReplyMsg(`Found ${parts.join(", ")}${movedCount ? ` · moved ${movedCount} to your campaign folder` : ""}.`);
        loadDetail(); loadRecipients();
      } else if (!silent) setReplyMsg("No replies or receipts from these recipients found in your inbox.");
    } catch { if (!silent) setReplyMsg("Couldn't check inbox (Outlook sign-in / permission needed)."); }
    finally { setCheckingReplies(false); }
  };

  // Auto-file this campaign's read receipts + replies into the campaign folder (and record them)
  // as soon as it's opened — so campaign mail leaves the inbox without a manual click. Runs once
  // per campaign, after its recipients have loaded.
  const autoFiledRef = useRef(false);
  useEffect(() => { autoFiledRef.current = false; }, [campaignId]);
  useEffect(() => {
    if (autoFiledRef.current) return;
    // Wait for the campaign subject (detail) too — the filer matches replies by subject.
    if (!detail?.subject || !recipients.length || !accounts || accounts.length === 0) return;
    autoFiledRef.current = true;
    checkReplies(true);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [recipients, accounts, detail]);

  const tiles = detail ? [
    { metric: "sent", label: "Sent", value: fmtNum(detail.sentCount), sub: `of ${fmtNum(detail.totalRecipients)}` },
    { metric: "opened", label: "Opened", value: fmtNum(detail.openedCount), sub: `${detail.openRate}% open rate` },
    { metric: "bounced", label: "Bounced", value: fmtNum(detail.bouncedCount), sub: `${detail.bounceRate}% bounce rate` },
    { metric: "replied", label: "Replied", value: fmtNum(detail.repliedCount), sub: `${detail.replyRate}% reply rate` },
    { metric: "unsubscribed", label: "Unsubscribed", value: fmtNum(detail.unsubscribedCount), sub: `${detail.unsubscribeRate}%` },
    { metric: "failed", label: "Failed", value: fmtNum(detail.failedCount), sub: detail.failedCount > 0 ? "needs attention" : "none" },
  ] : [];

  return (
    <div className="space-y-5">
      <button onClick={onBack} className="inline-flex items-center gap-1.5 text-sm font-medium hover:underline" style={{ color: PRIMARY }}>
        <ArrowLeft className="w-4 h-4" /> Back to campaigns
      </button>

      {detail && (
        <Card className="p-5 sm:p-6">
          <div className="flex flex-wrap items-start justify-between gap-3">
            <div className="min-w-0">
              <h2 className="text-lg font-semibold truncate" style={{ color: INK }}>{detail.subject || "(no subject)"}</h2>
              <p className="text-sm mt-0.5" style={{ color: BODY }}>From {detail.fromEmail || "—"} · {fmtDate(detail.createdAt)}</p>
            </div>
            <div className="flex items-center gap-2">
              <button onClick={checkReplies} disabled={checkingReplies}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-semibold hover:opacity-90 transition disabled:opacity-60"
                style={{ background: M.replied.tint, color: M.replied.c }}>
                <Reply className="w-3.5 h-3.5" /> {checkingReplies ? "Checking…" : "Check replies"}
              </button>
              <StatusPill status={detail.status} />
            </div>
          </div>
          {replyMsg && <p className="text-xs mt-2" style={{ color: BODY }}>{replyMsg}</p>}
          <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3 mt-5">
            {tiles.map((t) => <StatTile key={t.metric} {...t} />)}
          </div>
        </Card>
      )}

      <div className="flex flex-wrap gap-2">
        {FILTERS.map((f) => (
          <button key={f.key} onClick={() => setFilter(f.key)}
            className="px-3.5 py-1.5 rounded-full text-xs font-medium transition"
            style={filter === f.key ? { background: PRIMARY, color: "#fff" } : { background: "#fff", color: BODY, border: `1px solid ${BORDER}` }}>
            {f.label}
          </button>
        ))}
      </div>

      <Card className="overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="text-[11px] uppercase tracking-wider" style={{ background: SURFACE, color: MUTED }}>
              <tr>
                <th className="w-8" />
                <th className="text-left font-semibold px-4 py-3">Recipient</th>
                <th className="text-left font-semibold px-4 py-3">Engagement</th>
                <th className="text-left font-semibold px-4 py-3">Last activity</th>
                <th className="text-left font-semibold px-5 py-3">Sent</th>
              </tr>
            </thead>
            <tbody className="divide-y" style={{ borderColor: BORDER }}>
              {loading ? (
                <tr><td colSpan={5} className="px-5 py-12 text-center" style={{ color: MUTED }}>Loading…</td></tr>
              ) : recipients.length === 0 ? (
                <tr><td colSpan={5} className="px-5 py-12 text-center" style={{ color: MUTED }}>No recipients in this view.</td></tr>
              ) : recipients.map((r) => {
                const last = r.lastClickedAt || r.lastOpenedAt || r.repliedAt;
                return (
                  <React.Fragment key={r.id}>
                    <tr onClick={() => toggleExpand(r)} className="cursor-pointer transition-colors hover:bg-[#f7f9fb]">
                      <td className="px-2 py-3.5" style={{ color: MUTED }}>{expandedId === r.id ? <ChevronDown className="w-4 h-4" /> : <ChevronRight className="w-4 h-4" />}</td>
                      <td className="px-4 py-3.5">
                        <div className="font-medium" style={{ color: INK }}>{r.email}</div>
                        {r.status === "Failed" && r.error && <div className="text-[11px] mt-0.5 break-words max-w-md" style={{ color: M.failed.c }} title={r.error}>{r.error}</div>}
                      </td>
                      <td className="px-4 py-3.5"><RecipientChips r={r} /></td>
                      <td className="px-4 py-3.5 text-xs" style={{ color: BODY }}>{last ? fmtDate(last) : "—"}</td>
                      <td className="px-5 py-3.5 text-xs whitespace-nowrap" style={{ color: BODY }}>{r.sentAt ? fmtDate(r.sentAt) : "—"}</td>
                    </tr>
                    {expandedId === r.id && (
                      <tr style={{ background: "#fbfcfd" }}>
                        <td /><td colSpan={4} className="px-4 pb-4 pt-1"><Timeline events={events[r.id]} /></td>
                      </tr>
                    )}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

// ── Main page ─────────────────────────────────────────────────────────────────
export default function EmailTracking() {
  const { instance, accounts } = useMsal();
  const [overview, setOverview] = useState(null);
  const [campaigns, setCampaigns] = useState([]);
  const [loading, setLoading] = useState(true);
  const [selected, setSelected] = useState(null);
  const [filingAll, setFilingAll] = useState(false);
  const [fileAllMsg, setFileAllMsg] = useState("");

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const [ov, list] = await Promise.all([apiClient.get("/EmailCampaign/overview"), apiClient.get("/EmailCampaign")]);
      setOverview(ov.data || null);
      setCampaigns(Array.isArray(list.data) ? list.data : []);
    } catch { setOverview(null); setCampaigns([]); }
    finally { setLoading(false); }
  }, []);

  useEffect(() => { load(); }, [load]);

  // Sweep the inbox once and file every campaign's read receipts + replies into the campaign
  // folder — across ALL campaigns — so no campaign mail is left sitting in the inbox.
  const fileAllCampaignMail = async () => {
    if (!accounts || accounts.length === 0) { setFileAllMsg("Sign in to Outlook first."); return; }
    if (!campaigns.length) { setFileAllMsg("No campaigns yet."); return; }
    setFilingAll(true); setFileAllMsg("");
    try {
      let tok, canMove = true;
      try { tok = await instance.acquireTokenSilent({ account: accounts[0], scopes: ["https://graph.microsoft.com/Mail.ReadWrite"] }); }
      catch { canMove = false; tok = await instance.acquireTokenSilent({ account: accounts[0], scopes: ["https://graph.microsoft.com/Mail.Read"] }); }
      const res = await fetch("https://graph.microsoft.com/v1.0/me/messages?$top=250&$select=id,from,subject,bodyPreview,receivedDateTime&$orderby=receivedDateTime desc", { headers: { Authorization: `Bearer ${tok.accessToken}` } });
      const data = await res.json();
      const messages = data.value || [];
      const folderId = canMove ? await ensureCampaignFolderId(tok.accessToken) : null;
      const processedIds = new Set();   // shared across campaigns so no message is handled twice
      let totalMoved = 0, hitCampaigns = 0;
      for (const c of campaigns) {
        try {
          const rr = await apiClient.get(`/EmailCampaign/${c.id}/recipients`, { params: { page: 1, pageSize: 1000 } });
          const recips = Array.isArray(rr.data?.items) ? rr.data.items : [];
          if (!recips.length) continue;
          const { movedCount } = await scanAndFileCampaign(tok.accessToken, messages, recips, c.id, c.subject, canMove, folderId, processedIds);
          if (movedCount) { totalMoved += movedCount; hitCampaigns++; }
        } catch { /* skip this campaign */ }
      }
      setFileAllMsg(
        !canMove ? "Grant Outlook mail permission to move messages."
          : totalMoved ? `Filed ${totalMoved} message${totalMoved === 1 ? "" : "s"} from ${hitCampaigns} campaign${hitCampaigns === 1 ? "" : "s"} into your campaign folder.`
            : "No new campaign mail found in your inbox."
      );
      load();
    } catch { setFileAllMsg("Couldn't file campaign mail (Outlook sign-in / permission needed)."); }
    finally { setFilingAll(false); }
  };

  const o = overview || {};
  const tiles = [
    { metric: "campaigns", label: "Campaigns", value: fmtNum(o.campaigns), sub: `${fmtNum(o.recipients)} recipients` },
    { metric: "sent", label: "Sent", value: fmtNum(o.sent), sub: o.failed > 0 ? `${fmtNum(o.failed)} failed` : "delivered" },
    { metric: "opened", label: "Opened", value: fmtNum(o.opened), sub: `${o.openRate || 0}% open rate` },
    { metric: "bounced", label: "Bounced", value: fmtNum(o.bounced), sub: `${o.bounceRate || 0}% bounce rate` },
    { metric: "replied", label: "Replied", value: fmtNum(o.replied), sub: `${o.replyRate || 0}% reply rate` },
    { metric: "unsubscribed", label: "Unsubscribed", value: fmtNum(o.unsubscribed), sub: `${o.unsubscribeRate || 0}%` },
  ];

  return (
    <div className="h-full w-full overflow-auto font-[poppins,sans-serif]" style={{ background: SURFACE }}>
      <div className="w-full max-w-[1500px] mx-auto p-4 sm:p-6 lg:p-8 space-y-5">
        <div className="flex items-center justify-between gap-3">
          <div>
            <h1 className="text-xl sm:text-2xl font-bold tracking-tight" style={{ color: INK }}>Email Tracking</h1>
            <p className="text-sm mt-0.5" style={{ color: BODY }}>Delivery, opens, clicks, replies and unsubscribes across your campaigns.</p>
          </div>
          <div className="flex items-center gap-2">
            {selected == null && (
              <button onClick={fileAllCampaignMail} disabled={filingAll}
                className="inline-flex items-center gap-2 px-3.5 py-2 rounded-md text-sm font-medium hover:opacity-90 transition disabled:opacity-60"
                style={{ background: M.replied.tint, color: M.replied.c }}>
                <Inbox className="w-4 h-4" /> {filingAll ? "Filing…" : "File campaign mail"}
              </button>
            )}
            <button onClick={load} className="inline-flex items-center gap-2 px-3.5 py-2 rounded-md text-sm font-medium bg-white hover:bg-[#f7f9fb] transition" style={{ color: PRIMARY, border: `1px solid ${BORDER}` }}>
              <RefreshCw className={`w-4 h-4 ${loading ? "animate-spin" : ""}`} /> Refresh
            </button>
          </div>
        </div>
        {selected == null && fileAllMsg && <p className="text-xs -mt-2" style={{ color: BODY }}>{fileAllMsg}</p>}

        {selected != null ? (
          <CampaignDetail campaignId={selected} onBack={() => { setSelected(null); load(); }} />
        ) : (
          <>
            <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3 sm:gap-4">
              {tiles.map((t) => <StatTile key={t.metric} {...t} />)}
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
              {/* Engagement funnel — insight card with a header band + tapering funnel bars */}
              <div className="lg:col-span-2 rounded-xl border overflow-hidden shadow-sm" style={{ borderColor: BORDER, background: "#fff" }}>
                <div className="px-5 py-3.5 flex items-center gap-2.5" style={{ background: "linear-gradient(180deg,#f6f9ff,#ffffff)", borderBottom: `1px solid ${BORDER}` }}>
                  <span className="w-7 h-7 rounded-lg flex items-center justify-center" style={{ background: M.sent.tint, color: M.sent.c }}><Filter className="w-[15px] h-[15px]" /></span>
                  <h3 className="text-sm font-semibold" style={{ color: INK }}>Engagement funnel</h3>
                  <span className="text-xs ml-auto" style={{ color: MUTED }}>of {fmtNum(o.sent)} delivered</span>
                </div>
                <div className="p-5 sm:p-6">
                  {o.sent > 0 ? (
                    <div className="space-y-3">
                      {[
                        { metric: "sent", label: "Delivered", count: o.sent },
                        { metric: "opened", label: "Opened", count: o.opened },
                        { metric: "bounced", label: "Bounced", count: o.bounced },
                        { metric: "replied", label: "Replied", count: o.replied },
                        { metric: "unsubscribed", label: "Unsubscribed", count: o.unsubscribed },
                      ].map((s) => {
                        const p = pctv(s.count, o.sent);
                        return (
                          <div key={s.metric}>
                            <div className="flex items-center justify-between mb-1">
                              <span className="text-xs font-medium" style={{ color: BODY }}>{s.label}</span>
                              <span className="text-xs tabular-nums" style={{ color: MUTED }}><b style={{ color: INK }}>{fmtNum(s.count)}</b> · {p}%</span>
                            </div>
                            <div className="h-3.5 rounded-full overflow-hidden" style={{ background: "#f1f4f7" }}>
                              <div className="h-full rounded-full transition-all duration-500" style={{ width: `${Math.max(2, p)}%`, background: M[s.metric].l }} />
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  ) : <div className="py-10 text-center text-sm" style={{ color: MUTED }}>No delivered emails yet.</div>}
                </div>
              </div>

              {/* Delivery health — insight card with a donut gauge */}
              <div className="rounded-xl border overflow-hidden shadow-sm" style={{ borderColor: BORDER, background: "#fff" }}>
                <div className="px-5 py-3.5 flex items-center gap-2.5" style={{ background: "linear-gradient(180deg,#f1fbf7,#ffffff)", borderBottom: `1px solid ${BORDER}` }}>
                  <span className="w-7 h-7 rounded-lg flex items-center justify-center" style={{ background: M.opened.tint, color: M.opened.c }}><Gauge className="w-[15px] h-[15px]" /></span>
                  <h3 className="text-sm font-semibold" style={{ color: INK }}>Delivery health</h3>
                </div>
                <div className="p-5 sm:p-6 flex flex-col items-center justify-center">
                  {(() => {
                    const dp = pctv(o.sent, (o.sent || 0) + (o.failed || 0));
                    const rad = 74, circ = 2 * Math.PI * rad, dash = (dp / 100) * circ;
                    return (
                      <div className="relative" style={{ width: 172, height: 172 }}>
                        <svg width="172" height="172" style={{ transform: "rotate(-90deg)" }}>
                          <circle cx="86" cy="86" r={rad} fill="none" stroke="#eef1f4" strokeWidth="15" />
                          <circle cx="86" cy="86" r={rad} fill="none" stroke={M.opened.c} strokeWidth="15" strokeLinecap="round" strokeDasharray={`${dash} ${circ}`} />
                        </svg>
                        <div className="absolute inset-0 flex flex-col items-center justify-center">
                          <span className="text-4xl font-bold tabular-nums" style={{ color: INK }}>{dp}%</span>
                          <span className="text-[11px]" style={{ color: MUTED }}>delivered</span>
                        </div>
                      </div>
                    );
                  })()}
                  <div className="flex items-center gap-6 mt-5 text-sm">
                    <span className="inline-flex items-center gap-2"><span className="w-2.5 h-2.5 rounded-full" style={{ background: M.opened.c }} /><span style={{ color: BODY }}><b style={{ color: INK }}>{fmtNum(o.sent)}</b> sent</span></span>
                    <span className="inline-flex items-center gap-2"><span className="w-2.5 h-2.5 rounded-full" style={{ background: M.failed.c }} /><span style={{ color: BODY }}><b style={{ color: INK }}>{fmtNum(o.failed)}</b> failed</span></span>
                  </div>
                </div>
              </div>
            </div>

            <Card className="overflow-hidden">
              <div className="px-5 py-3.5 flex items-center gap-2" style={{ borderBottom: `1px solid ${BORDER}` }}>
                <Inbox className="w-4 h-4" style={{ color: MUTED }} />
                <p className="text-sm font-semibold" style={{ color: INK }}>Campaigns</p>
                <span className="text-xs" style={{ color: MUTED }}>({fmtNum(campaigns.length)})</span>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead className="text-[11px] uppercase tracking-wider" style={{ background: SURFACE, color: MUTED }}>
                    <tr>
                      <th className="text-left font-semibold px-5 py-3">Campaign</th>
                      <th className="text-left font-semibold px-4 py-3">Status</th>
                      <th className="text-right font-semibold px-4 py-3">Sent</th>
                      <th className="text-right font-semibold px-4 py-3">Opened</th>
                      <th className="text-right font-semibold px-4 py-3">Bounced</th>
                      <th className="text-right font-semibold px-4 py-3">Replied</th>
                      <th className="text-right font-semibold px-5 py-3">Created</th>
                      <th className="w-8" />
                    </tr>
                  </thead>
                  <tbody className="divide-y" style={{ borderColor: BORDER }}>
                    {loading ? (
                      [...Array(4)].map((_, i) => <tr key={i}><td colSpan={8} className="px-5 py-4"><div className="h-4 rounded animate-pulse" style={{ background: "#eef1f4" }} /></td></tr>)
                    ) : campaigns.length === 0 ? (
                      <tr><td colSpan={8} className="px-5 py-16 text-center">
                        <Megaphone className="w-8 h-8 mx-auto mb-2" style={{ color: "#cfd7df" }} />
                        <p className="text-sm font-medium" style={{ color: BODY }}>No campaigns yet</p>
                        <p className="text-xs mt-1" style={{ color: MUTED }}>Use the composer's "Send tracked campaign" to see it here.</p>
                      </td></tr>
                    ) : campaigns.map((c) => (
                      <tr key={c.id} onClick={() => setSelected(c.id)} className="cursor-pointer transition-colors group hover:bg-[#f4f8ff]">
                        <td className="px-5 py-3.5 max-w-xs">
                          <div className="font-semibold truncate" style={{ color: INK }}>{c.subject || "(no subject)"}</div>
                          <div className="text-xs truncate" style={{ color: MUTED }}>{c.fromEmail || c.createdBy || "—"}</div>
                        </td>
                        <td className="px-4 py-3.5"><StatusPill status={c.status} /></td>
                        <td className="px-4 py-3.5 text-right tabular-nums" style={{ color: BODY }}>{fmtNum(c.sentCount)}<span style={{ color: MUTED }}>/{fmtNum(c.totalRecipients)}</span></td>
                        <td className="px-4 py-3.5 text-right"><span className="font-semibold tabular-nums" style={{ color: M.opened.c }}>{c.openRate}%</span> <span className="text-xs" style={{ color: MUTED }}>({fmtNum(c.openedCount)})</span></td>
                        <td className="px-4 py-3.5 text-right"><span className="font-semibold tabular-nums" style={{ color: M.bounced.c }}>{fmtNum(c.bouncedCount)}</span> <span className="text-xs" style={{ color: MUTED }}>({c.bounceRate}%)</span></td>
                        <td className="px-4 py-3.5 text-right"><span className="font-semibold tabular-nums" style={{ color: M.replied.c }}>{fmtNum(c.repliedCount)}</span></td>
                        <td className="px-5 py-3.5 text-right text-xs whitespace-nowrap" style={{ color: BODY }}>{fmtDate(c.createdAt)}</td>
                        <td className="px-2 py-3.5" style={{ color: "#cfd7df" }}><ChevronRight className="w-4 h-4" /></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Card>
          </>
        )}
      </div>
    </div>
  );
}
