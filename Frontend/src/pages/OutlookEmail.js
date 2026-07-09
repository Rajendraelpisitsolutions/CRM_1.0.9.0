import React, { useState, useEffect, useRef, useMemo, useCallback } from "react";
import {
  Inbox, Send, Trash2, Edit2, Reply, ReplyAll, Forward, ArrowLeft,
  Mail, MailOpen, Paperclip, Plus, Pencil, FileText, Archive,
  AlertTriangle, Flag, Tag, Bell, Settings, Search, Sun, Moon,
  ChevronDown, ChevronRight, ChevronLeft, Star, Clock, Calendar,
  Users, CheckSquare, Download, Filter,
  X, Printer, RefreshCw, Link, UserPlus, Eye,
  Check, Pin, Activity, User,
  MessageSquare, FolderOpen, AlertCircle,
  AtSign, ArrowUpDown,
  MapPin, Video, Link2, Smile, Copy, Globe, Mic, ExternalLink, Bookmark, Lock, Menu,
  CheckCircle2, Image as ImageIcon,
} from "lucide-react";
import { FiTrash2 } from "react-icons/fi";
import apiClient from "../api/client";
import { useMsal } from "@azure/msal-react";
import DOMPurify from "dompurify";

// ─── Helpers ─────────────────────────────────────────────────────────────────
const TRANSPARENT_GIF = "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7";

// A template packs a body, a footer and an optional company logo (a data URL) into the
// single Template.Body column using sentinels, so no schema change is needed. It splits
// back into { body, footer, logo } on load (the subject lives in Template.Subject).
// `text` is kept as an alias of `footer` for older callers.
const LOGO_OPEN = "[[LOGO]]";
const LOGO_CLOSE = "[[/LOGO]]";
const FOOTER_OPEN = "[[FOOTER]]";
const FOOTER_CLOSE = "[[/FOOTER]]";
const ATTACH_OPEN = "[[ATTACH]]";
const ATTACH_CLOSE = "[[/ATTACH]]";
function buildTemplateBody(body, footer, logo, attachments) {
  let out = "";
  if (logo) out += `${LOGO_OPEN}${logo}${LOGO_CLOSE}`;
  if (attachments && attachments.length) out += `${ATTACH_OPEN}${JSON.stringify(attachments)}${ATTACH_CLOSE}`;
  out += `${FOOTER_OPEN}${footer || ""}${FOOTER_CLOSE}`;
  out += (body || "");
  return out;
}
function parseTemplateBody(raw) {
  let b = raw || "";
  let logo = "", attachments = [];
  if (b.startsWith(LOGO_OPEN)) {
    const end = b.indexOf(LOGO_CLOSE);
    if (end !== -1) { logo = b.slice(LOGO_OPEN.length, end); b = b.slice(end + LOGO_CLOSE.length); }
  }
  if (b.startsWith(ATTACH_OPEN)) {
    const end = b.indexOf(ATTACH_CLOSE);
    if (end !== -1) {
      try { attachments = JSON.parse(b.slice(ATTACH_OPEN.length, end)) || []; } catch { attachments = []; }
      b = b.slice(end + ATTACH_CLOSE.length);
    }
  }
  if (b.startsWith("\n")) b = b.slice(1);
  if (b.startsWith(FOOTER_OPEN)) {
    // New format: [[FOOTER]]footer[[/FOOTER]]body
    const end = b.indexOf(FOOTER_CLOSE);
    if (end !== -1) {
      const footer = b.slice(FOOTER_OPEN.length, end);
      let body = b.slice(end + FOOTER_CLOSE.length);
      if (body.startsWith("\n")) body = body.slice(1);
      return { body, footer, logo, attachments, text: footer };
    }
  }
  // Legacy format: the remaining text was the footer (no separate body).
  return { body: "", footer: b, logo, attachments, text: b };
}

// Ensures an Outlook mail folder with the given name exists (creating it once if not) and
// returns its id. Campaign copies are filed here (instead of Sent Items) so a big send
// doesn't clog the mailbox. Returns null on any failure, so the caller falls back cleanly.
async function ensureMailFolderId(accessToken, name) {
  try {
    const listRes = await window.fetch(
      "https://graph.microsoft.com/v1.0/me/mailFolders?$top=100&$select=id,displayName",
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    const list = await listRes.json();
    const found = (list.value || []).find(
      (f) => (f.displayName || "").toLowerCase() === name.toLowerCase()
    );
    if (found) return found.id;
    const createRes = await window.fetch("https://graph.microsoft.com/v1.0/me/mailFolders", {
      method: "POST",
      headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
      body: JSON.stringify({ displayName: name }),
    });
    const created = await createRes.json();
    return created.id || null;
  } catch {
    return null;
  }
}

function replaceCidImages(html, attachments = []) {
  if (!html) return html;
  let result = html;
  attachments.forEach(att => {
    if (!att.contentId) return;
    const cleanCid = att.contentId.replace(/[<>]/g, "");
    if (!att.contentBytes || !att.contentType) {
      [
        `cid:${cleanCid}`, `cid:${cleanCid}:1`, `cid:${cleanCid}:2`,
        `cid:${cleanCid}@`, att.contentId, `cid:${att.contentId}`,
      ].forEach(p => { result = result.split(p).join(TRANSPARENT_GIF); });
      return;
    }
    const base64Src = `data:${att.contentType};base64,${att.contentBytes}`;
    [
      `cid:${cleanCid}`, `cid:${cleanCid}:1`, `cid:${cleanCid}:2`,
      `cid:${cleanCid}@`, att.contentId, `cid:${att.contentId}`,
      `<${cleanCid}>`, `"cid:${cleanCid}"`, `'cid:${cleanCid}'`, cleanCid,
    ].forEach(p => { result = result.split(p).join(base64Src); });
  });
  result = result.replace(/src\s*=\s*["']cid:[^"']*["']/gi, `src="${TRANSPARENT_GIF}"`);
  return result;
}

function stripHtml(html) {
  if (!html) return "";
  return html.replace(/<[^>]*>/g, "").replace(/&nbsp;/g, " ").trim();
}

function splitEmailBody(html) {
  if (!html) return [html];
  const marked = html
    .replace(/<div[^>]*id\s*=\s*["']divRplyFwdMsg["'][^>]*>/gi, "<!--SPLIT--><div id=\"divRplyFwdMsg\">")
    .replace(/<hr\s*[^>]*>/gi, (m) => `<!--SPLIT-->${m}`);
  const parts = marked.split("<!--SPLIT-->").map(s => s.trim()).filter(Boolean);
  if (parts.length <= 1) return [html];
  return parts;
}

function fmtDateTime(d) {
  if (!d) return "";
  const date = new Date(d), now = new Date();
  const diff = now - date, h = Math.floor(diff / 3600000);
  if (h < 24) return date.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });
  if (h < 48) return "Yesterday";
  if (h < 168) return date.toLocaleDateString("en-US", { weekday: "short" });
  return date.toLocaleDateString("en-US", { month: "short", day: "numeric" });
}

function fmtTimeAgo(d) {
  if (!d) return "";
  const date = new Date(d), now = new Date();
  const diff = now - date;
  if (diff < 0) {
    const mins = Math.round(-diff / 60000);
    if (mins < 60) return `in ${mins}m`;
    const hrs = Math.round(-diff / 3600000);
    return hrs < 24 ? `in ${hrs}h` : `in ${Math.round(-diff / 86400000)}d`;
  }
  const mins = Math.floor(diff / 60000);
  if (mins < 1) return "Just now";
  if (mins < 60) return `${mins}m ago`;
  const hrs = Math.floor(diff / 3600000);
  if (hrs < 24) return `${hrs}h ago`;
  const days = Math.floor(diff / 86400000);
  return days < 7 ? `${days}d ago` : date.toLocaleDateString("en-US", { month: "short", day: "numeric" });
}

function fmtFileSize(bytes) {
  if (!bytes) return "";
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1048576) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1048576).toFixed(1)} MB`;
}

// ─── Calendar helpers ─────────────────────────────────────────────────────────
function calDateKey(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
function calAddDays(d, n) {
  const r = new Date(d); r.setDate(r.getDate() + n); return r;
}
function getMonthGrid(year, month) {
  const firstDay = new Date(year, month, 1);
  const lastDay = new Date(year, month + 1, 0);
  const startDow = firstDay.getDay();
  const days = [];
  for (let i = startDow - 1; i >= 0; i--) days.push({ date: new Date(year, month, -i), curr: false });
  for (let d = 1; d <= lastDay.getDate(); d++) days.push({ date: new Date(year, month, d), curr: true });
  while (days.length < 42) days.push({ date: new Date(year, month + 1, days.length - startDow - lastDay.getDate() + 1), curr: false });
  return days;
}

function getInitials(name) {
  if (!name) return "?";
  const parts = name.trim().split(/\s+/);
  return parts.length >= 2
    ? (parts[0][0] + parts[parts.length - 1][0]).toUpperCase()
    : name[0].toUpperCase();
}

const AVATAR_COLORS = [
  "bg-blue-500", "bg-purple-500", "bg-green-500", "bg-rose-500",
  "bg-orange-500", "bg-teal-500", "bg-indigo-500", "bg-pink-500",
];
function getAvatarColor(name) {
  if (!name) return AVATAR_COLORS[0];
  let h = 0;
  for (let i = 0; i < name.length; i++) h = (h * 31 + name.charCodeAt(i)) % AVATAR_COLORS.length;
  return AVATAR_COLORS[h];
}

const CATEGORY_COLORS = {
  red: { bg: "bg-red-100", dot: "bg-red-500", text: "text-red-700", label: "Red category" },
  orange: { bg: "bg-orange-100", dot: "bg-orange-500", text: "text-orange-700", label: "Orange category" },
  yellow: { bg: "bg-yellow-100", dot: "bg-yellow-400", text: "text-yellow-700", label: "Yellow category" },
  green: { bg: "bg-green-100", dot: "bg-green-500", text: "text-green-700", label: "Green category" },
  blue: { bg: "bg-blue-100", dot: "bg-blue-500", text: "text-blue-700", label: "Blue category" },
  purple: { bg: "bg-purple-100", dot: "bg-purple-500", text: "text-purple-700", label: "Purple category" },
};

// ─── Isolated email body renderer ────────────────────────────────────────────
function EmailBodyFrame({ html, isDark }) {
  const frameId = useRef("ef-" + Math.random().toString(36).slice(2));
  const [frameHeight, setFrameHeight] = useState(120);

  useEffect(() => {
    const id = frameId.current;
    const handler = (e) => {
      if (e.data?.type === "email-frame-resize" && e.data?.id === id) {
        setFrameHeight(Math.max(120, (e.data.height || 0) + 16));
      }
    };
    window.addEventListener("message", handler);
    return () => window.removeEventListener("message", handler);
  }, []);

  const sanitized = DOMPurify.sanitize(html || "", {
    ADD_DATA_URI_TAGS: ["img"],
    FORBID_TAGS: ["script", "noscript", "object", "embed", "applet", "form"],
    FORBID_ATTR: [
      "onerror", "onload", "onclick", "ondblclick", "onmousedown", "onmouseup", "onmouseover",
      "onmouseout", "onmousemove", "onkeydown", "onkeyup", "onkeypress", "onchange", "onsubmit",
      "onreset", "onselect", "onblur", "onfocus", "onabort", "onscroll", "onresize", "onunload",
      "onbeforeunload", "ondragstart", "ondrop", "oncontextmenu", "onwheel", "oninput", "oninvalid",
    ],
  });

  const id = frameId.current;
  const darkCss = isDark
    ? "html,body,body *{background-color:transparent!important;color:#ffffff!important;}"
    : "";

  const resizeScript = `
(function(){
  var ID="${id}";
  function send(){
    var h=document.body?document.body.scrollHeight:0;
    window.parent.postMessage({type:"email-frame-resize",id:ID,height:h},"*");
  }
  send();
  window.addEventListener("load",send);
  document.querySelectorAll("img").forEach(function(img){
    if(!img.complete){img.addEventListener("load",send);img.addEventListener("error",send);}
  });
  new MutationObserver(send).observe(document.documentElement,{subtree:true,childList:true,attributes:true});
  [200,600,1200].forEach(function(t){setTimeout(send,t);});
})();`;

  const srcdoc = [
    "<!DOCTYPE html><html><head>",
    '<meta charset="UTF-8">',
    '<meta name="viewport" content="width=device-width,initial-scale=1">',
    "<style>",
    "html,body{margin:0;padding:8px;font-size:15px;line-height:1.55;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;word-wrap:break-word;overflow-x:hidden;}",
    "img{max-width:100%;height:auto;display:inline-block;}",
    "table{border-collapse:collapse;max-width:100%;}",
    "a{color:inherit;word-break:break-all;}",
    darkCss,
    "</style></head><body>",
    sanitized,
    "<script>", resizeScript, "<\/script>",
    "</body></html>",
  ].join("");

  return (
    <iframe
      srcDoc={srcdoc}
      title="email-body"
      sandbox="allow-scripts"
      style={{ width: "100%", border: "none", height: frameHeight + "px", display: "block" }}
    />
  );
}

// ─── File type icon helper ────────────────────────────────────────────────────
function FileIcon({ name, size = 14 }) {
  const ext = (name || "").split(".").pop().toLowerCase();
  const colorMap = {
    pdf: "text-red-500", doc: "text-blue-600", docx: "text-blue-600",
    xls: "text-green-600", xlsx: "text-green-600", ppt: "text-orange-500",
    pptx: "text-orange-500", jpg: "text-purple-500", jpeg: "text-purple-500",
    png: "text-purple-500", gif: "text-pink-500", zip: "text-yellow-600",
    txt: "text-gray-500",
  };
  return <FileText size={size} className={colorMap[ext] || "text-gray-400"} />;
}

// ─── App-rail item ────────────────────────────────────────────────────────────
function RailItem({ icon: Icon, label, active, onClick, badge, isDark }) {
  return (
    <button
      onClick={onClick}
      title={label}
      className={`relative flex flex-col items-center justify-center w-12 h-12 rounded-xl transition-all ${active
        ? "bg-blue-500 text-white"
        : isDark
          ? "text-gray-300 hover:bg-white/10 hover:text-white"
          : "text-gray-700 hover:bg-gray-100 hover:text-gray-900"
        }`}
    >
      <Icon size={20} />
      {badge > 0 && (
        <span className="absolute top-1 right-1 w-4 h-4 bg-red-500 text-white text-[9px] rounded-full flex items-center justify-center font-bold">
          {badge > 9 ? "9+" : badge}
        </span>
      )}
    </button>
  );
}

// ─── Folder row in sidebar ────────────────────────────────────────────────────
function FolderRow({ icon: Icon, label, active, count, pinned, onPin, onClick, isDark }) {
  return (
    <button
      onClick={onClick}
      className={`group w-full flex items-center gap-2.5 px-3 py-2 rounded-lg text-left text-sm transition ${active
        ? isDark ? "bg-blue-500 text-white" : "bg-blue-100 text-blue-700"
        : isDark ? "text-gray-300 hover:bg-gray-700" : "text-gray-700 hover:bg-gray-100"
        }`}
    >
      <Icon size={16} className="flex-shrink-0" />
      <span className="flex-1 truncate">{label}</span>
      {count > 0 && (
        <span className={`text-[11px] px-1.5 py-0.5 rounded-full font-medium ${active ? "bg-white/20 text-white" : isDark ? "bg-gray-600 text-gray-200" : "bg-blue-100 text-blue-700"}`}>
          {count}
        </span>
      )}
      {onPin && (
        <Pin
          size={12}
          onClick={e => { e.stopPropagation(); onPin(); }}
          className={`hidden group-hover:block flex-shrink-0 ${pinned ? "text-blue-500" : "text-gray-400"}`}
        />
      )}
    </button>
  );
}

// ─── Main Component ───────────────────────────────────────────────────────────
function OutlookEmail({ onToast }) {
  const { instance, accounts } = useMsal();

  // ── Auth ──────────────────────────────────────────────────────────────────
  const [accessToken, setAccessToken] = useState("");

  // ── App Shell ─────────────────────────────────────────────────────────────
  const [activeApp, setActiveApp] = useState("Mail");
  const [isDark, setIsDark] = useState(false);

  // ── Mail Sidebar ──────────────────────────────────────────────────────────
  const [activeFolder, setActiveFolder] = useState("Inbox");
  const [expandedSections, setExpandedSections] = useState({ favorites: true, folders: true, groups: false });
  const [pinnedFolders, setPinnedFolders] = useState(["Inbox", "Sent Items"]);

  // ── Campaign folder (shared with the composer via localStorage) ─────────────
  // The Outlook folder tracked-campaign copies (and their replies) are filed into.
  const [campaignFolder, setCampaignFolder] = useState(() => {
    try { return localStorage.getItem("_crm_campaign_folder") || "CRM Campaigns"; } catch { return "CRM Campaigns"; }
  });
  const [outlookFolders, setOutlookFolders] = useState([]);
  const [campaignFolderSaved, setCampaignFolderSaved] = useState(false);
  const [editCampaignFolder, setEditCampaignFolder] = useState(false);
  const persistCampaignFolder = (name) => {
    setCampaignFolder(name);
    try { localStorage.setItem("_crm_campaign_folder", (name || "").trim() || "CRM Campaigns"); } catch { /* ignore */ }
    setCampaignFolderSaved(true);
  };
  // Pull the user's existing Outlook folder names on demand (for the picker's suggestions).
  const loadOutlookFolders = useCallback(() => {
    if (!accessToken || outlookFolders.length) return;
    window.fetch("https://graph.microsoft.com/v1.0/me/mailFolders?$top=100&$select=displayName", { headers: { Authorization: `Bearer ${accessToken}` } })
      .then((r) => r.json())
      .then((d) => setOutlookFolders((d.value || []).map((f) => f.displayName).filter(Boolean)))
      .catch(() => {});
  }, [accessToken, outlookFolders.length]);

  // ── Email Data ────────────────────────────────────────────────────────────
  const [inboxEmails, setInboxEmails] = useState([]);
  const [sentMails, setSentMails] = useState([]);
  const [trashMails, setTrashMails] = useState([]);
  const [draftMails, setDraftMails] = useState([]);
  const [junkMails, setJunkMails] = useState([]);
  const [archiveMails, setArchiveMails] = useState([]);
  const [campaignMails, setCampaignMails] = useState([]);
  const [apiTemplates, setApiTemplates] = useState([]);

  // ── Email List UI ─────────────────────────────────────────────────────────
  const [emailSearch, setEmailSearch] = useState("");
  const [inboxTab, setInboxTab] = useState("Focused");
  const [sortBy, setSortBy] = useState("date");
  const [filterBy, setFilterBy] = useState("all");
  const [showSortMenu, setShowSortMenu] = useState(false);
  const [showFilterMenu, setShowFilterMenu] = useState(false);

  // ── Responsive / Mobile ───────────────────────────────────────────────────
  const [showMobileSidebar, setShowMobileSidebar] = useState(false);
  const [mobileView, setMobileView] = useState("list"); // "list" | "email"

  // ── Email Actions ─────────────────────────────────────────────────────────
  const [openedMailId, setOpenedMailId] = useState(null);
  const [selectedMailIds, setSelectedMailIds] = useState([]);
  // Outlook-style reading view: collapse the folder pane while a mail is open so
  // the reading pane gets more room. A toggle in the app rail overrides it.
  const [showFolderPane, setShowFolderPane] = useState(true);
  useEffect(() => { setShowFolderPane(!openedMailId); }, [openedMailId]);
  const [flaggedIds, setFlaggedIds] = useState({});
  const [readOverrides, setReadOverrides] = useState({});
  const [mailCategories, setMailCategories] = useState({});
  const [showCategoryMenu, setShowCategoryMenu] = useState(null);
  const [showSnoozeMenu, setShowSnoozeMenu] = useState(null);

  // ── Delete ────────────────────────────────────────────────────────────────
  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false);
  const [deleteCandidateIds, setDeleteCandidateIds] = useState([]);

  // ── Reading Pane ──────────────────────────────────────────────────────────
  const [toExpanded, setToExpanded] = useState(false);
  const [ccExpanded, setCcExpanded] = useState(false);
  const [showCrmPanel, setShowCrmPanel] = useState(false);
  const [crmTab, setCrmTab] = useState("actions");

  // ── Thread / Conversation ─────────────────────────────────────────────────
  const [threadMessages, setThreadMessages] = useState([]);
  const [threadExpandedIds, setThreadExpandedIds] = useState(new Set());
  const [threadBodies, setThreadBodies] = useState({});

  // ── Compose ───────────────────────────────────────────────────────────────
  const [showCompose, setShowCompose] = useState(false);
  // True while the compose's Templates & Tags panel is open — used to hide the
  // mail list so the bigger template panel + compose can take over that space.
  const [composeTemplatesOpen, setComposeTemplatesOpen] = useState(false);
  useEffect(() => { if (!showCompose) setComposeTemplatesOpen(false); }, [showCompose]);
  const [replyData, setReplyData] = useState(null);

  // ── Inline reply ─────────────────────────────────────────────────────────
  const [inlineReply, setInlineReply] = useState(null);
  const [inlineToInput, setInlineToInput] = useState("");
  const [inlineBody, setInlineBody] = useState("");
  const [inlineSending, setInlineSending] = useState(false);
  const [inlineError, setInlineError] = useState("");
  const [inlineSendDropdown, setInlineSendDropdown] = useState(false);
  const [scheduleModal, setScheduleModal] = useState(false);
  const [scheduleDateTime, setScheduleDateTime] = useState("");
  const [schedulingInProgress, setSchedulingInProgress] = useState(false);

  // ── Side Panels ───────────────────────────────────────────────────────────
  const [showNotifications, setShowNotifications] = useState(false);
  const [showSettings, setShowSettings] = useState(false);

  // ── Templates ─────────────────────────────────────────────────────────────
  const [openedTemplateIdx, setOpenedTemplateIdx] = useState(null);
  const [showAddTemplate, setShowAddTemplate] = useState(false);
  // A template is "assigned" to an email (stored in CreatedBy). That email is the
  // key used to auto-load a user's template on compose.
  const loggedInEmail = (typeof window !== "undefined" ? sessionStorage.getItem("userEmail") : "") || "";
  const [newTemplate, setNewTemplate] = useState({ name: "", subject: "", body: "", footer: "", assignedEmail: "", logo: "", attachments: [] });
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [editTemplateIdx, setEditTemplateIdx] = useState(null);
  const [editTemplate, setEditTemplate] = useState({ subject: "", body: "", footer: "", assignedEmail: "", logo: "", attachments: [] });

  // A user only sees / manages templates assigned to them (CreatedBy == their email);
  // legacy templates with no assignment are owned if their footer carries the email.
  const ownsTemplate = (tpl) => {
    const cb = (tpl.createdBy ?? tpl.CreatedBy ?? "").toLowerCase().trim();
    const em = loggedInEmail.toLowerCase().trim();
    if (cb) return cb === em;
    const hay = `${tpl.name ?? ""} ${tpl.body ?? tpl.Body ?? ""}`.toLowerCase();
    return !!em && hay.includes(em);
  };
  const myTemplates = Array.isArray(apiTemplates) ? apiTemplates.filter(ownsTemplate) : [];

  // ── Message-box popup (replaces toast for created/updated/deleted) ─────────
  const [popup, setPopup] = useState(null); // { message, type: "success" | "error" }
  const showPopup = (message, type = "success") => {
    setPopup({ message, type });
    setTimeout(() => setPopup(prev => (prev && prev.message === message ? null : prev)), 2200);
  };

  // ── Delete-template confirm message box (replaces window.confirm) ──────────
  const [templateToDelete, setTemplateToDelete] = useState(null); // the tpl object
  const confirmDeleteTemplate = async () => {
    const tpl = templateToDelete;
    if (!tpl) return;
    setTemplateToDelete(null);
    try {
      await apiClient.delete(`/Template/${tpl.templateId ?? tpl.TemplateId}`);
      apiClient.get("/Template").then(r => setApiTemplates(Array.isArray(r.data) ? r.data : [])).catch(() => { });
      setOpenedTemplateIdx(null);
      showPopup("Template deleted");
    } catch {
      showPopup("Could not delete template", "error");
    }
  };

  // ── Settings ──────────────────────────────────────────────────────────────
  const [settingsTab, setSettingsTab] = useState("account");
  const [signature, setSignature] = useState("");
  const [autoReply, setAutoReply] = useState({ enabled: false, message: "" });

  // ── Copilot ───────────────────────────────────────────────────────────────
  const [copilotInput, setCopilotInput] = useState("");
  const [copilotHistory, setCopilotHistory] = useState([
    { role: "assistant", text: "Hi! I can summarize emails, draft replies, extract tasks, and more. Select an email to get started." },
  ]);

  // ── Calendar ──────────────────────────────────────────────────────────────
  const _today = new Date();
  const [calView, setCalView] = useState("Month");
  const [calDate, setCalDate] = useState(new Date(_today.getFullYear(), _today.getMonth(), 1));
  const [miniCalDate, setMiniCalDate] = useState(new Date(_today.getFullYear(), _today.getMonth(), 1));
  const [selectedCalDay, setSelectedCalDay] = useState(new Date(_today));
  const [showEventForm, setShowEventForm] = useState(false);
  const [eventFormDate, setEventFormDate] = useState(null);
  const [newEventTitle, setNewEventTitle] = useState("");
  const [newEventStart, setNewEventStart] = useState("12:00");
  const [newEventEnd, setNewEventEnd] = useState("12:30");
  const [newEventColor, setNewEventColor] = useState("blue");
  const [, setNewEventAttendees] = useState("");
  const [newEventLocation, setNewEventLocation] = useState("");
  const [outlookPeople, setOutlookPeople] = useState([]);
  const [evSelectedAttendees, setEvSelectedAttendees] = useState([]);
  const [evAttendeeSearch, setEvAttendeeSearch] = useState("");
  const [showEvAttendeePicker, setShowEvAttendeePicker] = useState(false);
  const [newEventDescription, setNewEventDescription] = useState("");
  const [newEventAllDay, setNewEventAllDay] = useState(false);
  const [newEventOnline, setNewEventOnline] = useState(false);
  const [newEventAgenda, setNewEventAgenda] = useState("");
  const [evBusyStatus, setEvBusyStatus] = useState("Busy");
  const [showEvBusyDd, setShowEvBusyDd] = useState(false);
  const [evReminder, setEvReminder] = useState("15 minutes before");
  const [showEvReminderDd, setShowEvReminderDd] = useState(false);
  const [evCategory, setEvCategory] = useState(null);
  const [showEvCategoryDd, setShowEvCategoryDd] = useState(false);
  const [evPrivacy, setEvPrivacy] = useState("Not private");
  const [showEvPrivacyDd, setShowEvPrivacyDd] = useState(false);
  const [calEvents, setCalEvents] = useState([]);
  const [calLoading, setCalLoading] = useState(false);
  const [selectedEvent, setSelectedEvent] = useState(null);

  // ── Notifications ─────────────────────────────────────────────────────────
  const [notifications, setNotifications] = useState([]);
  const [notifLoading, setNotifLoading] = useState(false);
  const [notifRefreshKey, setNotifRefreshKey] = useState(0);
  const unreadNotifCount = notifications.filter(n => !n.read).length;

  // ─── Reset on folder change ───────────────────────────────────────────────
  useEffect(() => {
    setEmailSearch("");
    setOpenedMailId(null);
    setSelectedMailIds([]);
    setShowAddTemplate(false);
    setShowCrmPanel(false);
    setShowCategoryMenu(null);
    setShowSnoozeMenu(null);
  }, [activeFolder]);

  useEffect(() => {
    setToExpanded(false);
    setCcExpanded(false);
    setShowCategoryMenu(null);
  }, [openedMailId]);

  // ─── Close menus on outside click ─────────────────────────────────────────
  useEffect(() => {
    const handler = () => {
      setShowSortMenu(false);
      setShowFilterMenu(false);
      setShowCategoryMenu(null);
      setShowSnoozeMenu(null);
    };
    document.addEventListener("click", handler);
    return () => document.removeEventListener("click", handler);
  }, []);

  // ─── Auth token ───────────────────────────────────────────────────────────
  const accountId = accounts[0]?.localAccountId ?? null;
  useEffect(() => {
    const getToken = async () => {
      if (!accountId) { setAccessToken(""); return; }
      try {
        const r = await instance.acquireTokenSilent({
          account: accounts[0],
          scopes: [
            "https://graph.microsoft.com/User.Read",
            "https://graph.microsoft.com/Mail.Read",
            "https://graph.microsoft.com/Mail.ReadWrite",
            "https://graph.microsoft.com/Mail.Send",
            "https://graph.microsoft.com/Contacts.Read",
            "https://graph.microsoft.com/Calendars.ReadWrite",
          ],
        });
        setAccessToken(r.accessToken);
      } catch (e) {
        if (e?.name === "InteractionRequiredAuthError") {
          setAccessToken("");
        } else if (e?.errorCode !== "interaction_in_progress") {
          console.warn("[OutlookEmail] Silent token failed:", e?.errorCode || e?.name);
          setAccessToken("");
        }
      }
    };
    getToken();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [accountId]);

  // ─── Notifications fetch ──────────────────────────────────────────────────
  useEffect(() => {
    const fetchNotifs = async () => {
      if (!accessToken) return;
      setNotifLoading(true);
      const notifs = [];
      const h = { Authorization: `Bearer ${accessToken}` };

      try {
        const r = await fetch(
          "https://graph.microsoft.com/v1.0/me/messages" +
          "?$filter=isRead eq false&$top=3&$orderby=receivedDateTime desc" +
          "&$select=id,subject,bodyPreview,from,receivedDateTime,isRead",
          { headers: h }
        );
        if (r.ok) {
          const d = await r.json();
          (d.value || []).forEach(msg => {
            const sender = msg.from?.emailAddress?.name || msg.from?.emailAddress?.address || "Unknown";
            notifs.push({
              id: `email-${msg.id}`, type: "email",
              title: `New email from ${sender}`,
              body: (msg.bodyPreview || msg.subject || "").slice(0, 90),
              time: fmtTimeAgo(msg.receivedDateTime),
              read: !!msg.isRead, graphId: msg.id,
            });
          });
        }
      } catch { }

      try {
        const now = new Date();
        const end24h = new Date(now.getTime() + 24 * 60 * 60 * 1000);
        const r = await fetch(
          `https://graph.microsoft.com/v1.0/me/calendarView` +
          `?startDateTime=${now.toISOString()}&endDateTime=${end24h.toISOString()}` +
          `&$select=id,subject,start,end&$top=3&$orderby=start/dateTime`,
          { headers: h }
        );
        if (r.ok) {
          const d = await r.json();
          (d.value || []).forEach(ev => {
            const startDt = new Date(ev.start.dateTime.endsWith("Z") ? ev.start.dateTime : ev.start.dateTime + "Z");
            const diffMin = Math.round((startDt - now) / 60000);
            notifs.push({
              id: `meeting-${ev.id}`, type: "meeting",
              title: ev.subject || "Upcoming meeting",
              body: diffMin <= 0 ? "Happening now" : diffMin < 60 ? `Starts in ${diffMin} min` : `Starts in ${Math.round(diffMin / 60)}h`,
              time: fmtTimeAgo(ev.start.dateTime.endsWith("Z") ? ev.start.dateTime : ev.start.dateTime + "Z"),
              read: diffMin > 30, graphId: ev.id,
            });
          });
        }
      } catch { }

      setNotifications(notifs);
      setNotifLoading(false);
    };
    fetchNotifs();
  }, [accessToken, notifRefreshKey]);

  // ─── Outlook contacts fetch ───────────────────────────────────────────────
  useEffect(() => {
    if (!accessToken || !accountId) return;
    const fetchOutlookPeople = async () => {
      try {
        const contactsRes = await fetch(
          "https://graph.microsoft.com/v1.0/me/contacts?$top=200&$select=displayName,emailAddresses,jobTitle,companyName&$orderby=displayName",
          { headers: { Authorization: `Bearer ${accessToken}` } }
        );
        if (!contactsRes.ok) return;
        const d = await contactsRes.json();
        const seen = new Set();
        const merged = [];
        for (const c of (d.value || [])) {
          const email = c.emailAddresses?.[0]?.address;
          if (!email || seen.has(email.toLowerCase())) continue;
          seen.add(email.toLowerCase());
          merged.push({ name: c.displayName || email, email, title: c.jobTitle || c.companyName || "", source: "contact" });
        }
        setOutlookPeople(merged);
      } catch { }
    };
    fetchOutlookPeople();
  }, [accessToken, accountId]);

  // ─── Inbox fetch ─────────────────────────────────────────────────────────
  useEffect(() => {
    const load = async () => {
      if (!accessToken || activeFolder !== "Inbox") return;
      try {
        const res = await fetch(
          "https://graph.microsoft.com/v1.0/me/messages?" +
          "$top=50&$orderby=receivedDateTime desc" +
          "&$select=id,subject,bodyPreview,from,sender,toRecipients,ccRecipients,receivedDateTime,hasAttachments,isRead,importance",
          { headers: { Authorization: `Bearer ${accessToken}` } }
        );
        if (!res.ok) { setInboxEmails([]); return; }
        const data = await res.json();
        setInboxEmails((data.value || []).map(msg => {
          const senderName = msg.from?.emailAddress?.name || msg.sender?.emailAddress?.name || msg.from?.emailAddress?.address || "";
          return {
            id: msg.id, subject: msg.subject,
            body: "", bodyPreview: msg.bodyPreview || "", attachments: [],
            hasAttachments: msg.hasAttachments, bodyLoaded: false,
            toEmail: msg.toRecipients?.map(r => r.emailAddress.address).join(", ") || "",
            ccEmail: msg.ccRecipients?.map(r => r.emailAddress.address).join(", ") || "",
            receivedDateTime: msg.receivedDateTime, sender: senderName, from: senderName,
            senderEmail: msg.from?.emailAddress?.address || msg.sender?.emailAddress?.address || "",
            isRead: msg.isRead, importance: msg.importance,
          };
        }));
      } catch (e) { console.error("Inbox fetch error:", e); setInboxEmails([]); }
    };
    load();
  }, [accessToken, activeFolder]);

  // ─── Sent fetch ───────────────────────────────────────────────────────────
  useEffect(() => {
    const load = async () => {
      if (activeFolder !== "Sent" || !accessToken) return;
      try {
        const res = await fetch(
          "https://graph.microsoft.com/v1.0/me/mailfolders/sentitems/messages?" +
          "$top=50&$orderby=sentDateTime desc" +
          "&$select=id,subject,bodyPreview,from,toRecipients,ccRecipients,sentDateTime,hasAttachments",
          { headers: { Authorization: `Bearer ${accessToken}` } }
        );
        if (!res.ok) { setSentMails([]); return; }
        const data = await res.json();
        setSentMails((data.value || []).map(msg => {
          const senderName = msg.from?.emailAddress?.name || msg.from?.emailAddress?.address || "";
          return {
            id: msg.id, subject: msg.subject,
            body: "", bodyPreview: msg.bodyPreview || "", attachments: [],
            hasAttachments: msg.hasAttachments, bodyLoaded: false,
            toEmail: msg.toRecipients?.map(r => r.emailAddress.address).join(", ") || "",
            ccEmail: msg.ccRecipients?.map(r => r.emailAddress.address).join(", ") || "",
            sentDateTime: msg.sentDateTime, sender: senderName, from: senderName,
            senderEmail: msg.from?.emailAddress?.address || "", isRead: true,
          };
        }));
      } catch (e) { console.error("Sent fetch error:", e); setSentMails([]); }
    };
    load();
  }, [activeFolder, accessToken]);

  // ─── Trash fetch ──────────────────────────────────────────────────────────
  useEffect(() => {
    const load = async () => {
      if (activeFolder !== "Deleted" || !accessToken) return;
      try {
        const res = await fetch(
          "https://graph.microsoft.com/v1.0/me/mailfolders/deleteditems/messages?" +
          "$top=50&$orderby=receivedDateTime desc" +
          "&$select=id,subject,bodyPreview,from,sender,toRecipients,ccRecipients,receivedDateTime,hasAttachments",
          { headers: { Authorization: `Bearer ${accessToken}` } }
        );
        if (!res.ok) { setTrashMails([]); return; }
        const data = await res.json();
        setTrashMails((data.value || []).map(msg => {
          const senderName = msg.from?.emailAddress?.name || msg.sender?.emailAddress?.name || msg.from?.emailAddress?.address || "";
          return {
            id: msg.id, subject: msg.subject,
            body: "", bodyPreview: msg.bodyPreview || "", attachments: [],
            hasAttachments: msg.hasAttachments, bodyLoaded: false,
            toEmail: msg.toRecipients?.map(r => r.emailAddress.address).join(", ") || "",
            ccEmail: msg.ccRecipients?.map(r => r.emailAddress.address).join(", ") || "",
            receivedDateTime: msg.receivedDateTime, sender: senderName, from: senderName,
            senderEmail: msg.from?.emailAddress?.address || msg.sender?.emailAddress?.address || "",
          };
        }));
      } catch (e) { console.error("Trash fetch error:", e); setTrashMails([]); }
    };
    load();
  }, [activeFolder, accessToken]);

  // ─── Drafts fetch ─────────────────────────────────────────────────────────
  useEffect(() => {
    const load = async () => {
      if (activeFolder !== "Drafts" || !accessToken) return;
      const res = await fetch(
        "https://graph.microsoft.com/v1.0/me/mailfolders/drafts/messages?" +
        "$top=50&$orderby=lastModifiedDateTime desc" +
        "&$select=id,subject,body,bodyPreview,from,toRecipients,ccRecipients,lastModifiedDateTime",
        { headers: { Authorization: `Bearer ${accessToken}` } }
      );
      if (!res.ok) { setDraftMails([]); return; }
      const data = await res.json();
      setDraftMails((data.value || []).map(msg => ({
        id: msg.id, subject: msg.subject, body: msg.body?.content || "",
        attachments: [],
        toEmail: msg.toRecipients?.map(r => r.emailAddress.address).join(", ") || "",
        ccEmail: msg.ccRecipients?.map(r => r.emailAddress.address).join(", ") || "",
        receivedDateTime: msg.lastModifiedDateTime,
        sender: msg.from?.emailAddress?.name || "(Draft)",
        from: msg.from?.emailAddress?.name || "(Draft)",
        senderEmail: msg.from?.emailAddress?.address || "",
        isDraft: true,
      })));
    };
    load();
  }, [activeFolder, accessToken]);

  // ─── Junk fetch ───────────────────────────────────────────────────────────
  useEffect(() => {
    const load = async () => {
      if (activeFolder !== "Junk" || !accessToken) return;
      const res = await fetch(
        "https://graph.microsoft.com/v1.0/me/mailfolders/junkemail/messages?" +
        "$top=50&$orderby=receivedDateTime desc" +
        "&$select=id,subject,body,bodyPreview,from,sender,toRecipients,ccRecipients,receivedDateTime,hasAttachments&$expand=attachments",
        { headers: { Authorization: `Bearer ${accessToken}` } }
      );
      if (!res.ok) { setJunkMails([]); return; }
      const data = await res.json();
      setJunkMails((data.value || []).map(msg => ({
        id: msg.id, subject: msg.subject, body: msg.body?.content || "",
        attachments: msg.attachments || [],
        toEmail: msg.toRecipients?.map(r => r.emailAddress.address).join(", ") || "",
        ccEmail: msg.ccRecipients?.map(r => r.emailAddress.address).join(", ") || "",
        receivedDateTime: msg.receivedDateTime,
        sender: msg.from?.emailAddress?.name || msg.sender?.emailAddress?.name || msg.from?.emailAddress?.address || "",
        from: msg.from?.emailAddress?.name || msg.sender?.emailAddress?.name || "",
        senderEmail: msg.from?.emailAddress?.address || msg.sender?.emailAddress?.address || "",
      })));
    };
    load();
  }, [activeFolder, accessToken]);

  // ─── Archive fetch ────────────────────────────────────────────────────────
  useEffect(() => {
    const load = async () => {
      if (activeFolder !== "Archive" || !accessToken) return;
      const res = await fetch(
        "https://graph.microsoft.com/v1.0/me/mailfolders/archive/messages?" +
        "$top=50&$orderby=receivedDateTime desc" +
        "&$select=id,subject,body,bodyPreview,from,sender,toRecipients,ccRecipients,receivedDateTime,hasAttachments&$expand=attachments",
        { headers: { Authorization: `Bearer ${accessToken}` } }
      );
      if (!res.ok) { setArchiveMails([]); return; }
      const data = await res.json();
      setArchiveMails((data.value || []).map(msg => ({
        id: msg.id, subject: msg.subject, body: msg.body?.content || "",
        attachments: msg.attachments || [],
        toEmail: msg.toRecipients?.map(r => r.emailAddress.address).join(", ") || "",
        ccEmail: msg.ccRecipients?.map(r => r.emailAddress.address).join(", ") || "",
        receivedDateTime: msg.receivedDateTime,
        sender: msg.from?.emailAddress?.name || msg.sender?.emailAddress?.name || msg.from?.emailAddress?.address || "",
        from: msg.from?.emailAddress?.name || msg.sender?.emailAddress?.name || "",
        senderEmail: msg.from?.emailAddress?.address || msg.sender?.emailAddress?.address || "",
      })));
    };
    load();
  }, [activeFolder, accessToken]);

  // ─── Campaign folder fetch ────────────────────────────────────────────────
  // Loads the messages inside the user's chosen campaign folder (sent copies + moved replies)
  // by resolving the folder's id from its display name, then listing its messages.
  useEffect(() => {
    const load = async () => {
      if (activeFolder !== "Campaign" || !accessToken) return;
      try {
        const name = (campaignFolder || "").trim() || "CRM Campaigns";
        const fRes = await fetch(
          "https://graph.microsoft.com/v1.0/me/mailFolders?$top=100&$select=id,displayName",
          { headers: { Authorization: `Bearer ${accessToken}` } }
        );
        const fData = await fRes.json();
        const folder = (fData.value || []).find(f => (f.displayName || "").toLowerCase() === name.toLowerCase());
        if (!folder) { setCampaignMails([]); return; }
        const res = await fetch(
          `https://graph.microsoft.com/v1.0/me/mailFolders/${folder.id}/messages?` +
          "$top=50&$orderby=receivedDateTime desc" +
          "&$select=id,subject,bodyPreview,from,sender,toRecipients,ccRecipients,receivedDateTime,sentDateTime,hasAttachments,isRead,isDraft",
          { headers: { Authorization: `Bearer ${accessToken}` } }
        );
        if (!res.ok) { setCampaignMails([]); return; }
        const data = await res.json();
        const myEmail = (accounts?.[0]?.username || sessionStorage.getItem("userEmail") || "").toLowerCase();
        setCampaignMails((data.value || []).map(msg => {
          const fromAddr = (msg.from?.emailAddress?.address || msg.sender?.emailAddress?.address || "").toLowerCase();
          const senderName = msg.from?.emailAddress?.name || msg.sender?.emailAddress?.name || msg.from?.emailAddress?.address || "";
          const toEmail = msg.toRecipients?.map(r => r.emailAddress.address).join(", ") || "";
          // These are mostly *sent* copies (from = me, or blank on a moved draft) — show who the
          // mail went TO rather than a blank "Unknown" (the reading pane already has its own "To:"
          // line, so no prefix here). Genuine replies (from someone else) keep the sender's name.
          const outgoing = !fromAddr || fromAddr === myEmail;
          const display = outgoing ? (toEmail || senderName || "Me") : (senderName || fromAddr || toEmail);
          return {
            id: msg.id, subject: msg.subject,
            body: "", bodyPreview: msg.bodyPreview || "", attachments: [],
            hasAttachments: msg.hasAttachments, bodyLoaded: false,
            toEmail,
            ccEmail: msg.ccRecipients?.map(r => r.emailAddress.address).join(", ") || "",
            receivedDateTime: msg.receivedDateTime || msg.sentDateTime, sentDateTime: msg.sentDateTime,
            sender: display, from: display,
            senderEmail: msg.from?.emailAddress?.address || msg.sender?.emailAddress?.address || "",
            isRead: msg.isRead, isDraft: msg.isDraft,
          };
        }));
      } catch (e) { console.error("Campaign folder fetch error:", e); setCampaignMails([]); }
    };
    load();
  }, [activeFolder, accessToken, campaignFolder]);

  // ─── Templates fetch ──────────────────────────────────────────────────────
  useEffect(() => {
    if (activeFolder === "Templates") {
      apiClient.get("/Template")
        .then(r => setApiTemplates(Array.isArray(r.data) ? r.data : []))
        .catch(() => setApiTemplates([]));
    }
  }, [activeFolder]);

  // ─── Calendar events fetch ────────────────────────────────────────────────
  useEffect(() => {
    const fetchCalEvents = async () => {
      if (!accessToken || activeApp !== "Calendar") return;
      setCalLoading(true);
      try {
        const y = calDate.getFullYear(), mo = calDate.getMonth();
        const pad2 = n => String(n).padStart(2, "0");
        const fmtLocal = d =>
          `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}T${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`;
        const startISO = fmtLocal(new Date(y, mo, 1, 0, 0, 0));
        const endISO = fmtLocal(new Date(y, mo + 1, 0, 23, 59, 59));
        const res = await fetch(
          `https://graph.microsoft.com/v1.0/me/calendarView` +
          `?startDateTime=${encodeURIComponent(startISO)}&endDateTime=${encodeURIComponent(endISO)}` +
          `&$select=id,subject,start,end,color,categories&$top=200`,
          { headers: { Authorization: `Bearer ${accessToken}` } }
        );
        if (!res.ok) { setCalLoading(false); return; }
        const data = await res.json();
        const COLOR_MAP = {
          blue: "blue", green: "green", red: "red", orange: "orange",
          purple: "purple", teal: "teal", pink: "red", yellow: "orange",
          grape: "purple", cyan: "teal",
        };
        const mapped = (data.value || []).map(ev => {
          const toLocal = s => new Date(s.endsWith("Z") ? s : s + "Z");
          const st = toLocal(ev.start.dateTime);
          const en = toLocal(ev.end.dateTime);
          return {
            id: ev.id, graphId: ev.id,
            title: ev.subject || "(No title)",
            date: calDateKey(st),
            start: `${String(st.getHours()).padStart(2, "0")}:${String(st.getMinutes()).padStart(2, "0")}`,
            end: `${String(en.getHours()).padStart(2, "0")}:${String(en.getMinutes()).padStart(2, "0")}`,
            color: COLOR_MAP[ev.color] || (ev.categories?.length ? "purple" : "blue"),
            _sortKey: st.getTime(),
          };
        });
        mapped.sort((a, b) => a._sortKey - b._sortKey);
        setCalEvents(mapped);
      } catch (e) { console.error("Calendar fetch error:", e); }
      setCalLoading(false);
    };
    fetchCalEvents();
  }, [accessToken, activeApp, calDate]);

  // ─── Reset thread state when switching emails ─────────────────────────────
  useEffect(() => {
    setThreadMessages([]);
    setThreadExpandedIds(new Set());
    setThreadBodies({});
  }, [openedMailId]);

  // ─── Lazy-load full body + attachments ───────────────────────────────────
  useEffect(() => {
    const load = async () => {
      if (!accessToken || !openedMailId) return;
      const folderMap = { Inbox: inboxEmails, Sent: sentMails, Deleted: trashMails, Drafts: draftMails, Junk: junkMails, Archive: archiveMails, Campaign: campaignMails };
      const arr = folderMap[activeFolder] || [];
      const idx = arr.findIndex(m => m.id === openedMailId);
      if (idx === -1) return;
      const msg = arr[idx];
      if (msg.bodyLoaded) {
        if (msg.conversationId && threadMessages.length === 0) fetchThread(msg.conversationId, openedMailId);
        return;
      }
      try {
        const r = await fetch(
          `https://graph.microsoft.com/v1.0/me/messages/${msg.id}?$expand=attachments&$select=id,body,attachments,conversationId`,
          { headers: { Authorization: `Bearer ${accessToken}` } }
        );
        if (!r.ok) return;
        const data = await r.json();
        const htmlBody = data.body?.content || "";
        const attachments = data.attachments || [];
        const body = replaceCidImages(htmlBody, attachments);
        const convId = data.conversationId || "";
        const updatedMsg = { ...msg, body, attachments, bodyLoaded: true, conversationId: convId };
        const setters = { Inbox: setInboxEmails, Sent: setSentMails, Deleted: setTrashMails, Drafts: setDraftMails, Junk: setJunkMails, Archive: setArchiveMails, Campaign: setCampaignMails };
        const setter = setters[activeFolder];
        if (setter) setter(prev => { const c = [...prev]; c[idx] = updatedMsg; return c; });
        if (convId) fetchThread(convId, openedMailId);
      } catch (e) { console.error("Body load error:", e); }
    };

    const fetchThread = async (convId, currentId) => {
      try {
        const r = await fetch(
          `https://graph.microsoft.com/v1.0/me/messages` +
          `?$filter=conversationId eq '${encodeURIComponent(convId).replace(/%27/g, "'")}'` +
          `&$count=true&$select=id,subject,from,sender,toRecipients,ccRecipients,receivedDateTime,bodyPreview&$top=50`,
          { headers: { Authorization: `Bearer ${accessToken}`, "ConsistencyLevel": "eventual" } }
        );
        if (!r.ok) return;
        const data = await r.json();
        const msgs = (data.value || []).map(m => ({
          id: m.id, subject: m.subject,
          sender: m.from?.emailAddress?.name || m.sender?.emailAddress?.name || m.from?.emailAddress?.address || "",
          senderEmail: m.from?.emailAddress?.address || "",
          toEmail: (m.toRecipients || []).map(x => x.emailAddress.address).join(", "),
          ccEmail: (m.ccRecipients || []).map(x => x.emailAddress.address).join(", "),
          receivedDateTime: m.receivedDateTime, bodyPreview: m.bodyPreview || "",
        })).sort((a, b) => new Date(a.receivedDateTime) - new Date(b.receivedDateTime));
        if (msgs.length > 1) {
          setThreadMessages(msgs);
          setThreadExpandedIds(new Set([currentId]));
        }
      } catch { }
    };

    load();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [openedMailId, accessToken]);

  const loadThreadBody = async (msgId) => {
    if (threadBodies[msgId]) return;
    try {
      const r = await fetch(
        `https://graph.microsoft.com/v1.0/me/messages/${msgId}?$select=id,body`,
        { headers: { Authorization: `Bearer ${accessToken}` } }
      );
      if (!r.ok) return;
      const data = await r.json();
      const html = replaceCidImages(data.body?.content || "", []);
      setThreadBodies(prev => ({ ...prev, [msgId]: html }));
    } catch { }
  };

  const toggleThreadMsg = (msgId) => {
    setThreadExpandedIds(prev => {
      const next = new Set(prev);
      if (next.has(msgId)) { next.delete(msgId); }
      else { next.add(msgId); loadThreadBody(msgId); }
      return next;
    });
  };

  // ─── Computed ─────────────────────────────────────────────────────────────
  const getCurrentMails = () => {
    const map = { Inbox: inboxEmails, Sent: sentMails, Deleted: trashMails, Drafts: draftMails, Junk: junkMails, Archive: archiveMails, Campaign: campaignMails };
    return map[activeFolder] || [];
  };

  const openedMail = getCurrentMails().find(m => m.id === openedMailId) || null;

  const getFilteredMails = () => {
    let mails = getCurrentMails();
    const q = emailSearch.trim().toLowerCase();
    if (q) {
      mails = mails.filter(m =>
        (m.subject || "").toLowerCase().includes(q) ||
        (m.sender || m.from || "").toLowerCase().includes(q) ||
        (m.senderEmail || "").toLowerCase().includes(q) ||
        stripHtml(m.body || "").toLowerCase().includes(q)
      );
    }
    if (filterBy === "unread") mails = mails.filter(m => {
      const isRead = readOverrides[m.id] !== undefined ? readOverrides[m.id] : m.isRead;
      return !isRead;
    });
    if (filterBy === "flagged") mails = mails.filter(m => flaggedIds[m.id]);
    if (filterBy === "attachments") mails = mails.filter(m => m.attachments?.length > 0);
    if (sortBy === "sender") mails = [...mails].sort((a, b) => (a.sender || "").localeCompare(b.sender || ""));
    if (sortBy === "subject") mails = [...mails].sort((a, b) => (a.subject || "").localeCompare(b.subject || ""));
    return mails;
  };

  const folderCounts = {
    Inbox: inboxEmails.filter(m => !m.isRead).length,
    Drafts: draftMails.length,
    Junk: junkMails.length,
  };

  // ─── Handlers ─────────────────────────────────────────────────────────────
  const confirmDelete = ids => {
    const arr = Array.isArray(ids) ? ids.filter(Boolean) : [ids];
    if (!arr.length) return;
    setDeleteCandidateIds(arr);
    setShowDeleteConfirm(true);
  };

  const performDelete = async () => {
    const ids = deleteCandidateIds;
    if (!ids.length) { setShowDeleteConfirm(false); return; }
    if (accessToken) {
      await Promise.all(ids.map(id =>
        fetch(`https://graph.microsoft.com/v1.0/me/messages/${id}`, {
          method: "DELETE", headers: { Authorization: `Bearer ${accessToken}` },
        }).catch(() => { })
      ));
    }
    const filter = prev => prev.filter(m => !ids.includes(m.id));
    setInboxEmails(filter); setSentMails(filter); setTrashMails(filter);
    setDraftMails(filter); setJunkMails(filter); setArchiveMails(filter); setCampaignMails(filter);
    setSelectedMailIds(p => p.filter(id => !ids.includes(id)));
    setOpenedMailId(null);
    setDeleteCandidateIds([]);
    setShowDeleteConfirm(false);
    if (onToast) onToast(`${ids.length} message(s) deleted`, "success");
  };

  const archiveMail = async mailId => {
    if (!accessToken) return;
    try {
      await fetch(`https://graph.microsoft.com/v1.0/me/messages/${mailId}/move`, {
        method: "POST",
        headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
        body: JSON.stringify({ destinationId: "archive" }),
      });
      setInboxEmails(p => p.filter(m => m.id !== mailId));
      setOpenedMailId(null);
      if (onToast) onToast("Email archived", "success");
    } catch { }
  };

  const toggleFlag = mailId => setFlaggedIds(p => ({ ...p, [mailId]: !p[mailId] }));

  const markAsRead = async mail => {
    if (!mail) return;
    const cur = readOverrides[mail.id] !== undefined ? readOverrides[mail.id] : mail.isRead;
    if (cur) return;
    setReadOverrides(p => ({ ...p, [mail.id]: true }));
    if (accessToken) {
      fetch(`https://graph.microsoft.com/v1.0/me/messages/${mail.id}`, {
        method: "PATCH",
        headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
        body: JSON.stringify({ isRead: true }),
      }).catch(() => { });
    }
  };

  const toggleRead = async mail => {
    if (!mail) return;
    const cur = readOverrides[mail.id] !== undefined ? readOverrides[mail.id] : mail.isRead;
    setReadOverrides(p => ({ ...p, [mail.id]: !cur }));
    if (accessToken) {
      fetch(`https://graph.microsoft.com/v1.0/me/messages/${mail.id}`, {
        method: "PATCH",
        headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
        body: JSON.stringify({ isRead: !cur }),
      }).catch(() => { });
    }
  };

  const setCategory = (mailId, color) => {
    setMailCategories(p => ({ ...p, [mailId]: color === p[mailId] ? null : color }));
    setShowCategoryMenu(null);
  };

  const openInlineReply = (type) => {
    if (!openedMail) return;
    let toEmail = "";
    let ccEmail = "";
    if (type === "reply") {
      toEmail = openedMail.senderEmail || openedMail.sender || "";
    } else if (type === "replyAll") {
      const toList = Array.from(new Set([
        openedMail.senderEmail,
        ...(openedMail.toEmail || "").split(/[,;]/).map(e => e.trim()).filter(Boolean),
      ]));
      toEmail = toList.join(", ");
      ccEmail = openedMail.ccEmail || "";
    }
    const subject = type === "forward"
      ? (openedMail.subject?.startsWith("Fwd:") ? openedMail.subject : `Fwd: ${openedMail.subject || ""}`)
      : (openedMail.subject?.startsWith("Re:") ? openedMail.subject : `Re: ${openedMail.subject || ""}`);
    setInlineReply({ type, subject, ccEmail, originalBody: openedMail.body });
    setInlineToInput(toEmail);
    setInlineBody("");
    setInlineError("");
  };

  const handleReply = () => openInlineReply("reply");
  const handleReplyAll = () => openInlineReply("replyAll");
  const handleForward = () => openInlineReply("forward");

  const handleInlineSend = async () => {
    const toEmails = inlineToInput.split(/[,;]/).map(e => e.trim()).filter(Boolean);
    if (!toEmails.length) { setInlineError("Please enter at least one recipient."); return; }
    setInlineSending(true);
    setInlineError("");
    try {
      const userHtml = inlineBody.replace(/\n/g, "<br>").replace(/ {2}/g, "&nbsp;&nbsp;");
      const quoteBlock = inlineReply.originalBody
        ? `<br><br><div style="border-left:2px solid #e5e7eb;padding-left:12px;margin-top:8px;color:#6b7280;font-size:13px;">${inlineReply.originalBody}</div>`
        : "";
      const ccEmails = (inlineReply.ccEmail || "").split(/[,;]/).map(e => e.trim()).filter(Boolean);
      const payload = {
        message: {
          subject: inlineReply.subject,
          body: { contentType: "HTML", content: userHtml + quoteBlock },
          toRecipients: toEmails.map(a => ({ emailAddress: { address: a } })),
          ccRecipients: ccEmails.map(a => ({ emailAddress: { address: a } })),
        },
      };
      const res = await window.fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
        method: "POST",
        headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      if (res.ok || res.status === 202) {
        setInlineReply(null);
        setInlineBody("");
        setInlineToInput("");
        if (onToast) onToast("Reply sent!", "success");
      } else {
        const err = await res.json().catch(() => ({}));
        setInlineError(err?.error?.message || "Failed to send.");
      }
    } catch (e) {
      setInlineError(e.message || "Failed to send.");
    } finally {
      setInlineSending(false);
    }
  };

  const buildInlinePayload = () => {
    const userHtml = inlineBody.replace(/\n/g, "<br>").replace(/ {2}/g, "&nbsp;&nbsp;");
    const quoteBlock = inlineReply?.originalBody
      ? `<br><br><div style="border-left:2px solid #e5e7eb;padding-left:12px;margin-top:8px;color:#6b7280;font-size:13px;">${inlineReply.originalBody}</div>`
      : "";
    return { html: userHtml + quoteBlock, subject: inlineReply?.subject || "" };
  };

  const handleScheduleSend = async () => {
    if (!scheduleDateTime) return;
    const scheduledAt = new Date(scheduleDateTime);
    const delay = scheduledAt.getTime() - Date.now();
    if (delay <= 0) { setInlineError("Please select a future date and time."); return; }
    const toEmails = inlineToInput.split(/[,;]/).map(e => e.trim()).filter(Boolean);
    if (!toEmails.length) { setInlineError("Please enter at least one recipient."); return; }
    setSchedulingInProgress(true);
    try {
      const { html, subject } = buildInlinePayload();
      const ccEmails = (inlineReply?.ccEmail || "").split(/[,;]/).map(e => e.trim()).filter(Boolean);
      const draftRes = await window.fetch("https://graph.microsoft.com/v1.0/me/messages", {
        method: "POST",
        headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
        body: JSON.stringify({
          subject,
          body: { contentType: "HTML", content: html },
          toRecipients: toEmails.map(a => ({ emailAddress: { address: a } })),
          ccRecipients: ccEmails.map(a => ({ emailAddress: { address: a } })),
        }),
      });
      if (draftRes.ok) {
        const draft = await draftRes.json();
        setTimeout(async () => {
          await window.fetch(`https://graph.microsoft.com/v1.0/me/messages/${draft.id}/send`, {
            method: "POST", headers: { Authorization: `Bearer ${accessToken}` },
          }).catch(() => { });
        }, delay);
        setScheduleModal(false);
        setInlineReply(null); setInlineBody(""); setInlineToInput("");
        if (onToast) onToast(`Email scheduled for ${scheduledAt.toLocaleString()}`, "success");
      } else {
        const err = await draftRes.json().catch(() => ({}));
        setInlineError(err?.error?.message || "Failed to schedule email.");
      }
    } catch (e) {
      setInlineError(e.message || "Failed to schedule.");
    } finally {
      setSchedulingInProgress(false);
    }
  };

  const handleMailMerge = async () => {
    const toEmails = inlineToInput.split(/[,;]/).map(e => e.trim()).filter(Boolean);
    if (!toEmails.length) { setInlineError("Please enter at least one recipient."); return; }
    setInlineSending(true); setInlineError("");
    const { html, subject } = buildInlinePayload();
    let sent = 0;
    for (const addr of toEmails) {
      try {
        const res = await window.fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
          method: "POST",
          headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
          body: JSON.stringify({
            message: {
              subject, body: { contentType: "HTML", content: html },
              toRecipients: [{ emailAddress: { address: addr } }],
            }
          }),
        });
        if (res.ok || res.status === 202) sent++;
      } catch { }
    }
    setInlineSending(false);
    setInlineReply(null); setInlineBody(""); setInlineToInput("");
    if (onToast) onToast(`Mail merge sent to ${sent} of ${toEmails.length} recipient${toEmails.length > 1 ? "s" : ""}.`, "success");
  };

  const handleMailSent = mail => {
    const newMail = {
      id: Date.now().toString(), subject: mail.subject, body: mail.body,
      toEmail: mail.to?.join(", ") || "", ccEmail: mail.cc?.join(", ") || "",
      sentDateTime: new Date().toISOString(), sender: accounts[0]?.name || "Me",
      from: accounts[0]?.name || "Me", isRead: true,
    };
    setSentMails(p => [newMail, ...p]);
  };

  // Reads picked files into base64 attachment objects for a template (create/edit).
  // They're persisted inside the template, so total size is capped.
  const addTemplateAttachments = async (fileList, setter, current) => {
    const files = Array.from(fileList || []);
    if (!files.length) return;
    const readOne = (file) => new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = () => resolve({ name: file.name, contentType: file.type || "application/octet-stream", contentBytes: (String(reader.result || "").split(",")[1] || ""), size: file.size });
      reader.onerror = () => resolve(null);
      reader.readAsDataURL(file);
    });
    const read = (await Promise.all(files.map(readOne))).filter(Boolean);
    const existingTotal = (current || []).reduce((s, a) => s + (a.size || 0), 0);
    const addTotal = read.reduce((s, a) => s + (a.size || 0), 0);
    if (existingTotal + addTotal > 4 * 1024 * 1024) { showPopup("Attachments too large (max ~4 MB total per template)", "error"); return; }
    setter((p) => ({ ...p, attachments: [...(p.attachments || []), ...read] }));
  };

  const handleAddTemplate = async e => {
    e.preventDefault();
    setIsSubmitting(true);
    try {
      const name = (newTemplate.subject || newTemplate.body || newTemplate.footer || "")
        .trim().slice(0, 40) || "Template";
      // CreatedBy = the logged-in user's email: each user owns only their own templates.
      // Subject is stored on its own; body + footer + logo are packed into Body.
      const r = await apiClient.post("/Template", {
        name,
        subject: newTemplate.subject || "",
        body: buildTemplateBody(newTemplate.body, newTemplate.footer, newTemplate.logo, newTemplate.attachments),
        createdBy: loggedInEmail,
      });
      if (r.status >= 200 && r.status < 300) {
        setShowAddTemplate(false);
        setNewTemplate({ name: "", subject: "", body: "", footer: "", assignedEmail: "", logo: "", attachments: [] });
        apiClient.get("/Template").then(r => setApiTemplates(Array.isArray(r.data) ? r.data : [])).catch(() => { });
        showPopup("Template created");
      }
    } catch (err) {
      const status = err?.response?.status;
      const tooBig = status === 413 || status === 400;
      const msg = tooBig
        ? "Template too large to save — the attachment exceeds the server upload limit. Use a smaller file."
        : (err?.response?.data?.title || err?.response?.data?.message ||
           (typeof err?.response?.data === "string" ? err.response.data : "") ||
           err?.message || "Could not create template");
      showPopup(String(msg).slice(0, 220), "error");
    }
    setIsSubmitting(false);
  };

  const toggleSelectMail = id => setSelectedMailIds(p => p.includes(id) ? p.filter(i => i !== id) : [...p, id]);
  const selectAll = () => {
    const all = getFilteredMails().map(m => m.id);
    const allSel = all.every(id => selectedMailIds.includes(id));
    setSelectedMailIds(allSel ? [] : all);
  };

  const copilotSend = () => {
    if (!copilotInput.trim()) return;
    const userMsg = { role: "user", text: copilotInput };
    let reply = "";
    const q = copilotInput.toLowerCase();
    if (q.includes("summarize") || q.includes("summary")) {
      reply = openedMail
        ? `Summary of "${openedMail.subject}":\n\nThis email is from ${openedMail.sender} and discusses: ${stripHtml(openedMail.body).slice(0, 200)}...`
        : "Please select an email first so I can summarize it.";
    } else if (q.includes("reply") || q.includes("draft")) {
      reply = openedMail
        ? `Suggested reply to "${openedMail.subject}":\n\nThank you for your email. I have reviewed the information you provided and will respond in detail shortly.\n\nBest regards`
        : "Please select an email first so I can draft a reply.";
    } else if (q.includes("task") || q.includes("action")) {
      reply = openedMail
        ? `Extracted action items from "${openedMail.subject}":\n\n• Review and respond to this email\n• Follow up by end of week\n• Schedule meeting if needed`
        : "Please select an email first to extract tasks.";
    } else {
      reply = "I can help you summarize emails, draft replies, extract action items, and more. Try asking me to 'summarize this email' or 'draft a reply'.";
    }
    setCopilotHistory(p => [...p, userMsg, { role: "assistant", text: reply }]);
    setCopilotInput("");
  };

  // ─── Notification handlers ────────────────────────────────────────────────
  const markAllNotifsRead = () => setNotifications(p => p.map(n => ({ ...n, read: true })));
  const markNotifRead = id => setNotifications(p => p.map(n => n.id === id ? { ...n, read: true } : n));

  // ─── Calendar API handlers ────────────────────────────────────────────────
  const handleSaveEvent = async () => {
    if (!newEventTitle.trim()) return;
    const date = eventFormDate || selectedCalDay;
    const pad = n => String(n).padStart(2, "0");
    const [sh, sm] = newEventStart.split(":").map(Number);
    const [eh, em] = newEventEnd.split(":").map(Number);
    const dateStr = `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
    const tempId = `local-${Date.now()}`;
    const newEv = {
      id: tempId, graphId: null,
      title: newEventTitle.trim(),
      date: calDateKey(date),
      start: newEventStart, end: newEventEnd,
      color: newEventColor,
    };
    setCalEvents(p => [...p, newEv]);
    setNewEventTitle(""); setNewEventColor("blue");
    setNewEventAttendees(""); setNewEventLocation(""); setNewEventDescription("");
    setNewEventAllDay(false); setNewEventOnline(false); setNewEventAgenda("");
    setEvBusyStatus("Busy"); setEvReminder("15 minutes before");
    setEvCategory(null); setEvPrivacy("Not private");
    setEvSelectedAttendees([]); setEvAttendeeSearch(""); setShowEvAttendeePicker(false);
    setShowEventForm(false);
    if (accessToken) {
      try {
        const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
        const showAsMap = { "Free": "free", "Working elsewhere": "workingElsewhere", "Tentative": "tentative", "Busy": "busy", "Out of office": "oof" };
        const reminderMap = { "Don't remind me": -1, "At time of event": 0, "5 minutes before": 5, "15 minutes before": 15, "30 minutes before": 30, "1 hour before": 60, "2 hours before": 120, "12 hours before": 720, "1 day before": 1440, "1 week before": 10080 };
        const reminderMins = reminderMap[evReminder] ?? 15;
        const body = {
          subject: newEv.title,
          start: { dateTime: newEventAllDay ? `${dateStr}T00:00:00` : `${dateStr}T${pad(sh)}:${pad(sm)}:00`, timeZone: tz },
          end: { dateTime: newEventAllDay ? `${dateStr}T23:59:00` : `${dateStr}T${pad(eh)}:${pad(em)}:00`, timeZone: tz },
          isOnlineMeeting: newEventOnline,
          showAs: showAsMap[evBusyStatus] || "busy",
          sensitivity: evPrivacy === "Private" ? "private" : "normal",
          isReminderOn: reminderMins >= 0,
          ...(reminderMins >= 0 ? { reminderMinutesBeforeStart: reminderMins } : {}),
          ...(evCategory ? { categories: [evCategory] } : {}),
          ...(newEventOnline ? { onlineMeetingProvider: "teamsForBusiness" } : {}),
          location: { displayName: newEventLocation || (newEventOnline ? "Microsoft Teams" : "") },
          ...(newEventDescription ? { body: { contentType: "text", content: newEventDescription } } : {}),
          ...(evSelectedAttendees.length ? { attendees: evSelectedAttendees.map(a => ({ emailAddress: { address: a.email, name: a.name }, type: "required" })) } : {}),
        };
        const r = await fetch("https://graph.microsoft.com/v1.0/me/events", {
          method: "POST",
          headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
          body: JSON.stringify(body),
        });
        if (r.ok) {
          const created = await r.json();
          setCalEvents(p => p.map(e => e.id === tempId ? { ...e, id: created.id, graphId: created.id } : e));
        }
      } catch (e) { console.error("Create event error:", e); }
    }
  };

  const handleDeleteEvent = async ev => {
    setCalEvents(p => p.filter(e => e.id !== ev.id));
    setSelectedEvent(null);
    const gId = ev.graphId || ev.id;
    if (accessToken && gId && !String(gId).startsWith("local-")) {
      fetch(`https://graph.microsoft.com/v1.0/me/events/${gId}`, {
        method: "DELETE",
        headers: { Authorization: `Bearer ${accessToken}` },
      }).catch(e => console.error("Delete event error:", e));
    }
  };

  // ─── Theme helpers ────────────────────────────────────────────────────────
  const th = {
    bg: isDark ? "bg-gray-900" : "bg-white",
    text: isDark ? "text-white" : "text-gray-900",
    textMuted: isDark ? "text-white" : "text-gray-500",
    border: isDark ? "border-gray-600" : "border-gray-200",
    surface: isDark ? "bg-gray-800" : "bg-gray-50",
    hover: isDark ? "hover:bg-gray-700" : "hover:bg-gray-100",
    input: isDark ? "bg-gray-700 border-gray-600 text-white placeholder-gray-400" : "bg-white border-gray-300 text-gray-900 placeholder-gray-400",
    card: isDark ? "bg-gray-800 border-gray-600" : "bg-white border-gray-200",
  };

  // ─── Highlight helper ─────────────────────────────────────────────────────
  const hl = text => {
    const q = emailSearch.trim().toLowerCase();
    if (!q || !text) return text;
    const str = String(text);
    const idx = str.toLowerCase().indexOf(q);
    if (idx === -1) return str;
    return <span>{str.slice(0, idx)}<mark className="bg-yellow-200 text-yellow-900 font-semibold rounded px-0.5">{str.slice(idx, idx + q.length)}</mark>{str.slice(idx + q.length)}</span>;
  };

  // ─── All folder definitions ───────────────────────────────────────────────
  const FOLDERS = [
    { key: "Inbox", label: "Inbox", icon: Inbox },
    { key: "Sent", label: "Sent Items", icon: Send },
    { key: "Drafts", label: "Drafts", icon: Edit2 },
    { key: "Deleted", label: "Deleted Items", icon: Trash2 },
    { key: "Junk", label: "Junk Email", icon: AlertTriangle },
    { key: "Archive", label: "Archive", icon: Archive },
    { key: "Outbox", label: "Outbox", icon: MailOpen },
    { key: "History", label: "Conversation History", icon: MessageSquare },
  ];

  const EMAIL_FOLDERS = ["Inbox", "Sent", "Drafts", "Deleted", "Junk", "Archive", "Outbox", "History", "Campaign"];
  const isEmailFolder = EMAIL_FOLDERS.includes(activeFolder);

  // Open the compose inline in the reading pane. If we're not on a mail folder
  // (e.g. Templates), switch to Inbox first so it never falls back to a modal overlay.
  const startCompose = (rd = null) => {
    if (!EMAIL_FOLDERS.includes(activeFolder)) setActiveFolder("Inbox");
    setOpenedMailId(null);
    setReplyData(rd);
    setShowCompose(true);
  };

  // Leaving the mail folders (e.g. opening Templates) closes any open compose,
  // so the Templates page shows instead of the New Message panel.
  useEffect(() => { if (!isEmailFolder) setShowCompose(false); }, [isEmailFolder]);

  // Arrived here from the Contacts "Send Email" button: open the compose window
  // immediately so the user lands on the composer, not the inbox.
  useEffect(() => {
    try {
      if (localStorage.getItem("openComposeOnLoad") === "1") {
        localStorage.removeItem("openComposeOnLoad");
        startCompose();
      }
    } catch (e) { /* ignore */ }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // ─── Render ────────────────────────────────────────────────────────────────
  return (
    <div className={`flex flex-col h-full ${th.bg} ${th.text} overflow-hidden relative${isDark ? " outlook-dark" : ""}`}>
      {isDark && (
        <style>{`
          .outlook-dark * { color: #ffffff !important; }
          .outlook-dark input, .outlook-dark textarea, .outlook-dark select { color: #ffffff !important; }
          .outlook-dark input::placeholder, .outlook-dark textarea::placeholder { color: rgba(255,255,255,0.5) !important; }
          .outlook-dark .text-blue-600, .outlook-dark .text-blue-500, .outlook-dark .text-blue-400 { color: #93c5fd !important; }
          .outlook-dark .text-red-500, .outlook-dark .text-red-600 { color: #fca5a5 !important; }
          .outlook-dark .text-green-600, .outlook-dark .text-emerald-600 { color: #6ee7b7 !important; }
          .outlook-dark .text-orange-500, .outlook-dark .text-orange-400 { color: #fdba74 !important; }
          .outlook-dark .text-purple-500, .outlook-dark .text-purple-600 { color: #c4b5fd !important; }
          .outlook-dark .text-yellow-400, .outlook-dark .text-yellow-500 { color: #fde047 !important; }
          .outlook-dark .text-white { color: #ffffff !important; }
          .outlook-dark .outlook-email-body, .outlook-dark .outlook-email-body * { color: #ffffff !important; background-color: transparent !important; background: transparent !important; }
        `}</style>
      )}

      {/* ── Message Box Popup (created / updated / deleted) ───────────────── */}
      {popup && (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-[60] p-4" onClick={() => setPopup(null)}>
          <div className={`${th.card} border w-full max-w-xs rounded-xl shadow-xl p-6 text-center`} onClick={e => e.stopPropagation()}>
            <div className={`w-14 h-14 mx-auto mb-4 rounded-full flex items-center justify-center ${popup.type === "error" ? (isDark ? "bg-red-500/15 text-red-400" : "bg-red-100 text-red-600") : (isDark ? "bg-green-500/15 text-green-400" : "bg-green-100 text-green-600")}`}>
              {popup.type === "error" ? <AlertCircle size={28} /> : <Check size={30} strokeWidth={3} />}
            </div>
            <p className={`text-base font-semibold ${th.text}`}>{popup.message}</p>
            <button onClick={() => setPopup(null)}
              className="mt-5 px-6 py-2 text-sm rounded-lg bg-blue-500 text-white hover:bg-blue-600 transition">
              OK
            </button>
          </div>
        </div>
      )}

      {/* ── Delete Template Confirm Message Box ───────────────────────────── */}
      {templateToDelete && (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-[60] p-4" onClick={() => setTemplateToDelete(null)}>
          <div className={`${th.card} border w-full max-w-sm rounded-xl shadow-xl p-6`} onClick={e => e.stopPropagation()}>
            <div className={`w-14 h-14 mx-auto mb-4 rounded-full flex items-center justify-center ${isDark ? "bg-red-500/15 text-red-400" : "bg-red-100 text-red-600"}`}>
              <Trash2 size={26} />
            </div>
            <h4 className={`text-base font-semibold text-center mb-1 ${th.text}`}>Delete template</h4>
            <p className={`text-sm text-center ${th.textMuted} mb-5`}>
              Are you sure you want to delete{templateToDelete.name || templateToDelete.body ? ` "${(templateToDelete.name || templateToDelete.body).slice(0, 40)}"` : " this template"}? This can't be undone.
            </p>
            <div className="flex justify-center gap-3">
              <button onClick={() => setTemplateToDelete(null)}
                className={`px-5 py-2 text-sm rounded-lg ${isDark ? "bg-gray-700 text-gray-200 hover:bg-gray-600" : "bg-gray-100 text-gray-700 hover:bg-gray-200"}`}>
                Cancel
              </button>
              <button onClick={confirmDeleteTemplate}
                className="px-5 py-2 text-sm rounded-lg bg-red-600 text-white hover:bg-red-700">
                Delete
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── Delete Confirm Modal ─────────────────────────────────────────── */}
      {showDeleteConfirm && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
          <div className={`${th.card} border w-full max-w-sm rounded-xl shadow-xl p-6`}>
            <h4 className="text-base font-semibold mb-2">Delete messages</h4>
            <p className={`text-sm ${th.textMuted} mb-5`}>
              Are you sure you want to permanently delete {deleteCandidateIds.length} message(s)?
            </p>
            <div className="flex justify-end gap-3">
              <button onClick={() => { setShowDeleteConfirm(false); setDeleteCandidateIds([]); }}
                className={`px-4 py-2 text-sm rounded-lg ${isDark ? "bg-gray-700 text-gray-200 hover:bg-gray-600" : "bg-gray-100 text-gray-700 hover:bg-gray-200"}`}>
                Cancel
              </button>
              <button onClick={performDelete}
                className="px-4 py-2 text-sm rounded-lg bg-red-600 text-white hover:bg-red-700">
                Delete
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── Settings Modal ───────────────────────────────────────────────── */}
      {showSettings && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-2 sm:p-4">
          <div className={`${th.card} border w-full max-w-2xl max-h-[95vh] rounded-xl shadow-xl flex flex-col overflow-hidden`}>
            <div className={`flex items-center justify-between px-4 py-3 border-b ${th.border} shrink-0`}>
              <h2 className="text-base font-semibold">Settings</h2>
              <button onClick={() => setShowSettings(false)} className={`p-2 rounded-lg ${th.hover}`}><X size={18} /></button>
            </div>
            <div className="flex flex-1 overflow-hidden min-h-0">
              {/* Settings nav — horizontal scroll on mobile */}
              <nav className={`flex flex-row md:flex-col md:w-44 border-b md:border-b-0 md:border-r ${th.border} p-2 md:p-3 gap-1 shrink-0 overflow-x-auto md:overflow-x-visible`}>
                {[
                  { key: "account", label: "Account", icon: User },
                  { key: "signature", label: "Signature", icon: Pencil },
                  { key: "theme", label: "Theme", icon: Sun },
                  { key: "notifications", label: "Notifications", icon: Bell },
                  { key: "rules", label: "Rules", icon: Filter },
                  { key: "autoreply", label: "Auto Reply", icon: RefreshCw },
                ].map(({ key, label, icon: Icon }) => (
                  <button key={key} onClick={() => setSettingsTab(key)}
                    className={`flex items-center gap-2 px-3 py-2 rounded-lg text-xs md:text-sm whitespace-nowrap transition shrink-0 ${settingsTab === key ? "bg-blue-500 text-white" : `${th.hover} ${th.text}`}`}>
                    <Icon size={14} />{label}
                  </button>
                ))}
              </nav>
              {/* Settings content */}
              <div className="flex-1 overflow-y-auto p-4 md:p-6">
                {settingsTab === "account" && (
                  <div className="space-y-4">
                    <h3 className="font-semibold text-base">Account Settings</h3>
                    <div className={`flex items-center gap-4 p-4 rounded-xl border ${th.border}`}>
                      <div className={`w-12 h-12 md:w-14 md:h-14 rounded-full ${getAvatarColor(accounts[0]?.name || "")} flex items-center justify-center text-white text-lg md:text-xl font-bold shrink-0`}>
                        {getInitials(accounts[0]?.name || accounts[0]?.username || "U")}
                      </div>
                      <div className="min-w-0">
                        <p className="font-medium truncate">{accounts[0]?.name || "User"}</p>
                        <p className={`text-sm ${th.textMuted} truncate`}>{accounts[0]?.username || ""}</p>
                        <span className="inline-block mt-1 px-2 py-0.5 bg-green-100 text-green-700 text-xs rounded-full">Connected</span>
                      </div>
                    </div>
                  </div>
                )}
                {settingsTab === "signature" && (
                  <div className="space-y-4">
                    <h3 className="font-semibold text-base">Email Signature</h3>
                    <textarea
                      value={signature}
                      onChange={e => setSignature(e.target.value)}
                      rows={6}
                      className={`w-full px-4 py-3 border rounded-xl text-sm ${th.input} resize-y focus:outline-none focus:ring-2 focus:ring-blue-500`}
                      placeholder="Enter your email signature..."
                    />
                    <button className="px-4 py-2 bg-blue-500 text-white rounded-lg text-sm hover:bg-blue-600">Save Signature</button>
                  </div>
                )}
                {settingsTab === "theme" && (
                  <div className="space-y-4">
                    <h3 className="font-semibold text-base">Theme</h3>
                    <div className="flex gap-4">
                      {[{ v: false, label: "Light", icon: Sun }, { v: true, label: "Dark", icon: Moon }].map(({ v, label, icon: Icon }) => (
                        <button key={label} onClick={() => setIsDark(v)}
                          className={`flex-1 flex flex-col items-center gap-2 p-4 rounded-xl border-2 transition ${isDark === v ? "border-blue-500 bg-blue-50 text-blue-700" : `border-transparent ${th.border} ${th.hover}`}`}>
                          <Icon size={24} /><span className="text-sm font-medium">{label}</span>
                        </button>
                      ))}
                    </div>
                  </div>
                )}
                {settingsTab === "notifications" && (
                  <div className="space-y-4">
                    <h3 className="font-semibold text-base">Notifications</h3>
                    {[
                      { label: "New email alerts", desc: "Get notified for new emails" },
                      { label: "Meeting reminders", desc: "Reminders 15 min before meetings" },
                      { label: "Mentions", desc: "When you're @mentioned" },
                      { label: "Task alerts", desc: "Due date reminders" },
                    ].map(({ label, desc }) => (
                      <div key={label} className={`flex items-center justify-between p-3 rounded-lg border ${th.border}`}>
                        <div className="min-w-0 mr-3">
                          <p className="text-sm font-medium">{label}</p>
                          <p className={`text-xs ${th.textMuted}`}>{desc}</p>
                        </div>
                        <div className="w-10 h-6 bg-blue-500 rounded-full flex items-center justify-end px-0.5 cursor-pointer shrink-0">
                          <div className="w-5 h-5 bg-white rounded-full shadow" />
                        </div>
                      </div>
                    ))}
                  </div>
                )}
                {settingsTab === "autoreply" && (
                  <div className="space-y-4">
                    <h3 className="font-semibold text-base">Automatic Replies</h3>
                    <label className="flex items-center gap-3 cursor-pointer">
                      <input type="checkbox" checked={autoReply.enabled} onChange={e => setAutoReply(p => ({ ...p, enabled: e.target.checked }))} className="w-4 h-4" />
                      <span className="text-sm">Send automatic replies</span>
                    </label>
                    {autoReply.enabled && (
                      <textarea
                        value={autoReply.message}
                        onChange={e => setAutoReply(p => ({ ...p, message: e.target.value }))}
                        rows={5}
                        className={`w-full px-4 py-3 border rounded-xl text-sm ${th.input} resize-y focus:outline-none focus:ring-2 focus:ring-blue-500`}
                        placeholder="I'm currently out of office..."
                      />
                    )}
                    <button className="px-4 py-2 bg-blue-500 text-white rounded-lg text-sm hover:bg-blue-600">Save</button>
                  </div>
                )}
                {settingsTab === "rules" && (
                  <div className="space-y-4">
                    <h3 className="font-semibold text-base">Mail Rules</h3>
                    <p className={`text-sm ${th.textMuted}`}>Create rules to automatically organize your incoming mail.</p>
                    <button className="flex items-center gap-2 px-4 py-2 bg-blue-500 text-white rounded-lg text-sm hover:bg-blue-600">
                      <Plus size={14} /> New Rule
                    </button>
                    <div className={`p-8 rounded-xl border ${th.border} text-center`}>
                      <Filter size={32} className={`mx-auto mb-2 ${th.textMuted}`} />
                      <p className={`text-sm ${th.textMuted}`}>No rules yet. Create one to get started.</p>
                    </div>
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ── Notification Panel ───────────────────────────────────────────── */}
      {showNotifications && (
        <div className="fixed inset-0 z-40" onClick={() => setShowNotifications(false)}>
          {/* Notification panel attached to bell icon */}
          <div
            className={`absolute top-16 right-4 w-[min(100vw-1rem,24rem)] ${th.card} border ${th.border} rounded-xl shadow-xl overflow-hidden flex flex-col`}
            style={{ maxHeight: "calc(100vh - 4.5rem)" }}
            onClick={e => e.stopPropagation()}>
            <div className={`flex items-center justify-between px-4 py-3 border-b ${th.border} shrink-0`}>
              <div className="flex items-center gap-2">
                <span className="font-semibold text-sm">Notifications</span>
                {notifLoading && <RefreshCw size={12} className="animate-spin text-blue-400" />}
              </div>
              <button onClick={() => setShowNotifications(false)} className={`p-1 rounded ${th.hover}`}><X size={16} /></button>
            </div>
            <div className="overflow-y-auto flex-1">
              {notifLoading && notifications.length === 0 ? (
                <div className="flex flex-col items-center justify-center py-10 gap-2">
                  <RefreshCw size={20} className={`animate-spin ${th.textMuted}`} />
                  <p className={`text-xs ${th.textMuted}`}>Loading notifications…</p>
                </div>
              ) : notifications.length === 0 ? (
                <div className="flex flex-col items-center justify-center py-10 gap-2">
                  <Bell size={28} className={`${th.textMuted} opacity-30`} />
                  <p className={`text-xs ${th.textMuted}`}>No notifications</p>
                </div>
              ) : (
                notifications.map(n => (
                  <div key={n.id}
                    onClick={() => markNotifRead(n.id)}
                    className={`px-4 py-3 border-b ${th.border} ${th.hover} cursor-pointer flex gap-3 ${!n.read ? isDark ? "bg-blue-900/20" : "bg-blue-50" : ""}`}>
                    <div className={`w-8 h-8 rounded-full flex items-center justify-center flex-shrink-0
                      ${n.type === "email" ? "bg-blue-100 text-blue-600" :
                        n.type === "meeting" ? "bg-green-100 text-green-600" :
                          n.type === "task" ? "bg-orange-100 text-orange-600" :
                            "bg-purple-100 text-purple-600"}`}>
                      {n.type === "email" ? <Mail size={14} /> :
                        n.type === "meeting" ? <Calendar size={14} /> :
                          n.type === "task" ? <CheckSquare size={14} /> :
                            <AtSign size={14} />}
                    </div>
                    <div className="flex-1 min-w-0">
                      <p className={`text-sm font-medium truncate ${!n.read ? "text-blue-600" : th.text}`}>{n.title}</p>
                      <p className={`text-xs ${th.textMuted} truncate`}>{n.body}</p>
                      <p className={`text-xs ${th.textMuted} mt-0.5`}>{n.time}</p>
                    </div>
                    {!n.read && <div className="w-2 h-2 bg-blue-500 rounded-full mt-1.5 flex-shrink-0" />}
                  </div>
                ))
              )}
            </div>
            <div className={`px-4 py-2 border-t ${th.border} flex items-center justify-between shrink-0`}>
              <button onClick={markAllNotifsRead} className="text-xs text-blue-500 hover:underline">
                Mark all as read
              </button>
              <button
                onClick={() => setNotifRefreshKey(p => p + 1)}
                className={`text-xs ${th.textMuted} hover:text-blue-500 flex items-center gap-1`}>
                <RefreshCw size={11} />Refresh
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── Top Header ───────────────────────────────────────────────────── */}
      <header className={`h-12 ${isDark ? "bg-gray-950 border-gray-800 text-gray-100" : "bg-white border-gray-200 text-gray-900"} border-b flex items-center justify-between px-2 sm:px-3 gap-2 shrink-0 z-20`}>
        {/* Left: Hamburger + Logo */}
        <div className="flex items-center gap-1.5 shrink-0">
          <button onClick={() => setShowMobileSidebar(p => !p)}
            className={`md:hidden p-2 rounded-lg transition ${isDark ? "text-gray-300 hover:text-white hover:bg-white/10" : "text-gray-700 hover:text-gray-900 hover:bg-gray-100"}`}>
            <Menu size={18} />
          </button>
          <div className={`w-7 h-7 rounded flex items-center justify-center shrink-0 ${isDark ? "bg-white/10" : "bg-gray-100"}`}>
            <Mail size={16} className={isDark ? "text-white" : "text-gray-700"} />
          </div>
          <span className={`font-semibold text-sm hidden sm:block ${isDark ? "text-white" : "text-gray-900"}`}>Outlook</span>
        </div>

        {/* Center: Search */}
        <div className="flex-1 max-w-lg mx-1 sm:mx-2">
          <div className="relative">
            <Search size={14} className={`absolute left-2.5 top-1/2 -translate-y-1/2 pointer-events-none ${isDark ? "text-gray-400" : "text-gray-500"}`} />
            <input
              type="text"
              placeholder="Search..."
              value={emailSearch}
              onChange={e => { setEmailSearch(e.target.value); setOpenedMailId(null); }}
              className={`w-full pl-8 pr-7 py-1.5 text-sm rounded-lg border transition ${isDark ? "bg-gray-800 text-gray-100 placeholder-gray-400 border-gray-700 focus:bg-gray-700 focus:border-gray-600" : "bg-gray-100 text-gray-900 placeholder-gray-500 border-gray-300 focus:bg-white focus:border-gray-400"}`}
            />
            {emailSearch && (
              <button onClick={() => setEmailSearch("")} className={`absolute right-2 top-1/2 -translate-y-1/2 ${isDark ? "text-gray-400 hover:text-white" : "text-gray-500 hover:text-gray-900"}`}>
                <X size={14} />
              </button>
            )}
          </div>
        </div>

        {/* Right: Actions */}
        <div className="flex items-center gap-0.5 shrink-0" />
      </header>

      {/* ── Mobile Sidebar Overlay ───────────────────────────────────────── */}
      {showMobileSidebar && (
        <div className="fixed inset-0 z-50 md:hidden" onClick={() => setShowMobileSidebar(false)}>
          <div className="absolute inset-0 bg-black/50" />
          <div className={`absolute left-0 top-0 bottom-0 w-72 ${th.surface} border-r ${th.border} flex flex-col overflow-hidden`}
            onClick={e => e.stopPropagation()}>
            <div className={`flex items-center justify-between px-4 py-3 border-b ${th.border} shrink-0`}>
              <span className="font-semibold text-sm">Folders</span>
              <button onClick={() => setShowMobileSidebar(false)} className={`p-2 rounded-lg ${th.hover}`}><X size={16} /></button>
            </div>
            <div className="p-3 shrink-0">
              <button onClick={() => { startCompose(); setShowMobileSidebar(false); }}
                className="w-full flex items-center justify-center gap-2 bg-blue-500 hover:bg-blue-600 text-white text-sm px-3 py-3 rounded-xl transition font-medium">
                <Plus size={16} />New email
              </button>
            </div>
            <nav className="flex-1 overflow-y-auto px-2 pb-4 space-y-0.5">
              {[
                { key: "Inbox", label: "Inbox", icon: Mail, count: folderCounts.Inbox },
                { key: "Sent", label: "Sent", icon: Send },
                { key: "Drafts", label: "Drafts", icon: Pencil },
                { key: "Archive", label: "Archive", icon: Archive },
                { key: "Junk", label: "Junk", icon: AlertCircle },
                { key: "Deleted", label: "Trash", icon: Trash2 },
              ].map(({ key, label, icon: Icon, count }) => (
                <button key={key} onClick={() => { setActiveFolder(key); setOpenedMailId(null); setMobileView("list"); setShowMobileSidebar(false); }}
                  className={`w-full flex items-center gap-3 px-3 py-3 rounded-lg text-sm transition ${activeFolder === key ? "bg-blue-500 text-white" : `${th.hover} ${th.text}`}`}>
                  <Icon size={16} />{label}
                  {count > 0 && <span className="ml-auto text-xs font-semibold bg-blue-500 text-white px-1.5 py-0.5 rounded-full">{count}</span>}
                </button>
              ))}
              <div className={`my-2 border-t ${th.border}`} />
              <button onClick={() => { setActiveFolder("Templates"); setOpenedMailId(null); setShowMobileSidebar(false); }}
                className={`w-full flex items-center gap-3 px-3 py-3 rounded-lg text-sm transition ${activeFolder === "Templates" ? "bg-blue-500 text-white" : `${th.hover} ${th.text}`}`}>
                <FileText size={16} />Templates
              </button>
              <div className={`my-2 border-t ${th.border}`} />
              {/* Settings in mobile sidebar */}
              <button onClick={() => { setShowSettings(true); setShowMobileSidebar(false); }}
                className={`w-full flex items-center gap-3 px-3 py-3 rounded-lg text-sm transition ${th.hover} ${th.text}`}>
                <Settings size={16} />Settings
              </button>
              <button onClick={() => setIsDark(p => !p)}
                className={`w-full flex items-center gap-3 px-3 py-3 rounded-lg text-sm transition ${th.hover} ${th.text}`}>
                {isDark ? <Sun size={16} /> : <Moon size={16} />}{isDark ? "Light mode" : "Dark mode"}
              </button>
            </nav>
          </div>
        </div>
      )}

      {/* ── Main Body ────────────────────────────────────────────────────── */}
      <div className="flex flex-1 overflow-hidden min-h-0">

        {/* ── Left App Rail (desktop only) ────────────────────────────────── */}
        <nav className={`hidden md:flex flex-col items-center py-3 gap-1 w-14 shrink-0 ${isDark ? "bg-gray-950 border-r border-gray-800" : "bg-white border-r border-gray-200"}`}>
          <button onClick={() => setShowFolderPane(p => !p)}
            title={showFolderPane ? "Collapse folder pane" : "Expand folder pane"}
            className={`w-10 h-10 flex items-center justify-center rounded-lg transition mb-1 ${isDark ? "text-gray-300 hover:bg-white/10 hover:text-white" : "text-gray-700 hover:bg-gray-100 hover:text-gray-900"}`}>
            <Menu size={20} />
          </button>
          {[
            { key: "Mail", icon: Mail, badge: folderCounts.Inbox },
            { key: "Calendar", icon: Calendar },
            { key: "People", icon: Users },
            { key: "Tasks", icon: CheckSquare },
            { key: "Files", icon: FolderOpen },
          ].map(({ key, icon: Icon, badge }) => (
            <RailItem key={key} icon={Icon} label={key} active={activeApp === key}
              onClick={() => setActiveApp(key)} badge={badge || 0} isDark={isDark} />
          ))}
          <div className="flex-1" />
          <button
            onClick={() => setIsDark(p => !p)}
            title={isDark ? "Switch to light mode" : "Switch to dark mode"}
            className={`w-10 h-10 flex items-center justify-center rounded-lg transition ${isDark ? "text-gray-300 hover:bg-white/10 hover:text-white" : "text-gray-700 hover:bg-gray-100 hover:text-gray-900"}`}
          >
            {isDark ? <Sun size={20} /> : <Moon size={20} />}
          </button>
        </nav>

        {/* ── Mail App ─────────────────────────────────────────────────── */}
        {activeApp === "Mail" && (
          <>
            {/* ── Mail Sidebar (desktop only; collapses while reading a mail) ── */}
            {showFolderPane && (
            <aside className={`hidden md:flex md:w-56 ${th.surface} border-r ${th.border} flex-col shrink-0`}>
              <div className="p-3">
                <button onClick={() => startCompose()}
                  className="w-full flex items-center justify-center gap-2 bg-blue-500 hover:bg-blue-600 text-white text-sm px-3 py-2.5 rounded-xl transition shadow-sm font-medium">
                  <Plus size={16} />New email
                </button>
              </div>

              <nav className="flex-1 overflow-y-auto px-2 space-y-0.5">
                <div>
                  <button
                    onClick={() => setExpandedSections(p => ({ ...p, favorites: !p.favorites }))}
                    className={`w-full flex items-center gap-1.5 px-2 py-1.5 text-xs font-semibold uppercase tracking-wider ${th.textMuted} ${th.hover} rounded-lg`}>
                    {expandedSections.favorites ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
                    Favorites
                  </button>
                  {expandedSections.favorites && pinnedFolders.map(f => {
                    const folder = FOLDERS.find(fd => fd.key === f || fd.label === f);
                    if (!folder) return null;
                    return (
                      <FolderRow key={f} icon={Star} label={folder.label} active={activeFolder === folder.key}
                        count={folderCounts[folder.key] || 0} pinned
                        onClick={() => { setActiveFolder(folder.key); setOpenedMailId(null); }}
                        isDark={isDark} />
                    );
                  })}
                </div>

                <div className={`my-2 border-t ${th.border}`} />

                <div>
                  <button
                    onClick={() => setExpandedSections(p => ({ ...p, folders: !p.folders }))}
                    className={`w-full flex items-center gap-1.5 px-2 py-1.5 text-xs font-semibold uppercase tracking-wider ${th.textMuted} ${th.hover} rounded-lg`}>
                    {expandedSections.folders ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
                    Folders
                  </button>
                  {expandedSections.folders && FOLDERS.map(({ key, label, icon }) => (
                    <FolderRow key={key} icon={icon} label={label} active={activeFolder === key}
                      count={folderCounts[key] || 0}
                      onPin={() => setPinnedFolders(p => p.includes(label) ? p.filter(x => x !== label) : [...p, label])}
                      pinned={pinnedFolders.includes(label)}
                      onClick={() => { setActiveFolder(key); setOpenedMailId(null); }}
                      isDark={isDark} />
                  ))}
                  <button className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-lg text-sm ${th.textMuted} ${th.hover} transition`}>
                    <Plus size={14} /><span>Add folder</span>
                  </button>
                </div>

                <div className={`my-2 border-t ${th.border}`} />

                <FolderRow icon={FileText} label="Templates" active={activeFolder === "Templates"}
                  onClick={() => { setActiveFolder("Templates"); setOpenedMailId(null); }}
                  isDark={isDark} />

                <div className={`my-2 border-t ${th.border}`} />

                {/* ── Campaign folder: click to view sent copies + replies; pencil to rename ── */}
                <div className="px-1">
                  <div className={`flex items-center justify-between px-2 py-1.5 text-xs font-semibold uppercase tracking-wider ${th.textMuted}`}>
                    <span>Campaign folder</span>
                    <button type="button" onClick={() => { setEditCampaignFolder(v => !v); loadOutlookFolders(); }}
                      className={`${th.hover} rounded p-0.5`} title="Change campaign folder">
                      <Pencil size={12} />
                    </button>
                  </div>
                  <FolderRow icon={FolderOpen} label={campaignFolder || "CRM Campaigns"}
                    active={activeFolder === "Campaign"}
                    count={campaignMails.length}
                    onClick={() => { setActiveFolder("Campaign"); setOpenedMailId(null); }}
                    isDark={isDark} />
                  {editCampaignFolder && (
                    <div className="px-2 pt-1 pb-1">
                      <input
                        list="crm-campaign-folders"
                        value={campaignFolder}
                        onChange={(e) => persistCampaignFolder(e.target.value)}
                        onFocus={loadOutlookFolders}
                        placeholder="CRM Campaigns"
                        className={`w-full px-2.5 py-1.5 text-sm rounded-lg border outline-none focus:ring-2 focus:ring-blue-500 ${isDark ? "bg-gray-800 border-gray-600 text-gray-100" : "bg-white border-gray-300 text-gray-800"}`}
                      />
                      <datalist id="crm-campaign-folders">
                        {outlookFolders.map((f) => <option key={f} value={f} />)}
                      </datalist>
                      <p className={`text-[11px] mt-1 leading-snug ${th.textMuted}`}>
                        {campaignFolderSaved ? "Saved ✓ — " : ""}Tracked-campaign sends &amp; their replies are filed here (created if new).
                      </p>
                    </div>
                  )}
                </div>

                <div className={`my-2 border-t ${th.border}`} />

                <div>
                  <button
                    onClick={() => setExpandedSections(p => ({ ...p, groups: !p.groups }))}
                    className={`w-full flex items-center gap-1.5 px-2 py-1.5 text-xs font-semibold uppercase tracking-wider ${th.textMuted} ${th.hover} rounded-lg`}>
                    {expandedSections.groups ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
                    Groups
                  </button>
                  {expandedSections.groups && (
                    <div className={`px-3 py-3 text-xs ${th.textMuted} text-center`}>No groups yet</div>
                  )}
                </div>
              </nav>
            </aside>
            )}

            {/* ── Mobile folder tab bar ──────────────────────────────── */}
            {/* Only show when NOT in email reading view on mobile */}
            {/* <div className={`${mobileView === "email" && openedMailId ? "hidden" : "md:hidden"} absolute left-0 right-0 border-b ${th.border} ${th.surface} px-2 py-2 overflow-x-auto shrink-0 z-10`}
              style={{ top: "48px" }}>
              <div className="flex gap-1.5 min-w-max">
                {[...FOLDERS, { key: "Templates", label: "Templates", icon: FileText }].map(({ key, label }) => (
                  <button key={key} onClick={() => { setActiveFolder(key); setOpenedMailId(null); setMobileView("list"); }}
                    className={`px-3 py-1.5 rounded-full text-xs whitespace-nowrap font-medium transition ${activeFolder === key ? "bg-blue-500 text-white" : isDark ? "bg-gray-700 text-gray-300" : "bg-gray-100 text-gray-700"}`}>
                    {label}{folderCounts[key] ? ` (${folderCounts[key]})` : ""}
                  </button>
                ))}
              </div>
            </div> */}

            {/* ── Email Area (with top offset on mobile for folder tabs) ── */}
            {/* On mobile, folder tab bar is 44px tall, so we offset the content */}
            <div className="flex flex-1 overflow-hidden min-h-0">

              {/* Email List + Reading Pane */}
              {isEmailFolder && (
                <>
                  {/* ── Email List Panel ─────────────────────────────── */}
                  {/* Mobile: show list when mobileView==="list"; Desktop: always show */}
                  <div className={`
  ${showCompose && composeTemplatesOpen
      ? "hidden"
      : `${(openedMailId && mobileView === "email") || showCompose ? "hidden md:flex" : "flex md:flex"}`}
  w-full md:w-96 border-r ${th.border} flex-col ${th.bg} shrink-0
`}>

                    {/* List header */}
                    <div className={`border-b ${th.border} ${th.surface} shrink-0`}>
                      <div className="h-12 px-3 flex items-center gap-2">
                        <input type="checkbox" aria-label="Select all"
                          checked={(() => { const m = getFilteredMails(); return m.length > 0 && m.every(x => selectedMailIds.includes(x.id)); })()}
                          onChange={selectAll}
                          className="w-4 h-4 rounded border-gray-300 shrink-0" />
                        <h2 className={`flex-1 text-sm font-semibold truncate ${th.text}`}>
                          {activeFolder === "Deleted" ? "Deleted Items"
                            : activeFolder === "Campaign" ? (campaignFolder || "CRM Campaigns")
                              : activeFolder}
                        </h2>
                        <div className="flex items-center gap-0.5">
                          {selectedMailIds.length > 0 && (
                            <>
                              <button onClick={() => confirmDelete(selectedMailIds)} title="Delete"
                                className={`p-2 rounded-lg text-red-500 ${th.hover} transition`}><FiTrash2 size={15} /></button>
                              {activeFolder === "Inbox" && (
                                <button onClick={() => { selectedMailIds.forEach(id => archiveMail(id)); }} title="Archive"
                                  className={`p-2 rounded-lg ${th.textMuted} ${th.hover} transition`}><Archive size={15} /></button>
                              )}
                              <button onClick={() => selectedMailIds.forEach(id => toggleRead(getCurrentMails().find(m => m.id === id)))} title="Mark read/unread"
                                className={`p-2 rounded-lg ${th.textMuted} ${th.hover} transition`}><MailOpen size={15} /></button>
                              <button onClick={() => selectedMailIds.forEach(id => toggleFlag(id))} title="Flag"
                                className={`p-2 rounded-lg ${th.textMuted} ${th.hover} transition`}><Flag size={15} /></button>
                            </>
                          )}
                          {/* Sort */}
                          <div className="relative" onClick={e => e.stopPropagation()}>
                            <button onClick={() => setShowSortMenu(p => !p)} title="Sort"
                              className={`p-2 rounded-lg ${th.textMuted} ${th.hover} transition`}><ArrowUpDown size={15} /></button>
                            {showSortMenu && (
                              <div className={`absolute right-0 mt-2 top-10 w-40 ${th.card} border ${th.border} rounded-xl shadow-xl z-30 overflow-hidden`}>
                                {[["date", "Date"], ["sender", "Sender"], ["subject", "Subject"]].map(([v, l]) => (
                                  <button key={v} onClick={() => { setSortBy(v); setShowSortMenu(false); }}
                                    className={`w-full flex items-center gap-2 px-3 py-2.5 text-sm ${th.hover} ${sortBy === v ? "text-blue-500" : th.text}`}>
                                    {sortBy === v && <Check size={12} />}{l}
                                  </button>
                                ))}
                              </div>
                            )}
                          </div>
                          {/* Filter */}
                          <div className="relative" onClick={e => e.stopPropagation()}>
                            <button onClick={() => setShowFilterMenu(p => !p)} title="Filter"
                              className={`p-2 rounded-lg transition ${filterBy !== "all" ? "text-blue-500" : th.textMuted} ${th.hover}`}><Filter size={15} /></button>
                            {showFilterMenu && (
                              <div className={`absolute right-0 top-10 w-44 ${th.card} border ${th.border} rounded-xl shadow-xl z-30 overflow-hidden`}>
                                {[["all", "All"], ["unread", "Unread"], ["flagged", "Flagged"], ["attachments", "Has attachments"]].map(([v, l]) => (
                                  <button key={v} onClick={() => { setFilterBy(v); setShowFilterMenu(false); }}
                                    className={`w-full flex items-center gap-2 px-3 py-2.5 text-sm ${th.hover} ${filterBy === v ? "text-blue-500" : th.text}`}>
                                    {filterBy === v && <Check size={12} />}{l}
                                  </button>
                                ))}
                              </div>
                            )}
                          </div>
                        </div>
                      </div>

                      {/* Focused / Other tabs for Inbox */}
                      {activeFolder === "Inbox" && (
                        <div className={`flex border-t ${th.border}`}>
                          {["Focused", "Other"].map(tab => (
                            <button key={tab} onClick={() => setInboxTab(tab)}
                              className={`flex-1 py-2.5 text-sm font-medium transition border-b-2 ${inboxTab === tab ? "border-blue-600 text-blue-600" : `border-transparent ${th.textMuted} ${th.hover}`}`}>
                              {tab}
                            </button>
                          ))}
                        </div>
                      )}
                    </div>

                    {/* Email list items */}
                    <div className={`flex-1 overflow-y-auto ${isDark ? "bg-[#1b1a19]" : "bg-white"}`}>
                      {(() => {
                        const mails = getFilteredMails();
                        if (!mails.length) return (
                          <div className="flex flex-col items-center justify-center h-full py-12 px-6 text-center">
                            <div className={`w-14 h-14 rounded-full flex items-center justify-center mb-3 ${isDark ? "bg-gray-700" : "bg-gray-100"}`}>
                              <Inbox size={24} className={`${th.textMuted} opacity-70`} />
                            </div>
                            <p className={`text-sm font-medium ${th.text}`}>{emailSearch ? "No results found" : "No messages"}</p>
                            <p className={`text-xs mt-1 ${th.textMuted}`}>{emailSearch ? `Nothing matches "${emailSearch}"` : "Your messages will appear here"}</p>
                          </div>
                        );
                        return mails.map(mail => {
                          const isRead = readOverrides[mail.id] !== undefined ? readOverrides[mail.id] : mail.isRead;
                          const isFlagged = flaggedIds[mail.id];
                          const cat = mailCategories[mail.id];
                          const isSelected = selectedMailIds.includes(mail.id);
                          const isOpen = openedMailId === mail.id;
                          return (
                            <div key={mail.id}
                              onClick={() => {
                                setShowCompose(false);
                                setReplyData(null);
                                setOpenedMailId(mail.id);
                                markAsRead(mail);
                                setInlineReply(null);
                                setInlineBody("");
                                setInlineError("");
                                setMobileView("email");
                              }}
                              className={`relative px-3 py-3 cursor-pointer border-b transition ${isDark ? "border-[#3d3b39]" : "border-[#E1E1E1]"} ${isOpen ? isDark ? "bg-[#252423]" : "bg-[#EBF3FB]" : isDark ? "hover:bg-[#252423]" : "hover:bg-[#F3F2F1]"} ${isSelected ? isDark ? "bg-[#3d3b39]" : "bg-[#EBF3FB]" : ""}`}>
                              {cat && <div className={`absolute left-0 top-0 bottom-0 w-1 ${CATEGORY_COLORS[cat].dot}`} />}
                              <div className="flex items-start gap-2.5">
                                <input type="checkbox" checked={isSelected}
                                  onChange={e => { e.stopPropagation(); toggleSelectMail(mail.id); }}
                                  onClick={e => e.stopPropagation()}
                                  className="w-4 h-4 mt-1 rounded border-gray-300 shrink-0" />
                                <div className={`w-9 h-9 rounded-full ${getAvatarColor(mail.sender || mail.from || "")} flex items-center justify-center text-white text-xs font-bold shrink-0`}>
                                  {getInitials(mail.sender || mail.from || "?")}
                                </div>
                                <div className="flex-1 min-w-0">
                                  <div className="flex items-center justify-between gap-1 mb-0.5">
                                    <span className={`text-sm truncate ${!isRead ? "font-semibold" : ""} ${th.text}`}>
                                      {hl(mail.sender || mail.from || mail.toEmail || "No sender")}
                                    </span>
                                    <div className="flex items-center gap-1 shrink-0">
                                      {mail.importance === "high" && <AlertCircle size={12} className="text-red-500" />}
                                      {isFlagged && <Flag size={12} className="text-orange-400 fill-orange-400" />}
                                      {mail.attachments?.length > 0 && <Paperclip size={12} className={th.textMuted} />}
                                      <span className={`text-xs ${th.textMuted}`}>{fmtDateTime(mail.receivedDateTime || mail.sentDateTime)}</span>
                                    </div>
                                  </div>
                                  <div className="flex items-center gap-1.5 mb-0.5">
                                    {!isRead && <div className="w-2 h-2 bg-blue-500 rounded-full shrink-0" />}
                                    <span className={`text-sm truncate ${!isRead ? "font-semibold" : ""} ${th.text}`}>
                                      {hl(mail.subject || "(No subject)")}
                                    </span>
                                  </div>
                                  <p className={`text-xs ${th.textMuted} truncate`}>
                                    {hl(stripHtml(mail.body || "").slice(0, 100))}
                                  </p>
                                </div>
                              </div>
                            </div>
                          );
                        });
                      })()}
                    </div>
                  </div>

                  {/* ── Reading Pane ─────────────────────────────────── */}
                  {/* Mobile: full-screen when mobileView==="email" */}
                  <div className={`
                    ${(mobileView === "email" && openedMailId) || showCompose ? "flex" : "hidden"}
                    md:flex flex-1 overflow-hidden min-h-0
                  `}>
                    <div className={`flex flex-1 flex-col ${isDark ? "bg-[#1b1a19]" : "bg-[#F3F2F1]"} overflow-hidden min-w-0`}>
                      {showCompose ? (
                        <Email
                          inline
                          accessToken={accessToken}
                          onClose={() => { setShowCompose(false); setReplyData(null); }}
                          replyData={replyData}
                          onTemplatesOpenChange={setComposeTemplatesOpen}
                          onMailSent={mail => { handleMailSent(mail); setShowCompose(false); setReplyData(null); }}
                        />
                      ) : openedMail ? (
                        <>
                          {/* Mobile back button */}
                          <div className={`md:hidden px-3 pt-2 pb-1 ${isDark ? "bg-[#1b1a19]" : "bg-[#F3F2F1]"} shrink-0`}>
                            <button onClick={() => { setOpenedMailId(null); setMobileView("list"); }}
                              className="flex items-center gap-1.5 text-sm text-blue-600 font-medium py-1">
                              <ArrowLeft size={16} />Back
                            </button>
                          </div>

                          {/* ── Conversation Header ── */}
                          <div className={`${isDark ? "bg-[#252423] border-[#3d3b39]" : "bg-white border-[#E1E1E1]"} border-b px-3 md:px-5 pt-3 pb-0 shrink-0`}>
                            {/* Subject row */}
                            <div className="flex items-start gap-2 mb-2">
                              <h2 className={`flex-1 text-[14px] md:text-[15px] font-semibold leading-snug ${isDark ? "text-white" : "text-[#1f1f1f]"}`}>
                                {openedMail.subject || "(No subject)"}
                              </h2>
                              {mailCategories[openedMail.id] && (
                                <span className={`text-xs px-1.5 py-0.5 rounded font-medium shrink-0 mt-0.5 ${CATEGORY_COLORS[mailCategories[openedMail.id]].bg} ${CATEGORY_COLORS[mailCategories[openedMail.id]].text}`}>
                                  {CATEGORY_COLORS[mailCategories[openedMail.id]].label}
                                </span>
                              )}
                              <button onClick={() => confirmDelete([openedMail.id])}
                                className={`p-1.5 rounded ${isDark ? "text-gray-400 hover:text-red-400 hover:bg-[#3d3b39]" : "text-[#616161] hover:text-red-500 hover:bg-[#F3F2F1]"} transition shrink-0`} title="Delete">
                                <Trash2 size={15} />
                              </button>
                            </div>

                            {/* Participants + metadata */}
                            <div className="flex items-center gap-2 mb-2 flex-wrap">
                              {threadMessages.length > 1 ? (
                                <div className="flex -space-x-1.5 mr-1">
                                  {[...new Set(threadMessages.map(m => m.sender))].slice(0, 4).map((name, i) => (
                                    <div key={i} className={`w-6 h-6 rounded-full ${getAvatarColor(name)} border-2 ${isDark ? "border-[#252423]" : "border-white"} flex items-center justify-center text-white text-[9px] font-bold shrink-0`}>
                                      {getInitials(name || "?")}
                                    </div>
                                  ))}
                                </div>
                              ) : (
                                <div className={`w-6 h-6 rounded-full ${getAvatarColor(openedMail.sender || "")} flex items-center justify-center text-white text-[9px] font-bold shrink-0`}>
                                  {getInitials(openedMail.sender || openedMail.from || "?")}
                                </div>
                              )}
                              <span className={`text-xs truncate max-w-[55vw] md:max-w-none ${isDark ? "text-gray-300" : "text-[#424242]"}`}>
                                {threadMessages.length > 1
                                  ? [...new Set(threadMessages.map(m => m.sender))].slice(0, 3).join(", ") + (threadMessages.length > 3 ? ` +${threadMessages.length - 3}` : "")
                                  : (openedMail.sender || openedMail.from || "Unknown")}
                              </span>
                              {threadMessages.length > 1 && (
                                <span className={`text-xs shrink-0 ${isDark ? "text-gray-400" : "text-[#616161]"}`}>
                                  · {threadMessages.length} msgs
                                </span>
                              )}
                              <span className={`text-xs shrink-0 ml-auto ${isDark ? "text-gray-400" : "text-[#616161]"}`}>
                                {fmtDateTime(openedMail.receivedDateTime || openedMail.sentDateTime)}
                              </span>
                            </div>

                            {/* Action toolbar — horizontally scrollable on mobile */}
                            <div className="flex items-center overflow-x-auto -mx-1 pb-0.5 gap-0.5 scrollbar-none">
                              {(() => {
                                const btn = (onClick, icon, label, extra = "") => (
                                  <button onClick={onClick}
                                    className={`flex items-center gap-1 px-2 py-1.5 text-[13px] rounded transition whitespace-nowrap shrink-0 ${isDark ? "text-gray-300 hover:bg-[#3d3b39]" : "text-[#424242] hover:bg-[#EDEBE9]"} ${extra}`}>
                                    {icon}{label && <span className="hidden sm:inline">{label}</span>}
                                  </button>
                                );
                                return <>
                                  {btn(handleReply, <Reply size={14} />, "Reply")}
                                  {btn(handleReplyAll, <ReplyAll size={14} />, "Reply all")}
                                  {btn(handleForward, <Forward size={14} />, "Forward")}
                                  <div className={`w-px h-4 mx-0.5 shrink-0 ${isDark ? "bg-[#3d3b39]" : "bg-[#D1D1D1]"}`} />
                                  {activeFolder === "Inbox" && btn(() => archiveMail(openedMail.id), <Archive size={14} />, "Archive")}
                                  <button onClick={() => toggleFlag(openedMail.id)} title="Flag"
                                    className={`flex items-center gap-1 px-2 py-1.5 text-[13px] rounded transition shrink-0 ${flaggedIds[openedMail.id] ? "text-orange-500" : isDark ? "text-gray-300 hover:bg-[#3d3b39]" : "text-[#424242] hover:bg-[#EDEBE9]"}`}>
                                    <Flag size={14} /><span className="hidden sm:inline">{flaggedIds[openedMail.id] ? "Unflag" : "Flag"}</span>
                                  </button>
                                  {btn(() => toggleRead(openedMail), <MailOpen size={14} />, (readOverrides[openedMail.id] !== undefined ? readOverrides[openedMail.id] : openedMail.isRead) ? "Unread" : "Read")}
                                  <div className="relative shrink-0" onClick={e => e.stopPropagation()}>
                                    {btn(() => setShowCategoryMenu(p => p === openedMail.id ? null : openedMail.id), <Tag size={14} />, "Categorize")}
                                    {showCategoryMenu === openedMail.id && (
                                      <div className={`absolute left-0 top-full mt-0.5 w-44 ${isDark ? "bg-[#252423] border-[#3d3b39]" : "bg-white border-[#D1D1D1]"} border rounded shadow-lg z-30 py-1`}>
                                        {Object.entries(CATEGORY_COLORS).map(([key, { dot, label }]) => (
                                          <button key={key} onClick={() => setCategory(openedMail.id, key)}
                                            className={`w-full flex items-center gap-2 px-3 py-2 text-sm ${isDark ? "text-gray-200 hover:bg-[#3d3b39]" : "text-[#1f1f1f] hover:bg-[#F3F2F1]"}`}>
                                            <div className={`w-2.5 h-2.5 rounded-full ${dot}`} />
                                            {label}
                                            {mailCategories[openedMail.id] === key && <Check size={11} className="ml-auto text-blue-500" />}
                                          </button>
                                        ))}
                                      </div>
                                    )}
                                  </div>
                                  <div className="relative shrink-0" onClick={e => e.stopPropagation()}>
                                    {btn(() => setShowSnoozeMenu(p => p === openedMail.id ? null : openedMail.id), <Clock size={14} />, "Snooze")}
                                    {showSnoozeMenu === openedMail.id && (
                                      <div className={`absolute left-0 top-full mt-0.5 w-44 ${isDark ? "bg-[#252423] border-[#3d3b39]" : "bg-white border-[#D1D1D1]"} border rounded shadow-lg z-30 py-1 overflow-hidden`}>
                                        {["Tomorrow morning", "Tomorrow afternoon", "This weekend", "Next week", "In 2 weeks"].map(t => (
                                          <button key={t} onClick={() => { if (onToast) onToast(`Snoozed: ${t}`, "success"); setShowSnoozeMenu(null); }}
                                            className={`w-full px-3 py-2 text-sm text-left ${isDark ? "text-gray-200 hover:bg-[#3d3b39]" : "text-[#1f1f1f] hover:bg-[#F3F2F1]"}`}>{t}</button>
                                        ))}
                                      </div>
                                    )}
                                  </div>
                                  {btn(() => window.print(), <Printer size={14} />, null, "hidden md:flex")}
                                  <div className={`w-px h-4 mx-0.5 shrink-0 ${isDark ? "bg-[#3d3b39]" : "bg-[#D1D1D1]"} ml-auto`} />
                                  <button onClick={() => setShowCrmPanel(p => !p)}
                                    className={`flex items-center gap-1 px-2 py-1.5 text-[13px] rounded transition shrink-0 ${showCrmPanel ? "text-purple-600 bg-purple-50" : isDark ? "text-gray-300 hover:bg-[#3d3b39]" : "text-[#424242] hover:bg-[#EDEBE9]"}`}>
                                    <Activity size={14} /><span className="hidden sm:inline">CRM</span>
                                  </button>
                                </>;
                              })()}
                            </div>
                          </div>

                          {/* Scroll + body area */}
                          <div className="flex-1 overflow-y-auto">
                            {/* ── Inline reply compose ── */}
                            {inlineReply && (
                              <div className={`px-3 md:px-4 py-3 ${isDark ? "bg-[#252423] border-[#3d3b39]" : "bg-white border-[#E1E1E1]"} border-b`}>
                                <div className="flex items-center justify-between mb-3">
                                  <span className={`text-sm font-semibold ${th.text}`}>
                                    {inlineReply.type === "reply" ? "Reply" : inlineReply.type === "replyAll" ? "Reply All" : "Forward"}
                                  </span>
                                  <button onClick={() => { setInlineReply(null); setInlineBody(""); setInlineError(""); }}
                                    className={`p-1.5 rounded-lg ${th.hover} ${th.textMuted} transition`} title="Discard">
                                    <X size={14} />
                                  </button>
                                </div>

                                <div className={`flex items-center gap-2 px-3 py-2 border ${th.border} rounded-lg mb-2 ${isDark ? "bg-gray-800" : "bg-white"}`}>
                                  <span className={`text-xs font-semibold ${th.textMuted} w-6 flex-shrink-0`}>To</span>
                                  <input
                                    value={inlineToInput}
                                    onChange={e => setInlineToInput(e.target.value)}
                                    placeholder="Recipients…"
                                    className={`flex-1 text-sm bg-transparent outline-none ${th.text} placeholder:text-gray-400 min-w-0`}
                                  />
                                </div>

                                {inlineReply.ccEmail && (
                                  <div className={`flex items-center gap-2 px-3 py-2 border ${th.border} rounded-lg mb-2 ${isDark ? "bg-gray-800" : "bg-white"}`}>
                                    <span className={`text-xs font-semibold ${th.textMuted} w-6 flex-shrink-0`}>Cc</span>
                                    <span className={`text-sm ${th.textMuted} truncate min-w-0`}>{inlineReply.ccEmail}</span>
                                  </div>
                                )}

                                <textarea
                                  value={inlineBody}
                                  onChange={e => { setInlineBody(e.target.value); e.target.style.height = 'auto'; e.target.style.height = e.target.scrollHeight + 'px'; }}
                                  placeholder="Write your reply…"
                                  style={{ minHeight: '80px', overflow: 'hidden' }}
                                  className={`w-full px-3 py-2.5 text-sm border ${th.border} rounded-lg resize-none outline-none transition ${th.text} ${isDark ? "bg-gray-800" : "bg-white"} placeholder:text-gray-400`}
                                />

                                {inlineError && <p className="mt-1.5 text-xs text-red-500">{inlineError}</p>}

                                <div className="flex items-center gap-2 mt-3 flex-wrap">
                                  {/* Send split-button */}
                                  <div className="relative flex-shrink-0">
                                    <div className="flex rounded-lg overflow-hidden shadow-sm">
                                      <button onClick={handleInlineSend} disabled={inlineSending}
                                        className="flex items-center gap-2 px-4 py-2 bg-blue-500 hover:bg-blue-600 text-white text-sm font-medium transition disabled:opacity-50">
                                        <Send size={14} />{inlineSending ? "Sending…" : "Send"}
                                      </button>
                                      <button onClick={() => setInlineSendDropdown(p => !p)}
                                        className="flex items-center px-2 py-2 bg-blue-500 hover:bg-blue-600 text-white border-l border-slate-600 transition">
                                        <ChevronDown size={14} />
                                      </button>
                                    </div>
                                    {inlineSendDropdown && (
                                      <div className={`absolute left-0 bottom-full mb-1 w-48 border rounded-xl shadow-xl z-50 overflow-hidden ${isDark ? "bg-gray-800 border-gray-600" : "bg-white border-gray-200"}`}
                                        onMouseLeave={() => setInlineSendDropdown(false)}>
                                        <button onClick={() => { setInlineSendDropdown(false); handleInlineSend(); }}
                                          className={`w-full flex items-center gap-3 px-4 py-3 text-sm ${th.hover} ${th.text} transition`}>
                                          <Send size={14} className="text-blue-500 flex-shrink-0" />Send
                                        </button>
                                        <button onClick={() => { setInlineSendDropdown(false); setScheduleDateTime(""); setScheduleModal(true); }}
                                          className={`w-full flex items-center gap-3 px-4 py-3 text-sm ${th.hover} ${th.text} transition`}>
                                          <Clock size={14} className="text-blue-500 flex-shrink-0" />Schedule send
                                        </button>
                                        <button onClick={() => { setInlineSendDropdown(false); handleMailMerge(); }}
                                          className={`w-full flex items-center gap-3 px-4 py-3 text-sm ${th.hover} ${th.text} transition`}>
                                          <Users size={14} className="text-blue-500 flex-shrink-0" />Start mail merge
                                        </button>
                                      </div>
                                    )}
                                  </div>
                                  <button onClick={() => { setInlineReply(null); setInlineBody(""); setInlineError(""); setInlineSendDropdown(false); }}
                                    className={`flex items-center gap-2 px-4 py-2 text-sm font-medium ${th.textMuted} ${th.hover} rounded-lg transition`}>
                                    <Trash2 size={14} />Discard
                                  </button>
                                </div>
                              </div>
                            )}

                            {/* ── Email thread / single email ── */}
                            <div className="px-3 md:px-4 py-3 space-y-2">
                              {!openedMail.bodyLoaded && (
                                <div className="flex items-center gap-2 text-xs text-gray-400 py-2 animate-pulse">
                                  <RefreshCw size={12} className="animate-spin" />Loading…
                                </div>
                              )}

                              {threadMessages.length > 1 ? (
                                threadMessages.map((tm) => {
                                  const isCurrentMsg = tm.id === openedMail.id;
                                  const isExpanded = threadExpandedIds.has(tm.id);
                                  const bodyHtml = isCurrentMsg
                                    ? (openedMail.body || openedMail.bodyPreview || "")
                                    : (threadBodies[tm.id] || "");
                                  const cleanHtml = bodyHtml.replace(/http:/g, "https:").replace(/border-left\s*:\s*[^;'"<>]*(?:#[0-9a-fA-F]{3,8}|rgb[^;)]+\)|rgba[^;)]+\)|blue|navy|#3b82f6|cornflowerblue)[^;'"<>]*;?/gi, "border-left:2px solid #e5e7eb;");
                                  return (
                                    <div key={tm.id} className={`${isDark ? "bg-[#252423] border-[#3d3b39]" : "bg-white border-[#E1E1E1]"} border rounded overflow-hidden`}>
                                      <button
                                        onClick={() => isCurrentMsg ? null : toggleThreadMsg(tm.id)}
                                        className={`w-full flex items-center gap-3 px-3 py-3 text-left ${isCurrentMsg ? "cursor-default" : isDark ? "hover:bg-[#3d3b39]" : "hover:bg-[#F3F2F1]"} transition`}
                                      >
                                        <div className={`w-8 h-8 rounded-full ${getAvatarColor(tm.sender)} flex items-center justify-center text-white text-xs font-bold shrink-0`}>
                                          {getInitials(tm.sender || "?")}
                                        </div>
                                        <div className="flex-1 min-w-0">
                                          <div className="flex items-center justify-between gap-2">
                                            <span className={`text-[13px] font-semibold truncate ${isDark ? "text-white" : "text-[#1f1f1f]"}`}>{tm.sender}</span>
                                            <span className={`text-xs shrink-0 ${isDark ? "text-gray-400" : "text-[#616161]"}`}>{tm.receivedDateTime ? new Date(tm.receivedDateTime).toLocaleString([], { month: "short", day: "numeric", hour: "2-digit", minute: "2-digit" }) : ""}</span>
                                          </div>
                                          {!isExpanded && <p className={`text-xs truncate mt-0.5 ${isDark ? "text-gray-400" : "text-[#616161]"}`}>{tm.bodyPreview}</p>}
                                          {isExpanded && tm.toEmail && (
                                            <div className={`text-xs mt-0.5 ${isDark ? "text-gray-400" : "text-[#616161]"}`}>
                                              <span className="font-medium">To:</span> {tm.toEmail}
                                            </div>
                                          )}
                                        </div>
                                        {!isCurrentMsg && (
                                          <ChevronDown size={13} className={`shrink-0 ${isDark ? "text-gray-400" : "text-[#616161]"} transition-transform ${isExpanded ? "rotate-180" : ""}`} />
                                        )}
                                      </button>
                                      {isExpanded && (
                                        <div className={`border-t ${isDark ? "border-[#3d3b39]" : "border-[#E1E1E1]"}`}>
                                          {!bodyHtml && !isCurrentMsg ? (
                                            <p className={`text-xs ${isDark ? "text-gray-400" : "text-[#616161]"} py-3 px-4`}>Loading…</p>
                                          ) : (
                                            <EmailBodyFrame html={cleanHtml} isDark={isDark} />
                                          )}
                                        </div>
                                      )}
                                    </div>
                                  );
                                })
                              ) : (
                                <div className={`${isDark ? "bg-[#252423] border-[#3d3b39]" : "bg-white border-[#E1E1E1]"} border rounded overflow-hidden`}>
                                  {/* Email sub-header */}
                                  <div className={`px-3 md:px-4 pt-3 pb-2 border-b ${isDark ? "border-[#3d3b39]" : "border-[#E1E1E1]"}`}>
                                    <div className="flex items-start gap-3">
                                      <div className={`w-8 h-8 rounded-full ${getAvatarColor(openedMail.sender || "")} flex items-center justify-center text-white text-xs font-bold shrink-0`}>
                                        {getInitials(openedMail.sender || openedMail.from || "?")}
                                      </div>
                                      <div className="flex-1 min-w-0">
                                        <div className="flex items-center justify-between gap-2">
                                          <span className={`text-[13px] font-semibold truncate ${isDark ? "text-white" : "text-[#1f1f1f]"}`}>{openedMail.sender || openedMail.from || "Unknown"}</span>
                                          <span className={`text-xs shrink-0 ${isDark ? "text-gray-400" : "text-[#616161]"}`}>{fmtDateTime(openedMail.receivedDateTime || openedMail.sentDateTime)}</span>
                                        </div>
                                        {openedMail.senderEmail && (
                                          <div className={`text-xs truncate ${isDark ? "text-gray-400" : "text-[#616161]"}`}>{openedMail.senderEmail}</div>
                                        )}
                                        <div className={`text-xs mt-0.5 ${isDark ? "text-gray-400" : "text-[#616161]"}`}>
                                          {(() => {
                                            const list = (openedMail.toEmail || "").split(/[,;]/).map(e => e.trim()).filter(Boolean);
                                            const vis = toExpanded ? list : list.slice(0, 2);
                                            const extra = list.length - 2;
                                            return <><span className="font-medium">To:</span>{" "}{vis.join(", ")}
                                              {!toExpanded && extra > 0 && <button onClick={() => setToExpanded(true)} className="text-blue-500 ml-1">+{extra} more</button>}
                                            </>;
                                          })()}
                                        </div>
                                        {openedMail.ccEmail && (
                                          <div className={`text-xs mt-0.5 ${isDark ? "text-gray-400" : "text-[#616161]"}`}>
                                            <span className="font-medium">Cc:</span>{" "}
                                            {(() => {
                                              const list = openedMail.ccEmail.split(/[,;]/).map(e => e.trim()).filter(Boolean);
                                              const vis = ccExpanded ? list : list.slice(0, 2);
                                              const extra = list.length - 2;
                                              return <>{vis.join(", ")}{!ccExpanded && extra > 0 && <button onClick={() => setCcExpanded(true)} className="text-blue-500 ml-1">+{extra}</button>}</>;
                                            })()}
                                          </div>
                                        )}
                                      </div>
                                    </div>
                                  </div>
                                  {/* Body */}
                                  <EmailBodyFrame
                                    html={(openedMail.body || openedMail.bodyPreview || "").replace(/http:/g, "https:").replace(/border-left\s*:\s*[^;'"<>]*(?:#[0-9a-fA-F]{3,8}|rgb[^;)]+\)|rgba[^;)]+\)|blue|navy|#3b82f6|cornflowerblue)[^;'"<>]*;?/gi, "border-left:2px solid #e5e7eb;")}
                                    isDark={isDark}
                                  />
                                  {/* Attachments */}
                                  {openedMail.attachments?.length > 0 && (
                                    <div className={`px-3 md:px-4 pb-4 pt-2 border-t ${isDark ? "border-[#3d3b39]" : "border-[#E1E1E1]"}`}>
                                      <p className={`text-xs font-semibold mb-2 flex items-center gap-1.5 ${isDark ? "text-gray-300" : "text-[#424242]"}`}>
                                        <Paperclip size={12} />Attachments ({openedMail.attachments.length})
                                      </p>
                                      <div className="flex flex-wrap gap-2">
                                        {openedMail.attachments.map((att, i) => {
                                          const name = att.name || att.fileName || `attachment-${i}`;
                                          const size = att.size ? fmtFileSize(att.size) : "";
                                          return (
                                            <div key={name + i} className={`flex items-center gap-2 px-2.5 py-2 border ${isDark ? "border-[#3d3b39] bg-[#1b1a19] hover:bg-[#3d3b39]" : "border-[#E1E1E1] bg-white hover:bg-[#F3F2F1]"} rounded text-sm group transition max-w-[calc(50%-4px)]`}>
                                              <FileIcon name={name} size={14} />
                                              <div className="min-w-0">
                                                <p className={`truncate max-w-[100px] text-xs font-medium ${isDark ? "text-gray-200" : "text-[#1f1f1f]"}`}>{name}</p>
                                                {size && <p className={`text-xs ${isDark ? "text-gray-400" : "text-[#616161]"}`}>{size}</p>}
                                              </div>
                                              <div className="flex gap-1 opacity-0 group-hover:opacity-100 transition">
                                                {att.contentBytes ? (
                                                  <>
                                                    <a href={`data:${att.contentType || "application/octet-stream"};base64,${att.contentBytes}`}
                                                      download={name} title="Download"
                                                      className={`p-1 rounded ${isDark ? "text-gray-400 hover:text-blue-400" : "text-[#616161] hover:text-blue-600"}`}><Download size={13} /></a>
                                                    <button onClick={() => window.open(`data:${att.contentType};base64,${att.contentBytes}`, "_blank")} title="Preview"
                                                      className={`p-1 rounded ${isDark ? "text-gray-400 hover:text-green-400" : "text-[#616161] hover:text-green-600"}`}><Eye size={13} /></button>
                                                  </>
                                                ) : (
                                                  <span className={`p-1 ${isDark ? "text-gray-400" : "text-[#616161]"}`}><Download size={13} /></span>
                                                )}
                                              </div>
                                            </div>
                                          );
                                        })}
                                      </div>
                                    </div>
                                  )}
                                </div>
                              )}
                            </div>
                          </div>
                        </>
                      ) : (
                        <div className="flex-1 flex flex-col items-center justify-center">
                          <MailOpen size={56} className={`mb-4 ${th.textMuted} opacity-30`} />
                          <p className={`text-base font-medium ${th.textMuted}`}>Select an email to read</p>
                          <p className={`text-sm ${th.textMuted} mt-1 opacity-60`}>Nothing selected</p>
                        </div>
                      )}
                    </div>

                    {/* ── CRM Extension Panel (desktop only) ───────────── */}
                    {showCrmPanel && openedMail && (
                      <div className={`hidden md:flex w-64 xl:w-72 border-l ${th.border} ${th.surface} flex-col overflow-y-auto shrink-0`}>
                        <div className={`flex items-center justify-between px-4 py-3 border-b ${th.border}`}>
                          <div className="flex items-center gap-2">
                            <Activity size={15} className="text-purple-500" />
                            <span className="text-sm font-semibold">CRM</span>
                          </div>
                          <button onClick={() => setShowCrmPanel(false)} className={`p-1 rounded-lg ${th.hover} ${th.textMuted}`}><X size={15} /></button>
                        </div>
                        <div className={`flex border-b ${th.border}`}>
                          {["actions", "timeline"].map(t => (
                            <button key={t} onClick={() => setCrmTab(t)}
                              className={`flex-1 py-2 text-xs font-medium capitalize transition border-b-2 ${crmTab === t ? "border-purple-500 text-purple-600" : `border-transparent ${th.textMuted} ${th.hover}`}`}>
                              {t}
                            </button>
                          ))}
                        </div>
                        <div className="p-3 flex-1">
                          {crmTab === "actions" && (
                            <div className="space-y-2">
                              <p className={`text-xs font-semibold ${th.textMuted} uppercase tracking-wide mb-3`}>Quick Actions</p>
                              {[
                                { icon: Link, label: "Link to Contact", color: "text-blue-500", hover: isDark ? "hover:bg-blue-900/20" : "hover:bg-blue-50" },
                                { icon: UserPlus, label: "Create Lead", color: "text-green-500", hover: isDark ? "hover:bg-green-900/20" : "hover:bg-green-50" },
                                { icon: Tag, label: "Create Ticket", color: "text-orange-500", hover: isDark ? "hover:bg-orange-900/20" : "hover:bg-orange-50" },
                                { icon: CheckSquare, label: "Add Follow-up Task", color: "text-purple-500", hover: isDark ? "hover:bg-purple-900/20" : "hover:bg-purple-50" },
                                { icon: Calendar, label: "Schedule Meeting", color: "text-indigo-500", hover: isDark ? "hover:bg-indigo-900/20" : "hover:bg-indigo-50" },
                              ].map(({ icon: Icon, label, color, hover }) => (
                                <button key={label}
                                  onClick={() => { if (onToast) onToast(`${label} — coming soon`, "info"); }}
                                  className={`w-full flex items-center gap-2.5 px-3 py-2.5 border ${th.border} rounded-xl text-sm ${th.text} ${th.bg} ${hover} transition text-left`}>
                                  <Icon size={15} className={color} />{label}
                                </button>
                              ))}
                              <div className={`mt-4 pt-4 border-t ${th.border}`}>
                                <p className={`text-xs font-semibold ${th.textMuted} uppercase tracking-wide mb-2`}>Email Info</p>
                                <div className={`space-y-1.5 text-xs ${th.textMuted}`}>
                                  <div className="flex gap-2"><span className="font-medium w-10 shrink-0">From</span><span className="truncate">{openedMail.senderEmail || openedMail.sender || "—"}</span></div>
                                  <div className="flex gap-2"><span className="font-medium w-10 shrink-0">To</span><span className="truncate">{(openedMail.toEmail || "—").split(",")[0]}</span></div>
                                  <div className="flex gap-2"><span className="font-medium w-10 shrink-0">Date</span><span>{fmtDateTime(openedMail.receivedDateTime || openedMail.sentDateTime)}</span></div>
                                  {openedMail.attachments?.length > 0 && (
                                    <div className="flex gap-2"><span className="font-medium w-10 shrink-0">Files</span><span>{openedMail.attachments.length} attachment(s)</span></div>
                                  )}
                                </div>
                              </div>
                              {flaggedIds[openedMail.id] && (
                                <div className="mt-3 px-3 py-2 bg-orange-50 border border-orange-200 rounded-xl">
                                  <p className="text-xs text-orange-600 font-medium flex items-center gap-1.5"><Flag size={12} className="fill-orange-500" />Flagged for follow-up</p>
                                </div>
                              )}
                            </div>
                          )}
                          {crmTab === "timeline" && (
                            <div>
                              <p className={`text-xs font-semibold ${th.textMuted} uppercase tracking-wide mb-3`}>Activity History</p>
                              <div className="space-y-3">
                                {[
                                  { icon: Mail, label: "Email received", time: fmtDateTime(openedMail.receivedDateTime || openedMail.sentDateTime), color: "bg-blue-100 text-blue-600" },
                                  { icon: User, label: "Contact viewed", time: "Just now", color: "bg-gray-100 text-gray-600" },
                                ].map((item, i) => (
                                  <div key={i} className="flex gap-2.5">
                                    <div className={`w-7 h-7 rounded-full ${item.color} flex items-center justify-center shrink-0 mt-0.5`}>
                                      <item.icon size={13} />
                                    </div>
                                    <div>
                                      <p className={`text-xs font-medium ${th.text}`}>{item.label}</p>
                                      <p className={`text-xs ${th.textMuted}`}>{item.time}</p>
                                    </div>
                                  </div>
                                ))}
                              </div>
                            </div>
                          )}
                        </div>
                      </div>
                    )}
                  </div>
                </>
              )}

              {/* ── Templates section ────────────────────────────────── */}
              {activeFolder === "Templates" && (
                <div className="flex flex-1 overflow-hidden">
                  {/* Template list */}
                  <div className={`${openedTemplateIdx !== null || showAddTemplate ? "hidden md:flex" : "flex"} w-full md:w-80 border-r ${th.border} flex-col ${th.bg}`}>
                    <div className={`h-12 border-b ${th.border} px-4 flex items-center justify-between ${th.surface} shrink-0`}>
                      <h2 className="text-sm font-semibold">Templates</h2>
                      <button onClick={() => { setShowAddTemplate(true); setOpenedTemplateIdx(null); setNewTemplate({ name: "", subject: "", body: "", footer: "", assignedEmail: loggedInEmail, logo: "", attachments: [] }); }}
                        className="flex items-center gap-1.5 bg-blue-500 hover:bg-blue-600 text-white px-3 py-1.5 rounded-lg text-xs transition">
                        <Plus size={13} />New
                      </button>
                    </div>
                    <div className="flex-1 overflow-y-auto">
                      {!myTemplates.length ? (
                        <div className="flex flex-col items-center justify-center h-full py-12 px-6 text-center">
                          <div className={`w-14 h-14 rounded-full flex items-center justify-center mb-3 ${isDark ? "bg-gray-700" : "bg-gray-100"}`}>
                            <Edit2 size={22} className={`${th.textMuted} opacity-70`} />
                          </div>
                          <p className={`text-sm font-medium ${th.text}`}>No templates yet</p>
                          <p className={`text-xs mt-1 ${th.textMuted}`}>Create a template to reuse it later</p>
                        </div>
                      ) : myTemplates.map((tpl, idx) => (
                        <div key={tpl.name + idx}
                          onClick={() => { setOpenedTemplateIdx(idx); setShowAddTemplate(false); }}
                          className={`border-b ${th.border} px-4 py-3 cursor-pointer transition ${th.hover} ${openedTemplateIdx === idx && !showAddTemplate ? isDark ? "bg-slate-700/30 border-l-2 border-l-slate-400" : "bg-slate-50 border-l-2 border-l-slate-700" : ""}`}>
                          <p className={`text-sm font-medium truncate ${th.text}`}>{tpl.name || tpl.body || "Untitled template"}</p>
                        </div>
                      ))}
                    </div>
                  </div>
                  {/* Template detail */}
                  <div className={`${openedTemplateIdx !== null || showAddTemplate ? "flex" : "hidden md:flex"} flex-1 flex-col ${th.bg} overflow-hidden`}>
                    {/* Mobile back button for templates */}
                    {(openedTemplateIdx !== null || showAddTemplate) && (
                      <div className="md:hidden px-4 pt-3 shrink-0">
                        <button onClick={() => { setOpenedTemplateIdx(null); setShowAddTemplate(false); }}
                          className="flex items-center gap-1.5 text-sm text-blue-600 font-medium py-1">
                          <ArrowLeft size={15} />Templates
                        </button>
                      </div>
                    )}
                    {showAddTemplate ? (
                      <>
                        <div className={`h-12 border-b ${th.border} px-4 md:px-6 flex items-center ${th.surface} shrink-0`}>
                          <h2 className="text-sm font-semibold">New Template</h2>
                        </div>
                        <div className="flex-1 overflow-y-auto p-4 md:p-6">
                          <form onSubmit={handleAddTemplate} className="max-w-2xl space-y-5">
                            <div>
                              <label className={`block text-sm font-medium ${th.textMuted} mb-1.5`}>Assigned to</label>
                              <div className={`px-4 py-2.5 border rounded-xl text-sm ${th.input} opacity-80`}>{loggedInEmail || "you"}</div>
                              <p className={`text-xs ${th.textMuted} mt-1`}>Saved under your account — only you can see and use it, and it auto-loads on new email.</p>
                            </div>
                            <div>
                              <label className={`block text-sm font-medium ${th.textMuted} mb-1.5`}>Subject</label>
                              <input type="text" value={newTemplate.subject}
                                onChange={e => setNewTemplate(p => ({ ...p, subject: e.target.value }))}
                                placeholder="Email subject..."
                                className={`w-full px-4 py-2.5 border rounded-xl text-sm ${th.input} focus:outline-none focus:ring-2 focus:ring-blue-500`} />
                            </div>
                            <div>
                              <label className={`block text-sm font-medium ${th.textMuted} mb-1.5`}>Body</label>
                              <textarea value={newTemplate.body}
                                onChange={e => setNewTemplate(p => ({ ...p, body: e.target.value }))}
                                rows={6}
                                placeholder="Email body..."
                                className={`w-full px-4 py-2.5 border rounded-xl text-sm ${th.input} resize-y focus:outline-none focus:ring-2 focus:ring-blue-500`} />
                            </div>
                            <div>
                              <label className={`block text-sm font-medium ${th.textMuted} mb-1.5`}>Attachments (optional)</label>
                              {(newTemplate.attachments || []).length > 0 && (
                                <div className="space-y-1.5 mb-2">
                                  {newTemplate.attachments.map((a, i) => (
                                    <div key={i} className={`flex items-center justify-between gap-2 px-3 py-2 border rounded-lg text-sm ${th.input}`}>
                                      <span className="truncate flex items-center gap-2 min-w-0"><Paperclip size={14} className="shrink-0" /><span className="truncate">{a.name}</span></span>
                                      <button type="button" onClick={() => setNewTemplate(p => ({ ...p, attachments: (p.attachments || []).filter((_, idx) => idx !== i) }))} className="text-xs text-red-600 hover:underline shrink-0">Remove</button>
                                    </div>
                                  ))}
                                </div>
                              )}
                              <label className={`inline-flex items-center gap-2 px-4 py-2.5 border rounded-xl text-sm cursor-pointer ${th.input} ${th.hover}`}>
                                <Paperclip size={16} /><span>Add files</span>
                                <input type="file" multiple className="hidden" onChange={(e) => { addTemplateAttachments(e.target.files, setNewTemplate, newTemplate.attachments); e.target.value = ""; }} />
                              </label>
                              <p className={`text-xs ${th.textMuted} mt-1`}>Attached to the email when this template is used (max ~4 MB total).</p>
                            </div>
                            <div>
                              <label className={`block text-sm font-medium ${th.textMuted} mb-1.5`}>Footer</label>
                              <textarea value={newTemplate.footer}
                                onChange={e => setNewTemplate(p => ({ ...p, footer: e.target.value }))}
                                rows={4}
                                placeholder="Footer content (signature, company info)..."
                                className={`w-full px-4 py-2.5 border rounded-xl text-sm ${th.input} resize-y focus:outline-none focus:ring-2 focus:ring-blue-500`} />
                            </div>
                            <div>
                              <label className={`block text-sm font-medium ${th.textMuted} mb-1.5`}>Company Logo (optional)</label>
                              {newTemplate.logo ? (
                                <div className="flex items-center gap-3">
                                  <img src={newTemplate.logo} alt="Company logo" className="h-14 max-w-[180px] object-contain border rounded-lg p-1 bg-white" />
                                  <button type="button" onClick={() => setNewTemplate(p => ({ ...p, logo: "" }))}
                                    className="px-3 py-1.5 text-xs rounded-lg border border-red-200 text-red-600 hover:bg-red-50">Remove</button>
                                </div>
                              ) : (
                                <label className={`inline-flex items-center gap-2 px-4 py-2.5 border rounded-xl text-sm cursor-pointer ${th.input} ${th.hover}`}>
                                  <ImageIcon size={16} />
                                  <span>Upload logo</span>
                                  <input type="file" accept="image/*" className="hidden"
                                    onChange={e => {
                                      const file = e.target.files && e.target.files[0];
                                      if (!file) return;
                                      if (file.size > 512 * 1024) { showPopup("Logo too large (max 512 KB)", "error"); return; }
                                      const reader = new FileReader();
                                      reader.onload = () => setNewTemplate(p => ({ ...p, logo: String(reader.result || "") }));
                                      reader.readAsDataURL(file);
                                    }} />
                                </label>
                              )}
                              <p className={`text-xs ${th.textMuted} mt-1`}>Appears in the email footer when this template is used.</p>
                            </div>
                            <div className="flex gap-3 flex-wrap">
                              <button type="submit" disabled={isSubmitting}
                                className="px-5 py-2.5 bg-blue-500 hover:bg-blue-600 text-white text-sm rounded-xl transition disabled:opacity-50">
                                {isSubmitting ? "Creating..." : "Create Template"}
                              </button>
                              <button type="button" onClick={() => setShowAddTemplate(false)}
                                className={`px-5 py-2.5 border ${th.border} text-sm rounded-xl ${th.hover} ${th.text}`}>Cancel</button>
                            </div>
                          </form>
                        </div>
                      </>
                    ) : (() => {
                      const tpl = myTemplates[openedTemplateIdx];
                      if (editTemplateIdx === openedTemplateIdx && tpl) return (
                        <>
                          <div className={`h-12 border-b ${th.border} px-4 md:px-6 flex items-center ${th.surface} shrink-0`}>
                            <h2 className="text-sm font-semibold">Edit Template</h2>
                          </div>
                          <div className="flex-1 overflow-y-auto p-4 md:p-6">
                            <form onSubmit={async e => {
                              e.preventDefault(); setIsSubmitting(true);
                              try {
                                await apiClient.put(`/Template/${tpl.templateId ?? tpl.TemplateId}`, { name: (editTemplate.subject || editTemplate.body || editTemplate.footer || "").trim().slice(0, 40) || "Template", subject: editTemplate.subject || "", body: buildTemplateBody(editTemplate.body, editTemplate.footer, editTemplate.logo, editTemplate.attachments), createdBy: (editTemplate.assignedEmail || loggedInEmail || "").trim(), isActive: true });
                                apiClient.get("/Template").then(r => setApiTemplates(Array.isArray(r.data) ? r.data : [])).catch(() => { });
                                setEditTemplateIdx(null);
                                showPopup("Template updated");
                              } catch (err) {
                                const status = err?.response?.status;
                                const tooBig = status === 413 || status === 400;
                                const msg = tooBig
                                  ? "Template too large to save — the attachment exceeds the server upload limit. Use a smaller file."
                                  : (err?.response?.data?.title || err?.response?.data?.message ||
                                     (typeof err?.response?.data === "string" ? err.response.data : "") ||
                                     err?.message || "Could not update template");
                                showPopup(String(msg).slice(0, 220), "error");
                              }
                              setIsSubmitting(false);
                            }} className="max-w-2xl space-y-5">
                              <div>
                                <label className={`block text-sm font-medium ${th.textMuted} mb-1.5`}>Assigned to</label>
                                <div className={`px-4 py-2.5 border rounded-xl text-sm ${th.input} opacity-80`}>{editTemplate.assignedEmail || loggedInEmail || "you"}</div>
                              </div>
                              <div>
                                <label className={`block text-sm font-medium ${th.textMuted} mb-1.5`}>Subject</label>
                                <input type="text" value={editTemplate.subject}
                                  onChange={e => setEditTemplate(p => ({ ...p, subject: e.target.value }))}
                                  placeholder="Email subject..."
                                  className={`w-full px-4 py-2.5 border rounded-xl text-sm ${th.input} focus:outline-none focus:ring-2 focus:ring-blue-500`} />
                              </div>
                              <div>
                                <label className={`block text-sm font-medium ${th.textMuted} mb-1.5`}>Body</label>
                                <textarea value={editTemplate.body}
                                  onChange={e => setEditTemplate(p => ({ ...p, body: e.target.value }))} rows={6}
                                  placeholder="Email body..."
                                  className={`w-full px-4 py-2.5 border rounded-xl text-sm ${th.input} resize-y focus:outline-none focus:ring-2 focus:ring-blue-500`} />
                              </div>
                              <div>
                                <label className={`block text-sm font-medium ${th.textMuted} mb-1.5`}>Attachments (optional)</label>
                                {(editTemplate.attachments || []).length > 0 && (
                                  <div className="space-y-1.5 mb-2">
                                    {editTemplate.attachments.map((a, i) => (
                                      <div key={i} className={`flex items-center justify-between gap-2 px-3 py-2 border rounded-lg text-sm ${th.input}`}>
                                        <span className="truncate flex items-center gap-2 min-w-0"><Paperclip size={14} className="shrink-0" /><span className="truncate">{a.name}</span></span>
                                        <button type="button" onClick={() => setEditTemplate(p => ({ ...p, attachments: (p.attachments || []).filter((_, idx) => idx !== i) }))} className="text-xs text-red-600 hover:underline shrink-0">Remove</button>
                                      </div>
                                    ))}
                                  </div>
                                )}
                                <label className={`inline-flex items-center gap-2 px-4 py-2.5 border rounded-xl text-sm cursor-pointer ${th.input} ${th.hover}`}>
                                  <Paperclip size={16} /><span>Add files</span>
                                  <input type="file" multiple className="hidden" onChange={(e) => { addTemplateAttachments(e.target.files, setEditTemplate, editTemplate.attachments); e.target.value = ""; }} />
                                </label>
                                <p className={`text-xs ${th.textMuted} mt-1`}>Attached to the email when this template is used (max ~4 MB total).</p>
                              </div>
                              <div>
                                <label className={`block text-sm font-medium ${th.textMuted} mb-1.5`}>Footer</label>
                                <textarea value={editTemplate.footer}
                                  onChange={e => setEditTemplate(p => ({ ...p, footer: e.target.value }))} rows={4}
                                  placeholder="Footer content (signature, company info)..."
                                  className={`w-full px-4 py-2.5 border rounded-xl text-sm ${th.input} resize-y focus:outline-none focus:ring-2 focus:ring-blue-500`} />
                              </div>
                              <div>
                                <label className={`block text-sm font-medium ${th.textMuted} mb-1.5`}>Company Logo (optional)</label>
                                {editTemplate.logo ? (
                                  <div className="flex items-center gap-3">
                                    <img src={editTemplate.logo} alt="Company logo" className="h-14 max-w-[180px] object-contain border rounded-lg p-1 bg-white" />
                                    <button type="button" onClick={() => setEditTemplate(p => ({ ...p, logo: "" }))}
                                      className="px-3 py-1.5 text-xs rounded-lg border border-red-200 text-red-600 hover:bg-red-50">Remove</button>
                                  </div>
                                ) : (
                                  <label className={`inline-flex items-center gap-2 px-4 py-2.5 border rounded-xl text-sm cursor-pointer ${th.input} ${th.hover}`}>
                                    <ImageIcon size={16} />
                                    <span>Upload logo</span>
                                    <input type="file" accept="image/*" className="hidden"
                                      onChange={e => {
                                        const file = e.target.files && e.target.files[0];
                                        if (!file) return;
                                        if (file.size > 512 * 1024) { showPopup("Logo too large (max 512 KB)", "error"); return; }
                                        const reader = new FileReader();
                                        reader.onload = () => setEditTemplate(p => ({ ...p, logo: String(reader.result || "") }));
                                        reader.readAsDataURL(file);
                                      }} />
                                  </label>
                                )}
                              </div>
                              <div className="flex gap-3 flex-wrap">
                                <button type="submit" disabled={isSubmitting}
                                  className="px-5 py-2.5 bg-blue-500 hover:bg-blue-600 text-white text-sm rounded-xl transition disabled:opacity-50">
                                  {isSubmitting ? "Saving..." : "Save Changes"}
                                </button>
                                <button type="button" onClick={() => setEditTemplateIdx(null)}
                                  className={`px-5 py-2.5 border ${th.border} text-sm rounded-xl ${th.hover} ${th.text}`}>Cancel</button>
                              </div>
                            </form>
                          </div>
                        </>
                      );
                      return tpl ? (
                        <>
                          <div className={`h-12 border-b ${th.border} px-4 md:px-6 flex items-center justify-between ${th.surface} shrink-0`}>
                            <h2 className="text-sm font-semibold truncate">{tpl.name || tpl.body || "Untitled template"}</h2>
                            <div className="flex gap-2 shrink-0">
                              <button onClick={() => { const parsed = parseTemplateBody(tpl.body); setEditTemplateIdx(openedTemplateIdx); setEditTemplate({ subject: tpl.subject ?? tpl.Subject ?? "", body: parsed.body, footer: parsed.footer, assignedEmail: tpl.createdBy ?? tpl.CreatedBy ?? "", logo: parsed.logo, attachments: parsed.attachments || [] }); }}
                                className={`flex items-center gap-1.5 px-3 py-1.5 text-xs rounded-lg ${th.hover} ${th.text}`}><Pencil size={13} />Edit</button>
                              <button onClick={() => setTemplateToDelete(tpl)}
                                className="flex items-center gap-1.5 px-3 py-1.5 text-xs rounded-lg hover:bg-red-50 text-red-500">
                                <Trash2 size={13} />Delete
                              </button>
                            </div>
                          </div>
                          <div className="flex-1 overflow-y-auto p-4 md:p-6">
                            <div className="max-w-2xl space-y-5">
                              <div>
                                <label className={`block text-xs font-semibold ${th.textMuted} uppercase tracking-wide mb-2`}>Assigned Email</label>
                                <div className={`px-4 py-3 border ${th.border} rounded-xl text-sm ${th.text} ${isDark ? "bg-gray-800" : "bg-gray-50"}`}>{(tpl.createdBy ?? tpl.CreatedBy) || "— not assigned —"}</div>
                              </div>
                              {(() => { const { body, footer, logo } = parseTemplateBody(tpl.body); const subj = tpl.subject ?? tpl.Subject ?? ""; return (
                              <>
                              {subj && (
                                <div>
                                  <label className={`block text-xs font-semibold ${th.textMuted} uppercase tracking-wide mb-2`}>Subject</label>
                                  <div className={`px-4 py-3 border ${th.border} rounded-xl text-sm ${th.text} ${isDark ? "bg-gray-800" : "bg-gray-50"}`}>{subj}</div>
                                </div>
                              )}
                              {body && (
                                <div>
                                  <label className={`block text-xs font-semibold ${th.textMuted} uppercase tracking-wide mb-2`}>Body</label>
                                  <div className={`px-4 py-4 border ${th.border} rounded-xl text-sm ${th.text} whitespace-pre-wrap min-h-24 ${isDark ? "bg-gray-800" : "bg-gray-50"}`}>{body}</div>
                                </div>
                              )}
                              <div>
                                <label className={`block text-xs font-semibold ${th.textMuted} uppercase tracking-wide mb-2`}>Footer</label>
                                  <div className={`flex gap-4 items-start px-4 py-4 border ${th.border} rounded-xl text-sm ${th.text} min-h-32 ${isDark ? "bg-gray-800" : "bg-gray-50"}`}>
                                    {logo && <img src={logo} alt="Company logo" className="h-16 max-w-[160px] object-contain shrink-0" />}
                                    <div className="whitespace-pre-wrap flex-1 min-w-0">{footer}</div>
                                  </div>
                              </div>
                              </>
                              ); })()}
                              <button onClick={() => { startCompose({ type: "new", toEmail: "", subject: tpl.subject ?? tpl.Subject ?? "", body: tpl.body }); }}
                                className="flex items-center gap-2 px-4 py-2.5 bg-blue-500 hover:bg-blue-600 text-white text-sm rounded-xl transition">
                                <Mail size={14} />Use Template
                              </button>
                            </div>
                          </div>
                        </>
                      ) : (
                        <div className="flex-1 flex flex-col items-center justify-center">
                          <Edit2 size={48} className={`mb-4 ${th.textMuted} opacity-30`} />
                          <p className={`text-base ${th.textMuted}`}>Select a template</p>
                        </div>
                      );
                    })()}
                  </div>
                </div>
              )}
            </div>
          </>
        )}

        {/* ── Calendar ─────────────────────────────────────────────────── */}
        {activeApp === "Calendar" && (() => {
          const today = new Date(_today);
          const eventsForDay = (d) => calEvents.filter(e => e.date === calDateKey(d));
          const monthGrid = getMonthGrid(calDate.getFullYear(), calDate.getMonth());
          const miniGrid = getMonthGrid(miniCalDate.getFullYear(), miniCalDate.getMonth());
          const evColor = (c) => ({ blue: "bg-blue-500", green: "bg-green-600", red: "bg-red-500", orange: "bg-orange-500", purple: "bg-purple-500", teal: "bg-teal-500" }[c] || "bg-blue-500");

          return (
            <div className="flex flex-1 overflow-hidden">

              {/* ── Left: Mini calendar (desktop only) ─────────────────── */}
              <aside className={`hidden md:flex flex-col w-60 border-r ${th.border} ${th.surface} shrink-0`}>
                <div className="p-3">
                  <div className="flex items-center justify-between mb-2 px-1">
                    <button onClick={() => setMiniCalDate(d => new Date(d.getFullYear(), d.getMonth() - 1, 1))}
                      className={`p-1 rounded ${th.hover} ${th.textMuted}`}><ChevronLeft size={14} /></button>
                    <span className={`text-xs font-semibold ${th.text}`}>
                      {miniCalDate.toLocaleDateString("en-US", { year: "numeric", month: "long" })}
                    </span>
                    <button onClick={() => setMiniCalDate(d => new Date(d.getFullYear(), d.getMonth() + 1, 1))}
                      className={`p-1 rounded ${th.hover} ${th.textMuted}`}><ChevronRight size={14} /></button>
                  </div>
                  <div className="grid grid-cols-7 mb-1">
                    {["S", "M", "T", "W", "T", "F", "S"].map((d, i) => (
                      <div key={i} className={`text-center text-[10px] font-semibold ${th.textMuted} py-0.5`}>{d}</div>
                    ))}
                  </div>
                  <div className="grid grid-cols-7 gap-y-0.5">
                    {miniGrid.map((cell, i) => {
                      const isToday = cell.date.toDateString() === today.toDateString();
                      const isSel = cell.date.toDateString() === selectedCalDay.toDateString();
                      return (
                        <button key={i}
                          onClick={() => { setSelectedCalDay(cell.date); setCalDate(new Date(cell.date.getFullYear(), cell.date.getMonth(), 1)); setMiniCalDate(new Date(cell.date.getFullYear(), cell.date.getMonth(), 1)); }}
                          className={`flex items-center justify-center w-7 h-7 mx-auto rounded-full text-[11px] transition font-medium
                            ${isToday ? "bg-blue-500 text-white" :
                              isSel ? isDark ? "bg-gray-600 text-white" : "bg-gray-200 text-gray-800" :
                                !cell.curr ? th.textMuted + " opacity-40" :
                                  `${th.hover} ${th.text}`}`}>
                          {cell.date.getDate()}
                        </button>
                      );
                    })}
                  </div>
                </div>

                <div className={`mx-3 border-t ${th.border}`} />

                <button className="flex items-center gap-2 px-4 py-2.5 text-sm text-blue-500 hover:text-blue-600 transition">
                  <Plus size={14} />Add calendar
                </button>

                <div className={`mx-3 border-t ${th.border}`} />

                <div className="flex-1 overflow-y-auto p-2 space-y-1">
                  {(accounts.length ? accounts.slice(0, 2) : [{ username: "My Calendar", name: "Me" }]).map((acc, i) => (
                    <div key={i} className="mb-1">
                      <button className={`w-full flex items-center gap-1.5 px-2 py-1.5 rounded ${th.hover} text-xs font-semibold ${th.textMuted}`}>
                        <ChevronRight size={11} />
                        <span className="truncate max-w-[150px]">{acc.username || acc.name}</span>
                      </button>
                      <div className="ml-4 space-y-0.5">
                        {[{ label: "Calendar", color: "bg-blue-500" }, { label: "Birthdays", color: "bg-green-500" }, { label: "Holidays", color: "bg-teal-500" }].map(({ label, color }) => (
                          <div key={label} className={`flex items-center gap-2 px-2 py-1 rounded ${th.hover} cursor-pointer`}>
                            <div className={`w-3 h-3 rounded-sm ${color} flex items-center justify-center shrink-0`}>
                              <Check size={8} className="text-white" />
                            </div>
                            <span className={`text-xs ${th.text}`}>{label}</span>
                          </div>
                        ))}
                        <button className="px-2 py-0.5 text-xs text-blue-500 hover:underline block">Show all</button>
                      </div>
                    </div>
                  ))}
                </div>
              </aside>

              {/* ── Main calendar area ────────────────────────────────── */}
              <div className={`flex flex-1 flex-col overflow-hidden ${th.bg}`}>

                {/* Top toolbar */}
                <div className={`border-b ${th.border} px-2 md:px-3 py-2 flex items-center gap-1.5 md:gap-2 flex-wrap shrink-0`}>
                  <button
                    onClick={() => { setEventFormDate(selectedCalDay); setShowEventForm(true); }}
                    className="flex items-center gap-1.5 px-3 py-1.5 bg-blue-500 hover:bg-blue-600 text-white text-sm rounded-lg font-medium transition shrink-0">
                    <Plus size={14} />New event <ChevronDown size={11} />
                  </button>
                  <div className={`hidden md:block w-px h-5 ${isDark ? "bg-gray-600" : "bg-gray-300"}`} />
                  {/* View switcher — scrollable on mobile */}
                  <div className={`flex border ${th.border} rounded-lg overflow-hidden text-xs md:text-sm shrink-0`}>
                    {["Day", "Week", "Month"].map(v => (
                      <button key={v} onClick={() => setCalView(v)}
                        className={`px-2 md:px-3 py-1.5 font-medium transition border-r last:border-r-0 ${th.border} ${calView === v ? isDark ? "bg-gray-600 text-white" : "bg-gray-100 text-gray-900" : `${th.text} ${th.hover}`}`}>
                        {v}
                      </button>
                    ))}
                    <button onClick={() => setCalView("Work week")}
                      className={`hidden md:block px-3 py-1.5 font-medium transition border-r last:border-r-0 ${th.border} ${calView === "Work week" ? isDark ? "bg-gray-600 text-white" : "bg-gray-100 text-gray-900" : `${th.text} ${th.hover}`}`}>
                      Work week
                    </button>
                  </div>
                  <button
                    onClick={() => setCalDate(d => new Date(d.getFullYear(), d.getMonth(), 1))}
                    title="Refresh events"
                    className={`flex items-center gap-1 px-2.5 py-1.5 text-xs md:text-sm border ${th.border} rounded-lg ${th.hover} ${th.text} ${calLoading ? "opacity-50 cursor-not-allowed" : ""} shrink-0`}>
                    <RefreshCw size={13} className={calLoading ? "animate-spin" : ""} />
                    <span className="hidden md:inline">Refresh</span>
                  </button>
                </div>

                {/* Sub-header: Today + navigation */}
                <div className={`border-b ${th.border} px-2 md:px-4 py-2 flex items-center gap-2 shrink-0`}>
                  <button
                    onClick={() => { setCalDate(new Date(today.getFullYear(), today.getMonth(), 1)); setSelectedCalDay(today); setMiniCalDate(new Date(today.getFullYear(), today.getMonth(), 1)); }}
                    className={`flex items-center gap-1 px-2.5 py-1.5 text-xs md:text-sm border ${th.border} rounded-lg ${th.hover} ${th.text} font-medium shrink-0`}>
                    <Calendar size={13} />Today
                  </button>
                  <div className="flex items-center shrink-0">
                    <button onClick={() => setCalDate(d => new Date(d.getFullYear(), d.getMonth() - 1, 1))}
                      className={`p-1.5 rounded-lg ${th.hover} ${th.textMuted}`}><ChevronLeft size={16} /></button>
                    <button onClick={() => setCalDate(d => new Date(d.getFullYear(), d.getMonth() + 1, 1))}
                      className={`p-1.5 rounded-lg ${th.hover} ${th.textMuted}`}><ChevronRight size={16} /></button>
                  </div>
                  <h2 className={`text-sm md:text-lg font-semibold ${th.text} truncate`}>
                    {calDate.toLocaleDateString("en-US", { year: "numeric", month: "long" })}
                  </h2>
                </div>

                {/* ── Month grid ────────────────────────────────────── */}
                {calView === "Month" && (
                  <div className="flex-1 overflow-auto flex flex-col">
                    <div className={`grid grid-cols-7 border-b ${th.border} shrink-0 sticky top-0 z-10 ${th.surface}`}>
                      {/* Abbreviated on mobile */}
                      {["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"].map((d, i) => (
                        <div key={i} className={`py-1.5 text-center text-[10px] md:text-xs font-semibold ${th.textMuted} border-r last:border-r-0 ${th.border}`}>
                          <span className="md:hidden">{d}</span>
                          <span className="hidden md:inline">{["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"][i]}</span>
                        </div>
                      ))}
                    </div>
                    <div className="flex-1 grid grid-rows-6">
                      {Array.from({ length: 6 }, (_, wk) => {
                        const week = monthGrid.slice(wk * 7, wk * 7 + 7);
                        return (
                          <div key={wk} className={`grid grid-cols-7 border-b last:border-b-0 ${th.border}`} style={{ minHeight: 60 }}>
                            {week.map((cell, di) => {
                              const isToday = cell.date.toDateString() === today.toDateString();
                              const isSel = cell.date.toDateString() === selectedCalDay.toDateString();
                              const dayEvs = eventsForDay(cell.date);
                              return (
                                <div key={di}
                                  onClick={() => { setSelectedCalDay(cell.date); }}
                                  onDoubleClick={() => { setSelectedCalDay(cell.date); setEventFormDate(cell.date); setShowEventForm(true); }}
                                  className={`border-r last:border-r-0 ${th.border} p-0.5 md:p-1 cursor-pointer transition group ${!cell.curr ? "opacity-40" : ""} ${isSel && !isToday ? isDark ? "bg-gray-800" : "bg-blue-50/50" : isDark ? "hover:bg-gray-800" : "hover:bg-gray-50"}`}>
                                  <div className="mb-0.5">
                                    <span className={`inline-flex items-center justify-center w-6 h-6 md:w-7 md:h-7 text-xs md:text-sm rounded-full font-medium
                                      ${isToday ? "bg-blue-500 text-white font-bold" :
                                        isSel ? isDark ? "bg-gray-600 text-white" : "bg-blue-100 text-blue-700" :
                                          th.text}`}>
                                      {cell.date.getDate()}
                                    </span>
                                  </div>
                                  <div className="space-y-0.5">
                                    {dayEvs.slice(0, 2).map(ev => (
                                      <div key={ev.id}
                                        onClick={e => { e.stopPropagation(); setSelectedEvent(ev); }}
                                        className={`text-[9px] md:text-xs px-1 py-0.5 rounded text-white font-medium truncate cursor-pointer hover:opacity-90 transition ${evColor(ev.color)}`}>
                                        <span className="hidden md:inline opacity-80 mr-0.5 text-[10px]">{ev.start}</span>{ev.title}
                                      </div>
                                    ))}
                                    {dayEvs.length > 2 && (
                                      <p className={`text-[9px] md:text-[11px] ${th.textMuted} px-0.5`}>+{dayEvs.length - 2}</p>
                                    )}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}

                {/* ── Day view ──────────────────────────────────────── */}
                {calView === "Day" && (
                  <div className="flex-1 overflow-auto">
                    <div className={`px-3 py-3 border-b ${th.border} text-sm font-semibold ${th.text}`}>
                      {selectedCalDay.toLocaleDateString("en-US", { weekday: "long", year: "numeric", month: "long", day: "numeric" })}
                    </div>
                    <div className="relative" style={{ minHeight: 1200 }}>
                      {Array.from({ length: 24 }, (_, h) => (
                        <div key={h} className={`flex border-b ${th.border}`} style={{ height: 50 }}>
                          <div className={`w-14 md:w-16 shrink-0 text-right pr-2 md:pr-3 pt-1 text-[10px] md:text-xs ${th.textMuted}`}>
                            {h === 0 ? "" : `${h % 12 || 12} ${h < 12 ? "AM" : "PM"}`}
                          </div>
                          <div className="flex-1 relative" />
                        </div>
                      ))}
                      {eventsForDay(selectedCalDay).map(ev => {
                        const [sh, sm] = ev.start.split(":").map(Number);
                        const [eh, em] = (ev.end || ev.start).split(":").map(Number);
                        const top = (sh * 60 + sm) / 60 * 50;
                        const height = Math.max(((eh * 60 + em) - (sh * 60 + sm)) / 60 * 50, 25);
                        return (
                          <div key={ev.id}
                            onClick={() => setSelectedEvent(ev)}
                            className={`absolute left-16 right-2 rounded-lg px-2 py-1 text-white text-xs font-medium shadow cursor-pointer hover:opacity-90 transition ${evColor(ev.color)}`}
                            style={{ top, height }}>
                            <p className="truncate font-semibold">{ev.title}</p>
                            <p className="opacity-80 text-[10px]">{ev.start} – {ev.end || ""}</p>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}

                {/* ── Week / Work week view ─────────────────────────── */}
                {(calView === "Week" || calView === "Work week") && (
                  <div className="flex-1 overflow-auto flex flex-col">
                    {(() => {
                      const dow = selectedCalDay.getDay();
                      const weekStart = calAddDays(selectedCalDay, -dow);
                      const cols = calView === "Work week" ? [1, 2, 3, 4, 5].map(i => calAddDays(weekStart, i)) : Array.from({ length: 7 }, (_, i) => calAddDays(weekStart, i));
                      return (
                        <>
                          <div className={`grid border-b ${th.border} shrink-0`} style={{ gridTemplateColumns: `48px repeat(${cols.length}, 1fr)` }}>
                            <div className="border-r border-transparent" />
                            {cols.map((d, i) => {
                              const isTod = d.toDateString() === today.toDateString();
                              return (
                                <div key={i} className={`py-1.5 text-center border-r last:border-r-0 ${th.border}`}>
                                  <p className={`text-[10px] font-medium ${th.textMuted}`}>{d.toLocaleDateString("en-US", { weekday: "short" })}</p>
                                  <span className={`inline-flex items-center justify-center w-6 h-6 rounded-full text-xs font-semibold mx-auto mt-0.5 ${isTod ? "bg-blue-500 text-white" : th.text}`}>
                                    {d.getDate()}
                                  </span>
                                </div>
                              );
                            })}
                          </div>
                          <div className="flex-1 overflow-auto relative" style={{ minHeight: 1200 }}>
                            {Array.from({ length: 24 }, (_, h) => (
                              <div key={h} className={`flex border-b ${th.border}`} style={{ height: 50 }}>
                                <div className={`w-12 shrink-0 text-right pr-1.5 pt-1 text-[9px] md:text-[10px] ${th.textMuted}`}>
                                  {h > 0 ? `${h % 12 || 12}${h < 12 ? "a" : "p"}` : ""}
                                </div>
                                {cols.map((_, ci) => <div key={ci} className={`flex-1 border-r last:border-r-0 ${th.border}`} />)}
                              </div>
                            ))}
                            {cols.map((d, ci) => eventsForDay(d).map(ev => {
                              const [sh, sm] = ev.start.split(":").map(Number);
                              const [eh, em] = (ev.end || ev.start).split(":").map(Number);
                              const top = (sh * 60 + sm) / 60 * 50;
                              const height = Math.max(((eh * 60 + em) - (sh * 60 + sm)) / 60 * 50, 25);
                              return (
                                <div key={ev.id}
                                  onClick={() => setSelectedEvent(ev)}
                                  className={`absolute rounded px-1 py-0.5 text-white text-[10px] font-medium shadow cursor-pointer hover:opacity-90 transition ${evColor(ev.color)}`}
                                  style={{ top, height, left: `calc(${48}px + ${ci * (100 / cols.length)}%)`, width: `calc(${100 / cols.length}% - 2px)` }}>
                                  <p className="truncate">{ev.title}</p>
                                </div>
                              );
                            }))}
                          </div>
                        </>
                      );
                    })()}
                  </div>
                )}
              </div>

              {/* ── Calendar loading overlay ─────────────────────── */}
              {calLoading && (
                <div className="absolute inset-0 flex items-center justify-center bg-black/20 z-30 pointer-events-none">
                  <div className={`flex items-center gap-2.5 px-4 py-2.5 rounded-xl shadow-lg ${isDark ? "bg-gray-800 text-gray-200" : "bg-white text-gray-700"} text-sm font-medium`}>
                    <RefreshCw size={15} className="animate-spin text-blue-500" />
                    Loading events…
                  </div>
                </div>
              )}

              {/* ── Event detail popup ────────────────────────────── */}
              {selectedEvent && (
                <div className="fixed inset-0 z-50 p-4 flex items-center justify-center" onClick={() => setSelectedEvent(null)}>
                  <div
                    className={`w-full max-w-sm ${th.card} border ${th.border} rounded-xl shadow-xl p-5`}
                    onClick={e => e.stopPropagation()}>
                    <div className="flex items-start gap-3 mb-4">
                      <div className={`w-3 h-3 rounded-sm mt-1.5 shrink-0 ${evColor(selectedEvent.color)}`} />
                      <div className="flex-1 min-w-0">
                        <h3 className={`font-semibold ${th.text} leading-snug`}>{selectedEvent.title}</h3>
                        <p className={`text-xs ${th.textMuted} mt-1`}>
                          {new Date(selectedEvent.date + "T00:00:00").toLocaleDateString("en-US", { weekday: "long", year: "numeric", month: "long", day: "numeric" })}
                        </p>
                        <p className={`text-xs ${th.textMuted}`}>{selectedEvent.start}{selectedEvent.end ? ` – ${selectedEvent.end}` : ""}</p>
                      </div>
                      <button onClick={() => setSelectedEvent(null)} className={`p-1 rounded ${th.hover} ${th.textMuted} shrink-0`}><X size={15} /></button>
                    </div>
                    <div className="flex justify-between items-center pt-3 border-t border-gray-200/30">
                      <button
                        onClick={() => handleDeleteEvent(selectedEvent)}
                        className="flex items-center gap-1.5 px-3 py-2 text-sm text-red-500 hover:bg-red-50 rounded-lg transition">
                        <Trash2 size={14} />Delete
                      </button>
                      <button onClick={() => setSelectedEvent(null)}
                        className={`px-3 py-2 text-sm border ${th.border} rounded-lg ${th.hover} ${th.text}`}>
                        Close
                      </button>
                    </div>
                  </div>
                </div>
              )}

              {/* ── New Event Form Modal ──────────────────────────── */}
              {showEventForm && (() => {
                const evDate = eventFormDate || selectedCalDay;
                const evDateLabel = evDate ? evDate.toLocaleDateString("en-US", { weekday: "short", year: "numeric", month: "2-digit", day: "2-digit" }).replace(/\//g, "-") : "";
                const busyOptions = ["Free", "Working elsewhere", "Tentative", "Busy", "Out of office"];
                const reminderOptions = ["Don't remind me", "At time of event", "5 minutes before", "15 minutes before", "30 minutes before", "1 hour before", "2 hours before", "12 hours before", "1 day before", "1 week before"];
                const categoryOptions = [
                  { label: "Blue category", color: "#0078D4" }, { label: "Green category", color: "#107C10" },
                  { label: "Orange category", color: "#CA5010" }, { label: "Purple category", color: "#5C2D91" },
                  { label: "Red category", color: "#D13438" }, { label: "Yellow category", color: "#C19C00" },
                ];
                const busyDotColor = (s) => s === "Free" ? "bg-green-400" : s === "Working elsewhere" ? "bg-yellow-400" : s === "Tentative" ? "bg-blue-400" : s === "Out of office" ? "bg-purple-400" : "bg-[#6264A7]";
                const closeAllDd = () => { setShowEvBusyDd(false); setShowEvReminderDd(false); setShowEvCategoryDd(false); setShowEvPrivacyDd(false); };
                const [sh, sm] = newEventStart.split(":").map(Number);
                const [eh, em] = newEventEnd.split(":").map(Number);
                const gridStart = 7;
                const topPx = Math.max(0, ((sh - gridStart) * 60 + sm) / 60 * 48);
                const heightPx = Math.max(28, ((eh * 60 + em) - (sh * 60 + sm)) / 60 * 48);
                return (
                  <div className="fixed inset-0 bg-black/75 z-50 flex items-end md:items-center justify-center p-0 md:p-4" onClick={() => setShowEventForm(false)}>
                    {/* Full screen on mobile, modal on desktop */}
                    <div className="bg-[#1f1f2e] border-t md:border border-white/10 rounded-t-2xl md:rounded-xl shadow-xl w-full md:max-w-5xl flex flex-col overflow-hidden"
                      style={{ height: "92vh" }}
                      onClick={e => e.stopPropagation()}>

                      {/* ── Top toolbar ── */}
                      <div className="flex items-center justify-between px-3 py-2 bg-[#141424] border-b border-white/10 shrink-0">
                        <div className="flex items-center gap-0.5 overflow-x-auto">
                          <button onClick={handleSaveEvent}
                            disabled={!newEventTitle.trim()}
                            className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-md mr-1 transition-colors shrink-0 ${newEventTitle.trim() ? "bg-blue-500 hover:bg-blue-600 text-white" : "bg-blue-500/30 text-white/35 cursor-not-allowed"}`}>
                            <Bookmark size={13} />Save
                          </button>
                          <button className="flex items-center gap-1.5 px-3 py-1.5 bg-white/10 border border-white/20 text-white text-xs font-semibold rounded-md shrink-0">
                            <Calendar size={13} className="text-blue-400" />Event
                          </button>
                          <div className="w-px h-5 bg-white/15 mx-1 shrink-0" />
                          {/* Busy dropdown */}
                          <div className="relative shrink-0">
                            <button onClick={() => { closeAllDd(); setShowEvBusyDd(p => !p); }}
                              className="flex items-center gap-1 px-2 py-1.5 text-white/60 hover:bg-white/8 text-xs rounded-md transition-colors">
                              <span className={`w-2 h-2 rounded-full shrink-0 ${busyDotColor(evBusyStatus)}`} />
                              <span className="hidden sm:inline">{evBusyStatus}</span><ChevronDown size={11} />
                            </button>
                            {showEvBusyDd && (
                              <><div className="fixed inset-0 z-40" onClick={() => setShowEvBusyDd(false)} />
                                <div className="absolute top-full left-0 mt-1 w-48 bg-[#1e1f38] border border-white/15 rounded-lg shadow-xl z-50 py-1">
                                  {busyOptions.map(opt => (
                                    <button key={opt} onClick={() => { setEvBusyStatus(opt); setShowEvBusyDd(false); }}
                                      className="w-full flex items-center gap-3 px-3 py-2.5 text-xs text-white/80 hover:bg-white/10 transition-colors">
                                      <span className={`w-2 h-2 rounded-full shrink-0 ${busyDotColor(opt)}`} />
                                      <span className="flex-1">{opt}</span>
                                      {evBusyStatus === opt && <Check size={12} className="text-blue-400" />}
                                    </button>
                                  ))}
                                </div></>
                            )}
                          </div>
                          <div className="relative shrink-0">
                            <button onClick={() => { closeAllDd(); setShowEvReminderDd(p => !p); }}
                              className="flex items-center gap-0.5 px-2 py-1.5 text-white/45 hover:bg-white/8 rounded-md transition-colors">
                              <Bell size={13} /><ChevronDown size={11} />
                            </button>
                            {showEvReminderDd && (
                              <><div className="fixed inset-0 z-40" onClick={() => setShowEvReminderDd(false)} />
                                <div className="absolute top-full left-0 mt-1 w-52 bg-[#1e1f38] border border-white/15 rounded-lg shadow-xl z-50 py-1 max-h-64 overflow-y-auto">
                                  {reminderOptions.map(opt => (
                                    <button key={opt} onClick={() => { setEvReminder(opt); setShowEvReminderDd(false); }}
                                      className="w-full flex items-center gap-3 px-3 py-2.5 text-xs text-white/80 hover:bg-white/10 transition-colors">
                                      <span className="w-3.5 shrink-0 flex justify-center">{evReminder === opt && <Check size={12} className="text-blue-400" />}</span>
                                      {opt}
                                    </button>
                                  ))}
                                </div></>
                            )}
                          </div>
                          <div className="relative shrink-0">
                            <button onClick={() => { closeAllDd(); setShowEvCategoryDd(p => !p); }}
                              className="flex items-center gap-0.5 px-2 py-1.5 text-white/45 hover:bg-white/8 rounded-md transition-colors">
                              <Tag size={13} /><ChevronDown size={11} />
                            </button>
                            {showEvCategoryDd && (
                              <><div className="fixed inset-0 z-40" onClick={() => setShowEvCategoryDd(false)} />
                                <div className="absolute top-full left-0 mt-1 w-52 bg-[#1e1f38] border border-white/15 rounded-lg shadow-xl z-50 py-1">
                                  {categoryOptions.map(({ label, color }) => (
                                    <button key={label} onClick={() => { setEvCategory(c => c === label ? null : label); setShowEvCategoryDd(false); }}
                                      className="w-full flex items-center gap-3 px-3 py-2.5 text-xs text-white/80 hover:bg-white/10 transition-colors">
                                      <span style={{ color }} className="shrink-0"><Tag size={13} /></span>
                                      <span className="flex-1">{label}</span>
                                      {evCategory === label && <Check size={12} className="text-blue-400" />}
                                    </button>
                                  ))}
                                </div></>
                            )}
                          </div>
                          <div className="relative shrink-0">
                            <button onClick={() => { closeAllDd(); setShowEvPrivacyDd(p => !p); }}
                              className="flex items-center gap-0.5 px-2 py-1.5 text-white/45 hover:bg-white/8 rounded-md transition-colors">
                              <Lock size={13} /><ChevronDown size={11} />
                            </button>
                            {showEvPrivacyDd && (
                              <><div className="fixed inset-0 z-40" onClick={() => setShowEvPrivacyDd(false)} />
                                <div className="absolute top-full left-0 mt-1 w-40 bg-[#1e1f38] border border-white/15 rounded-lg shadow-xl z-50 py-1">
                                  {["Private", "Not private"].map(opt => (
                                    <button key={opt} onClick={() => { setEvPrivacy(opt); setShowEvPrivacyDd(false); }}
                                      className="w-full flex items-center gap-3 px-3 py-2.5 text-xs text-white/80 hover:bg-white/10 transition-colors">
                                      <span className="w-3.5 shrink-0 flex justify-center">{evPrivacy === opt && <Check size={12} className="text-blue-400" />}</span>
                                      {opt}
                                    </button>
                                  ))}
                                </div></>
                            )}
                          </div>
                        </div>
                        <button onClick={() => setShowEventForm(false)} className="p-2 hover:bg-white/10 rounded-md text-white/35 transition-colors ml-2 shrink-0">
                          <X size={15} />
                        </button>
                      </div>

                      {/* ── Body ── */}
                      <div className="flex flex-1 overflow-hidden min-h-0">
                        {/* LEFT: form */}
                        <div className="flex-1 overflow-y-auto px-4 md:px-7 py-4 md:py-5">
                          {/* Calendar selector */}
                          <div className="flex items-center gap-2 mb-4 pb-3 border-b border-white/8">
                            <span className="w-3 h-3 rounded-full bg-blue-500 shrink-0" />
                            <span className="text-sm text-white/70 truncate">
                              Calendar ({accounts[0]?.username || "your account"})
                            </span>
                            <ChevronDown size={13} className="text-white/35 shrink-0" />
                          </div>

                          {/* Title */}
                          <div className="border-b border-white/20 mb-4 pb-1">
                            <input autoFocus type="text" placeholder="Add title"
                              value={newEventTitle}
                              onChange={e => setNewEventTitle(e.target.value)}
                              className="w-full bg-transparent text-xl md:text-[22px] font-light text-white placeholder-white/25 focus:outline-none" />
                          </div>

                          {/* Attendees */}
                          <div className="relative border-b border-white/8 px-1 py-2">
                            <div className="flex items-start gap-3">
                              <Users size={18} className="text-white/35 shrink-0 mt-2" />
                              <div className="flex-1 min-w-0">
                                {evSelectedAttendees.length > 0 && (
                                  <div className="flex flex-wrap gap-1.5 mb-2 pt-1">
                                    {evSelectedAttendees.map(p => {
                                      const avatarColors = ["bg-blue-600", "bg-emerald-600", "bg-rose-600", "bg-amber-600", "bg-violet-600", "bg-cyan-600", "bg-pink-600", "bg-teal-600"];
                                      const ac = avatarColors[(p.name || p.email).charCodeAt(0) % avatarColors.length];
                                      const ini = (p.name || p.email).split(" ").map(w => w[0]).join("").slice(0, 2).toUpperCase();
                                      return (
                                        <span key={p.email} className="flex items-center gap-1.5 bg-blue-500/15 border border-blue-500/30 text-white text-xs px-2 py-1 rounded-full max-w-[160px]">
                                          <span className={`w-4 h-4 rounded-full ${ac} flex items-center justify-center text-[8px] font-bold shrink-0`}>{ini}</span>
                                          <span className="truncate">{p.name || p.email}</span>
                                          <button onMouseDown={e => { e.preventDefault(); setEvSelectedAttendees(a => a.filter(x => x.email !== p.email)); }} className="text-white/35 hover:text-white/80 shrink-0 ml-0.5">
                                            <X size={10} />
                                          </button>
                                        </span>
                                      );
                                    })}
                                  </div>
                                )}
                                <input type="text"
                                  placeholder={evSelectedAttendees.length ? "Add more attendees..." : "Invite required attendees"}
                                  value={evAttendeeSearch}
                                  onChange={e => { setEvAttendeeSearch(e.target.value); setShowEvAttendeePicker(true); }}
                                  onFocus={() => setShowEvAttendeePicker(true)}
                                  onBlur={() => setTimeout(() => setShowEvAttendeePicker(false), 150)}
                                  className="w-full bg-transparent text-sm text-white placeholder-white/30 focus:outline-none py-1" />
                              </div>
                              <AtSign size={15} className="text-white/25 shrink-0 mt-2" />
                            </div>

                            {showEvAttendeePicker && (() => {
                              const q = evAttendeeSearch.toLowerCase().trim();
                              const pool = outlookPeople.filter(p => !evSelectedAttendees.some(a => a.email.toLowerCase() === p.email.toLowerCase()));
                              const suggestions = q
                                ? pool.filter(p => p.name.toLowerCase().includes(q) || p.email.toLowerCase().includes(q))
                                : pool.slice(0, 12);
                              const avatarColors = ["bg-blue-600", "bg-emerald-600", "bg-rose-600", "bg-amber-600", "bg-violet-600", "bg-cyan-600", "bg-pink-600", "bg-teal-600", "bg-orange-600", "bg-indigo-600"];
                              const addAttendee = p => {
                                setEvSelectedAttendees(a => [...a, { name: p.name, email: p.email }]);
                                setEvAttendeeSearch("");
                                setShowEvAttendeePicker(true);
                              };
                              return (
                                <div className="absolute left-0 right-0 top-full mt-1 bg-[#1a1b2e] border border-white/12 rounded-xl shadow-xl z-50 overflow-hidden" style={{ maxHeight: 240, overflowY: "auto" }}>
                                  <div className="flex items-center justify-between px-4 py-2 border-b border-white/8 bg-[#15162a]">
                                    <span className="text-[10px] font-semibold text-white/40 uppercase tracking-wider">
                                      {q ? "Search results" : "Suggested from Outlook"}
                                    </span>
                                  </div>
                                  {!q && outlookPeople.length === 0 && (
                                    <div className="flex items-center justify-center py-4 gap-2">
                                      <div className="w-4 h-4 border-2 border-blue-500 border-t-transparent rounded-full animate-spin" />
                                      <span className="text-xs text-white/35">Fetching contacts…</span>
                                    </div>
                                  )}
                                  {q && suggestions.length === 0 && (
                                    <div className="px-4 py-4 text-center">
                                      <p className="text-xs text-white/40">No contacts match "{evAttendeeSearch}"</p>
                                    </div>
                                  )}
                                  {suggestions.map((p, idx) => {
                                    const ini = p.name.split(" ").map(w => w[0]).join("").slice(0, 2).toUpperCase() || "?";
                                    const ac = avatarColors[p.name.charCodeAt(0) % avatarColors.length];
                                    return (
                                      <button key={p.email + idx}
                                        onMouseDown={e => { e.preventDefault(); addAttendee(p); }}
                                        className="w-full flex items-center gap-3 px-4 py-2.5 hover:bg-white/8 transition-colors text-left">
                                        <div className={`w-8 h-8 rounded-full ${ac} flex items-center justify-center text-xs font-bold text-white shrink-0`}>
                                          {ini}
                                        </div>
                                        <div className="flex-1 min-w-0">
                                          <p className="text-sm text-white font-medium truncate">{p.name}</p>
                                          <p className="text-[11px] text-white/45 truncate">{p.email}</p>
                                        </div>
                                      </button>
                                    );
                                  })}
                                </div>
                              );
                            })()}
                          </div>

                          {/* Date / Time */}
                          <div className="flex items-center gap-3 py-3 border-b border-white/8 px-1 flex-wrap">
                            <Clock size={18} className="text-white/35 shrink-0" />
                            <div className="flex items-center gap-2 flex-wrap text-sm text-white">
                              <span className="text-white/60 text-xs">{evDateLabel}</span>
                              {!newEventAllDay && (
                                <>
                                  <input type="time" value={newEventStart}
                                    onChange={e => setNewEventStart(e.target.value)}
                                    className="bg-transparent text-sm text-white focus:outline-none [color-scheme:dark]" />
                                  <span className="text-white/30">-</span>
                                  <input type="time" value={newEventEnd}
                                    onChange={e => setNewEventEnd(e.target.value)}
                                    className="bg-transparent text-sm text-white focus:outline-none [color-scheme:dark]" />
                                </>
                              )}
                            </div>
                            <label className="flex items-center gap-2 cursor-pointer shrink-0 ml-auto">
                              <div onClick={() => setNewEventAllDay(p => !p)}
                                className={`w-9 h-5 rounded-full transition-colors relative cursor-pointer ${newEventAllDay ? "bg-blue-500" : "bg-white/20"}`}>
                                <div className={`absolute top-0.5 w-4 h-4 bg-white rounded-full shadow transition-transform ${newEventAllDay ? "translate-x-4" : "translate-x-0.5"}`} />
                              </div>
                              <span className="text-xs text-white/50">All day</span>
                            </label>
                          </div>

                          {/* Location */}
                          <div className="flex items-center gap-3 py-3 border-b border-white/8 px-1">
                            <MapPin size={18} className="text-white/35 shrink-0" />
                            <input type="text" placeholder="Add a room or location"
                              value={newEventLocation}
                              onChange={e => setNewEventLocation(e.target.value)}
                              className="flex-1 bg-transparent text-sm text-white placeholder-white/30 focus:outline-none min-w-0" />
                          </div>

                          {/* Teams meeting toggle */}
                          <div className="flex items-center gap-3 py-3 border-b border-white/8 px-1">
                            <Video size={18} className="text-white/35 shrink-0" />
                            <div onClick={() => setNewEventOnline(p => !p)}
                              className={`w-9 h-5 rounded-full transition-colors relative cursor-pointer ${newEventOnline ? "bg-blue-500" : "bg-white/20"}`}>
                              <div className={`absolute top-0.5 w-4 h-4 bg-white rounded-full shadow transition-transform ${newEventOnline ? "translate-x-4" : "translate-x-0.5"}`} />
                            </div>
                            <span className="text-sm text-white/70 font-medium">Teams meeting</span>
                          </div>

                          {/* Description */}
                          <div className="mt-4 border border-white/10 rounded-xl overflow-hidden">
                            <textarea rows={5} placeholder="Add a description or notes..."
                              value={newEventDescription}
                              onChange={e => setNewEventDescription(e.target.value)}
                              className="w-full bg-[#1a1a2e] px-4 py-3 text-sm text-white placeholder-white/25 focus:outline-none resize-none" />
                          </div>

                          {/* Agenda */}
                          <div className="flex items-center gap-3 py-3 mt-3 border-t border-white/8 px-1">
                            <FileText size={18} className="text-white/35 shrink-0" />
                            <input type="text" placeholder="Add an agenda"
                              value={newEventAgenda}
                              onChange={e => setNewEventAgenda(e.target.value)}
                              className="flex-1 bg-transparent text-sm text-white placeholder-white/30 focus:outline-none min-w-0" />
                          </div>
                        </div>

                        {/* RIGHT: Day preview (desktop only) */}
                        <div className="hidden md:flex w-64 border-l border-white/10 flex-col shrink-0 bg-[#14142a]">
                          <div className="flex items-center justify-between px-3 py-2.5 border-b border-white/10 shrink-0">
                            <span className="text-[11px] text-white/60 font-medium">
                              {evDate ? evDate.toLocaleDateString("en-US", { weekday: "short", day: "2-digit", month: "short", year: "numeric" }) : ""}
                            </span>
                            <button className="p-1 hover:bg-white/10 rounded text-white/35 transition-colors"><ExternalLink size={13} /></button>
                          </div>
                          <div className="flex-1 overflow-y-auto relative">
                            {Array.from({ length: 14 }, (_, i) => {
                              const h = i + 7;
                              const h12 = h === 12 ? 12 : h > 12 ? h - 12 : h;
                              const ap = h < 12 ? "AM" : "PM";
                              return (
                                <div key={h} className="flex border-b border-white/5" style={{ height: 48 }}>
                                  <div className="w-14 text-right pr-2 pt-1 text-[9px] text-white/30 shrink-0 leading-tight select-none">{h12} {ap}</div>
                                  <div className="flex-1 border-l border-white/5" />
                                </div>
                              );
                            })}
                            {!newEventAllDay && (
                              <div className="absolute left-14 right-1 rounded-md bg-[#c84b4b] px-2 py-1 text-white text-[11px] font-medium shadow-lg flex flex-col justify-center"
                                style={{ top: topPx, height: heightPx }}>
                                <span className="truncate font-semibold leading-tight">{newEventTitle || "New event"}</span>
                                <span className="truncate opacity-80 leading-tight">{newEventStart} – {newEventEnd}</span>
                              </div>
                            )}
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                );
              })()}
            </div>
          );
        })()}

        {/* ── People stub ───────────────────────────────────────────────── */}
        {activeApp === "People" && (
          <div className={`flex-1 flex flex-col items-center justify-center ${th.bg}`}>
            <Users size={64} className={`mb-4 ${th.textMuted} opacity-30`} />
            <p className={`text-xl font-semibold ${th.text}`}>People</p>
            <p className={`text-sm ${th.textMuted} mt-2`}>Your contacts are available in the Contacts module.</p>
          </div>
        )}

        {/* ── Tasks stub ────────────────────────────────────────────────── */}
        {activeApp === "Tasks" && (
          <div className={`flex-1 flex flex-col items-center justify-center ${th.bg}`}>
            <CheckSquare size={64} className={`mb-4 ${th.textMuted} opacity-30`} />
            <p className={`text-xl font-semibold ${th.text}`}>Tasks</p>
            <p className={`text-sm ${th.textMuted} mt-2`}>Manage your tasks and to-do items here.</p>
          </div>
        )}

        {/* ── Files stub ────────────────────────────────────────────────── */}
        {activeApp === "Files" && (
          <div className={`flex-1 flex flex-col items-center justify-center ${th.bg}`}>
            <FolderOpen size={64} className={`mb-4 ${th.textMuted} opacity-30`} />
            <p className={`text-xl font-semibold ${th.text}`}>Files</p>
            <p className={`text-sm ${th.textMuted} mt-2`}>Access your OneDrive files and email attachments.</p>
          </div>
        )}
      </div>

      {/* Compose renders inline in the reading pane (beside the mail list) for email
          folders only — see the showCompose branch there. No modal overlay, so opening
          Templates (or any non-mail view) shows that page, not a New Message panel. */}

      {/* ── Schedule Send Modal ──────────────────────────────────────────── */}
      {scheduleModal && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-end md:items-center justify-center z-[1300] p-0 md:p-4">
          <div className={`${th.card} border ${th.border} rounded-t-2xl md:rounded-xl shadow-xl w-full md:max-w-sm overflow-hidden`}>
            <div className={`px-5 py-4 border-b ${th.border} flex items-center justify-between`}>
              <h3 className={`text-base font-semibold ${th.text} flex items-center gap-2`}>
                <Clock size={16} className="text-blue-500" />Schedule send
              </h3>
              <button onClick={() => setScheduleModal(false)} className={`p-2 rounded-lg ${th.hover} ${th.textMuted} transition`}>
                <X size={14} />
              </button>
            </div>
            <div className="px-5 py-4">
              <label className={`text-xs font-medium ${th.textMuted} mb-2 block`}>Select date and time</label>
              <input
                type="datetime-local"
                value={scheduleDateTime}
                onChange={e => setScheduleDateTime(e.target.value)}
                min={new Date(Date.now() + 60000).toISOString().slice(0, 16)}
                className={`w-full px-3 py-2.5 border ${th.border} rounded-lg text-sm ${th.text} ${isDark ? "bg-gray-800" : "bg-white"} outline-none focus:border-transparent focus:ring-2 focus:ring-blue-500`}
              />
              {inlineError && <p className="mt-2 text-xs text-red-500">{inlineError}</p>}
            </div>
            <div className={`px-5 py-4 border-t ${th.border} flex items-center justify-end gap-2`}>
              <button onClick={() => { setScheduleModal(false); setInlineError(""); }}
                className={`px-4 py-2.5 text-sm font-medium ${th.textMuted} ${th.hover} rounded-lg transition`}>
                Cancel
              </button>
              <button onClick={handleScheduleSend} disabled={!scheduleDateTime || schedulingInProgress}
                className="flex items-center gap-2 px-4 py-2.5 bg-blue-500 hover:bg-blue-600 text-white text-sm font-medium rounded-lg transition disabled:opacity-50">
                <Clock size={14} />{schedulingInProgress ? "Scheduling…" : "Schedule"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── Mobile Bottom Navigation ─────────────────────────────────────── */}
      <nav className={`md:hidden flex items-center justify-around border-t ${th.border} ${th.surface} shrink-0 z-20`} style={{ height: 56, paddingBottom: "env(safe-area-inset-bottom)" }}>
        {[
          { key: "Mail", icon: Mail, badge: folderCounts.Inbox },
          { key: "Calendar", icon: Calendar, badge: 0 },
          { key: "People", icon: Users, badge: 0 },
          { key: "Tasks", icon: CheckSquare, badge: 0 },
          { key: "Files", icon: FolderOpen, badge: 0 },
        ].map(({ key, icon: Icon, badge }) => (
          <button key={key} onClick={() => setActiveApp(key)}
            className={`relative flex flex-col items-center justify-center gap-0.5 flex-1 h-full transition ${activeApp === key ? "text-blue-600" : th.textMuted}`}>
            <Icon size={20} />
            <span className="text-[10px] font-medium">{key}</span>
            {badge > 0 && (
              <span className="absolute top-2 right-[calc(50%-16px)] w-4 h-4 bg-red-500 text-white text-[9px] rounded-full flex items-center justify-center font-bold">
                {badge > 99 ? "99+" : badge}
              </span>
            )}
          </button>
        ))}
      </nav>
    </div>
  );
}

export default OutlookEmail;


// ─── Email compose component (merged from former Email.js) ──────────────────
const QUOTE_COLORS = ["bg-blue-500","bg-red-500","bg-green-500","bg-purple-500","bg-orange-500","bg-teal-500","bg-pink-500","bg-indigo-500"];
function quoteAvatarColor(name) {
  if (!name) return QUOTE_COLORS[0];
  let h = 0;
  for (let i = 0; i < name.length; i++) h = (h * 31 + name.charCodeAt(i)) & 0xffff;
  return QUOTE_COLORS[h % QUOTE_COLORS.length];
}
function quoteInitials(name) {
  if (!name) return "?";
  const parts = name.trim().split(/\s+/);
  return (parts.length > 1 ? parts[0][0] + parts[parts.length - 1][0] : name.slice(0, 2)).toUpperCase();
}
function quoteFileSize(bytes) {
  if (!bytes) return "";
  if (bytes >= 1024 * 1024) return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
  if (bytes >= 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${bytes} B`;
}

export function Email({ accessToken: accessTokenProp, onClose, onMailSent, replyData, inline = false, onTemplatesOpenChange }) {
  const { instance, accounts } = useMsal();
  const [accessToken, setAccessToken] = useState(accessTokenProp || "");

  // Templates state
  const [templates, setTemplates] = useState([]);
  const [selectedTemplate, setSelectedTemplate] = useState(""); // Applied template
  // Modal state for template
  const [modalSelectedTemplate, setModalSelectedTemplate] = useState("");
  // Sync modal state with applied state when opening modal

  // States
  const [toInput, setToInput] = useState("");   // confirmed recipients, comma-separated
  const [toDraft, setToDraft] = useState("");   // the address currently being typed
  const [ccInput, setCcInput] = useState("");
  const [subject, setSubject] = useState("");
  const [body, setBody] = useState("");
  // The footer is one unit: a signature text + an optional company logo, kept
  // separate from the message body so they move/send together as a single block.
  const [footerText, setFooterText] = useState("");
  const [footerLogo, setFooterLogo] = useState("");
  const [showCc, setShowCc] = useState(false);
  const [isSending, setIsSending] = useState(false);
  const [errors, setErrors] = useState({});
  const [successMessage, setSuccessMessage] = useState("");
  const [sendDropdown, setSendDropdown]           = useState(false);
  const [scheduleModal, setScheduleModal]         = useState(false);
  const [scheduleDateTime, setScheduleDateTime]   = useState("");
  const [schedulingInProgress, setSchedulingInProgress] = useState(false);
  // Show all emails modal for To field
  const [showAllToEmails, setShowAllToEmails] = useState(false);
  // Tag selection state
  const [tags, setTags] = useState([]);
  const [tagsLoading, setTagsLoading] = useState(false);
  const [tagsError, setTagsError] = useState("");
  const [selectedTags, setSelectedTags] = useState([]); // Applied tags
  // Modal state for tags
  const [modalSelectedTags, setModalSelectedTags] = useState([]);

  // Templates & Tags Modal state
  const [showTemplatesModal, setShowTemplatesModal] = useState(false);
  // Let the host (OutlookEmail) react when the picker opens (e.g. widen the area).
  useEffect(() => { if (onTemplatesOpenChange) onTemplatesOpenChange(showTemplatesModal); }, [showTemplatesModal, onTemplatesOpenChange]);
  const [searchQuery, setSearchQuery] = useState("");

  // Original email quote (reply/replyAll mode) — shown visually, not in textarea
  const [quoteHtml, setQuoteHtml] = useState("");

  // File attachment and image states
  const [attachments, setAttachments] = useState([]);
  const [images, setImages] = useState([]);
  const attachmentInputRef = React.useRef(null);
  const imageInputRef = React.useRef(null);

  // Email suggestions state
  const [emailSuggestions, setEmailSuggestions] = useState([]);
  const [showEmailSuggestions, setShowEmailSuggestions] = useState(false);
  const [filteredEmailSuggestions, setFilteredEmailSuggestions] = useState([]);
  const toInputRef = React.useRef(null);

  async function fetchApi(url, options) {
    try {
      if (/^https?:\/\//i.test(url)) return window.fetch(url, options);
      const path = url.startsWith("/") ? url : `/${url}`;
      const method = (options && options.method) || "GET";
      if (method === "GET") {
        const res = await apiClient.get(path, { params: options && options.params });
        return { ok: res.status >= 200 && res.status < 300, status: res.status, json: async () => res.data, text: async () => JSON.stringify(res.data) };
      }
      const res = await apiClient.request({ url: path, method, data: options && options.body, headers: options && options.headers });
      return { ok: res.status >= 200 && res.status < 300, status: res.status, json: async () => res.data, text: async () => JSON.stringify(res.data) };
    } catch (err) {
      return Promise.reject(err);
    }
  }

  const fetch = fetchApi;

  //  DRAG & DROP
  const [isDragging, setIsDragging] = useState(false);

  const handleDragOver = (e) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);

    const droppedFiles = Array.from(e.dataTransfer.files);

    const imageFiles = droppedFiles.filter((file) =>
      file.type.startsWith("image/")
    );
    const otherFiles = droppedFiles.filter(
      (file) => !file.type.startsWith("image/")
    );

    if (imageFiles.length > 0) {
      setImages((prev) => [...prev, ...imageFiles]);
    }

    if (otherFiles.length > 0) {
      setAttachments((prev) => [...prev, ...otherFiles]);
    }
  };

  // Filter templates and tags based on search
  const filteredTemplates = templates.filter(
    (tpl) =>
      (tpl.name || "").toLowerCase().includes(searchQuery.toLowerCase()) ||
      ((tpl.templateType ?? tpl.TemplateType) &&
        (tpl.templateType ?? tpl.TemplateType).toLowerCase().includes(searchQuery.toLowerCase())) ||
      ((tpl.body ?? tpl.Body) &&
        (tpl.body ?? tpl.Body).toLowerCase().includes(searchQuery.toLowerCase()))
  );

  const filteredTags = tags.filter((tag) =>
    tag.toLowerCase().includes(searchQuery.toLowerCase())
  );

  // Handler functions for modal — selecting reflects immediately (template footer
  // into the message box, tag emails into the To field), not only on "Apply".
  const handleTemplateSelect = (templateName) => {
    setModalSelectedTemplate(templateName);
    setSelectedTemplate(templateName);
  };

  const handleTagToggle = (tag) => {
    const next = modalSelectedTags.includes(tag)
      ? modalSelectedTags.filter((t) => t !== tag)
      : [...modalSelectedTags, tag];
    setModalSelectedTags(next);
    setSelectedTags(next);
  };

  const handleApply = () => {
    setSelectedTemplate(modalSelectedTemplate);
    setSelectedTags(modalSelectedTags);
    setShowTemplatesModal(false);
  };

  const handleClearAll = () => {
    setModalSelectedTemplate("");
    setModalSelectedTags([]);
    setSelectedTemplate("");
    setSelectedTags([]);
  };

  // Get access token from MSAL (redirect-safe, no popup)
useEffect(() => {
  console.log(
    "Email: Getting token. Prop:",
    !!accessTokenProp,
    "Accounts:",
    accounts.length
  );

  // 1️⃣ If token was passed in, just use it
  if (accessTokenProp) {
    console.log("Email: Using provided access token");
    setAccessToken(accessTokenProp);
    return;
  }

  // 2️⃣ No account → nothing to do
  if (accounts.length === 0) {
    console.warn("Email: No accounts found");
    setAccessToken("");
    return;
  }

  const request = {
    account: accounts[0],
    scopes: [
      "https://graph.microsoft.com/User.Read",
      "https://graph.microsoft.com/Mail.Read",
      "https://graph.microsoft.com/Mail.ReadWrite",
      "https://graph.microsoft.com/Mail.Send",
    ],
  };

  const getToken = async () => {
    try {
      // ✅ Silent first
      const response = await instance.acquireTokenSilent(request);
      console.log("Email: Token acquired (silent)");
      setAccessToken(response.accessToken);
    } catch (error) {
      console.warn(
        "Email: Silent token failed, redirecting:",
        error.errorCode
      );

      // ✅ Redirect fallback (NO popup)
      instance.acquireTokenRedirect(request);
    }
  };

  getToken();
}, [accessTokenProp, instance, accounts]);

  // Manage body overflow when modal is open
  useEffect(() => {
    // Prevent body scroll when modal is open
    document.body.style.overflow = "hidden";
    return () => {
      // Restore body scroll when modal is closed/unmounted
      document.body.style.overflow = "unset";
    };
  }, []);

  // Fetch templates on mount
  useEffect(() => {
    apiClient
      .get(`/Template`)
      .then((res) => {
        const templatesArray = Array.isArray(res.data) ? res.data : [];
        // Active templates that belong to the logged-in user only.
        const me = (sessionStorage.getItem("userEmail") || accounts[0]?.username || "").toLowerCase().trim();
        const mine = templatesArray.filter((t) => {
          if (t.IsActive === false || t.isActive === false) return false;
          const cb = (t.createdBy ?? t.CreatedBy ?? "").toLowerCase().trim();
          if (cb) return cb === me;
          const hay = `${t.name ?? ""} ${t.body ?? t.Body ?? ""}`.toLowerCase();
          return !!me && hay.includes(me);
        });
        console.debug("[Email] Templates loaded (mine):", mine);
        setTemplates(mine);
      })
      .catch((err) => {
        console.error("[Email] Failed to load templates:", err);
        setTemplates([]);
      });
  }, []);

  // Auto-select the template that belongs to the logged-in user (their signature)
  // for a fresh new email — matched primarily by the user's EMAIL appearing in the
  // template (footer/name) or its CreatedBy. Skips replies/forwards and never
  // overrides a template the user picked themselves.
  useEffect(() => {
    if (!templates.length || selectedTemplate) return;
    const t = replyData?.type;
    if (t === "reply" || t === "replyAll" || t === "forward") return;
    if (replyData?.body) return; // "Use Template" already supplied a specific template

    // The logged-in user's email(s): CRM session login and the connected mailbox.
    const emails = [sessionStorage.getItem("userEmail"), accounts[0]?.username]
      .map((e) => (e || "").toLowerCase().trim())
      .filter(Boolean);

    let match = null;
    // Primary: the template that belongs to or carries the logged-in email.
    if (emails.length) {
      match = templates.find((tpl) => {
        const createdBy = (tpl.createdBy ?? tpl.CreatedBy ?? "").toLowerCase();
        if (emails.includes(createdBy)) return true;
        const hay = `${tpl.name ?? ""} ${tpl.body ?? tpl.Body ?? ""}`.toLowerCase();
        return emails.some((em) => hay.includes(em));
      });
    }

    // Fallback: match by the user's name / email handle if no email match.
    if (!match) {
      const name = (accounts[0]?.name || sessionStorage.getItem("userName") || "").toLowerCase().trim();
      const local = (emails[0] || "").split("@")[0];
      const tokens = [name, ...name.split(/\s+/), local, ...local.split(/[._-]+/)]
        .map((s) => s.trim())
        .filter((s) => s.length >= 3);
      if (tokens.length) {
        match = templates.find((tpl) => {
          const hay = `${tpl.name ?? ""} ${tpl.body ?? tpl.Body ?? ""}`.toLowerCase();
          return tokens.some((tok) => hay.includes(tok));
        });
      }
    }

    if (match) {
      setSelectedTemplate(match.name);
      setModalSelectedTemplate(match.name);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [templates, accounts]);

  // Fetch tags on mount (so button shows if there are any tags)
  useEffect(() => {
    setTagsLoading(true);
    setTagsError("");
    apiClient
      .get(`/Contact/tags/all`)
      .then((res) => {
        const tagsArray = Array.isArray(res.data) ? res.data : [];
        console.debug("[Email] Tags loaded:", tagsArray);
        setTags(tagsArray);
        setTagsLoading(false);
      })
      .catch((err) => {
        console.error("[Email] Failed to load tags:", err);
        setTags([]);
        setTagsError("Could not load tags");
        setTagsLoading(false);
      });
  }, []);

  // Fetch contact emails for autocomplete using paged endpoint (avoids loading all contacts)
  useEffect(() => {
    const fetchContactEmails = async () => {
      try {
        const allEmails = [];
        let page = 1;
        const pageSize = 100;
        while (true) {
          const res = await apiClient.get(`/Contact?page=${page}&pageSize=${pageSize}`);
          const items = Array.isArray(res.data?.items) ? res.data.items : [];
          items.forEach((contact) => {
            const workEmail = contact.WorkEmail || contact.workEmail || "";
            if (typeof workEmail === "string" && workEmail.trim()) {
              allEmails.push(workEmail.trim());
            }
          });
          if (items.length < pageSize) break;
          page++;
          if (page > 10) break; // safety cap at 1000 emails
        }
        setEmailSuggestions([...new Set(allEmails)].sort());
      } catch (error) {
        setEmailSuggestions([]);
      }
    };
    fetchContactEmails();
  }, []);

  // When a template is selected, reflect its footer (body) and subject into the
  // message box. The templates list from GET /Template already carries full data,
  // so we resolve it locally instead of re-fetching by name — `/Template/{name}`
  // isn't a valid route (the API only exposes /Template/{id} and /Template/name/{name}).
  useEffect(() => {
    if (!selectedTemplate) return;
    const tpl = templates.find((t) => (t.name ?? t.Name) === selectedTemplate);
    if (tpl) {
      const { body, footer, logo, attachments } = parseTemplateBody(tpl.body ?? tpl.Body ?? "");
      setSubject(tpl.subject ?? tpl.Subject ?? "");
      if (body) setBody(body);
      setFooterText(footer);
      setFooterLogo(logo);
      if (Array.isArray(attachments) && attachments.length) setAttachments(attachments);
    }
  }, [selectedTemplate, templates]);

  // On mount, check for selectedContactEmails in localStorage to prefill To field (one-shot, clear after read)
  useEffect(() => {
    try {
      const raw = localStorage.getItem("selectedContactEmails");
      if (raw) {
        localStorage.removeItem("selectedContactEmails");
        const arr = JSON.parse(raw);
        if (Array.isArray(arr) && arr.length > 0) {
          setToInput(arr.join(", "));
        }
      }
    } catch (e) {}
  }, []);

  // Convert HTML to plain text while preserving structure
  const htmlToPlainText = (html) => {
    if (!html) return "";
    let text = html;
    // Convert common HTML tags to newlines
    text = text.replace(/<br\s*\/?>/gi, "\n");
    text = text.replace(/<\/p>/gi, "\n");
    text = text.replace(/<\/div>/gi, "\n");
    text = text.replace(/<\/li>/gi, "\n");
    // Remove all remaining HTML tags
    text = text.replace(/<[^>]*>/g, "");
    // Decode HTML entities
    text = text.replace(/&nbsp;/g, " ");
    text = text.replace(/&lt;/g, "<");
    text = text.replace(/&gt;/g, ">");
    text = text.replace(/&amp;/g, "&");
    text = text.replace(/&quot;/g, '"');
    text = text.replace(/&#39;/g, "'");
    // Clean up multiple newlines
    text = text.replace(/\n\n\n+/g, "\n\n");
    return text.trim();
  };

  // Handle reply data when provided
  useEffect(() => {
    if (replyData) {
      setToInput(replyData.toEmail || "");
      setCcInput(replyData.ccEmail || "");
      setSubject(replyData.subject || "");

      if (replyData.type === "new" && replyData.body) {
        // "Use Template" passed a packed template body (body + footer + optional logo + attachments).
        const { body, footer, logo, attachments } = parseTemplateBody(replyData.body);
        if (body) setBody(body);
        setFooterText(footer);
        setFooterLogo(logo);
        if (Array.isArray(attachments) && attachments.length) setAttachments(attachments);
        setQuoteHtml("");
      } else if (replyData.type === "reply" || replyData.type === "replyAll") {
        // Body starts empty — user types reply above the quoted original
        setBody("");
        setQuoteHtml(replyData.originalBody || "");
      } else if (replyData.type === "forward") {
        const plainTextBody = htmlToPlainText(replyData.originalBody);
        const forwardText = `\n\n---\nForwarded message:\nFrom: ${replyData.originalMail.sender}\nTo: ${replyData.originalMail.toEmail}\nDate: ${new Date(replyData.originalMail.receivedDateTime).toLocaleString()}\nSubject: ${replyData.originalMail.subject}\n\n${plainTextBody}`;
        setBody(forwardText);
        setQuoteHtml("");
        if (replyData.attachments && replyData.attachments.length > 0) {
          setAttachments(replyData.attachments);
        }
      }

      if (replyData.ccEmail) {
        setShowCc(true);
      }
    }
  }, [replyData]);

  // Append a confirmed address to the To list (kept comma-separated).
  const commitToEmail = (email) => {
    const clean = (email || "").trim();
    if (!clean) return;
    setToInput((prev) => (prev.trim() ? `${prev.trim()}, ${clean}` : clean));
  };

  // Handle To input change with email autocomplete. Only the small draft string
  // changes per keystroke — the (potentially large) confirmed list is untouched,
  // so typing stays fast even with many recipients.
  const handleToInputChange = (e) => {
    const value = e.target.value;
    const lastChar = value[value.length - 1];

    // Separator commits the typed address.
    if (lastChar === "," || lastChar === ";") {
      commitToEmail(value.slice(0, -1));
      setToDraft("");
      setShowEmailSuggestions(false);
      setFilteredEmailSuggestions([]);
      return;
    }

    setToDraft(value);

    const lastEmail = value.trim().toLowerCase();
    if (lastEmail) {
      // Cap at 8 matches with an early exit — avoids scanning/rendering the whole list.
      const filtered = [];
      for (let i = 0; i < emailSuggestions.length && filtered.length < 8; i++) {
        if (emailSuggestions[i].toLowerCase().includes(lastEmail)) {
          filtered.push(emailSuggestions[i]);
        }
      }
      setFilteredEmailSuggestions(filtered);
      setShowEmailSuggestions(filtered.length > 0);
    } else {
      setShowEmailSuggestions(false);
      setFilteredEmailSuggestions([]);
    }
  };

  // Handle selecting an email from suggestions
  const handleSelectEmailSuggestion = (email) => {
    commitToEmail(email);
    setToDraft("");
    setShowEmailSuggestions(false);
    setFilteredEmailSuggestions([]);
    if (toInputRef.current) {
      toInputRef.current.focus();
    }
  };

  // Helper: parse a user-provided string of emails into an array
  const parseEmails = (input) => {
    if (!input) return [];
    return input
      .split(/[;,\n]+/) // split on comma, semicolon or newline
      .map((s) => s.trim())
      .filter((s) => s.length > 0);
  };

  // Confirmed recipients, parsed once per change (avoids re-splitting on every render).
  const toEmails = useMemo(() => parseEmails(toInput), [toInput]);
  // Recipients including the in-progress draft — used when sending.
  const allToEmails = useMemo(
    () => parseEmails(toDraft ? `${toInput}, ${toDraft}` : toInput),
    [toInput, toDraft]
  );

  const validateEmail = (email) => {
    // simple regex for basic validation
    const re =
      /^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return re.test(String(email).toLowerCase());
  };

  // Handle attachment file selection
  const handleAttachmentSelect = (e) => {
    const files = Array.from(e.target.files || []);
    setAttachments([...attachments, ...files]);
    if (attachmentInputRef.current) {
      attachmentInputRef.current.value = "";
    }
  };

  // Handle image file selection
  const handleImageSelect = (e) => {
    const files = Array.from(e.target.files || []);
    const imageFiles = files.filter((file) => file.type.startsWith("image/"));
    setImages([...images, ...imageFiles]);
    if (imageInputRef.current) {
      imageInputRef.current.value = "";
    }
  };

  // Remove attachment
  const removeAttachment = (index) => {
    setAttachments(attachments.filter((_, i) => i !== index));
  };

  // Remove image
  const removeImage = (index) => {
    setImages(images.filter((_, i) => i !== index));
  };

  // Convert file to base64
  const fileToBase64 = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = () => {
        const base64 = reader.result.split(",")[1];
        resolve(base64);
      };
      reader.onerror = reject;
    });
  };

  // The footer as one block: company logo on the LEFT, signature text on the right.
  const footerBlockHtml = () => {
    if (!footerText && !footerLogo) return "";
    const textHtml = (footerText || "").replace(/\n/g, "<br>").replace(/ {2}/g, "&nbsp;&nbsp;");
    if (footerLogo) {
      return `<table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;margin-top:18px;"><tr>`
        + `<td valign="middle" style="padding-right:14px;"><img src="${footerLogo}" alt="logo" style="display:block;max-width:140px;max-height:90px;" /></td>`
        + `<td valign="middle" style="font-size:15px;color:#374151;line-height:1.6;">${textHtml}</td>`
        + `</tr></table>`;
    }
    return `<div style="margin-top:18px;font-size:15px;color:#374151;line-height:1.6;">${textHtml}</div>`;
  };

  // Handle Form Submission
  const handleSubmit = async (e) => {
    e.preventDefault();
    setIsSending(true);

    // Parse inputs into arrays (include any address still being typed)
    const toEmails = allToEmails;
    const ccEmails = parseEmails(ccInput);

    // Frontend validation for empty or invalid fields
    let validationErrors = {};
    if (toEmails.length === 0)
      validationErrors.to = "At least one recipient is required";
    if (!subject.trim()) validationErrors.subject = "Subject is required";
    if (!body.trim() && !quoteHtml) validationErrors.body = "Email body is required";

    // Validate email formats
    const invalidTo = toEmails.filter((addr) => !validateEmail(addr));
    if (invalidTo.length > 0)
      validationErrors.to = `Invalid To address(es): ${invalidTo.join(", ")}`;
    const invalidCc = ccEmails.filter((addr) => !validateEmail(addr));
    if (invalidCc.length > 0)
      validationErrors.cc = `Invalid Cc address(es): ${invalidCc.join(", ")}`;

    // Update error state and clear success message
    setErrors(validationErrors);
    setSuccessMessage("");

    // If any validation failed, stop submission here
    if (Object.keys(validationErrors).length > 0) {
      setIsSending(false);
      return;
    }

    // Must be signed in to Outlook (Microsoft Graph) to send. Without a token every
    // Graph sendMail returns 401, which is exactly why all recipients showed "failed".
    // Block here with a clear message instead of creating a campaign full of failures.
    if (!accessToken) {
      setErrors({
        apiError:
          accounts && accounts.length > 0
            ? "Your Outlook session has expired. Open the Email menu to sign in again, then resend."
            : "You're not signed in to Outlook. Open the Email menu (left sidebar) to sign in with Microsoft, then resend.",
      });
      setSuccessMessage("");
      setIsSending(false);
      return;
    }

    // Prepare attachments for Microsoft Graph API
    let attachmentsPayload = [];
    try {
      for (const file of attachments) {
        // Template/reply attachments arrive already base64-encoded ({name, contentBytes});
        // freshly-picked files are File objects that need reading.
        if (file && file.contentBytes) {
          attachmentsPayload.push({
            "@odata.type": "#microsoft.graph.fileAttachment",
            name: file.name,
            contentBytes: file.contentBytes,
            contentType: file.contentType || "application/octet-stream",
          });
        } else {
          const base64 = await fileToBase64(file);
          attachmentsPayload.push({
            "@odata.type": "#microsoft.graph.fileAttachment",
            name: file.name,
            contentBytes: base64,
          });
        }
      }

      for (const file of images) {
        const base64 = await fileToBase64(file);
        attachmentsPayload.push({
          "@odata.type": "#microsoft.graph.fileAttachment",
          name: file.name,
          contentBytes: base64,
        });
      }
    } catch (error) {
      console.error("Error converting files to base64:", error);
      setErrors({ apiError: "Failed to process attachments/images" });
      setIsSending(false);
      return;
    }

    // Build the final HTML.
    const userHtml = body.replace(/\n/g, "<br>").replace(/ {2}/g, "&nbsp;&nbsp;");
    const finalHtml = quoteHtml
      ? `${userHtml}${footerBlockHtml()}<br><br><div style="border-left:2px solid #e5e7eb;padding-left:12px;margin-top:8px;color:#6b7280;font-size:13px;">${quoteHtml}</div>`
      : `${userHtml}${footerBlockHtml()}`;

    // Normal send — ONE message to all recipients, no tracking.
    // (For open/click tracking, use the Send dropdown → "Send tracked campaign".)
    try {
      const emailPayload = {
        message: {
          subject,
          body: { contentType: "HTML", content: finalHtml },
          toRecipients: toEmails.map((addr) => ({ emailAddress: { address: addr } })),
          ccRecipients: ccEmails.length > 0 ? ccEmails.map((addr) => ({ emailAddress: { address: addr } })) : [],
        },
      };
      if (attachmentsPayload.length > 0) emailPayload.message.attachments = attachmentsPayload;

      const response = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
        method: "POST",
        headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
        body: JSON.stringify(emailPayload),
      });

      if (response.ok || response.status === 202) {
        setSuccessMessage("Email sent successfully!");
        setErrors({});
        setToInput(""); setCcInput(""); setSubject(""); setBody("");
        setQuoteHtml(""); setSelectedTemplate(""); setSelectedTags([]);
        setAttachments([]); setImages([]);
        try { localStorage.removeItem("selectedContactEmails"); } catch (e) {}
        if (typeof onMailSent === "function") onMailSent({ subject, body, to: toEmails, cc: ccEmails });
        setIsSending(false);
        setTimeout(() => onClose(), 1000);
      } else {
        const errorData = await response.json();
        console.error("Graph API error response:", errorData);
        setErrors({ apiError: errorData.error?.message || "Failed to send email" });
        setIsSending(false);
      }
    } catch (error) {
      console.error("Network/parse error when sending email:", error);
      setErrors({ apiError: error.message || "Failed to send email" });
      setIsSending(false);
    }
  };

  // Build send payload (shared by schedule + mail merge)
  const buildPayload = () => {
    const toEmails = allToEmails;
    const ccEmails = parseEmails(ccInput);
    const userHtml = body.replace(/\n/g, "<br>").replace(/ {2}/g, "&nbsp;&nbsp;");
    const finalHtml = quoteHtml
      ? `${userHtml}${footerBlockHtml()}<br><br><div style="border-left:2px solid #e5e7eb;padding-left:12px;margin-top:8px;color:#6b7280;font-size:13px;">${quoteHtml}</div>`
      : `${userHtml}${footerBlockHtml()}`;
    return { toEmails, ccEmails, html: finalHtml };
  };

  const handleScheduleSend = async () => {
    if (!scheduleDateTime) return;
    const scheduledAt = new Date(scheduleDateTime);
    const delay = scheduledAt.getTime() - Date.now();
    if (delay <= 0) { setErrors({ apiError: "Please select a future date and time." }); return; }
    const { toEmails, ccEmails, html } = buildPayload();
    if (!toEmails.length) { setErrors({ to: "At least one recipient is required." }); return; }
    setSchedulingInProgress(true);
    try {
      const draftRes = await fetch("https://graph.microsoft.com/v1.0/me/messages", {
        method: "POST",
        headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
        body: JSON.stringify({
          subject,
          body: { contentType: "HTML", content: html },
          toRecipients: toEmails.map(a => ({ emailAddress: { address: a } })),
          ccRecipients: ccEmails.map(a => ({ emailAddress: { address: a } })),
        }),
      });
      if (draftRes.ok) {
        const draft = await draftRes.json();
        setTimeout(async () => {
          await fetch(`https://graph.microsoft.com/v1.0/me/messages/${draft.id}/send`, {
            method: "POST", headers: { Authorization: `Bearer ${accessToken}` },
          }).catch(() => {});
        }, delay);
        setScheduleModal(false);
        setSuccessMessage(`Scheduled for ${scheduledAt.toLocaleString()}`);
        setTimeout(() => onClose(), 1200);
      } else {
        const err = await draftRes.json().catch(() => ({}));
        setErrors({ apiError: err?.error?.message || "Failed to schedule email." });
      }
    } catch (e) {
      setErrors({ apiError: e.message || "Failed to schedule." });
    } finally {
      setSchedulingInProgress(false);
    }
  };

  const handleMailMerge = async () => {
    const { toEmails, ccEmails, html } = buildPayload();
    if (!toEmails.length) { setErrors({ to: "At least one recipient is required." }); return; }
    if (!subject.trim()) { setErrors({ subject: "Subject is required." }); return; }
    setIsSending(true); setErrors({});
    let sent = 0;
    for (const addr of toEmails) {
      try {
        const res = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
          method: "POST",
          headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
          body: JSON.stringify({ message: {
            subject,
            body: { contentType: "HTML", content: html },
            toRecipients: [{ emailAddress: { address: addr } }],
            ccRecipients: ccEmails.map(a => ({ emailAddress: { address: a } })),
          }}),
        });
        if (res.ok || res.status === 202) sent++;
      } catch {}
    }
    setIsSending(false);
    setSuccessMessage(`Mail merge sent to ${sent} of ${toEmails.length} recipient${toEmails.length !== 1 ? "s" : ""}.`);
    setTimeout(() => onClose(), 1200);
  };

  // TRACKED CAMPAIGN — sends each recipient their OWN email with a unique open pixel + tracked
  // links, so opens/clicks/unsubscribes are attributed per person on the Email Tracking page.
  // Sending is browser-driven (uses your Outlook token), and each outcome is reported back so the
  // tracking page shows real delivery. The plain Send button (above) does NOT track.
  const handleTrackedCampaign = async () => {
    const { toEmails, ccEmails, html } = buildPayload();
    if (!toEmails.length) { setErrors({ to: "At least one recipient is required." }); return; }
    if (!subject.trim()) { setErrors({ subject: "Subject is required." }); return; }
    if (!accessToken) {
      setErrors({
        apiError: accounts && accounts.length > 0
          ? "Your Outlook session has expired. Open the Email menu to sign in again, then resend."
          : "You're not signed in to Outlook. Open the Email menu (left sidebar) to sign in with Microsoft, then resend.",
      });
      return;
    }
    setIsSending(true); setErrors({});

    // 1) Register the campaign + get each recipient's tracked HTML (open pixel + links woven in).
    let tracked = null;
    try {
      // Tell the backend the exact origin we reach the API at, so the tracking / unsubscribe
      // links in the email resolve to the backend in production (not the SPA host).
      let publicBaseUrl = (apiClient?.defaults?.baseURL || "").replace(/\/api\/?$/i, "");
      if (!publicBaseUrl || publicBaseUrl.startsWith("/")) publicBaseUrl = window.location.origin;
      const prep = await apiClient.post("/EmailCampaign/prepare", {
        subject,
        body: html,
        publicBaseUrl,
        // The Subscribe/Unsubscribe buttons land on this app's /email/* pages (always reachable),
        // which record via the working /api path.
        subscribeBaseUrl: window.location.origin,
        recipients: toEmails.map((e) => ({ email: e })),
      });
      const items = Array.isArray(prep.data?.items) ? prep.data.items : null;
      if (items && items.length) tracked = { campaignId: prep.data.campaignId, items };
    } catch (e) {
      console.error("Tracking prepare failed:", e);
    }
    if (!tracked) {
      setErrors({ apiError: "Couldn't start the tracked campaign. Is the backend running?" });
      setIsSending(false);
      return;
    }

    // Get (or create) a dedicated Outlook folder to file campaign copies into, so a large
    // send doesn't pile up in Sent Items / the mailbox. Falls back to Sent Items if unavailable.
    let campaignFolderName = "CRM Campaigns";
    try { campaignFolderName = (localStorage.getItem("_crm_campaign_folder") || "CRM Campaigns").trim() || "CRM Campaigns"; } catch { /* default */ }
    const campaignFolderId = await ensureMailFolderId(accessToken, campaignFolderName);

    // Build the attachment payload once (shared by every recipient). Handles both freshly-picked
    // File objects and pre-encoded template attachments ({name, contentBytes}).
    let campaignAttachments = [];
    try {
      for (const file of attachments) {
        if (file && file.contentBytes) {
          campaignAttachments.push({ "@odata.type": "#microsoft.graph.fileAttachment", name: file.name, contentBytes: file.contentBytes, contentType: file.contentType || "application/octet-stream" });
        } else {
          const base64 = await fileToBase64(file);
          campaignAttachments.push({ "@odata.type": "#microsoft.graph.fileAttachment", name: file.name, contentBytes: base64 });
        }
      }
      for (const file of images) {
        const base64 = await fileToBase64(file);
        campaignAttachments.push({ "@odata.type": "#microsoft.graph.fileAttachment", name: file.name, contentBytes: base64 });
      }
    } catch (e) { console.error("Campaign attachment prep failed:", e); }

    // Helper: pull a readable error message out of a failed Graph response.
    const readGraphErr = async (r) => {
      try {
        const txt = await r.text();
        const parsed = txt && txt.trim().startsWith("{") ? JSON.parse(txt) : null;
        return parsed?.error?.message || (txt ? txt.slice(0, 300) : "") || `HTTP ${r.status}`;
      } catch { return `HTTP ${r.status}`; }
    };

    // 2) Send each recipient their own tracked email, collecting outcomes.
    // When a campaign folder is chosen we send via draft → send, and remember the message id so
    // we can FILE it into the folder *afterwards* (step 2b). Filing has to wait until the send has
    // actually finished — POST /send returns 202 (async), so moving the message immediately can
    // grab it while it's still a draft mid-send, which Outlook then shows as an "unknown message".
    const results = [];
    const sentIds = [];   // ids of successfully sent messages to file into the campaign folder
    for (let i = 0; i < tracked.items.length; i++) {
      const item = tracked.items[i];
      const message = {
        subject,
        body: { contentType: "HTML", content: item.html },
        toRecipients: [{ emailAddress: { address: item.email } }],
        ccRecipients: i === 0 ? ccEmails.map((a) => ({ emailAddress: { address: a } })) : [],
        ...(campaignAttachments.length ? { attachments: campaignAttachments } : {}),
        // Ask for read + delivery receipts so we can confirm opens/delivery even when
        // the recipient's client blocks the tracking pixel (images turned off).
        isReadReceiptRequested: true,
        isDeliveryReceiptRequested: true,
      };
      try {
        let ok = false, errMsg = null;
        if (campaignFolderId) {
          // Draft → send produces a REAL sent message (POSTing to a folder's /messages only ever
          // creates a *draft*, which is why copies showed as "This message hasn't been sent").
          const draftRes = await fetch("https://graph.microsoft.com/v1.0/me/messages", {
            method: "POST",
            headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
            body: JSON.stringify(message),
          });
          if (!(draftRes.ok || draftRes.status === 201)) {
            errMsg = await readGraphErr(draftRes);
          } else {
            const draft = await draftRes.json();
            const mid = draft.id;
            // Send the draft — Graph delivers it and moves it to Sent Items.
            const sendRes = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${mid}/send`, {
              method: "POST",
              headers: { Authorization: `Bearer ${accessToken}` },
            });
            ok = sendRes.ok || sendRes.status === 202;
            if (!ok) errMsg = await readGraphErr(sendRes);
            else if (mid) sentIds.push(mid);   // file it into the campaign folder after the loop
          }
        } else {
          // No campaign folder chosen — plain send, kept in Sent Items.
          const res = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
            method: "POST",
            headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
            body: JSON.stringify({ message, saveToSentItems: true }),
          });
          ok = res.ok || res.status === 202;
          if (!ok) errMsg = await readGraphErr(res);
        }
        results.push({ recipientId: item.recipientId, status: ok ? "Sent" : "Failed", error: errMsg });
      } catch (err) {
        results.push({ recipientId: item.recipientId, status: "Failed", error: err.message });
      }
    }

    // 2b) File the sent copies into the campaign folder — only once each send has finished, i.e.
    // the message is no longer a draft (isDraft === false) and has landed in Sent Items. Moving a
    // message still mid-send is what produced the "unknown message" items. Best-effort per message:
    // if it never settles, the email is still sent and just stays in Sent Items.
    if (campaignFolderId && sentIds.length) {
      const wait = (ms) => new Promise((r) => setTimeout(r, ms));
      for (const mid of sentIds) {
        try {
          let ready = false;
          for (let a = 0; a < 8 && !ready; a++) {
            const chk = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${mid}?$select=isDraft`, {
              headers: { Authorization: `Bearer ${accessToken}` },
            });
            if (chk.ok) {
              const j = await chk.json();
              if (j.isDraft === false) { ready = true; break; }
            }
            await wait(500);
          }
          if (ready) {
            await fetch(`https://graph.microsoft.com/v1.0/me/messages/${mid}/move`, {
              method: "POST",
              headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
              body: JSON.stringify({ destinationId: campaignFolderId }),
            });
          }
        } catch { /* filing is best-effort */ }
      }
    }

    // 3) Report outcomes so the Email Tracking page reflects delivery.
    try { await apiClient.post(`/EmailCampaign/${tracked.campaignId}/status`, { results }); } catch (e) {}

    const sent = results.filter((r) => r.status === "Sent").length;
    setIsSending(false);
    if (sent > 0) {
      setSuccessMessage(`Sent & tracking ${sent} of ${results.length} email${results.length !== 1 ? "s" : ""}. Follow opens & clicks on the Email Tracking page.`);
      setErrors({});
      setToInput(""); setCcInput(""); setSubject(""); setBody(""); setQuoteHtml("");
      setSelectedTags([]); setAttachments([]); setImages([]);
      try { localStorage.removeItem("selectedContactEmails"); } catch (e) {}
      setTimeout(() => onClose(), 1500);
    } else {
      setErrors({ apiError: `All sends failed — ${results[0]?.error || "check your Outlook sign-in / permissions."}` });
    }
  };

  // When tags are applied, reflect the matching contacts' emails into the To field.
  useEffect(() => {
    if (!selectedTags.length) return;
    apiClient
      .get(`/Contact/tags/emails`, { params: { tags: selectedTags.join(",") } })
      .then((res) => {
        // Backend returns a comma-joined string; be defensive about other shapes.
        let raw = res.data;
        if (Array.isArray(raw)) raw = raw.join(",");
        else if (raw && typeof raw === "object") raw = Object.values(raw).join(",");
        const emailArr = String(raw ?? "")
          .split(/[,;\n]+/)
          .map((e) => e.trim())
          .filter((e) => e.length > 0);
        if (emailArr.length === 0) {
          // No emails for these tags — don't wipe what's already in To.
          console.warn("[Email] No contact emails found for tags:", selectedTags);
          return;
        }
        // Reflect the selected tags' contact emails into the To field.
        setToInput([...new Set(emailArr)].join(", "));
      })
      .catch((err) => {
        console.error("[Email] Failed to load emails for tags:", selectedTags, err);
      });
  }, [selectedTags]);


  // Static light-mode theme (Email compose always light)
  const d = {
    bg: "bg-white", surface: "bg-gray-50", text: "text-gray-900",
    muted: "text-gray-600", border: "border-gray-200", borderSm: "border-gray-100",
    input: "bg-white border-gray-300 text-gray-900 placeholder-gray-400 focus:ring-blue-500",
    hover: "hover:bg-gray-100", hoverBlue: "hover:bg-blue-50",
    chip: "bg-white border-gray-300 text-gray-700",
    tag: "bg-blue-100 text-blue-800 border-blue-300",
  };
  const isDark = false;

  return (
    <div className={inline ? "w-full h-full flex flex-col min-h-0" : "fixed inset-0 bg-black/40 backdrop-blur-sm flex items-center justify-center z-[1100] p-3 sm:p-4"}>
      <div className={inline ? "bg-white w-full h-full flex flex-col overflow-hidden min-h-0" : "bg-white rounded-xl shadow-xl w-full max-w-2xl sm:max-w-4xl max-h-[95vh] sm:max-h-[90vh] flex flex-col overflow-hidden"}>
        {/* Header */}
        <div className={`flex items-center justify-between px-4 sm:px-6 py-3 sm:py-4 border-b ${d.border} ${d.surface}`}>
          <h2 className={`text-lg sm:text-xl font-bold ${d.text}`}>
            {replyData?.type === "reply" ? "Reply"
              : replyData?.type === "replyAll" ? "Reply All"
              : replyData?.type === "forward" ? "Forward"
              : "New Message"}
          </h2>
          <button
            onClick={onClose}
            className={`${d.muted} hover:text-gray-300 transition-colors p-1.5 ${d.hover} rounded-lg`}
            aria-label="Close"
          >
            <X size={20} />
          </button>
        </div>

        {/* Form */}
        <form
          onSubmit={handleSubmit}
          className="flex flex-col flex-1 overflow-hidden"
        >
          {/* Selected template + tags, with the Tags picker button (hidden while picking) */}
          {!showTemplatesModal && (tags.length > 0 || templates.length > 1) && (
            <div className={`px-4 sm:px-6 py-3 border-b ${isDark ? "bg-blue-900/20 border-blue-800/40" : "bg-blue-50 border-blue-100"}`}>
              <div className="flex items-start gap-3">
                <button
                  type="button"
                  onClick={() => setShowTemplatesModal(true)}
                  className={`px-3 py-2 ${isDark ? "bg-gray-700 text-blue-400 border-blue-700 hover:bg-gray-600" : "bg-white text-blue-600 border-blue-200 hover:bg-blue-50"} border rounded-lg transition flex items-center gap-2 shadow-sm text-xs sm:text-sm font-medium flex-shrink-0`}
                >
                  <Tag size={16} />
                  Tags
                </button>

                {(selectedTemplate || selectedTags.length > 0) ? (
                  <div className="flex-1 min-w-0 flex flex-wrap content-start gap-2 max-h-[5.5rem] overflow-y-auto pr-1">
                    {selectedTemplate && (
                      <span className="inline-flex items-center gap-1.5 pl-2.5 pr-1 py-1 rounded-full bg-blue-100 text-blue-700 border border-blue-200 text-xs font-medium">
                        <FileText size={12} className="flex-shrink-0" />
                        <span className="truncate max-w-[220px]">{selectedTemplate}</span>
                        <button
                          type="button"
                          onClick={() => setSelectedTemplate("")}
                          className="flex-shrink-0 w-4 h-4 inline-flex items-center justify-center rounded-full hover:bg-blue-200 font-bold leading-none"
                          aria-label="Remove template"
                        >
                          ×
                        </button>
                      </span>
                    )}
                    {selectedTags.map((tag) => (
                      <span
                        key={tag}
                        className="inline-flex items-center gap-1.5 pl-2.5 pr-1 py-1 rounded-full bg-emerald-100 text-emerald-700 border border-emerald-200 text-xs font-medium"
                      >
                        <Tag size={11} className="flex-shrink-0" />
                        <span className="truncate max-w-[200px]">{tag}</span>
                        <button
                          type="button"
                          onClick={() => setSelectedTags(selectedTags.filter((t) => t !== tag))}
                          className="flex-shrink-0 w-4 h-4 inline-flex items-center justify-center rounded-full hover:bg-emerald-200 font-bold leading-none"
                          aria-label="Remove tag"
                        >
                          ×
                        </button>
                      </span>
                    ))}
                  </div>
                ) : (
                  <span className={`flex-1 self-center text-xs ${d.muted}`}>Click the Tags button to add recipients by tag.</span>
                )}
              </div>
            </div>
          )}
          {/* Compose body: optional Templates & Tags side panel (inline) + the fields */}
          <div className="flex flex-1 min-h-0 overflow-hidden">
          {/* Templates & Tags — inline side panel when composing in-pane, modal otherwise */}
          {showTemplatesModal && (
            <div className={inline
              ? `flex-1 min-w-0 h-full flex`
              : "fixed inset-0 bg-black/40 backdrop-blur-sm flex items-center justify-center z-[1200] p-3 sm:p-4"}>
              <div className={inline
                ? `${d.bg} w-full h-full overflow-hidden flex flex-col`
                : `${d.bg} rounded-xl shadow-xl w-full max-w-xl sm:max-w-2xl max-h-[90vh] overflow-hidden flex flex-col`}>
                {/* Modal Header */}
                <div className={`px-4 sm:px-6 py-3 sm:py-4 border-b ${d.border} ${isDark ? "bg-blue-900/20" : "bg-blue-50"}`}>
                  <div className="flex items-center justify-between gap-3">
                    <div className="flex items-center gap-3">
                      <div className="p-2 bg-blue-500 rounded-lg flex-shrink-0">
                        <FileText size={20} className="text-white" />
                      </div>
                      <div className="min-w-0">
                        <h2 className={`text-base sm:text-lg font-bold ${d.text}`}>Tags</h2>
                        <p className={`text-xs sm:text-sm ${d.muted} mt-0.5`}>
                          Pick tags to add their contacts to the recipients
                        </p>
                      </div>
                    </div>
                    <button
                      type="button"
                      onClick={() => setShowTemplatesModal(false)}
                      className={`p-1.5 ${d.hover} rounded-lg transition-colors flex-shrink-0 ${d.muted}`}
                      aria-label="Close"
                    >
                      <X size={20} />
                    </button>
                  </div>
                </div>

                {/* Search Bar */}
                <div className={`px-4 sm:px-6 py-2 sm:py-3 ${d.surface} border-b ${d.border}`}>
                  <div className="relative">
                    <Search size={16} className={`absolute left-3 top-1/2 transform -translate-y-1/2 ${d.muted} flex-shrink-0`} />
                    <input
                      type="text"
                      placeholder="Search tags..."
                      className={`w-full pl-9 pr-3 py-2 border rounded-lg focus:outline-none focus:border-blue-500 focus:ring-2 transition-all text-sm ${d.input}`}
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                    />
                  </div>
                </div>

                {/* Content Area */}
                <div className="flex-1 overflow-y-auto px-4 sm:px-6 py-3 sm:py-4">
                  {/* Templates Section — only when the user has more than one of their own to switch between */}
                  {templates.length > 1 && (
                    <div className="mb-6">
                      <div className="flex items-center gap-2 mb-3">
                        <FileText size={16} className="text-blue-500 flex-shrink-0" />
                        <h3 className={`${d.text} text-sm font-semibold`}>Templates</h3>
                        <span className={`text-xs ${d.muted} ${isDark ? "bg-gray-700" : "bg-gray-100"} px-2 py-0.5 rounded-full ml-auto`}>
                          {filteredTemplates.length}
                        </span>
                      </div>

                      {filteredTemplates.length === 0 ? (
                        <p className={`${d.muted} text-xs sm:text-sm py-3 sm:py-4 text-center ${d.surface} rounded-lg`}>
                          No templates found
                        </p>
                      ) : (
                        <div className="space-y-2">
                          {filteredTemplates.map((tpl) => (
                            <button
                              key={tpl.name}
                              type="button"
                              onClick={() => handleTemplateSelect(tpl.name)}
                              className={`w-full text-left px-3 sm:px-4 py-2 sm:py-3 rounded-lg border transition-all text-sm ${
                                modalSelectedTemplate === tpl.name
                                  ? isDark ? "bg-blue-900/30 border-blue-500" : "bg-blue-50 border-blue-500 shadow-sm"
                                  : `${d.bg} ${d.border} ${d.hoverBlue}`
                              }`}
                            >
                              <div className="flex items-start justify-between gap-2">
                                <div className="flex-1 min-w-0">
                                  <div className="flex items-center gap-2 flex-wrap">
                                    <span className={`${d.text} font-medium text-sm`}>{tpl.name}</span>
                                    {tpl.TemplateType && (
                                      <span className={`text-xs px-2 py-0.5 rounded flex-shrink-0 ${isDark ? "bg-blue-900/40 text-blue-300" : "bg-blue-100 text-blue-700"}`}>
                                        {tpl.TemplateType}
                                      </span>
                                    )}
                                    {modalSelectedTemplate === tpl.name && (
                                      <CheckCircle2 size={16} className="text-blue-500 flex-shrink-0 ml-auto" />
                                    )}
                                  </div>
                                  {(tpl.body ?? tpl.Body) && (
                                    <p className={`${d.muted} text-xs mt-1 line-clamp-2`}>{tpl.body ?? tpl.Body}</p>
                                  )}
                                </div>
                              </div>
                            </button>
                          ))}
                        </div>
                      )}
                    </div>
                  )}

                  {/* Tags Section */}
                  {!tagsLoading && !tagsError && tags.length > 0 && (
                    <div>
                      <div className="flex items-center gap-2 mb-3">
                        <Tag size={16} className="text-emerald-500 flex-shrink-0" />
                        <h3 className={`${d.text} text-sm font-semibold`}>Tags</h3>
                        <div className="ml-auto flex items-center gap-2">
                          <button
                            type="button"
                            onClick={() => {
                              const allSelected = filteredTags.length > 0 && filteredTags.every((t) => modalSelectedTags.includes(t));
                              setModalSelectedTags(allSelected
                                ? modalSelectedTags.filter((t) => !filteredTags.includes(t))
                                : [...new Set([...modalSelectedTags, ...filteredTags])]);
                            }}
                            className="text-xs font-semibold text-emerald-600 hover:text-emerald-700"
                          >
                            {filteredTags.length > 0 && filteredTags.every((t) => modalSelectedTags.includes(t)) ? "Clear all" : "Select all"}
                          </button>
                          <span className={`text-xs ${d.muted} ${isDark ? "bg-gray-700" : "bg-gray-100"} px-2 py-0.5 rounded-full`}>
                            {filteredTags.length}
                          </span>
                        </div>
                      </div>

                      {filteredTags.length === 0 ? (
                        <p className={`${d.muted} text-xs sm:text-sm py-3 sm:py-4 text-center ${d.surface} rounded-lg`}>
                          No tags found
                        </p>
                      ) : (
                        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2">
                          {filteredTags.map((tag) => (
                            <button
                              key={tag}
                              type="button"
                              onClick={() => handleTagToggle(tag)}
                              className={`w-full text-left px-3 sm:px-4 py-2 sm:py-2.5 rounded-lg border transition-all text-sm flex items-center gap-2 ${
                                modalSelectedTags.includes(tag)
                                  ? isDark ? "bg-emerald-900/30 border-emerald-500 text-emerald-300" : "bg-emerald-50 border-emerald-500 text-emerald-700 shadow-sm"
                                  : `${d.bg} ${d.border} ${d.text} hover:border-emerald-400 ${isDark ? "hover:bg-emerald-900/20" : "hover:bg-emerald-50"}`
                              }`}
                            >
                              <Tag size={14} className="flex-shrink-0 text-emerald-500" />
                              <span className="flex-1 truncate font-medium">{tag}</span>
                              {modalSelectedTags.includes(tag) && (
                                <CheckCircle2 size={16} className="flex-shrink-0 text-emerald-500" />
                              )}
                            </button>
                          ))}
                        </div>
                      )}
                    </div>
                  )}

                  {tagsLoading && (
                    <div className={`${d.muted} text-sm text-center py-4`}>
                      Loading tags...
                    </div>
                  )}

                  {tagsError && (
                    <div className="text-red-500 text-sm text-center py-4">
                      {tagsError}
                    </div>
                  )}
                </div>

                {/* Modal Footer */}
                <div className={`px-4 sm:px-6 py-3 sm:py-4 border-t ${d.border} ${d.surface}`}>
                  <div className="flex items-center justify-between gap-3">
                    <button
                      type="button"
                      onClick={handleClearAll}
                      className={`px-3 py-2 ${d.text} rounded-lg transition-colors text-xs sm:text-sm font-medium ${d.hover}`}
                    >
                      Clear All
                    </button>
                    <div className="flex gap-2">
                      <button
                        type="button"
                        onClick={() => setShowTemplatesModal(false)}
                        className={`px-3 py-2 border ${d.border} ${d.text} text-xs sm:text-sm font-medium rounded-lg ${d.hover} transition-colors`}
                      >
                        Cancel
                      </button>
                      <button
                        type="button"
                        onClick={handleApply}
                        className="px-3 py-2 bg-blue-500 text-white text-xs sm:text-sm font-medium rounded-lg hover:bg-blue-600 transition-colors shadow-sm"
                      >
                        Apply
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}

            {/* Compose fields — full width; hidden while the full-width Tags picker is open */}
            <div className={`${inline && showTemplatesModal ? "hidden" : "flex"} flex-col flex-1 min-w-0 overflow-hidden`}>

          {/* Recipients */}
          <div className={`px-4 sm:px-6 py-3`}>
            <div className={`flex items-center gap-3 py-1 border-b-2 border-gray-300 focus-within:border-blue-500 transition-colors`}>
              <span className={`inline-flex items-center px-3 py-1.5 rounded-lg border ${d.border} text-xs sm:text-sm font-semibold ${d.text} flex-shrink-0`}>To</span>
              <div className="flex-1 relative min-w-0">
                {/* chips + input (borderless; the row shows an underline) */}
                <div
                  onClick={() => toInputRef.current && toInputRef.current.focus()}
                  className="flex flex-wrap items-center gap-1.5 w-full cursor-text text-sm"
                >
                  {toEmails.slice(0, 2).map((email, idx) => (
                    <span
                      key={email + idx}
                      className={`${d.tag} pl-2 pr-1 py-0.5 rounded-full text-xs font-medium border flex items-center gap-1 flex-shrink-0`}
                    >
                      <span className="truncate max-w-[160px]">{email}</span>
                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation();
                          setToInput(toEmails.filter((_, i) => i !== idx).join(", "));
                        }}
                        className="hover:text-red-600 font-bold leading-none"
                        aria-label="Remove email"
                      >
                        ×
                      </button>
                    </span>
                  ))}
                  {toEmails.length > 2 && (
                    <button
                      type="button"
                      onClick={(e) => {
                        e.stopPropagation();
                        setShowAllToEmails(true);
                      }}
                      className="bg-blue-50 text-blue-600 px-2 py-0.5 rounded-full text-xs font-medium border border-blue-300 hover:bg-blue-100 transition-colors flex-shrink-0"
                    >
                      +{toEmails.length - 2} more
                    </button>
                  )}
                  {toEmails.length > 0 && (
                    <button
                      type="button"
                      onClick={(e) => {
                        e.stopPropagation();
                        setToInput("");
                        setToDraft("");
                      }}
                      className="text-[11px] text-gray-400 hover:text-red-600 font-semibold flex-shrink-0 px-1"
                      title="Clear all recipients"
                    >
                      Clear all
                    </button>
                  )}
                  <input
                    ref={toInputRef}
                    type="text"
                    placeholder={toEmails.length ? "" : "Recipients"}
                    value={toDraft}
                    onChange={handleToInputChange}
                    onKeyDown={(e) => {
                      if (e.key === "Backspace" && !toDraft && toEmails.length) {
                        setToInput(toEmails.slice(0, -1).join(", "));
                      }
                    }}
                    onFocus={() => {
                      const q = toDraft.trim().toLowerCase();
                      if (q && emailSuggestions.some((em) => em.toLowerCase().includes(q))) {
                        setShowEmailSuggestions(true);
                      }
                    }}
                    className={`flex-1 min-w-[120px] border-0 outline-none bg-transparent px-1 py-1.5 text-sm ${d.text} placeholder-gray-400`}
                  />
                </div>

                {/* Email Suggestions Dropdown */}
                {showEmailSuggestions && filteredEmailSuggestions.length > 0 && (
                  <div className={`absolute top-full left-0 right-0 mt-1 ${d.bg} border ${d.border} rounded-lg shadow-lg z-[100] max-h-[200px] overflow-y-auto`}>
                    {filteredEmailSuggestions.map((email, idx) => (
                      <button
                        key={idx}
                        type="button"
                        onClick={() => handleSelectEmailSuggestion(email)}
                        className={`w-full text-left px-3 py-2 ${d.hoverBlue} border-b ${d.borderSm} last:border-b-0 text-sm ${d.text} transition-colors`}
                      >
                        {email}
                      </button>
                    ))}
                  </div>
                )}
              </div>
              <button
                type="button"
                onClick={() => setShowCc(!showCc)}
                className="text-blue-600 hover:text-blue-700 text-xs sm:text-sm px-2 font-medium flex-shrink-0"
              >
                Cc
              </button>
            </div>
            {errors.to && (
              <p className="text-red-600 text-xs ml-8 sm:ml-[52px] -mt-1 mb-2">
                {errors.to}
              </p>
            )}

            {showCc && (
              <>
                <div className={`flex items-center gap-3 py-1 mt-2 border-b-2 border-gray-300 focus-within:border-blue-500 transition-colors`}>
                  <span className={`inline-flex items-center px-3 py-1.5 rounded-lg border ${d.border} text-xs sm:text-sm font-semibold ${d.text} flex-shrink-0`}>Cc</span>
                  <input
                    type="text"
                    value={ccInput}
                    onChange={(e) => setCcInput(e.target.value)}
                    placeholder="Cc recipients"
                    className={`flex-1 min-w-0 border-0 outline-none bg-transparent px-1 py-1.5 text-sm ${d.text} placeholder-gray-400`}
                  />
                </div>
                {errors.cc && (
                  <p className="text-red-600 text-xs mt-1 mb-1">
                    {errors.cc}
                  </p>
                )}
              </>
            )}

            <div className={`py-1 mt-2 border-b-2 border-gray-300 focus-within:border-blue-500 transition-colors`}>
              <input
                id="email-subject"
                type="text"
                value={subject}
                onChange={(e) => setSubject(e.target.value)}
                placeholder="Add a subject"
                className={`w-full border-0 outline-none bg-transparent px-1 py-1.5 text-sm ${d.text} placeholder-gray-400`}
              />
            </div>
            {errors.subject && (
              <p className="text-red-600 text-xs mt-1">
                {errors.subject}
              </p>
            )}
          </div>

          {/* Message Body */}
          <div className="flex-1 overflow-y-auto px-4 sm:px-6 py-3 sm:py-4 flex flex-col gap-3">
            {/* Message body + footer in a single box (footer = logo on the left + text) */}
            <div
              className={`flex flex-col border rounded-lg overflow-hidden transition-all focus-within:ring-2 focus-within:ring-slate-500 focus-within:border-transparent ${
                quoteHtml ? "min-h-[220px]" : "flex-1 min-h-[480px]"
              } ${isDragging ? "border-blue-500 " + (isDark ? "bg-blue-900/20" : "bg-blue-50") : d.input}`}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
            >
              <textarea
                value={body}
                onChange={(e) => setBody(e.target.value)}
                placeholder={quoteHtml ? "Type your reply here…" : "Type your message here… or drop files"}
                className={`flex-1 w-full resize-none text-base leading-relaxed bg-transparent outline-none border-0 p-3 ${d.text} placeholder-gray-400 ${quoteHtml ? "min-h-[160px]" : "min-h-[340px]"}`}
              />
              {(footerText || footerLogo) && (
                <div className={`flex gap-3 items-center p-3 border-t ${d.border}`}>
                  {footerLogo && (
                    <img src={footerLogo} alt="Company logo" className={`h-16 max-w-[130px] object-contain shrink-0 self-center border-r ${d.border} pr-3`} />
                  )}
                  <textarea
                    value={footerText}
                    onChange={(e) => setFooterText(e.target.value)}
                    rows={Math.max(1, (footerText || "").split("\n").length)}
                    placeholder="Footer / signature"
                    className={`flex-1 w-full resize-none text-[15px] bg-transparent outline-none border-0 p-0 self-center leading-6 ${d.text} placeholder-gray-400`}
                  />
                </div>
              )}
            </div>
            {errors.body && (
              <p className="text-red-600 text-xs -mt-2">{errors.body}</p>
            )}

            {/* Original email quoted block (reply/replyAll mode) */}
            {quoteHtml && replyData && (
              <div className={`border ${d.border} rounded-xl overflow-hidden shadow-sm`}>
                {/* Sender row */}
                <div className={`flex items-start gap-3 px-4 py-3 ${d.surface} border-b ${d.border}`}>
                  <div className={`w-9 h-9 rounded-full ${quoteAvatarColor(replyData.originalMail?.sender)} flex items-center justify-center text-white font-semibold text-sm flex-shrink-0`}>
                    {quoteInitials(replyData.originalMail?.sender)}
                  </div>
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center justify-between gap-2 mb-1">
                      <span className={`text-sm font-semibold ${d.text} truncate`}>{replyData.originalMail?.sender}</span>
                      <span className={`text-[11px] ${d.muted} flex-shrink-0`}>
                        {replyData.originalMail?.receivedDateTime ? new Date(replyData.originalMail.receivedDateTime).toLocaleString() : ""}
                      </span>
                    </div>
                    <div className={`flex flex-wrap items-center gap-1 text-xs ${d.muted}`}>
                      <span className="font-medium">To:</span>
                      {(replyData.originalMail?.toEmail || "").split(/[,;]/).map(e => e.trim()).filter(Boolean).map((e, i) => (
                        <span key={i} className={`px-1.5 py-0.5 rounded ${isDark ? "bg-gray-700 text-gray-300" : "bg-gray-100 text-gray-600"}`}>{e}</span>
                      ))}
                    </div>
                    {replyData.originalMail?.ccEmail && (
                      <div className={`flex flex-wrap items-center gap-1 mt-0.5 text-xs ${d.muted}`}>
                        <span className="font-medium">Cc:</span>
                        {replyData.originalMail.ccEmail.split(/[,;]/).map(e => e.trim()).filter(Boolean).map((e, i) => (
                          <span key={i} className={`px-1.5 py-0.5 rounded ${isDark ? "bg-gray-700 text-gray-300" : "bg-gray-100 text-gray-600"}`}>{e}</span>
                        ))}
                      </div>
                    )}
                  </div>
                </div>

                {/* Attachments row */}
                {replyData.originalMail?.attachments?.length > 0 && (
                  <div className={`px-4 py-2.5 border-b ${d.border} ${d.surface} flex flex-wrap gap-2`}>
                    {replyData.originalMail.attachments.map((att, i) => {
                      const name = att.name || att.fileName || `attachment-${i}`;
                      const sz = quoteFileSize(att.size);
                      return (
                        <div key={i} className={`flex items-center gap-1.5 px-2.5 py-1.5 border ${d.border} rounded-lg ${d.bg} text-xs`}>
                          <Paperclip size={12} className={`${d.muted} flex-shrink-0`} />
                          <span className={`${d.text} font-medium truncate max-w-[140px]`}>{name}</span>
                          {sz && <span className={d.muted}>{sz}</span>}
                        </div>
                      );
                    })}
                  </div>
                )}

                {/* Email body */}
                <div
                  className={`px-4 py-3 text-sm ${d.muted} max-h-52 overflow-y-auto leading-relaxed ${d.bg}`}
                  dangerouslySetInnerHTML={{ __html: quoteHtml }}
                />
              </div>
            )}
          </div>
          

          {/* Attachments Preview */}
          {attachments.length > 0 && (
            <div className={`px-4 sm:px-6 py-2 border-t ${d.border} ${d.surface}`}>
              <p className={`text-xs font-semibold ${d.text} mb-1.5`}>
                Attachments ({attachments.length})
              </p>
              <div className="flex gap-2 overflow-x-auto pb-1">
                {attachments.map((file, idx) => (
                  <div
                    key={idx}
                    className={`flex items-center gap-2 ${d.chip} px-2 py-1 rounded text-xs border flex-shrink-0`}
                  >
                    <span className={`${d.text} truncate text-xs max-w-[160px]`}>{file.name}</span>
                    <button
                      type="button"
                      onClick={() => removeAttachment(idx)}
                      className="text-red-600 hover:text-red-800 font-bold flex-shrink-0"
                      aria-label="Remove attachment"
                    >
                      ×
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Images Preview */}
          {images.length > 0 && (
            <div className={`px-4 sm:px-6 py-2 border-t ${d.border} ${d.surface}`}>
              <p className={`text-xs font-semibold ${d.text} mb-1.5`}>
                Images ({images.length})
              </p>
              <div className="flex gap-2 overflow-x-auto pb-1">
                {images.map((file, idx) => (
                  <div key={idx} className="relative group flex-shrink-0">
                    <img
                      src={URL.createObjectURL(file)}
                      alt={file.name}
                      className="h-12 w-12 object-cover rounded border border-gray-300"
                    />
                    <button
                      type="button"
                      onClick={() => removeImage(idx)}
                      className="absolute -top-1.5 -right-1.5 bg-red-600 text-white rounded-full w-4 h-4 flex items-center justify-center opacity-0 group-hover:opacity-100 transition text-[10px] leading-none"
                      aria-label="Remove image"
                    >
                      ×
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Error & Success Messages */}
          {errors.apiError && (
            <div className="px-4 sm:px-6 py-2 sm:py-3 bg-red-50 border-t border-red-200">
              <p className="text-red-600 text-xs sm:text-sm">{errors.apiError}</p>
            </div>
          )}

          {successMessage && (
            <div className="px-4 sm:px-6 py-2 sm:py-3 bg-emerald-50 border-t border-emerald-200">
              <p className="text-emerald-700 text-xs sm:text-sm">{successMessage}</p>
            </div>
          )}

          {/* Footer */}
          <div className={`px-4 sm:px-6 py-3 border-t ${d.border} flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3 ${d.surface}`}>
            <div className="flex items-center gap-1 flex-wrap">
              {/* Send split-button with dropdown */}
              <div className="relative flex-shrink-0">
                <div className="flex rounded-lg overflow-hidden shadow-sm">
                  <button
                    type="submit"
                    disabled={isSending}
                    className="flex items-center gap-2 bg-blue-500 text-white px-3 sm:px-4 py-2 hover:bg-blue-600 transition disabled:bg-slate-400 disabled:cursor-not-allowed text-xs sm:text-sm font-medium h-10 sm:h-auto"
                  >
                    <Send size={15} />
                    {isSending ? "Sending..." : "Send"}
                  </button>
                  <button
                    type="button"
                    onClick={() => setSendDropdown(p => !p)}
                    className="flex items-center px-2 bg-blue-500 hover:bg-blue-600 text-white border-l border-slate-600 transition h-10 sm:h-auto"
                  >
                    <ChevronDown size={14} />
                  </button>
                </div>

                {sendDropdown && (
                  <div className={`absolute left-0 bottom-full mb-1 w-48 ${d.bg} border ${d.border} rounded-xl shadow-xl z-[60] overflow-visible`}
                    onMouseLeave={() => setSendDropdown(false)}>
                    <button type="submit" onClick={() => setSendDropdown(false)}
                      className={`w-full flex items-center gap-3 px-4 py-2.5 text-sm ${d.text} ${d.hover} transition`}>
                      <Send size={14} className="text-blue-500 flex-shrink-0" />Send
                    </button>
                    <button type="button" onClick={() => { setSendDropdown(false); setScheduleDateTime(""); setScheduleModal(true); }}
                      className={`w-full flex items-center gap-3 px-4 py-2.5 text-sm ${d.text} ${d.hover} transition`}>
                      <Clock size={14} className="text-blue-500 flex-shrink-0" />Schedule send
                    </button>
                    <button type="button" onClick={() => { setSendDropdown(false); handleMailMerge(); }}
                      className={`w-full flex items-center gap-3 px-4 py-2.5 text-sm ${d.text} ${d.hover} transition`}>
                      <Users size={14} className="text-blue-500 flex-shrink-0" />Start mail merge
                    </button>
                    <button type="button" onClick={() => { setSendDropdown(false); handleTrackedCampaign(); }}
                      className={`w-full flex items-center gap-3 px-4 py-2.5 text-sm ${d.text} ${d.hover} transition`}>
                      <Send size={14} className="text-emerald-500 flex-shrink-0" />Send tracked campaign
                    </button>
                  </div>
                )}
              </div>

              <button
                type="button"
                onClick={() => attachmentInputRef.current?.click()}
                className={`p-2 ${d.muted} ${d.hover} rounded-lg transition h-10 w-10 flex items-center justify-center`}
                title="Attach file"
                aria-label="Attach file"
              >
                <Paperclip size={18} />
              </button>
              <input
                ref={attachmentInputRef}
                type="file"
                multiple
                onChange={handleAttachmentSelect}
                className="hidden"
              />

              <button
                type="button"
                className={`p-2 ${d.muted} ${d.hover} rounded-lg transition h-10 w-10 flex items-center justify-center`}
                title="Insert emoji"
                aria-label="Insert emoji"
              >
                <Smile size={18} />
              </button>

              <button
                type="button"
                onClick={() => imageInputRef.current?.click()}
                className={`p-2 ${d.muted} ${d.hover} rounded-lg transition h-10 w-10 flex items-center justify-center`}
                title="Insert image"
                aria-label="Insert image"
              >
                <ImageIcon size={18} />
              </button>
              <input
                ref={imageInputRef}
                type="file"
                multiple
                accept="image/*"
                onChange={handleImageSelect}
                className="hidden"
              />
            </div>

            <button
              type="button"
              onClick={onClose}
              className={`${d.muted} px-3 sm:px-4 py-2 ${d.hover} rounded-lg transition text-xs sm:text-sm font-medium`}
            >
              Discard
            </button>
          </div>
            </div>{/* /compose fields column */}
          </div>{/* /compose body row */}
        </form>

        {/* All Recipients Modal */}
        {showAllToEmails && (
          <div className="fixed inset-0 bg-black/40 backdrop-blur-sm z-[200] flex items-center justify-center p-3 sm:p-4">
            <div className={`${d.bg} rounded-xl shadow-xl w-full max-w-lg max-h-[90vh] overflow-hidden flex flex-col`}>
              {/* Modal Header */}
              <div className={`px-4 sm:px-6 py-3 sm:py-4 border-b ${d.border} flex items-center justify-between gap-3`}>
                <h2 className={`text-base sm:text-lg font-bold ${d.text}`}>All Recipients ({toEmails.length})</h2>
                <button
                  onClick={() => setShowAllToEmails(false)}
                  className={`${d.muted} p-1.5 ${d.hover} rounded-lg transition flex-shrink-0`}
                  aria-label="Close"
                >
                  <X size={20} />
                </button>
              </div>

              {/* Modal Body - Scrollable */}
              <div className="flex-1 overflow-y-auto px-4 sm:px-6 py-3 sm:py-4">
                {/* Input field for adding more emails */}
                <div className="mb-4 relative">
                  <input
                    type="text"
                    placeholder="Add more emails..."
                    value={toDraft}
                    onChange={handleToInputChange}
                    onFocus={() => setShowEmailSuggestions(filteredEmailSuggestions.length > 0)}
                    className={`w-full px-3 py-2 border rounded-lg focus:outline-none focus:border-blue-500 focus:ring-2 transition-all text-sm ${d.input}`}
                  />

                  {/* Suggestions in modal */}
                  {showEmailSuggestions && filteredEmailSuggestions.length > 0 && (
                    <div className={`absolute top-full left-0 right-0 mt-1 ${d.bg} border ${d.border} rounded-lg shadow-lg z-[210] max-h-[150px] overflow-y-auto`}>
                      {filteredEmailSuggestions.map((email, idx) => (
                        <button
                          key={idx}
                          type="button"
                          onClick={() => handleSelectEmailSuggestion(email)}
                          className={`w-full text-left px-3 py-2 ${d.hoverBlue} border-b ${d.borderSm} last:border-b-0 text-xs sm:text-sm ${d.text} transition-colors`}
                        >
                          {email}
                        </button>
                      ))}
                    </div>
                  )}
                </div>

                {/* List of all emails */}
                <div className="space-y-2">
                  {toEmails.map((email, idx) => (
                    <div
                      key={email + idx}
                      className={`flex items-center justify-between p-2 sm:p-3 rounded-lg border ${isDark ? "bg-blue-900/20 border-blue-800" : "bg-blue-50 border-blue-200"}`}
                    >
                      <span className={`text-xs sm:text-sm ${d.text} flex-1 break-all`}>{email}</span>
                      <button
                        type="button"
                        onClick={() => {
                          setToInput(toEmails.filter((_, i) => i !== idx).join(", "));
                        }}
                        className="ml-2 text-red-600 hover:text-red-700 hover:bg-red-100 p-1 rounded transition flex-shrink-0"
                        title="Remove email"
                        aria-label="Remove email"
                      >
                        <X size={16} />
                      </button>
                    </div>
                  ))}
                </div>

                {toEmails.length === 0 && (
                  <p className={`text-center ${d.muted} text-xs sm:text-sm py-6 sm:py-8`}>No recipients added yet</p>
                )}
              </div>

              {/* Modal Footer */}
              <div className={`px-4 sm:px-6 py-3 sm:py-4 border-t ${d.border} flex items-center justify-end gap-2 ${d.surface}`}>
                <button
                  type="button"
                  onClick={() => setShowAllToEmails(false)}
                  className={`px-3 py-2 ${d.text} border ${d.border} rounded-lg ${d.hover} transition text-xs sm:text-sm font-medium`}
                >
                  Close
                </button>
              </div>
            </div>
          </div>
        )}

        {/* ── Schedule Send Modal ── */}
        {scheduleModal && (
          <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-[1400] p-4">
            <div className={`${d.bg} rounded-xl shadow-xl w-full max-w-sm overflow-hidden border ${d.border}`}>
              <div className={`px-5 py-4 border-b ${d.border} flex items-center justify-between`}>
                <h3 className={`text-base font-semibold ${d.text} flex items-center gap-2`}>
                  <Clock size={16} className="text-blue-500" />Schedule send
                </h3>
                <button type="button" onClick={() => setScheduleModal(false)}
                  className={`p-1.5 rounded-lg ${d.hover} ${d.muted} transition`}>
                  <X size={14} />
                </button>
              </div>
              <div className="px-5 py-4">
                <label className={`text-xs font-medium ${d.muted} mb-2 block`}>Select date and time</label>
                <input
                  type="datetime-local"
                  value={scheduleDateTime}
                  onChange={e => setScheduleDateTime(e.target.value)}
                  min={new Date(Date.now() + 60000).toISOString().slice(0, 16)}
                  className={`w-full px-3 py-2 border rounded-lg text-sm outline-none focus:border-blue-500 focus:ring-2 ${d.input}`}
                />
                {errors.apiError && <p className="mt-2 text-xs text-red-500">{errors.apiError}</p>}
              </div>
              <div className={`px-5 py-3 border-t ${d.border} flex items-center justify-end gap-2 ${d.surface}`}>
                <button type="button" onClick={() => setScheduleModal(false)}
                  className={`px-4 py-2 text-sm font-medium ${d.muted} ${d.hover} rounded-lg transition`}>
                  Cancel
                </button>
                <button type="button" onClick={handleScheduleSend}
                  disabled={!scheduleDateTime || schedulingInProgress}
                  className="flex items-center gap-2 px-4 py-2 bg-blue-500 hover:bg-blue-600 text-white text-sm font-medium rounded-lg transition disabled:opacity-50">
                  <Clock size={14} />{schedulingInProgress ? "Scheduling…" : "Schedule"}
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
