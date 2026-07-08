import React, { useState, useEffect, useRef } from "react";
import { useNavigate } from "react-router-dom";
import { useMsal } from "@azure/msal-react";
//import WindowLogo from "../assets/Window.png";
import { FaRecycle } from "react-icons/fa";
import tokenUtils from "../auth/tokenUtils";

/* ── UI SVG icons ───────────────────────────────────────────── */
const FiMenu = ({ size = 24 }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
  >
    <line x1="3" y1="6" x2="21" y2="6" />
    <line x1="3" y1="12" x2="21" y2="12" />
    <line x1="3" y1="18" x2="21" y2="18" />
  </svg>
);

const FiTripleChevron = ({ size = 18, className = "" }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    className={className}
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
  >
    <polyline points="11 17 6 12 11 7" />
    <polyline points="17 17 12 12 17 7" />
    <polyline points="23 17 18 12 23 7" />
  </svg>
);


/* ── Custom branded SVG icons ───────────────────────────────── */
const HomeIcon = ({ size = 16, className = "" }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <path d="M3 9.5L12 3l9 6.5V20a1 1 0 01-1 1H5a1 1 0 01-1-1V9.5z" />
    <path d="M9 21V12h6v9" />
  </svg>
);

const AccountsIcon = ({ size = 16, className = "" }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <rect x="2" y="7" width="20" height="14" rx="2" />
    <path d="M16 7V5a2 2 0 00-2-2h-4a2 2 0 00-2 2v2" />
    <line x1="12" y1="12" x2="12" y2="16" />
    <line x1="10" y1="14" x2="14" y2="14" />
  </svg>
);

const ContactsIcon = ({ size = 16, className = "" }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2" />
    <circle cx="9" cy="7" r="4" />
    <path d="M23 21v-2a4 4 0 00-3-3.87M16 3.13a4 4 0 010 7.75" />
  </svg>
);

const ProductsIcon = ({ size = 16, className = "" }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <path d="M21 16V8a2 2 0 00-1-1.73l-7-4a2 2 0 00-2 0l-7 4A2 2 0 002 8v8a2 2 0 001 1.73l7 4a2 2 0 002 0l7-4A2 2 0 0021 16z" />
    <polyline points="3.27 6.96 12 12.01 20.73 6.96" />
    <line x1="12" y1="22.08" x2="12" y2="12" />
  </svg>
);

const DealsIcon = ({ size = 16, className = "" }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <polyline points="22 7 13.5 15.5 8.5 10.5 2 17" />
    <polyline points="16 7 22 7 22 13" />
  </svg>
);

const OutlookIcon = ({ size = 16, className = "" }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <rect x="2" y="4" width="20" height="16" rx="2" />
    <polyline points="2,4 12,13 22,4" />
  </svg>
);

const TeamsIcon = ({ size = 16, className = "" }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2" />
    <circle cx="9" cy="7" r="4" />
    <path d="M23 21v-2a4 4 0 00-3-3.87" />
    <path d="M16 3.13a4 4 0 010 7.75" />
  </svg>
);

const RecycleBinIcon = ({ size = 16, className = "" }) => (
  <FaRecycle size={size} className={className} />
);

const ProfileIcon = ({ size = 16, className = "" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
    <path d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z" />
  </svg>
);

const ImportIcon = ({ size = 16, className = "" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
    <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4" />
    <polyline points="7 10 12 15 17 10" />
    <line x1="12" y1="15" x2="12" y2="3" />
  </svg>
);

const ManageUsersIcon = ({ size = 16, className = "" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
    <path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2" />
    <circle cx="9" cy="7" r="4" />
    <path d="M23 21v-2a4 4 0 00-3-3.87M16 3.13a4 4 0 010 7.75" />
  </svg>
);

const AuditIcon = ({ size = 16, className = "" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
    <path d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-3 7h3m-3 4h3m-6-4h.01M9 16h.01" />
  </svg>
);

const BellIcon = ({ size = 16, className = "" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
    <path d="M18 8A6 6 0 006 8c0 7-3 9-3 9h18s-3-2-3-9" />
    <path d="M13.73 21a2 2 0 01-3.46 0" />
  </svg>
);

const ActivityIcon = ({ size = 16, className = "" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
    <polyline points="22 12 18 12 15 21 9 3 6 12 2 12" />
  </svg>
);

function Sidebar({ activeContent, openOrActivateTab, setShowPanel }) {
  const { instance, accounts } = useMsal();
  const navigate = useNavigate();
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const sidebarRef = useRef(null);

  // Admin-only sidebar items (Import Data, Manage Users, Audit Logs) — read role from the JWT.
  const isAdmin = (() => {
    try {
      const t = tokenUtils.getToken();
      const d = t ? tokenUtils.decodeToken(t) : null;
      return (d?.role || d?.Role || "").toString().toLowerCase() === "admin";
    } catch (_) {
      return false;
    }
  })();

  // Close sidebar when clicking outside on mobile
  useEffect(() => {
    function handleClickOutside(event) {
      if (sidebarRef.current && !sidebarRef.current.contains(event.target)) {
        setSidebarOpen(false);
      }
    }

    if (sidebarOpen) {
      document.addEventListener("mousedown", handleClickOutside);
    }

    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, [sidebarOpen]);

  const handleNavClick = (tabId, path) => {
    openOrActivateTab(tabId);
    setShowPanel(false);
    navigate(path);
    setSidebarOpen(false); // Close sidebar on mobile after navigation
  };

  // Sidebar is segregated by priority into labelled sections. Admin-only items
  // (admin: true) are hidden for non-admins; a section with no visible items is
  // dropped entirely (heading included).
  const navSections = [
    {
      heading: "Main",
      items: [
        { id: "home", label: "Home", icon: HomeIcon, path: "/dashboard" },
        { id: "accounts", label: "Accounts", icon: AccountsIcon, path: "/dashboard/Accounts" },
        { id: "contacts", label: "Contacts", icon: ContactsIcon, path: "/dashboard/contacts" },
        { id: "products", label: "Products", icon: ProductsIcon, path: "/dashboard/Products" },
        { id: "deals", label: "Deals", icon: DealsIcon, path: "/dashboard/Deals" },
      ],
    },
    {
      heading: "Communication",
      items: [
        {
          id: "outlookEmail",
          label: "Email",
          icon: OutlookIcon,
          path: "/dashboard/OutlookEmail",
          onClick: async () => {
            if (accounts.length === 0) {
              try {
                await instance.loginRedirect({
                  scopes: [
                    "https://graph.microsoft.com/User.Read",
                    "https://graph.microsoft.com/Mail.Read",
                    "https://graph.microsoft.com/Mail.ReadWrite",
                    "https://graph.microsoft.com/Mail.Send",
                  ],
                });
                handleNavClick("outlookEmail", "/dashboard/OutlookEmail");
              } catch (error) {
                console.error("Login failed:", error);
              }
            } else {
              handleNavClick("outlookEmail", "/dashboard/OutlookEmail");
            }
          },
        },
        {
          id: "emailTracking",
          label: "Email Tracking",
          icon: ({ size = 16, className = "" }) => (
            <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
              <path d="M3 3v18h18" />
              <rect x="7" y="11" width="3" height="7" />
              <rect x="12" y="7" width="3" height="11" />
              <rect x="17" y="4" width="3" height="14" />
            </svg>
          ),
          path: "/dashboard/EmailTracking",
        },
        { id: "notifications", label: "Notifications", icon: BellIcon, path: "/dashboard/notifications" },
        { id: "activity", label: "Recent Activity", icon: ActivityIcon, path: "/dashboard/activity" },
        { id: "teams", label: "Teams", icon: TeamsIcon, path: "/dashboard/Teams" },
      ],
    },
    {
      heading: "Administration",
      items: [
        // Moved out of the header account menu → now live in the sidebar (Log out stays in the header).
        {
          id: "import",
          label: "Import Data",
          icon: ImportIcon,
          path: "/Dashboard/Import",
          admin: true,
          onClick: () => {
            navigate("/Dashboard/Import");
            setShowPanel(false);
            setSidebarOpen(false);
          },
        },
        { id: "users", label: "Manage Users", icon: ManageUsersIcon, path: "/dashboard/users", admin: true },
        { id: "auditLogs", label: "Audit Logs", icon: AuditIcon, path: "/dashboard/audit-logs", admin: true },
      ],
    },
  ];

  // Bottom utility items (below a divider, no heading). Profile replaces the old
  // "Settings" button — both opened the same page, so only one is kept.
  const bottomItems = [
    {
      id: "profile",
      label: "Profile",
      icon: ProfileIcon,
      path: "/Dashboard/profile",
      onClick: () => {
        navigate("/Dashboard/profile");
        setShowPanel(false);
        setSidebarOpen(false);
      },
    },
    { id: "recycle-bin", label: "Recycle Bin", icon: RecycleBinIcon, path: "/dashboard/recycle-bin" },
  ];

  // Shared nav-link renderer so sections and the bottom utilities stay identical.
  const renderNavLink = (item) => {
    const Icon = item.icon;
    const isActive = activeContent === item.id;
    return (
      <a
        key={item.id}
        href={item.path}
        onClick={(e) => {
          e.preventDefault();
          if (item.onClick) item.onClick();
          else handleNavClick(item.id, item.path);
        }}
        className={`
          group relative w-full flex items-center gap-2.5 rounded-lg text-[13px]
          transition-colors duration-150 cursor-pointer
          ${sidebarCollapsed ? "justify-center py-2" : "px-2.5 py-2"}
          ${
            isActive
              ? "bg-slate-100 text-slate-900 font-medium"
              : "text-slate-600 font-normal hover:bg-slate-100/70 hover:text-slate-900"
          }
        `}
        title={item.label}
      >
        {/* Active indicator — a soft neutral bar, not a filled blue box */}
        {isActive && !sidebarCollapsed && (
          <span className="absolute left-0 top-1/2 -translate-y-1/2 h-5 w-[3px] rounded-r-full bg-slate-700" />
        )}
        <Icon
          size={17}
          className={`flex-shrink-0 transition-colors ${
            isActive ? "text-slate-700" : "text-slate-400 group-hover:text-slate-600"
          }`}
        />
        {!sidebarCollapsed && <span className="truncate">{item.label}</span>}
      </a>
    );
  };

  return (
    <>
      {/* Mobile Hamburger Button - visible only on mobile */}
      <button
        onClick={() => setSidebarOpen(!sidebarOpen)}
        className="lg:hidden fixed left-1/2 bottom-4 z-[120] -translate-x-1/2 bg-transparent p-0 text-slate-500 hover:text-slate-700"
        aria-label="Toggle navigation menu"
      >
        <FiMenu size={26} />
      </button>

      {/* Mobile Sidebar Overlay - visible only on mobile */}
      {sidebarOpen && (
        <div
          className="fixed inset-0 bg-black/50 lg:hidden z-[90]"
          onClick={() => setSidebarOpen(false)}
        />
      )}

      {/* Sidebar Navigation */}
      <nav
        ref={sidebarRef}
        className={`
          fixed lg:relative
          left-0 top-14 lg:top-0
          h-[calc(125vh-3.5rem)] lg:h-full
          bg-white backdrop-blur-sm
          border-r border-slate-200
          z-[100] lg:z-10
          transform transition-all duration-300 ease-in-out
          flex flex-col
          ${sidebarOpen ? "translate-x-0 w-44" : "-translate-x-full lg:translate-x-0"}
          ${sidebarCollapsed ? "lg:w-16" : "lg:w-44"}
        `}
      >
        {/* Sidebar Header — collapse toggle (neutral ghost button, no accent colour).
            py-4 + a 36px button = 68px, matching the content tab-bar header so the
            two bottom borders line up exactly. */}
        <div className="px-2.5 py-4 border-b border-slate-200 flex items-center flex-shrink-0">
          <button
            onClick={() => setSidebarCollapsed(!sidebarCollapsed)}
            className={`
              hidden lg:flex items-center rounded-lg
              text-slate-500 hover:bg-slate-100 hover:text-slate-800
              transition-colors duration-200
              ${sidebarCollapsed ? "w-9 h-9 mx-auto justify-center" : "w-full h-9 px-2.5 justify-start"}
            `}
            aria-label={sidebarCollapsed ? "Expand sidebar" : "Collapse sidebar"}
            title={sidebarCollapsed ? "Expand sidebar" : "Collapse sidebar"}
          >
            <span className={`transition-transform duration-300 ${sidebarCollapsed ? "" : "rotate-180"}`}>
              <FiTripleChevron size={16} />
            </span>
            {!sidebarCollapsed && <span className="ml-2 text-[13px] font-medium">Collapse</span>}
          </button>
        </div>

        {/* Navigation Items — grouped by priority into labelled sections */}
        <div className="flex-1 overflow-y-auto p-3 space-y-1">
          {navSections.map((section, si) => {
            const items = section.items.filter((item) => !item.admin || isAdmin);
            if (items.length === 0) return null;
            return (
              <div key={section.heading}>
                {/* Section heading (or a divider when collapsed to icons) */}
                {sidebarCollapsed ? (
                  si > 0 && <div className="mx-2 my-2.5 border-t border-slate-200/80" />
                ) : (
                  <p className={`px-2.5 pb-1.5 text-[10px] font-semibold uppercase tracking-[0.1em] text-slate-400 select-none ${si === 0 ? "pt-1" : "pt-4"}`}>
                    {section.heading}
                  </p>
                )}
                <div className="space-y-0.5">
                  {items.map((item) => renderNavLink(item))}
                </div>
              </div>
            );
          })}
        </div>

        {/* Divider */}
        <div className="mx-3 my-2 border-t border-slate-200" />

        {/* Bottom section: Profile, then Recycle Bin */}
        <div className="px-3 pb-4 flex-shrink-0 space-y-1">
          {bottomItems.map((item) => renderNavLink(item))}
        </div>
      </nav>
    </>
  );
}

export default Sidebar;