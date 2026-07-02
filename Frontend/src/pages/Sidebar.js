import React, { useState, useEffect, useRef } from "react";
import { useNavigate } from "react-router-dom";
import { useMsal } from "@azure/msal-react";
//import WindowLogo from "../assets/Window.png";
import { FaRecycle } from "react-icons/fa";

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

const SettingsIcon = ({ size = 16, className = "" }) => (
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
    <circle cx="12" cy="12" r="3" />
    <path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 01-2.83 2.83l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-4 0v-.09a1.65 1.65 0 00-1-1.51 1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83-2.83l.06-.06a1.65 1.65 0 00.33-1.82 1.65 1.65 0 00-1.51-1H3a2 2 0 010-4h.09a1.65 1.65 0 001.51-1 1.65 1.65 0 00-.33-1.82L4.31 4.1A2 2 0 017.14 1.27l.06.06a1.65 1.65 0 001.82.33H9a1.65 1.65 0 001-1.51V3a2 2 0 014 0v.09c.09.59.56 1.09 1.15 1.33.59.24 1.29.1 1.82-.33l.06-.06A2 2 0 0119.4 4.1l-.06.06a1.65 1.65 0 00-.33 1.82V7c.41.22.79.5 1.13.83.34.33.62.72.85 1.14.23.53.08 1.23-.33 1.82-.41.6-.41 1.34 0 1.95z" />
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

function Sidebar({ activeContent, openOrActivateTab, setShowPanel }) {
  const { instance, accounts } = useMsal();
  const navigate = useNavigate();
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const sidebarRef = useRef(null);

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

  const navItems = [
    { id: "home", label: "Home", icon: HomeIcon, path: "/dashboard" },
    { id: "accounts", label: "Accounts", icon: AccountsIcon, path: "/dashboard/Accounts" },
    { id: "contacts", label: "Contacts", icon: ContactsIcon, path: "/dashboard/contacts" },
    { id: "products", label: "Products", icon: ProductsIcon, path: "/dashboard/Products" },
    { id: "deals", label: "Deals", icon: DealsIcon, path: "/dashboard/Deals" },
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
    { id: "teams", label: "Teams", icon: TeamsIcon, path: "/dashboard/Teams" },
  ];

  // Recycle Bin lives at the bottom of the sidebar, just above Settings
  const recycleBinItem = {
    id: "recycle-bin",
    label: "Recycle Bin",
    icon: RecycleBinIcon,
    path: "/dashboard/recycle-bin",
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
        {/* Sidebar Header */}
        <div className="h-[70px] px-3 border-b border-slate-200 flex items-center justify-center flex-shrink-0">
          <button
            onClick={() => setSidebarCollapsed(!sidebarCollapsed)}
            className={`
              hidden lg:flex items-center justify-center
              transition-all duration-300
              rounded-lg
              ${
                sidebarCollapsed
                  ? "w-10 h-10 bg-gray-100 text-blue-500 hover:bg-gray-150"
                  : "w-full h-10 px-3 bg-gray-100 text-blue-500 hover:bg-gray-150"
              }
            `}
            aria-label={sidebarCollapsed ? "Expand sidebar" : "Collapse sidebar"}
            title={sidebarCollapsed ? "Expand sidebar" : "Collapse sidebar"}
          >
            <span className={`transition-transform duration-300 ${sidebarCollapsed ? "" : "rotate-180"}`}>
              <FiTripleChevron size={18} />
            </span>

            {!sidebarCollapsed && <span className="ml-2 text-sm font-semibold">Collapse</span>}
          </button>
        </div>

        {/* Navigation Items */}
        <div className="flex-1 overflow-y-auto p-3 space-y-2">
          {navItems.map((item) => {
            const Icon = item.icon;
            const isActive = activeContent === item.id;
            return (
              <a
                key={item.id}
                href={item.path}
                onClick={(e) => {
                  e.preventDefault();
                  if (item.onClick) {
                    item.onClick();
                  } else {
                    handleNavClick(item.id, item.path);
                  }
                }}
                className={`
                  w-full flex items-center gap-2 px-2 py-1.5 rounded-lg transition-all duration-200 cursor-pointer
                  ${
                    isActive
                      ? "bg-blue-100 text-black/65 border border-gray-300 font-medium"
                      : "text-slate-600 hover:bg-slate-100 transition-colors"
                  }
                `}
                title={item.label}
              >
                <Icon size={16} className="flex-shrink-0" />
                {!sidebarCollapsed && (
                  <span className="text-xs font-medium truncate">{item.label}</span>
                )}
              </a>
            );
          })}
        </div>

        {/* Divider */}
        <div className="mx-3 my-3 border-t border-slate-200" />

        {/* Bottom section: Recycle Bin, then Settings */}
        <div className="px-3 pb-4 flex-shrink-0 space-y-1">
          {/* Recycle Bin */}
          <a
            href={recycleBinItem.path}
            onClick={(e) => {
              e.preventDefault();
              handleNavClick(recycleBinItem.id, recycleBinItem.path);
            }}
            className={`
              w-full flex items-center gap-2 px-2 py-1.5 rounded-lg transition-all duration-200 cursor-pointer
              ${
                activeContent === recycleBinItem.id
                  ? "bg-blue-100 text-black/65 border border-gray-300 font-medium"
                  : "text-slate-600 hover:bg-slate-100 transition-colors"
              }
            `}
            title={recycleBinItem.label}
          >
            <RecycleBinIcon size={16} className="flex-shrink-0" />
            {!sidebarCollapsed && (
              <span className="text-xs font-medium truncate">{recycleBinItem.label}</span>
            )}
          </a>

          {/* Settings (opens profile as requested) */}
          <button
            onClick={() => {
              navigate("/dashboard/profile");
              setShowPanel(false);
              setSidebarOpen(false);
            }}
            className="flex items-center gap-2 w-full text-slate-600 hover:bg-slate-100 px-2 py-2 rounded"
            title="Settings"
          >
            <SettingsIcon className="text-slate-900" />
            {!sidebarCollapsed && <span className="text-xs font-medium">Settings</span>}
          </button>
        </div>
      </nav>
    </>
  );
}

export default Sidebar;