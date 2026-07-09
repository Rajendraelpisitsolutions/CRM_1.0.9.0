import React, { useEffect, useRef, useState } from "react";
import { useSearchParams } from "react-router-dom";
import apiClient from "../api/client";

// Public page recipients land on from the Subscribe / Unsubscribe buttons in an email.
// It records the choice through the working /api path and shows a confirmation — so it
// never depends on a plain /api/track/* path being reachable in production.

const COMPANY = "Elpis IT Solutions";

export default function EmailSubscription({ mode }) {
  const [params] = useSearchParams();
  const token = params.get("token");
  const [status, setStatus] = useState("loading"); // loading | done | invalid | error
  const ran = useRef(false);

  useEffect(() => {
    if (ran.current) return; // guard against double-invoke in StrictMode
    ran.current = true;
    if (!token) { setStatus("invalid"); return; }
    const url =
      mode === "subscribe"
        ? "/EmailCampaign/subscription/subscribe"
        : "/EmailCampaign/subscription/unsubscribe";
    apiClient
      .post(url, { token })
      .then((r) => setStatus(r?.data?.found === false ? "invalid" : "done"))
      .catch(() => setStatus("error"));
  }, [token, mode]);

  const isUnsub = mode !== "subscribe";

  const heading =
    status === "invalid" ? "Link invalid or expired"
      : status === "error" ? "Something went wrong"
        : isUnsub ? "Thanks for confirming"
          : "You're subscribed!";

  const message =
    status === "invalid" ? "This subscription link is invalid or has already expired. You can close this window."
      : status === "error" ? "We couldn't process your request just now. Please try the link again in a moment."
        : isUnsub
          ? <>You have been unsubscribed and will <strong>not receive any further emails</strong> from <strong>{COMPANY}</strong>.</>
          : <>You'll continue to receive further emails related to <strong>{COMPANY}</strong>.</>;

  const accent = status === "invalid" || status === "error" ? "#9ca3af" : isUnsub ? "#4b5563" : "#2563eb";
  const tint = status === "invalid" || status === "error" ? "#f3f4f6" : isUnsub ? "#f3f4f6" : "#eff6ff";

  return (
    <div style={{ minHeight: "100vh", background: "#f8fafc", display: "flex", alignItems: "center", justifyContent: "center", padding: 16, fontFamily: "Arial, Helvetica, sans-serif" }}>
      <div style={{ maxWidth: 460, width: "100%", background: "#fff", border: "1px solid #e5e7eb", borderRadius: 18, padding: 40, textAlign: "center", boxShadow: "0 10px 30px rgba(0,0,0,0.06)" }}>
        <div style={{ fontSize: 20, fontWeight: 700, color: "#111827", marginBottom: 20 }}>{COMPANY}</div>

        {status === "loading" ? (
          <>
            <div style={{ width: 40, height: 40, margin: "8px auto 16px", border: "3px solid #e5e7eb", borderTopColor: "#2563eb", borderRadius: "50%", animation: "es-spin 0.8s linear infinite" }} />
            <p style={{ color: "#6b7280", margin: 0 }}>Processing your request…</p>
            <style>{"@keyframes es-spin{to{transform:rotate(360deg)}}"}</style>
          </>
        ) : (
          <>
            <div style={{ width: 60, height: 60, borderRadius: "50%", background: tint, color: accent, fontSize: 30, lineHeight: "60px", margin: "0 auto 16px" }}>
              {status === "invalid" || status === "error" ? "!" : "✓"}
            </div>
            <h2 style={{ color: "#111827", margin: "6px 0 10px", fontSize: 22 }}>{heading}</h2>
            <p style={{ color: "#6b7280", lineHeight: 1.6, margin: 0 }}>{message} You can close this window.</p>
          </>
        )}
      </div>
    </div>
  );
}
