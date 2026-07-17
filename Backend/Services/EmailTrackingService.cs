using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Elpis_CRM.Data;
using Elpis_CRM.Model;
using Microsoft.EntityFrameworkCore;

namespace Elpis_CRM.Services
{
    /// <summary>
    /// Persistence + logic for email tracking: creating campaigns and their per-recipient rows,
    /// weaving the open pixel / click-through links / unsubscribe link into each recipient's HTML,
    /// recording the open/click/unsubscribe hits, and producing the analytics rollups.
    /// </summary>
    public class EmailTrackingService
    {
        private readonly AppDbContext _db;

        public EmailTrackingService(AppDbContext db)
        {
            _db = db;
        }

        /// <summary>True when the address is a well-formed email: exactly one @, a non-empty local
        /// part and a domain ending in a 2+ letter TLD, with no whitespace. A false result → BOUNCE.</summary>
        public static bool IsValidEmailFormat(string? email)
        {
            if (string.IsNullOrWhiteSpace(email)) return false;
            var e = email.Trim();
            if (e.Contains(' ') || e.Contains('\t')) return false;
            return Regex.IsMatch(e, @"^[^@\s]+@[^@\s]+\.[A-Za-z]{2,}$");
        }

        // ─── CREATE ──────────────────────────────────────────────────────────────

        /// <summary>
        /// Creates a queued campaign and one Pending recipient row (with a fresh tracking token)
        /// per address. De-duplicates addresses case-insensitively. The background sender picks
        /// the campaign up from here.
        /// </summary>
        public async Task<EmailCampaignModel> CreateCampaignAsync(
            string? subject,
            string? bodyHtml,
            string? fromEmail,
            string? createdBy,
            long? createdById,
            IEnumerable<(string Email, long? ContactId)> recipients,
            string status = "Queued")
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var cleaned = new List<(string Email, long? ContactId)>();
            foreach (var r in recipients ?? Enumerable.Empty<(string, long?)>())
            {
                var email = (r.Email ?? "").Trim();
                // Keep invalid addresses too (missing @, bad/absent domain) — they must show up as a
                // BOUNCE on the dashboard, not be silently discarded. Only truly empty entries are skipped.
                if (email.Length == 0) continue;
                if (seen.Add(email)) cleaned.Add((email, r.ContactId));
            }

            // Honour opt-outs: never build a campaign for anyone who has unsubscribed.
            if (cleaned.Count > 0)
            {
                var emails = cleaned.Select(c => c.Email).ToList();
                var suppressed = await _db.EmailRecipients
                    .Where(x => x.Unsubscribed && emails.Contains(x.Email))
                    .Select(x => x.Email)
                    .Distinct()
                    .ToListAsync();
                if (suppressed.Count > 0)
                {
                    var supSet = new HashSet<string>(suppressed, StringComparer.OrdinalIgnoreCase);
                    cleaned = cleaned.Where(c => !supSet.Contains(c.Email)).ToList();
                }
            }

            var campaign = new EmailCampaignModel
            {
                Subject = subject,
                BodyHtml = bodyHtml,
                FromEmail = fromEmail,
                CreatedBy = createdBy,
                CreatedById = createdById,
                CreatedAt = DateTime.UtcNow,
                Status = status,
                TotalRecipients = cleaned.Count
            };
            _db.EmailCampaigns.Add(campaign);
            await _db.SaveChangesAsync(); // populates campaign.Id

            var now = DateTime.UtcNow;
            var rows = cleaned.Select(c => new EmailRecipientModel
            {
                CampaignId = campaign.Id,
                Email = c.Email,
                ContactId = c.ContactId,
                TrackingToken = Guid.NewGuid().ToString("N"),
                // Invalid-format address → mark it Bounced right away (@, domain and TLD required).
                Status = IsValidEmailFormat(c.Email) ? "Pending" : "Bounced",
                Error = IsValidEmailFormat(c.Email) ? null : "Invalid email address"
            }).ToList();

            _db.EmailRecipients.AddRange(rows);
            await _db.SaveChangesAsync();

            // Tally the format-bounces and log a Bounce event for each so the dashboard reflects them.
            var badRows = rows.Where(r => r.Status == "Bounced").ToList();
            if (badRows.Count > 0)
            {
                campaign.BouncedCount += badRows.Count;
                foreach (var r in badRows)
                    _db.EmailEvents.Add(new EmailEventModel
                    {
                        RecipientId = r.Id,
                        CampaignId = campaign.Id,
                        Type = "Bounce",
                        OccurredAt = now
                    });
                await _db.SaveChangesAsync();
            }

            return campaign;
        }

        /// <summary>
        /// Records the per-recipient send outcome reported by a browser-driven send: flips each
        /// recipient to Sent/Failed, logs a Sent/Failed event, bumps the campaign tallies, and marks
        /// the campaign Completed once every recipient has a terminal status. Idempotent per recipient
        /// (a recipient already Sent/Failed is skipped) so retried reports don't double-count.
        /// </summary>
        public async Task RecordSendResultsAsync(long campaignId, IEnumerable<(long RecipientId, string Status, string? Error)> results)
        {
            var campaign = await _db.EmailCampaigns.FindAsync(campaignId);
            if (campaign == null) return;

            var byId = results
                .GroupBy(r => r.RecipientId)
                .ToDictionary(g => g.Key, g => g.Last());

            var recips = await _db.EmailRecipients
                .Where(r => r.CampaignId == campaignId && byId.Keys.Contains(r.Id))
                .ToListAsync();

            var now = DateTime.UtcNow;
            foreach (var r in recips)
            {
                if (r.Status == "Sent" || r.Status == "Failed" || r.Status == "Bounced") continue; // already terminal
                var res = byId[r.Id];
                bool sent = string.Equals(res.Status, "Sent", StringComparison.OrdinalIgnoreCase);
                // A bounce = an invalid / undeliverable address. Tracked separately from a generic Failed.
                bool bounced = string.Equals(res.Status, "Bounced", StringComparison.OrdinalIgnoreCase);
                r.Status = sent ? "Sent" : (bounced ? "Bounced" : "Failed");
                r.SentAt = sent ? now : r.SentAt;
                if (!sent) r.Error = res.Error;
                if (sent) campaign.SentCount++;
                else if (bounced) campaign.BouncedCount++;
                else campaign.FailedCount++;

                _db.EmailEvents.Add(new EmailEventModel
                {
                    RecipientId = r.Id,
                    CampaignId = campaignId,
                    Type = sent ? "Sent" : (bounced ? "Bounce" : "Failed"),
                    OccurredAt = now
                });
            }

            // Completed once every recipient has a terminal outcome.
            if (campaign.SentCount + campaign.FailedCount + campaign.BouncedCount >= campaign.TotalRecipients)
            {
                campaign.Status = "Completed";
                campaign.CompletedAt = now;
            }

            await _db.SaveChangesAsync();
        }

        // ─── HTML INSTRUMENTATION ────────────────────────────────────────────────

        private static readonly Regex HrefRegex =
            new("href\\s*=\\s*\"(?<url>[^\"]+)\"", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// Returns the recipient's HTML with every http(s) link rewritten to go through the click
        /// tracker, a hidden 1×1 open pixel appended, and an unsubscribe footer added. mailto:, tel:,
        /// and in-page anchors are left alone. <paramref name="baseUrl"/> is the public backend origin
        /// (e.g. https://crm.elpisitsolutions.com) — the recipient's mail client must be able to reach it.
        /// When <paramref name="footerOnly"/> is true, ONLY the Subscribe/Unsubscribe footer is added:
        /// the click-tracking link rewrite and the open pixel are skipped, so the normal Send button can
        /// carry the (un)subscribe option without turning every email into a fully tracked one.
        /// </summary>
        public string BuildTrackedHtml(string? bodyHtml, string token, string baseUrl, string? subBaseUrl = null, bool footerOnly = false)
        {
            var b = (baseUrl ?? "").TrimEnd('/');
            // The (un)subscribe links go through the React app (/email/...), which is always reachable
            // and records via the working /api path — sidestepping any /api/track routing issue.
            var sub = string.IsNullOrWhiteSpace(subBaseUrl) ? b : subBaseUrl.TrimEnd('/');
            var html = bodyHtml ?? "";

            // 1) Rewrite links → click tracker (keeps the real destination as ?u=). Skipped for a
            //    footer-only send, which leaves the recipient's links exactly as written.
            if (!footerOnly)
            {
                html = HrefRegex.Replace(html, m =>
                {
                    var url = m.Groups["url"].Value;
                    if (url.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase) ||
                        url.StartsWith("tel:", StringComparison.OrdinalIgnoreCase) ||
                        url.StartsWith("#") ||
                        url.StartsWith(b, StringComparison.OrdinalIgnoreCase))
                    {
                        return m.Value; // leave internal / non-web links untouched
                    }
                    var tracked = $"{b}/api/track/c/{token}?u={Uri.EscapeDataString(url)}";
                    return $"href=\"{tracked}\"";
                });
            }

            // 2) Subscription footer — clear Unsubscribe / Subscribe BUTTONS (email-safe styled links).
            //    Always added, so every recipient can opt out (or back in) from any send.
            html += $"<div style=\"margin-top:24px;padding-top:16px;border-top:1px solid #eee;" +
                    $"text-align:center;font-family:Arial,Helvetica,sans-serif;\">" +
                    $"<p style=\"font-size:12px;color:#9ca3af;margin:0 0 12px;\">Don't want to receive these emails?</p>" +
                    $"<a href=\"{sub}/email/unsubscribe?token={token}\" style=\"display:inline-block;background:#4b5563;" +
                    $"color:#ffffff;text-decoration:none;font-size:13px;font-weight:600;padding:10px 20px;" +
                    $"border-radius:8px;\">Unsubscribe</a>" +
                    $"&nbsp;&nbsp;" +
                    $"<a href=\"{sub}/email/subscribe?token={token}\" style=\"display:inline-block;background:#eff6ff;" +
                    $"color:#2563eb;text-decoration:none;font-size:13px;font-weight:600;padding:10px 20px;" +
                    $"border-radius:8px;border:1px solid #bfdbfe;\">Subscribe</a></div>";

            // 3) Open pixel (last, so it loads after the body). Skipped for a footer-only send.
            if (!footerOnly)
            {
                html += $"<img src=\"{b}/api/track/o/{token}\" width=\"1\" height=\"1\" " +
                        $"style=\"display:none;width:1px;height:1px;\" alt=\"\" />";
            }

            return html;
        }

        // ─── EVENT RECORDING (hit by the recipient's mail client, no auth) ─────────

        public async Task<EmailRecipientModel?> GetByTokenAsync(string token) =>
            await _db.EmailRecipients.FirstOrDefaultAsync(r => r.TrackingToken == token);

        public async Task RecordOpenAsync(string token, string? ip, string? userAgent)
        {
            var r = await _db.EmailRecipients.FirstOrDefaultAsync(x => x.TrackingToken == token);
            if (r == null) return;

            var now = DateTime.UtcNow;
            bool firstOpen = r.FirstOpenedAt == null;
            r.OpenCount++;
            r.LastOpenedAt = now;
            if (firstOpen) r.FirstOpenedAt = now;

            _db.EmailEvents.Add(new EmailEventModel
            {
                RecipientId = r.Id,
                CampaignId = r.CampaignId,
                Type = "Open",
                OccurredAt = now,
                IpAddress = ip,
                UserAgent = Trim(userAgent, 512)
            });

            if (firstOpen)
            {
                var c = await _db.EmailCampaigns.FindAsync(r.CampaignId);
                if (c != null) c.OpenedCount++;
            }

            await _db.SaveChangesAsync();
        }

        public async Task RecordClickAsync(string token, string url, string? ip, string? userAgent)
        {
            var r = await _db.EmailRecipients.FirstOrDefaultAsync(x => x.TrackingToken == token);
            if (r == null) return;

            var now = DateTime.UtcNow;
            bool firstClick = r.FirstClickedAt == null;
            r.ClickCount++;
            r.LastClickedAt = now;
            if (firstClick) r.FirstClickedAt = now;

            // A click implies an open even if the pixel was blocked.
            if (r.FirstOpenedAt == null)
            {
                r.FirstOpenedAt = now;
                r.LastOpenedAt = now;
                r.OpenCount++;
            }

            _db.EmailEvents.Add(new EmailEventModel
            {
                RecipientId = r.Id,
                CampaignId = r.CampaignId,
                Type = "Click",
                Url = Trim(url, 2000),
                OccurredAt = now,
                IpAddress = ip,
                UserAgent = Trim(userAgent, 512)
            });

            if (firstClick)
            {
                var c = await _db.EmailCampaigns.FindAsync(r.CampaignId);
                if (c != null) c.ClickedCount++;
            }

            await _db.SaveChangesAsync();
        }

        public async Task RecordUnsubscribeAsync(string token, string? ip, string? userAgent)
        {
            var r = await _db.EmailRecipients.FirstOrDefaultAsync(x => x.TrackingToken == token);
            if (r == null || r.Unsubscribed) return;

            var now = DateTime.UtcNow;

            // Mark this address unsubscribed on every campaign it appears in, so future sends skip it.
            var rows = await _db.EmailRecipients
                .Where(x => x.Email == r.Email && !x.Unsubscribed)
                .ToListAsync();
            foreach (var x in rows) { x.Unsubscribed = true; x.UnsubscribedAt = now; }

            _db.EmailEvents.Add(new EmailEventModel
            {
                RecipientId = r.Id,
                CampaignId = r.CampaignId,
                Type = "Unsubscribe",
                OccurredAt = now,
                IpAddress = ip,
                UserAgent = Trim(userAgent, 512)
            });

            var c = await _db.EmailCampaigns.FindAsync(r.CampaignId);
            if (c != null) c.UnsubscribedCount++;

            // Reflect the opt-out on the CRM contact record.
            if (r.ContactId != null)
            {
                var ct = await _db.Contacts.FindAsync(r.ContactId.Value);
                if (ct != null) ct.SubscriptionStatus = "Unsubscribed";
            }

            await _db.SaveChangesAsync();
        }

        /// <summary>Re-subscribes an address (clears the opt-out everywhere + on the CRM contact) so it
        /// resumes receiving campaigns. Hit from the "Subscribe" link in the email footer.</summary>
        public async Task RecordSubscribeAsync(string token, string? ip, string? userAgent)
        {
            var r = await _db.EmailRecipients.FirstOrDefaultAsync(x => x.TrackingToken == token);
            if (r == null) return;

            var now = DateTime.UtcNow;
            var rows = await _db.EmailRecipients
                .Where(x => x.Email == r.Email && x.Unsubscribed)
                .ToListAsync();
            foreach (var x in rows) { x.Unsubscribed = false; x.UnsubscribedAt = null; }

            _db.EmailEvents.Add(new EmailEventModel
            {
                RecipientId = r.Id,
                CampaignId = r.CampaignId,
                Type = "Subscribe",
                OccurredAt = now,
                IpAddress = ip,
                UserAgent = Trim(userAgent, 512)
            });

            if (r.ContactId != null)
            {
                var ct = await _db.Contacts.FindAsync(r.ContactId.Value);
                if (ct != null) ct.SubscriptionStatus = "Subscribed";
            }

            await _db.SaveChangesAsync();
        }

        /// <summary>
        /// Marks the given recipient addresses of a campaign as having replied (once each), logs a
        /// Reply event, and bumps the campaign's reply tally. Called after the client detects replies
        /// in the mailbox via Graph and reports the matching addresses.
        /// </summary>
        public async Task RecordRepliesAsync(long campaignId, IEnumerable<string> repliedEmails)
        {
            var set = new HashSet<string>(
                (repliedEmails ?? Enumerable.Empty<string>())
                    .Where(e => !string.IsNullOrWhiteSpace(e)).Select(e => e.Trim()),
                StringComparer.OrdinalIgnoreCase);
            if (set.Count == 0) return;

            var campaign = await _db.EmailCampaigns.FindAsync(campaignId);
            if (campaign == null) return;

            var recips = await _db.EmailRecipients
                .Where(r => r.CampaignId == campaignId && !r.Replied)
                .ToListAsync();

            var now = DateTime.UtcNow;
            foreach (var r in recips)
            {
                if (!set.Contains(r.Email)) continue;
                r.Replied = true;
                r.RepliedAt = now;
                campaign.RepliedCount++;
                _db.EmailEvents.Add(new EmailEventModel
                {
                    RecipientId = r.Id,
                    CampaignId = campaignId,
                    Type = "Reply",
                    OccurredAt = now
                });
            }

            await _db.SaveChangesAsync();
        }

        /// <summary>
        /// Records inbox-derived signals the client reports after scanning the sender's mailbox via
        /// Graph: read receipts (→ open, works even when the pixel is blocked), delivery receipts
        /// (→ delivered), and genuine replies. Each is applied at most once per recipient.
        /// </summary>
        public async Task RecordReceiptsAsync(
            long campaignId,
            IEnumerable<string>? opens,
            IEnumerable<string>? delivered,
            IEnumerable<string>? replies,
            IEnumerable<string>? bounces = null)
        {
            HashSet<string> ToSet(IEnumerable<string>? e) => new(
                (e ?? Enumerable.Empty<string>()).Where(x => !string.IsNullOrWhiteSpace(x)).Select(x => x.Trim()),
                StringComparer.OrdinalIgnoreCase);

            var openSet = ToSet(opens);
            var delivSet = ToSet(delivered);
            var replySet = ToSet(replies);
            var bounceSet = ToSet(bounces);
            if (openSet.Count == 0 && delivSet.Count == 0 && replySet.Count == 0 && bounceSet.Count == 0) return;

            var campaign = await _db.EmailCampaigns.FindAsync(campaignId);
            if (campaign == null) return;

            var recips = await _db.EmailRecipients
                .Where(r => r.CampaignId == campaignId)
                .ToListAsync();

            var now = DateTime.UtcNow;
            foreach (var r in recips)
            {
                // Read receipt → first open (independent of the tracking pixel).
                if (openSet.Contains(r.Email) && r.FirstOpenedAt == null)
                {
                    r.FirstOpenedAt = now;
                    r.LastOpenedAt = now;
                    r.OpenCount++;
                    campaign.OpenedCount++;
                    _db.EmailEvents.Add(new EmailEventModel
                    {
                        RecipientId = r.Id,
                        CampaignId = campaignId,
                        Type = "Open",
                        OccurredAt = now,
                        UserAgent = "Read receipt"
                    });
                }

                // Delivery receipt → delivered (once).
                if (delivSet.Contains(r.Email) && !r.Delivered)
                {
                    r.Delivered = true;
                    r.DeliveredAt = now;
                    campaign.DeliveredCount++;
                    _db.EmailEvents.Add(new EmailEventModel
                    {
                        RecipientId = r.Id,
                        CampaignId = campaignId,
                        Type = "Delivered",
                        OccurredAt = now
                    });
                }

                // Genuine reply (once).
                if (replySet.Contains(r.Email) && !r.Replied)
                {
                    r.Replied = true;
                    r.RepliedAt = now;
                    campaign.RepliedCount++;
                    _db.EmailEvents.Add(new EmailEventModel
                    {
                        RecipientId = r.Id,
                        CampaignId = campaignId,
                        Type = "Reply",
                        OccurredAt = now
                    });
                }

                // Bounce-back (NDR / "Undeliverable") → the address was undeliverable (once). This
                // catches valid-looking addresses that don't actually exist. If it had been counted
                // as Sent, undo that — it never really got delivered.
                if (bounceSet.Contains(r.Email) && r.Status != "Bounced")
                {
                    if (string.Equals(r.Status, "Sent", StringComparison.OrdinalIgnoreCase) && campaign.SentCount > 0)
                        campaign.SentCount--;
                    r.Status = "Bounced";
                    if (string.IsNullOrEmpty(r.Error)) r.Error = "Undeliverable (bounced back)";
                    campaign.BouncedCount++;
                    _db.EmailEvents.Add(new EmailEventModel
                    {
                        RecipientId = r.Id,
                        CampaignId = campaignId,
                        Type = "Bounce",
                        OccurredAt = now
                    });
                }
            }

            await _db.SaveChangesAsync();
        }

        /// <summary>
        /// Turns a raw User-Agent into a friendly device + client, and flags image-proxy pre-fetches.
        /// Gmail/Yahoo proxy-preload every pixel, so those "opens" are automated, not human reads —
        /// surfacing that is more honest than Freshsales, which just counts them as opens.
        /// </summary>
        public static (string Device, string Client, bool Proxy) DescribeUa(string? ua)
        {
            if (string.IsNullOrWhiteSpace(ua)) return ("Unknown", "Unknown", false);
            var u = ua.ToLowerInvariant();
            bool proxy = u.Contains("googleimageproxy") || u.Contains("ggpht")
                         || u.Contains("yahoomailproxy") || u.Contains("mailproxy");
            string client = proxy ? "Gmail image proxy"
                : u.Contains("outlook") ? "Outlook"
                : (u.Contains("apple mail") || u.Contains("applemail")) ? "Apple Mail"
                : u.Contains("thunderbird") ? "Thunderbird"
                : u.Contains("edg") ? "Edge"
                : u.Contains("chrome") ? "Chrome"
                : u.Contains("firefox") ? "Firefox"
                : u.Contains("safari") ? "Safari"
                : "Email client";
            string device = u.Contains("iphone") ? "iPhone"
                : u.Contains("ipad") ? "iPad"
                : u.Contains("android") ? "Android"
                : u.Contains("windows") ? "Windows"
                : (u.Contains("macintosh") || u.Contains("mac os")) ? "Mac"
                : u.Contains("linux") ? "Linux"
                : proxy ? "Proxy server"
                : "Desktop";
            return (device, client, proxy);
        }

        private static string? Trim(string? s, int max) =>
            string.IsNullOrEmpty(s) ? s : (s.Length <= max ? s : s.Substring(0, max));
    }
}
