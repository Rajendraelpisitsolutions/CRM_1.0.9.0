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
                if (email.Length == 0 || !email.Contains('@')) continue;
                if (seen.Add(email)) cleaned.Add((email, r.ContactId));
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

            var rows = cleaned.Select(c => new EmailRecipientModel
            {
                CampaignId = campaign.Id,
                Email = c.Email,
                ContactId = c.ContactId,
                TrackingToken = Guid.NewGuid().ToString("N"),
                Status = "Pending"
            }).ToList();

            _db.EmailRecipients.AddRange(rows);
            await _db.SaveChangesAsync();

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
                if (r.Status == "Sent" || r.Status == "Failed") continue; // already terminal
                var res = byId[r.Id];
                bool sent = string.Equals(res.Status, "Sent", StringComparison.OrdinalIgnoreCase);
                r.Status = sent ? "Sent" : "Failed";
                r.SentAt = sent ? now : r.SentAt;
                if (!sent) r.Error = res.Error;
                if (sent) campaign.SentCount++; else campaign.FailedCount++;

                _db.EmailEvents.Add(new EmailEventModel
                {
                    RecipientId = r.Id,
                    CampaignId = campaignId,
                    Type = sent ? "Sent" : "Failed",
                    OccurredAt = now
                });
            }

            // Completed once every recipient has a terminal outcome.
            if (campaign.SentCount + campaign.FailedCount >= campaign.TotalRecipients)
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
        /// </summary>
        public string BuildTrackedHtml(string? bodyHtml, string token, string baseUrl)
        {
            var b = (baseUrl ?? "").TrimEnd('/');
            var html = bodyHtml ?? "";

            // 1) Rewrite links → click tracker (keeps the real destination as ?u=).
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

            // 2) Unsubscribe footer.
            html += $"<div style=\"margin-top:18px;padding-top:10px;border-top:1px solid #eee;" +
                    $"font-size:12px;color:#9ca3af;font-family:Arial,Helvetica,sans-serif;\">" +
                    $"If you no longer wish to receive these emails, you can " +
                    $"<a href=\"{b}/api/track/u/{token}\" style=\"color:#6b7280;\">unsubscribe</a>.</div>";

            // 3) Open pixel (last, so it loads after the body).
            html += $"<img src=\"{b}/api/track/o/{token}\" width=\"1\" height=\"1\" " +
                    $"style=\"display:none;width:1px;height:1px;\" alt=\"\" />";

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
            r.Unsubscribed = true;
            r.UnsubscribedAt = now;

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
