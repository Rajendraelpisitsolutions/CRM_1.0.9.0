using System;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using Elpis_CRM.Data;
using Elpis_CRM.Model.DTOs;
using Elpis_CRM.Services;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace Elpis_CRM.Controllers
{
    /// <summary>
    /// Authenticated endpoints that power the Email Tracking page: queueing a tracked bulk send and
    /// reading the analytics (overview totals, per-campaign summaries, and per-recipient drill-down).
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
    public class EmailCampaignController : ControllerBase
    {
        private readonly AppDbContext _db;
        private readonly EmailTrackingService _tracking;
        private readonly IConfiguration _config;

        public EmailCampaignController(AppDbContext db, EmailTrackingService tracking, IConfiguration config)
        {
            _db = db;
            _tracking = tracking;
            _config = config;
        }

        private string BaseUrl() =>
            (_config["Tracking:PublicBaseUrl"] ?? $"{Request.Scheme}://{Request.Host}").TrimEnd('/');

        /// <summary>The signed-in user's email — every read/write is scoped to campaigns they created,
        /// so one user never sees another user's tracking.</summary>
        private string? CurrentUserEmail() =>
            User?.FindFirst(ClaimTypes.Email)?.Value ?? User?.Identity?.Name;

        private Task<bool> OwnsCampaignAsync(long id)
        {
            var me = CurrentUserEmail();
            return _db.EmailCampaigns.AnyAsync(c => c.Id == id && c.CreatedBy == me);
        }

        /// <summary>Queues a tracked bulk email. The background sender delivers it recipient-by-recipient.</summary>
        [HttpPost]
        public async Task<IActionResult> Create([FromBody] EmailCampaignCreateDto dto)
        {
            if (dto == null || dto.Recipients == null || dto.Recipients.Count == 0)
                return BadRequest("At least one recipient is required.");
            if (string.IsNullOrWhiteSpace(dto.Subject))
                return BadRequest("Subject is required.");

            var email = User?.FindFirst(ClaimTypes.Email)?.Value ?? User?.Identity?.Name;
            var from = string.IsNullOrWhiteSpace(dto.FromEmail) ? email : dto.FromEmail;

            var campaign = await _tracking.CreateCampaignAsync(
                dto.Subject,
                dto.Body,
                from,
                email,
                null,
                dto.Recipients.Select(r => (r.Email, r.ContactId)));

            return Ok(new { campaign.Id, campaign.Status, campaign.TotalRecipients });
        }

        /// <summary>
        /// Prepares a tracked campaign for a BROWSER-driven send: creates the campaign (status
        /// "Sending" so the background sender ignores it) and returns each recipient's id and the
        /// HTML with their unique open pixel / tracked links woven in. The caller then sends each
        /// item via Graph and reports back to <c>status</c>. This lets the normal Send button and
        /// its bulk/tag flows be tracked without depending on the app-only background sender.
        /// </summary>
        [HttpPost("prepare")]
        public async Task<IActionResult> Prepare([FromBody] EmailCampaignCreateDto dto)
        {
            if (dto == null || dto.Recipients == null || dto.Recipients.Count == 0)
                return BadRequest("At least one recipient is required.");

            var email = User?.FindFirst(ClaimTypes.Email)?.Value ?? User?.Identity?.Name;
            var from = string.IsNullOrWhiteSpace(dto.FromEmail) ? email : dto.FromEmail;

            var campaign = await _tracking.CreateCampaignAsync(
                dto.Subject, dto.Body, from, email, null,
                dto.Recipients.Select(r => (r.Email, r.ContactId)),
                status: "Sending");

            // Prefer the origin the frontend actually reaches the API at, so tracking links resolve
            // to the backend in production regardless of how PublicBaseUrl is configured.
            var baseUrl = !string.IsNullOrWhiteSpace(dto.PublicBaseUrl)
                ? dto.PublicBaseUrl.TrimEnd('/')
                : BaseUrl();
            // The (un)subscribe links target the SPA; default to the API origin if not provided.
            var subBaseUrl = !string.IsNullOrWhiteSpace(dto.SubscribeBaseUrl)
                ? dto.SubscribeBaseUrl.TrimEnd('/')
                : baseUrl;
            var recipients = await _db.EmailRecipients.AsNoTracking()
                .Where(r => r.CampaignId == campaign.Id)
                .ToListAsync();

            var items = recipients.Select(r => new
            {
                recipientId = r.Id,
                email = r.Email,
                html = _tracking.BuildTrackedHtml(dto.Body, r.TrackingToken, baseUrl, subBaseUrl)
            });

            return Ok(new { campaignId = campaign.Id, items });
        }

        /// <summary>Records the per-recipient send outcomes reported after a browser-driven tracked send.</summary>
        [HttpPost("{id:long}/status")]
        public async Task<IActionResult> Status(long id, [FromBody] SendStatusDto dto)
        {
            if (dto?.Results == null || dto.Results.Count == 0) return Ok();
            if (!await OwnsCampaignAsync(id)) return NotFound();
            await _tracking.RecordSendResultsAsync(
                id, dto.Results.Select(r => (r.RecipientId, r.Status, r.Error)));
            return Ok();
        }

        /// <summary>Overall totals across every campaign (the top cards on the analytics page).</summary>
        [HttpGet("overview")]
        public async Task<IActionResult> Overview()
        {
            var me = CurrentUserEmail();
            var c = await _db.EmailCampaigns.AsNoTracking().Where(x => x.CreatedBy == me).ToListAsync();
            int sent = c.Sum(x => x.SentCount);
            int opened = c.Sum(x => x.OpenedCount);
            int clicked = c.Sum(x => x.ClickedCount);
            int unsub = c.Sum(x => x.UnsubscribedCount);
            int failed = c.Sum(x => x.FailedCount);
            int recipients = c.Sum(x => x.TotalRecipients);
            int replied = c.Sum(x => x.RepliedCount);
            int delivered = c.Sum(x => x.DeliveredCount);

            return Ok(new
            {
                campaigns = c.Count,
                recipients,
                sent,
                failed,
                opened,
                clicked,
                unsubscribed = unsub,
                replied,
                delivered,
                openRate = Rate(opened, sent),
                clickRate = Rate(clicked, sent),
                replyRate = Rate(replied, sent),
                deliveryRate = Rate(delivered, sent),
                unsubscribeRate = Rate(unsub, sent)
            });
        }

        /// <summary>Every campaign, newest first, each with its counts and computed rates.</summary>
        [HttpGet]
        public async Task<IActionResult> List()
        {
            var me = CurrentUserEmail();
            var campaigns = await _db.EmailCampaigns.AsNoTracking()
                .Where(x => x.CreatedBy == me)
                .OrderByDescending(x => x.CreatedAt)
                .ToListAsync();

            var result = campaigns.Select(x => new
            {
                x.Id,
                x.Subject,
                x.FromEmail,
                x.CreatedBy,
                x.CreatedAt,
                x.CompletedAt,
                x.Status,
                x.TotalRecipients,
                x.SentCount,
                x.FailedCount,
                x.OpenedCount,
                x.ClickedCount,
                x.UnsubscribedCount,
                x.RepliedCount,
                x.DeliveredCount,
                openRate = Rate(x.OpenedCount, x.SentCount),
                clickRate = Rate(x.ClickedCount, x.SentCount),
                replyRate = Rate(x.RepliedCount, x.SentCount)
            });

            return Ok(result);
        }

        /// <summary>A single campaign's summary.</summary>
        [HttpGet("{id:long}")]
        public async Task<IActionResult> Get(long id)
        {
            var me = CurrentUserEmail();
            var x = await _db.EmailCampaigns.AsNoTracking().FirstOrDefaultAsync(c => c.Id == id && c.CreatedBy == me);
            if (x == null) return NotFound();

            return Ok(new
            {
                x.Id,
                x.Subject,
                x.BodyHtml,
                x.FromEmail,
                x.CreatedBy,
                x.CreatedAt,
                x.CompletedAt,
                x.Status,
                x.TotalRecipients,
                x.SentCount,
                x.FailedCount,
                x.OpenedCount,
                x.ClickedCount,
                x.UnsubscribedCount,
                x.RepliedCount,
                x.DeliveredCount,
                openRate = Rate(x.OpenedCount, x.SentCount),
                clickRate = Rate(x.ClickedCount, x.SentCount),
                replyRate = Rate(x.RepliedCount, x.SentCount),
                unsubscribeRate = Rate(x.UnsubscribedCount, x.SentCount)
            });
        }

        /// <summary>The per-recipient drill-down for a campaign (who was sent, opened, clicked, unsubscribed).</summary>
        [HttpGet("{id:long}/recipients")]
        public async Task<IActionResult> Recipients(long id, [FromQuery] string? filter = null,
            [FromQuery] int page = 1, [FromQuery] int pageSize = 100)
        {
            page = Math.Max(1, page);
            pageSize = Math.Clamp(pageSize, 1, 500);

            if (!await OwnsCampaignAsync(id))
                return Ok(new { total = 0, page, pageSize, items = new List<object>() });

            var q = _db.EmailRecipients.AsNoTracking().Where(r => r.CampaignId == id);
            q = (filter ?? "").ToLowerInvariant() switch
            {
                "opened" => q.Where(r => r.FirstOpenedAt != null),
                "clicked" => q.Where(r => r.FirstClickedAt != null),
                "replied" => q.Where(r => r.Replied),
                "delivered" => q.Where(r => r.Delivered),
                "unsubscribed" => q.Where(r => r.Unsubscribed),
                "unopened" => q.Where(r => r.Status == "Sent" && r.FirstOpenedAt == null),
                "failed" => q.Where(r => r.Status == "Failed"),
                _ => q
            };

            int total = await q.CountAsync();
            var items = await q
                .OrderByDescending(r => r.LastOpenedAt ?? r.SentAt)
                .Skip((page - 1) * pageSize)
                .Take(pageSize)
                .Select(r => new
                {
                    r.Id,
                    r.Email,
                    r.ContactId,
                    r.Status,
                    r.SentAt,
                    r.Error,
                    r.OpenCount,
                    r.FirstOpenedAt,
                    r.LastOpenedAt,
                    r.ClickCount,
                    r.FirstClickedAt,
                    r.LastClickedAt,
                    r.Unsubscribed,
                    r.UnsubscribedAt,
                    r.Replied,
                    r.RepliedAt,
                    r.Delivered,
                    r.DeliveredAt
                })
                .ToListAsync();

            return Ok(new { total, page, pageSize, items });
        }

        /// <summary>Full event timeline for one recipient (opens, clicks, reply, etc.), newest first,
        /// each enriched with the device/client and a flag for automated image-proxy pre-fetches.</summary>
        [HttpGet("recipient/{recipientId:long}/events")]
        public async Task<IActionResult> RecipientEvents(long recipientId)
        {
            var r = await _db.EmailRecipients.AsNoTracking().FirstOrDefaultAsync(x => x.Id == recipientId);
            if (r == null) return NotFound();
            if (!await OwnsCampaignAsync(r.CampaignId)) return NotFound();

            var events = await _db.EmailEvents.AsNoTracking()
                .Where(e => e.RecipientId == recipientId)
                .OrderByDescending(e => e.OccurredAt)
                .ToListAsync();

            return Ok(new
            {
                recipient = new
                {
                    r.Id, r.Email, r.ContactId, r.Status, r.SentAt, r.Error,
                    r.OpenCount, r.FirstOpenedAt, r.LastOpenedAt,
                    r.ClickCount, r.FirstClickedAt, r.LastClickedAt,
                    r.Unsubscribed, r.UnsubscribedAt, r.Replied, r.RepliedAt
                },
                events = events.Select(e =>
                {
                    var (device, client, proxy) = EmailTrackingService.DescribeUa(e.UserAgent);
                    return new { e.Type, e.OccurredAt, e.Url, e.IpAddress, device, client, proxy };
                })
            });
        }

        /// <summary>Records the recipient addresses that replied (detected by the client via Graph).</summary>
        [HttpPost("{id:long}/replies")]
        public async Task<IActionResult> Replies(long id, [FromBody] RepliesDto dto)
        {
            if (dto?.Emails == null || dto.Emails.Count == 0) return Ok();
            if (!await OwnsCampaignAsync(id)) return NotFound();
            await _tracking.RecordRepliesAsync(id, dto.Emails);
            return Ok();
        }

        /// <summary>
        /// Records the results of a client-side inbox scan: read-receipt opens (confirm opens even
        /// when the tracking pixel is blocked), delivery receipts, and genuine replies.
        /// </summary>
        [HttpPost("{id:long}/receipts")]
        public async Task<IActionResult> Receipts(long id, [FromBody] ReceiptsDto dto)
        {
            if (dto == null) return Ok();
            if (!await OwnsCampaignAsync(id)) return NotFound();
            await _tracking.RecordReceiptsAsync(id, dto.Opens, dto.Delivered, dto.Replies);
            return Ok();
        }

        /// <summary>
        /// Anonymous unsubscribe-by-token, called by the public /email/unsubscribe page. Routed under
        /// /api/EmailCampaign (which the frontend already reaches) so it works even where a plain
        /// /api/track/* path isn't proxied to the API. Idempotent; returns the recipient email when found.
        /// </summary>
        [AllowAnonymous]
        [HttpPost("subscription/unsubscribe")]
        public async Task<IActionResult> UnsubscribeByToken([FromBody] SubscriptionTokenDto dto)
        {
            var token = dto?.Token?.Trim();
            if (string.IsNullOrWhiteSpace(token)) return Ok(new { found = false });
            var r = await _tracking.GetByTokenAsync(token);
            if (r == null) return Ok(new { found = false });
            await _tracking.RecordUnsubscribeAsync(token,
                HttpContext.Connection.RemoteIpAddress?.ToString(),
                Request.Headers["User-Agent"].ToString());
            return Ok(new { found = true, email = r.Email });
        }

        /// <summary>Anonymous subscribe-by-token (re-subscribe), for the public /email/subscribe page.</summary>
        [AllowAnonymous]
        [HttpPost("subscription/subscribe")]
        public async Task<IActionResult> SubscribeByToken([FromBody] SubscriptionTokenDto dto)
        {
            var token = dto?.Token?.Trim();
            if (string.IsNullOrWhiteSpace(token)) return Ok(new { found = false });
            var r = await _tracking.GetByTokenAsync(token);
            if (r == null) return Ok(new { found = false });
            await _tracking.RecordSubscribeAsync(token,
                HttpContext.Connection.RemoteIpAddress?.ToString(),
                Request.Headers["User-Agent"].ToString());
            return Ok(new { found = true, email = r.Email });
        }

        private static double Rate(int part, int whole) =>
            whole <= 0 ? 0 : Math.Round(part * 100.0 / whole, 1);
    }

    public class SubscriptionTokenDto
    {
        public string? Token { get; set; }
    }

    public class RepliesDto
    {
        public List<string> Emails { get; set; } = new();
    }

    public class ReceiptsDto
    {
        public List<string> Opens { get; set; } = new();
        public List<string> Delivered { get; set; } = new();
        public List<string> Replies { get; set; } = new();
    }
}
