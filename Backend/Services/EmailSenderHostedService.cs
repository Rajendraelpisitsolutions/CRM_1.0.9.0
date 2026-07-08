using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Elpis_CRM.Data;
using Elpis_CRM.Model;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace Elpis_CRM.Services
{
    /// <summary>
    /// Background worker that delivers queued <see cref="EmailCampaignModel"/> sends recipient-by-recipient.
    /// Each recipient gets their own message with a unique open pixel / tracked links woven in, so the
    /// analytics page can attribute opens and clicks to individuals. Sends are paced under Microsoft
    /// Graph's ~30 messages/minute throttle and retried on HTTP 429 (honoring Retry-After), so a large
    /// blast (e.g. 9,000) drains steadily on its own without a browser staying open.
    ///
    /// Sending uses an app-only (client-credentials) Graph token, which requires the app registration to
    /// have the *application* permission Mail.Send granted with admin consent. If those credentials are
    /// missing/invalid the campaign is marked Failed with a clear error rather than crashing the loop.
    /// </summary>
    public class EmailSenderHostedService : BackgroundService
    {
        private readonly IServiceScopeFactory _scopeFactory;
        private readonly IConfiguration _config;
        private readonly IHttpClientFactory _httpFactory;
        private readonly ILogger<EmailSenderHostedService> _logger;

        private string? _token;
        private DateTimeOffset _tokenExpiry = DateTimeOffset.MinValue;

        public EmailSenderHostedService(
            IServiceScopeFactory scopeFactory,
            IConfiguration config,
            IHttpClientFactory httpFactory,
            ILogger<EmailSenderHostedService> logger)
        {
            _scopeFactory = scopeFactory;
            _config = config;
            _httpFactory = httpFactory;
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            // Small startup delay so the app finishes booting (and the tables get created) first.
            try { await Task.Delay(TimeSpan.FromSeconds(8), stoppingToken); } catch { }

            while (!stoppingToken.IsCancellationRequested)
            {
                bool workedSomething = false;
                try
                {
                    workedSomething = await ProcessNextQueuedCampaignAsync(stoppingToken);
                }
                catch (OperationCanceledException) { break; }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Email sender loop error");
                }

                // If we just finished a campaign, look again immediately; otherwise idle-poll.
                if (!workedSomething)
                {
                    try { await Task.Delay(TimeSpan.FromSeconds(10), stoppingToken); }
                    catch { break; }
                }
            }
        }

        private async Task<bool> ProcessNextQueuedCampaignAsync(CancellationToken ct)
        {
            using var scope = _scopeFactory.CreateScope();
            var db = scope.ServiceProvider.GetRequiredService<AppDbContext>();
            var tracking = scope.ServiceProvider.GetRequiredService<EmailTrackingService>();

            var campaign = await db.EmailCampaigns
                .FirstOrDefaultAsync(c => c.Status == "Queued", ct);
            if (campaign == null) return false;

            var baseUrl = (_config["Tracking:PublicBaseUrl"] ?? "").TrimEnd('/');
            int perMinute = int.TryParse(_config["Tracking:SendRatePerMinute"], out var pm) && pm > 0 ? pm : 30;
            var interval = TimeSpan.FromSeconds(60.0 / perMinute);

            campaign.Status = "Sending";
            await db.SaveChangesAsync(ct);

            if (string.IsNullOrWhiteSpace(campaign.FromEmail))
            {
                campaign.Status = "Failed";
                campaign.CompletedAt = DateTime.UtcNow;
                await db.SaveChangesAsync(ct);
                _logger.LogError("Campaign {Id} has no FromEmail; marked Failed.", campaign.Id);
                return true;
            }

            var recipients = await db.EmailRecipients
                .Where(r => r.CampaignId == campaign.Id && r.Status == "Pending")
                .OrderBy(r => r.Id)
                .ToListAsync(ct);

            _logger.LogInformation("Sending campaign {Id} to {Count} recipients at ~{Rate}/min.",
                campaign.Id, recipients.Count, perMinute);

            foreach (var r in recipients)
            {
                if (ct.IsCancellationRequested) break;
                var started = DateTimeOffset.UtcNow;

                try
                {
                    r.Status = "Sending";
                    if (r.Unsubscribed)
                    {
                        throw new InvalidOperationException("Recipient has unsubscribed.");
                    }

                    var html = tracking.BuildTrackedHtml(campaign.BodyHtml, r.TrackingToken, baseUrl);
                    await SendViaGraphAsync(campaign.FromEmail!, r.Email, campaign.Subject ?? "", html, ct);

                    r.Status = "Sent";
                    r.SentAt = DateTime.UtcNow;
                    campaign.SentCount++;
                    db.EmailEvents.Add(new EmailEventModel
                    {
                        RecipientId = r.Id,
                        CampaignId = campaign.Id,
                        Type = "Sent",
                        OccurredAt = DateTime.UtcNow
                    });
                }
                catch (Exception ex)
                {
                    r.Status = "Failed";
                    r.Error = ex.Message.Length > 1000 ? ex.Message.Substring(0, 1000) : ex.Message;
                    campaign.FailedCount++;
                    db.EmailEvents.Add(new EmailEventModel
                    {
                        RecipientId = r.Id,
                        CampaignId = campaign.Id,
                        Type = "Failed",
                        OccurredAt = DateTime.UtcNow
                    });
                    _logger.LogWarning("Send failed for {Email} in campaign {Id}: {Err}",
                        r.Email, campaign.Id, ex.Message);
                }

                await db.SaveChangesAsync(ct);

                // Pace to stay under the per-minute limit.
                var elapsed = DateTimeOffset.UtcNow - started;
                var wait = interval - elapsed;
                if (wait > TimeSpan.Zero)
                {
                    try { await Task.Delay(wait, ct); } catch { break; }
                }
            }

            campaign.Status = ct.IsCancellationRequested ? "Queued" : "Completed";
            if (!ct.IsCancellationRequested) campaign.CompletedAt = DateTime.UtcNow;
            await db.SaveChangesAsync(ct);

            _logger.LogInformation("Campaign {Id} finished: {Sent} sent, {Failed} failed.",
                campaign.Id, campaign.SentCount, campaign.FailedCount);
            return true;
        }

        // ─── Graph app-only sending ────────────────────────────────────────────────

        private async Task SendViaGraphAsync(string from, string to, string subject, string html, CancellationToken ct)
        {
            var token = await GetAppTokenAsync(ct);
            var http = _httpFactory.CreateClient();

            var payload = new
            {
                message = new
                {
                    subject,
                    body = new { contentType = "HTML", content = html },
                    toRecipients = new[] { new { emailAddress = new { address = to } } }
                },
                saveToSentItems = true
            };
            var json = JsonSerializer.Serialize(payload);
            var url = $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(from)}/sendMail";

            const int maxAttempts = 5;
            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                using var req = new HttpRequestMessage(HttpMethod.Post, url);
                req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                req.Content = new StringContent(json, Encoding.UTF8, "application/json");

                using var res = await http.SendAsync(req, ct);
                if (res.IsSuccessStatusCode || (int)res.StatusCode == 202)
                    return;

                if (res.StatusCode == HttpStatusCode.TooManyRequests && attempt < maxAttempts)
                {
                    var retry = res.Headers.RetryAfter?.Delta ?? TimeSpan.FromSeconds(30);
                    _logger.LogInformation("Graph 429; backing off {Sec}s (attempt {A}).", retry.TotalSeconds, attempt);
                    await Task.Delay(retry, ct);
                    continue;
                }

                if (res.StatusCode == HttpStatusCode.Unauthorized && attempt < maxAttempts)
                {
                    _token = null; // force refresh, then retry once
                    token = await GetAppTokenAsync(ct);
                    continue;
                }

                var err = await res.Content.ReadAsStringAsync(ct);
                throw new Exception($"Graph sendMail {(int)res.StatusCode}: {Truncate(err, 500)}");
            }
        }

        private async Task<string> GetAppTokenAsync(CancellationToken ct)
        {
            if (_token != null && DateTimeOffset.UtcNow < _tokenExpiry.AddMinutes(-2))
                return _token;

            var tenant = _config["AzureAd:TenantId"];
            var clientId = _config["AzureAd:ClientId"];
            var secret = _config["AzureAd:ClientSecret"];
            if (string.IsNullOrWhiteSpace(tenant) || string.IsNullOrWhiteSpace(clientId) || string.IsNullOrWhiteSpace(secret))
                throw new Exception("App-only Graph credentials (AzureAd:TenantId/ClientId/ClientSecret) are not configured.");

            var http = _httpFactory.CreateClient();
            var form = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                ["client_id"] = clientId!,
                ["client_secret"] = secret!,
                ["scope"] = "https://graph.microsoft.com/.default",
                ["grant_type"] = "client_credentials"
            });

            using var res = await http.PostAsync(
                $"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token", form, ct);
            var body = await res.Content.ReadAsStringAsync(ct);
            if (!res.IsSuccessStatusCode)
                throw new Exception($"App token request failed ({(int)res.StatusCode}): {Truncate(body, 500)}");

            using var doc = JsonDocument.Parse(body);
            _token = doc.RootElement.GetProperty("access_token").GetString();
            var expiresIn = doc.RootElement.TryGetProperty("expires_in", out var e) ? e.GetInt32() : 3600;
            _tokenExpiry = DateTimeOffset.UtcNow.AddSeconds(expiresIn);
            return _token!;
        }

        private static string Truncate(string s, int max) =>
            string.IsNullOrEmpty(s) ? "" : (s.Length <= max ? s : s.Substring(0, max));
    }
}
