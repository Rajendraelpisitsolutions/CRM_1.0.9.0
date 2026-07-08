using System;
using System.Threading.Tasks;
using Elpis_CRM.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace Elpis_CRM.Controllers
{
    /// <summary>
    /// Public, unauthenticated endpoints that recipients' mail clients hit: the open pixel, the
    /// click-through redirect, and the unsubscribe link. Reachable at BOTH /track/* and
    /// /api/track/* — the emails use the /api/track/* form so they ride the same reverse-proxy
    /// rule that already forwards /api to the backend in production (a plain /track/* path is
    /// typically NOT proxied to the API, which is why opens/clicks were never recorded).
    /// These must be reachable from the public internet.
    /// </summary>
    [ApiController]
    [AllowAnonymous]
    [Route("track")]
    [Route("api/track")]
    public class EmailTrackingController : ControllerBase
    {
        private readonly EmailTrackingService _tracking;

        public EmailTrackingController(EmailTrackingService tracking)
        {
            _tracking = tracking;
        }

        // 1×1 transparent GIF returned by the open pixel.
        private static readonly byte[] PixelGif =
        {
            0x47,0x49,0x46,0x38,0x39,0x61,0x01,0x00,0x01,0x00,0x80,0x00,0x00,
            0xFF,0xFF,0xFF,0x00,0x00,0x00,0x21,0xF9,0x04,0x01,0x00,0x00,0x00,0x00,
            0x2C,0x00,0x00,0x00,0x00,0x01,0x00,0x01,0x00,0x00,0x02,0x02,0x44,0x01,0x00,0x3B
        };

        /// <summary>Open pixel. Logs the open then returns a 1×1 GIF (no-cache so re-opens count).</summary>
        [HttpGet("o/{token}")]
        public async Task<IActionResult> Open(string token)
        {
            try { await _tracking.RecordOpenAsync(token, ClientIp(), UserAgent()); }
            catch { /* never fail the image request */ }

            Response.Headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0";
            Response.Headers["Pragma"] = "no-cache";
            return File(PixelGif, "image/gif");
        }

        /// <summary>Click tracker. Logs the click then 302-redirects to the real destination (?u=).</summary>
        [HttpGet("c/{token}")]
        public async Task<IActionResult> Click(string token, [FromQuery] string u)
        {
            string dest = u ?? "";
            // Only redirect to real web links; ignore anything else to avoid open-redirect abuse.
            bool ok = Uri.TryCreate(dest, UriKind.Absolute, out var uri) &&
                      (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps);

            try { if (ok) await _tracking.RecordClickAsync(token, dest, ClientIp(), UserAgent()); }
            catch { /* never block the redirect */ }

            return ok ? Redirect(dest) : NotFound();
        }

        /// <summary>Unsubscribe link. Records the opt-out and shows a confirmation page.</summary>
        [HttpGet("u/{token}")]
        public async Task<IActionResult> Unsubscribe(string token)
        {
            try { await _tracking.RecordUnsubscribeAsync(token, ClientIp(), UserAgent()); }
            catch { /* show the page regardless */ }

            return Content(ConfirmPage(
                "Unsubscribed",
                "Unsubscribed successfully",
                "You have been unsubscribed. You will not receive any further emails from us. " +
                "Changed your mind? Use the Subscribe link in any earlier email to opt back in."
            ), "text/html");
        }

        /// <summary>Subscribe link. Clears the opt-out and shows a confirmation page.</summary>
        [HttpGet("s/{token}")]
        public async Task<IActionResult> Subscribe(string token)
        {
            try { await _tracking.RecordSubscribeAsync(token, ClientIp(), UserAgent()); }
            catch { /* show the page regardless */ }

            return Content(ConfirmPage(
                "Subscribed",
                "Subscribed successfully",
                "You're subscribed. You'll keep receiving the emails we send. " +
                "You can unsubscribe at any time using the link in our emails."
            ), "text/html");
        }

        // Small self-contained confirmation page shown to recipients.
        private static string ConfirmPage(string title, string heading, string message) =>
            "<!doctype html><html><head><meta charset=\"utf-8\">" +
            "<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">" +
            $"<title>{title}</title></head>" +
            "<body style=\"font-family:Arial,Helvetica,sans-serif;background:#f8fafc;margin:0;\">" +
            "<div style=\"max-width:480px;margin:12vh auto;background:#fff;border:1px solid #e5e7eb;" +
            "border-radius:16px;padding:36px;text-align:center;\">" +
            "<div style=\"width:56px;height:56px;border-radius:50%;background:#ecfdf5;color:#059669;" +
            "font-size:30px;line-height:56px;margin:0 auto 12px;\">✓</div>" +
            $"<h2 style=\"color:#111827;margin:8px 0;\">{heading}</h2>" +
            $"<p style=\"color:#6b7280;line-height:1.6;\">{message} You can close this window.</p>" +
            "</div></body></html>";

        private string? ClientIp() => HttpContext.Connection.RemoteIpAddress?.ToString();
        private string? UserAgent() => Request.Headers["User-Agent"].ToString();
    }
}
