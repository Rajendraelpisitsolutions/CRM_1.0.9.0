using System;
using System.Threading.Tasks;
using Elpis_CRM.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace Elpis_CRM.Controllers
{
    /// <summary>
    /// Public, unauthenticated endpoints that recipients' mail clients hit: the open pixel, the
    /// click-through redirect, and the unsubscribe link. Routed at /track/* (not /api) so the URLs
    /// embedded in emails stay short. These must be reachable from the public internet.
    /// </summary>
    [ApiController]
    [AllowAnonymous]
    [Route("track")]
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

        /// <summary>Unsubscribe link. Records the opt-out and shows a small confirmation page.</summary>
        [HttpGet("u/{token}")]
        public async Task<IActionResult> Unsubscribe(string token)
        {
            try { await _tracking.RecordUnsubscribeAsync(token, ClientIp(), UserAgent()); }
            catch { /* show the page regardless */ }

            const string page = "<!doctype html><html><head><meta charset=\"utf-8\">" +
                "<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">" +
                "<title>Unsubscribed</title></head>" +
                "<body style=\"font-family:Arial,Helvetica,sans-serif;background:#f8fafc;margin:0;\">" +
                "<div style=\"max-width:480px;margin:12vh auto;background:#fff;border:1px solid #e5e7eb;" +
                "border-radius:16px;padding:36px;text-align:center;\">" +
                "<div style=\"font-size:40px\">✓</div>" +
                "<h2 style=\"color:#111827;margin:8px 0;\">You've been unsubscribed</h2>" +
                "<p style=\"color:#6b7280;\">You won't receive further emails from this list. " +
                "You can close this window.</p></div></body></html>";
            return Content(page, "text/html");
        }

        private string? ClientIp() => HttpContext.Connection.RemoteIpAddress?.ToString();
        private string? UserAgent() => Request.Headers["User-Agent"].ToString();
    }
}
