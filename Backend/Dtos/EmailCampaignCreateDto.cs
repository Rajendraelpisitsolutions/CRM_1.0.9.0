using System.Collections.Generic;

namespace Elpis_CRM.Model.DTOs
{
    /// <summary>Payload the frontend posts to queue a tracked bulk email.</summary>
    public class EmailCampaignCreateDto
    {
        public string? Subject { get; set; }
        /// <summary>The email body as HTML.</summary>
        public string? Body { get; set; }
        /// <summary>The mailbox to send from (defaults to the signed-in user server-side).</summary>
        public string? FromEmail { get; set; }
        /// <summary>The public origin the frontend reaches the API at (e.g. https://crm.example.com).
        /// The open pixel / click / (un)subscribe links are built on this so they resolve to the
        /// backend in production. Falls back to the Tracking:PublicBaseUrl config when absent.</summary>
        public string? PublicBaseUrl { get; set; }
        /// <summary>The app (SPA) origin, used to build the /email/subscribe and /email/unsubscribe
        /// links so recipients land on the React page (which records via the working /api path).</summary>
        public string? SubscribeBaseUrl { get; set; }
        public List<EmailRecipientInputDto> Recipients { get; set; } = new();
    }

    public class EmailRecipientInputDto
    {
        public string Email { get; set; } = string.Empty;
        public long? ContactId { get; set; }
    }

    /// <summary>Per-recipient send outcomes the browser reports back after a hybrid tracked send.</summary>
    public class SendStatusDto
    {
        public List<SendResultDto> Results { get; set; } = new();
    }

    public class SendResultDto
    {
        public long RecipientId { get; set; }
        /// <summary>"Sent" or "Failed".</summary>
        public string Status { get; set; } = "Sent";
        public string? Error { get; set; }
    }
}
