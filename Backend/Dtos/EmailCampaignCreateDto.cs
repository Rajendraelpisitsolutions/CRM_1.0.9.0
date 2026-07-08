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
