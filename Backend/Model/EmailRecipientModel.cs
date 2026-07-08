using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Elpis_CRM.Model
{
    /// <summary>
    /// One recipient of a campaign. Carries the unguessable <see cref="TrackingToken"/> that the
    /// open pixel, click links and unsubscribe link embed, plus the per-person engagement tallies.
    /// </summary>
    [Table("EmailRecipients")]
    public class EmailRecipientModel
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public long Id { get; set; }

        public long CampaignId { get; set; }

        public string Email { get; set; } = string.Empty;

        /// <summary>The CRM contact this address belongs to, when known.</summary>
        public long? ContactId { get; set; }

        /// <summary>Unguessable per-recipient token embedded in the pixel / links (a GUID).</summary>
        public string TrackingToken { get; set; } = string.Empty;

        /// <summary>Pending | Sending | Sent | Failed.</summary>
        public string Status { get; set; } = "Pending";

        public DateTime? SentAt { get; set; }
        public string? Error { get; set; }

        // Engagement.
        public int OpenCount { get; set; }
        public DateTime? FirstOpenedAt { get; set; }
        public DateTime? LastOpenedAt { get; set; }

        public int ClickCount { get; set; }
        public DateTime? FirstClickedAt { get; set; }
        public DateTime? LastClickedAt { get; set; }

        public bool Unsubscribed { get; set; }
        public DateTime? UnsubscribedAt { get; set; }

        public bool Replied { get; set; }
        public DateTime? RepliedAt { get; set; }

        public EmailCampaignModel? Campaign { get; set; }
    }
}
