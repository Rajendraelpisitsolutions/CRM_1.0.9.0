using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Elpis_CRM.Model
{
    /// <summary>
    /// One bulk-email send. Holds the shared subject/body and the running tallies that the
    /// analytics page reads. Each recipient it was sent to is an <see cref="EmailRecipientModel"/>.
    /// </summary>
    [Table("EmailCampaigns")]
    public class EmailCampaignModel
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public long Id { get; set; }

        public string? Subject { get; set; }
        public string? BodyHtml { get; set; }

        /// <summary>The mailbox the emails are sent from (the signed-in user's address).</summary>
        public string? FromEmail { get; set; }

        public string? CreatedBy { get; set; }
        public long? CreatedById { get; set; }
        public DateTime CreatedAt { get; set; }
        public DateTime? CompletedAt { get; set; }

        /// <summary>Draft | Queued | Sending | Completed | Failed.</summary>
        public string Status { get; set; } = "Queued";

        // Running tallies (updated as the send progresses and as tracking events arrive).
        public int TotalRecipients { get; set; }
        public int SentCount { get; set; }
        public int FailedCount { get; set; }
        public int OpenedCount { get; set; }
        public int ClickedCount { get; set; }
        public int UnsubscribedCount { get; set; }
        public int RepliedCount { get; set; }
        public int DeliveredCount { get; set; }
        /// <summary>Sends that bounced — an invalid / undeliverable address.</summary>
        public int BouncedCount { get; set; }
    }
}
