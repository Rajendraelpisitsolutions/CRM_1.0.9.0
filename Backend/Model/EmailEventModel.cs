using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Elpis_CRM.Model
{
    /// <summary>
    /// A single tracked interaction with a sent email — one row per open, click, unsubscribe,
    /// send or failure. Powers the timeline drill-down on the analytics page.
    /// </summary>
    [Table("EmailEvents")]
    public class EmailEventModel
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public long Id { get; set; }

        public long RecipientId { get; set; }

        /// <summary>Denormalized so campaign-level rollups don't need a join.</summary>
        public long CampaignId { get; set; }

        /// <summary>Sent | Failed | Open | Click | Unsubscribe.</summary>
        public string Type { get; set; } = string.Empty;

        /// <summary>The destination url for Click events; null otherwise.</summary>
        public string? Url { get; set; }

        public DateTime OccurredAt { get; set; }

        public string? IpAddress { get; set; }
        public string? UserAgent { get; set; }
    }
}
