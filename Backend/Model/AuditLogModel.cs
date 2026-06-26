using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Elpis_CRM.Model
{
    /// <summary>
    /// An append-only record of a change made through the application: who did it, when, what
    /// entity it affected, the action (Insert/Update/Delete/Login) and the field-level changes.
    /// </summary>
    [Table("AuditLogs")]
    public class AuditLogModel
    {
        [Key]
        public long Id { get; set; }

        /// <summary>Entity affected (e.g. "Contact", "Deal", "Login").</summary>
        [StringLength(100)]
        public string? EntityName { get; set; }

        /// <summary>Primary key of the affected row (string, since keys vary). Null for login events.</summary>
        [StringLength(100)]
        public string? EntityId { get; set; }

        /// <summary>Insert | Update | Delete | Login.</summary>
        [StringLength(20)]
        public string? Action { get; set; }

        /// <summary>Email of the user who performed the action (from the JWT).</summary>
        [StringLength(150)]
        public string? ChangedBy { get; set; }

        [StringLength(150)]
        public string? ChangedByName { get; set; }

        [StringLength(50)]
        public string? ChangedByRole { get; set; }

        [Column(TypeName = "datetime2")]
        public DateTime ChangedAt { get; set; }

        /// <summary>JSON: for updates a map of field -> { old, new }; for inserts the new values.</summary>
        public string? Changes { get; set; }

        [StringLength(64)]
        public string? IpAddress { get; set; }
    }
}
