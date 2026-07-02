using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Elpis_CRM.Model
{
    [Table("Entities")]
    public class RecycleBinItemModel
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public long Id { get; set; }

        [Required]
        [StringLength(100)]
        public string EntityType { get; set; } = string.Empty;

        [Required]
        [StringLength(100)]
        public string EntityId { get; set; } = string.Empty;

        [StringLength(500)]
        public string? DisplayName { get; set; }

        public string? Details { get; set; }

        public DateTime DeletedAt { get; set; } = DateTime.Now;

        [StringLength(200)]
        public string? DeletedBy { get; set; }

        public DateTime? RestoredAt { get; set; }

        [StringLength(200)]
        public string? RestoredBy { get; set; }

        public bool IsRestored { get; set; } = false;

        public string? Payload { get; set; }
    }
}