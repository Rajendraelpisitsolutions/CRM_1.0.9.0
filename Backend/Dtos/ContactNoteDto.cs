using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Elpis_CRM.Model.DTOs
{
    [NotMapped]
    public class ContactNoteDto
    {
        public int Id { get; set; }
        public string? Description { get; set; }
        public long? CreatedById { get; set; }
        public DateTime? CreatedAt { get; set; }
        public DateTime? UpdatedAt { get; set; }
        [Required]
        public long ContactId { get; set; }

        // Deals this note is shared with (from NoteTargets). Empty for legacy notes.
        public List<NoteShareTargetDto> SharedWithDeals { get; set; } = new();
    }
}
