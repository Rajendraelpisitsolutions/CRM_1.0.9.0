using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Elpis_CRM.Model
{
    [Table("Notes")]
    public class NotesModel
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        public string? Description { get; set; }
        public long? CreatedById { get; set; }
        public DateTime? CreatedAt { get; set; }
        public DateTime? UpdatedAt { get; set; }
        public long? ContactId { get; set; }
        public string? RelatedToType { get; set; }
        public long? DealId { get; set; }

        // True for notes created under the shared model (Option A): the note is a single
        // row and its associations live in NoteTargets. Legacy notes stay false and keep
        // their old per-copy behaviour, so existing data is read exactly as before.
        public bool IsShared { get; set; }

        [NotMapped]
        public List<long>? MirrorToContactIds { get; set; }

        [NotMapped]
        public List<long>? MirrorToDealIds { get; set; }
    }
}