using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Elpis_CRM.Model
{
    /// <summary>
    /// One association of a shared note (Option A) to a record it appears under.
    /// A deal note shared to two contacts is one Notes row plus three NoteTargets:
    /// one (Deal, dealId) origin row and one (Contact, contactId) row per contact.
    /// Deleting the note cascades and removes its targets.
    /// </summary>
    [Table("NoteTargets")]
    public class NoteTargetModel
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public long Id { get; set; }

        // FK to Notes.Id (int primary key on NotesModel).
        public int NoteId { get; set; }

        // "Deal" or "Contact".
        public string TargetType { get; set; } = string.Empty;

        // Deals.Id or Contacts.ContactId, matching TargetType.
        public long TargetId { get; set; }

        public NotesModel? Note { get; set; }
    }
}
