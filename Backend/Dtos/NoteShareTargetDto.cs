namespace Elpis_CRM.Model.DTOs
{
    /// <summary>
    /// A record a shared note is associated with, shown on the note card
    /// (e.g. the contact names under a deal note, or the deal names under a contact note).
    /// </summary>
    public class NoteShareTargetDto
    {
        public long Id { get; set; }
        public string Name { get; set; } = string.Empty;
    }
}
