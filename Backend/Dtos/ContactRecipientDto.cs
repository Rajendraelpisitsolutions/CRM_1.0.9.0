namespace Elpis_CRM.Dtos
{
    /// <summary>
    /// One addressable email address together with the contact it belongs to. Exists because the
    /// bulk-email flow needs the contact's identity — not just the address — to build a per-person
    /// greeting ("Dear Mr. Ravi Kumar,"); the older tags/emails endpoint returns bare addresses
    /// and cannot support that.
    /// </summary>
    public class ContactRecipientDto
    {
        /// <summary>Primary key of the contact this address belongs to.</summary>
        public long ContactId { get; set; }

        /// <summary>Display name ("FirstName LastName"), trimmed; empty when the contact has neither.</summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>A single email address for the contact. A contact with several addresses yields several rows.</summary>
        public string Email { get; set; } = string.Empty;
    }
}
