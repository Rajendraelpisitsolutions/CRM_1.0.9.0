using System;

namespace Elpis_CRM.Model
{
    /// <summary>
    /// Lightweight template row for list views — everything except Body.
    /// Bodies carry base64 inline images (hundreds of KB to MB each) and the
    /// database is remote, so lists ship metadata only and the body is fetched
    /// per template via GET /api/Template/{id} when actually needed.
    /// </summary>
    public class TemplateListItemDto
    {
        public int TemplateId { get; set; }
        public string? Name { get; set; }
        public string? Subject { get; set; }
        public string? TemplateType { get; set; }
        public DateTime? CreatedAt { get; set; }
        public string? CreatedBy { get; set; }
        public DateTime? UpdatedAt { get; set; }
        public bool IsActive { get; set; }
        public bool IsDefault { get; set; }
    }
}
