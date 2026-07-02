namespace Elpis_CRM.Dtos
{
    public class RecycleBinItemDto
    {
        public string EntityType { get; set; } = string.Empty;
        public string EntityId { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string? Details { get; set; }
        public DateTime? DeletedAt { get; set; }
        public string? DeletedBy { get; set; }
        public DateTime? RestoredAt { get; set; }
        public string? RestoredBy { get; set; }
    }
}
