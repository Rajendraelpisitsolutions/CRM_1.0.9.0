using Elpis_CRM.Data;
using Elpis_CRM.Dtos;
using Elpis_CRM.Model;
using Microsoft.EntityFrameworkCore;
using System.Text.Json;

namespace Elpis_CRM.Services
{
    /// <summary>
    /// Provides services for managing recycle bin operations such as
    /// creating entries, retrieving deleted items, restoring entities,
    /// and permanently deleting them.
    /// </summary>
    public class RecycleBinService
    {
        private readonly AppDbContext _db;

        /// <summary>
        /// Initializes a new instance of the <see cref="RecycleBinService"/> class.
        /// </summary>
        /// <param name="db">Application database context.</param>
        public RecycleBinService(AppDbContext db)
        {
            _db = db;
        }

        /// <summary>
        /// Retrieves all deleted items that have not been restored.
        /// </summary>
        /// <returns>A list of recycle bin items.</returns>
        public async Task<List<RecycleBinItemDto>> GetDeletedItemsAsync()
        {
            var items = await _db.RecycleBinItems
                .Where(x => !x.IsRestored)
                .OrderByDescending(x => x.DeletedAt)
                .ToListAsync();

            return items.Select(x => new RecycleBinItemDto
            {
                EntityType = x.EntityType,
                EntityId = x.EntityId,
                DisplayName = x.DisplayName ?? x.EntityType,
                Details = x.Details,
                DeletedAt = x.DeletedAt,
                DeletedBy = x.DeletedBy,
                RestoredAt = x.RestoredAt,
                RestoredBy = x.RestoredBy
            }).ToList();
        }

        /// <summary>
        /// Restores a deleted entity from the recycle bin using its stored payload.
        /// </summary>
        /// <param name="entityType">The type of entity to restore.</param>
        /// <param name="entityId">The unique identifier of the entity.</param>
        /// <param name="restoredBy">The user performing the restore operation.</param>
        /// <returns>
        /// <c>true</c> if the entity was restored successfully; otherwise, <c>false</c>.
        /// </returns>
        public async Task<bool> RestoreAsync(string entityType, string entityId, string restoredBy)
        {
            var entry = await _db.RecycleBinItems.FirstOrDefaultAsync(x => x.EntityType == entityType && x.EntityId == entityId && !x.IsRestored);
            if (entry == null) return false;

            try
            {
                switch (entityType)
                {
                    case "Account":
                        if (!string.IsNullOrWhiteSpace(entry.Payload))
                        {
                            var account = JsonSerializer.Deserialize<AccountModel>(entry.Payload);
                            if (account != null)
                            {
                                account.AccountId = long.Parse(entityId);
                                _db.Accounts.Add(account);
                            }
                        }
                        break;
                    case "Contact":
                        if (!string.IsNullOrWhiteSpace(entry.Payload))
                        {
                            var contact = JsonSerializer.Deserialize<ContactModel>(entry.Payload);
                            if (contact != null)
                            {
                                contact.ContactId = long.Parse(entityId);
                                _db.Contacts.Add(contact);
                            }
                        }
                        break;
                    case "Product":
                        if (!string.IsNullOrWhiteSpace(entry.Payload))
                        {
                            var product = JsonSerializer.Deserialize<ProductsModel>(entry.Payload);
                            if (product != null)
                            {
                                product.ProductId = int.Parse(entityId);
                                _db.Products.Add(product);
                            }
                        }
                        break;
                    case "Deal":
                        if (!string.IsNullOrWhiteSpace(entry.Payload))
                        {
                            var deal = JsonSerializer.Deserialize<DealModel>(entry.Payload);
                            if (deal != null)
                            {
                                deal.Id = long.Parse(entityId);
                                _db.Deals.Add(deal);
                            }
                        }
                        break;
                    case "Note":
                        if (!string.IsNullOrWhiteSpace(entry.Payload))
                        {
                            var note = JsonSerializer.Deserialize<NotesModel>(entry.Payload);

                            if (note != null)
                            {
                                note.Id = 0;
                                _db.Notes.Add(note);
                            }
                        }
                        break;
                    default:
                        return false;
                }
            }
            catch
            {
                return false;
            }

            entry.IsRestored = true;
            entry.RestoredAt = ToIndianTime(DateTime.UtcNow);
            entry.RestoredBy = restoredBy;
            await _db.SaveChangesAsync();
            return true;
        }

        /// <summary>
        /// Permanently deletes an entity and removes its recycle bin entry.
        /// </summary>
        /// <param name="entityType">The type of entity to delete.</param>
        /// <param name="entityId">The unique identifier of the entity.</param>
        /// <returns>
        /// <c>true</c> if the entity was permanently deleted; otherwise, <c>false</c>.
        /// </returns>
        public async Task<bool> PermanentDeleteAsync(string entityType, string entityId)
        {
            var entry = await _db.RecycleBinItems.FirstOrDefaultAsync(x => x.EntityType == entityType && x.EntityId == entityId && !x.IsRestored);
            if (entry == null) return false;

            switch (entityType)
            {
                case "Account":
                    var account = await _db.Accounts.FirstOrDefaultAsync(x => x.AccountId.ToString() == entityId);
                    if (account != null) _db.Accounts.Remove(account);
                    break;
                case "Contact":
                    var contact = await _db.Contacts.FirstOrDefaultAsync(x => x.ContactId.ToString() == entityId);
                    if (contact != null) _db.Contacts.Remove(contact);
                    break;
                case "Product":
                    var product = await _db.Products.FirstOrDefaultAsync(x => x.ProductId.ToString() == entityId);
                    if (product != null) _db.Products.Remove(product);
                    break;
                case "Deal":
                    var deal = await _db.Deals.FirstOrDefaultAsync(x => x.Id.ToString() == entityId);
                    if (deal != null) _db.Deals.Remove(deal);
                    break;
                case "Note":
                    var note = await _db.Notes.FirstOrDefaultAsync(x => x.Id.ToString() == entityId);
                    if (note != null) _db.Notes.Remove(note);
                    break;
                default:
                    return false;
            }

            _db.RecycleBinItems.Remove(entry);
            await _db.SaveChangesAsync();
            return true;
        }

        /// <summary>
        /// Converts a UTC date and time to Indian Standard Time (IST).
        /// Falls back to adding 5 hours and 30 minutes if the time zone cannot be resolved.
        /// </summary>
        /// <param name="utcDateTime">The UTC date and time.</param>
        /// <returns>The converted Indian Standard Time.</returns>
        private static DateTime ToIndianTime(DateTime utcDateTime)
        {
            try
            {
                var indiaTimeZone = TimeZoneInfo.FindSystemTimeZoneById("India Standard Time");
                return TimeZoneInfo.ConvertTimeFromUtc(utcDateTime, indiaTimeZone);
            }
            catch
            {
                return utcDateTime.AddHours(5.5);
            }
        }

        /// <summary>
        /// Creates a new recycle bin entry for a deleted entity.
        /// </summary>
        /// <param name="entityType">The type of the deleted entity.</param>
        /// <param name="entityId">The unique identifier of the deleted entity.</param>
        /// <param name="displayName">The display name of the entity.</param>
        /// <param name="details">Additional details about the entity.</param>
        /// <param name="deletedBy">The user who deleted the entity.</param>
        /// <param name="payload">The serialized entity data used for restoration.</param>
        /// <returns>A task representing the asynchronous operation.</returns>
        public async Task CreateEntryAsync(string entityType, string entityId, string displayName, string? details, string? deletedBy, object? payload = null)
        {
            var entry = new RecycleBinItemModel
            {
                EntityType = entityType,
                EntityId = entityId,
                DisplayName = displayName,
                Details = details,
                DeletedAt = ToIndianTime(DateTime.UtcNow),
                DeletedBy = deletedBy,
                Payload = payload == null ? null : JsonSerializer.Serialize(payload)
            };

            _db.RecycleBinItems.Add(entry);
            await _db.SaveChangesAsync();
        }
    }
}