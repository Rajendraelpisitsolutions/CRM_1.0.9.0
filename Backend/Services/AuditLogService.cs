using System;
using System.Linq;
using System.Threading.Tasks;
using Elpis_CRM.Data;
using Microsoft.EntityFrameworkCore;

namespace Elpis_CRM.Service
{
    /// <summary>
    /// Read-side access to the audit trail. Supports filtering by entity, action, user and a free-text
    /// search, returned newest-first and paged.
    /// </summary>
    public class AuditLogService
    {
        private readonly AppDbContext _db;

        public AuditLogService(AppDbContext db)
        {
            _db = db;
        }

        /// <summary>
        /// Returns a page of audit entries (newest first) plus the total matching count.
        /// </summary>
        public async Task<object> GetAsync(string? entity, string? action, string? user, string? search,
            int page = 1, int pageSize = 50)
        {
            if (page < 1) page = 1;
            pageSize = Math.Clamp(pageSize, 1, 200);

            var q = _db.AuditLogs.AsNoTracking().AsQueryable();

            if (!string.IsNullOrWhiteSpace(entity))
                q = q.Where(a => a.EntityName == entity);

            if (!string.IsNullOrWhiteSpace(action))
                q = q.Where(a => a.Action == action);

            if (!string.IsNullOrWhiteSpace(user))
                q = q.Where(a => a.ChangedBy != null && a.ChangedBy.ToLower().Contains(user.ToLower()));

            if (!string.IsNullOrWhiteSpace(search))
            {
                var s = search.ToLower();
                q = q.Where(a =>
                    (a.ChangedBy != null && a.ChangedBy.ToLower().Contains(s)) ||
                    (a.ChangedByName != null && a.ChangedByName.ToLower().Contains(s)) ||
                    (a.EntityName != null && a.EntityName.ToLower().Contains(s)) ||
                    (a.EntityId != null && a.EntityId.ToLower().Contains(s)) ||
                    (a.Action != null && a.Action.ToLower().Contains(s)) ||
                    (a.Changes != null && a.Changes.ToLower().Contains(s)));
            }

            var total = await q.CountAsync();
            var items = await q
                .OrderByDescending(a => a.ChangedAt)
                .Skip((page - 1) * pageSize)
                .Take(pageSize)
                .ToListAsync();

            // ChangedAt is stored as UTC (DateTime.UtcNow) in a datetime2 column, which EF
            // reads back with Kind=Unspecified — serialized without a 'Z' and then mis-parsed
            // as local time on the client. Re-stamp as UTC so the JSON carries the 'Z' and the
            // UI's IST (Asia/Kolkata) conversion is applied to the correct instant.
            foreach (var a in items)
                a.ChangedAt = DateTime.SpecifyKind(a.ChangedAt, DateTimeKind.Utc);

            return new { items, totalCount = total, page, pageSize };
        }

        /// <summary>Distinct entity names present in the log, for a filter dropdown.</summary>
        public async Task<System.Collections.Generic.List<string>> GetEntitiesAsync()
        {
            return await _db.AuditLogs
                .Where(a => a.EntityName != null)
                .Select(a => a.EntityName!)
                .Distinct()
                .OrderBy(x => x)
                .ToListAsync();
        }
    }
}
