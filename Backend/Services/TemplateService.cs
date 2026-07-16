using Elpis_CRM.Data;
using Elpis_CRM.Model;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Elpis_CRM.Service
{
    /// <summary>
    /// Service class to manage CRUD operations and retrieval for email/template models.
    /// </summary>
    public class TemplateService
    {
        private readonly AppDbContext _templateDbContext;

        /// <summary>
        /// Initializes a new instance of the <see cref="TemplateService"/> class.
        /// </summary>
        /// <param name="templateDbContext">The database context used for template operations.</param>
        public TemplateService(AppDbContext templateDbContext)
        {
            _templateDbContext = templateDbContext;
        }

        /// <summary>
        /// Inserts a template, setting both CreatedAt and UpdatedAt to the current UTC time before saving.
        /// </summary>
        /// <param name="template">The template to persist.</param>
        /// <returns>The same instance after saving, now carrying its database-generated ID and timestamps.</returns>
        public async Task<TemplateModel> AddTemplate(TemplateModel template)
        {
            template.CreatedAt = DateTime.UtcNow;
            template.UpdatedAt = DateTime.UtcNow;

            _templateDbContext.Templates.Add(template);
            await _templateDbContext.SaveChangesAsync();
            return template;
        }

        /// <summary>
        /// Loads every template regardless of its active state.
        /// </summary>
        /// <returns>All templates as a list; empty when the table holds no rows.</returns>
        public async Task<List<TemplateModel>> GetAllAsync()
        {
            return await _templateDbContext.Templates.ToListAsync();
        }

        /// <summary>
        /// Fetches a template by primary key using the context's identity-map-aware Find.
        /// </summary>
        /// <param name="templateId">Primary key to look up.</param>
        /// <returns>The matching template, or null when no row has that ID.</returns>
        public async Task<TemplateModel?> GetByIdAsync(int templateId)
        {
            return await _templateDbContext.Templates.FindAsync(templateId);
        }

        /// <summary>
        /// Returns the first template whose Name equals the given value.
        /// </summary>
        /// <param name="name">Exact name to match.</param>
        /// <returns>The first matching template, or null when none match.</returns>
        public async Task<TemplateModel?> GetByNameAsync(string name)
        {
            return await _templateDbContext.Templates
                                           .FirstOrDefaultAsync(t => t.Name == name);
        }

        /// <summary>
        /// Copies name, subject, body, type and active flag onto the existing row and bumps UpdatedAt to now (UTC).
        /// CreatedAt and the ID are left untouched.
        /// </summary>
        /// <param name="templateId">Primary key of the template to update.</param>
        /// <param name="template">Source of the new field values.</param>
        /// <returns>The updated template, or null when no row has that ID.</returns>
        public async Task<TemplateModel?> UpdateTemplate(int templateId, TemplateModel template)
        {
            var existing = await _templateDbContext.Templates.FindAsync(templateId);
            if (existing == null)
            {
                return null;
            }

            existing.Name = template.Name;
            existing.Subject = template.Subject;
            existing.Body = template.Body;
            existing.TemplateType = template.TemplateType;
            // CreatedBy carries the email the template is assigned to — allow reassignment.
            if (!string.IsNullOrWhiteSpace(template.CreatedBy))
            {
                existing.CreatedBy = template.CreatedBy;
            }
            existing.UpdatedAt = DateTime.UtcNow;
            existing.IsActive = template.IsActive;

            await _templateDbContext.SaveChangesAsync();
            return existing;
        }

        /// <summary>
        /// The logged-in user's active templates, WITHOUT bodies. Filtering runs in SQL and
        /// Body is never selected, so the remote database only ships a few metadata rows —
        /// bodies (base64 inline images, up to MBs each) are fetched per template by ID when
        /// actually opened/applied. Ownership: CreatedBy equals the email; legacy templates
        /// with no CreatedBy are owned when the email appears in their name or body.
        /// </summary>
        /// <param name="email">The owner's email (the value stored in CreatedBy).</param>
        /// <returns>The user's active templates as lightweight list items.</returns>
        public async Task<List<TemplateListItemDto>> GetMineAsync(string email)
        {
            var owner = (email ?? string.Empty).Trim();
            return await _templateDbContext.Templates
                .Where(t => t.IsActive &&
                       (t.CreatedBy == owner ||
                        ((t.CreatedBy == null || t.CreatedBy == "") && owner != "" &&
                         ((t.Name != null && t.Name.Contains(owner)) ||
                          (t.Body != null && t.Body.Contains(owner))))))
                .Select(t => new TemplateListItemDto
                {
                    TemplateId = t.TemplateId,
                    Name = t.Name,
                    Subject = t.Subject,
                    TemplateType = t.TemplateType,
                    CreatedAt = t.CreatedAt,
                    CreatedBy = t.CreatedBy,
                    UpdatedAt = t.UpdatedAt,
                    IsActive = t.IsActive,
                    IsDefault = t.IsDefault,
                })
                .ToListAsync();
        }

        /// <summary>
        /// Marks (or unmarks) a template as its owner's default — the one auto-loaded into a
        /// fresh compose. Setting it clears the flag on the owner's other templates so each
        /// owner (CreatedBy email) has at most one default. IsDefault is managed only here;
        /// UpdateTemplate deliberately leaves it untouched so edits don't drop the flag.
        /// </summary>
        /// <param name="templateId">Primary key of the template to flag.</param>
        /// <param name="isDefault">True to make it the owner's default; false to clear it.</param>
        /// <returns>The updated template, or null when no row has that ID.</returns>
        public async Task<TemplateModel?> SetDefaultAsync(int templateId, bool isDefault)
        {
            var template = await _templateDbContext.Templates.FindAsync(templateId);
            if (template == null)
            {
                return null;
            }

            if (isDefault)
            {
                var owner = template.CreatedBy;
                var otherDefaults = await _templateDbContext.Templates
                    .Where(t => t.TemplateId != templateId && t.IsDefault && t.CreatedBy == owner)
                    .ToListAsync();
                foreach (var other in otherDefaults)
                {
                    other.IsDefault = false;
                    other.UpdatedAt = DateTime.UtcNow;
                }
            }

            template.IsDefault = isDefault;
            template.UpdatedAt = DateTime.UtcNow;
            await _templateDbContext.SaveChangesAsync();
            return template;
        }

        /// <summary>
        /// Hard-deletes the template with the given ID, if it exists.
        /// </summary>
        /// <param name="templateId">Primary key of the template to remove.</param>
        /// <returns>True when a row was found and deleted; false when the ID was not present.</returns>
        public async Task<bool> DeleteTemplate(int templateId)
        {
            var template = await _templateDbContext.Templates.FindAsync(templateId);

            if (template == null)
            {
                return false;
            }
            _templateDbContext.Templates.Remove(template);
            await _templateDbContext.SaveChangesAsync();
            return true;
        }

        /// <summary>
        /// Loads only templates whose IsActive flag is true.
        /// </summary>
        /// <returns>The active templates; empty when none are active.</returns>
        public async Task<List<TemplateModel>> GetActiveTemplatesAsync()
        {
            return await _templateDbContext.Templates
                                           .Where(t => t.IsActive)
                                           .ToListAsync();
        }

        /// <summary>
        /// Loads every template whose TemplateType exactly equals the given value, ignoring active state.
        /// </summary>
        /// <param name="templateType">Template type to filter on.</param>
        /// <returns>The matching templates; empty when none share that type.</returns>
        public async Task<List<TemplateModel>> GetByTypeAsync(string templateType)
        {
            return await _templateDbContext.Templates
                                           .Where(t => t.TemplateType == templateType)
                                           .ToListAsync();
        }
    }
}
