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
