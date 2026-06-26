using Elpis_CRM.Model;
using Elpis_CRM.Service;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Elpis_CRM.Controllers
{
    /// <summary>
    /// Controller for managing templates in the CRM system.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class TemplateController : ControllerBase
    {
        private readonly TemplateService _templateService;

        /// <summary>
        /// Initializes a new instance of the <see cref="TemplateController"/> class.
        /// </summary>
        /// <param name="templateService">Service for template operations.</param>
        public TemplateController(TemplateService templateService)
        {
            _templateService = templateService;
        }

        /// <summary>
        /// Persists a new template; the service stamps created/updated timestamps before saving.
        /// </summary>
        /// <param name="template">Template payload (name, subject, body, type, active flag).</param>
        /// <returns>The saved template including its generated ID and timestamps.</returns>
        /// <response code="201">Template created; a Location header points to the new resource.</response>
        [HttpPost]
        public async Task<ActionResult<TemplateModel>> Create([FromBody] TemplateModel template)
        {
            var created = await _templateService.AddTemplate(template);
            return CreatedAtAction(nameof(GetById), new { id = created.TemplateId }, created);
        }

        /// <summary>
        /// Returns every stored template, both active and inactive, in no particular order.
        /// </summary>
        /// <returns>The full list of templates, wrapped in 200 OK.</returns>
        [HttpGet]
        public async Task<ActionResult<List<TemplateModel>>> GetAll()
        {
            return Ok(await _templateService.GetAllAsync());
        }

        /// <summary>
        /// Looks up a single template by its primary key.
        /// </summary>
        /// <param name="id">Primary key of the template.</param>
        /// <returns>The matching template, or a 404 when no template has that ID.</returns>
        /// <response code="200">Template found.</response>
        /// <response code="404">No template exists with the given ID.</response>
        [HttpGet("{id:int}")]
        public async Task<ActionResult<TemplateModel>> GetById(int id)
        {
            var template = await _templateService.GetByIdAsync(id);
            if (template == null)
            {
                return NotFound($"No template found for ID '{id}'.");
            }
            return Ok(template);
        }

        /// <summary>
        /// Finds the first template whose name matches exactly.
        /// </summary>
        /// <param name="name">Exact template name to match.</param>
        /// <returns>The matching template, or a 404 when no template carries that name.</returns>
        /// <response code="200">Template found.</response>
        /// <response code="404">No template matches the given name.</response>
        [HttpGet("name/{name}")]
        public async Task<ActionResult<TemplateModel>> GetByName(string name)
        {
            var template = await _templateService.GetByNameAsync(name);
            if (template == null)
            {
                return NotFound($"No template found for name '{name}'.");
            }
            return Ok(template);
        }

        /// <summary>
        /// Returns only templates flagged as active, omitting any with IsActive set to false.
        /// </summary>
        /// <returns>The active templates, wrapped in 200 OK (empty list if none are active).</returns>
        [HttpGet("active")]
        public async Task<ActionResult<List<TemplateModel>>> GetActiveTemplates()
        {
            var templates = await _templateService.GetActiveTemplatesAsync();
            return Ok(templates);
        }

        /// <summary>
        /// Returns all templates whose type matches the supplied value (for example "Email" or "SMS").
        /// </summary>
        /// <param name="templateType">Template type to filter on.</param>
        /// <returns>The matching templates, wrapped in 200 OK (empty list if none match).</returns>
        [HttpGet("type/{templateType}")]
        public async Task<ActionResult<List<TemplateModel>>> GetByType(string templateType)
        {
            var templates = await _templateService.GetByTypeAsync(templateType);
            return Ok(templates);
        }

        /// <summary>
        /// Overwrites an existing template's name, subject, body, type and active flag, refreshing its UpdatedAt timestamp.
        /// </summary>
        /// <param name="id">Primary key of the template to update.</param>
        /// <param name="template">Payload carrying the new field values.</param>
        /// <returns>A plain "Updated successfully" message, or a 404 when the ID is unknown.</returns>
        /// <response code="200">Template updated.</response>
        /// <response code="404">No template exists with the given ID.</response>
        [HttpPut("{id:int}")]
        public async Task<ActionResult<TemplateModel>> UpdateTemplate(int id, [FromBody] TemplateModel template)
        {
            var updated = await _templateService.UpdateTemplate(id, template);
            if (updated == null)
            {
                return NotFound($"Template with ID '{id}' not found.");
            }
            return Ok("Updated successfully");
        }

        /// <summary>
        /// Permanently removes the template with the given ID.
        /// </summary>
        /// <param name="id">Primary key of the template to delete.</param>
        /// <returns>A plain "Deleted Successfully" message, or a 404 when the ID is unknown.</returns>
        /// <response code="200">Template deleted.</response>
        /// <response code="404">No template exists with the given ID.</response>
        [HttpDelete("{id:int}")]
        public async Task<ActionResult> Delete(int id)
        {
            var result = await _templateService.DeleteTemplate(id);
            if (!result)
            {
                return NotFound($"Template with ID '{id}' not found.");
            }
            return Ok("Deleted Successfully");
        }
    }
}
