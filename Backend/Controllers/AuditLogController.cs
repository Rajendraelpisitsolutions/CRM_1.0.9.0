using System.Threading.Tasks;
using Elpis_CRM.Service;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace Elpis_CRM.Controllers
{
    /// <summary>
    /// Read-only access to the audit trail. Restricted to Admins.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin")]
    public class AuditLogController : ControllerBase
    {
        private readonly AuditLogService _auditLogService;

        public AuditLogController(AuditLogService auditLogService)
        {
            _auditLogService = auditLogService;
        }

        /// <summary>
        /// Returns a paged, newest-first list of audit entries, optionally filtered by entity, action,
        /// user or a free-text search.
        /// </summary>
        /// <response code="200">Audit entries returned.</response>
        /// <response code="401">No or invalid bearer token.</response>
        /// <response code="403">Caller is not an Admin.</response>
        [HttpGet]
        public async Task<IActionResult> Get(
            [FromQuery] string? entity,
            [FromQuery] string? action,
            [FromQuery] string? user,
            [FromQuery] string? search,
            [FromQuery] int page = 1,
            [FromQuery] int pageSize = 50)
        {
            var result = await _auditLogService.GetAsync(entity, action, user, search, page, pageSize);
            return Ok(result);
        }

        /// <summary>Distinct entity names present in the log (for a filter dropdown).</summary>
        [HttpGet("entities")]
        public async Task<IActionResult> GetEntities()
        {
            return Ok(await _auditLogService.GetEntitiesAsync());
        }
    }
}
