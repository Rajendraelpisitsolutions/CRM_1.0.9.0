using Elpis_CRM.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using System.Security.Claims;

namespace Elpis_CRM.Controllers
{
    /// <summary>
    /// Provides APIs for viewing, restoring, and permanently deleting items
    /// from the recycle bin.
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    public class RecycleBinController : ControllerBase
    {
        private readonly RecycleBinService _recycleBinService;

        /// <summary>
        /// Initializes a new instance of the <see cref="RecycleBinController"/> class.
        /// </summary>
        /// <param name="recycleBinService">Service used to manage recycle bin operations.</param>
        public RecycleBinController(RecycleBinService recycleBinService)
        {
            _recycleBinService = recycleBinService;
        }

        /// <summary>
        /// Retrieves all deleted items currently available in the recycle bin.
        /// </summary>
        /// <returns>A list of deleted items.</returns>
        [HttpGet]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<IActionResult> GetDeletedItems()
        {
            var items = await _recycleBinService.GetDeletedItemsAsync();
            return Ok(items);
        }

        /// <summary>
        /// Restores a deleted entity from the recycle bin.
        /// </summary>
        /// <param name="entityType">The type of entity to restore.</param>
        /// <param name="entityId">The unique identifier of the entity.</param>
        /// <returns>A success message if restored; otherwise, a not found response.</returns>
        [HttpPost("{entityType}/{entityId}/restore")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin")]
        public async Task<IActionResult> Restore(string entityType, string entityId)
        {
            var restoredBy = User?.Identity?.Name ?? User?.FindFirst(ClaimTypes.Email)?.Value ?? "System";
            var success = await _recycleBinService.RestoreAsync(entityType, entityId, restoredBy);

            if (!success)
            {
                return NotFound(new { message = "Deleted item not found." });
            }

            return Ok(new { message = "Restored successfully." });
        }

        /// <summary>
        /// Permanently deletes an entity from the recycle bin.
        /// </summary>
        /// <param name="entityType">The type of entity to delete.</param>
        /// <param name="entityId">The unique identifier of the entity.</param>
        /// <returns>A success message if deleted; otherwise, a not found response.</returns>
        [HttpDelete("{entityType}/{entityId}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin")]
        public async Task<IActionResult> PermanentDelete(string entityType, string entityId)
        {
            var success = await _recycleBinService.PermanentDeleteAsync(entityType, entityId);

            if (!success)
            {
                return NotFound(new { message = "Deleted item not found." });
            }

            return Ok(new { message = "Removed permanently." });
        }
    }
}