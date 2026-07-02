using Elpis_CRM.Model;
using Elpis_CRM.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Security.Claims;
using Elpis_CRM.Data;
using Microsoft.EntityFrameworkCore;

namespace Elpis_CRM.Controllers
{
    /// <summary>
    /// Provides API endpoints to manage deals and deal analytics.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class DealController : ControllerBase
    {
        private const string JwtScheme = "Bearer";
        private readonly DealsService _dealService;
        private readonly AppDbContext _dealsDb;
        private readonly RecycleBinService _recycleBinService;

        /// <summary>
        /// Initializes a new instance of the <see cref="DealController"/>.
        /// </summary>
        /// <param name="dealService">Service for deal operations.</param>
        /// <param name="dealsDb">Application database context.</param>
        public DealController(DealsService dealService, AppDbContext dealsDb, RecycleBinService recycleBinService)
        {
            _dealService = dealService;
            _dealsDb = dealsDb;
            _recycleBinService = recycleBinService;
        }

        /// <summary>
        /// Fetches a single deal by ID, with its linked contacts resolved into the contact fields.
        /// </summary>
        /// <param name="dealId">Unique identifier of the deal.</param>
        /// <returns>The matching deal.</returns>
        /// <response code="200">Deal found</response>
        /// <response code="404">No deal exists with the given ID</response>
        [HttpGet("{dealId:long}")]
        [Authorize(AuthenticationSchemes = JwtScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<DealModel>> GetById(long dealId)
        {
            try
            {
                var deal = await _dealService.GetByIdAsync(dealId);
                return Ok(deal);
            }
            catch (KeyNotFoundException)
            {
                return NotFound($"Deal with ID '{dealId}' not found.");
            }
        }

        /// <summary>
        /// Returns all deals (no pagination, since the Kanban board needs every card across all stages),
        /// optionally narrowed by a server-side search term.
        /// </summary>
        /// <param name="search">Optional term (min 2 chars); split into words that must each match one of
        /// Name, AccountName, ContactName, Stage, Owner, Territory, Tags, or Pipeline.</param>
        /// <returns>Flat list of deals ordered by most recent activity, with contact fields enriched.</returns>
        /// <response code="200">Deals retrieved successfully</response>
        [HttpGet]
        [Authorize(AuthenticationSchemes = JwtScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<List<DealModel>>> GetAll([FromQuery] string? search = null)
        {
            var deals = await _dealService.GetAllAsync(search);
            return Ok(deals);
        }

        /// <summary>
        /// Typeahead search across deal name, account, contact, stage, owner, territory, and tags;
        /// a purely numeric query also matches by deal ID. Returns an empty list for queries under 2 characters.
        /// </summary>
        /// <param name="q">Search term; must be at least 2 characters after trimming or no results are returned.</param>
        /// <param name="limit">Maximum rows to return; clamped to the range 1-100 (default 50).</param>
        /// <returns>Matching deals, newest-activity first, with contact fields resolved from deal-contact links.</returns>
        /// <response code="200">Search completed (list may be empty)</response>
        [HttpGet("search")]
        [Authorize(AuthenticationSchemes = JwtScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<List<DealModel>>> Search([FromQuery] string q, [FromQuery] int limit = 50)
        {
            var results = await _dealService.SearchAsync(q, limit);
            return Ok(results);
        }

        /// <summary>
        /// Returns the distinct set of tags across all deals, splitting each deal's comma-separated Tags field.
        /// </summary>
        /// <returns>Deduplicated, trimmed tag values.</returns>
        /// <response code="200">Tags retrieved successfully</response>
        [HttpGet("tags/all")]
        [Authorize(AuthenticationSchemes = JwtScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<List<string>>> GetAllTagsAsync()
        {
            var tags = await _dealService.GetAllTagsAsync();
            return Ok(tags);
        }

        /// <summary>
        /// Returns the deals carrying any of the supplied tags. A deal matches if at least one of its
        /// comma-separated tags equals one of the requested tags.
        /// </summary>
        /// <param name="tags">Comma-separated list of tags to match against.</param>
        /// <returns>Deals that share at least one of the given tags.</returns>
        /// <response code="200">Matching deals retrieved successfully</response>
        /// <response code="400">Tags parameter is missing or blank</response>
        [HttpGet("tags/dealName")]
        [Authorize(AuthenticationSchemes = JwtScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<string>> GetDealByTag([FromQuery] string tags)
        {
            if (string.IsNullOrWhiteSpace(tags))
            {
                return BadRequest("Tags parameter is required.");
            }

            var dealName = await _dealService.GetDealsByTagsAsync(tags);
            return Ok(dealName);
        }

        /// <summary>
        /// Counts deals created since the start of the current week (weeks run Monday through Sunday).
        /// </summary>
        /// <returns>Number of deals created this week.</returns>
        /// <response code="200">Deal count retrieved successfully</response>
        [HttpGet("ThisWeek/count")]
        [Authorize(AuthenticationSchemes = JwtScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<int>> GetThisWeekDealCount()
        {
            int count = await _dealService.GetThisWeekDealCountAsync();
            return Ok(count);
        }

        /// <summary>
        /// Provides a three-week trend of deal-creation counts for charting.
        /// </summary>
        /// <returns>A three-element list ordered [two weeks ago, last week, this week].</returns>
        /// <response code="200">Weekly deal analytics retrieved successfully</response>
        [HttpGet("AllWeeks/Analysis")]
        [Authorize(AuthenticationSchemes = JwtScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<List<int>>> GetAllWeeksDeals()
        {
            var result = await _dealService.GetDealsForAllWeeksAsync();
            return Ok(result);
        }

        /// <summary>
        /// Creates a deal, stamping CreatedBy/CreatedById from the caller's JWT claims before delegating to
        /// the service (which applies stage, pipeline, and close-date defaults and links contacts).
        /// </summary>
        /// <param name="deal">Deal payload from the request body.</param>
        /// <returns>The persisted deal, including its generated ID.</returns>
        /// <response code="201">Deal created; Location header points at the new deal</response>
        /// <response code="400">Model validation failed or a referenced account/contact does not exist</response>
        /// <response code="500">Unexpected error while saving the deal</response>
        [HttpPost]
        [Authorize(AuthenticationSchemes = JwtScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<DealModel>> Create([FromBody] DealModel deal)
        {
            try
            {
                var createdByClaimValue = User.FindFirst(ClaimTypes.Name)?.Value;
                var createdByIdClaimValue = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;

                if (!string.IsNullOrWhiteSpace(createdByClaimValue))
                {
                    deal.CreatedBy = createdByClaimValue;
                }

                if (long.TryParse(createdByIdClaimValue, out var createdById))
                {
                    deal.CreatedById = createdById;
                }

                Console.WriteLine($"[DealsController.Create] Incoming payload: Name='{deal.Name}', Value={deal.DealValue}, AccountId={deal.AccountId}");

                if (!ModelState.IsValid)
                {
                    var errors = ModelState.Values.SelectMany(v => v.Errors).Select(e => e.ErrorMessage);
                    Console.WriteLine($"[DealsController.Create] ModelState invalid: {string.Join(", ", errors)}");
                    return BadRequest(ModelState);
                }

                var created = await _dealService.AddAsync(deal);
                Console.WriteLine($"[DealsController.Create] SUCCESS: Created deal ID={created.Id}");

                return CreatedAtAction(
                    nameof(GetById),
                    new { dealId = created.Id },
                    created
                );
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine($"[DealsController.Create] VALIDATION ERROR: {ex.Message}");
                return BadRequest(new { error = ex.Message });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DealsController.Create] ERROR: {ex.Message}");
                if (ex.InnerException != null){
                    Console.WriteLine($"[DealsController.Create] Inner: {ex.InnerException.Message}");
                }
                return StatusCode(500, new
                {
                    error = "Internal server error",
                    details = ex.Message,
                    inner = ex.InnerException?.Message
                });
            }
        }

        /// <summary>
        /// Overwrites an existing deal with the supplied values, stamping UpdatedBy/UpdatedById from the
        /// caller's JWT claims and re-syncing its contact links.
        /// </summary>
        /// <param name="dealId">ID of the deal to update.</param>
        /// <param name="deal">New field values for the deal.</param>
        /// <returns>The updated deal.</returns>
        /// <response code="200">Deal updated successfully</response>
        /// <response code="404">No deal exists with the given ID</response>
        [HttpPut("{dealId:long}")]
        [Authorize(AuthenticationSchemes = JwtScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<DealModel>> Update(long dealId, [FromBody] DealModel deal)
        {
            try
            {
                var updatedByClaim = User.FindFirst(ClaimTypes.Name)?.Value;
                if (!string.IsNullOrWhiteSpace(updatedByClaim))
                {
                    deal.UpdatedBy = updatedByClaim;
                }

                if (long.TryParse(User.FindFirst(ClaimTypes.NameIdentifier)?.Value, out var updatedById))
                {
                    deal.UpdatedById = updatedById;
                }

                var updated = await _dealService.UpdateAsync(dealId, deal);
                return Ok(updated);
            }
            catch (KeyNotFoundException)
            {
                return NotFound($"Deal with ID '{dealId}' not found.");
            }
        }

        /// <summary>
        /// Removes a deal along with its related call logs and meetings. Admin only.
        /// </summary>
        /// <param name="dealId">ID of the deal to delete.</param>
        /// <returns>A plain confirmation message.</returns>
        /// <response code="200">Deal deleted successfully</response>
        /// <response code="404">No deal exists with the given ID</response>
        [HttpDelete("{dealId:long}")]
        [Authorize(AuthenticationSchemes = JwtScheme, Roles = "Admin")]
        public async Task<ActionResult> Delete(long dealId)
        {
            try
            {
                var deal = await _dealService.GetByIdAsync(dealId);
                await _dealService.DeleteAsync(dealId, User?.Identity?.Name ?? User?.FindFirst(ClaimTypes.Email)?.Value ?? "System");
                await _recycleBinService.CreateEntryAsync("Deal", dealId.ToString(), deal?.Name ?? "Unnamed Deal", "Deal deleted", User?.Identity?.Name ?? User?.FindFirst(ClaimTypes.Email)?.Value ?? "System", deal);
                return Ok("Deleted Successfully");
            }
            catch (KeyNotFoundException)
            {
                return NotFound($"Deal with ID '{dealId}' not found.");
            }
        }

        /// <summary>
        /// Returns the deals whose CreatedBy matches the caller's name claim, ordered by most recent
        /// activity (UpdatedAt, falling back to CreatedAt).
        /// </summary>
        /// <returns>The current user's deals, newest activity first.</returns>
        /// <response code="200">User's deals retrieved successfully</response>
        /// <response code="400">The name claim is missing from the token</response>
        [Authorize(AuthenticationSchemes = JwtScheme)]
        [HttpGet("my-deals")]
        public async Task<IActionResult> GetMyDeals()
        {
            var createdByClaimValue = User.FindFirst(ClaimTypes.Name)?.Value;
            if (string.IsNullOrEmpty(createdByClaimValue))
            {
                return BadRequest("User name not found in token");
            }

            var deals = await _dealsDb.Deals
                .Where(d => d.CreatedBy == createdByClaimValue)
                .OrderByDescending(d => d.UpdatedAt ?? d.CreatedAt)
                .ThenByDescending(d => d.CreatedAt)
                .ToListAsync();

            return Ok(deals);
        }
        /// <summary>
        /// Returns the deals belonging to the named pipeline, matched case-insensitively and ordered by
        /// most recent activity. An empty or unknown pipeline yields an empty list.
        /// </summary>
        /// <param name="pipeline">Pipeline name to filter on (for example: software, hardware, default).</param>
        /// <returns>Deals in the pipeline, with contact fields enriched.</returns>
        /// <response code="200">Deals retrieved successfully</response>
        /// <response code="500">Unexpected error while querying deals</response>
        [HttpGet("pipeline/deals")]
        [Authorize(AuthenticationSchemes = JwtScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<List<DealModel>>> GetDealsByPipeline([FromQuery] string pipeline)
        {
            try
            {
                Console.WriteLine($"[DealsController.GetDealsByPipeline] Called with pipeline: '{pipeline}'");
                var deals = await _dealService.GetByPipelineAsync(pipeline);
                Console.WriteLine($"[DealsController.GetDealsByPipeline] Returned {deals.Count} deals");
                return Ok(deals);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DealsController.GetDealsByPipeline] Error: {ex.Message}");
                return StatusCode(500, new { error = ex.Message });
            }
        }

        /// <summary>
        /// Lists every deal associated with a contact, matching both rows linked through the deal-contact
        /// join table and legacy deals carrying the contact directly in ContactId.
        /// </summary>
        /// <param name="contactId">ID of the contact whose deals are requested.</param>
        /// <returns>The contact's deals ordered by most recent activity, with contact fields enriched.</returns>
        /// <response code="200">Deals retrieved (list may be empty)</response>
        [HttpGet("contact/{contactId}")]
        public async Task<IActionResult> GetByContact(long contactId)
        {
            var deals = await _dealService.GetByContactIdAsync(contactId);
            return Ok(deals);
        }
    }
}
