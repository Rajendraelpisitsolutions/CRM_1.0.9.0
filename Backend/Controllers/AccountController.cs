using Elpis_CRM.Model;
using Elpis_CRM.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authentication.JwtBearer;

namespace Elpis_CRM.Controllers
{
    /// <summary>
    /// CRUD and bulk-management endpoints for CRM accounts. All actions require a JWT bearer token;
    /// read operations are open to Admin/Manager/User while deletes are restricted to Admin.
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    public class AccountController : ControllerBase
    {
        private readonly AccountService _accountService;

        /// <summary>
        /// Creates the controller with the account data/service layer it delegates to.
        /// </summary>
        /// <param name="accountService">Service that performs the account queries and mutations against the database.</param>
        public AccountController(AccountService accountService)
        {
            _accountService = accountService;
        }

        /// <summary>
        /// Returns one page of accounts, newest first, optionally narrowed by a free-text search.
        /// The search splits on spaces and every token must match at least one field (AND across tokens);
        /// terms shorter than two characters are ignored. Page is floored to 1 and pageSize is clamped to 1-500.
        /// </summary>
        /// <param name="page">1-based page number; values below 1 are treated as 1.</param>
        /// <param name="pageSize">Rows per page, clamped to the 1-500 range.</param>
        /// <param name="search">Optional space-separated search text; matched against name, phone, website, city, country, sales owner, territory, tags and industry.</param>
        /// <returns>A wrapper carrying the page metadata (Page, PageSize, TotalCount, TotalPages) and the matching Items.</returns>
        /// <response code="200">The requested page of accounts and paging totals.</response>
        [HttpGet]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme,Roles = "Admin,Manager,User")]
        public async Task<IActionResult> GetAll([FromQuery] int page     = 1, [FromQuery] int pageSize = 150,[FromQuery] string? search = null)
        {
            page     = Math.Max(page, 1);
            pageSize = Math.Clamp(pageSize, 1, 500);

            var (items, totalCount) = await _accountService.GetAllAsync(page, pageSize, search);

            return Ok(new
            {
                Page      = page,
                PageSize  = pageSize,
                TotalCount = totalCount,
                TotalPages = (int)Math.Ceiling((double)totalCount / pageSize),
                Items     = items
            });
        }

        /// <summary>
        /// Typeahead lookup used by autocomplete UIs: returns accounts ordered by name, capped at
        /// <paramref name="limit"/>. A purely numeric query also matches on AccountId; queries under
        /// two characters yield an empty list.
        /// </summary>
        /// <param name="q">Search text (minimum two characters after trimming).</param>
        /// <param name="limit">Maximum rows to return; the service clamps this to 1-100.</param>
        /// <returns>Matching accounts, name-ascending; empty when the query is too short.</returns>
        /// <response code="200">The matching accounts (possibly an empty list).</response>
        [HttpGet("search")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<List<AccountModel>>> Search([FromQuery] string q, [FromQuery] int limit = 50)
        {
            var results = await _accountService.SearchAsync(q, limit);
            return Ok(results);
        }

        /// <summary>
        /// Lists every account whose CreatedAt falls on the given calendar day. Only the date part is
        /// used; the time component is ignored and matching is done over the [date, date+1day) UTC window.
        /// </summary>
        /// <param name="createdAt">Day to filter on; the time-of-day portion is discarded (yyyy-MM-dd is sufficient).</param>
        /// <returns>Accounts created on that day; empty when none match.</returns>
        /// <response code="200">Accounts created on the requested day.</response>
        [HttpGet("createdAt")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<List<AccountModel>>> GetByCreatedAt([FromQuery] DateTime createdAt)
        {
            var accounts = await _accountService.GetAccountsByCreatedAtAsync(createdAt);
            return Ok(accounts);
        }

        /// <summary>
        /// Returns the distinct set of tags in use across all accounts. The Tags column stores a
        /// comma-separated string per account; this flattens, trims and de-duplicates them.
        /// </summary>
        /// <returns>Each unique tag exactly once, unsorted.</returns>
        /// <response code="200">The distinct list of tags.</response>
        [HttpGet("tags/all")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<List<string>>> GetAllTagsAsync()
        {
            var tags = await _accountService.GetAllTagsAsync();
            return Ok(tags);
        }

        /// <summary>
        /// Returns accounts tagged with any of the supplied tags (OR semantics). Tags are matched
        /// exactly after trimming each account's comma-separated tag list.
        /// </summary>
        /// <param name="tags">Comma-separated tags to match against; required and non-blank.</param>
        /// <returns>Accounts carrying at least one of the requested tags.</returns>
        /// <response code="200">Accounts matching one or more of the tags.</response>
        /// <response code="400">The tags parameter was missing or blank.</response>
        [HttpGet("tags/accounts")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<List<AccountModel>>> GetAccountsByTagsAsync([FromQuery] string tags)
        {
            if (string.IsNullOrWhiteSpace(tags))
                return BadRequest("Tags parameter is required.");
            var accounts = await _accountService.GetAccountsByTagsAsync(tags);
            return Ok(accounts);
        }

        /// <summary>
        /// Fetches a single account by its identifier.
        /// </summary>
        /// <param name="id">AccountId to look up.</param>
        /// <returns>The matching account.</returns>
        /// <response code="200">The requested account.</response>
        /// <response code="404">No account exists with that id.</response>
        [HttpGet("{id:long}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<AccountModel>> GetById(long id)
        {
            var account = await _accountService.GetByIdAsync(id);
            if (account == null) return NotFound();
            return Ok(account);
        }

        /// <summary>
        /// Creates a new account. Any client-supplied AccountId is discarded (forced to 0) so the
        /// service always generates a fresh identifier, and CreatedAt/UpdatedAt are stamped server-side.
        /// </summary>
        /// <param name="account">The account to create; AccountId is ignored.</param>
        /// <returns>The persisted account including its generated id, with a Location header pointing at <see cref="GetById"/>.</returns>
        /// <response code="201">The account was created.</response>
        /// <response code="400">The request body was missing.</response>
        [HttpPost]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<AccountModel>> Create([FromBody] AccountModel account)
        {
            if (account == null) return BadRequest("Account payload is required.");
            account.AccountId = 0;
            var created = await _accountService.AddAsync(account);
            return CreatedAtAction(nameof(GetById), new { id = created.AccountId }, created);
        }

        /// <summary>
        /// Overwrites the mutable fields of an existing account with the supplied values and refreshes
        /// UpdatedAt. The route id, not any id in the body, identifies the record.
        /// </summary>
        /// <param name="id">AccountId of the record to update.</param>
        /// <param name="account">New field values to apply.</param>
        /// <returns>The updated account.</returns>
        /// <response code="200">The account was updated.</response>
        /// <response code="404">No account exists with that id.</response>
        [HttpPut("{id:long}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<AccountModel>> Update(long id, [FromBody] AccountModel account)
        {
            var updated = await _accountService.UpdateAsync(id, account);
            if (updated == null) return NotFound();
            return Ok(updated);
        }

        /// <summary>
        /// Deletes an account and, to satisfy the foreign-key constraint, cascades the delete to any
        /// Deals linked to it. Admin only.
        /// </summary>
        /// <param name="id">AccountId to delete.</param>
        /// <response code="204">The account (and its related deals) were deleted.</response>
        /// <response code="404">No account exists with that id.</response>
        [HttpDelete("{id:long}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin")]
        public async Task<IActionResult> Delete(long id)
        {
            var deleted = await _accountService.DeleteAsync(id);
            if (!deleted) return NotFound();
            return NoContent();
        }

        /// <summary>
        /// Wipes every account from the database, first removing all Deals to clear foreign-key
        /// constraints. Destructive and irreversible; Admin only.
        /// </summary>
        /// <returns>A success message, or a 500 payload describing the failure.</returns>
        /// <response code="200">All accounts (and deals) were deleted.</response>
        /// <response code="500">Deletion failed; the body carries the error message.</response>
        [HttpDelete("delete-all")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin")]
        public async Task<ActionResult> DeleteAllAsync()
        {
            try
            {
                await _accountService.DeleteAllAsync();
                return Ok(new { message = "All accounts deleted successfully" });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = "Error deleting accounts", error = ex.Message });
            }
        }

        // POST: api/account/bulk-delete  body: { "ids": [1, 2, 3] }
        /// <summary>
        /// Deletes several accounts whose ids are passed in the request body, cascading to their related
        /// Deals first. Admin only.
        /// </summary>
        /// <param name="request">Payload whose <c>ids</c> list names the accounts to delete.</param>
        /// <returns>A message with the number of accounts actually deleted, or a 500 payload on failure.</returns>
        /// <response code="200">Deletion succeeded; body reports the deleted count.</response>
        /// <response code="400">No ids were supplied.</response>
        /// <response code="500">Deletion failed; the body carries the error message.</response>
        [HttpPost("bulk-delete")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin")]
        public async Task<ActionResult> BulkDeleteAsync([FromBody] DeleteBulkRequest request)
        {
            var idsToDelete = request?.ids;
            if (idsToDelete == null || idsToDelete.Count == 0)
                return BadRequest("At least one account ID is required.");
            try
            {
                int deletedCount = await _accountService.DeleteMultipleAsync(idsToDelete);
                return Ok(new { message = $"{deletedCount} accounts deleted successfully", deletedCount });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = "Error deleting accounts", error = ex.Message });
            }
        }

        // DELETE: api/account/bulk?ids=1&ids=2&ids=3
        /// <summary>
        /// Query-string variant of the bulk delete: removes the accounts named by repeated <c>ids</c>
        /// parameters, cascading to their related Deals first. Admin only.
        /// </summary>
        /// <param name="ids">Account ids to delete, supplied as repeated query parameters.</param>
        /// <returns>A message with the number of accounts actually deleted, or a 500 payload on failure.</returns>
        /// <response code="200">Deletion succeeded; body reports the deleted count.</response>
        /// <response code="400">No ids were supplied.</response>
        /// <response code="500">Deletion failed; the body carries the error message.</response>
        [HttpDelete("bulk")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin")]
        public async Task<ActionResult> BulkDeleteQueryAsync([FromQuery] List<long> ids)
        {
            if (ids == null || ids.Count == 0)
                return BadRequest("At least one account ID is required.");
            try
            {
                int deletedCount = await _accountService.DeleteMultipleAsync(ids);
                return Ok(new { message = $"{deletedCount} accounts deleted successfully", deletedCount });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = "Error deleting accounts", error = ex.Message });
            }
        }
    }
}
