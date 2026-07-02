using Elpis_CRM.Model;
using Elpis_CRM.Service;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using System.Collections.Generic;
using System.Threading.Tasks;
using Elpis_CRM.Services;

namespace Elpis_CRM.Controllers
{
    /// <summary>
    /// Provides API endpoints to manage contacts and contact-related tags.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class ContactController : ControllerBase
    {
        private readonly ContactService _contactService;
        private readonly RecycleBinService _recycleBinService;

        /// <summary>
        /// Initializes a new instance of the <see cref="ContactController"/>.
        /// </summary>
        /// <param name="contactService">Service for contact operations.</param>
        public ContactController(ContactService contactService, RecycleBinService recycleBinService)
        {
            _contactService = contactService;
            _recycleBinService = recycleBinService;
        }

        /// <summary>
        /// Returns one page of contacts ordered newest-first, optionally narrowed by a token-based
        /// search where every whitespace-separated term must match at least one contact field.
        /// </summary>
        /// <param name="page">1-based page number; values below 1 are clamped to 1.</param>
        /// <param name="pageSize">Rows per page; clamped to the range 1–500.</param>
        /// <param name="search">Optional search text; ignored unless it has at least 2 non-blank characters.</param>
        /// <returns>An object carrying the page, page size, total count, total page count and the page of contacts.</returns>
        /// <response code="200">Page returned (the item list may be empty).</response>
        [HttpGet]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<IActionResult> GetAllAsync(
            [FromQuery] int page = 1,
            [FromQuery] int pageSize = 150,
            [FromQuery] string? search = null)
        {
            // Validation
            page = Math.Max(page, 1);
            pageSize = Math.Clamp(pageSize, 1, 500);

            var (contacts, totalCount) =
                await _contactService.GetAllAsync(page, pageSize, search);

            return Ok(new
            {
                Page = page,
                PageSize = pageSize,
                TotalCount = totalCount,
                TotalPages = (int)Math.Ceiling((double)totalCount / pageSize),
                Items = contacts
            });
        }

        /// <summary>
        /// Typeahead lookup across name, email, phone and account fields, ordered by last name then
        /// first name; a purely numeric query also matches on contact ID.
        /// </summary>
        /// <param name="q">Search text; returns an empty list when it has fewer than 2 non-blank characters.</param>
        /// <param name="limit">Maximum rows to return; clamped to the range 1–100 by the service.</param>
        /// <returns>Matching contacts, or an empty list when the query is too short.</returns>
        /// <response code="200">Matches returned (possibly empty).</response>
        [HttpGet("search")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<List<ContactModel>>> SearchAsync([FromQuery] string q, [FromQuery] int limit = 50)
        {
            var results = await _contactService.SearchAsync(q, limit);
            return Ok(results);
        }

        /// <summary>
        /// Returns contact counts bucketed by lifecycle stage for dashboard charts, with stages folded
        /// into prospect, engaged, customer, promoter and other (null/blank/unrecognized stages count as other).
        /// </summary>
        /// <returns>A dictionary keyed by the five fixed stage buckets with their contact counts.</returns>
        /// <response code="200">Counts returned.</response>
        [HttpGet("stats/lifecycle")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<Dictionary<string, int>>> GetLifeCycleStatsAsync()
        {
            var stats = await _contactService.GetLifeCycleStageCountsAsync();
            return Ok(stats);
        }

        /// <summary>
        /// Returns every contact whose CreatedAt falls within the calendar day of the given date; the
        /// time component is ignored and only the date portion is used as the matching window.
        /// </summary>
        /// <param name="createdAt">Day to match on; only the date part is significant (yyyy-MM-dd recommended).</param>
        /// <returns>Contacts created on that day, or an empty list when none match.</returns>
        /// <response code="200">Contacts returned (possibly empty).</response>
        [HttpGet("createdAt")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<List<ContactModel>>> GetByCreatedAt([FromQuery] DateTime createdAt)
        {
            var contacts = await _contactService.GetContactsByCreatedAtAsync(createdAt);
            return Ok(contacts);
        }

        /// <summary>
        /// Returns the distinct set of individual tags used across all contacts, obtained by splitting each
        /// contact's comma-separated Tags field and trimming and de-duplicating the values.
        /// </summary>
        /// <returns>The unique tag names; empty when no contact carries any tag.</returns>
        /// <response code="200">Tags returned.</response>
        [HttpGet("tags/all")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<List<string>>> GetAllTagsAsync()
        {
            var tags = await _contactService.GetAllTagsAsync();
            return Ok(tags);
        }

        /// <summary>
        /// Collects the email addresses of every contact carrying at least one of the requested tags,
        /// pulling from both the WorkEmail and the multi-value Emails field and returning them
        /// de-duplicated (case-insensitive) as a single comma-joined string.
        /// </summary>
        /// <param name="tags">Comma-separated tags; a contact matches if any of its tags is in this set.</param>
        /// <returns>A comma-joined list of distinct email addresses; empty string when nothing matches.</returns>
        /// <response code="200">Emails returned (possibly an empty string).</response>
        /// <response code="400">The tags parameter was missing or blank.</response>
        [HttpGet("tags/emails")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<string>> GetEmailsByTagsAsync([FromQuery] string tags)
        {
            if (string.IsNullOrWhiteSpace(tags))
            {
                return BadRequest("Tags parameter is required.");
            }

            var emails = await _contactService.GetEmailsByTagsAsync(tags);
            return Ok(emails);
        }

        /// <summary>
        /// Returns the full contacts that carry at least one of the requested tags. Despite the route name,
        /// this yields complete contact records, identical to the tags/contacts endpoint.
        /// </summary>
        /// <param name="tags">Comma-separated tags; a contact matches if any of its tags is in this set.</param>
        /// <returns>The matching contacts, or an empty list when none match.</returns>
        /// <response code="200">Contacts returned (possibly empty).</response>
        /// <response code="400">The tags parameter was missing or blank.</response>
        [HttpGet("tags/contactName")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<string>> GetContactsByTags([FromQuery] string tags)
        {
            if (string.IsNullOrWhiteSpace(tags))
            {
                return BadRequest("Tags parameter is required.");
            }

            var contactNames = await _contactService.GetContactsByTagsAsync(tags);
            return Ok(contactNames);
        }

        /// <summary>
        /// Returns the full contacts that carry at least one of the requested tags, matching each
        /// contact's comma-separated Tags against the supplied set.
        /// </summary>
        /// <param name="tags">Comma-separated tags; a contact matches if any of its tags is in this set.</param>
        /// <returns>The matching contacts, or an empty list when none match.</returns>
        /// <response code="200">Contacts returned (possibly empty).</response>
        /// <response code="400">The tags parameter was missing or blank.</response>
        [HttpGet("tags/contacts")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<List<ContactModel>>> GetContactsByTagsFullAsync([FromQuery] string tags)
        {
            if (string.IsNullOrWhiteSpace(tags))
            {
                return BadRequest("Tags parameter is required.");
            }

            var contacts = await _contactService.GetContactsByTagsAsync(tags);
            return Ok(contacts);
        }

        /// <summary>
        /// Looks up a single contact by its primary key.
        /// </summary>
        /// <param name="contactId">Primary key of the contact.</param>
        /// <returns>The matching contact.</returns>
        /// <response code="200">Contact found.</response>
        /// <response code="404">No contact has that ID.</response>
        [HttpGet("{contactId:long}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<ContactModel>> GetByIdAsync(long contactId)
        {
            var contact = await _contactService.GetByIdAsync(contactId);

            if (contact == null)
            {
                return NotFound($"Contact with ID '{contactId}' not found.");
            }

            return Ok(contact);
        }

        /// <summary>
        /// Returns the contacts belonging to one account, ordered by name and capped for large accounts;
        /// each result is a lightweight projection (ID, name, work email and account only).
        /// </summary>
        /// <param name="accountId">The account whose contacts are requested.</param>
        /// <param name="q">Optional name/email/phone filter applied within the account.</param>
        /// <param name="limit">Maximum rows to return; clamped to the range 1–500 by the service.</param>
        /// <returns>The account's contacts as trimmed projections, or an empty list when none match.</returns>
        /// <response code="200">Contacts returned (possibly empty).</response>
        [HttpGet("account/{accountId:long}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<List<ContactModel>>> GetContactsByAccountIdAsync(long accountId, [FromQuery] string? q = null, [FromQuery] int limit = 300)
        {
            var contacts = await _contactService.GetContactsByAccountIdAsync(accountId, q, limit);
            return Ok(contacts);
        }

        /// <summary>
        /// Creates a contact: sets timestamps, resolves the account FK from the account name when only the
        /// name was supplied, rejects duplicate emails, and (when the ID already exists) updates instead.
        /// </summary>
        /// <param name="contact">Contact payload; any client-supplied EnquiryNo is ignored on create.</param>
        /// <param name="generateEnquiryNo">When true, assigns a fresh sequential EnquiryNo (e.g. EITSPL-EQ-003); otherwise it is left null.</param>
        /// <returns>The persisted contact, including its generated ID and EnquiryNo.</returns>
        /// <response code="200">Contact created (or an existing same-ID contact updated).</response>
        /// <response code="400">Payload was null or a contact with the same email already exists.</response>
        /// <response code="500">An unexpected error occurred while saving.</response>
        [HttpPost]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<ContactModel>> CreateAsync(
    [FromBody] ContactModel contact,
    [FromQuery] bool generateEnquiryNo = false,
    [FromQuery] bool generateEstimatedQuote = false)
        {
            if (contact == null)
            {
                return BadRequest("Contact data cannot be null.");
            }

            try
            {
                var created = await _contactService.AddAsync(contact, generateEnquiryNo, generateEstimatedQuote);
                return Ok(created);
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(new
                {
                    message = ex.Message
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new
                {
                    message = "An unexpected error occurred.",
                    detail = ex.InnerException?.Message ?? ex.Message
                });
            }
        }


        /// <summary>
        /// Overwrites an existing contact's editable fields and refreshes its activity timestamps; CreatedBy
        /// info is preserved and EnquiryNo is never taken from the payload (only optionally generated).
        /// </summary>
        /// <param name="contactId">Primary key of the contact to update.</param>
        /// <param name="contact">New field values; account FK is re-resolved from the account name when needed.</param>
        /// <param name="generateEnquiryNo">When true, assigns a sequential EnquiryNo only if the contact currently has none; the existing value is otherwise left untouched.</param>
        /// <returns>A confirmation message wrapping the updated contact.</returns>
        /// <response code="200">Contact updated.</response>
        /// <response code="400">Payload was null.</response>
        /// <response code="404">No contact has that ID.</response>
        [HttpPut("{contactId:long}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult> UpdateAsync(
            long contactId,
            [FromBody] ContactModel contact,
            [FromQuery] bool generateEnquiryNo = false,
            [FromQuery] bool generateEstimatedQuote = false)
        {
            if (contact == null)
            {
                return BadRequest("Contact data cannot be null.");
            }

            var updated = await _contactService.UpdateAsync(contactId, contact, generateEnquiryNo, generateEstimatedQuote);

            if (updated == null)
            {
                return NotFound($"Contact with ID '{contactId}' not found.");
            }

            return Ok(new
            {
                Message = "Contact updated successfully",
                Contact = updated
            });
        }

        /// <summary>
        /// Deletes a contact along with its related call logs, tasks and notes; restricted to the Admin role.
        /// </summary>
        /// <param name="contactId">Primary key of the contact to delete.</param>
        /// <returns>A confirmation message.</returns>
        /// <response code="200">Contact deleted.</response>
        /// <response code="404">No contact has that ID.</response>
        [HttpDelete("{contactId:long}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin")]
        public async Task<ActionResult> DeleteAsync(long contactId)
        {
            var contact = await _contactService.GetByIdAsync(contactId);
            if (contact == null)
            {
                return NotFound($"Contact with ID '{contactId}' not found.");
            }

            var deleted = await _contactService.DeleteAsync(contactId, User?.Identity?.Name ?? User?.FindFirst(System.Security.Claims.ClaimTypes.Email)?.Value ?? "System");

            if (!deleted)
            {
                return NotFound($"Contact with ID '{contactId}' not found.");
            }

            await _recycleBinService.CreateEntryAsync("Contact", contactId.ToString(), $"{contact.FirstName} {contact.LastName}".Trim(), "Contact deleted", User?.Identity?.Name ?? User?.FindFirst(System.Security.Claims.ClaimTypes.Email)?.Value ?? "System", contact);
            return Ok(new
            {
                Message = "Contact deleted successfully"
            });
        }

        /// <summary>
        /// Streams a contact's stored business-card image as JPEG, selecting the front or back side.
        /// </summary>
        /// <param name="contactId">Primary key of the contact owning the image.</param>
        /// <param name="imageType">Which side to return; "front" or "back" (case-insensitive).</param>
        /// <returns>The image bytes as an image/jpeg file response.</returns>
        /// <response code="200">Image returned.</response>
        /// <response code="404">Contact not found, the image type is unknown, or that side has no stored image.</response>
        [HttpGet("{contactId:long}/image/{imageType}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<IActionResult> GetImage(long contactId, string imageType)
        {
            var contact = await _contactService.GetByIdAsync(contactId);

            if (contact == null)
                return NotFound("Contact not found.");

            byte[]? imageBytes = imageType.ToLower() switch
            {
                "front" => contact.FrontImage,
                "back" => contact.BackImage,
                _ => null
            };

            if (imageBytes == null)
                return NotFound($"{imageType} image not found.");

            return File(imageBytes, "image/jpeg");
        }

        /// <summary>
        /// Returns the EnquiryNo values for the given contact IDs, skipping any contact that has no
        /// enquiry number assigned, so the result may be shorter than the input list.
        /// </summary>
        /// <param name="contactIds">Contact IDs to look up enquiry numbers for.</param>
        /// <returns>The non-blank enquiry numbers found among those contacts.</returns>
        /// <response code="200">Enquiry numbers returned (possibly empty).</response>
        [HttpPost("enquiry-numbers")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<List<string>>> GetEnquiryNumbers(
    [FromBody] List<long> contactIds)
        {
            var enquiryNumbers = await _contactService.GetEnquiryNumbersAsync(contactIds);
            return Ok(enquiryNumbers);
        }

        /// <summary>
        /// Returns the EstimatedQuote values for the given contact IDs, skipping any contact that has
        /// no estimated quote assigned, so the result may be shorter than the input list.
        /// </summary>
        /// <param name="contactIds">Contact IDs to look up estimated quotes for.</param>
        /// <returns>The non-blank estimated-quote numbers found among those contacts.</returns>
        /// <response code="200">Estimated quotes returned (possibly empty).</response>
        [HttpPost("estimated-quotes")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<List<string>>> GetEstimatedQuotes(
    [FromBody] List<long> contactIds)
        {
            var estimatedQuotes = await _contactService.GetEstimatedQuotesAsync(contactIds);
            return Ok(estimatedQuotes);
        }


    }
}
