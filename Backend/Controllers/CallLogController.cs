using Elpis_CRM.Model;
using Elpis_CRM.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using System.Threading.Tasks;

namespace Elpis_CRM.Controllers
{
    /// <summary>
    /// Exposes API endpoints for recording and managing call logs.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class CallLogController : ControllerBase
    {
        private readonly CallLogService _callLogService;

        /// <summary>
        /// Initializes a new instance of the <see cref="CallLogController"/>.
        /// </summary>
        /// <param name="callLogService">Service that handles call log persistence and queries.</param>
        public CallLogController(CallLogService callLogService)
        {
            _callLogService = callLogService;
        }

        /// <summary>
        /// Returns every call log in the system, unfiltered.
        /// </summary>
        /// <returns>The full list of call logs; empty when none exist.</returns>
        /// <response code="200">Call logs retrieved successfully.</response>
        // GET: api/CallLog
        [HttpGet]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<IActionResult> GetAll()
        {
            var logs = await _callLogService.GetAllCallLogsAsync();
            return Ok(logs);
        }

        /// <summary>
        /// Returns the call logs recorded against a contact, newest first.
        /// </summary>
        /// <param name="contactId">Identifier of the contact whose calls are wanted.</param>
        /// <returns>The contact's call logs ordered by creation time descending; empty when the contact has none.</returns>
        /// <response code="200">Call logs retrieved successfully.</response>
        // GET: api/CallLog/contact/5
        [HttpGet("contact/{contactId}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<IActionResult> GetByContactId(long contactId)
        {
            var logs = await _callLogService.GetCallLogsByContactIdAsync(contactId);
            return Ok(logs);
        }

        /// <summary>
        /// Looks up a single call log by its identifier.
        /// </summary>
        /// <param name="id">Primary key of the call log.</param>
        /// <returns>The matching call log, or a 404 when no log carries that ID.</returns>
        /// <response code="200">Call log found.</response>
        /// <response code="404">No call log exists with the given ID.</response>
        // GET: api/CallLog/5
        [HttpGet("{id}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<IActionResult> GetById(long id)
        {
            var log = await _callLogService.GetCallLogByIdAsync(id);

            if (log == null)
            {
                return NotFound(new { message = $"Call log with ID {id} not found" });
            }
            return Ok(log);
        }

        /// <summary>
        /// Records a new call log; the service stamps <c>CreatedAt</c> only when the client leaves it unset.
        /// </summary>
        /// <param name="callLog">Call log to persist, supplied in the request body.</param>
        /// <returns>The stored call log with its generated ID, plus a Location header pointing at <see cref="GetById"/>.</returns>
        /// <response code="201">Call log created successfully.</response>
        /// <response code="400">The request body was missing or could not be bound.</response>
        // POST: api/CallLog
        [HttpPost]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<IActionResult> Create([FromBody] CallLogModel callLog)
        {
            if (callLog == null)
            {
                return BadRequest("Invalid call log data");
            }

            var createdLog = await _callLogService.CreateCallLogAsync(callLog);

            return CreatedAtAction(
                nameof(GetById),
                new { id = createdLog.CallLogId },
                createdLog
            );
        }

        /// <summary>
        /// Updates an existing call log; the route ID must match the body's <c>CallLogId</c>.
        /// </summary>
        /// <param name="id">Primary key from the route; must equal <paramref name="callLog"/>.CallLogId.</param>
        /// <param name="callLog">Replacement values for the editable call log fields.</param>
        /// <returns>The updated call log, a 400 on a missing body or ID mismatch, or a 404 when the ID does not exist.</returns>
        /// <response code="200">Call log updated successfully.</response>
        /// <response code="400">Body was missing or its CallLogId does not match the route ID.</response>
        /// <response code="404">No call log exists with the given ID.</response>
        // PUT: api/CallLog/5
        [HttpPut("{id}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<IActionResult> Update(long id, [FromBody] CallLogModel callLog)
        {
            if (callLog == null || id != callLog.CallLogId)
            {
                return BadRequest("CallLog ID mismatch");
            }

            var updatedLog = await _callLogService.UpdateCallLogAsync(callLog);

            if (updatedLog == null)
            {
                return NotFound(new { message = $"Call log with ID {id} not found" });
            }

            return Ok(updatedLog);
        }

        /// <summary>
        /// Permanently removes the call log with the given ID. Restricted to the Admin role.
        /// </summary>
        /// <param name="id">Primary key of the call log to delete.</param>
        /// <returns>A confirmation message on success, or a 404 when the ID does not exist.</returns>
        /// <response code="200">Call log deleted successfully.</response>
        /// <response code="404">No call log exists with the given ID.</response>
        // DELETE: api/CallLog/5
        [HttpDelete("{id}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin")]
        public async Task<IActionResult> Delete(long id)
        {
            var deleted = await _callLogService.DeleteCallLogAsync(id);

            if (!deleted)
            {
                return NotFound(new { message = $"Call log with ID {id} not found" });
            }

            return Ok(new { message = "Call log deleted successfully" });
        }
    }
}