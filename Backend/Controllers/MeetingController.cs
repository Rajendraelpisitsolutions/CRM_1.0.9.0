using Elpis_CRM.Model;
using Elpis_CRM.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Elpis_CRM.Controllers
{
    /// <summary>
    /// Exposes CRUD endpoints for CRM meetings, including lookups by contact, account, and upcoming schedule.
    /// All routes require a JWT bearer token; deletion is restricted to Admins.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class MeetingController : ControllerBase
    {
        private readonly MeetingService _meetingService;

        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingController"/> with the meeting persistence service.
        /// </summary>
        /// <param name="meetingService">Service that handles meeting storage and retrieval.</param>
        public MeetingController(MeetingService meetingService)
        {
            _meetingService = meetingService;
        }

        /// <summary>
        /// Returns every meeting in the system, unfiltered.
        /// </summary>
        /// <returns>The full list of meetings.</returns>
        /// <response code="200">Meetings retrieved.</response>
        [HttpGet]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<List<MeetingModel>>> GetAll()
        {
            var meetings = await _meetingService.GetAllMeetingsAsync();
            return Ok(meetings);
        }

        /// <summary>
        /// Fetches a single meeting by its identifier.
        /// </summary>
        /// <param name="id">Primary key of the meeting.</param>
        /// <returns>The matching meeting, or a 404 if no meeting has that id.</returns>
        /// <response code="200">Meeting found.</response>
        /// <response code="404">No meeting exists with the given id.</response>
        [HttpGet("{id:int}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<MeetingModel>> GetById(int id)
        {
            var meeting = await _meetingService.GetMeetingByIdAsync(id);

            if (meeting == null)
            {
                return NotFound($"Meeting with ID '{id}' not found.");
            }
            return Ok(meeting);
        }

        /// <summary>
        /// Lists all meetings linked to a specific contact.
        /// </summary>
        /// <param name="contactId">Identifier of the contact whose meetings are requested.</param>
        /// <returns>Meetings associated with the contact; an empty list if none match.</returns>
        /// <response code="200">Meetings retrieved (possibly empty).</response>
        [HttpGet("contact/{contactId:long}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<List<MeetingModel>>> GetByContactId(long contactId)
        {
            var meetings = await _meetingService.GetMeetingsByContactIdAsync(contactId);
            return Ok(meetings);
        }

        /// <summary>
        /// Lists all meetings linked to a specific account.
        /// </summary>
        /// <param name="accountId">Identifier of the account whose meetings are requested.</param>
        /// <returns>Meetings associated with the account; an empty list if none match.</returns>
        /// <response code="200">Meetings retrieved (possibly empty).</response>
        [HttpGet("account/{accountId:long}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<List<MeetingModel>>> GetByAccountId(long accountId)
        {
            var meetings = await _meetingService.GetMeetingsByAccountIdAsync(accountId);
            return Ok(meetings);
        }

        /// <summary>
        /// Returns meetings whose start time is now or later (UTC), ordered by start time ascending.
        /// </summary>
        /// <returns>Future meetings sorted soonest-first; an empty list if none are scheduled.</returns>
        /// <response code="200">Upcoming meetings retrieved.</response>
        [HttpGet("upcoming")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<List<MeetingModel>>> GetUpcoming()
        {
            var meetings = await _meetingService.GetUpcomingMeetingsAsync();
            return Ok(meetings);
        }

        /// <summary>
        /// Persists a new meeting after model validation, stamping creation and update timestamps server-side.
        /// </summary>
        /// <param name="meeting">Meeting payload; the AddVideoConference flag is stored as-is and no external conferencing link is generated.</param>
        /// <returns>The created meeting with its assigned id, exposed via the GetById route.</returns>
        /// <response code="201">Meeting created.</response>
        /// <response code="400">The submitted meeting failed model validation.</response>
        [HttpPost]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<MeetingModel>> Create([FromBody] MeetingModel meeting)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }
            var created = await _meetingService.CreateMeetingAsync(meeting);

            return CreatedAtAction(nameof(GetById), new { id = created.MeetingId }, created);
        }

        /// <summary>
        /// Overwrites the editable fields of an existing meeting and refreshes its UpdatedAt timestamp.
        /// </summary>
        /// <param name="id">Primary key of the meeting to update.</param>
        /// <param name="meeting">Meeting payload carrying the new field values.</param>
        /// <returns>The updated meeting, or a 404 if no meeting has the given id.</returns>
        /// <response code="200">Meeting updated.</response>
        /// <response code="404">No meeting exists with the given id.</response>
        [HttpPut("{id:int}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<MeetingModel>> Update(int id, [FromBody] MeetingModel meeting)
        {
            var updated = await _meetingService.UpdateMeetingAsync(id, meeting);

            if (updated == null)
            {
                return NotFound($"Meeting with ID '{id}' not found.");
            }
            return Ok(updated);
        }

        /// <summary>
        /// Permanently removes a meeting. Restricted to Admins.
        /// </summary>
        /// <param name="id">Primary key of the meeting to delete.</param>
        /// <returns>A success message, or a 404 if no meeting has the given id.</returns>
        /// <response code="200">Meeting deleted.</response>
        /// <response code="404">No meeting exists with the given id.</response>
        [HttpDelete("{id:int}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin")]
        public async Task<ActionResult> Delete(int id)
        {
            var deleted = await _meetingService.DeleteMeetingAsync(id);

            if (!deleted)
            {
                return NotFound($"Meeting with ID '{id}' not found.");
            }
            return Ok("Deleted Successfully");
        }
    }
}