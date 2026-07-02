using Elpis_CRM.Model;
using Elpis_CRM.Services;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;
using Elpis_CRM.Model.DTOs;

namespace Elpis_CRM.Controllers
{
    /// <summary>
    /// Exposes CRUD endpoints for notes and lookups scoped to a contact or a deal.
    /// All actions require an authenticated Admin, Manager, or User; deletion is further restricted to Admin.
    /// </summary>
    [Route("api/[controller]")]
[ApiController]
[Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
    public class NotesController : ControllerBase
    {
        private readonly NotesService _notesService;
        private readonly RecycleBinService _recycleBinService;

        /// <summary>
        /// Creates the controller with the note service used to persist and query notes.
        /// </summary>
        /// <param name="notesService">Service that handles note storage, mirroring, and lookups.</param>
        public NotesController(NotesService notesService,RecycleBinService recycleBinService)
        {
            _notesService = notesService;
            _recycleBinService = recycleBinService;
        }

        /// <summary>
        /// Creates a note. When the note targets a deal panel, the service also writes
        /// independent contact copies for any ids the caller selected (and vice versa for contact notes).
        /// </summary>
        /// <param name="note">Note payload; may carry MirrorToDealIds/MirrorToContactIds to fan out independent copies.</param>
        /// <returns>201 with the stored note and a Location header pointing at the new resource.</returns>
        /// <response code="201">The note was created.</response>
        /// <response code="400">The request body was missing or could not be bound.</response>
        [HttpPost]
        public async Task<IActionResult> Add([FromBody] NotesModel note)
        {
            if (note == null)
            {
                return BadRequest("Invalid note data");
            }
            var created = await _notesService.AddNote(note);

            return CreatedAtAction(nameof(GetById), new { id = created.Id }, created);
        }

        /// <summary>
        /// Returns every note in the system, newest first.
        /// </summary>
        /// <returns>200 with the full list of notes (empty list if none exist).</returns>
        /// <response code="200">Notes retrieved.</response>
        [HttpGet]
        public async Task<IActionResult> GetAll()
        {
            var notes = await _notesService.GetAllNotes();
            return Ok(notes);
        }

        /// <summary>
        /// Fetches a single note by its primary key.
        /// </summary>
        /// <param name="id">Identifier of the note to retrieve.</param>
        /// <returns>200 with the note, or 404 if no note has that id.</returns>
        /// <response code="200">The note was found.</response>
        /// <response code="404">No note exists with the given id.</response>
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetById(int id)
        {
            var note = await _notesService.GetById(id);

            if (note == null)
            {
                return NotFound("Note not found");
            }
            return Ok(note);
        }

        /// <summary>
        /// Lists the notes attached to a contact as DTOs, ordered newest first.
        /// </summary>
        /// <param name="contactId">Identifier of the contact whose notes are requested.</param>
        /// <returns>200 with the contact's notes (empty list if the contact has none).</returns>
        /// <response code="200">Contact notes retrieved.</response>
        [HttpGet("contact/{contactId:long}")]
        public async Task<ActionResult<List<ContactNoteDto>>> GetByContact(long contactId)
        {
            var notes = await _notesService.GetNotesByContactAsync(contactId);
            return Ok(notes);
        }


        /// <summary>
        /// Overwrites the description and author of a single note. Because mirrored copies are
        /// independent rows, this affects only the targeted note, not any copies of it.
        /// </summary>
        /// <param name="id">Identifier of the note to update.</param>
        /// <param name="note">Payload supplying the new description and author.</param>
        /// <returns>200 with the updated note, or 404 if the note does not exist.</returns>
        /// <response code="200">The note was updated.</response>
        /// <response code="404">No note exists with the given id.</response>
        [HttpPut("{id:int}")]
        public async Task<IActionResult> Update(int id, [FromBody] NotesModel note)
        {
            var updated = await _notesService.UpdateNote(id, note);

            if (updated == null)
            {
                return NotFound("Note not found");
            }
            return Ok(updated);
        }

        /// <summary>
        /// Permanently removes a single note. Admin-only; mirrored copies are left untouched.
        /// </summary>
        /// <param name="id">Identifier of the note to delete.</param>
        /// <returns>200 with a confirmation message, or 404 if the note does not exist.</returns>
        /// <response code="200">The note was deleted.</response>
        /// <response code="404">No note exists with the given id.</response>
        [HttpDelete("{id:int}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin")]
        public async Task<IActionResult> Delete(int id)
        {
            var note = await _notesService.GetById(id);

            if (note == null)
            {
                return NotFound("Note not found");
            }

            var deleted = await _notesService.DeleteNote(
                id,
                User?.Identity?.Name ??
                User?.FindFirst(System.Security.Claims.ClaimTypes.Email)?.Value ??
                "System");

            if (!deleted)
            {
                return NotFound("Note not found");
            }

            await _recycleBinService.CreateEntryAsync(
             "Note",
             note.Id.ToString(),
             note.RelatedToType == "Contact" ? "Contact Note" : "Deal Note",
             note.Description,
             User?.Identity?.Name ??
             User?.FindFirst(System.Security.Claims.ClaimTypes.Email)?.Value ??
             "System",
             note);

            return Ok("Note deleted successfully");
        }

        /// <summary>
        /// Lists the notes attached to a deal as DTOs, ordered newest first.
        /// </summary>
        /// <param name="dealId">Identifier of the deal whose notes are requested.</param>
        /// <returns>200 with the deal's notes (empty list if the deal has none).</returns>
        /// <response code="200">Deal notes retrieved.</response>
        [HttpGet("deal/{dealId:long}")]
        public async Task<ActionResult<List<DealNoteDto>>> GetByDeal(long dealId)
        {
            var notes = await _notesService.GetNotesByDealAsync(dealId);
            return Ok(notes);
        }
    }
}