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
    /// Provides API endpoints to manage appointments.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class AppointmentsController : ControllerBase
    {
        private readonly AppointmentsService _appointmentsService;

        /// <summary>
        /// Initializes a new instance of the <see cref="AppointmentsController"/>.
        /// </summary>
        /// <param name="appointmentsService">Service for appointment operations.</param>
        public AppointmentsController(AppointmentsService appointmentsService)
        {
            _appointmentsService = appointmentsService;
        }

        /// <summary>
        /// Returns every appointment in the system, unfiltered.
        /// </summary>
        /// <returns>The full list of appointments; empty when none exist.</returns>
        /// <response code="200">Appointments retrieved successfully.</response>
        [HttpGet]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<List<AppointmentsModel>>> GetAll()
        {
            var appointments = await _appointmentsService.GetAllAppointmentsAsync();
            return Ok(appointments);
        }

        /// <summary>
        /// Looks up a single appointment by its identifier.
        /// </summary>
        /// <param name="id">Primary key of the appointment.</param>
        /// <returns>The matching appointment, or a 404 when no appointment carries that ID.</returns>
        /// <response code="200">Appointment found.</response>
        /// <response code="404">No appointment exists with the given ID.</response>
        [HttpGet("{id:int}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<AppointmentsModel>> GetById(int id)
        {
            var appointment = await _appointmentsService.GetAppointmentByIdAsync(id);
            if (appointment == null)
            {
                return NotFound($"Appointment with ID '{id}' not found.");
            }
            return Ok(appointment);
        }

        /// <summary>
        /// Returns the appointments linked to a given contact.
        /// </summary>
        /// <param name="contactId">Identifier of the contact whose appointments are wanted.</param>
        /// <returns>Appointments for that contact; empty when the contact has none.</returns>
        /// <response code="200">Appointments retrieved successfully.</response>
        [HttpGet("contact/{contactId:int}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<List<AppointmentsModel>>> GetByContactId(int contactId)
        {
            var appointments = await _appointmentsService.GetAppointmentsByContactIdAsync(contactId);
            return Ok(appointments);
        }

        /// <summary>
        /// Returns the appointments linked to a given account.
        /// </summary>
        /// <param name="accountId">Identifier of the account whose appointments are wanted.</param>
        /// <returns>Appointments for that account; empty when the account has none.</returns>
        /// <response code="200">Appointments retrieved successfully.</response>
        [HttpGet("account/{accountId:int}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<List<AppointmentsModel>>> GetByAccountId(int accountId)
        {
            var appointments = await _appointmentsService.GetAppointmentsByAccountIdAsync(accountId);
            return Ok(appointments);
        }

        /// <summary>
        /// Creates an appointment after validating the model; creation timestamps are set by the service.
        /// </summary>
        /// <param name="appointment">Appointment to persist, supplied in the request body.</param>
        /// <returns>The stored appointment with its generated ID, plus a Location header pointing at <see cref="GetById"/>.</returns>
        /// <response code="201">Appointment created successfully.</response>
        /// <response code="400">The submitted appointment failed model validation.</response>
        [HttpPost]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<AppointmentsModel>> Create([FromBody] AppointmentsModel appointment)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            var created = await _appointmentsService.CreateAppointmentAsync(appointment);
            return CreatedAtAction(nameof(GetById), new { id = created.AppointmentId }, created);
        }

        /// <summary>
        /// Applies the supplied changes to the appointment with the given ID; only a fixed subset of fields is updated.
        /// </summary>
        /// <param name="id">Primary key of the appointment to modify.</param>
        /// <param name="appointment">Replacement values for the editable appointment fields.</param>
        /// <returns>The updated appointment, or a 404 when the ID does not exist.</returns>
        /// <response code="200">Appointment updated successfully.</response>
        /// <response code="404">No appointment exists with the given ID.</response>
        [HttpPut("{id:int}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        public async Task<ActionResult<AppointmentsModel>> Update(int id, [FromBody] AppointmentsModel appointment)
        {
            var updated = await _appointmentsService.UpdateAppointmentAsync(id, appointment);
            if (updated == null)
            {
                return NotFound($"Appointment with ID '{id}' not found.");
            }
            return Ok(updated);
        }

        /// <summary>
        /// Permanently removes the appointment with the given ID. Restricted to Admin and Manager roles.
        /// </summary>
        /// <param name="id">Primary key of the appointment to delete.</param>
        /// <returns>A confirmation message on success, or a 404 when the ID does not exist.</returns>
        /// <response code="200">Appointment deleted successfully.</response>
        /// <response code="404">No appointment exists with the given ID.</response>
        [HttpDelete("{id:int}")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager")]
        public async Task<ActionResult> Delete(int id)
        {
            var deleted = await _appointmentsService.DeleteAppointmentAsync(id);
            if (!deleted)
            {
                return NotFound($"Appointment with ID '{id}' not found.");
            }
            return Ok("Deleted Successfully");
        }
    }
}
