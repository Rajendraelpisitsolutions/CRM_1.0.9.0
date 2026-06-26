using Elpis_CRM.Dtos;
using Elpis_CRM.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Threading.Tasks;

namespace Elpis_CRM.Controllers
{
    /// <summary>
    /// Public endpoint that receives "Contact Us" form submissions from the
    /// external website. Matches the submitter against existing contacts and
    /// either creates a deal on the existing contact, or creates a new contact
    /// and a deal for it.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class ContactUsController : ControllerBase
    {
        private readonly ContactUsService _contactUsService;

        /// <summary>
        /// Creates the controller with the service that processes Contact Us submissions.
        /// </summary>
        /// <param name="contactUsService">Service that matches/creates the contact and opens a deal.</param>
        public ContactUsController(ContactUsService contactUsService)
        {
            _contactUsService = contactUsService;
        }

        /// <summary>
        /// Anonymous endpoint that processes a public Contact Us submission: ensures the account, matches or
        /// creates the contact, and opens a new deal against it. Argument-validation failures from the service
        /// (e.g. missing first name or company) surface as 400.
        /// </summary>
        /// <param name="form">The form fields entered by the website visitor.</param>
        /// <returns>A result describing whether the contact and account were matched or created, plus the new deal's ID and name.</returns>
        /// <response code="200">Submission processed; contact and deal resolved.</response>
        /// <response code="400">Form was null, failed model validation, or a required field was missing.</response>
        /// <response code="500">An unexpected error occurred while processing the submission.</response>
        [HttpPost]
        [AllowAnonymous]
        public async Task<IActionResult> SubmitAsync([FromBody] ContactUsFormDto form)
        {
            if (form == null)
            {
                return BadRequest("Form data cannot be null.");
            }

            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            try
            {
                var result = await _contactUsService.SubmitAsync(form);
                return Ok(result);
            }
            catch (ArgumentException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = ex.Message, detail = ex.InnerException?.Message });
            }
        }
    }
}
