using Elpis_CRM.Model;
using Elpis_CRM.Services;
using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;

namespace Elpis_CRM.Controllers
{
    /// <summary>
    /// Handles forgot password operations such as OTP generation,
    /// OTP validation, and password reset.
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    public class ForgotPasswordController : ControllerBase
    {
        private readonly ForgotPasswordService _forgotPasswordService;

        public ForgotPasswordController(ForgotPasswordService forgotPasswordService)
        {
            _forgotPasswordService = forgotPasswordService;
        }

        /// <summary>
        /// Generates a 4-digit OTP for the given email, keeps it server-side for 5 minutes, and emails it to the user.
        /// Note: the service returns a success message even when the address has no matching account.
        /// </summary>
        /// <param name="request">Carries the target <c>Email</c>; required and must be non-blank.</param>
        /// <returns>200 with the service's status message, or 400 if the email is missing.</returns>
        /// <response code="200">OTP generated and the email send attempted.</response>
        /// <response code="400">Email was not supplied.</response>
        [HttpPost("generate-otp")]
        public async Task<IActionResult> GenerateOtp([FromBody] EmailRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.Email))
            {
                return BadRequest(new { Message = "Email is required" });
            }
            var result = await _forgotPasswordService.GenerateOtpAsync(request.Email);
            return Ok(new { Message = result });
        }

        /// <summary>
        /// Checks the submitted OTP against the stored one for the email and, if it matches and is unexpired,
        /// flags it as verified so the password can subsequently be reset.
        /// </summary>
        /// <param name="request">Carries the <c>Email</c> and the numeric <c>Otp</c> to verify.</param>
        /// <returns>200 when the OTP is valid; otherwise 400 with the reason (not found, invalid or expired).</returns>
        /// <response code="200">OTP matched and marked validated.</response>
        /// <response code="400">Email missing, OTP not found, or OTP invalid/expired.</response>
        [HttpPost("validate-otp")]
        public IActionResult ValidateOtp([FromBody] OtpRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.Email))
            {
                return BadRequest(new { Message = "Email is required" });
            }
            var result = _forgotPasswordService.ValidateOtp(request.Email, request.Otp);

            if (result == "OTP validated successfully")
            {
                return Ok(new { Message = result });
            }
            return BadRequest(new { Message = result });
        }

        /// <summary>
        /// Resets the account's password, but only if the email's OTP was previously validated; the OTP entry is
        /// then discarded so it cannot be reused.
        /// </summary>
        /// <param name="request">Carries the <c>Email</c> and the <c>NewPassword</c> to store.</param>
        /// <returns>200 on success; otherwise 400 with the reason (OTP not validated or user not found).</returns>
        /// <response code="200">Password updated and the OTP cleared.</response>
        /// <response code="400">Email missing, OTP not validated, or no user for that email.</response>
        [HttpPost("update-password")]
        public async Task<IActionResult> UpdatePassword([FromBody] ResetPasswordRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.Email))
            {
                return BadRequest(new { Message = "Email is required" });
            }
            var result = await _forgotPasswordService.UpdatePasswordAsync(request.Email, request.NewPassword);

            if (result == "Password updated successfully")
            {
                return Ok(new { Message = result });
            }
            return BadRequest(new { Message = result });
        }
    }
}
