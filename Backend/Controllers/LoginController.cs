using Elpis_CRM.Data;
using Elpis_CRM.Model;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text;
using Microsoft.AspNetCore.Authentication.JwtBearer;

namespace Elpis_CRM.Controllers
{
    /// <summary>
    /// Authentication and user-account management: credential login with JWT issuance,
    /// self-service profile/password updates, and Admin-only CRUD over login accounts.
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    public class LoginController : ControllerBase
    {
        private readonly AppDbContext _loginDb;
        private readonly IConfiguration _configuration;
        private readonly ILogger<LoginController> _logger;
        private readonly IWebHostEnvironment _environment;

        public LoginController(
            AppDbContext loginDb,
            IConfiguration configuration,
            ILogger<LoginController> logger,
            IWebHostEnvironment environment)
        {
            _loginDb = loginDb;
            _configuration = configuration;
            _logger = logger;
            _environment = environment;
        }

        /// <summary>
        /// Authenticates an active user by email-or-phone and password (compared as plain text, case-sensitive)
        /// and, on success, issues an HMAC-SHA256 signed JWT valid for 2 hours carrying the user's email, name,
        /// id and role, returned alongside basic profile fields.
        /// </summary>
        /// <param name="login">Credentials: <c>EmailOrPhone</c> (matched against either column) and <c>Password</c>; both are trimmed.</param>
        /// <returns>200 with the token and profile; otherwise an error message describing the failure.</returns>
        /// <response code="200">Credentials valid; returns token, role, name, email and loginId.</response>
        /// <response code="400">Body missing, email/phone or password blank, or the password did not match.</response>
        /// <response code="404">No active user matches the supplied email or phone.</response>
        /// <response code="500">JWT signing key is not configured, or an unexpected error occurred.</response>
        [HttpPost("check")]
        public async Task<IActionResult> LoginCheck([FromBody] LoginCheckModel login)
        {
            if (login == null)
            {
                return BadRequest(new { message = "Login data is required." });
            }

            if (string.IsNullOrWhiteSpace(login.EmailOrPhone) || string.IsNullOrWhiteSpace(login.Password))
            {
                return BadRequest(new { message = "Email/Phone and password are required." });
            }

            try
            {
                var emailOrPhone = login.EmailOrPhone.Trim();
                var password = login.Password.Trim();

                var user = await _loginDb.Login
                    .FirstOrDefaultAsync(u =>
                        (u.Email == emailOrPhone || u.Phone == emailOrPhone) &&
                        u.IsActive);

                if (user == null)
                {
                    return NotFound(new { message = "User not found." });
                }

                if (!string.Equals(user.Password, password, StringComparison.Ordinal))
                {
                    return BadRequest(new { message = "Incorrect password." });
                }

                var jwtKey = _configuration["Jwt:Key"];
                if (string.IsNullOrWhiteSpace(jwtKey))
                {
                    _logger.LogError("JWT key is missing from configuration.");
                    return StatusCode(500, new { message = "JWT configuration is missing." });
                }

                var roleValue = (user.Role ?? "").Trim();

                var claims = new List<Claim>
                {
                    new Claim(ClaimTypes.Email, user.Email ?? ""),
                    new Claim(ClaimTypes.Name, user.Name ?? ""),
                    new Claim(ClaimTypes.NameIdentifier, user.LoginId.ToString()),
                    new Claim(ClaimTypes.Role, roleValue),
                    new Claim("role", roleValue.ToLowerInvariant())
                };

                var key = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(jwtKey));

                var token = new JwtSecurityToken(
                    expires: DateTime.UtcNow.AddHours(2),
                    claims: claims,
                    signingCredentials: new SigningCredentials(key, SecurityAlgorithms.HmacSha256)
                );

                // Audit the successful login (the SaveChanges hook skips AuditLogs rows themselves).
                // Also store the logged-in person's record so the entry captures who signed in.
                _loginDb.AuditLogs.Add(new Elpis_CRM.Model.AuditLogModel
                {
                    EntityName = "Login",
                    EntityId = user.LoginId.ToString(),
                    Action = "Login",
                    ChangedBy = user.Email,
                    ChangedByName = user.Name,
                    ChangedByRole = roleValue,
                    ChangedAt = DateTime.UtcNow,
                    IpAddress = HttpContext.Connection.RemoteIpAddress?.ToString(),
                    Changes = System.Text.Json.JsonSerializer.Serialize(new
                    {
                        LoginId = user.LoginId,
                        Name = user.Name,
                        Email = user.Email,
                        Phone = user.Phone,
                        Role = roleValue
                    })
                });
                await _loginDb.SaveChangesAsync();

                return Ok(new
                {
                    token = new JwtSecurityTokenHandler().WriteToken(token),
                    role = roleValue,
                    name = user.Name,
                    email = user.Email,
                    loginId = user.LoginId
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Login check failed for {EmailOrPhone}", login.EmailOrPhone);
                return StatusCode(500, new
                {
                    message = "Server error during login check.",
                    detail = _environment.IsDevelopment()
                        ? ex.InnerException?.Message ?? ex.Message
                        : null
                });
            }
        }

        /// <summary>
        /// Lists active users that have a name (id, name and email only), de-duplicated, for populating
        /// dropdowns and autocomplete pickers. Requires a valid bearer token.
        /// </summary>
        /// <returns>200 with the slimmed-down user list (may be empty).</returns>
        /// <response code="200">User list returned.</response>
        /// <response code="401">No or invalid bearer token.</response>
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme)]
        [HttpGet("all-users")]
        public async Task<IActionResult> GetAllUsers()
        {
            var users = await _loginDb.Login
                .Where(u => u.IsActive && !string.IsNullOrEmpty(u.Name))
                .Select(u => new { name = u.Name, email = u.Email, loginId = u.LoginId })
                .Distinct()
                .ToListAsync();

            return Ok(users);
        }

        /// <summary>
        /// Creates a new login account, stamping created/updated timestamps and marking it active.
        /// Rejects the request if the email is already registered. Admin role required.
        /// </summary>
        /// <param name="login">Full account record; <c>Email</c> must be unique. The password is persisted as supplied.</param>
        /// <returns>200 with a confirmation message and the new loginId.</returns>
        /// <response code="200">Account created.</response>
        /// <response code="409">The email is already registered.</response>
        /// <response code="401">No or invalid bearer token.</response>
        /// <response code="403">Caller is authenticated but not an Admin.</response>
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin")]
        [HttpPost("create")]
        public async Task<IActionResult> Create([FromBody] LoginModel login)
        {
            bool exists = await _loginDb.Login.AnyAsync(u => u.Email == login.Email);
            if (exists)
            {
                return Conflict("Email is already registered.");
            }

            login.CreatedAt = DateTime.UtcNow;
            login.UpdatedAt = DateTime.UtcNow;
            login.IsActive = true;

            _loginDb.Login.Add(login);
            await _loginDb.SaveChangesAsync();

            return Ok(new
            {
                message = "Account created successfully.",
                loginId = login.LoginId
            });
        }

        /// <summary>
        /// Returns the profile of the caller identified by the <c>NameIdentifier</c> claim in their token,
        /// reloaded fresh from the database.
        /// </summary>
        /// <returns>200 with the current user's profile.</returns>
        /// <response code="200">Profile returned.</response>
        /// <response code="401">Token missing or has no usable user-id claim.</response>
        /// <response code="404">No account exists for the id in the token.</response>
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme)]
        [HttpGet("me")]
        public async Task<IActionResult> GetLoggedInUser()
        {
            var loginIdStr = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            if (!int.TryParse(loginIdStr, out var loginId))
                return Unauthorized(new { message = "User ID not found in token." });

            var user = await _loginDb.Login
                .Where(u => u.LoginId == loginId)
                .Select(u => new
                {
                    u.LoginId,
                    u.Name,
                    u.Email,
                    u.Phone,
                    u.Role,
                    u.IsActive,
                    u.CreatedAt,
                    u.UpdatedAt
                })
                .FirstOrDefaultAsync();

            if (user == null)
                return NotFound(new { message = "User not found." });

            return Ok(user);
        }

        /// <summary>
        /// Updates the caller's own name, email and/or phone (only non-blank fields are applied; phone may be
        /// cleared by sending an empty string). Blocks the change if the new email is already used by another account.
        /// </summary>
        /// <param name="model">Partial profile; blank/omitted name and email are ignored, a non-null phone is always written (trimmed).</param>
        /// <returns>200 with a confirmation and the updated profile.</returns>
        /// <response code="200">Profile updated.</response>
        /// <response code="401">Token missing or has no usable user-id claim.</response>
        /// <response code="404">No account exists for the id in the token.</response>
        /// <response code="409">The requested email already belongs to another account.</response>
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme)]
        [HttpPut("me")]
        public async Task<IActionResult> UpdateProfile([FromBody] UpdateProfileModel model)
        {
            var loginIdStr = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
            if (!int.TryParse(loginIdStr, out var loginId))
            {
                return Unauthorized(new { message = "User ID not found in token." });
            }
            var user = await _loginDb.Login.FindAsync(loginId);
            if (user == null)
            {
                return NotFound(new { message = "User not found." });
            }
            if (!string.IsNullOrWhiteSpace(model.Email) && model.Email != user.Email)
            {
                var emailTaken = await _loginDb.Login.AnyAsync(u => u.Email == model.Email && u.LoginId != loginId);
                if (emailTaken)
                {
                    return Conflict(new { message = "Email is already in use by another account." });
                }
            }

            if (!string.IsNullOrWhiteSpace(model.Name)) user.Name = model.Name.Trim();
            if (!string.IsNullOrWhiteSpace(model.Email)) user.Email = model.Email.Trim();
            if (model.Phone != null) user.Phone = model.Phone.Trim();
            user.UpdatedAt = DateTime.UtcNow;

            await _loginDb.SaveChangesAsync();

            return Ok(new
            {
                message = "Profile updated successfully.",
                user = new
                {
                    user.LoginId,
                    user.Name,
                    user.Email,
                    user.Phone,
                    user.Role,
                    user.IsActive,
                    user.CreatedAt,
                    user.UpdatedAt
                }
            });
        }

        /// <summary>
        /// Returns every login account, active or not, with profile and audit fields but without passwords. Admin role required.
        /// </summary>
        /// <returns>200 with the full account list.</returns>
        /// <response code="200">Accounts returned.</response>
        /// <response code="401">No or invalid bearer token.</response>
        /// <response code="403">Caller is authenticated but not an Admin.</response>
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin")]
        [HttpGet("all")]
        public async Task<IActionResult> GetAllAccounts()
        {
            var users = await _loginDb.Login
                .Select(u => new
                {
                    u.LoginId,
                    u.Name,
                    u.Email,
                    u.Phone,
                    u.Role,
                    u.IsActive,
                    u.CreatedAt,
                    u.UpdatedAt
                })
                .ToListAsync();

            return Ok(users);
        }

        /// <summary>
        /// Fetches a single account's profile and audit fields (no password) by its login id. Admin role required.
        /// </summary>
        /// <param name="id">Login id of the account to retrieve.</param>
        /// <returns>200 with the account, or 404 if no such id exists.</returns>
        /// <response code="200">Account found.</response>
        /// <response code="404">No account with the given id.</response>
        /// <response code="401">No or invalid bearer token.</response>
        /// <response code="403">Caller is authenticated but not an Admin.</response>
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin")]
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetById(int id)
        {
            var user = await _loginDb.Login
                .Where(u => u.LoginId == id)
                .Select(u => new
                {
                    u.LoginId,
                    u.Name,
                    u.Email,
                    u.Phone,
                    u.Role,
                    u.IsActive,
                    u.CreatedAt,
                    u.UpdatedAt
                })
                .FirstOrDefaultAsync();

            if (user == null)
            {
                return NotFound($"User with ID {id} not found.");
            }

            return Ok(user);
        }

        /// <summary>
        /// Overwrites an account's name, email, phone, role and active flag and refreshes its updated timestamp.
        /// The password is intentionally left untouched. Admin role required.
        /// </summary>
        /// <param name="id">Login id of the account to update.</param>
        /// <param name="loginModel">New field values; the supplied password is ignored.</param>
        /// <returns>200 with a confirmation and the updated fields, or 404 if the id is unknown.</returns>
        /// <response code="200">Account updated.</response>
        /// <response code="404">No account with the given id.</response>
        /// <response code="401">No or invalid bearer token.</response>
        /// <response code="403">Caller is authenticated but not an Admin.</response>
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin")]
        [HttpPut("{id:int}")]
        public async Task<IActionResult> Update(int id, [FromBody] LoginModel loginModel)
        {
            var existing = await _loginDb.Login.FindAsync(id);

            if (existing == null)
            {
                return NotFound($"User with ID {id} not found.");
            }

            existing.Name = loginModel.Name;
            existing.Email = loginModel.Email;
            existing.Phone = loginModel.Phone;
            existing.Role = loginModel.Role;
            existing.IsActive = loginModel.IsActive;
            existing.UpdatedAt = DateTime.UtcNow;

            // Only update password if provided
            //if (!string.IsNullOrEmpty(loginModel.Password))
            //{
            //    existing.Password = loginModel.Password;
            //}

            await _loginDb.SaveChangesAsync();

            return Ok(new
            {
                message = "User updated successfully.",
                user = new
                {
                    existing.LoginId,
                    existing.Name,
                    existing.Email,
                    existing.Phone,
                    existing.Role,
                    existing.IsActive
                }
            });
        }

        /// <summary>
        /// Permanently removes an account from the database (hard delete). Admin role required.
        /// </summary>
        /// <param name="id">Login id of the account to delete.</param>
        /// <returns>200 with a confirmation, or 404 if the id is unknown.</returns>
        /// <response code="200">Account deleted.</response>
        /// <response code="404">No account with the given id.</response>
        /// <response code="401">No or invalid bearer token.</response>
        /// <response code="403">Caller is authenticated but not an Admin.</response>
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin")]
        [HttpDelete("{id:int}")]
        public async Task<IActionResult> Delete(int id)
        {
            var user = await _loginDb.Login.FindAsync(id);

            if (user == null)
            {
                return NotFound($"User with ID {id} not found.");
            }

            _loginDb.Login.Remove(user);
            await _loginDb.SaveChangesAsync();

            return Ok(new { message = "User deleted successfully." });
        }

        /// <summary>
        /// Soft-deletes an account by clearing its active flag (preventing login) and bumping the updated timestamp,
        /// leaving the record in place. Admin role required.
        /// </summary>
        /// <param name="id">Login id of the account to deactivate.</param>
        /// <returns>200 with a confirmation, or 404 if the id is unknown.</returns>
        /// <response code="200">Account deactivated.</response>
        /// <response code="404">No account with the given id.</response>
        /// <response code="401">No or invalid bearer token.</response>
        /// <response code="403">Caller is authenticated but not an Admin.</response>
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin")]
        [HttpPut("{id:int}/deactivate")]
        public async Task<IActionResult> Deactivate(int id)
        {
            var user = await _loginDb.Login.FindAsync(id);

            if (user == null)
            {
                return NotFound($"User with ID {id} not found.");
            }

            user.IsActive = false;
            user.UpdatedAt = DateTime.UtcNow;
            await _loginDb.SaveChangesAsync();

            return Ok(new { message = "User deactivated successfully." });
        }

        /// <summary>
        /// Re-enables a previously deactivated account by setting its active flag and bumping the updated timestamp.
        /// Admin role required.
        /// </summary>
        /// <param name="id">Login id of the account to activate.</param>
        /// <returns>200 with a confirmation, or 404 if the id is unknown.</returns>
        /// <response code="200">Account activated.</response>
        /// <response code="404">No account with the given id.</response>
        /// <response code="401">No or invalid bearer token.</response>
        /// <response code="403">Caller is authenticated but not an Admin.</response>
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin")]
        [HttpPut("{id:int}/activate")]
        public async Task<IActionResult> Activate(int id)
        {
            var user = await _loginDb.Login.FindAsync(id);

            if (user == null)
            {
                return NotFound($"User with ID {id} not found.");
            }

            user.IsActive = true;
            user.UpdatedAt = DateTime.UtcNow;
            await _loginDb.SaveChangesAsync();

            return Ok(new { message = "User activated successfully." });
        }

        /// <summary>
        /// Sets a new password for the caller, located by the email claim in their token. Note: the new value is
        /// stored as-is and the model's current password is not verified before the change.
        /// </summary>
        /// <param name="model">Carries <c>NewPassword</c>, which replaces the stored password.</param>
        /// <returns>200 with a confirmation message.</returns>
        /// <response code="200">Password changed.</response>
        /// <response code="401">Token has no email claim.</response>
        /// <response code="404">No account matches the email in the token.</response>
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme)]
        [HttpPut("change-password")]
        public async Task<IActionResult> ChangePassword([FromBody] ChangePasswordModel model)
        {
            var email = User.FindFirst(ClaimTypes.Email)?.Value;

            if (string.IsNullOrEmpty(email))
            {
                return Unauthorized("User email not found in token.");
            }

            var user = await _loginDb.Login.FirstOrDefaultAsync(u => u.Email == email);

            if (user == null)
            {
                return NotFound("User not found.");
            }

            user.Password = model.NewPassword;
            user.UpdatedAt = DateTime.UtcNow;
            await _loginDb.SaveChangesAsync();

            return Ok(new { message = "Password changed successfully." });
        }
    }
}
