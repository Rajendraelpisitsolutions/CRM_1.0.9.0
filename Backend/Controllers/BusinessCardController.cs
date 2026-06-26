using Elpis_CRM.Model;
using Elpis_CRM.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authentication.JwtBearer;

namespace Elpis_CRM.Controllers
{
    /// <summary>
    /// Manages business card scanning operations using Azure Document Intelligence.
    /// Extracts structured contact information from business card images.
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    public class BusinessCardController : ControllerBase
    {
        private readonly AzureBusinessCardService? _businessCardService;
        private readonly AccountService _accountService;

        /// <summary>
        /// Initializes a new instance of the <see cref="BusinessCardController"/>.
        /// </summary>
        /// <param name="businessCardService">Azure Document Intelligence scanning service; null when the integration is not configured, in which case scan endpoints return 501.</param>
        /// <param name="accountService">Account service used to persist accounts created from a scanned card.</param>
        public BusinessCardController(
            AzureBusinessCardService? businessCardService,
            AccountService accountService)
        {
            _businessCardService = businessCardService;
            _accountService = accountService;
        }

        /// <summary>
        /// Runs the uploaded card image through Azure Document Intelligence and returns the extracted
        /// account fields (company name, website, phone, address) without persisting anything.
        /// </summary>
        /// <param name="file">Business card image (JPEG, PNG, TIFF, BMP, or WebP); must be non-empty and at most 5 MB.</param>
        /// <param name="cancellationToken">Token used to cancel the Azure analysis call.</param>
        /// <returns>An account populated from the card; fields the model could not detect are left empty.</returns>
        /// <response code="200">Business card scanned successfully</response>
        /// <response code="400">File is missing, empty, too large, or has an unsupported format</response>
        /// <response code="422">Azure analysis failed or could not extract document</response>
        /// <response code="501">Azure Document Intelligence service not configured</response>
        /// <response code="500">Server error or Azure service unavailable</response>
        [HttpPost("scan")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager,User")]
        [RequestSizeLimit(5 * 1024 * 1024)]
        [RequestFormLimits(MultipartBodyLengthLimit = 5 * 1024 * 1024)]
        [ProducesResponseType(typeof(AccountModel), StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(ProblemDetails), StatusCodes.Status400BadRequest)]
        [ProducesResponseType(typeof(ProblemDetails), StatusCodes.Status422UnprocessableEntity)]
        [ProducesResponseType(typeof(ProblemDetails), StatusCodes.Status501NotImplemented)]
        [ProducesResponseType(typeof(ProblemDetails), StatusCodes.Status500InternalServerError)]
        public async Task<ActionResult<AccountModel>> Scan(
            IFormFile? file,
            CancellationToken cancellationToken)
        {
            if (_businessCardService is null)
                return StatusCode(StatusCodes.Status501NotImplemented, 
                    new ProblemDetails
                    {
                        Status = StatusCodes.Status501NotImplemented,
                        Title = "Business Card Service Not Configured",
                        Detail = "Azure Document Intelligence is not configured. Please add credentials to appsettings.json.",
                        Instance = HttpContext.Request.Path
                    });

            if (file is null || file.Length == 0)
                return BadRequest(new ProblemDetails
                {
                    Status = StatusCodes.Status400BadRequest,
                    Title = "No file provided",
                    Detail = "Please attach a business card image file using the 'file' form field.",
                    Instance = HttpContext.Request.Path
                });

            try
            {
                var scannedAccount = await _businessCardService.ScanAsync(file, cancellationToken);
                return Ok(scannedAccount);
            }
            catch (ArgumentException ex)
            {
                return BadRequest(new ProblemDetails
                {
                    Status = StatusCodes.Status400BadRequest,
                    Title = "Invalid file",
                    Detail = ex.Message,
                    Instance = HttpContext.Request.Path
                });
            }
            catch (InvalidOperationException ex)
            {
                return StatusCode(StatusCodes.Status422UnprocessableEntity, 
                    new ProblemDetails
                    {
                        Status = StatusCodes.Status422UnprocessableEntity,
                        Title = "Processing Error",
                        Detail = ex.Message,
                        Instance = HttpContext.Request.Path
                    });
            }
        }

        /// <summary>
        /// Scans the card via Azure Document Intelligence and persists the extracted data as a new account.
        /// Restricted to Admins and Managers. The Location header points at the Scan action.
        /// </summary>
        /// <param name="file">Business card image (JPEG, PNG, TIFF, BMP, or WebP); must be non-empty and at most 5 MB.</param>
        /// <param name="cancellationToken">Token used to cancel the Azure analysis call.</param>
        /// <returns>The persisted account, including its database-assigned id.</returns>
        /// <response code="201">Account created successfully from business card</response>
        /// <response code="400">File is missing, empty, too large, or has an unsupported format</response>
        /// <response code="422">Azure analysis failed or could not extract document</response>
        /// <response code="501">Azure Document Intelligence service not configured</response>
        /// <response code="500">Server error or Azure service unavailable</response>
        [HttpPost("scan-and-create")]
        [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin,Manager")]
        [RequestSizeLimit(5 * 1024 * 1024)]
        [RequestFormLimits(MultipartBodyLengthLimit = 5 * 1024 * 1024)]
        [ProducesResponseType(typeof(AccountModel), StatusCodes.Status201Created)]
        [ProducesResponseType(typeof(ProblemDetails), StatusCodes.Status400BadRequest)]
        [ProducesResponseType(typeof(ProblemDetails), StatusCodes.Status422UnprocessableEntity)]
        [ProducesResponseType(typeof(ProblemDetails), StatusCodes.Status501NotImplemented)]
        [ProducesResponseType(typeof(ProblemDetails), StatusCodes.Status500InternalServerError)]
        public async Task<ActionResult<AccountModel>> ScanAndCreate(
            IFormFile? file,
            CancellationToken cancellationToken)
        {
            if (_businessCardService is null)
                return StatusCode(StatusCodes.Status501NotImplemented, 
                    new ProblemDetails
                    {
                        Status = StatusCodes.Status501NotImplemented,
                        Title = "Business Card Service Not Configured",
                        Detail = "Azure Document Intelligence is not configured. Please add credentials to appsettings.json.",
                        Instance = HttpContext.Request.Path
                    });

            if (file is null || file.Length == 0)
                return BadRequest(new ProblemDetails
                {
                    Status = StatusCodes.Status400BadRequest,
                    Title = "No file provided",
                    Detail = "Please attach a business card image file using the 'file' form field.",
                    Instance = HttpContext.Request.Path
                });

            try
            {
                var scannedAccount = await _businessCardService.ScanAsync(file, cancellationToken);
                var created = await _accountService.AddAsync(scannedAccount);
                return CreatedAtAction(nameof(Scan), new { id = created.AccountId }, created);
            }
            catch (ArgumentException ex)
            {
                return BadRequest(new ProblemDetails
                {
                    Status = StatusCodes.Status400BadRequest,
                    Title = "Invalid file",
                    Detail = ex.Message,
                    Instance = HttpContext.Request.Path
                });
            }
            catch (InvalidOperationException ex)
            {
                return StatusCode(StatusCodes.Status422UnprocessableEntity, 
                    new ProblemDetails
                    {
                        Status = StatusCodes.Status422UnprocessableEntity,
                        Title = "Processing Error",
                        Detail = ex.Message,
                        Instance = HttpContext.Request.Path
                    });
            }
        }
    }
}
