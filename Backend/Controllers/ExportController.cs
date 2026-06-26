using Elpis_CRM.DTOs;
using Elpis_CRM.Services;
using Microsoft.AspNetCore.Mvc;

namespace Elpis_CRM.Controllers
{
    /// <summary>
    /// Generates downloadable Excel (.xlsx) exports of CRM accounts and contacts,
    /// honoring optional search, tag, and column-selection filters.
    /// </summary>
    [ApiController]
    [Route("api/export")]
    public class ExportController : ControllerBase
    {
        private readonly ExportService _exportService;
        private readonly ILogger<ExportController> _logger;

        private const string ExcelContentType =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        public ExportController(ExportService exportService, ILogger<ExportController> logger)
        {
            _exportService = exportService;
            _logger = logger;
        }

        // ─── POST /api/export/accounts ──────────────────────────────────────────

        /// <summary>
        /// Builds an Excel workbook of accounts and returns it as a timestamped
        /// <c>accounts_yyyyMMdd_HHmmss.xlsx</c> download. The optional <c>Search</c>
        /// term matches name/city/country/phone/website, <c>Tag</c> filters by tag,
        /// and <c>Columns</c> restricts which fields become worksheet columns.
        /// </summary>
        /// <param name="request">Export filters: search text, tag, and the column subset to include.</param>
        /// <returns>The generated .xlsx file, or a 500 error payload if export fails.</returns>
        /// <response code="200">Returns the accounts workbook as a file download.</response>
        /// <response code="500">An unexpected error occurred while building the export.</response>
        [HttpPost("accounts")]
        [ProducesResponseType(typeof(FileContentResult), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        public async Task<IActionResult> ExportAccounts([FromBody] ExportRequest request)
        {
            try
            {
                byte[] fileBytes = await _exportService.ExportAccountsAsync(
                    request.Search,
                    request.Tag,
                    request.Columns);

                string fileName = $"accounts_{DateTime.UtcNow:yyyyMMdd_HHmmss}.xlsx";
                return File(fileBytes, ExcelContentType, fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error exporting accounts");
                return StatusCode(StatusCodes.Status500InternalServerError,
                    new { message = "An error occurred while exporting accounts." });
            }
        }

        // ─── POST /api/export/contacts ──────────────────────────────────────────

        /// <summary>
        /// Builds an Excel workbook of contacts and returns it as a timestamped
        /// <c>contacts_yyyyMMdd_HHmmss.xlsx</c> download. The optional <c>Search</c>
        /// term matches first/last name, work email, work phone, city, and country,
        /// and <c>Columns</c> restricts which fields become worksheet columns. The
        /// request's <c>Tag</c> value is ignored here, as contacts are not tagged.
        /// </summary>
        /// <param name="request">Export filters: search text and the column subset to include (tag is not used).</param>
        /// <returns>The generated .xlsx file, or a 500 error payload if export fails.</returns>
        /// <response code="200">Returns the contacts workbook as a file download.</response>
        /// <response code="500">An unexpected error occurred while building the export.</response>
        [HttpPost("contacts")]
        [ProducesResponseType(typeof(FileContentResult), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        public async Task<IActionResult> ExportContacts([FromBody] ExportRequest request)
        {
            try
            {
                byte[] fileBytes = await _exportService.ExportContactsAsync(
                    request.Search,
                    request.Columns);

                string fileName = $"contacts_{DateTime.UtcNow:yyyyMMdd_HHmmss}.xlsx";
                return File(fileBytes, ExcelContentType, fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error exporting contacts");
                return StatusCode(StatusCodes.Status500InternalServerError,
                    new { message = "An error occurred while exporting contacts." });
            }
        }
    }
}
