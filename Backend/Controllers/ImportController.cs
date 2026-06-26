using Elpis_CRM.Services;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace Elpis_CRM.Controllers;

/// <summary>
/// Admin-only bulk import endpoints that accept multipart file uploads (CSV/Excel) and stream
/// them into CRM tables via <see cref="ImportService"/>. Uploads are capped at 50 MB and every
/// route requires the JWT "Admin" role except the anonymous <see cref="Status"/> health probe.
/// </summary>
[ApiController]
[Route("api/import")]
[Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme, Roles = "Admin")]
[RequestSizeLimit(52_428_800)]               // 50 MB hard cap — sync with ImportOptions if changed
[RequestFormLimits(MultipartBodyLengthLimit = 52_428_800)]
public sealed class ImportController(
    ImportService importService,
    DealsService dealsService,
    ILogger<ImportController> logger) : ControllerBase
{

    /// <summary>Imports an uploaded CSV/Excel file into the Accounts table.</summary>
    /// <param name="file">The multipart-uploaded spreadsheet; headers are mapped to Account columns.</param>
    /// <param name="ct">Cancellation token tied to the request lifetime.</param>
    /// <returns>200 with import counts on success, 400 when the file is missing/invalid or the bulk insert fails.</returns>
    /// <response code="200">Import completed; body reports rows imported, skipped, elapsed time and any row errors.</response>
    /// <response code="400">File was empty/missing or the import failed.</response>
    [HttpPost("accounts"), Consumes("multipart/form-data")]
    public Task<IActionResult> Accounts(IFormFile file, CancellationToken ct)
        => RunImport(file, "Accounts", ct);

    /// <summary>Imports an uploaded CSV/Excel file into the Contacts table.</summary>
    /// <param name="file">The multipart-uploaded spreadsheet; headers are mapped to Contact columns.</param>
    /// <param name="ct">Cancellation token tied to the request lifetime.</param>
    /// <returns>200 with import counts on success, 400 when the file is missing/invalid or the bulk insert fails.</returns>
    /// <response code="200">Import completed; body reports rows imported, skipped, elapsed time and any row errors.</response>
    /// <response code="400">File was empty/missing or the import failed.</response>
    [HttpPost("contacts"), Consumes("multipart/form-data")]
    public Task<IActionResult> Contacts(IFormFile file, CancellationToken ct)
        => RunImport(file, "Contacts", ct);

    /// <summary>Imports an uploaded CSV/Excel file into the Deals table.</summary>
    /// <param name="file">The multipart-uploaded spreadsheet; headers are mapped to Deal columns.</param>
    /// <param name="ct">Cancellation token tied to the request lifetime.</param>
    /// <returns>200 with import counts on success, 400 when the file is missing/invalid or the bulk insert fails.</returns>
    /// <response code="200">Import completed; body reports rows imported, skipped, elapsed time and any row errors.</response>
    /// <response code="400">File was empty/missing or the import failed.</response>
    [HttpPost("deals"), Consumes("multipart/form-data")]
    public Task<IActionResult> Deals(IFormFile file, CancellationToken ct)
        => RunImport(file, "Deals", ct);

    /// <summary>
    /// Links existing deals to contacts from a two-column file (DealId + ContactId), resolving the
    /// contact display name from the Contacts table. Delegates to <see cref="DealsService.ImportContactLinksAsync"/>
    /// and reports rows updated/skipped plus any per-row errors.
    /// </summary>
    /// <param name="file">The multipart-uploaded CSV/Excel file containing DealId and ContactId columns.</param>
    /// <param name="ct">Cancellation token tied to the request lifetime.</param>
    /// <returns>200 with update counts on success, 400 when the file is empty/missing or the link operation reports failure.</returns>
    /// <response code="200">Links applied; body reports rows updated, skipped, elapsed time and any row errors.</response>
    /// <response code="400">File was empty/missing or the link operation failed.</response>
    [HttpPost("deals/link-contacts"), Consumes("multipart/form-data")]
    public async Task<IActionResult> LinkDealContacts(IFormFile file, CancellationToken ct)
    {
        if (file is null || file.Length == 0)
            return BadRequest(new { error = "No file provided or file is empty." });

        logger.LogInformation(
            "Deal contact link import | {File} | {Size:N0} bytes | User: {User}",
            file.FileName, file.Length, User.Identity?.Name ?? "unknown");

        var result = await dealsService.ImportContactLinksAsync(file, ct);

        if (result.Success)
        {
            return Ok(new
            {
                success = true,
                table = "DealContactLinks",
                rowsImported = result.RowsUpdated,
                rowsSkipped = result.RowsSkipped,
                elapsedMs = (long)result.Elapsed.TotalMilliseconds,
                rowErrors = result.RowErrors,
            });
        }

        return BadRequest(new
        {
            success = false,
            table = "DealContactLinks",
            error = result.Error,
            rowsImported = result.RowsUpdated,
            rowsSkipped = result.RowsSkipped,
            rowErrors = result.RowErrors,
        });
    }

    /// <summary>Imports an uploaded CSV/Excel file into the CallLog table.</summary>
    /// <param name="file">The multipart-uploaded spreadsheet; headers are mapped to CallLog columns.</param>
    /// <param name="ct">Cancellation token tied to the request lifetime.</param>
    /// <returns>200 with import counts on success, 400 when the file is missing/invalid or the bulk insert fails.</returns>
    /// <response code="200">Import completed; body reports rows imported, skipped, elapsed time and any row errors.</response>
    /// <response code="400">File was empty/missing or the import failed.</response>
    [HttpPost("calllogs"), Consumes("multipart/form-data")]
    public Task<IActionResult> CallLogs(IFormFile file, CancellationToken ct)
        => RunImport(file, "CallLogs", ct);

    /// <summary>Imports an uploaded CSV/Excel file into the Notes table.</summary>
    /// <param name="file">The multipart-uploaded spreadsheet; headers are mapped to Note columns.</param>
    /// <param name="ct">Cancellation token tied to the request lifetime.</param>
    /// <returns>200 with import counts on success, 400 when the file is missing/invalid or the bulk insert fails.</returns>
    /// <response code="200">Import completed; body reports rows imported, skipped, elapsed time and any row errors.</response>
    /// <response code="400">File was empty/missing or the import failed.</response>
[HttpPost("notes"), Consumes("multipart/form-data")]
    public Task<IActionResult> Notes(IFormFile file, CancellationToken ct)
        => RunImport(file, "Notes", ct);

    /// <summary>Imports an uploaded CSV/Excel file into the Tasks table.</summary>
    /// <param name="file">The multipart-uploaded spreadsheet; headers are mapped to Task columns.</param>
    /// <param name="ct">Cancellation token tied to the request lifetime.</param>
    /// <returns>200 with import counts on success, 400 when the file is missing/invalid or the bulk insert fails.</returns>
    /// <response code="200">Import completed; body reports rows imported, skipped, elapsed time and any row errors.</response>
    /// <response code="400">File was empty/missing or the import failed.</response>
    [HttpPost("tasks"), Consumes("multipart/form-data")]
    public Task<IActionResult> Tasks(IFormFile file, CancellationToken ct)
        => RunImport(file, "Tasks", ct);

    // Add more tables by:
    //   1. Adding an ImportTableSchema entry to SchemaRegistry.All in ImportService.cs
    //   2. Adding an endpoint here that calls RunImport(file, "TableName", ct)
    //
    // Example:
    // [HttpPost("products"), Consumes("multipart/form-data")]
    // public Task<IActionResult> Products(IFormFile file, CancellationToken ct)
    //     => RunImport(file, "Products", ct);

    // ── Status (public — useful for health probes / front-end upload hints) ───

    /// <summary>
    /// Anonymous health/capability probe describing the import service: the tables that can be
    /// imported, the 50 MB size cap, the accepted file extensions and the default batch size.
    /// Useful for front-end upload hints and uptime checks.
    /// </summary>
    /// <returns>200 with a static descriptor of import capabilities.</returns>
    /// <response code="200">Service is running; body lists allowed tables, max file size, supported formats and batch size.</response>
    [HttpGet("status"), AllowAnonymous]
    public IActionResult Status() => Ok(new
    {
        status = "running",
        allowedTables = new[] { "Accounts", "Contacts", "Deals", "DealContactLinks", "CallLogs", "Notes", "Tasks" },
        maxFileSizeMb = 50,
        supportedFormats = new[] { ".xlsx", ".xlsb", ".xls", ".csv" },
        defaultBatchSize = 5_000,
    });


    /// <summary>
    /// Shared pipeline behind every table endpoint: rejects empty uploads, logs the request, builds
    /// fixed <see cref="ImportOptions"/> (5 000-row batches, no bad-row skipping, 3 retries, 5 MB
    /// index-disable threshold), runs the import and shapes the success/failure JSON response.
    /// </summary>
    /// <param name="file">The uploaded file to import.</param>
    /// <param name="table">Target table key matching a registered import schema (e.g. "Accounts").</param>
    /// <param name="ct">Cancellation token tied to the request lifetime.</param>
    /// <returns>Ok with row counts when the import succeeds; otherwise BadRequest with the error and any row errors.</returns>
    private async Task<IActionResult> RunImport(
        IFormFile file, string table, CancellationToken ct)
    {
        if (file is null || file.Length == 0)
            return BadRequest(new { error = "No file provided or file is empty." });

        logger.LogInformation(
            "Import request → [{Table}] | {File} | {Size:N0} bytes | User: {User}",
            table, file.FileName, file.Length, User.Identity?.Name ?? "unknown");

        // Customise per-request if needed (e.g. read batch size from a header / query param).
        var options = new ImportOptions
        {
            BatchSize = 5_000,
            SkipBadRows = false,
            MaxRetryAttempts = 3,
            IndexDisableThresholdBytes = 5 * 1_024 * 1_024, // 5 MB
        };

        var result = await importService.ImportAsync(file, table, options, ct);

        if (result.Success)
        {
            logger.LogInformation(
                "Import succeeded → [{Table}] | {Rows:N0} rows in {Elapsed:0.00}s",
                table, result.RowsImported, result.Elapsed.TotalSeconds);

            return Ok(new
            {
                success = true,
                table,
                rowsImported = result.RowsImported,
                rowsSkipped = result.RowsSkipped,
                elapsedMs = (long)result.Elapsed.TotalMilliseconds,
                rowErrors = result.RowErrors,
            });
        }

        logger.LogWarning(
            "Import failed → [{Table}] | {Error}", table, result.Error);

        return BadRequest(new
        {
            success = false,
            table,
            error = result.Error,
            rowErrors = result.RowErrors,
        });
    }
}