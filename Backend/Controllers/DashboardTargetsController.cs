using Elpis_CRM.Dtos;
using Microsoft.AspNetCore.Mvc;
using System.Text.Json;

namespace Elpis_CRM.Controllers
{
    /// <summary>
    /// Reads and updates the quarterly dashboard sales targets, which are persisted in a single JSON file
    /// (Data/dashboard-targets.json) under the content root rather than in the database. A process-wide
    /// semaphore serializes file access so concurrent reads/writes don't corrupt the file.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class DashboardTargetsController : ControllerBase
    {
        private static readonly SemaphoreSlim _fileLock = new SemaphoreSlim(1, 1);
        private readonly string _filePath;

        /// <summary>
        /// Resolves the targets file path under the application's content root (Data/dashboard-targets.json).
        /// </summary>
        /// <param name="env">Hosting environment used to locate the content root.</param>
        public DashboardTargetsController(IWebHostEnvironment env)
        {
            // ContentRootPath is always correct on any server
            _filePath = Path.Combine(
                env.ContentRootPath,
                "Data",
                "dashboard-targets.json"
            );
        }

        /// <summary>
        /// Returns the raw targets JSON file contents verbatim (as application/json), taking the file lock
        /// for the duration of the read.
        /// </summary>
        /// <returns>The file's JSON content.</returns>
        /// <response code="200">The current dashboard targets as JSON.</response>
        /// <response code="404">The targets file does not exist.</response>
        [HttpGet("dashboard-targets")]
        public async Task<IActionResult> GetTargets()
        {
            await _fileLock.WaitAsync();

            try
            {
                if (!System.IO.File.Exists(_filePath))
                    return NotFound("Targets file not found.");

                var json = await System.IO.File.ReadAllTextAsync(_filePath);

                return Content(json, "application/json");
            }
            finally
            {
                _fileLock.Release();
            }
        }

        /// <summary>
        /// Sets the target for a single quarter and rewrites the JSON file, preserving the other quarters'
        /// values. The whole read-modify-write runs under the file lock; the quarter name is matched
        /// case-insensitively against q1-q4.
        /// </summary>
        /// <param name="quarter">Which quarter to update: q1, q2, q3 or q4 (case-insensitive).</param>
        /// <param name="request">Body carrying the new <c>Target</c> value for that quarter.</param>
        /// <returns>The full set of targets after the update.</returns>
        /// <response code="200">The quarter was updated; returns all four targets.</response>
        /// <response code="400">The quarter segment was not one of q1-q4.</response>
        /// <response code="404">The targets file does not exist.</response>
        [HttpPut("dashboard-targets/{quarter}")]
        public async Task<IActionResult> UpdateQuarter(
            string quarter,
            [FromBody] QuarterTargetDto request)
        {
            await _fileLock.WaitAsync();

            try
            {
                if (!System.IO.File.Exists(_filePath))
                    return NotFound("Targets file not found.");

                var json = await System.IO.File.ReadAllTextAsync(_filePath);

                var options = new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true,
                    WriteIndented = true
                };

                var targets = JsonSerializer.Deserialize<DashboardTargetsDto>(
                    json,
                    options
                ) ?? new DashboardTargetsDto();

                switch (quarter.ToLower())
                {
                    case "q1":
                        targets.Q1 = request.Target;
                        break;

                    case "q2":
                        targets.Q2 = request.Target;
                        break;

                    case "q3":
                        targets.Q3 = request.Target;
                        break;

                    case "q4":
                        targets.Q4 = request.Target;
                        break;

                    default:
                        return BadRequest("Invalid quarter. Use q1–q4.");
                }

                var updatedJson = JsonSerializer.Serialize(targets, options);

                await System.IO.File.WriteAllTextAsync(_filePath, updatedJson);

                return Ok(targets);
            }
            finally
            {
                _fileLock.Release();
            }
        }
    }
}