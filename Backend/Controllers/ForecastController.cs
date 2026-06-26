using Elpis_CRM.Services;
using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;

namespace Elpis_CRM.Controllers
{
    /// <summary>
    /// Exposes sales forecast figures (deal counts per period), per-period deal name lists,
    /// and a six-month actual-vs-pipeline trend, scoped by month, quarter, or year.
    /// </summary>
    [ApiController]
    [Route("api/forecast")]
    public class ForecastController : ControllerBase
    {
        private readonly ForecastService _forecastService;

        /// <summary>
        /// Initializes the controller with the forecast service that performs the underlying calculations.
        /// </summary>
        /// <param name="service">Service that computes forecast totals, deal names, and trend data.</param>
        public ForecastController(ForecastService service)
        {
            _forecastService = service;
        }

        /// <summary>
        /// Returns the forecast for one calendar month: a period label and the count of deals created within it.
        /// </summary>
        /// <param name="month">Full month name (e.g., "January"); parsed using the invariant culture.</param>
        /// <param name="year">Four-digit year.</param>
        /// <returns>An object with the period label and total deal count for that month.</returns>
        /// <response code="200">Monthly forecast retrieved successfully.</response>
            //  AFTER
        [HttpGet("month/{month}/{year}")]
        public async Task<IActionResult> Monthly(string month, int year)
            => Ok(await _forecastService.MonthlyAsync(month, year));


        /// <summary>
        /// Returns the forecast for one quarter: a month-range label and the count of deals created across its three months.
        /// </summary>
        /// <param name="quarter">Quarter token ("Q1"–"Q4", case-insensitive); other values cause a service-level error.</param>
        /// <param name="year">Four-digit year.</param>
        /// <returns>An object with the quarter's period label and total deal count.</returns>
        /// <response code="200">Quarterly forecast retrieved successfully.</response>

        //  AFTER
        [HttpGet("quarter/{quarter}/{year}")]
        public async Task<IActionResult> Quarterly(string quarter, int year)
            => Ok(await _forecastService.QuarterlyAsync(quarter, year));

        /// <summary>
        /// Returns the forecast for a full calendar year: the year label and the count of deals created within it.
        /// </summary>
        /// <param name="year">Four-digit year.</param>
        /// <returns>An object with the year label and total deal count.</returns>
        /// <response code="200">Yearly forecast retrieved successfully.</response>

        // AFTER
        [HttpGet("year/{year}")]
        public async Task<IActionResult> Yearly(int year)
            => Ok(await _forecastService.YearlyAsync(year));

        /// <summary>
        /// Lists the distinct names of deals created during the given month.
        /// </summary>
        /// <param name="monthName">Full month name (e.g., "January"), supplied as a query parameter.</param>
        /// <param name="year">Four-digit year, supplied as a query parameter.</param>
        /// <returns>The distinct deal names for that month (names may be null).</returns>
        /// <response code="200">Monthly deal names retrieved successfully.</response>
        [HttpGet("monthly/deals")]
        public async Task<IActionResult> GetMonthlyDeals([FromQuery] string monthName, [FromQuery] int year)
        {
            var deals = await _forecastService.MonthlyDealNamesAsync(monthName, year);
            return Ok(deals);
        }

        /// <summary>
        /// Lists the distinct names of deals created during the given quarter.
        /// </summary>
        /// <param name="quarter">Quarter token ("Q1"–"Q4", case-insensitive), supplied as a query parameter.</param>
        /// <param name="year">Four-digit year, supplied as a query parameter.</param>
        /// <returns>The distinct deal names for that quarter (names may be null).</returns>
        /// <response code="200">Quarterly deal names retrieved successfully.</response>
        [HttpGet("quarterly/deals")]
        public async Task<IActionResult> GetQuarterlyDeals([FromQuery] string quarter, [FromQuery] int year)
        {
            var deals = await _forecastService.QuarterlyDealNamesAsync(quarter, year);
            return Ok(deals);
        }

        /// <summary>
        /// Lists the distinct names of deals created during the given year.
        /// </summary>
        /// <param name="year">Four-digit year, supplied as a query parameter.</param>
        /// <returns>The distinct deal names for that year (names may be null).</returns>
        /// <response code="200">Yearly deal names retrieved successfully.</response>
        [HttpGet("yearly/deals")]
        public async Task<IActionResult> GetYearlyDeals([FromQuery] int year)
        {
            var deals = await _forecastService.YearlyDealNamesAsync(year);
            return Ok(deals);
        }

        /// <summary>
        /// Returns the last six months of trend data, each month carrying actual revenue from
        /// closed/paid deals alongside the probability-weighted value of still-open pipeline.
        /// </summary>
        /// <returns>A per-month series with the month label, actual revenue, and expected pipeline figures.</returns>
        /// <response code="200">Trend data retrieved successfully.</response>
        [HttpGet("sixmonthtrend")]
        public async Task<IActionResult> SixMonthTrend()
        {
            var trendData = await _forecastService.SixMonthTrendAsync();
            return Ok(trendData);
        }
    }
}
