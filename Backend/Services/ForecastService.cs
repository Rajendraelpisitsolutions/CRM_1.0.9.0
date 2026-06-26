using Elpis_CRM.Data;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

namespace Elpis_CRM.Services
{
    /// <summary>
    /// Service to generate sales forecasts and deal summaries for monthly, quarterly, and yearly periods.
    /// </summary>
    public class ForecastService
    {
        private readonly AppDbContext _context;

        /// <summary>
        /// Initializes a new instance of <see cref="ForecastService"/>.
        /// </summary>
        /// <param name="context">The database context.</param>
        public ForecastService(AppDbContext context)
        {
            _context = context;
        }

        #region Totals

        /// <summary>
        /// Counts deals created within the given month (by <c>CreatedAt</c>) and labels the period with the localized month name and year.
        /// </summary>
        /// <param name="monthName">Full month name (e.g., "January"), parsed via the invariant culture.</param>
        /// <param name="year">Four-digit year (e.g., 2026).</param>
        /// <returns>An anonymous object exposing <c>period</c> (label) and <c>totalDeals</c> (count).</returns>
        public async Task<object> MonthlyAsync(string monthName, int year)
        {
            int month = ParseMonthName(monthName);
            DateTime start = new DateTime(year, month, 1, 0, 0, 0, DateTimeKind.Utc);
            DateTime end = start.AddMonths(1).AddDays(-1);

            return await BuildResultAsync(start, end, $"{CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month)} {year}");
        }

        /// <summary>
        /// Counts deals created across the quarter's three months (by <c>CreatedAt</c>) and labels the period with the start–end month range.
        /// </summary>
        /// <param name="quarter">Quarter token "Q1"–"Q4" (case-insensitive).</param>
        /// <param name="year">Four-digit year.</param>
        /// <returns>An anonymous object exposing <c>period</c> (label) and <c>totalDeals</c> (count).</returns>
        /// <exception cref="ArgumentException">Thrown when <paramref name="quarter"/> is not Q1–Q4.</exception>
        public async Task<object> QuarterlyAsync(string quarter, int year)
        {
            int startMonth = quarter.ToUpper() switch
            {
                "Q1" => 1,
                "Q2" => 4,
                "Q3" => 7,
                "Q4" => 10,
                _ => throw new ArgumentException("Invalid quarter")
            };

            DateTime start = new DateTime(year, startMonth, 1, 0, 0, 0, DateTimeKind.Utc);
            DateTime end = start.AddMonths(3).AddDays(-1);

            string label = $"{start:MMM}–{end:MMM} {year}";
            return await BuildResultAsync(start, end, label);
        }

        /// <summary>
        /// Counts deals created during the full calendar year (by <c>CreatedAt</c>), labeling the period with the year.
        /// </summary>
        /// <param name="year">Four-digit year.</param>
        /// <returns>An anonymous object exposing <c>period</c> (label) and <c>totalDeals</c> (count).</returns>
        public async Task<object> YearlyAsync(int year)
        {
            DateTime start = new DateTime(year, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            DateTime end = new DateTime(year, 12, 31, 23, 59, 59, DateTimeKind.Utc);

            return await BuildResultAsync(start, end, year.ToString());
        }

        /// <summary>
        /// Counts deals whose <c>CreatedAt</c> falls within the inclusive [start, end] range and pairs the count with the supplied label.
        /// </summary>
        /// <param name="start">Inclusive lower bound of the creation date range.</param>
        /// <param name="end">Inclusive upper bound of the creation date range.</param>
        /// <param name="label">Human-readable period label returned as <c>period</c>.</param>
        /// <returns>An anonymous object exposing <c>period</c> and <c>totalDeals</c>.</returns>
        private async Task<object> BuildResultAsync(DateTime start, DateTime end, string label)
        {
            int totalDeals = await _context.Deals
                .CountAsync(d => d.CreatedAt >= start && d.CreatedAt <= end);

            return new
            {
                period = label,
                totalDeals
            };
        }

        #endregion

        #region Six Month Trend

        /// <summary>
        /// Builds a six-month trend grouped by month. For each month, actual revenue sums the
        /// <c>DealValue</c> of deals that have a close date and a "Paid"/"Won" payment status,
        /// while expected pipeline sums open deals' <c>DealValue × (Probability / 100)</c>.
        /// Deals are bucketed by their closed month when set, otherwise their expected-close month.
        /// </summary>
        /// <returns>An ordered list of per-month objects, each with <c>month</c>, <c>actualRevenue</c>, and <c>expectedPipeline</c>.</returns>
        public async Task<List<object>> SixMonthTrendAsync()
        {
            var endDate = DateTime.UtcNow.Date;
            var startDate = endDate.AddMonths(-6).Date;

            // Load all relevant deals first to avoid EF complex query issues
            var deals = await _context.Deals
                .Where(d => (d.ClosedDate.HasValue && d.ClosedDate.Value.Date >= startDate) ||
                           (d.ExpectedCloseDate.HasValue && d.ExpectedCloseDate.Value.Date >= startDate))
                .ToListAsync();

            var trendData = deals
                .GroupBy(d => new {
                    Year = d.ClosedDate?.Year ?? d.ExpectedCloseDate!.Value.Year,
                    Month = d.ClosedDate?.Month ?? d.ExpectedCloseDate!.Value.Month
                })
                .Select(g => {
                    var actualRevenue = g.Where(d => d.ClosedDate.HasValue &&
                                                    (d.PaymentStatus == "Paid" || d.PaymentStatus == "Won") &&
                                                    d.DealValue.HasValue)
                                        .Sum(d => (double)(d.DealValue!.Value));

                    var expectedPipeline = g.Where(d => !d.ClosedDate.HasValue &&
                                                       d.ExpectedCloseDate.HasValue &&
                                                       d.Probability.HasValue &&
                                                       d.DealValue.HasValue)
                                          .Sum(d => (double)(d.DealValue!.Value * (d.Probability!.Value / 100)));

                    string monthName = g.Key.Month switch
                    {
                        1 => "Jan",
                        2 => "Feb",
                        3 => "Mar",
                        4 => "Apr",
                        5 => "May",
                        6 => "Jun",
                        7 => "Jul",
                        8 => "Aug",
                        9 => "Sep",
                        10 => "Oct",
                        11 => "Nov",
                        12 => "Dec",
                        _ => "???"
                    };

                    return new
                    {
                        month = $"{monthName} {g.Key.Year}",
                        actualRevenue = actualRevenue,
                        expectedPipeline = expectedPipeline
                    };
                })
                .OrderBy(x => x.month)
                .ToList<object>();

            return trendData;
        }

        #endregion

        #region Deal Names

        /// <summary>
        /// Returns the distinct names of deals created during the specified month (matched on <c>CreatedAt</c>).
        /// </summary>
        /// <param name="monthName">Full month name (e.g., "January"), parsed via the invariant culture.</param>
        /// <param name="year">Four-digit year.</param>
        /// <returns>Distinct deal names for the month; entries may be null.</returns>
        public async Task<List<string?>> MonthlyDealNamesAsync(string monthName, int year)
        {
            int month = ParseMonthName(monthName);
            DateTime start = new DateTime(year, month, 1);
            DateTime end = start.AddMonths(1).AddDays(-1);

            return await GetDealNamesForPeriodAsync(start, end);
        }

        /// <summary>
        /// Returns the distinct names of deals created during the specified quarter (matched on <c>CreatedAt</c>).
        /// </summary>
        /// <param name="quarter">Quarter token "Q1"–"Q4" (case-insensitive).</param>
        /// <param name="year">Four-digit year.</param>
        /// <returns>Distinct deal names for the quarter; entries may be null.</returns>
        /// <exception cref="ArgumentException">Thrown when <paramref name="quarter"/> is not Q1–Q4.</exception>
        public async Task<List<string?>> QuarterlyDealNamesAsync(string quarter, int year)
        {
            int startMonth = quarter.ToUpper() switch
            {
                "Q1" => 1,
                "Q2" => 4,
                "Q3" => 7,
                "Q4" => 10,
                _ => throw new ArgumentException("Invalid quarter")
            };

            DateTime start = new DateTime(year, startMonth, 1);
            DateTime end = start.AddMonths(3).AddDays(-1);

            return await GetDealNamesForPeriodAsync(start, end);
        }

        /// <summary>
        /// Returns the distinct names of deals created during the specified year (matched on <c>CreatedAt</c>).
        /// </summary>
        /// <param name="year">Four-digit year.</param>
        /// <returns>Distinct deal names for the year; entries may be null.</returns>
        public async Task<List<string?>> YearlyDealNamesAsync(int year)
        {
            DateTime start = new DateTime(year, 1, 1);
            DateTime end = new DateTime(year, 12, 31);

            return await GetDealNamesForPeriodAsync(start, end);
        }

        /// <summary>
        /// Queries deals whose <c>CreatedAt</c> falls in the inclusive [start, end] range and returns their distinct names.
        /// </summary>
        /// <param name="start">Inclusive lower bound of the creation date range.</param>
        /// <param name="end">Inclusive upper bound of the creation date range.</param>
        /// <returns>Distinct deal names within the range; entries may be null.</returns>
        private async Task<List<string?>> GetDealNamesForPeriodAsync(DateTime start, DateTime end)
        {
            return await _context.Deals
                .Where(d => d.CreatedAt >= start && d.CreatedAt <= end)
                .Select(d => d.Name)
                .Distinct()
                .ToListAsync();
        }

        #endregion

        #region Helpers

        /// <summary>
        /// Converts a full month name to its 1–12 number, trimming whitespace and stripping any embedded quotes first.
        /// </summary>
        /// <param name="monthName">Full month name (e.g., "January"), interpreted with the invariant culture.</param>
        /// <returns>The month number, 1 through 12.</returns>
        /// <exception cref="ArgumentException">Thrown when <paramref name="monthName"/> is null, empty, or whitespace.</exception>
        /// <exception cref="FormatException">Thrown when the text is not a recognizable full month name.</exception>
        private int ParseMonthName(string monthName)
        {
            if (string.IsNullOrWhiteSpace(monthName))
                throw new ArgumentException("Month name cannot be empty.");

            monthName = monthName.Trim().Replace("\"", "");
            return DateTime.ParseExact(monthName, "MMMM", CultureInfo.InvariantCulture).Month;
        }

        #endregion
    }
}
