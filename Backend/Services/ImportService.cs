using CsvHelper;
using Microsoft.Data.SqlClient;
using Sylvan.Data.Csv;
using Sylvan.Data.Excel;
using System.Data;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Text;

namespace Elpis_CRM.Services;


/// <summary>
/// Outcome of an import run: overall success, the number of rows written versus skipped, the wall-clock
/// duration, and (when applicable) a top-level error message and per-row error details.
/// </summary>
/// <param name="Success">True when the bulk insert completed; false when validation or insertion failed.</param>
/// <param name="RowsImported">Rows actually written to the destination table.</param>
/// <param name="RowsSkipped">Rows read but not written (e.g. skipped bad rows).</param>
/// <param name="Elapsed">Total time taken, measured from the start of <see cref="ImportService.ImportAsync"/>.</param>
/// <param name="Error">Top-level failure reason, or null on success.</param>
/// <param name="RowErrors">Individual row-level error messages, or null when there were none.</param>
public sealed record ImportResult(
    bool Success,
    int RowsImported,
    int RowsSkipped,
    TimeSpan Elapsed,
    string? Error = null,
    IReadOnlyList<string>? RowErrors = null);

/// <summary>Progress snapshot reported during an import: rows read so far and batches completed.</summary>
/// <param name="RowsProcessed">Cumulative count of rows read from the source.</param>
/// <param name="BatchesCompleted">Number of full batches processed so far.</param>
public sealed record ImportProgress(int RowsProcessed, int BatchesCompleted);

/// <summary>Tunable knobs controlling batching, error handling, retries and large-file index management for an import.</summary>
public sealed record ImportOptions
{
    /// <summary>Rows per SqlBulkCopy batch. 5 000–10 000 is optimal for most workloads.</summary>
    public int BatchSize { get; init; } = 5_000;

    /// <summary>When true, rows that throw during conversion are skipped and logged instead of aborting the import.</summary>
    public bool SkipBadRows { get; init; } = false;

    /// <summary>Max retry attempts on transient SQL errors (exponential back-off: 2 s, 4 s, 8 s…).</summary>
    public int MaxRetryAttempts { get; init; } = 3;

    /// <summary>
    /// Files larger than this (bytes) will have non-clustered indexes disabled before insert
    /// and rebuilt after. Acts as a proxy for "large file" since we stream without knowing row count.
    /// Default: 5 MB ≈ ~50 k rows for typical CRM data.
    /// </summary>
    public long IndexDisableThresholdBytes { get; init; } = 5 * 1_024 * 1_024;

    /// <summary>Optional progress sink — called after every <see cref="BatchSize"/> rows.</summary>
    public IProgress<ImportProgress>? Progress { get; init; }
}



/// <summary>
/// Defines one destination column for an import: its SQL name, CLR type, nullability, and the
/// pre-compiled converter that turns a raw cell value into a SQL-ready object.
/// </summary>
internal sealed class ImportColumnDef
{
    /// <summary>The destination column name in the SQL table.</summary>
    public required string SqlColumnName { get; init; }

    /// <summary>The CLR type SqlBulkCopy reports for this column (e.g. string, long, DateTime).</summary>
    public required Type ClrType { get; init; }

    /// <summary>When false, the source file must supply this column or the import is rejected; when true, a missing column maps to DBNull.</summary>
    public bool IsNullable { get; init; }

    /// <summary>
    /// Pre-compiled value converter: raw Excel / CSV cell value → SQL-ready object.
    /// Called in the hot row-read loop. Must have ZERO reflection, ZERO regex,
    /// ZERO per-call heap allocations beyond the inescapable boxing of value types.
    /// </summary>
    public required Func<object?, object?> Converter { get; init; }
}

/// <summary>The full import schema for one table: the physical SQL table name and its ordered column definitions.</summary>
internal sealed class ImportTableSchema
{
    /// <summary>The physical SQL table the rows are written to (may differ from the import key, e.g. "CallLogs" → "CallLog").</summary>
    public required string TableName { get; init; }

    /// <summary>Ordered destination columns; identity columns are intentionally omitted because SqlBulkCopy rejects them.</summary>
    public required ImportColumnDef[] Columns { get; init; }
}



/// <summary>
/// Static catalogue of every importable table and its column converters. <see cref="All"/> is the
/// single source of truth consulted by <see cref="ImportService"/>; add a table here (plus a
/// controller endpoint) to make it importable.
/// </summary>
internal static class SchemaRegistry
{
    /// <summary>String converter: maps null/DBNull to DBNull, otherwise the value's string representation.</summary>
    private static readonly Func<object?, object?> Str =
        static v => v is null or DBNull ? (object)DBNull.Value : v.ToString()!;

    /// <summary>Builds an int converter; the nullable variant returns DBNull for blank/unparseable input, the non-nullable variant returns 0.</summary>
    /// <param name="nullable">Whether unparseable or empty cells become DBNull (true) or 0 (false).</param>
    /// <returns>A converter that coerces int/long/double/string cells to an int (or DBNull/0).</returns>
    private static Func<object?, object?> Int(bool nullable) => nullable
        ? static v => v switch
        {
            null or DBNull => (object)DBNull.Value,
            int i => (object)i,
            long l => (object)(int)l,
            double d => (object)(int)d,
            string s => int.TryParse(s, out var r) ? (object)r : (object)DBNull.Value,
            _ => (object)Convert.ToInt32(v)
        }
        : static v => v switch
        {
            null or DBNull => (object)0,
            int i => (object)i,
            long l => (object)(int)l,
            double d => (object)(int)d,
            string s => int.TryParse(s, out var r) ? (object)r : (object)0,
            _ => (object)Convert.ToInt32(v)
        };

    /// <summary>Builds a long converter; the nullable variant returns DBNull for blank/unparseable input, the non-nullable variant returns 0.</summary>
    /// <param name="nullable">Whether unparseable or empty cells become DBNull (true) or 0L (false).</param>
    /// <returns>A converter that coerces long/int/double/string cells to a long (or DBNull/0).</returns>
    private static Func<object?, object?> Long(bool nullable) => nullable
        ? static v => v switch
        {
            null or DBNull => (object)DBNull.Value,
            long l => (object)l,
            int i => (object)(long)i,
            double d => (object)(long)d,
            string s => long.TryParse(s, out var r) ? (object)r : (object)DBNull.Value,
            _ => (object)Convert.ToInt64(v)
        }
        : static v => v switch
        {
            null or DBNull => (object)0L,
            long l => (object)l,
            int i => (object)(long)i,
            double d => (object)(long)d,
            string s => long.TryParse(s, out var r) ? (object)r : (object)0L,
            _ => (object)Convert.ToInt64(v)
        };

    /// <summary>Builds a decimal converter; strings are parsed with invariant culture and <c>NumberStyles.Any</c>, falling back to DBNull (nullable) or 0 (non-nullable).</summary>
    /// <param name="nullable">Whether unparseable or empty cells become DBNull (true) or 0m (false).</param>
    /// <returns>A converter that coerces decimal/double/int/long/string cells to a decimal (or DBNull/0).</returns>
    private static Func<object?, object?> Dec(bool nullable) => nullable
        ? static v => v switch
        {
            null or DBNull => (object)DBNull.Value,
            decimal dm => (object)dm,
            double d => (object)(decimal)d,
            int i => (object)(decimal)i,
            long l => (object)(decimal)l,
            string s => decimal.TryParse(s,
                                  System.Globalization.NumberStyles.Any,
                                  System.Globalization.CultureInfo.InvariantCulture,
                                  out var r)
                              ? (object)r : (object)DBNull.Value,
            _ => (object)Convert.ToDecimal(v)
        }
        : static v => v switch
        {
            null or DBNull => (object)0m,
            decimal dm => (object)dm,
            double d => (object)(decimal)d,
            int i => (object)(decimal)i,
            long l => (object)(decimal)l,
            string s => decimal.TryParse(s,
                                  System.Globalization.NumberStyles.Any,
                                  System.Globalization.CultureInfo.InvariantCulture,
                                  out var r)
                              ? (object)r : (object)0m,
            _ => (object)Convert.ToDecimal(v)
        };

    /// <summary>
    /// Builds a datetime converter. Excel OADate doubles are converted directly; strings are parsed
    /// as invariant/UTC (a trailing " UTC" suffix is stripped first). Unparseable values fall back to
    /// DBNull when nullable, or to <see cref="DateTime.UtcNow"/> when not.
    /// </summary>
    /// <param name="nullable">Whether blank/unparseable cells become DBNull (true) or the current UTC time (false).</param>
    /// <returns>A converter that coerces DateTime/double/string cells to a DateTime.</returns>
    private static Func<object?, object?> Dt(bool nullable)
    {
        static object? ParseDateString(string s, bool isNullable)
        {
            // Handle strings with " UTC" suffix (e.g., "2026-01-02 09:55:31 UTC")
            var trimmed = s.EndsWith(" UTC") ? s[..^4].Trim() : s;
            var parsed = DateTime.TryParse(trimmed,
                                  System.Globalization.CultureInfo.InvariantCulture,
                                  System.Globalization.DateTimeStyles.AssumeUniversal,
                                  out var r);
            return parsed ? (object)r : (isNullable ? (object)DBNull.Value : (object)DateTime.UtcNow);
        }

        return nullable
            ? static v => v switch
            {
                null or DBNull => (object)DBNull.Value,
                DateTime dt => (object)dt,
                double d => (object)DateTime.FromOADate(d),
                string s => ParseDateString(s, true),
                _ => (object)Convert.ToDateTime(v)
            }
            : static v => v switch
            {
                null or DBNull => (object)DateTime.UtcNow,
                DateTime dt => (object)dt,
                double d => (object)DateTime.FromOADate(d),
                string s => ParseDateString(s, false),
                _ => (object)Convert.ToDateTime(v)
            };
    }

    // ── Date-only (SQL date column) — stored as DateTime at midnight ─────────

    // private static readonly Func<object?, object?> DateOnly =
    //     static v => v switch
    //     {
    //         null or DBNull => (object)DBNull.Value,
    //         DateTime dt => (object)dt.Date,
    //         double d => (object)DateTime.FromOADate(d).Date,
    //         string s => DateTime.TryParse(s,
    //                               System.Globalization.CultureInfo.InvariantCulture,
    //                               System.Globalization.DateTimeStyles.None,
    //                               out var r)
    //                           ? (object)r.Date : (object)DBNull.Value,
    //         _ => (object)DBNull.Value
    //     };
 /// <summary>
    /// Converter for SQL <c>date</c> columns: keeps only the date component. Excel OADate doubles and
    /// strings (with optional " UTC" suffix, parsed as universal time) are accepted; anything
    /// unparseable becomes DBNull.
    /// </summary>
    private static readonly Func<object?, object?> DateOnly =
        static v => v switch
        {
            null or DBNull => (object)DBNull.Value,
            DateTime dt => (object)dt.Date,
            double d => (object)DateTime.FromOADate(d).Date,
            string s => DateTime.TryParse(
                              s.EndsWith(" UTC", StringComparison.OrdinalIgnoreCase) ? s[..^4] : s,
                              System.Globalization.CultureInfo.InvariantCulture,
                              System.Globalization.DateTimeStyles.AssumeUniversal | System.Globalization.DateTimeStyles.AdjustToUniversal,
                              out var r)
                          ? (object)r.Date : (object)DBNull.Value,
            _ => (object)DBNull.Value
        };
    // ── Table definitions ────────────────────────────────────────────────────
    // Rules:
    //   • Identity columns (AccountId, ContactId, Id) are OMITTED — SqlBulkCopy rejects them.
    //   • is_nullable = 0  → IsNullable = false  (file MUST supply this column)
    //   • is_nullable = 1  → IsNullable = true   (missing cell → DBNull → DB DEFAULT / NULL)
    //   • nvarchar         → Str
    //   • int nullable     → Int(true)   | int not-null → Int(false)
    //   • decimal nullable → Dec(true)
    //   • datetime nullable→ Dt(true)
    //   • date nullable    → DateOnly   (SQL date type; we strip the time component)

    /// <summary>
    /// Case-insensitive registry mapping an import key (Accounts, Contacts, Deals, DealContactLinks,
    /// CallLogs, Notes, Tasks) to its table schema. Identity columns are deliberately excluded so
    /// SqlBulkCopy can write the remaining columns.
    /// </summary>
    public static readonly IReadOnlyDictionary<string, ImportTableSchema> All =
        new Dictionary<string, ImportTableSchema>(StringComparer.OrdinalIgnoreCase)
        {
            
            // Accounts  (AccountId is NOT identity — must be imported from file)
            
            ["Accounts"] = new()
            {
                TableName = "Accounts",
                Columns =
                [
                    new() { SqlColumnName = "AccountId",                 ClrType = typeof(long),     IsNullable = false,  Converter = Long(true) },
                    new() { SqlColumnName = "Name",                      ClrType = typeof(string),   IsNullable = false, Converter = Str       },
                    new() { SqlColumnName = "NumberOfEmployees",         ClrType = typeof(int),      IsNullable = true,  Converter = Int(true) },
                    new() { SqlColumnName = "AnnualRevenue",             ClrType = typeof(decimal),  IsNullable = true,  Converter = Dec(true) },
                    new() { SqlColumnName = "Website",                   ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Phone",                     ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "DisplayPhone",              ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Address",                   ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "City",                      ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "State",                     ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Zipcode",                   ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Country",                   ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Facebook",                  ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Twitter",                   ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "LinkedIn",                  ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "LastContactedTime",         ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "LastContactedMode",         ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "LastActivityType",          ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "LastActivityDate",          ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "RecentNote",                ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "IndustryType",              ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "BusinessType",              ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Territory",                 ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "CreatedAt",                 ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "UpdatedAt",                 ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "ActiveSalesSequences",      ClrType = typeof(int),      IsNullable = true,  Converter = Int(true) },
                    new() { SqlColumnName = "CompletedSalesSequences",   ClrType = typeof(int),      IsNullable = true,  Converter = Int(true) },
                    new() { SqlColumnName = "LastAssignedAt",            ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "ParentAccountId",           ClrType = typeof(long),      IsNullable = true,  Converter = Long(true) },
                    new() { SqlColumnName = "ParentAccount",             ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "SalesOwnerId",              ClrType = typeof(long),     IsNullable = true,  Converter = Long(true) },
                    new() { SqlColumnName = "SalesOwner",                ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "CreatedById",               ClrType = typeof(long),      IsNullable = true,  Converter = Long(true) },
                    new() { SqlColumnName = "CreatedBy",                 ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "UpdatedById",               ClrType = typeof(long),      IsNullable = true,  Converter = Long(true) },
                    new() { SqlColumnName = "UpdatedBy",                 ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "ImportID",                  ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Tags",                      ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                ]
            },

            // ════════════════════════════════════════════════════════════════
            // Contacts  (ContactId is NOT identity — must be imported from file)
            // ════════════════════════════════════════════════════════════════
            ["Contacts"] = new()
            {
                TableName = "Contacts",
                Columns =
                [
                    new() { SqlColumnName = "ContactId",                 ClrType = typeof(long),     IsNullable = false,  Converter = Long(true) },
                    new() { SqlColumnName = "FirstName",                  ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "LastName",                   ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "JobTitle",                   ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "WorkPhone",                  ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Mobile",                     ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "LastSeenOnChat",             ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "Locale",                     ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "TotalChatSessions",          ClrType = typeof(int),      IsNullable = true,  Converter = Int(true) },
                    new() { SqlColumnName = "FirstSeenOnChat",            ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "ExternalID",                 ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "TimeZone",                   ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Address",                    ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "City",                       ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "State",                      ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Zipcode",                    ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Country",                    ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Facebook",                   ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Twitter",                    ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "LinkedIn",                   ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Medium",                     ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Keyword",                    ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "LastContactedTime",          ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "LastContactedMode",          ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "LastActivityType",           ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "LastActivityDate",           ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "LastSeenOnWeb",              ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "Score",                      ClrType = typeof(int),      IsNullable = true,  Converter = Int(true) },
                    new() { SqlColumnName = "SubscriptionStatus",         ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "UnsubscribeReason",          ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "OtherUnsubscribeReasons",    ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "WhatsAppSubscriptionStatus", ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "SMSSubscriptionStatus",      ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "RecentNote",                 ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "CreatedAt",                  ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "UpdatedAt",                  ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "Source",                     ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Campaign",                   ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Territory",                  ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "LifeCycleStage",             ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Status",                     ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "LostReason",                 ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "OriginalCampaign",           ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "OriginalMedium",             ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "OriginalSource",             ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "CreatedThroughCampaign",     ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "CreatedFromMedium",          ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "CreatedFromSource",          ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "MostRecentCampaign",         ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "MostRecentMedium",           ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "MostRecentSource",           ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "WorkEmail",                  ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "ActiveSalesSequences",       ClrType = typeof(int),      IsNullable = true,  Converter = Int(true) },
                    new() { SqlColumnName = "CompletedSalesSequences",    ClrType = typeof(int),      IsNullable = true,  Converter = Int(true) },
                    new() { SqlColumnName = "CustomerFit",                ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "WebForms",                   ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "LastAssignedAt",             ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "AccountId",                  ClrType = typeof(long),     IsNullable = true,  Converter = Long(true) },
                    new() { SqlColumnName = "Account",                    ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "SalesOwnerId",               ClrType = typeof(long),     IsNullable = true,  Converter = Long(true) },
                    new() { SqlColumnName = "SalesOwner",                 ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "CreatedById",                ClrType = typeof(long),     IsNullable = true,  Converter = Long(true) },
                    new() { SqlColumnName = "CreatedBy",                  ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "UpdatedById",                ClrType = typeof(long),     IsNullable = true,  Converter = Long(true) },
                    new() { SqlColumnName = "UpdatedBy",                  ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "ImportID",                   ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Emails",                     ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Products",                   ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Message",                    ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Tags",                       ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Lists",                      ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                ]
            },

            // ════════════════════════════════════════════════════════════════
            // Deals  (Id is NOT identity — must be imported from file)
            // ════════════════════════════════════════════════════════════════
            ["Deals"] = new()
            {
                TableName = "Deals",
                Columns =
                [
                    new() { SqlColumnName = "Id",                      ClrType = typeof(long),     IsNullable = false,  Converter = Long(true) },
                    new() { SqlColumnName = "Name",                    ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "DealValue",               ClrType = typeof(decimal),  IsNullable = true,  Converter = Dec(true) },
                    new() { SqlColumnName = "Currency",                ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "DealValueInBaseCurrency", ClrType = typeof(decimal),  IsNullable = true,  Converter = Dec(true) },
                    new() { SqlColumnName = "ExpectedCloseDate",       ClrType = typeof(DateTime), IsNullable = true,  Converter = DateOnly  },
                    new() { SqlColumnName = "ClosedDate",              ClrType = typeof(DateTime), IsNullable = true,  Converter = DateOnly  },
                    new() { SqlColumnName = "Probability",             ClrType = typeof(int),      IsNullable = true,  Converter = Int(true) },
                    new() { SqlColumnName = "LastActivityType",        ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "LastActivityDate",        ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "RecentNote",              ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "CreatedAt",               ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "UpdatedAt",               ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "DealPipeline",            ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "DealStage",               ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "LostReason",              ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "PaymentStatus",           ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Territory",               ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Type",                    ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Source",                  ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Campaign",                ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "ForecastCategory",        ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "ActiveSalesSequences",    ClrType = typeof(int),      IsNullable = true,  Converter = Int(true) },
                    new() { SqlColumnName = "CompletedSalesSequences", ClrType = typeof(int),      IsNullable = true,  Converter = Int(true) },
                    new() { SqlColumnName = "WebForm",                 ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "UpcomingActivities",      ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "DealStageUpdatedAt",      ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "LastAssignedAt",          ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "ExpectedDealValue",       ClrType = typeof(decimal),  IsNullable = true,  Converter = Dec(true) },
                    new() { SqlColumnName = "AccountId",               ClrType = typeof(long),     IsNullable = true,  Converter = Long(true) },
                    new() { SqlColumnName = "AccountName",             ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "ContactId",               ClrType = typeof(long),     IsNullable = true,  Converter = Long(true) },
                    new() { SqlColumnName = "ContactName",             ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "SalesOwnerId",            ClrType = typeof(long),     IsNullable = true,  Converter = Long(true) },
                    new() { SqlColumnName = "SalesOwner",              ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "CreatedById",             ClrType = typeof(long),     IsNullable = true,  Converter = Long(true) },
                    new() { SqlColumnName = "CreatedBy",               ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "UpdatedById",             ClrType = typeof(long),     IsNullable = true,  Converter = Long(true) },
                    new() { SqlColumnName = "UpdatedBy",               ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "ImportID",                ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "EnquiryNumber",           ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Tags",                    ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "AgeInDays",               ClrType = typeof(int),      IsNullable = true,  Converter = Int(true) },
                ]
            },

            // CallLog (CallLogId is IDENTITY — excluded)
            ["DealContactLinks"] = new()
            {
                TableName = "DealContactLinks",
                Columns =
                [
                    new() { SqlColumnName = "DealId",    ClrType = typeof(long), IsNullable = false, Converter = Long(true) },
                    new() { SqlColumnName = "ContactId", ClrType = typeof(long), IsNullable = false, Converter = Long(true) },
                ]
            },

            ["CallLogs"] = new()
            {
                TableName = "CallLog",
                Columns =
                [
                    new() { SqlColumnName = "CallOwner",          ClrType = typeof(string),   IsNullable = true,  Converter = Str        },
                    new() { SqlColumnName = "CreatedAt",          ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)   },
                    new() { SqlColumnName = "CallDirection",      ClrType = typeof(string),   IsNullable = true,  Converter = Str        },
                    new() { SqlColumnName = "CallStatus",         ClrType = typeof(string),   IsNullable = true,  Converter = Str        },
                    new() { SqlColumnName = "CallDuration",       ClrType = typeof(string),   IsNullable = true,  Converter = Str        },
                    new() { SqlColumnName = "Outcome",            ClrType = typeof(string),   IsNullable = true,  Converter = Str        },
                    new() { SqlColumnName = "Phone",              ClrType = typeof(string),   IsNullable = true,  Converter = Str        },
                    new() { SqlColumnName = "CallType",           ClrType = typeof(string),   IsNullable = true,  Converter = Str        },
                    new() { SqlColumnName = "Notes",              ClrType = typeof(string),   IsNullable = true,  Converter = Str        },
                    new() { SqlColumnName = "AssociatedWithCall", ClrType = typeof(string),   IsNullable = true,  Converter = Str        },
                    new() { SqlColumnName = "ContactId",          ClrType = typeof(long),     IsNullable = true,  Converter = Long(true) },
                    new() { SqlColumnName = "AccountId",          ClrType = typeof(long),     IsNullable = true,  Converter = Long(true) },
                    new() { SqlColumnName = "DealId",             ClrType = typeof(long),     IsNullable = true,  Converter = Long(true) },
                ]
            },

            // Notes (Id is IDENTITY — excluded)
            ["Notes"] = new()
            {
                TableName = "Notes",
                Columns =
                [
                    new() { SqlColumnName = "Description",   ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "CreatedById",   ClrType = typeof(long),      IsNullable = true,  Converter = Long(true) },
                    new() { SqlColumnName = "CreatedAt",     ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "UpdatedAt",     ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "ContactId",     ClrType = typeof(long),     IsNullable = true,  Converter = Long(true) },
                    new() { SqlColumnName = "RelatedToType", ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                ]
            },

            // Tasks (Id is IDENTITY — excluded)
            ["Tasks"] = new()
            {
                TableName = "Tasks",
                Columns =
                [
                    new() { SqlColumnName = "Title",            ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Description",      ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "Status",           ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "TaskType",         ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "DueDate",          ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "CompletedDate",    ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "Outcome",          ClrType = typeof(string),   IsNullable = true,  Converter = Str       },
                    new() { SqlColumnName = "OwnerId",          ClrType = typeof(long),     IsNullable = true,  Converter = Long(true) },
                    new() { SqlColumnName = "CreatedById",      ClrType = typeof(long),     IsNullable = true,  Converter = Long(true) },
                    new() { SqlColumnName = "UpdatedById",      ClrType = typeof(long),     IsNullable = true,  Converter = Long(true) },
                    new() { SqlColumnName = "CreatedAt",        ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "UpdatedAt",        ClrType = typeof(DateTime), IsNullable = true,  Converter = Dt(true)  },
                    new() { SqlColumnName = "ContactId",        ClrType = typeof(long),     IsNullable = true,  Converter = Long(true) },
                ]
            },
        };
    }

// ───────────────────────────────────────────────────────────────────────────────
// Header Mapper
// Called ONCE before the row loop. Never touches the hot path.
// ───────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Matches source file headers to schema columns by fuzzy-normalizing both sides. Invoked once per
/// import, before the row loop, so it never touches the hot path.
/// </summary>
internal static class HeaderMapper
{
    /// <summary>
    /// Normalizes a header string for fuzzy matching.
    /// Strips spaces, underscores, dashes, parentheses, slashes and lowercases.
    /// Uses stackalloc — zero heap allocation for headers ≤ 256 chars.
    /// "Account Name" → "accountname"  |  "account_name" → "accountname"
    /// </summary>
    /// <param name="header">The raw header text to canonicalize.</param>
    /// <returns>The lowercased header with separator characters removed.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static string Normalize(ReadOnlySpan<char> header)
    {
        Span<char> buf = stackalloc char[Math.Min(header.Length, 256)];
        int pos = 0;
        foreach (char c in header)
        {
            if (c is not (' ' or '_' or '-' or '(' or ')' or '/' or '\\'))
                buf[pos++] = char.ToLowerInvariant(c);
        }
        return new string(buf[..pos]);
    }

    /// <summary>
    /// Builds the source-column index map for a single import.
    /// Returns int[] where result[i] = source file column index for output column i,
    /// or -1 when the column is absent (nullable → sends DBNull; non-nullable → error).
    /// </summary>
    /// <param name="source">The opened source reader whose header names are matched against the schema.</param>
    /// <param name="schema">The destination table schema whose columns are being resolved.</param>
    /// <param name="errors">Receives one message per required column that is missing from the file; empty when the map is valid.</param>
    /// <returns>An index array aligned to <see cref="ImportTableSchema.Columns"/>; -1 marks an absent column.</returns>
    public static int[] BuildMap(IDataReader source, ImportTableSchema schema, out List<string> errors)
    {
        errors = [];

        // Build normalized lookup from file headers (one Dictionary allocation, done once)
        int fieldCount = source.FieldCount;
        var lookup = new Dictionary<string, int>(fieldCount, StringComparer.Ordinal);
        for (int i = 0; i < fieldCount; i++)
            lookup.TryAdd(Normalize(source.GetName(i)), i); // first occurrence wins on duplicates

        var map = new int[schema.Columns.Length];
        for (int i = 0; i < schema.Columns.Length; i++)
        {
            var col = schema.Columns[i];
            if (lookup.TryGetValue(Normalize(col.SqlColumnName), out int srcIdx))
            {
                map[i] = srcIdx;
            }
            else if (!col.IsNullable)
            {
                errors.Add($"Required column '{col.SqlColumnName}' not found in file.");
                map[i] = -1;
            }
            else
            {
                // Missing optional column → -1 → GetValue returns DBNull
                // SqlBulkCopy (without KeepNulls) will apply the DB column's DEFAULT.
                map[i] = -1;
            }
        }

        return map;
    }
}

// ───────────────────────────────────────────────────────────────────────────────
// MappedDataReader
// The IDataReader handed directly to SqlBulkCopy.WriteToServerAsync().
// HOT PATH: Read() + GetValue() are called 100 000+ times per import.
// Design constraints:
//   • No reflection           • No Regex           • No per-row heap allocations
//   • No dynamic type checks  • No dictionary reads inside GetValue
// ───────────────────────────────────────────────────────────────────────────────

/// <summary>
/// An <see cref="IDataReader"/> adapter handed straight to <c>SqlBulkCopy.WriteToServerAsync</c>. It
/// projects the source reader's columns onto the destination schema using a precomputed index map and
/// per-column converters, presenting columns in schema order. Built for the hot path: no reflection,
/// regex, or per-row allocations. SqlBulkCopy only exercises Read/GetValue/FieldCount; the remaining
/// IDataReader members exist solely to satisfy the interface.
/// </summary>
internal sealed class MappedDataReader : IDataReader
{
    private readonly IDataReader _src;
    private readonly int[] _map;   // output col i → source col index (-1 = DBNull)
    private readonly Func<object?, object?>[] _conv;  // pre-compiled converter per output column
    private readonly string[] _names;
    private readonly Type[] _types;
    private readonly IProgress<ImportProgress>? _progress;
    private readonly int _batchSize;

    /// <summary>Total rows read from the source so far; populated as <see cref="Read"/> advances.</summary>
    public int RowsRead { get; private set; }

    /// <summary>Rows read but deliberately not written. Tracked here for the caller to read after the import.</summary>
    public int RowsSkipped { get; private set; }

    /// <summary>Accumulated per-row error messages, surfaced in the final <see cref="ImportResult"/>.</summary>
    public List<string> RowErrors { get; } = [];

    /// <summary>
    /// Wraps a source reader with a column map and converters so SqlBulkCopy sees the destination schema.
    /// </summary>
    /// <param name="source">The underlying CSV/Excel reader; not owned by this instance (lifetime managed by the caller).</param>
    /// <param name="columnMap">Per-output-column source index (-1 = absent → DBNull), as produced by <see cref="HeaderMapper.BuildMap"/>.</param>
    /// <param name="columns">Destination column definitions, used to extract names, CLR types and converters.</param>
    /// <param name="batchSize">Row interval at which progress is reported.</param>
    /// <param name="progress">Optional sink notified once per completed batch.</param>
    public MappedDataReader(
        IDataReader source,
        int[] columnMap,
        ImportColumnDef[] columns,
        int batchSize = 5_000,
        IProgress<ImportProgress>? progress = null)
    {
        _src = source;
        _map = columnMap;
        _conv = Array.ConvertAll(columns, static c => c.Converter);
        _names = Array.ConvertAll(columns, static c => c.SqlColumnName);
        _types = Array.ConvertAll(columns, static c => c.ClrType);
        _batchSize = batchSize;
        _progress = progress;
    }

    // ── Hot path ─────────────────────────────────────────────────────────────

    /// <summary>Advances to the next source row, incrementing <see cref="RowsRead"/> and reporting progress every batch.</summary>
    /// <returns>True if a row was read; false at end of data.</returns>
    public bool Read()
    {
        bool ok = _src.Read();
        if (ok)
        {
            RowsRead++;
            if (_progress is not null && RowsRead % _batchSize == 0)
                _progress.Report(new ImportProgress(RowsRead, RowsRead / _batchSize));
        }
        return ok;
    }

    /// <summary>
    /// Returns the converted value for output column <paramref name="i"/>: DBNull for an absent column,
    /// DBNull if the source cell can't be read, otherwise the result of the column's converter.
    /// </summary>
    /// <param name="i">Zero-based destination column index.</param>
    /// <returns>The SQL-ready value, never a CLR null (uses <see cref="DBNull.Value"/> instead).</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public object GetValue(int i)
    {
        int srcIdx = _map[i];
        if (srcIdx < 0) return DBNull.Value;     // missing optional column

        object? raw;
        try { raw = _src.GetValue(srcIdx); }
        catch { return DBNull.Value; }           // unreadable cell → null-safe

        // Delegate dispatch — inlined by JIT after first few calls per column.
        // No reflection, no boxing beyond what the delegate itself decides.
        object? result = _conv[i](raw is DBNull ? null : raw);
        return result ?? DBNull.Value;
    }

    /// <summary>Reports whether output column <paramref name="i"/> is null — true for an absent column or a null source cell.</summary>
    /// <param name="i">Zero-based destination column index.</param>
    /// <returns>True when the value is DBNull.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public bool IsDBNull(int i)
    {
        int srcIdx = _map[i];
        return srcIdx < 0 || _src.IsDBNull(srcIdx);
    }

    // ── IDataReader plumbing (SqlBulkCopy only uses Read / GetValue / FieldCount) ──

    /// <summary>Number of destination columns exposed to SqlBulkCopy.</summary>
    public int FieldCount => _map.Length;

    /// <summary>Returns the destination SQL column name at the given index.</summary>
    /// <param name="i">Zero-based destination column index.</param>
    public string GetName(int i) => _names[i];

    /// <summary>Returns the CLR type declared for the destination column at the given index.</summary>
    /// <param name="i">Zero-based destination column index.</param>
    public Type GetFieldType(int i) => _types[i];

    /// <summary>Resolves a destination column name to its index (case-insensitive).</summary>
    /// <param name="name">The destination column name to look up.</param>
    /// <returns>The matching zero-based index.</returns>
    /// <exception cref="IndexOutOfRangeException">Thrown when no column matches <paramref name="name"/>.</exception>
    public int GetOrdinal(string name)
    {
        for (int i = 0; i < _names.Length; i++)
            if (string.Equals(_names[i], name, StringComparison.OrdinalIgnoreCase)) return i;
        throw new IndexOutOfRangeException($"Column '{name}' not found.");
    }

    /// <summary>Fills <paramref name="values"/> with the converted values of the current row.</summary>
    /// <param name="values">Destination buffer; only the first <see cref="FieldCount"/> slots (or fewer) are written.</param>
    /// <returns>The number of values copied.</returns>
    public int GetValues(object[] values)
    {
        int count = Math.Min(values.Length, FieldCount);
        for (int i = 0; i < count; i++) values[i] = GetValue(i);
        return count;
    }

    // Typed accessors — SqlBulkCopy routes through GetValue so these rarely fire,
    // but IDataReader requires them.
    public string GetString(int i) => (string)GetValue(i);
    public bool GetBoolean(int i) => (bool)GetValue(i);
    public byte GetByte(int i) => (byte)GetValue(i);
    public char GetChar(int i) => (char)GetValue(i);
    public short GetInt16(int i) => (short)GetValue(i);
    public int GetInt32(int i) => (int)GetValue(i);
    public long GetInt64(int i) => (long)GetValue(i);
    public float GetFloat(int i) => (float)GetValue(i);
    public double GetDouble(int i) => (double)GetValue(i);
    public decimal GetDecimal(int i) => (decimal)GetValue(i);
    public DateTime GetDateTime(int i) => (DateTime)GetValue(i);
    public Guid GetGuid(int i) => (Guid)GetValue(i);

    public string GetDataTypeName(int i) => _types[i].Name;
    public long GetBytes(int i, long fo, byte[]? buf, int bo, int len) => 0;
    public long GetChars(int i, long fo, char[]? buf, int bo, int len) => 0;
    public IDataReader GetData(int i) => throw new NotSupportedException();
    public DataTable? GetSchemaTable() => null;

    public object this[int i] => GetValue(i);
    public object this[string name] => GetValue(GetOrdinal(name));

    // Lifecycle — source lifetime is managed by ImportService.
    public void Close() { }
    public void Dispose() { }
    public int Depth => 0;
    public bool IsClosed => false;
    public int RecordsAffected => -1;
    public bool NextResult() => false;
}

// ───────────────────────────────────────────────────────────────────────────────
// Index Manager
// Disables non-clustered indexes before a large insert and rebuilds afterward.
// Runs on a separate SqlConnection command (not inside the BulkCopy).
// ───────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Disables a table's non-clustered indexes before a large bulk insert and rebuilds them afterward,
/// trading insert throughput against a one-time rebuild cost. Operates via standalone commands on the
/// supplied connection, outside the SqlBulkCopy. Primary keys and unique constraints are left untouched.
/// </summary>
internal static class IndexManager
{
    private const string FindQuery = """
        SELECT i.name
        FROM   sys.indexes  i
        JOIN   sys.objects  o ON i.object_id = o.object_id
        WHERE  o.name               = @t
          AND  i.type_desc          = 'NONCLUSTERED'
          AND  i.is_disabled        = 0
          AND  i.is_primary_key     = 0
          AND  i.is_unique_constraint = 0;
        """;

    /// <summary>
    /// Finds the table's enabled, non-clustered, non-key/non-unique indexes and disables each one.
    /// </summary>
    /// <param name="conn">An open connection to the target database.</param>
    /// <param name="table">The table whose indexes should be disabled.</param>
    /// <param name="ct">Cancellation token.</param>
    /// <returns>The names of the indexes that were disabled, to pass back to <see cref="RebuildAsync"/>.</returns>
    public static async Task<string[]> DisableAsync(
        SqlConnection conn, string table, CancellationToken ct)
    {
        var names = new List<string>();

        await using (var cmd = new SqlCommand(FindQuery, conn))
        {
            cmd.Parameters.Add("@t", SqlDbType.NVarChar, 128).Value = table;
            await using var rdr = await cmd.ExecuteReaderAsync(ct);
            while (await rdr.ReadAsync(ct))
                names.Add(rdr.GetString(0));
        }

        foreach (string name in names)
        {
            await using var cmd =
                new SqlCommand($"ALTER INDEX [{name}] ON [{table}] DISABLE;", conn);
            await cmd.ExecuteNonQueryAsync(ct);
        }

        return [.. names];
    }

    /// <summary>
    /// Rebuilds the named indexes (offline, with a 10-minute per-index timeout). Intended to undo a
    /// prior <see cref="DisableAsync"/> call once the bulk insert finishes.
    /// </summary>
    /// <param name="conn">An open connection to the target database.</param>
    /// <param name="table">The table whose indexes are being rebuilt.</param>
    /// <param name="indexes">The index names to rebuild (typically the output of <see cref="DisableAsync"/>).</param>
    /// <param name="ct">Cancellation token.</param>
    public static async Task RebuildAsync(
        SqlConnection conn, string table, string[] indexes, CancellationToken ct)
    {
        foreach (string name in indexes)
        {
            // Remove ONLINE = OFF if not supported by your SQL Server edition.
            await using var cmd = new SqlCommand(
                $"ALTER INDEX [{name}] ON [{table}] REBUILD WITH (ONLINE = OFF);", conn);
            cmd.CommandTimeout = 600; // index rebuilds can be slow on large tables
            await cmd.ExecuteNonQueryAsync(ct);
        }
    }
}

// ───────────────────────────────────────────────────────────────────────────────
// Import Service — public entry point
// Thread-safe: stateless beyond injected _cs/_log (both immutable after ctor).
// ───────────────────────────────────────────────────────────────────────────────

/// <summary>
/// High-throughput bulk-import service that streams an uploaded CSV/Excel file into a registered CRM
/// table via SqlBulkCopy. Validates the file type and schema, maps headers to columns, optionally
/// disables/rebuilds indexes for large files, and retries the insert on transient SQL errors. Stateless
/// after construction and therefore thread-safe.
/// </summary>
public sealed class ImportService
{
    private readonly string _cs;
    private readonly ILogger<ImportService> _log;

    private static readonly HashSet<string> AllowedExtensions =
        new(StringComparer.OrdinalIgnoreCase) { ".xlsx", ".xlsb", ".xls", ".csv" };

    /// <summary>
    /// Resolves and caches the "DefaultConnection" connection string for all imports.
    /// </summary>
    /// <param name="cfg">Application configuration supplying the connection string.</param>
    /// <param name="log">Logger for import lifecycle and error events.</param>
    /// <exception cref="InvalidOperationException">Thrown when "DefaultConnection" is not configured.</exception>
    public ImportService(IConfiguration cfg, ILogger<ImportService> log)
    {
        _cs = cfg.GetConnectionString("DefaultConnection")
               ?? throw new InvalidOperationException("'DefaultConnection' missing from configuration.");
        _log = log;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Public entry point
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Imports an uploaded file into the named table end-to-end: validates the extension and schema,
    /// opens a CSV or Excel reader, maps the file's headers to the destination columns (failing fast if
    /// a required column is missing), then bulk-inserts in batches with retry. For files at or above the
    /// configured threshold it disables non-clustered indexes first and always attempts to rebuild them
    /// afterward, even if the insert fails. Errors are returned in the result rather than thrown.
    /// </summary>
    /// <param name="file">The uploaded CSV/Excel file; must be non-empty with a supported extension.</param>
    /// <param name="tableName">Import key identifying a schema in <see cref="SchemaRegistry.All"/>.</param>
    /// <param name="options">Batch size, retry count, index threshold and progress sink; defaults are used when null.</param>
    /// <param name="ct">Cancellation token for the read/insert phases (index rebuild always runs to completion).</param>
    /// <returns>
    /// A successful <see cref="ImportResult"/> with row counts and timing, or a failed one whose
    /// <see cref="ImportResult.Error"/> describes the validation or insertion failure.
    /// </returns>
    public async Task<ImportResult> ImportAsync(
        IFormFile file,
        string tableName,
        ImportOptions? options = null,
        CancellationToken ct = default)
    {
        options ??= new ImportOptions();
        var sw = Stopwatch.StartNew();

        // ── Fast-fail validation ───────────────────────────────────────────────
        if (file is null || file.Length == 0)
            return Fail("File is empty or missing.", sw);

        var ext = Path.GetExtension(file.FileName);
        if (!AllowedExtensions.Contains(ext))
            return Fail($"Unsupported file type '{ext}'. Supported: {string.Join(", ", AllowedExtensions)}", sw);

        if (!SchemaRegistry.All.TryGetValue(tableName, out var schema))
            return Fail($"No import schema registered for table '{tableName}'.", sw);

        // ── Open the file reader ──────────────────────────────────────────────
        // ASP.NET Core spools uploads > 64 KB to disk, so OpenReadStream() is seekable.
        await using var stream = file.OpenReadStream();

        IDataReader srcReader;
        try
        {
            srcReader = ext.Equals(".csv", StringComparison.OrdinalIgnoreCase)
                ? await OpenCsvAsync(stream, ct)
                : OpenExcel(stream, ext);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "Cannot open file '{File}'", file.FileName);
            return Fail($"Cannot open file: {ex.Message}", sw);
        }

        // srcReader owns its resources; MappedDataReader is a non-owning view over it.
        using (srcReader)
        {
            // ── Build header map — executed ONCE, before the row loop ──────────
            var sourceIndices = HeaderMapper.BuildMap(srcReader, schema, out var mapErrors);
            if (mapErrors.Count > 0)
                return Fail(string.Join(" | ", mapErrors), sw);

            _log.LogInformation(
                "Import start → [{Table}] | {File} ({Size:N0} bytes) | {Cols} columns mapped",
                tableName, file.FileName, file.Length, schema.Columns.Length);

            var reader = new MappedDataReader(
                srcReader, sourceIndices, schema.Columns, options.BatchSize, options.Progress);

            // ── SQL connection + optional index management ─────────────────────
            await using var conn = new SqlConnection(_cs);
            await conn.OpenAsync(ct);

            bool disableIdx = file.Length >= options.IndexDisableThresholdBytes;
            string[]? disabledIdx = null;

            if (disableIdx)
            {
                _log.LogInformation(
                    "File ≥ threshold ({Bytes:N0} bytes) — disabling non-clustered indexes on [{Table}]…",
                    file.Length, schema.TableName);

                disabledIdx = await IndexManager.DisableAsync(conn, schema.TableName, ct);

                _log.LogInformation("Disabled {N} index(es).", disabledIdx.Length);
            }

            // ── Bulk insert with retry ────────────────────────────────────────
            Exception? bulkError = null;
            try
            {
                await RunWithRetryAsync(
                    _ => BulkInsertAsync(conn, schema.TableName, reader, options, ct),
                    options.MaxRetryAttempts, ct);
            }
            catch (Exception ex)
            {
                bulkError = ex;
                _log.LogError(ex, "Bulk insert failed for [{Table}]", schema.TableName);
            }
            finally
            {
                // Always attempt index rebuild — even after a failed import.
                if (disabledIdx?.Length > 0)
                {
                    _log.LogInformation(
                        "Rebuilding {N} index(es) on [{Table}]…", disabledIdx.Length, schema.TableName);
                    try
                    {
                        // Use CancellationToken.None — we must attempt rebuild regardless of caller cancel.
                        await IndexManager.RebuildAsync(conn, schema.TableName, disabledIdx, CancellationToken.None);
                    }
                    catch (Exception ex)
                    {
                        _log.LogError(ex,
                            "Index rebuild failed on [{Table}] — manual intervention required. " +
                            "Run: ALTER INDEX ALL ON [{Table}] REBUILD;", schema.TableName, schema.TableName);
                    }
                }
            }

            if (bulkError is not null)
                return Fail($"Bulk insert error: {bulkError.Message}", sw);

            sw.Stop();

            int imported = reader.RowsRead - reader.RowsSkipped;
            _log.LogInformation(
                "Import complete → [{Table}] | {Imported:N0} rows in {Elapsed:0.00}s ({Rate:N0} rows/s)",
                tableName, imported, sw.Elapsed.TotalSeconds,
                sw.Elapsed.TotalSeconds > 0 ? imported / sw.Elapsed.TotalSeconds : 0);

            return new ImportResult(
                Success: true,
                RowsImported: imported,
                RowsSkipped: reader.RowsSkipped,
                Elapsed: sw.Elapsed,
                RowErrors: reader.RowErrors.Count > 0 ? reader.RowErrors : null);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Bulk insert
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Streams the mapped reader into the destination table with SqlBulkCopy, using a table lock and an
    /// internal per-batch transaction and adding name-to-name column mappings. Constraint checking and
    /// triggers are deliberately not enabled, favouring throughput.
    /// </summary>
    /// <param name="conn">An open connection to the target database.</param>
    /// <param name="tableName">The physical destination table name (bracket-quoted internally).</param>
    /// <param name="reader">The schema-projected reader supplying rows and column names.</param>
    /// <param name="options">Supplies the batch size used by the bulk copy.</param>
    /// <param name="ct">Cancellation token for the write.</param>
    private static async Task BulkInsertAsync(
        SqlConnection conn,
        string tableName,
        MappedDataReader reader,
        ImportOptions options,
        CancellationToken ct)
    {
        // TableLock         → exclusive table lock for the duration; prevents row-lock overhead.
        // UseInternalTransaction → each batch is auto-committed; no external transaction needed.
        // CheckConstraints / FireTriggers intentionally omitted for maximum throughput.
        // Add them if your schema strictly requires constraint enforcement at insert time.
        const SqlBulkCopyOptions BulkOpts =
            SqlBulkCopyOptions.TableLock |
            SqlBulkCopyOptions.UseInternalTransaction;

        using var bulk = new SqlBulkCopy(conn, BulkOpts, externalTransaction: null)
        {
            DestinationTableName = $"[{tableName}]",
            BatchSize = options.BatchSize,
            EnableStreaming = true,
            BulkCopyTimeout = 600,
        };

        // Name-to-name column mappings — SqlBulkCopy resolves case-insensitively.
        for (int i = 0; i < reader.FieldCount; i++)
        {
            string colName = reader.GetName(i);
            bulk.ColumnMappings.Add(colName, colName);
        }

        await bulk.WriteToServerAsync(reader, ct);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // File readers
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Opens an Excel workbook as an <see cref="IDataReader"/>, choosing the Sylvan workbook type from
    /// the extension (.xlsx/.xlsb/.xls). Row 1 is treated as headers; reading starts at the first data row.
    /// </summary>
    /// <param name="stream">The uploaded file stream.</param>
    /// <param name="ext">The file extension, used to select the workbook format.</param>
    /// <returns>A reader positioned to read data rows.</returns>
    /// <exception cref="NotSupportedException">Thrown for an unrecognized Excel extension.</exception>
    private static IDataReader OpenExcel(Stream stream, string ext)
    {
        // Sylvan reads row 1 as headers by default; GetName(i) returns those values.
        // Read() starts at row 2 (first data row). No manual header handling needed.
        var type = ext.ToLowerInvariant() switch
        {
            ".xlsx" => ExcelWorkbookType.ExcelXml,
            ".xlsb" => ExcelWorkbookType.ExcelBinary,
            ".xls" => ExcelWorkbookType.Excel,
            _ => throw new NotSupportedException($"Unsupported Excel format '{ext}'.")
        };
        return ExcelDataReader.Create(stream, type);
    }

    /// <summary>
    /// Opens a CSV file as an <see cref="IDataReader"/> using Sylvan, treating the first row as headers,
    /// honouring a UTF-8 BOM (as written by Excel exports) and auto-detecting the delimiter.
    /// </summary>
    /// <param name="stream">The uploaded file stream.</param>
    /// <param name="_">Cancellation token (currently unused).</param>
    /// <returns>A reader positioned to read data rows.</returns>
    private static async Task<IDataReader> OpenCsvAsync(Stream stream, CancellationToken _)
    {
        // detectEncodingFromByteOrderMarks handles UTF-8 BOM from Excel CSV exports.
        var textReader = new StreamReader(
            stream,
            Encoding.UTF8,
            detectEncodingFromByteOrderMarks: true,
            bufferSize: 65_536,
            leaveOpen: true);

        var opts = new CsvDataReaderOptions
        {
            HasHeaders = true,
            BufferSize = 65_536,
            // Sylvan auto-detects comma / tab delimiter; override here if needed:
            // Delimiter = ',',
        };

        return await Sylvan.Data.Csv.CsvDataReader.CreateAsync(textReader, opts);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Retry with exponential back-off (transient SQL errors only)
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Runs <paramref name="action"/>, retrying only on transient SQL errors with exponential back-off
    /// (2 s, 4 s, 8 s …) up to <paramref name="maxAttempts"/>. Non-transient errors and the final
    /// attempt's failure propagate to the caller.
    /// </summary>
    /// <param name="action">The operation to run; receives the 1-based attempt number.</param>
    /// <param name="maxAttempts">Maximum number of attempts before giving up.</param>
    /// <param name="ct">Cancellation token observed during the back-off delay.</param>
    private static async Task RunWithRetryAsync(
        Func<int, Task> action, int maxAttempts, CancellationToken ct)
    {
        for (int attempt = 1; ; attempt++)
        {
            try
            {
                await action(attempt);
                return;
            }
            catch (SqlException ex) when (attempt < maxAttempts && IsTransient(ex))
            {
                // Exponential back-off: 2 s → 4 s → 8 s
                await Task.Delay(TimeSpan.FromSeconds(Math.Pow(2, attempt)), ct);
            }
        }
    }

    /// <summary>
    /// Returns true for known transient SQL Server / Azure SQL error codes (deadlock, timeout, dropped
    /// connection, Azure throttling/busy/unavailable). Non-transient errors (schema mismatch, FK
    /// violation, etc.) return false and propagate immediately.
    /// </summary>
    /// <param name="ex">The SQL exception to classify.</param>
    /// <returns>True if the error is worth retrying.</returns>
    private static bool IsTransient(SqlException ex) => ex.Number is
        1205 or  // deadlock victim
        -2 or  // timeout
        233 or  // connection broken
        10053 or
        10054 or
        10060 or
        40197 or  // Azure: service busy
        40501 or  // Azure: throttled
        40613;    // Azure: database unavailable

    // ─────────────────────────────────────────────────────────────────────────
    // Helpers
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>Builds a failed <see cref="ImportResult"/> with zero row counts and the elapsed time so far.</summary>
    /// <param name="error">The failure message to surface to the caller.</param>
    /// <param name="sw">The running stopwatch whose elapsed time is captured.</param>
    /// <returns>An unsuccessful result carrying <paramref name="error"/>.</returns>
    private static ImportResult Fail(string error, Stopwatch sw) =>
        new(false, 0, 0, sw.Elapsed, error);
}
