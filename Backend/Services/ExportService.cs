using Elpis_CRM.Data;
using ClosedXML.Excel;
using Elpis_CRM.Model;
using Microsoft.EntityFrameworkCore;
using System.Reflection;


namespace Elpis_CRM.Services
{
    /// <summary>
    /// Queries accounts and contacts from the database and renders the matching
    /// rows into in-memory ClosedXML Excel workbooks for download.
    /// </summary>
    public class ExportService
    {
        private readonly AppDbContext _context;

        /// <summary>
        /// Initializes the service with the database context used to read the records being exported.
        /// </summary>
        /// <param name="context">The application database context.</param>
        public ExportService(AppDbContext context)
        {
            _context = context;
        }

        // ─── Accounts ───────────────────────────────────────────────────────────

        /// <summary>
        /// Reads accounts (no tracking), narrows them by an optional case-insensitive
        /// search over name/city/country/phone/website and an optional tag substring,
        /// then renders the result into an "Accounts" Excel worksheet.
        /// </summary>
        /// <param name="search">Case-insensitive term matched against name, city, country, phone, and website; ignored when null or blank.</param>
        /// <param name="tag">Case-insensitive substring matched against the account's tags; ignored when null or blank.</param>
        /// <param name="columns">Property names to export as columns; when null or empty, every public property is included.</param>
        /// <returns>The .xlsx workbook as a byte array.</returns>
        public async Task<byte[]> ExportAccountsAsync(
            string? search,
            string? tag,
            List<string>? columns)
        {
            IQueryable<AccountModel> query = _context.Accounts.AsNoTracking();

            if (!string.IsNullOrWhiteSpace(search))
            {
                string s = search.Trim().ToLower();
                query = query.Where(a =>
                    (a.Name != null && a.Name.ToLower().Contains(s)) ||
                    (a.City != null && a.City.ToLower().Contains(s)) ||
                    (a.Country != null && a.Country.ToLower().Contains(s)) ||
                    (a.Phone != null && a.Phone.ToLower().Contains(s)) ||
                    (a.Website != null && a.Website.ToLower().Contains(s)));
            }

            if (!string.IsNullOrWhiteSpace(tag))
            {
                string t = tag.Trim().ToLower();
                query = query.Where(a => a.Tags != null && a.Tags.ToLower().Contains(t));
            }

            List<AccountModel> data = await query.ToListAsync();
            return BuildExcel(data, columns, "Accounts");
        }

        // ─── Contacts ───────────────────────────────────────────────────────────

        /// <summary>
        /// Reads contacts (no tracking), narrows them by an optional case-insensitive
        /// search over first/last name, work email, work phone, city, and country,
        /// then renders the result into a "Contacts" Excel worksheet.
        /// </summary>
        /// <param name="search">Case-insensitive term matched against first/last name, work email, work phone, city, and country; ignored when null or blank.</param>
        /// <param name="columns">Property names to export as columns; when null or empty, every public property is included.</param>
        /// <returns>The .xlsx workbook as a byte array.</returns>
        public async Task<byte[]> ExportContactsAsync(
            string? search,
            List<string>? columns)
        {
            IQueryable<ContactModel> query = _context.Contacts.AsNoTracking();

            if (!string.IsNullOrWhiteSpace(search))
            {
                string s = search.Trim().ToLower();
                query = query.Where(c =>
                    (c.FirstName != null && c.FirstName.ToLower().Contains(s)) ||
                    (c.LastName != null && c.LastName.ToLower().Contains(s)) ||
                    (c.WorkEmail != null && c.WorkEmail.ToLower().Contains(s)) ||
                    (c.WorkPhone != null && c.WorkPhone.ToLower().Contains(s)) ||
                    (c.City != null && c.City.ToLower().Contains(s)) ||
                    (c.Country != null && c.Country.ToLower().Contains(s)));
            }

            List<ContactModel> data = await query.ToListAsync();
            return BuildExcel(data, columns, "Contacts");
        }

        // ─── Generic Excel builder ───────────────────────────────────────────────

        /// <summary>
        /// Builds a single-sheet Excel workbook from a list of model objects using reflection.
        /// Only the properties listed in <paramref name="columns"/> are included (order-preserving,
        /// case-insensitive, unknown names skipped); if <paramref name="columns"/> is null/empty,
        /// ALL public instance properties are exported. The header row is styled, columns are
        /// auto-fit up to 60 characters, and the header is frozen.
        /// </summary>
        /// <typeparam name="T">The model type whose public instance properties become columns.</typeparam>
        /// <param name="rows">The records to write as data rows.</param>
        /// <param name="columns">Property names selecting which columns to export; null or empty exports all.</param>
        /// <param name="sheetName">Name applied to the generated worksheet.</param>
        /// <returns>The serialized .xlsx workbook as a byte array.</returns>
        private static byte[] BuildExcel<T>(
            IReadOnlyList<T> rows,
            List<string>? columns,
            string sheetName)
        {
            // Resolve which properties to export
            PropertyInfo[] allProps = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            PropertyInfo[] props = (columns != null && columns.Count > 0)
                ? columns
                    .Select(col => allProps.FirstOrDefault(p =>
                        string.Equals(p.Name, col, StringComparison.OrdinalIgnoreCase)))
                    .Where(p => p != null)
                    .Cast<PropertyInfo>()
                    .ToArray()
                : allProps;

            using var workbook = new XLWorkbook();
            IXLWorksheet ws = workbook.Worksheets.Add(sheetName);

            // Header row
            for (int c = 0; c < props.Length; c++)
            {
                IXLCell cell = ws.Cell(1, c + 1);
                cell.Value = props[c].Name;
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#4472C4");
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            // Data rows – stream values directly; avoid boxing where possible
            for (int r = 0; r < rows.Count; r++)
            {
                T row = rows[r];
                for (int c = 0; c < props.Length; c++)
                {
                    object? value = props[c].GetValue(row);
                    IXLCell cell = ws.Cell(r + 2, c + 1);
                    SetCellValue(cell, value);
                }
            }

            // Auto-fit after all data is written (one pass, avoids repeated resize)
            ws.ColumnsUsed().AdjustToContents(1, 60); // cap at 60 chars width

            // Freeze header row
            ws.SheetView.FreezeRows(1);

            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms.ToArray();
        }

        // ─── Helpers ────────────────────────────────────────────────────────────

        /// <summary>
        /// Writes a value into a cell using a type-appropriate representation: dates get a
        /// "yyyy-MM-dd HH:mm:ss" format, numeric and boolean values are written natively,
        /// null becomes an empty string, and anything else falls back to <c>ToString()</c>.
        /// </summary>
        /// <param name="cell">The target worksheet cell.</param>
        /// <param name="value">The property value to write; may be null.</param>
        private static void SetCellValue(IXLCell cell, object? value)
        {
            if (value is null)
            {
                cell.SetValue(string.Empty);
                return;
            }

            switch (value)
            {
                case DateTime dt:
                    cell.Value = dt;
                    cell.Style.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
                    break;
                case DateTimeOffset dto:
                    cell.Value = dto.DateTime;
                    cell.Style.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
                    break;
                case int i:    cell.Value = i;   break;
                case long l:   cell.Value = l;   break;
                case decimal d: cell.Value = d;  break;
                case double db: cell.Value = db; break;
                case bool b:   cell.Value = b;   break;
                default:       cell.Value = value.ToString(); break;
            }
        }
    }
}
