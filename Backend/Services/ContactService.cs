using Elpis_CRM.Data;
using Elpis_CRM.Dtos;
using Elpis_CRM.Model;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Threading.Tasks;

namespace Elpis_CRM.Services
{
    /// <summary>
    /// Data-access and business logic for contacts: paging and search, tag-based lookups, lifecycle
    /// stats, sequential enquiry-number generation, account-FK resolution, duplicate-email guarding
    /// and cascade-aware deletion.
    /// </summary>
    public class ContactService
    {
        private readonly AppDbContext _contactContext;

        /// <summary>
        /// Creates the service over the given EF Core database context.
        /// </summary>
        /// <param name="context">The application's database context.</param>
        public ContactService(AppDbContext context)
        {
            _contactContext = context;
        }

        private const string EnquiryNoPrefix = "EITSPL-MKT-EQ-";
        private const string EstimatedQuotePrefix = "Estimate Quote EST-";

        /// <summary>
        /// Generates the next sequential EnquiryNo (e.g. "EITSPL-MKT-EQ-003").
        /// Looks at the highest existing numeric suffix across all contacts and increments it.
        /// Retries on a rare race-condition collision (two requests generating the same number at once).
        /// </summary>
        private async Task<string> GenerateNextEnquiryNoAsync()
        {
            const int maxAttempts = 5;

            for (int attempt = 0; attempt < maxAttempts; attempt++)
            {
                // Pull existing EnquiryNo values that match our prefix.
                var existing = await _contactContext.Contacts
                    .Where(c => c.EnquiryNo != null && c.EnquiryNo.StartsWith(EnquiryNoPrefix))
                    .Select(c => c.EnquiryNo)
                    .ToListAsync();

                var maxNumber = 0;
                foreach (var val in existing)
                {
                    var suffix = val!.Substring(EnquiryNoPrefix.Length);
                    if (int.TryParse(suffix, out var num) && num > maxNumber)
                    {
                        maxNumber = num;
                    }
                }

                var nextNumber = maxNumber + 1;
                var candidate = $"{EnquiryNoPrefix}{nextNumber:D3}";

                // Make sure nobody else just took this number (basic race-condition guard).
                var collision = await _contactContext.Contacts
                    .AnyAsync(c => c.EnquiryNo == candidate);

                if (!collision)
                {
                    return candidate;
                }
                // else loop and retry with a fresh max
            }

            // Extremely unlikely fallback: append a short random suffix to avoid blocking the request.
            return $"{EnquiryNoPrefix}{DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() % 1000:D3}";
        }

        /// <summary>
        /// Generates the next sequential EstimatedQuote (e.g. "EST-003"), mirroring the
        /// EnquiryNo scheme: highest existing numeric suffix + 1, with a race-condition retry.
        /// </summary>
        private async Task<string> GenerateNextEstimatedQuoteAsync()
        {
            const int maxAttempts = 5;

            for (int attempt = 0; attempt < maxAttempts; attempt++)
            {
                var existing = await _contactContext.Contacts
                    .Where(c => c.EstimatedQuote != null && c.EstimatedQuote.StartsWith(EstimatedQuotePrefix))
                    .Select(c => c.EstimatedQuote)
                    .ToListAsync();

                var maxNumber = 0;
                foreach (var val in existing)
                {
                    var suffix = val!.Substring(EstimatedQuotePrefix.Length);
                    if (int.TryParse(suffix, out var num) && num > maxNumber)
                    {
                        maxNumber = num;
                    }
                }

                var candidate = $"{EstimatedQuotePrefix}{maxNumber + 1:D3}";

                var collision = await _contactContext.Contacts
                    .AnyAsync(c => c.EstimatedQuote == candidate);

                if (!collision)
                {
                    return candidate;
                }
            }

            return $"{EstimatedQuotePrefix}{DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() % 1000:D3}";
        }

        /// <summary>
        /// Returns one page of contacts ordered newest-first (by CreatedAt), plus the total matching count.
        /// When a search of at least 2 characters is given, it is tokenized on spaces and every token must
        /// match at least one searchable field (name, emails, phones, account, owner, enquiry no, territory, tags).
        /// </summary>
        /// <param name="page">1-based page number; values below 1 are coerced to 1.</param>
        /// <param name="pageSize">Rows per page; coerced into the range 1–500.</param>
        /// <param name="search">Optional search text; ignored when blank or shorter than 2 characters.</param>
        /// <returns>A tuple of the page of contacts and the total count across all pages.</returns>
        // Get Contacts with Pagination (optional server-side search)
        public async Task<(List<ContactModel> items, int totalCount)> GetAllAsync(int page = 1, int pageSize = 150, string? search = null)
        {
            if (page < 1) page = 1;
            if (pageSize < 1) pageSize = 150;
            if (pageSize > 500) pageSize = 500;

            var query = _contactContext.Contacts.AsNoTracking();

            // Token-based search: each word must match at least one field (AND across tokens).
            // "Appa Rao" → "appa" in FirstName AND "rao" in LastName → match.
            if (!string.IsNullOrWhiteSpace(search) && search.Trim().Length >= 2)
            {
                var tokens = search.Trim().ToLowerInvariant()
                    .Split(' ', StringSplitOptions.RemoveEmptyEntries);
                foreach (var tok in tokens)
                {
                    var t = tok;
                    query = query.Where(c =>
                        (c.FirstName  != null && c.FirstName.ToLower().Contains(t))  ||
                        (c.LastName   != null && c.LastName.ToLower().Contains(t))   ||
                        (c.WorkEmail  != null && c.WorkEmail.ToLower().Contains(t))  ||
                        (c.WorkPhone  != null && c.WorkPhone.ToLower().Contains(t))  ||
                        (c.Mobile     != null && c.Mobile.ToLower().Contains(t))     ||
                        (c.Account    != null && c.Account.ToLower().Contains(t))    ||
                        (c.SalesOwner != null && c.SalesOwner.ToLower().Contains(t)) ||
                        (c.EnquiryNo != null && c.EnquiryNo.ToLower().Contains(t)) ||
                        (c.EstimatedQuote != null && c.EstimatedQuote.ToLower().Contains(t)) ||
                        (c.Territory  != null && c.Territory.ToLower().Contains(t))  ||
                        (c.Tags       != null && c.Tags.ToLower().Contains(t)));
                }
            }

            var totalCount = await query.CountAsync();
            var items = await query
                .OrderByDescending(c => c.CreatedAt ?? DateTime.MinValue)
                .Skip((page - 1) * pageSize)
                .Take(pageSize)
                .ToListAsync();
            return (items, totalCount);
        }

        /// <summary>
        /// Fetches a single contact by its primary key.
        /// </summary>
        /// <param name="id">Primary key of the contact.</param>
        /// <returns>The matching contact, or null when none has that ID.</returns>
        // Get contact by ID
        public async Task<ContactModel?> GetByIdAsync(long id)
        {
            return await _contactContext.Contacts
                .FirstOrDefaultAsync(c => c.ContactId == id);
        }

        /// <summary>
        /// Returns the contacts under one account, ordered by first then last name and capped at <paramref name="limit"/>,
        /// optionally filtered by a substring on name/email/phone. Results are trimmed projections carrying only
        /// the ID, first/last name, work email and account.
        /// </summary>
        /// <param name="accountId">Account whose contacts are requested.</param>
        /// <param name="q">Optional case-insensitive substring filter on name, work email or phone.</param>
        /// <param name="limit">Maximum rows; clamped to the range 1–500.</param>
        /// <returns>The matching contact projections, or an empty list when none match.</returns>
        // Get contacts by account ID (optional name/email filter + cap for large accounts)
        public async Task<List<ContactModel>> GetContactsByAccountIdAsync(long accountId, string? q = null, int limit = 300)
        {
            limit = Math.Clamp(limit, 1, 500);
            var query = _contactContext.Contacts
                .AsNoTracking()
                .Where(c => c.AccountId == accountId);

            if (!string.IsNullOrWhiteSpace(q))
            {
                var t = q.Trim().ToLowerInvariant();
                if (t.Length >= 1)
                {
                    query = query.Where(c =>
                        (c.FirstName != null && c.FirstName.ToLower().Contains(t)) ||
                        (c.LastName != null && c.LastName.ToLower().Contains(t)) ||
                        (c.WorkEmail != null && c.WorkEmail.ToLower().Contains(t)) ||
                        (c.Mobile != null && c.Mobile.ToLower().Contains(t)) ||
                        (c.WorkPhone != null && c.WorkPhone.ToLower().Contains(t)));
                }
            }

            return await query
                .OrderBy(c => c.FirstName)
                .ThenBy(c => c.LastName)
                .Take(limit)
                .Select(c => new ContactModel
                {
                    ContactId = c.ContactId,
                    FirstName = c.FirstName,
                    LastName = c.LastName,
                    WorkEmail = c.WorkEmail,
                    Account = c.Account
                })
                .ToListAsync();
        }

        /// <summary>
        /// Global contact typeahead, ordered by last then first name. A numeric query also matches on contact ID;
        /// otherwise the text is tokenized and every token must match some name/email/phone/account field.
        /// </summary>
        /// <param name="q">Search text; an empty list is returned when it has fewer than 2 non-blank characters.</param>
        /// <param name="limit">Maximum rows; clamped to the range 1–100.</param>
        /// <returns>Matching contacts, or an empty list when the query is too short.</returns>
        public async Task<List<ContactModel>> SearchAsync(string? q, int limit = 50)
        {
            if (string.IsNullOrWhiteSpace(q) || q.Trim().Length < 2)
                return new List<ContactModel>();

            limit = Math.Clamp(limit, 1, 100);
            var term = q.Trim().ToLowerInvariant();
            var isNumeric = long.TryParse(term, out var idVal);

            var query = _contactContext.Contacts.AsNoTracking().AsQueryable();

            if (isNumeric)
            {
                query = query.Where(c =>
                    c.ContactId == idVal ||
                    (c.FirstName != null && c.FirstName.ToLower().Contains(term)) ||
                    (c.LastName != null && c.LastName.ToLower().Contains(term)) ||
                    (c.WorkEmail != null && c.WorkEmail.ToLower().Contains(term)) ||
                    (c.Account != null && c.Account.ToLower().Contains(term)));
            }
            else
            {
                // Token-based: each word must match some field
                var tokens = term.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                foreach (var tok in tokens)
                {
                    var t = tok;
                    query = query.Where(c =>
                        (c.FirstName != null && c.FirstName.ToLower().Contains(t)) ||
                        (c.LastName  != null && c.LastName.ToLower().Contains(t))  ||
                        (c.WorkEmail != null && c.WorkEmail.ToLower().Contains(t)) ||
                        (c.Mobile    != null && c.Mobile.ToLower().Contains(t))    ||
                        (c.WorkPhone != null && c.WorkPhone.ToLower().Contains(t)) ||
                        (c.EnquiryNo != null && c.EnquiryNo.ToLower().Contains(t)) ||
                        (c.EstimatedQuote != null && c.EstimatedQuote.ToLower().Contains(t)) ||
                        (c.Account   != null && c.Account.ToLower().Contains(t)));
                }
            }

            return await query
                .OrderBy(c => c.LastName)
                .ThenBy(c => c.FirstName)
                .Take(limit)
                .ToListAsync();
        }

        /// <summary>
        /// Groups all contacts by lifecycle stage in the database, then folds the stages into five fixed
        /// buckets (prospect, engaged, customer, promoter, other), routing null/blank/unrecognized stages to "other".
        /// </summary>
        /// <returns>A dictionary keyed by the five bucket names with their contact counts.</returns>
        public async Task<Dictionary<string, int>> GetLifeCycleStageCountsAsync()
        {
            var groups = await _contactContext.Contacts
                .AsNoTracking()
                .GroupBy(c => c.LifeCycleStage ?? "")
                .Select(g => new { Stage = g.Key, Count = g.Count() })
                .ToListAsync();

            var prospect = 0;
            var engaged = 0;
            var customer = 0;
            var promoter = 0;
            var other = 0;

            foreach (var row in groups)
            {
                var s = (row.Stage ?? "").Trim();
                if (string.IsNullOrEmpty(s))
                {
                    other += row.Count;
                    continue;
                }

                switch (s.ToLowerInvariant())
                {
                    case "prospect":
                        prospect += row.Count;
                        break;
                    case "engaged":
                        engaged += row.Count;
                        break;
                    case "customer":
                        customer += row.Count;
                        break;
                    case "promoter":
                        promoter += row.Count;
                        break;
                    default:
                        other += row.Count;
                        break;
                }
            }

            return new Dictionary<string, int>
            {
                ["prospect"] = prospect,
                ["engaged"] = engaged,
                ["customer"] = customer,
                ["promoter"] = promoter,
                ["other"] = other
            };
        }

        /// <summary>
        /// Builds the distinct set of individual tags across all contacts by splitting each contact's
        /// comma-separated Tags field, trimming each tag and removing duplicates.
        /// </summary>
        /// <returns>The unique tag names; empty when no contact carries any tag.</returns>
        // Get all unique tags
        public async Task<List<string>> GetAllTagsAsync()
        {
            var tags = await _contactContext.Contacts
                .Where(c => !string.IsNullOrEmpty(c.Tags))
                .Select(c => c.Tags)
                .ToListAsync();

            return tags
                .SelectMany(t => t.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
                .Select(t => t.Trim())
                .Where(t => t.Length > 0)
                .Distinct()
                .ToList();
        }

        /// <summary>
        /// Gathers the email addresses of contacts that carry at least one of the requested tags (matched by
        /// comparing each contact's comma-separated, trimmed tags against the requested set). Addresses are
        /// pulled from WorkEmail and from the Emails field (split on "," or ";"), then de-duplicated case-insensitively.
        /// </summary>
        /// <param name="tags">Comma-separated tags to match; a contact qualifies if any of its tags is in this set.</param>
        /// <returns>A comma-joined string of distinct email addresses; an empty string when nothing matches.</returns>
        // Splits a stored Tags string on "," or ";" into trimmed, non-empty tags.
        private static readonly char[] TagSeparators = { ',', ';' };
        private static IEnumerable<string> SplitTags(string? tags) =>
            (tags ?? "").Split(TagSeparators, StringSplitOptions.RemoveEmptyEntries).Select(t => t.Trim());

        // Builds a SQL-translatable predicate — c.Tags LIKE '%t1%' OR c.Tags LIKE '%t2%' … — so the
        // database narrows to the (usually small) tagged subset. Without this the old code pulled every
        // tagged contact — including its business-card image blobs — into memory just to match tags.
        private static Expression<Func<ContactModel, bool>> TagsContainAny(List<string> tags)
        {
            var param = Expression.Parameter(typeof(ContactModel), "c");
            var tagsProp = Expression.Property(param, nameof(ContactModel.Tags));
            var contains = typeof(string).GetMethod(nameof(string.Contains), new[] { typeof(string) })!;
            Expression body = Expression.Constant(false);
            foreach (var t in tags)
            {
                var notNull = Expression.NotEqual(tagsProp, Expression.Constant(null, typeof(string)));
                var like = Expression.Call(tagsProp, contains, Expression.Constant(t));
                body = Expression.OrElse(body, Expression.AndAlso(notNull, like));
            }
            return Expression.Lambda<Func<ContactModel, bool>>(body, param);
        }

        // Get emails by tags — returns both WorkEmail and the Emails field
        public async Task<string> GetEmailsByTagsAsync(string tags)
        {
            var selectedTags = SplitTags(tags).Where(t => t.Length > 0).ToList();
            if (selectedTags.Count == 0) return string.Empty;

            // Pre-filter in the DB and project to ONLY the tag + email columns (never the image blobs).
            var rows = await _contactContext.Contacts.AsNoTracking()
                .Where(TagsContainAny(selectedTags))
                .Where(c => c.WorkEmail != null || c.Emails != null)
                .Select(c => new { c.Tags, c.WorkEmail, c.Emails })
                .ToListAsync();

            var emails = new List<string>();
            foreach (var c in rows)
            {
                // Exact-match refine (a substring pre-filter can over-match, e.g. "VIP" vs "VIPer").
                if (!SplitTags(c.Tags).Any(tag => selectedTags.Contains(tag))) continue;

                if (!string.IsNullOrWhiteSpace(c.WorkEmail)) emails.Add(c.WorkEmail.Trim());
                if (!string.IsNullOrWhiteSpace(c.Emails))
                    emails.AddRange(c.Emails.Split(TagSeparators, StringSplitOptions.RemoveEmptyEntries).Select(e => e.Trim()));
            }

            return string.Join(",", emails.Where(e => !string.IsNullOrWhiteSpace(e)).Distinct(StringComparer.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Returns one row per (contact, email address) for every contact carrying at least one of the
        /// requested tags, each keeping the contact's display name. Mirrors <see cref="GetEmailsByTagsAsync"/>
        /// but retains the identity that one discards, which a per-recipient greeting needs.
        /// </summary>
        /// <param name="tags">Comma-separated tags to match; a contact qualifies if any of its tags is in this set.</param>
        /// <returns>The matching recipients, de-duplicated by address (case-insensitive); empty when nothing matches.</returns>
        public async Task<List<ContactRecipientDto>> GetRecipientsByTagsAsync(string tags)
        {
            var selectedTags = SplitTags(tags).Where(t => t.Length > 0).ToList();
            if (selectedTags.Count == 0) return new List<ContactRecipientDto>();

            // Pre-filter in the DB and project to only what's needed (never the image blobs).
            var rows = await _contactContext.Contacts.AsNoTracking()
                .Where(TagsContainAny(selectedTags))
                .Where(c => c.WorkEmail != null || c.Emails != null)
                .Select(c => new { c.ContactId, c.FirstName, c.LastName, c.Tags, c.WorkEmail, c.Emails })
                .ToListAsync();

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var recipients = new List<ContactRecipientDto>();
            foreach (var c in rows)
            {
                // Exact-match refine (a substring pre-filter can over-match, e.g. "VIP" vs "VIPer").
                if (!SplitTags(c.Tags).Any(tag => selectedTags.Contains(tag))) continue;

                var name = $"{c.FirstName} {c.LastName}".Trim();

                var addresses = new List<string>();
                if (!string.IsNullOrWhiteSpace(c.WorkEmail)) addresses.Add(c.WorkEmail.Trim());
                if (!string.IsNullOrWhiteSpace(c.Emails))
                    addresses.AddRange(c.Emails.Split(TagSeparators, StringSplitOptions.RemoveEmptyEntries).Select(e => e.Trim()));

                foreach (var email in addresses)
                {
                    // An address shared by two contacts is only emailed once — first contact wins,
                    // matching the de-duplication GetEmailsByTagsAsync already does.
                    if (string.IsNullOrWhiteSpace(email) || !seen.Add(email)) continue;
                    recipients.Add(new ContactRecipientDto { ContactId = c.ContactId, Name = name, Email = email });
                }
            }

            return recipients;
        }

        /// <summary>
        /// Returns the contacts that carry at least one of the requested tags. Narrows in the DB by
        /// substring, then refines for an exact tag match in memory; image blobs are stripped from the
        /// result so the payload stays small.
        /// </summary>
        /// <param name="tags">Comma-separated tags to match; a contact qualifies if any of its tags is in this set.</param>
        /// <returns>The matching contacts; an empty list when the argument is blank or nothing matches.</returns>
        // Get contacts by tags
        public async Task<List<ContactModel>> GetContactsByTagsAsync(string tags)
        {
            if (string.IsNullOrWhiteSpace(tags))
                return new List<ContactModel>();

            var selectedTags = SplitTags(tags).Where(t => t.Length > 0).ToList();
            if (selectedTags.Count == 0) return new List<ContactModel>();

            // DB-side pre-filter: only load contacts whose Tags contain the tag, not every tagged row.
            var contacts = await _contactContext.Contacts.AsNoTracking()
                .Where(TagsContainAny(selectedTags))
                .ToListAsync();

            var matched = contacts
                .Where(c => SplitTags(c.Tags).Any(tag => selectedTags.Contains(tag)))
                .ToList();

            // Don't ship the image blobs to the table — it never uses them (the slide-in fetches
            // images by id), and they bloat the response.
            foreach (var c in matched) { c.FrontImage = null; c.BackImage = null; }
            return matched;
        }

        /// <summary>
        /// Returns every contact's email addresses (WorkEmail + the multi-value Emails field) as a
        /// de-duplicated comma-joined string, optionally narrowed by a token search. Projected so it
        /// never loads the image blobs — used for "email all selected contacts" without paging.
        /// </summary>
        public async Task<string> GetAllEmailsAsync(string? search)
        {
            var query = _contactContext.Contacts.AsNoTracking()
                .Where(c => c.WorkEmail != null || c.Emails != null);

            if (!string.IsNullOrWhiteSpace(search) && search.Trim().Length >= 2)
            {
                var tokens = search.Trim().ToLowerInvariant().Split(' ', StringSplitOptions.RemoveEmptyEntries);
                foreach (var tok in tokens)
                {
                    var t = tok;
                    query = query.Where(c =>
                        (c.FirstName != null && c.FirstName.ToLower().Contains(t)) ||
                        (c.LastName != null && c.LastName.ToLower().Contains(t)) ||
                        (c.WorkEmail != null && c.WorkEmail.ToLower().Contains(t)) ||
                        (c.Emails != null && c.Emails.ToLower().Contains(t)) ||
                        (c.Account != null && c.Account.ToLower().Contains(t)) ||
                        (c.Tags != null && c.Tags.ToLower().Contains(t)));
                }
            }

            var rows = await query.Select(c => new { c.WorkEmail, c.Emails }).ToListAsync();

            var emails = new List<string>();
            foreach (var c in rows)
            {
                if (!string.IsNullOrWhiteSpace(c.WorkEmail)) emails.Add(c.WorkEmail.Trim());
                if (!string.IsNullOrWhiteSpace(c.Emails))
                    emails.AddRange(c.Emails.Split(TagSeparators, StringSplitOptions.RemoveEmptyEntries).Select(e => e.Trim()));
            }

            return string.Join(",", emails.Where(e => !string.IsNullOrWhiteSpace(e)).Distinct(StringComparer.OrdinalIgnoreCase));
        }

        /// <summary>
        /// When the client sends <see cref="ContactModel.Account"/> but not <see cref="ContactModel.AccountId"/>,
        /// resolve the FK from Accounts (fixes deals / GetContactsByAccountId empty lists).
        /// </summary>
        private async Task EnsureAccountIdResolvedAsync(ContactModel contact)
        {
            if (contact == null)
            {
                return;
            }
            if (contact.AccountId.HasValue && contact.AccountId.Value > 0)
            {
                return;
            }
            if (string.IsNullOrWhiteSpace(contact.Account)) 
            {
                return;
            }
            var trimmed = contact.Account.Trim();
            var acc = await _contactContext.Accounts.AsNoTracking().FirstOrDefaultAsync(a => a.Name != null && a.Name == trimmed);
            if (acc == null)
            {
                var lower = trimmed.ToLowerInvariant();
                acc = await _contactContext.Accounts
                    .AsNoTracking()
                    .FirstOrDefaultAsync(a => a.Name != null && a.Name.ToLower() == lower);
            }

            if (acc != null)
            {
                contact.AccountId = acc.AccountId;
            }
        }

        /// <summary>
        /// Persists a new contact: stamps the created/updated and activity timestamps, resolves the account FK
        /// from the account name when missing, and rejects the insert if any of its emails (WorkEmail or the
        /// comma-separated Emails field) already exists on another contact. If a positive ID is supplied and
        /// already exists (import scenario), the existing contact is updated instead; otherwise an ID is generated.
        /// </summary>
        /// <param name="contact">Contact to create; its EnquiryNo is ignored and set by this method.</param>
        /// <param name="generateEnquiryNo">When true, assigns a fresh sequential EnquiryNo; otherwise it is left null.</param>
        /// <returns>The saved contact (or the updated existing one when the ID already existed).</returns>
        /// <exception cref="InvalidOperationException">Thrown when a contact with the same email already exists.</exception>
        // Create contact
        public async Task<ContactModel> AddAsync(ContactModel contact, bool generateEnquiryNo = false, bool generateEstimatedQuote = false)
        {
            var now = DateTime.UtcNow;

            contact.CreatedAt = now;
            contact.UpdatedAt = now;

            if (contact.LastActivityDate == null) contact.LastActivityDate = now;
            if (contact.LastAssignedAt == null) contact.LastAssignedAt = now;
            if (contact.FirstSeenOnChat == null) contact.FirstSeenOnChat = now;
            if (contact.LastSeenOnChat == null) contact.LastSeenOnChat = now;
            if (contact.LastSeenOnWeb == null) contact.LastSeenOnWeb = now;
            if (contact.LastContactedTime == null) contact.LastContactedTime = now;

            // EnquiryNo is never accepted from the client on create.
            // It is either system-generated (if requested) or left blank.
            contact.EnquiryNo = generateEnquiryNo
                ? await GenerateNextEnquiryNoAsync()
                : null;

            // EstimatedQuote is entered manually by the user — keep the value from the payload.

            await EnsureAccountIdResolvedAsync(contact);

            // Duplicate Email Check
            // WorkEmail -> WorkEmail + Emails
            // Emails -> WorkEmail + Emails
            

            var emailList = new List<string>();

            if (!string.IsNullOrWhiteSpace(contact.WorkEmail))
            {
                emailList.AddRange(
                    contact.WorkEmail
                        .Split(',', StringSplitOptions.RemoveEmptyEntries)
                        .Select(e => e.Trim().ToLower())
                );
            }

            if (!string.IsNullOrWhiteSpace(contact.Emails))
            {
                emailList.AddRange(
                    contact.Emails
                        .Split(',', StringSplitOptions.RemoveEmptyEntries)
                        .Select(e => e.Trim().ToLower())
                );
            }

            emailList = emailList.Distinct().ToList();

            if (emailList.Any())
            {
                var contacts = await _contactContext.Contacts
                    .Select(c => new
                    {
                        c.WorkEmail,
                        c.Emails
                    })
                    .ToListAsync();

                var duplicateExists = contacts.Any(c =>
                {
                    var existingEmails = new List<string>();

                    if (!string.IsNullOrWhiteSpace(c.WorkEmail))
                    {
                        existingEmails.AddRange(
                            c.WorkEmail
                                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                                .Select(x => x.Trim().ToLower())
                        );
                    }

                    if (!string.IsNullOrWhiteSpace(c.Emails))
                    {
                        existingEmails.AddRange(
                            c.Emails
                                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                                .Select(x => x.Trim().ToLower())
                        );
                    }

                    return existingEmails.Intersect(emailList).Any();
                });

                if (duplicateExists)
                {
                    throw new InvalidOperationException("A contact with this email already exists.");
                }
            }

            // Generate ID only if not provided (manual add).
            // Import provides its own IDs from the Excel file.
            if (contact.ContactId <= 0)
            {
                contact.ContactId =
                    DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() * 1000 +
                    new Random().Next(0, 999);
            }
            else
            {
                // Import: if contact with this ID already exists, update instead
                var existing = await _contactContext.Contacts
                    .FirstOrDefaultAsync(c => c.ContactId == contact.ContactId);

                if (existing != null)
                {
                    return await UpdateAsync(contact.ContactId, contact) ?? contact;
                }
            }

            _contactContext.Contacts.Add(contact);
            await _contactContext.SaveChangesAsync();

            return contact;
        }

        /// <summary>
        /// Copies the editable fields from <paramref name="contact"/> onto the stored contact and refreshes its
        /// activity/updated timestamps. CreatedBy/UpdatedBy fields are intentionally preserved, the account FK is
        /// re-resolved from the account name when needed, and EnquiryNo is taken only from generation (never from
        /// the payload) and then only when requested and currently empty.
        /// </summary>
        /// <param name="id">Primary key of the contact to update.</param>
        /// <param name="contact">Source of the new field values.</param>
        /// <param name="generateEnquiryNo">When true, assigns a sequential EnquiryNo only if the stored contact has none.</param>
        /// <returns>The updated contact, or null when no contact has that ID.</returns>
        // Update contact
        // Update contact
        public async Task<ContactModel?> UpdateAsync(long id, ContactModel contact, bool generateEnquiryNo = false, bool generateEstimatedQuote = false)
        {
            var existing = await _contactContext.Contacts
                .FirstOrDefaultAsync(c => c.ContactId == id);

            if (existing == null)
            {
                return null;
            }

            await EnsureAccountIdResolvedAsync(contact);

            existing.FirstName = contact.FirstName;
            existing.LastName = contact.LastName;
            existing.JobTitle = contact.JobTitle;
            existing.WorkPhone = contact.WorkPhone;
            existing.Mobile = contact.Mobile;
            existing.WorkEmail = contact.WorkEmail;
            existing.Address = contact.Address;
            existing.City = contact.City;
            existing.State = contact.State;
            existing.Zipcode = contact.Zipcode;
            existing.Country = contact.Country;
            existing.LinkedIn = contact.LinkedIn;
            existing.Facebook = contact.Facebook;
            existing.Twitter = contact.Twitter;
            existing.Source = contact.Source;
            existing.Campaign = contact.Campaign;
            existing.Territory = contact.Territory;
            existing.LifeCycleStage = contact.LifeCycleStage;
            existing.Status = contact.Status;
            existing.LostReason = contact.LostReason;
            existing.Tags = contact.Tags;
            existing.Lists = contact.Lists;
            existing.Emails = contact.Emails;
            existing.Products = contact.Products;
            existing.Message = contact.Message;
            existing.AccountId = contact.AccountId;
            existing.Account = contact.Account;
            existing.SalesOwnerId = contact.SalesOwnerId;
            existing.SalesOwner = contact.SalesOwner;
            // existing.CreatedById = contact.CreatedById;         // IMMUTABLE - preserve original
            // existing.CreatedBy = contact.CreatedBy;             // IMMUTABLE - preserve original
            // existing.UpdatedById = contact.UpdatedById;         // Backend sets from current user
            // existing.UpdatedBy = contact.UpdatedBy;             // Backend sets from current user
            existing.ImportID = contact.ImportID;

            var now = DateTime.UtcNow;
            existing.LastActivityDate = now;
            existing.LastSeenOnChat = now;
            existing.LastSeenOnWeb = now;
            existing.LastContactedTime = now;
            existing.UpdatedAt = now;

            // EnquiryNo is only touched on update when explicitly requested via the
            // "Generate Enquiry Number" checkbox. Otherwise the existing value is
            // preserved as-is — we deliberately do NOT take it from the incoming
            // `contact` payload, since callers may omit/blank it unintentionally.
            if (generateEnquiryNo && string.IsNullOrWhiteSpace(existing.EnquiryNo))
            {
                existing.EnquiryNo = await GenerateNextEnquiryNoAsync();
            }

            // EstimatedQuote is entered manually — save the value the user typed (only when
            // provided, so editing other fields never blanks an existing quote). Editing an
            // existing quote goes through here; CLEARING one does not — a blank here is treated
            // as "not supplied", so removal has its own Admin-only endpoint
            // (see ClearEstimatedQuoteAsync).
            if (!string.IsNullOrWhiteSpace(contact.EstimatedQuote))
            {
                existing.EstimatedQuote = contact.EstimatedQuote;
            }

            await _contactContext.SaveChangesAsync();
            return existing;
        }

        /// <summary>
        /// Removes a contact's EstimatedQuote. Separate from <see cref="UpdateAsync"/> because that
        /// method treats a blank quote as "field not supplied" and so can never clear one — and
        /// because deleting a quote is Admin-only, which the controller enforces.
        /// </summary>
        /// <param name="contactId">Primary key of the contact whose quote is being removed.</param>
        /// <returns>True when the contact existed and its quote was cleared; false when no such contact.</returns>
        public async Task<bool> ClearEstimatedQuoteAsync(long contactId)
        {
            var existing = await _contactContext.Contacts.FindAsync(contactId);
            if (existing == null) return false;

            existing.EstimatedQuote = null;
            await _contactContext.SaveChangesAsync();
            return true;
        }

        /// <summary>
        /// Deletes a contact, first removing its related call logs, tasks and notes to satisfy FK constraints.
        /// If that cascade fails for any reason, it falls back to removing just the contact row.
        /// </summary>
        /// <param name="id">Primary key of the contact to delete.</param>
        /// <returns>True when a contact was deleted; false when no contact has that ID.</returns>
        // Delete contact
        public async Task<bool> DeleteAsync(long id, string deletedBy)
        {
            var contact = await _contactContext.Contacts
                .FirstOrDefaultAsync(c => c.ContactId == id);

            if (contact == null)
            {
                return false;
            }

            var relatedCallLogs = await _contactContext.CallLog
                .Where(cl => cl.ContactId == id)
                .ToListAsync();

            var relatedTasks = await _contactContext.Tasks
                .Where(t => t.ContactId == id)
                .ToListAsync();

            var relatedNotes = await _contactContext.Notes
                .Where(n => n.ContactId == id)
                .ToListAsync();

            if (relatedCallLogs.Any())
            {
                _contactContext.CallLog.RemoveRange(relatedCallLogs);
            }

            if (relatedTasks.Any())
            {
                _contactContext.Tasks.RemoveRange(relatedTasks);
            }

            if (relatedNotes.Any())
            {
                _contactContext.Notes.RemoveRange(relatedNotes);
            }

            _contactContext.Contacts.Remove(contact);
            await _contactContext.SaveChangesAsync();
            return true;
        }
        
        /// <summary>
        /// Returns all contacts created within the calendar day of <paramref name="createdAt"/>; the time
        /// component is dropped and matching spans from midnight up to (but not including) the next midnight.
        /// </summary>
        /// <param name="createdAt">Day to match; only the date part is used.</param>
        /// <returns>Contacts created on that day, or an empty list when none match.</returns>
            // Get contacts by CreatedAt date
        public async Task<List<ContactModel>> GetContactsByCreatedAtAsync(DateTime createdAt)
        {
            var start = createdAt.Date;
            var end = start.AddDays(1);

            return await _contactContext.Contacts
                .Where(c => c.CreatedAt >= start && c.CreatedAt < end)
                .ToListAsync();
        }

        /// <summary>
        /// Returns the enquiry numbers for the given contact IDs, omitting contacts whose EnquiryNo is null or
        /// blank, so the result may contain fewer entries than the input list.
        /// </summary>
        /// <param name="contactIds">Contact IDs to read enquiry numbers from.</param>
        /// <returns>The non-blank enquiry numbers found among those contacts.</returns>
        public async Task<List<string>> GetEnquiryNumbersAsync(List<long> contactIds)
        {
            return await _contactContext.Contacts
                .Where(c => contactIds.Contains(c.ContactId))
                .Select(c => c.EnquiryNo)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x!)
                .ToListAsync();
        }

        /// <summary>
        /// Returns the estimated-quote numbers for the given contact IDs, omitting contacts whose
        /// EstimatedQuote is null or blank (so the result may be shorter than the input list).
        /// </summary>
        /// <param name="contactIds">Contact IDs to read estimated-quote numbers from.</param>
        /// <returns>The non-blank estimated-quote numbers found among those contacts.</returns>
        public async Task<List<string>> GetEstimatedQuotesAsync(List<long> contactIds)
        {
            return await _contactContext.Contacts
                .Where(c => contactIds.Contains(c.ContactId))
                .Select(c => c.EstimatedQuote)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x!)
                .ToListAsync();
        }
    }
}