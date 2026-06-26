using Elpis_CRM.Data;
using Elpis_CRM.Model;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Elpis_CRM.Services
{
    /// <summary>
    /// Data-access layer for accounts: querying, searching, tag handling, create/update, and the
    /// cascade-aware deletes (related Deals are removed first to respect the foreign key).
    /// </summary>
    public class AccountService
    {
        private readonly AppDbContext _accountDb;

        /// <summary>
        /// Creates the service over the given EF Core database context.
        /// </summary>
        /// <param name="accountDb">Application database context exposing Accounts and Deals.</param>
        public AccountService(AppDbContext accountDb)
        {
            _accountDb = accountDb;
        }

        /// <summary>
        /// Returns one page of accounts ordered by CreatedAt descending (nulls last), with an optional
        /// token-based search where every whitespace-separated token must match one of the searchable
        /// fields (AND across tokens). Searches shorter than two characters are ignored. The query runs
        /// no-tracking; page is floored to 1 and pageSize clamped to 1-500.
        /// </summary>
        /// <param name="page">1-based page number; values below 1 become 1.</param>
        /// <param name="pageSize">Rows per page, clamped to 1-500.</param>
        /// <param name="search">Optional space-separated search text matched against name, phone, display phone, website, city, country, sales owner, territory, tags and industry.</param>
        /// <returns>A tuple of the page's Items and the TotalCount across all matches (before paging).</returns>
        public async Task<(List<AccountModel> Items, int TotalCount)> GetAllAsync(
            int page = 1, int pageSize = 150, string? search = null)
        {
            page     = Math.Max(page, 1);
            pageSize = Math.Clamp(pageSize, 1, 500);

            var query = _accountDb.Accounts.AsNoTracking();

            // Token-based search: each word must match at least one field (AND across tokens).
            if (!string.IsNullOrWhiteSpace(search) && search.Trim().Length >= 2)
            {
                var tokens = search.Trim().ToLowerInvariant()
                    .Split(' ', StringSplitOptions.RemoveEmptyEntries);
                foreach (var tok in tokens)
                {
                    var t = tok;
                    query = query.Where(a =>
                        (a.Name        != null && a.Name.ToLower().Contains(t))        ||
                        (a.Phone       != null && a.Phone.ToLower().Contains(t))       ||
                        (a.DisplayPhone!= null && a.DisplayPhone.ToLower().Contains(t))||
                        (a.Website     != null && a.Website.ToLower().Contains(t))     ||
                        (a.City        != null && a.City.ToLower().Contains(t))        ||
                        (a.Country     != null && a.Country.ToLower().Contains(t))     ||
                        (a.SalesOwner  != null && a.SalesOwner.ToLower().Contains(t))  ||
                        (a.Territory   != null && a.Territory.ToLower().Contains(t))   ||
                        (a.Tags        != null && a.Tags.ToLower().Contains(t))        ||
                        (a.IndustryType!= null && a.IndustryType.ToLower().Contains(t)));
                }
            }

            var totalCount = await query.CountAsync();

            var items = await query
                .OrderByDescending(a => a.CreatedAt ?? DateTime.MinValue)
                .Skip((page - 1) * pageSize)
                .Take(pageSize)
                .ToListAsync();

            return (items, totalCount);
        }

        /// <summary>
        /// Typeahead search returning accounts ordered by name, capped at <paramref name="limit"/> rows.
        /// A purely numeric query is treated as an exact AccountId match (in addition to name/phone/website);
        /// otherwise each token must match a searchable field. Runs no-tracking and returns an empty list
        /// when the trimmed query is under two characters.
        /// </summary>
        /// <param name="q">Search text; must be at least two characters after trimming.</param>
        /// <param name="limit">Maximum rows to return, clamped to 1-100.</param>
        /// <returns>Matching accounts, name-ascending; empty when the query is too short.</returns>
        public async Task<List<AccountModel>> SearchAsync(string? q, int limit = 50)
        {
            if (string.IsNullOrWhiteSpace(q) || q.Trim().Length < 2)
                return new List<AccountModel>();

            limit = Math.Clamp(limit, 1, 100);
            var term = q.Trim().ToLowerInvariant();
            var isNumeric = long.TryParse(term, out var idVal);

            var query = _accountDb.Accounts.AsNoTracking().AsQueryable();

            if (isNumeric)
            {
                query = query.Where(a =>
                    a.AccountId == idVal ||
                    (a.Name != null && a.Name.ToLower().Contains(term)) ||
                    (a.Phone != null && a.Phone.ToLower().Contains(term)) ||
                    (a.DisplayPhone != null && a.DisplayPhone.ToLower().Contains(term)) ||
                    (a.Website != null && a.Website.ToLower().Contains(term)));
            }
            else
            {
                var tokens = term.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                foreach (var tok in tokens)
                {
                    var t = tok;
                    query = query.Where(a =>
                        (a.Name        != null && a.Name.ToLower().Contains(t))        ||
                        (a.Phone       != null && a.Phone.ToLower().Contains(t))       ||
                        (a.DisplayPhone!= null && a.DisplayPhone.ToLower().Contains(t))||
                        (a.Website     != null && a.Website.ToLower().Contains(t))     ||
                        (a.Territory   != null && a.Territory.ToLower().Contains(t))   ||
                        (a.SalesOwner  != null && a.SalesOwner.ToLower().Contains(t))  ||
                        (a.IndustryType!= null && a.IndustryType.ToLower().Contains(t)));
                }
            }

            return await query
                .OrderBy(a => a.Name)
                .Take(limit)
                .ToListAsync();
        }

        /// <summary>
        /// Collects every distinct tag across all accounts. Each account stores its tags as a single
        /// comma-separated string; these are split, trimmed and de-duplicated into a flat list.
        /// </summary>
        /// <returns>The unique tags, each appearing once (unsorted).</returns>
        public async Task<List<string>> GetAllTagsAsync()
        {
            var tags = await _accountDb.Accounts
                .Where(a => !string.IsNullOrEmpty(a.Tags))
                .Select(a => a.Tags)
                .ToListAsync();

            return tags
                .SelectMany(t => t.Split(',', StringSplitOptions.RemoveEmptyEntries))
                .Select(t => t.Trim())
                .Distinct()
                .ToList();
        }

        /// <summary>
        /// Returns accounts that carry any of the requested tags (OR semantics). The requested string and
        /// each account's tag list are split on commas and trimmed, then matched exactly. Note the tag
        /// comparison is performed in memory after loading all tagged accounts.
        /// </summary>
        /// <param name="tags">Comma-separated tags to match.</param>
        /// <returns>Accounts having at least one of the supplied tags.</returns>
        public async Task<List<AccountModel>> GetAccountsByTagsAsync(string tags)
        {
            var selectedTags = tags.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                   .Select(t => t.Trim())
                                   .ToList();

            var accounts = await _accountDb.Accounts
                .Where(a => !string.IsNullOrEmpty(a.Tags))
                .ToListAsync();

            return accounts
                .Where(a => a.Tags.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                  .Select(t => t.Trim())
                                  .Any(tag => selectedTags.Contains(tag)))
                .ToList();
        }

        /// <summary>
        /// Loads a single tracked account by its identifier.
        /// </summary>
        /// <param name="accountId">AccountId to look up.</param>
        /// <returns>The matching account, or null if none exists.</returns>
        public async Task<AccountModel?> GetByIdAsync(long accountId)
        {
            return await _accountDb.Accounts
                .FirstOrDefaultAsync(a => a.AccountId == accountId);
        }

        /// <summary>
        /// Persists an account, covering both manual entry and import. When AccountId is non-positive a
        /// time-based unique id is generated; when a positive id is supplied that already exists, the call
        /// is redirected to <see cref="UpdateAsync"/> (upsert). CreatedAt/UpdatedAt are set to now and the
        /// LastContacted/LastActivity timestamps default to now when not provided.
        /// </summary>
        /// <param name="account">Account to insert or, on id collision, update.</param>
        /// <returns>The saved account, including its (possibly generated) id.</returns>
        public async Task<AccountModel> AddAsync(AccountModel account)
        {
            var now = DateTime.UtcNow;

            // Always ensure AccountId is set — never allow NULL insert
            if (account.AccountId <= 0)
            {
                account.AccountId = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() * 1000 + new Random().Next(0, 999);
            }
            else
            {
                // Imported account with pre-defined ID — check if it already exists
                var existingAccount = await _accountDb.Accounts
                    .FirstOrDefaultAsync(a => a.AccountId == account.AccountId);

                if (existingAccount != null)
                {
                    return await UpdateAsync(account.AccountId, account);
                }
            }

            account.CreatedAt = now;
            account.UpdatedAt = now;
            account.LastContactedTime = account.LastContactedTime ?? now;
            account.LastActivityDate = account.LastActivityDate ?? now;

            _accountDb.Accounts.Add(account);
            await _accountDb.SaveChangesAsync();

            return account;
        }

        /// <summary>
        /// Imports a batch of accounts that already carry their own ids, upserting each: existing ids have
        /// their fields overwritten, new ids are inserted, and rows with a non-positive id are recorded as
        /// failures. Per-row exceptions are captured rather than aborting the batch; all changes are saved
        /// once at the end, so a final SaveChanges failure can still roll back inserts/updates.
        /// </summary>
        /// <param name="accounts">Accounts to import; each must have a positive AccountId to be processed.</param>
        /// <returns>A summary with inserted, updated and failed records plus success/failure counts.</returns>
        public async Task<BulkImportResult> BulkImportAsync(List<AccountModel> accounts)
        {
            var result = new BulkImportResult
            {
                TotalRecords = accounts.Count,
                InsertedRecords = new List<AccountModel>(),
                UpdatedRecords = new List<AccountModel>(),
                FailedRecords = new List<ImportFailure>()
            };

            var now = DateTime.UtcNow;

            foreach (var account in accounts)
            {
                try
                {
                    // Validate account has an ID
                    if (account.AccountId <= 0)
                    {
                        result.FailedRecords.Add(new ImportFailure
                        {
                            AccountId = account.AccountId,
                            Name = account.Name,
                            Reason = "Invalid or missing AccountId for import"
                        });
                        continue;
                    }

                    // Check if account already exists
                    var existingAccount = await _accountDb.Accounts
                        .FirstOrDefaultAsync(a => a.AccountId == account.AccountId);

                    if (existingAccount != null)
                    {
                        // Update existing account
                        existingAccount.Name = account.Name;
                        existingAccount.NumberOfEmployees = account.NumberOfEmployees;
                        existingAccount.AnnualRevenue = account.AnnualRevenue;
                        existingAccount.Website = account.Website;
                        existingAccount.Phone = account.Phone;
                        existingAccount.DisplayPhone = account.DisplayPhone;
                        existingAccount.Address = account.Address;
                        existingAccount.City = account.City;
                        existingAccount.State = account.State;
                        existingAccount.Zipcode = account.Zipcode;
                        existingAccount.Country = account.Country;
                        existingAccount.Facebook = account.Facebook;
                        existingAccount.Twitter = account.Twitter;
                        existingAccount.LinkedIn = account.LinkedIn;
                        existingAccount.LastContactedMode = account.LastContactedMode;
                        existingAccount.LastContactedTime = account.LastContactedTime ?? now;
                        existingAccount.LastActivityType = account.LastActivityType;
                        existingAccount.LastActivityDate = account.LastActivityDate ?? now;
                        existingAccount.RecentNote = account.RecentNote;
                        existingAccount.IndustryType = account.IndustryType;
                        existingAccount.BusinessType = account.BusinessType;
                        existingAccount.Territory = account.Territory;
                        existingAccount.ActiveSalesSequences = account.ActiveSalesSequences;
                        existingAccount.CompletedSalesSequences = account.CompletedSalesSequences;
                        existingAccount.ParentAccountId = account.ParentAccountId;
                        existingAccount.ParentAccount = account.ParentAccount;
                        existingAccount.SalesOwnerId = account.SalesOwnerId;
                        existingAccount.SalesOwner = account.SalesOwner;
                        existingAccount.ImportID = account.ImportID;
                        existingAccount.Tags = account.Tags;
                        existingAccount.LastAssignedAt = account.LastAssignedAt;

                        existingAccount.UpdatedAt = now;

                        result.UpdatedRecords.Add(existingAccount);
                    }
                    else
                    {
                        // Insert new account with provided ID
                        account.CreatedAt = now;
                        account.UpdatedAt = now;
                        account.LastContactedTime = account.LastContactedTime ?? now;
                        account.LastActivityDate = account.LastActivityDate ?? now;

                        _accountDb.Accounts.Add(account);
                        result.InsertedRecords.Add(account);
                    }
                }
                catch (Exception ex)
                {
                    result.FailedRecords.Add(new ImportFailure
                    {
                        AccountId = account.AccountId,
                        Name = account.Name,
                        Reason = ex.Message
                    });
                }
            }

            // Save all changes
            await _accountDb.SaveChangesAsync();

            result.SuccessCount = result.InsertedRecords.Count + result.UpdatedRecords.Count;
            result.FailureCount = result.FailedRecords.Count;

            return result;
        }

        /// <summary>
        /// Returns the accounts created on a given calendar day, matching CreatedAt within the
        /// [date, date+1day) window so the time component of <paramref name="createdAt"/> is ignored.
        /// Runs no-tracking.
        /// </summary>
        /// <param name="createdAt">Day to filter on; only the date part is used.</param>
        /// <returns>Accounts created that day; empty when none match.</returns>
        public async Task<List<AccountModel>> GetAccountsByCreatedAtAsync(DateTime createdAt)
        {
            var start = createdAt.Date;
            var end = start.AddDays(1);

            return await _accountDb.Accounts
                .AsNoTracking()
                .Where(a => a.CreatedAt >= start && a.CreatedAt < end)
                .ToListAsync();
        }
        /// <summary>
        /// Copies the mutable fields from <paramref name="updatedAccount"/> onto the stored record and
        /// refreshes UpdatedAt. AccountId and CreatedAt are preserved; the record is matched by
        /// <paramref name="accountId"/>, not by any id on the incoming model.
        /// </summary>
        /// <param name="accountId">AccountId of the record to update.</param>
        /// <param name="updatedAccount">Source of the new field values.</param>
        /// <returns>The updated account, or null if no record has that id.</returns>
        public async Task<AccountModel?> UpdateAsync(long accountId, AccountModel updatedAccount)
        {
            var existing = await _accountDb.Accounts
                .FirstOrDefaultAsync(a => a.AccountId == accountId);

            if (existing == null)
            {
                return null;
            }

            existing.Name = updatedAccount.Name;
            existing.NumberOfEmployees = updatedAccount.NumberOfEmployees;
            existing.AnnualRevenue = updatedAccount.AnnualRevenue;
            existing.Website = updatedAccount.Website;
            existing.Phone = updatedAccount.Phone;
            existing.DisplayPhone = updatedAccount.DisplayPhone;
            existing.Address = updatedAccount.Address;
            existing.City = updatedAccount.City;
            existing.State = updatedAccount.State;
            existing.Zipcode = updatedAccount.Zipcode;
            existing.Country = updatedAccount.Country;
            existing.Facebook = updatedAccount.Facebook;
            existing.Twitter = updatedAccount.Twitter;
            existing.LinkedIn = updatedAccount.LinkedIn;
            existing.LastContactedMode = updatedAccount.LastContactedMode;
            existing.LastContactedTime = updatedAccount.LastContactedTime;
            existing.LastActivityType = updatedAccount.LastActivityType;
            existing.LastActivityDate = updatedAccount.LastActivityDate;
            existing.RecentNote = updatedAccount.RecentNote;
            existing.IndustryType = updatedAccount.IndustryType;
            existing.BusinessType = updatedAccount.BusinessType;
            existing.Territory = updatedAccount.Territory;
            existing.ActiveSalesSequences = updatedAccount.ActiveSalesSequences;
            existing.CompletedSalesSequences = updatedAccount.CompletedSalesSequences;
            existing.ParentAccountId = updatedAccount.ParentAccountId;
            existing.ParentAccount = updatedAccount.ParentAccount;
            existing.SalesOwnerId = updatedAccount.SalesOwnerId;
            existing.SalesOwner = updatedAccount.SalesOwner;
            existing.ImportID = updatedAccount.ImportID;
            existing.Tags = updatedAccount.Tags;
            existing.LastAssignedAt = updatedAccount.LastAssignedAt;

            var now = DateTime.UtcNow;
            existing.UpdatedAt = now;

            await _accountDb.SaveChangesAsync();
            return existing;
        }

        /// <summary>
        /// Deletes an account, first removing any Deals that reference it to satisfy the foreign-key
        /// constraint (cascading delete).
        /// </summary>
        /// <param name="accountId">AccountId to delete.</param>
        /// <returns>True if the account existed and was deleted; false if no such account.</returns>
        public async Task<bool> DeleteAsync(long accountId)
        {
            var account = await _accountDb.Accounts
                .FirstOrDefaultAsync(a => a.AccountId == accountId);

            if (account == null)
            {
                return false;
            }

            // Handle foreign key constraint: Delete or orphan related Deals first
            var relatedDeals = await _accountDb.Deals
                .Where(d => d.AccountId == accountId)
                .ToListAsync();

            if (relatedDeals.Any())
            {
                // Delete related Deals (cascading delete)
                _accountDb.Deals.RemoveRange(relatedDeals);
            }

            _accountDb.Accounts.Remove(account);
            await _accountDb.SaveChangesAsync();

            return true;
        }

        /// <summary>
        /// Removes every account from the database, deleting all Deals first (in a separate save) to clear
        /// the foreign-key constraint. Any failure is wrapped in an <see cref="InvalidOperationException"/>.
        /// </summary>
        public async Task DeleteAllAsync()
        {
            try
            {
                // First, delete all related Deals to handle foreign key constraints
                var allDeals = await _accountDb.Deals.ToListAsync();
                if (allDeals.Any())
                {
                    _accountDb.Deals.RemoveRange(allDeals);
                    await _accountDb.SaveChangesAsync();
                }

                // Then delete all Accounts
                var allAccounts = await _accountDb.Accounts.ToListAsync();
                if (allAccounts.Any())
                {
                    _accountDb.Accounts.RemoveRange(allAccounts);
                    await _accountDb.SaveChangesAsync();
                }
            }
            catch (Exception ex)
            {
                // Log error or rethrow for controller to handle
                throw new InvalidOperationException("Error deleting all accounts", ex);
            }
        }

        /// <summary>
        /// Deletes several accounts by id, removing their related Deals first (in a separate save) to
        /// satisfy the foreign key. A null or empty id list is a no-op. Failures are wrapped in an
        /// <see cref="InvalidOperationException"/>.
        /// </summary>
        /// <param name="accountIds">Ids of the accounts to delete.</param>
        /// <returns>The number of accounts actually removed (ids with no matching account are skipped).</returns>
        public async Task<int> DeleteMultipleAsync(List<long> accountIds)
        {
            try
            {
                if (accountIds == null || !accountIds.Any())
                {
                    return 0;
                }

                // Delete related Deals for all accounts first
                var relatedDeals = await _accountDb.Deals
                    .Where(d => accountIds.Contains(d.AccountId.Value))
                    .ToListAsync();

                if (relatedDeals.Any())
                {
                    _accountDb.Deals.RemoveRange(relatedDeals);
                    await _accountDb.SaveChangesAsync();
                }

                // Delete all the accounts
                var accountsToDelete = await _accountDb.Accounts
                    .Where(a => accountIds.Contains(a.AccountId))
                    .ToListAsync();

                int countToDelete = accountsToDelete.Count;
                if (countToDelete > 0)
                {
                    _accountDb.Accounts.RemoveRange(accountsToDelete);
                    await _accountDb.SaveChangesAsync();
                }

                return countToDelete;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Error deleting multiple accounts", ex);
            }
        }

        /// <summary>
        /// Produces a candidate AccountId for the UI to preview, derived from the current UTC time in
        /// Unix milliseconds. It is not reserved or persisted, so concurrent callers may receive the same
        /// or overlapping values.
        /// </summary>
        /// <returns>A time-based candidate id (current UTC Unix time in milliseconds).</returns>
        public Task<long> GetNextAccountIdAsync()
        {
            return Task.FromResult(DateTimeOffset.UtcNow.ToUnixTimeMilliseconds());
        }
    }

    // Helper classes for bulk import result
    public class BulkImportResult
    {
        public int TotalRecords { get; set; }
        public int SuccessCount { get; set; }
        public int FailureCount { get; set; }
        public List<AccountModel> InsertedRecords { get; set; } = new List<AccountModel>();
        public List<AccountModel> UpdatedRecords { get; set; } = new List<AccountModel>();
        public List<ImportFailure> FailedRecords { get; set; } = new List<ImportFailure>();
    }

    public class ImportFailure
    {
        public long AccountId { get; set; }
        public string? Name { get; set; }
        public string? Reason { get; set; }
    }
}
