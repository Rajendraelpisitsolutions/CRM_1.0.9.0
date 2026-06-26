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
    /// Handles persistence and retrieval of call log records.
    /// </summary>
    public class CallLogService
    {
        private readonly AppDbContext _callLogContext;

        /// <summary>
        /// Initializes a new instance of <see cref="CallLogService"/>.
        /// </summary>
        /// <param name="callLogContext">Database context used to read and write call logs.</param>
        public CallLogService(AppDbContext callLogContext)
        {
            _callLogContext = callLogContext;
        }

        /// <summary>
        /// Inserts a new call log, defaulting <c>CreatedAt</c> to the current UTC time only when the caller left it null.
        /// </summary>
        /// <param name="callLog">The call log to persist.</param>
        /// <returns>The same instance after saving, now carrying its database-generated ID.</returns>
        // Create Call Log
        public async Task<CallLogModel> CreateCallLogAsync(CallLogModel callLog)
        {
            // Only set CreatedAt to current time if not provided by the client
            if (callLog.CreatedAt == null)
            {
                callLog.CreatedAt = DateTime.UtcNow;
            }

            _callLogContext.CallLog.Add(callLog);
            await _callLogContext.SaveChangesAsync();

            return callLog;
        }

        /// <summary>
        /// Loads every call log from the database in no particular order.
        /// </summary>
        /// <returns>All call logs; an empty list when the table holds none.</returns>
        // Get All Call Logs
        public async Task<List<CallLogModel>> GetAllCallLogsAsync()
        {
            return await _callLogContext.CallLog.ToListAsync();
        }

        /// <summary>
        /// Fetches a single call log by primary key.
        /// </summary>
        /// <param name="id">Primary key of the call log.</param>
        /// <returns>The matching call log, or <c>null</c> when no row has that key.</returns>
        // Get CallLog By Id
        public async Task<CallLogModel?> GetCallLogByIdAsync(long id)
        {
            return await _callLogContext.CallLog
                .FirstOrDefaultAsync(c => c.CallLogId == id);
        }

        /// <summary>
        /// Returns the call logs for one contact, ordered newest first by <c>CreatedAt</c>.
        /// </summary>
        /// <param name="contactId">Identifier of the contact to match on.</param>
        /// <returns>The contact's call logs in descending creation order; empty when there are none.</returns>
        // Get CallLogs By ContactId
        public async Task<List<CallLogModel>> GetCallLogsByContactIdAsync(long contactId)
        {
            return await _callLogContext.CallLog
                .Where(c => c.ContactId == contactId)
                .OrderByDescending(c => c.CreatedAt)
                .ToListAsync();
        }

        /// <summary>
        /// Copies the editable call fields (owner, direction, status, duration, outcome, phone, type, notes, association)
        /// and the contact/account/deal links onto the stored record identified by <c>CallLogId</c>, then saves.
        /// Fields such as <c>CreatedAt</c> are left as they are.
        /// </summary>
        /// <param name="callLog">Source record; its <c>CallLogId</c> selects the row and its other fields supply the new values.</param>
        /// <returns>The saved record, or <c>null</c> when no call log matches the given <c>CallLogId</c>.</returns>
        // Update CallLog
        public async Task<CallLogModel?> UpdateCallLogAsync(CallLogModel callLog)
        {
            var existing = await _callLogContext.CallLog
                .FirstOrDefaultAsync(c => c.CallLogId == callLog.CallLogId);

            if (existing == null)
            {
                return null;
            }

            existing.CallOwner = callLog.CallOwner;
            existing.CallDirection = callLog.CallDirection;
            existing.CallStatus = callLog.CallStatus;
            existing.CallDuration = callLog.CallDuration;
            existing.Outcome = callLog.Outcome;
            existing.Phone = callLog.Phone;
            existing.CallType = callLog.CallType;
            existing.Notes = callLog.Notes;
            existing.AssociatedWithCall = callLog.AssociatedWithCall;
 
            existing.ContactId = callLog.ContactId;
            existing.AccountId = callLog.AccountId;
            existing.DealId = callLog.DealId;

            await _callLogContext.SaveChangesAsync();

            return existing;
        }

        /// <summary>
        /// Removes the call log with the given ID from the database, if it exists.
        /// </summary>
        /// <param name="id">Primary key of the call log to delete.</param>
        /// <returns><c>true</c> once the row is deleted; <c>false</c> when no such call log was found.</returns>
        // Delete CallLog
        public async Task<bool> DeleteCallLogAsync(long id)
        {
            var existing = await _callLogContext.CallLog
                .FirstOrDefaultAsync(c => c.CallLogId == id);

            if (existing == null)
            {
                return false;
            }

            _callLogContext.CallLog.Remove(existing);
            await _callLogContext.SaveChangesAsync();

            return true;
        }
    }
}