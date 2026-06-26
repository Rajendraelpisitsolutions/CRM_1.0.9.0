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
    /// Handles persistence for meetings against the EF Core <see cref="AppDbContext"/>.
    /// Plain database CRUD — no Microsoft Graph or Teams integration; conferencing details are stored as data only.
    /// </summary>
    public class MeetingService
    {
        private readonly AppDbContext _meetingDb;

        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingService"/> with the application database context.
        /// </summary>
        /// <param name="meetingDb">EF Core context backing the Meeting table.</param>
        public MeetingService(AppDbContext meetingDb)
        {
            _meetingDb = meetingDb;
        }

        /// <summary>
        /// Loads every meeting from the database.
        /// </summary>
        /// <returns>All stored meetings.</returns>
        public async Task<List<MeetingModel>> GetAllMeetingsAsync()
        {
            return await _meetingDb.Meeting.ToListAsync();
        }

        /// <summary>
        /// Looks up a meeting by primary key.
        /// </summary>
        /// <param name="meetingId">Identifier of the meeting.</param>
        /// <returns>The meeting, or null if no row matches.</returns>
        public async Task<MeetingModel?> GetMeetingByIdAsync(int meetingId)
        {
            return await _meetingDb.Meeting.FindAsync(meetingId);
        }

        /// <summary>
        /// Retrieves all meetings whose ContactId matches the supplied value.
        /// </summary>
        /// <param name="contactId">Contact identifier to filter on.</param>
        /// <returns>Matching meetings; an empty list if none are linked to the contact.</returns>
        public async Task<List<MeetingModel>> GetMeetingsByContactIdAsync(long contactId)
        {
            return await _meetingDb.Meeting
                .Where(m => m.ContactId == contactId)
                .ToListAsync();
        }

        /// <summary>
        /// Retrieves all meetings whose AccountId matches the supplied value.
        /// </summary>
        /// <param name="accountId">Account identifier to filter on.</param>
        /// <returns>Matching meetings; an empty list if none are linked to the account.</returns>
        public async Task<List<MeetingModel>> GetMeetingsByAccountIdAsync(long accountId)
        {
            return await _meetingDb.Meeting
                .Where(m => m.AccountId == accountId)
                .ToListAsync();
        }

        /// <summary>
        /// Returns meetings starting at or after the current UTC time, ordered by start time ascending.
        /// </summary>
        /// <returns>Future meetings sorted soonest-first; an empty list if none are scheduled.</returns>
        public async Task<List<MeetingModel>> GetUpcomingMeetingsAsync()
        {
            var now = DateTime.UtcNow;

            return await _meetingDb.Meeting
                .Where(m => m.From >= now)
                .OrderBy(m => m.From)
                .ToListAsync();
        }

        /// <summary>
        /// Inserts a new meeting, setting CreatedAt and UpdatedAt to the current UTC time before saving.
        /// </summary>
        /// <param name="meeting">Meeting to persist.</param>
        /// <returns>The saved meeting, including its database-generated identifier.</returns>
        public async Task<MeetingModel> CreateMeetingAsync(MeetingModel meeting)
        {
            meeting.CreatedAt = DateTime.UtcNow;
            meeting.UpdatedAt = DateTime.UtcNow;

            _meetingDb.Meeting.Add(meeting);
            await _meetingDb.SaveChangesAsync();

            return meeting;
        }

        /// <summary>
        /// Copies the editable fields (title, schedule, location, descriptive notes, and related entity ids)
        /// from the supplied meeting onto the stored row, then refreshes UpdatedAt. CreatedAt is left untouched.
        /// </summary>
        /// <param name="meetingId">Identifier of the meeting to modify.</param>
        /// <param name="meeting">Source of the new field values.</param>
        /// <returns>The updated meeting, or null if no meeting has the given id.</returns>
        public async Task<MeetingModel?> UpdateMeetingAsync(int meetingId, MeetingModel meeting)
        {
            var existing = await _meetingDb.Meeting.FindAsync(meetingId);

            if (existing == null)
                return null;

            existing.Title = meeting.Title;
            existing.From = meeting.From;
            existing.To = meeting.To;
            existing.TimeZone = meeting.TimeZone;
            existing.AddVideoConference = meeting.AddVideoConference;
            existing.Location = meeting.Location;
            existing.Description = meeting.Description;
            existing.Outcome = meeting.Outcome;
            existing.Notes = meeting.Notes;
            existing.RelatedTo = meeting.RelatedTo;
            existing.ContactId = meeting.ContactId;
            existing.AccountId = meeting.AccountId;
            existing.DealId = meeting.DealId;
           
            
            existing.UpdatedAt = DateTime.UtcNow;

            await _meetingDb.SaveChangesAsync();

            return existing;
        }

        /// <summary>
        /// Deletes the meeting with the given id if it exists.
        /// </summary>
        /// <param name="meetingId">Identifier of the meeting to remove.</param>
        /// <returns>True if a meeting was found and deleted; false if none matched.</returns>
        public async Task<bool> DeleteMeetingAsync(int meetingId)
        {
            var meeting = await _meetingDb.Meeting.FindAsync(meetingId);

            if (meeting == null)
                return false;

            _meetingDb.Meeting.Remove(meeting);
            await _meetingDb.SaveChangesAsync();

            return true;
        }
    }
}