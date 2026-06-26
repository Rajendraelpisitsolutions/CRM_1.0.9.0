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
    /// Service for managing appointments in the CRM system.
    /// </summary>
    public class AppointmentsService
    {
        private readonly AppDbContext _appointmentsDb;

        /// <summary>
        /// Initializes a new instance of <see cref="AppointmentsService"/>.
        /// </summary>
        /// <param name="appointmentsDb">Database context for appointments.</param>
        public AppointmentsService(AppDbContext appointmentsDb)
        {
            _appointmentsDb = appointmentsDb;
        }

        /// <summary>
        /// Loads every appointment from the database in no particular order.
        /// </summary>
        /// <returns>All appointments; an empty list when the table holds none.</returns>
        public async Task<List<AppointmentsModel>> GetAllAppointmentsAsync()
        {
            return await _appointmentsDb.Appointments.ToListAsync();
        }

        /// <summary>
        /// Fetches a single appointment by primary key.
        /// </summary>
        /// <param name="appointmentId">Primary key of the appointment.</param>
        /// <returns>The matching appointment, or <c>null</c> when no row has that key.</returns>
        public async Task<AppointmentsModel?> GetAppointmentByIdAsync(int appointmentId)
        {
            return await _appointmentsDb.Appointments.FindAsync(appointmentId);
        }

        /// <summary>
        /// Filters appointments down to those belonging to one contact.
        /// </summary>
        /// <param name="contactId">Identifier of the contact to match on.</param>
        /// <returns>Appointments whose <c>ContactId</c> equals the argument; empty when there are none.</returns>
        public async Task<List<AppointmentsModel>> GetAppointmentsByContactIdAsync(int contactId)
        {
            return await _appointmentsDb.Appointments
                .Where(a => a.ContactId == contactId)
                .ToListAsync();
        }

        /// <summary>
        /// Filters appointments down to those belonging to one account.
        /// </summary>
        /// <param name="accountId">Identifier of the account to match on.</param>
        /// <returns>Appointments whose <c>AccountId</c> equals the argument; empty when there are none.</returns>
        public async Task<List<AppointmentsModel>> GetAppointmentsByAccountIdAsync(int accountId)
        {
            return await _appointmentsDb.Appointments
                .Where(a => a.AccountId == accountId)
                .ToListAsync();
        }

        /// <summary>
        /// Persists a new appointment, stamping both <c>CreatedAt</c> and <c>UpdatedAt</c> with the current UTC time before saving.
        /// </summary>
        /// <param name="appointment">The appointment to insert.</param>
        /// <returns>The same instance after saving, now populated with its database-generated ID.</returns>
        public async Task<AppointmentsModel> CreateAppointmentAsync(AppointmentsModel appointment)
        {
            appointment.CreatedAt = DateTime.UtcNow;
            appointment.UpdatedAt = DateTime.UtcNow;

            _appointmentsDb.Appointments.Add(appointment);
            await _appointmentsDb.SaveChangesAsync();
            return appointment;
        }

        /// <summary>
        /// Copies a fixed set of editable fields (owner, contact details, deal amount, tags, emails, account/contact links, etc.)
        /// onto the stored appointment and refreshes <c>UpdatedAt</c> to the current UTC time. Other columns are left untouched.
        /// </summary>
        /// <param name="appointmentId">Primary key of the appointment to update.</param>
        /// <param name="appointment">Source object whose editable fields are copied across.</param>
        /// <returns>The saved appointment, or <c>null</c> when no row matches the ID.</returns>
        public async Task<AppointmentsModel?> UpdateAppointmentAsync(int appointmentId, AppointmentsModel appointment)
        {
            var existing = await _appointmentsDb.Appointments.FindAsync(appointmentId);

            if (existing == null)
            {
                return null;
            }

            existing.MeetingOwner = appointment.MeetingOwner;
            existing.UpdatedBy = appointment.UpdatedBy;
            existing.ContactEmailStatus = appointment.ContactEmailStatus;
            existing.ContactJobTitle = appointment.ContactJobTitle;
            existing.ContactAddress = appointment.ContactAddress;
            existing.OpenDealsAmount = appointment.OpenDealsAmount;
            existing.FirstContacted = appointment.FirstContacted;
            existing.FirstAssignedAt = appointment.FirstAssignedAt;
            existing.Tags = appointment.Tags;
            existing.EmailIDs = appointment.EmailIDs;
            existing.UpdatedAt = DateTime.UtcNow;
            existing.AccountId = appointment.AccountId;
            existing.ContactId = appointment.ContactId;

            await _appointmentsDb.SaveChangesAsync();
            return existing;
        }

        /// <summary>
        /// Removes the appointment with the given ID from the database, if it exists.
        /// </summary>
        /// <param name="appointmentId">Primary key of the appointment to delete.</param>
        /// <returns><c>true</c> once the row is deleted; <c>false</c> when no such appointment was found.</returns>
        public async Task<bool> DeleteAppointmentAsync(int appointmentId)
        {
            var appointment = await _appointmentsDb.Appointments.FindAsync(appointmentId);

            if (appointment == null)
            {
                return false;
            }

            _appointmentsDb.Appointments.Remove(appointment);
            await _appointmentsDb.SaveChangesAsync();
            return true;
        }
    }
}
