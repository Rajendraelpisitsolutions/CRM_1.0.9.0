using Elpis_CRM.Data;
using Elpis_CRM.Model;
using Elpis_CRM.Model.DTOs;
using Microsoft.EntityFrameworkCore;

namespace Elpis_CRM.Services
{
    /// <summary>
    /// Persistence and query logic for notes. A note lives against a contact or a deal; when the
    /// caller supplies mirror ids, the note is also copied to the other side as INDEPENDENT rows,
    /// so later edits or deletes to one copy never propagate to the others.
    /// </summary>
    public class NotesService
    {
        private readonly AppDbContext _notesDb;

        /// <summary>
        /// Creates the service over the given database context.
        /// </summary>
        /// <param name="notesDb">EF Core context backing the Notes table.</param>
        public NotesService(AppDbContext notesDb)
        {
            _notesDb = notesDb;
        }

        // ─────────────────────────────────────────────────────────────────────────
        // ADD
        //
        // New behaviour (replaces the old 1:1 auto-mirroring):
        //
        //  • Contact panel note  → saved ONLY against that contact. No mirroring.
        //
        //  • Deal panel note     → the caller (UI) explicitly says where the note
        //                          should also appear, via note.MirrorToContactIds.
        //                          - Always writes the original Deal note.
        //                          - For every id in MirrorToContactIds, writes an
        //                            INDEPENDENT copy against that contact.
        //                          - Copies are independent rows: editing/deleting
        //                            one later does NOT affect the others.
        //
        // Validation of "at least one destination chosen" is expected to happen in
        // the UI/controller before this is called; this method trusts its input.
        // ─────────────────────────────────────────────────────────────────────────
        /// <summary>
        /// Saves a note and stamps its timestamps. A contact-only note is written once, then an
        /// independent copy is added under each id in MirrorToDealIds; a deal-only note works the
        /// same way against MirrorToContactIds. If both or neither id is set, the note is stored as-is.
        /// Mirror copies are standalone rows and are not linked back to the original.
        /// </summary>
        /// <param name="note">Note to persist; ContactId/DealId decide the panel, and the MirrorTo* lists drive the independent copies.</param>
        /// <returns>The original note as saved, including its generated id and timestamps.</returns>
        public async Task<NotesModel> AddNote(NotesModel note)
        {
            note.CreatedAt = DateTime.UtcNow;
            note.UpdatedAt = DateTime.UtcNow;

            // ── Case 1: Contact panel — always contact-only, never mirrored ──────
            if (note.ContactId.HasValue && !note.DealId.HasValue)
            {
                note.RelatedToType = "Contact";

                _notesDb.Notes.Add(note);
                await _notesDb.SaveChangesAsync();

                var mirrorDealIds = (note.MirrorToDealIds ?? new List<long>())
                    .Distinct()
                    .ToList();

                foreach (var dealId in mirrorDealIds)
                {
                    var copy = new NotesModel
                    {
                        Description = note.Description,
                        CreatedById = note.CreatedById,
                        CreatedAt = note.CreatedAt,
                        UpdatedAt = note.UpdatedAt,
                        DealId = dealId,
                        ContactId = null,
                        RelatedToType = "Deal"
                    };

                    _notesDb.Notes.Add(copy);
                }

                if (mirrorDealIds.Count > 0)
                {
                    await _notesDb.SaveChangesAsync();
                }

                return note;
            }

            // ── Case 2: Deal panel — write the deal note, then independent ───────
            //            copies for every contact the user explicitly checked.
            if (note.DealId.HasValue && !note.ContactId.HasValue)
            {
                note.RelatedToType = "Deal";
                _notesDb.Notes.Add(note);
                await _notesDb.SaveChangesAsync();

                var mirrorContactIds = (note.MirrorToContactIds ?? new List<long>())
                    .Distinct()
                    .ToList();

                foreach (var contactId in mirrorContactIds)
                {
                    var copy = new NotesModel
                    {
                        Description = note.Description,
                        CreatedById = note.CreatedById,
                        CreatedAt = note.CreatedAt,
                        UpdatedAt = note.UpdatedAt,
                        ContactId = contactId,
                        DealId = null,
                        RelatedToType = "Contact",
                    };
                    _notesDb.Notes.Add(copy);
                }

                if (mirrorContactIds.Count > 0)
                {
                    await _notesDb.SaveChangesAsync();
                }

                return note;
            }

            // ── Case 3: Both ids set or neither — save as-is ─────────────────────
            _notesDb.Notes.Add(note);
            await _notesDb.SaveChangesAsync();
            return note;
        }

        // ─────────────────────────────────────────────────────────────────────────
        // UPDATE
        //
        // Copies are independent now, so this only ever touches the single note
        // being edited — no mirror lookup, no propagation.
        // ─────────────────────────────────────────────────────────────────────────
        /// <summary>
        /// Updates the description and author of one note and refreshes its UpdatedAt timestamp.
        /// Affects only the targeted row; mirror copies are independent and untouched.
        /// </summary>
        /// <param name="id">Identifier of the note to update.</param>
        /// <param name="note">Source of the new Description and CreatedById values.</param>
        /// <returns>The updated note, or null if no note has the given id.</returns>
        public async Task<NotesModel?> UpdateNote(int id, NotesModel note)
        {
            var existing = await _notesDb.Notes.FindAsync(id);
            if (existing == null) return null;

            existing.Description = note.Description;
            existing.CreatedById = note.CreatedById;
            existing.UpdatedAt = DateTime.UtcNow;

            await _notesDb.SaveChangesAsync();
            return existing;
        }

        // ─────────────────────────────────────────────────────────────────────────
        // DELETE
        //
        // Independent copies → plain delete of the single row, no mirror cleanup.
        // ─────────────────────────────────────────────────────────────────────────
        /// <summary>
        /// Removes a single note row. Mirror copies are independent, so they are left in place.
        /// </summary>
        /// <param name="id">Identifier of the note to delete.</param>
        /// <returns>True if the note was found and removed; false if no note has the given id.</returns>
        public async Task<bool> DeleteNote(int id, string deletedBy)
        {
            var note = await _notesDb.Notes.FindAsync(id);

            if (note == null)
                return false;

            _notesDb.Notes.Remove(note);

            await _notesDb.SaveChangesAsync();

            return true;
        }

        // ─────────────────────────────────────────────────────────────────────────
        // READ
        // ─────────────────────────────────────────────────────────────────────────
        /// <summary>
        /// Returns all notes ordered from newest to oldest by creation time.
        /// </summary>
        /// <returns>Every note, or an empty list when none exist.</returns>
        public async Task<List<NotesModel>> GetAllNotes()
        {
            return await _notesDb.Notes
                .OrderByDescending(n => n.CreatedAt)
                .ToListAsync();
        }

        /// <summary>
        /// Looks up a single note by its primary key.
        /// </summary>
        /// <param name="id">Identifier of the note.</param>
        /// <returns>The matching note, or null if it is not found.</returns>
        public async Task<NotesModel?> GetById(int id)
        {
            return await _notesDb.Notes.FindAsync(id);
        }

        /// <summary>
        /// Projects the notes belonging to a contact into DTOs, newest first.
        /// </summary>
        /// <param name="contactId">Identifier of the contact to filter by.</param>
        /// <returns>The contact's notes as <see cref="ContactNoteDto"/> items, or an empty list if there are none.</returns>
        public async Task<List<ContactNoteDto>> GetNotesByContactAsync(long contactId)
        {
            return await _notesDb.Notes
                .Where(n => n.ContactId == contactId)
                .OrderByDescending(n => n.CreatedAt)
                .Select(n => new ContactNoteDto
                {
                    Id = n.Id,
                    Description = n.Description,
                    CreatedById = n.CreatedById,
                    CreatedAt = n.CreatedAt,
                    UpdatedAt = n.UpdatedAt,
                    ContactId = n.ContactId!.Value,
                })
                .ToListAsync();
        }

        /// <summary>
        /// Projects the notes belonging to a deal into DTOs, newest first.
        /// </summary>
        /// <param name="dealId">Identifier of the deal to filter by.</param>
        /// <returns>The deal's notes as <see cref="DealNoteDto"/> items, or an empty list if there are none.</returns>
        public async Task<List<DealNoteDto>> GetNotesByDealAsync(long dealId)
        {
            return await _notesDb.Notes
                .Where(n => n.DealId == dealId)
                .OrderByDescending(n => n.CreatedAt)
                .Select(n => new DealNoteDto
                {
                    Id = n.Id,
                    Description = n.Description,
                    CreatedById = n.CreatedById,
                    CreatedAt = n.CreatedAt,
                    UpdatedAt = n.UpdatedAt,
                    DealId = n.DealId!.Value,
                })
                .ToListAsync();
        }
    }
}