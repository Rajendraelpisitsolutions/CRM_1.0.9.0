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

            // ── Case 1: Contact panel — one shared row, associations in NoteTargets ──
            if (note.ContactId.HasValue && !note.DealId.HasValue)
            {
                note.RelatedToType = "Contact";
                note.IsShared = true;

                _notesDb.Notes.Add(note);
                await _notesDb.SaveChangesAsync(); // populates note.Id

                var mirrorDealIds = (note.MirrorToDealIds ?? new List<long>())
                    .Where(id => id > 0)
                    .Distinct()
                    .ToList();

                var targets = new List<NoteTargetModel>
                {
                    new() { NoteId = note.Id, TargetType = "Contact", TargetId = note.ContactId.Value }
                };
                targets.AddRange(mirrorDealIds.Select(dealId => new NoteTargetModel
                {
                    NoteId = note.Id,
                    TargetType = "Deal",
                    TargetId = dealId
                }));

                _notesDb.NoteTargets.AddRange(targets);
                await _notesDb.SaveChangesAsync();
                return note;
            }

            // ── Case 2: Deal panel — one shared row, associations in NoteTargets ──
            if (note.DealId.HasValue && !note.ContactId.HasValue)
            {
                note.RelatedToType = "Deal";
                note.IsShared = true;

                _notesDb.Notes.Add(note);
                await _notesDb.SaveChangesAsync(); // populates note.Id

                var mirrorContactIds = (note.MirrorToContactIds ?? new List<long>())
                    .Where(id => id > 0)
                    .Distinct()
                    .ToList();

                var targets = new List<NoteTargetModel>
                {
                    new() { NoteId = note.Id, TargetType = "Deal", TargetId = note.DealId.Value }
                };
                targets.AddRange(mirrorContactIds.Select(contactId => new NoteTargetModel
                {
                    NoteId = note.Id,
                    TargetType = "Contact",
                    TargetId = contactId
                }));

                _notesDb.NoteTargets.AddRange(targets);
                await _notesDb.SaveChangesAsync();
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

            // 1) Always record the note's OWN record in NoteTargets. For a legacy note this
            //    migrates it into the shared model the first time it is edited; if the link
            //    already exists nothing is added (idempotent).
            if (existing.ContactId.HasValue && existing.ContactId.Value > 0)
                await EnsureTargetAsync(existing.Id, "Contact", existing.ContactId.Value);
            if (existing.DealId.HasValue && existing.DealId.Value > 0)
                await EnsureTargetAsync(existing.Id, "Deal", existing.DealId.Value);

            // 2) Checklist-driven sharing: the update UI shows the related contacts/deals and the
            //    user ticks who this note should also appear under. Each ticked id becomes a
            //    NoteTargets reference here (same shape as create's MirrorTo* on AddNote). This
            //    lets old duplicate-style notes be linked to their relations without new copies,
            //    so the contact/deal notes panels can segregate them. Additive and idempotent —
            //    an already-linked target is skipped and nothing is removed.
            if (note.MirrorToContactIds != null)
                foreach (var cid in note.MirrorToContactIds.Where(x => x > 0).Distinct())
                    await EnsureTargetAsync(existing.Id, "Contact", cid);
            if (note.MirrorToDealIds != null)
                foreach (var did in note.MirrorToDealIds.Where(x => x > 0).Distinct())
                    await EnsureTargetAsync(existing.Id, "Deal", did);

            await _notesDb.SaveChangesAsync();
            return existing;
        }

        /// <summary>
        /// Ensures a single (NoteId, TargetType, TargetId) association exists in NoteTargets.
        /// It checks both the rows already saved and any staged in this unit of work, and only
        /// adds the row when absent, so calling this repeatedly never creates duplicate links.
        /// The new row (if any) is staged on the context; the caller runs SaveChangesAsync.
        /// </summary>
        /// <param name="noteId">The note the association belongs to.</param>
        /// <param name="targetType">"Contact" or "Deal".</param>
        /// <param name="targetId">The Contacts.ContactId or Deals.Id being linked.</param>
        private async Task EnsureTargetAsync(int noteId, string targetType, long targetId)
        {
            // Already staged (added but not yet saved) in this same update?
            bool staged = _notesDb.NoteTargets.Local.Any(t =>
                t.NoteId == noteId && t.TargetType == targetType && t.TargetId == targetId);
            if (staged) return;

            bool exists = await _notesDb.NoteTargets.AnyAsync(t =>
                t.NoteId == noteId && t.TargetType == targetType && t.TargetId == targetId);

            if (!exists)
            {
                _notesDb.NoteTargets.Add(new NoteTargetModel
                {
                    NoteId = noteId,
                    TargetType = targetType,
                    TargetId = targetId
                });
            }
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
            // A note belongs under this contact if it is a legacy/origin contact note
            // (ContactId == contactId) OR a shared note linked to it via NoteTargets.
            var linkedNoteIds = await _notesDb.NoteTargets
                .Where(t => t.TargetType == "Contact" && t.TargetId == contactId)
                .Select(t => t.NoteId)
                .ToListAsync();

            var notes = await _notesDb.Notes
                .Where(n => n.ContactId == contactId || linkedNoteIds.Contains(n.Id))
                .OrderByDescending(n => n.CreatedAt)
                .ToListAsync();

            var noteIds = notes.Select(n => n.Id).ToList();

            // ── IF: new shared notes — associations come from NoteTargets. ───────
            var dealTargets = await _notesDb.NoteTargets
                .Where(t => noteIds.Contains(t.NoteId) && t.TargetType == "Deal")
                .ToListAsync();

            var targetIdsByNote = dealTargets
                .GroupBy(t => t.NoteId)
                .ToDictionary(g => g.Key, g => g.Select(t => t.TargetId).Distinct().ToList());

            // ── ELSE: legacy notes (no NoteTargets) — reconstruct the shared deals
            //          by matching this contact's related deals that have a note with
            //          the same text + timestamp (the old duplicate copies).
            var fallbackIdsByNote = await BuildLegacyDealMatchesAsync(
                contactId,
                notes.Where(n => !targetIdsByNote.ContainsKey(n.Id)).ToList());

            var allDealIds = dealTargets.Select(t => t.TargetId)
                .Concat(fallbackIdsByNote.SelectMany(kv => kv.Value))
                .Distinct()
                .ToList();

            var dealNames = await _notesDb.Deals
                .Where(d => allDealIds.Contains(d.Id))
                .Select(d => new { d.Id, d.Name })
                .ToListAsync();
            var dealNameById = dealNames.ToDictionary(d => d.Id, d => d.Name);

            NoteShareTargetDto DealRef(long id) => new()
            {
                Id = id,
                Name = dealNameById.TryGetValue(id, out var nm) && !string.IsNullOrWhiteSpace(nm)
                    ? nm!
                    : $"Deal {id}"
            };

            return notes.Select(n =>
            {
                var ids = targetIdsByNote.TryGetValue(n.Id, out var t)
                    ? t
                    : (fallbackIdsByNote.TryGetValue(n.Id, out var f) ? f : new List<long>());

                return new ContactNoteDto
                {
                    Id = n.Id,
                    Description = n.Description,
                    CreatedById = n.CreatedById,
                    CreatedAt = n.CreatedAt,
                    UpdatedAt = n.UpdatedAt,
                    ContactId = contactId,
                    SharedWithDeals = ids.Select(DealRef).ToList()
                };
            }).ToList();
        }

        /// <summary>
        /// Projects the notes belonging to a deal into DTOs, newest first.
        /// </summary>
        /// <param name="dealId">Identifier of the deal to filter by.</param>
        /// <returns>The deal's notes as <see cref="DealNoteDto"/> items, or an empty list if there are none.</returns>
        public async Task<List<DealNoteDto>> GetNotesByDealAsync(long dealId)
        {
            // A note belongs under this deal if it is a legacy/origin deal note
            // (DealId == dealId) OR a shared note linked to it via NoteTargets.
            var linkedNoteIds = await _notesDb.NoteTargets
                .Where(t => t.TargetType == "Deal" && t.TargetId == dealId)
                .Select(t => t.NoteId)
                .ToListAsync();

            var notes = await _notesDb.Notes
                .Where(n => n.DealId == dealId || linkedNoteIds.Contains(n.Id))
                .OrderByDescending(n => n.CreatedAt)
                .ToListAsync();

            var noteIds = notes.Select(n => n.Id).ToList();

            // ── IF: new shared notes — associations come from NoteTargets. ───────
            var contactTargets = await _notesDb.NoteTargets
                .Where(t => noteIds.Contains(t.NoteId) && t.TargetType == "Contact")
                .ToListAsync();

            var targetIdsByNote = contactTargets
                .GroupBy(t => t.NoteId)
                .ToDictionary(g => g.Key, g => g.Select(t => t.TargetId).Distinct().ToList());

            // ── ELSE: legacy notes (no NoteTargets) — reconstruct the shared
            //          contacts by matching this deal's linked contacts that have a
            //          note with the same text + timestamp (the old duplicate copies).
            var fallbackIdsByNote = await BuildLegacyContactMatchesAsync(
                dealId,
                notes.Where(n => !targetIdsByNote.ContainsKey(n.Id)).ToList());

            // Resolve every contact id we need a name for (targets + fallback).
            var allContactIds = contactTargets.Select(t => t.TargetId)
                .Concat(fallbackIdsByNote.SelectMany(kv => kv.Value))
                .Distinct()
                .ToList();

            var contacts = await _notesDb.Contacts
                .Where(c => allContactIds.Contains(c.ContactId))
                .Select(c => new { c.ContactId, c.FirstName, c.LastName })
                .ToListAsync();
            var nameById = contacts.ToDictionary(
                c => c.ContactId,
                c => string.Join(" ", new[] { c.FirstName, c.LastName }
                    .Where(s => !string.IsNullOrWhiteSpace(s))));

            NoteShareTargetDto ContactRef(long id) => new()
            {
                Id = id,
                Name = nameById.TryGetValue(id, out var nm) && !string.IsNullOrWhiteSpace(nm)
                    ? nm
                    : $"Contact {id}"
            };

            return notes.Select(n =>
            {
                var ids = targetIdsByNote.TryGetValue(n.Id, out var t)
                    ? t
                    : (fallbackIdsByNote.TryGetValue(n.Id, out var f) ? f : new List<long>());

                return new DealNoteDto
                {
                    Id = n.Id,
                    Description = n.Description,
                    CreatedById = n.CreatedById,
                    CreatedAt = n.CreatedAt,
                    UpdatedAt = n.UpdatedAt,
                    DealId = dealId,
                    SharedWithContacts = ids.Select(ContactRef).ToList()
                };
            }).ToList();
        }

        // ─────────────────────────────────────────────────────────────────────────
        // LEGACY FALLBACK HELPERS
        //
        // Before the shared-note model, sending a deal note to contacts created an
        // independent copy per contact (same Description + CreatedAt). These helpers
        // recover that link for old data by matching text+timestamp, scoped only to
        // the records actually related to the deal/contact — never a global search.
        // ─────────────────────────────────────────────────────────────────────────

        /// <summary>
        /// For legacy deal notes, finds which of the deal's linked contacts hold a
        /// note copy with the same Description and CreatedAt. Returns note id → contact ids.
        /// </summary>
        private async Task<Dictionary<int, List<long>>> BuildLegacyContactMatchesAsync(
            long dealId, List<NotesModel> legacyNotes)
        {
            var result = new Dictionary<int, List<long>>();
            if (legacyNotes.Count == 0) return result;

            // Contacts linked to this deal: join table + legacy direct ContactId.
            var linkedContactIds = await _notesDb.DealContactLinks
                .Where(l => l.DealId == dealId)
                .Select(l => l.ContactId)
                .ToListAsync();
            var directContactIds = await _notesDb.Deals
                .Where(d => d.Id == dealId && d.ContactId != null)
                .Select(d => d.ContactId!.Value)
                .ToListAsync();
            var candidateContactIds = linkedContactIds.Concat(directContactIds).Distinct().ToList();
            if (candidateContactIds.Count == 0) return result;

            // Those contacts' notes, used as the match pool.
            var contactNotes = await _notesDb.Notes
                .Where(n => n.ContactId != null && candidateContactIds.Contains(n.ContactId.Value))
                .Select(n => new { ContactId = n.ContactId!.Value, n.Description, n.CreatedAt })
                .ToListAsync();

            foreach (var ln in legacyNotes)
            {
                if (string.IsNullOrWhiteSpace(ln.Description)) continue; // avoid matching blanks
                var matches = contactNotes
                    .Where(cn => cn.Description == ln.Description && cn.CreatedAt == ln.CreatedAt)
                    .Select(cn => cn.ContactId)
                    .Distinct()
                    .ToList();
                if (matches.Count > 0) result[ln.Id] = matches;
            }

            return result;
        }

        /// <summary>
        /// For legacy contact notes, finds which of the contact's related deals hold a
        /// note copy with the same Description and CreatedAt. Returns note id → deal ids.
        /// </summary>
        private async Task<Dictionary<int, List<long>>> BuildLegacyDealMatchesAsync(
            long contactId, List<NotesModel> legacyNotes)
        {
            var result = new Dictionary<int, List<long>>();
            if (legacyNotes.Count == 0) return result;

            // Deals linked to this contact: join table + legacy direct ContactId.
            var linkedDealIds = await _notesDb.DealContactLinks
                .Where(l => l.ContactId == contactId)
                .Select(l => l.DealId)
                .ToListAsync();
            var directDealIds = await _notesDb.Deals
                .Where(d => d.ContactId == contactId)
                .Select(d => d.Id)
                .ToListAsync();
            var candidateDealIds = linkedDealIds.Concat(directDealIds).Distinct().ToList();
            if (candidateDealIds.Count == 0) return result;

            var dealNotes = await _notesDb.Notes
                .Where(n => n.DealId != null && candidateDealIds.Contains(n.DealId.Value))
                .Select(n => new { DealId = n.DealId!.Value, n.Description, n.CreatedAt })
                .ToListAsync();

            foreach (var ln in legacyNotes)
            {
                if (string.IsNullOrWhiteSpace(ln.Description)) continue;
                var matches = dealNotes
                    .Where(dn => dn.Description == ln.Description && dn.CreatedAt == ln.CreatedAt)
                    .Select(dn => dn.DealId)
                    .Distinct()
                    .ToList();
                if (matches.Count > 0) result[ln.Id] = matches;
            }

            return result;
        }
    }
}