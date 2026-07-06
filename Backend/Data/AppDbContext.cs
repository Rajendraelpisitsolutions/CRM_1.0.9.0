using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.ChangeTracking;
using Elpis_CRM.Model;

namespace Elpis_CRM.Data
{
    public class AppDbContext : DbContext
    {
        private readonly IHttpContextAccessor? _http;

        public AppDbContext(DbContextOptions<AppDbContext> options, IHttpContextAccessor? http = null) : base(options)
        {
            _http = http;
        }

        public DbSet<LoginModel> Login { get; set; }
        public DbSet<AccountModel> Accounts { get; set; }
        public DbSet<ContactModel> Contacts { get; set; }
        public DbSet<ProductsModel> Products { get; set; }
        public DbSet<DealModel> Deals { get; set; }
        public DbSet<DealContactLinkModel> DealContactLinks { get; set; }
        public DbSet<TemplateModel> Templates { get; set; }
        public DbSet<CallLogModel> CallLog { get; set; }
        public DbSet<TaskModel> Tasks { get; set; }
        public DbSet<MeetingModel> Meeting { get; set; }
        public DbSet<NotesModel> Notes { get; set; }
        public DbSet<NoteTargetModel> NoteTargets { get; set; }
        public DbSet<AppointmentsModel> Appointments { get; set; }
        public DbSet<AuditLogModel> AuditLogs { get; set; }
        public DbSet<RecycleBinItemModel> RecycleBinItems { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<ContactModel>()
                .Property(x => x.ContactId)
                .ValueGeneratedNever();

            modelBuilder.Entity<AccountModel>()
                .Property(x => x.AccountId)
                .ValueGeneratedNever();

            modelBuilder.Entity<DealModel>()
                .Property(x => x.Id)
                .ValueGeneratedNever();

            modelBuilder.Entity<DealContactLinkModel>()
                .HasKey(x => new { x.DealId, x.ContactId });

            modelBuilder.Entity<DealContactLinkModel>()
                .HasOne(x => x.Deal)
                .WithMany()
                .HasForeignKey(x => x.DealId)
                .OnDelete(DeleteBehavior.Cascade);

            modelBuilder.Entity<DealContactLinkModel>()
                .HasOne(x => x.Contact)
                .WithMany()
                .HasForeignKey(x => x.ContactId)
                .OnDelete(DeleteBehavior.Cascade);

            // Shared-note associations (Option A). Cascade so deleting a note clears its links;
            // the lookup index covers "which notes appear under this deal/contact".
            modelBuilder.Entity<NoteTargetModel>()
                .HasOne(x => x.Note)
                .WithMany()
                .HasForeignKey(x => x.NoteId)
                .OnDelete(DeleteBehavior.Cascade);

            modelBuilder.Entity<NoteTargetModel>()
                .HasIndex(x => new { x.TargetType, x.TargetId });

            modelBuilder.Entity<NoteTargetModel>()
                .HasIndex(x => new { x.NoteId, x.TargetType, x.TargetId })
                .IsUnique();

            modelBuilder.Entity<AuditLogModel>().HasIndex(x => new { x.EntityName, x.EntityId });
            modelBuilder.Entity<AuditLogModel>().HasIndex(x => x.ChangedAt);
        }

        // ── Audit logging ───────────────────────────────────────────────────────
        // Never log these (noise / secrets), and never log the AuditLogs table itself.
        private static readonly HashSet<string> SkipProps = new(StringComparer.OrdinalIgnoreCase)
        {
            "Password", "FrontImage", "BackImage"
        };

        public override async Task<int> SaveChangesAsync(CancellationToken cancellationToken = default)
        {
            var pending = BuildAuditEntries();
            var result = await base.SaveChangesAsync(cancellationToken);

            if (pending.Count > 0)
            {
                foreach (var (entry, log) in pending)
                {
                    // Inserts get their generated key after the first save; deleted entries are detached.
                    if (entry.State != EntityState.Detached)
                        log.EntityId = TryGetKey(entry);
                }
                AuditLogs.AddRange(pending.Select(p => p.Log));
                await base.SaveChangesAsync(cancellationToken);
            }

            return result;
        }

        private List<(EntityEntry Entry, AuditLogModel Log)> BuildAuditEntries()
        {
            var list = new List<(EntityEntry, AuditLogModel)>();

            var entries = ChangeTracker.Entries()
                .Where(e => e.Entity is not AuditLogModel
                            && (e.State == EntityState.Added || e.State == EntityState.Modified || e.State == EntityState.Deleted))
                .ToList();

            if (entries.Count == 0) return list;

            var (email, name, role, ip) = CurrentUser();
            var now = DateTime.UtcNow;

            foreach (var entry in entries)
            {
                string action;
                Dictionary<string, object?>? changes = null;

                if (entry.State == EntityState.Added)
                {
                    action = "Insert";
                    changes = new();
                    foreach (var p in entry.Properties)
                    {
                        if (SkipProps.Contains(p.Metadata.Name) || p.Metadata.ClrType == typeof(byte[])) continue;
                        changes[p.Metadata.Name] = p.CurrentValue;
                    }
                }
                else if (entry.State == EntityState.Deleted)
                {
                    action = "Delete";
                    // Capture the deleted row's values (so admins can see what was removed).
                    changes = new();
                    foreach (var p in entry.Properties)
                    {
                        if (SkipProps.Contains(p.Metadata.Name) || p.Metadata.ClrType == typeof(byte[])) continue;
                        changes[p.Metadata.Name] = p.OriginalValue;
                    }
                }
                else // Modified
                {
                    action = "Update";
                    changes = new();
                    foreach (var p in entry.Properties)
                    {
                        if (!p.IsModified) continue;
                        if (SkipProps.Contains(p.Metadata.Name) || p.Metadata.ClrType == typeof(byte[])) continue;
                        changes[p.Metadata.Name] = new { old = p.OriginalValue, @new = p.CurrentValue };
                    }
                    if (changes.Count == 0) continue; // nothing meaningful changed
                }

                var log = new AuditLogModel
                {
                    EntityName = entry.Entity.GetType().Name.Replace("Model", ""),
                    EntityId = TryGetKey(entry),
                    Action = action,
                    ChangedBy = email,
                    ChangedByName = name,
                    ChangedByRole = role,
                    ChangedAt = now,
                    IpAddress = ip,
                    Changes = changes != null ? JsonSerializer.Serialize(changes, JsonOpts) : null
                };
                list.Add((entry, log));
            }

            return list;
        }

        private static readonly JsonSerializerOptions JsonOpts = new() { WriteIndented = false };

        private static string? TryGetKey(EntityEntry entry)
        {
            var key = entry.Metadata.FindPrimaryKey();
            if (key == null) return null;
            try
            {
                var vals = key.Properties.Select(p => entry.Property(p.Name).CurrentValue?.ToString());
                return string.Join(",", vals);
            }
            catch { return null; }
        }

        private (string? email, string? name, string? role, string? ip) CurrentUser()
        {
            var user = _http?.HttpContext?.User;
            var email = user?.FindFirst(ClaimTypes.Email)?.Value;
            var name = user?.FindFirst(ClaimTypes.Name)?.Value;
            var role = user?.FindFirst(ClaimTypes.Role)?.Value;
            var ip = _http?.HttpContext?.Connection?.RemoteIpAddress?.ToString();
            return (string.IsNullOrWhiteSpace(email) ? "system" : email, name, role, ip);
        }
    }
}
