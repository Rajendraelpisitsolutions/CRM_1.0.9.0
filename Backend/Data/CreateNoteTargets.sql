-- ============================================================================
-- Option A: shared notes.
-- Adds Notes.IsShared and the NoteTargets association table.
-- Idempotent and non-destructive: it never touches or deletes existing rows;
-- existing notes keep IsShared = 0 and their current behaviour.
-- Run once against the CRM database.
-- ============================================================================

-- 1) Notes.IsShared — marks notes created under the shared model.
IF NOT EXISTS
(
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID(N'dbo.Notes', N'U')
      AND name = N'IsShared'
)
BEGIN
    ALTER TABLE dbo.Notes
        ADD IsShared BIT NOT NULL CONSTRAINT DF_Notes_IsShared DEFAULT (0);
END;

-- 2) NoteTargets — one row per record a shared note appears under.
IF OBJECT_ID(N'dbo.NoteTargets', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.NoteTargets
    (
        Id         BIGINT        NOT NULL IDENTITY(1,1),
        NoteId     INT           NOT NULL,
        TargetType NVARCHAR(20)  NOT NULL,   -- 'Deal' or 'Contact'
        TargetId   BIGINT        NOT NULL,   -- Deals.Id or Contacts.ContactId
        CONSTRAINT PK_NoteTargets PRIMARY KEY (Id),
        CONSTRAINT FK_NoteTargets_Notes
            FOREIGN KEY (NoteId) REFERENCES dbo.Notes(Id) ON DELETE CASCADE
    );

    -- Fast "which notes appear under this deal/contact" lookups.
    CREATE INDEX IX_NoteTargets_Target
        ON dbo.NoteTargets (TargetType, TargetId);

    -- Prevent duplicate associations for the same note + target.
    CREATE UNIQUE INDEX UX_NoteTargets_Note_Target
        ON dbo.NoteTargets (NoteId, TargetType, TargetId);
END;
