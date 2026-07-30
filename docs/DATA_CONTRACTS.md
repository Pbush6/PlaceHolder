# Purview PST Reports — Data Contracts

**Version:** 1.2.0.0
**Last updated:** 2026-07-29

This document defines every machine-readable contract between the core converter, launcher, GUI, tests, and Curt memory system.

## 1. Stdout pipe-delimited lines (core → launcher)

All lines are single-line, pipe-separated `Key=Value` fields. Optional `RunId=<32-lowercase-hex>` when `-RunId` is passed. The JSON schema rejects malformed RunIds and unknown properties on every machine-line type.

| Prefix | When emitted | Required fields |
|--------|--------------|-----------------|
| `CONVERSION_PROGRESS` | During PST scan | `ItemsAttempted`, `ItemsExported`, `FoldersScanned`, `ItemReadFailures`, `ElapsedSeconds`, `RatePerMinute`, `FolderPath` |
| `CONVERSION_STAGE` | Report lifecycle | `Stage` (+ optional `Written`, `Total` for `WritingReport`); includes `ImportingEmailDatabase` |
| `CONVERSION_RESULT` | Success only | Backward-compatible `OutputPath`, `LogPath`, `ItemsExported`, failure counters; typed `TeamsOutputPath`, `EmailOutputPath`, `CalendarOutputPath`, `ContactsOutputPath`, corresponding typed log paths, and per-report exported counts |
| `CONVERSION_ERROR` | Fatal failure (before throw) | `ExitCode`, `Message` (max 500 chars, no CR/LF/pipes) |

**Schema:** `docs/schemas/stdout-contract.schema.json`

**RunId gating (launcher):** GUI ignores any line whose `RunId` ≠ active run GUID.

## 2. Log file contract

Format: `yyyy-MM-ddTHH:mm:ss [INFO|WARN|ERROR] message`

- Single line per entry (CR/LF collapsed in messages)
- UTF-8 without BOM
- Truncated fresh each run

Summary line includes the `subfolderScanFailures` counter. The result line always carries all four typed output/log fields and per-report counts; paths for unselected reports are empty. The six-item sample exports 2 Teams, 2 Email, 1 Calendar, and 1 Contacts item (`ItemsExported=6`).

## 3. In-memory message record (core internal)

Not serialized to JSON; shape produced by `Get-MessageRecord`:

| Field | Type | Notes |
|-------|------|-------|
| `SortTime` | datetime | Sentinel dates (year ≥ 4500) treated as missing |
| `FolderPath` | string | Backslash-separated |
| `Subject`, `MessageClass` | string | |
| `SenderName`, `SenderEmail`, `SenderDisplay` | string | |
| `To`, `Cc` | string | |
| `Participants` | string[] | |
| `ParticipantsKey`, `ConversationKey`, `ConversationTitle` | string | |
| `SentOn`, `ReceivedTime`, `CreationTime` | datetime? | |
| `EntryId` | string | |
| `BodyText` | string | HTML-encoded at write time |
| `AttachmentsHtml` | string | Pre-rendered table HTML |

## 4. Output report contracts

**Teams report** required DOM markers: `participantMatchMode`, `startDateFilter`, `endDateFilter`, `sortOrder`, `conversation-toolbar`, `hero-credit` containing `By Patrick Bush`.

**Email report:** SQLite database named `Base_Email.db`, with `EmailMessages`, `EmailMessagesFts`, indexes, and FTS maintenance triggers. The viewer pages metadata and loads a body only when selected.

**Calendar report:** static HTML named `Base_Calendar.html`, with encoded appointment/meeting records and dataset-backed search/date/folder/type/all-day/recurrence filters.

**Contacts report:** static HTML named `Base_Contacts.html`, with encoded contact/distribution-list records and dataset-backed text search, folder, and category filters.

All user-controlled text passes `ConvertTo-HtmlEncodedText`.

## 5. Email NDJSON import

UTF-8 without BOM, one JSON object per physical line. Required data fields are `FolderPath`, `SenderName`, `SenderAddress`, `ToRecipients`, `CcRecipients`, `Subject`, `SentUtc`, `ReceivedUtc`, `Preview`, `BodyText`, `MessageClass`, `EntryId`, `ConversationId`, and `ConversationTopic`.

- Date values are ISO-8601 with timezone offset and are normalized to UTC by the importer.
- Preview is the first non-empty body line, bounded to 500 characters.
- `EntryId` is the unique duplicate key; missing values receive a deterministic SHA-256 fallback.
- Import writes `<output>.importing`, validates final row count, then replaces the destination.
- Staging NDJSON is removed only after importer exit code 0; failures log and retain its exact path.
- Email attachment summaries remain deferred in 1.2.0.0.

## 6. Launcher ↔ child process

| Concern | Contract |
|---------|----------|
| Child executable | `pwsh -Sta -NoPrompt -File <temp-core.ps1>` |
| Integrity | SHA256 of embedded bytes verified before launch |
| Stdout | Async queue (GUI) or synchronous drain (NoGui) |
| Stderr | Concurrent drain (NoGui) to prevent pipe deadlock |
| Cancel | `taskkill /T /F` + Job Object `KILL_ON_JOB_CLOSE` |
| Report flags | Explicit `TeamsReport`, `EmailReport`, `CalendarReport`, `ContactsReport`; all default `true`; CLI and GUI reject all four `false` |
| Legacy flag binding | A direct core/launcher call that explicitly binds `TeamsReport` and/or `EmailReport` while omitting both new flags uses legacy two-report behavior (`CalendarReport=false`, `ContactsReport=false`). Calls binding neither family retain all-four defaults. The current launcher always forwards all four flags explicitly. |
| Success | ExitCode 0, valid `CONVERSION_RESULT`, and every selected report's typed output and typed log fields are non-empty and reference existing files |
| Launch | Email `.db` opens in Email Reviewer; Teams/Calendar/Contacts HTML opens in the default browser, once per selected output and only after success |

The no-console release EXE is compiled with PS2EXE `-NoOutput`. It therefore
does not expose stdout or `CONVERSION_RESULT` to scripted consumers. Use the
source launcher or packaged Debug converter for machine-readable stdout.
Release `-NoGui` verification uses process exit, selected typed artifacts,
typed logs, and their summary content.

## 7. JSON state files

| File | Schema | Location |
|------|--------|----------|
| `inflight.json` | `inflight-state.schema.json` | `.cursor/state/` per workspace |
| `index.json` | `memory-index.schema.json` | `Cursor Output/Curt-Memory/` |

## 8. Version alignment

| Source | Expected |
|--------|----------|
| `build.ps1` default | `1.2.0.0` |
| `README.md` | `1.2.0.0` |
| Email Reviewer assembly/file version | `1.2.0.0` |
| Pester build and release tests | `1.2.0.0` |

Release 1.2.0.0 is aligned across source defaults, documentation, executable metadata, package naming, and release verification.

## Validation

Run: `pwsh -File scripts/Validate-DataContracts.ps1`
