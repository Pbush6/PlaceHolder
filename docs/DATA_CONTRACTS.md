# Purview Teams PST → HTML — Data Contracts

**Version:** 1.0.32.1  
**Last updated:** 2026-07-08

This document defines every machine-readable contract between the core converter, launcher, GUI, tests, and Curt memory system.

## 1. Stdout pipe-delimited lines (core → launcher)

All lines are single-line, pipe-separated `Key=Value` fields. Optional `RunId=<32-char-hex>` when `-RunId` is passed.

| Prefix | When emitted | Required fields |
|--------|--------------|-----------------|
| `CONVERSION_PROGRESS` | During PST scan | `ItemsAttempted`, `ItemsExported`, `FoldersScanned`, `ItemReadFailures`, `ElapsedSeconds`, `RatePerMinute`, `FolderPath` |
| `CONVERSION_STAGE` | Report lifecycle | `Stage` (+ optional `Written`, `Total` for `WritingReport`) |
| `CONVERSION_RESULT` | Success only | `OutputPath`, `LogPath`, `ItemsExported`, `ItemReadFailures`, `AttachmentReadFailures`, `SubfolderScanFailures` + optional `TeamsOutputPath`, `EmailOutputPath`, `TeamsLogPath`, `EmailLogPath`, `TeamsItemsExported`, `EmailItemsExported` |
| `CONVERSION_ERROR` | Fatal failure (before throw) | `ExitCode`, `Message` (max 500 chars, no CR/LF/pipes) |

**Schema:** `docs/schemas/stdout-contract.schema.json`

**RunId gating (launcher):** GUI ignores any line whose `RunId` ≠ active run GUID.

## 2. Log file contract

Format: `yyyy-MM-ddTHH:mm:ss [INFO|WARN|ERROR] message`

- Single line per entry (CR/LF collapsed in messages)
- UTF-8 without BOM
- Truncated fresh each run

Summary line includes: `subfolderScanFailures` counter. When both reports are enabled, the result line also carries the per-report paths and item counts for Teams and Email.

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

## 4. HTML report contract (grep-validated)

Required DOM markers: `participantMatchMode`, `startDateFilter`, `endDateFilter`, `sortOrder`, `conversation-toolbar`, `hero-credit` containing `By Patrick Bush`.

All user-controlled text passes `ConvertTo-HtmlEncodedText`.

## 5. Launcher ↔ child process

| Concern | Contract |
|---------|----------|
| Child executable | `pwsh -Sta -NoPrompt -File <temp-core.ps1>` |
| Integrity | SHA256 of embedded bytes verified before launch |
| Stdout | Async queue (GUI) or synchronous drain (NoGui) |
| Stderr | Concurrent drain (NoGui) to prevent pipe deadlock |
| Cancel | `taskkill /T /F` + Job Object `KILL_ON_JOB_CLOSE` |
| Success | ExitCode 0 **and** report file exists on disk |

## 6. JSON state files

| File | Schema | Location |
|------|--------|----------|
| `inflight.json` | `inflight-state.schema.json` | `.cursor/state/` per workspace |
| `index.json` | `memory-index.schema.json` | `Cursor Output/Curt-Memory/` |

## 7. Version alignment

| Source | Expected |
|--------|----------|
| `build.ps1` default | Bump on release |
| `README.md` | Match shipped version |
| Pester build test | Match `build.ps1 -Version` arg |

**Current drift to resolve before release:** README/build/test version pins.

## Validation

Run: `pwsh -File scripts/Validate-DataContracts.ps1`
