# Email Review Viewer prototype

Standalone, offline Windows viewer proof of concept for large email collections. It does not read PST files and does not change the existing PowerShell converter.

## Requirements

- Windows 10 or 11
- .NET 8 SDK to build (installed for this prototype under `%USERPROFILE%\.dotnet`)

Run these commands from the repository root:

```powershell
$dotnet = "$env:USERPROFILE\.dotnet\dotnet.exe"
& $dotnet build .\EmailReviewViewer\EmailReviewViewer.sln
& $dotnet test .\EmailReviewViewer\EmailReviewViewer.sln
```

## Generate a 25,000-message sample

```powershell
$dotnet = "$env:USERPROFILE\.dotnet\dotnet.exe"
& $dotnet run --project .\EmailReviewViewer\EmailReviewViewer.App -c Release -- `
  --generate .\EmailReviewViewer\sample-emails.db 25000
```

The generator uses a fixed seed and writes realistic senders, recipients, dates, subjects, previews, and bodies. Re-running it replaces the sample database.

## Run the benchmark

```powershell
& $dotnet run --project .\EmailReviewViewer\EmailReviewViewer.App -c Release -- `
  --benchmark .\EmailReviewViewer\sample-emails.db
```

The benchmark executes keyword, date, sender/recipient, and selected-folder filters. Each query requests only 50 list rows and does not select message bodies.

## Launch

From source:

```powershell
& $dotnet run --project .\EmailReviewViewer\EmailReviewViewer.App -c Release -- `
  .\EmailReviewViewer\sample-emails.db
```

From a published deliverable, keep `sample-emails.db` beside `EmailReviewViewer.App.exe` and double-click the EXE. A different database can be passed as its first command-line argument.

## Folder browser

- The left sidebar lists every full `FolderPath` as a readable leaf name with its absolute message count.
- Hover a folder to see its full path; this distinguishes duplicate leaf names such as multiple `Inbox` folders.
- Type in **Search folders** to narrow the folder list without changing the email results.
- Check one or more folders to limit the paged email query to those exact full paths.
- **All Folders** shows every email. Clearing the last checked folder returns to **All Folders**.
- Folder filtering is parameterized in SQLite; the UI does not load all matching emails or bodies into memory.

## Three-pane review layout

- The folder browser remains on the far left, the paged email results grid is in the center, and a dedicated full-message reader is on the right.
- Drag either vertical splitter to resize the folder list, result grid, or reading pane.
- Select a row to load its full subject, date/time, sender, recipients, folder, and complete plain-text body.
- The body is read-only, selectable, line-break preserving, and independently scrollable.
- Paging or changing filters clears the prior selection and reader safely. Rapid keyboard or mouse selection cancels superseded detail requests.
- Message bodies are retrieved by `Id` only after selection; result pages continue to load metadata and previews only.

## Architecture

- `EmailReviewViewer.App`: .NET 8 WinForms application and command-line sample/benchmark modes.
- `EmailRepository`: parameterized SQLite access, schema creation, FTS5 search, folder counts, paged list queries, and single-message detail retrieval.
- `EmailMessages`: metadata, preview, and body storage with date/sender/folder indexes.
- `EmailMessagesFts`: external-content FTS5 index over subject, body, sender, and recipients, maintained by triggers.
- The grid loads at most 200 metadata/preview rows per page. `BodyText` is fetched only after selecting one row.
- Filter edits are debounced; superseded page queries are cancelled.
- `EmailReviewViewer.Tests`: schema, FTS, date/party filter, paging, and detail retrieval tests.

## Prototype acceptance checklist

- [x] Standalone desktop viewer; converter integration intentionally deferred.
- [x] Offline SQLite + FTS5 storage and search.
- [x] Dense list with date, From, To, subject, and preview.
- [x] Reading pane retrieves one full body on selection.
- [x] Date, participant, and keyword filters.
- [x] Searchable multi-select folder sidebar with absolute message counts.
- [x] Database paging; no all-body UI materialization.
- [x] Deterministic 25,000-record generator and query benchmark modes.
- [x] Automated repository tests.
