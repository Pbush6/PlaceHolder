# Purview Teams PST to HTML Converter - Build Notes

Generated: 2026-07-03

## Version 1.2.0.0 packaging (2026-07-29)

- Builds the converter with the current embedded four-report core.
- Adds Calendar (`_Calendar.html`) and Contacts (`_Contacts.html`) static reports alongside Teams HTML and Email SQLite.
- Packages the final self-contained Email Reviewer, including the bottom-right `By Patrick Bush` credit.
- Builds the no-console release converter with PS2EXE `-NoOutput`; scripts that require stdout/`CONVERSION_RESULT` must use the source launcher or packaged Debug converter. Release `-NoGui` success is verified from exit code, typed artifacts/logs, and log summaries.
- Portable release name: `PurviewTeamsPstToHtmlApp-1.2.0.0-win-x64.zip`.

## Version 1.1.0.1 packaging (2026-07-13)

- Packaged viewer discovery starts from the running converter executable directory, with source and build fallbacks for development runs.
- Viewer launch uses Windows PowerShell-compatible native argument quoting for database paths containing spaces and punctuation.
- Automatic launch failures preserve conversion success and log the exact viewer path, database path, and error.
- Email Review Viewer can open and switch compatible databases from its File menu or Open Database button; startup without an argument shows a welcome state.

## Version 1.1.0.0 packaging (2026-07-13)

- Build converter EXEs with `pwsh -File .\build.ps1 -Version 1.1.0.0`.
- Publish the viewer framework-dependent for win-x64:
  `dotnet publish .\EmailReviewViewer\EmailReviewViewer.App\EmailReviewViewer.App.csproj -c Release -r win-x64 --self-contained false -o .\EmailReviewViewer\artifacts\publish\win-x64`.
- Deliver `PurviewTeamsPstToHtmlConverter.exe`, `PurviewTeamsPstToHtmlConverter_Debug.exe`, `README.md`, and the complete published `EmailReviewViewer\` folder together.
- The launcher resolves `EmailReviewViewer\EmailReviewViewer.App.exe` beside the converter and fails with an actionable message when Email output is selected but the viewer is missing.

## Cursor working-copy maintenance update (2026-07-08)

This build-notes file originally described the Hermes working-copy/output flow and is now partially historical.
For the current Cursor review/build pass, treat these as authoritative:

- Active working copy: `C:\Users\pbush\OneDrive - Perfection Learning\Documents\AI\Cursor Working Directory\PurviewTeamsPstToHtmlApp`
- Safe build output folder: `...\Cursor Working Directory\PurviewTeamsPstToHtmlApp\build`
- Do **not** overwrite the original shipped EXE in `Documents\AI` during review builds.
- Recommended hardening added in this pass: only treat `taskkill /T /F` as successful when it exits 0; otherwise fall back to `.Kill()`.

Older sections below still document prior packaging milestones, but their path references may point at the Hermes/original output locations rather than this Cursor working copy.

## Source and output

- Project folder: `C:\Users\pbush\OneDrive - Perfection Learning\Documents\AI\Hermes Working Directory\PurviewTeamsPstToHtmlApp`
- Original script: `C:\Users\pbush\Downloads\Convert-PurviewTeamsPstToHtml_1.0.ps1`
- Core converter: `src\Convert-PurviewTeamsPstToHtml.ps1`
- Launcher: `src\Start-PurviewTeamsPstToHtmlApp.ps1`
- Release EXE: `C:\Users\pbush\OneDrive - Perfection Learning\Documents\AI\Hermes Output\PurviewTeamsPstToHtmlConverter.exe`
- Debug EXE: `C:\Users\pbush\OneDrive - Perfection Learning\Documents\AI\Hermes Output\PurviewTeamsPstToHtmlConverter_Debug.exe`

## Implemented phases

1. Created project structure and original-script backup.
2. Added packaging-friendly core converter behavior:
   - `-NoPrompt` switch.
   - Clear failure when `-NoPrompt` is used without a valid PST path.
   - Machine-readable `CONVERSION_RESULT|...` output line.
3. Created WinForms launcher script.
   - Uses embedded Base64 copy of the core converter, so the EXE is self-contained for the converter logic.
   - GUI supports PST selection, output/log paths, participant defaults, Keep PST Attached, and sample-data mode.
   - `-NoGui` mode supports smoke testing and automation.
4. Installed/verified PS2EXE 1.0.18 for the current user.
5. Built debug and release EXEs.
6. Validated sample-data outputs.
7. Wrote README.

## Verification evidence

Parser checks:
- Core converter: OK.
- Launcher: OK.

Sample-data script/launcher runs:
- Core script produced sample HTML/log and `CONVERSION_RESULT`.
- Launcher script produced sample HTML/log and `CONVERSION_RESULT`.

EXE artifacts:
- `PurviewTeamsPstToHtmlConverter_Debug.exe`: 114,688 bytes.
- `PurviewTeamsPstToHtmlConverter.exe`: 122,880 bytes.

Sample reports:
- Debug sample HTML: 22,646 bytes.
- Release sample HTML: 22,648 bytes.
- Debug sample log: 549 bytes.
- Release sample log: 551 bytes.

HTML smoke checks passed:
- Report title present.
- Fireflies.ai sample appears.
- Exact-Conversation filter present.
- Resizable filter handle present.
- HTML/script sample is safely encoded.

Log smoke checks passed:
- `exported=3`.
- `itemReadFailures=0`.

## Known caveats

- The app is a packaged PowerShell utility, not a native rewrite.
- Real PST conversion still requires Outlook COM on Windows.
- The EXEs are unsigned and may trigger SmartScreen/AV warnings.
- Directly running PS2EXE console apps from Git Bash can appear to hang in this environment; validation was done via PowerShell `Start-Process` and output/file checks.
- GUI conversion currently runs synchronously; for very large PSTs the window may appear busy while Outlook COM scanning is active, but the converter writes log output and the process should complete.

## Fix: WinForms PlaceholderText compatibility

After the first packaged GUI run, the EXE reported:

`The property 'PlaceholderText' cannot be found on this object. Verify that the property exists and can be set.`

Cause: PS2EXE hosts the script on a Windows PowerShell/.NET WinForms surface where `System.Windows.Forms.TextBox.PlaceholderText` is not available. The source ran far enough to parse and NoGui tests passed, but the interactive GUI path failed when setting that property.

Fix applied:
- Removed `$participantsBox.PlaceholderText = ...`.
- Added a separate gray help `Label` under the participants textbox.
- Rebuilt debug and release EXEs as version `1.0.1.0`.

Post-fix verification:
- Launcher parser OK.
- No `PlaceholderText` references remain in launcher source.
- Source launcher sample-data run generated HTML/log and `CONVERSION_RESULT`.
- Debug EXE sample-data run generated HTML/log and `CONVERSION_RESULT`.
- Release EXE sample-data run generated HTML/log.
- HTML smoke checks passed: title, Fireflies.ai, Exact-Conversation, resize handle, HTML encoding.
- Log smoke checks passed: `exported=3`, `itemReadFailures=0`.

## Change: Remove Default participants GUI section

Requested change: remove the section asking for default participants.

Implementation:
- Removed the `Default participants` label, textbox, and help text from the WinForms GUI.
- Moved the checkboxes and action buttons upward to close the gap.
- The GUI now passes an empty default participant list to the converter.
- Command-line `-DefaultConversationParticipants` support remains in the underlying converter/launcher plumbing if needed later, but it is no longer exposed in the app window.
- Rebuilt debug and release EXEs as version `1.0.2.0`.

Post-change verification:
- Launcher parser OK.
- Search confirmed removed GUI terms are no longer in the launcher source.
- Source launcher `-NoGui -UseSampleData` run generated HTML/log and `CONVERSION_RESULT`.
- Debug EXE sample-data run passed.
- Release EXE sample-data run passed.
- HTML smoke checks passed: title, Fireflies.ai, Exact-Conversation, resize handle, HTML encoding.
- Log smoke checks passed: `exported=3`, `itemReadFailures=0`.

## Change: Add GUI progress bar

Requested change: add a progress bar for conversion.

Implementation:
- Added a WinForms `ProgressBar` and status `Label` below the action buttons.
- Conversion now runs through a `System.ComponentModel.BackgroundWorker` in the GUI path so the window can update while conversion runs.
- The progress indicator is stage/activity-based: start, preparing, reading PST / collecting messages, writing final report/log, complete/failure.
- Exact per-message percentage is not implemented because Outlook COM scanning does not know the final total across nested PST folders before scanning, and pre-scanning would double the slowest part of the process.
- Rebuilt debug and release EXEs as version `1.0.3.0`.

Build note:
- The first release rebuild attempt failed because the old release EXE was still running and locked. The running converter process was terminated, then the release EXE rebuilt successfully.

Post-change verification:
- Launcher parser OK.
- Source contains `ProgressBar`, `BackgroundWorker`, `ReportProgress`, and `Marquee` progress handling.
- Source launcher `-NoGui -UseSampleData` run generated HTML/log and `CONVERSION_RESULT`.
- Debug EXE sample-data run passed.
- Release EXE sample-data run passed.
- HTML smoke checks passed: title, Fireflies.ai, Exact-Conversation, resize handle, HTML encoding.
- Log smoke checks passed: `exported=3`, `itemReadFailures=0`.

## Change: Open generated HTML report automatically

Requested change: when conversion completes and the HTML file is created, open the HTML file.

Implementation:
- In the GUI completion handler, after a successful conversion and after resolving `OutputPath`, the app now calls `Start-Process -FilePath $script:lastReportPath` when the report file exists.
- The progress/status label changes to `Conversion completed successfully. Opening report...`.
- The log textbox records `Opened report: <path>` if the automatic open succeeds, or a nonfatal warning if opening fails.
- The success message now states that the report has been opened in the default browser.
- Rebuilt debug and release EXEs as version `1.0.4.0`.

Post-change verification:
- Launcher parser OK.
- Auto-open source references present: `Opening report`, `Opened report`, and `Start-Process -FilePath $script:lastReportPath`.
- Source launcher `-NoGui -UseSampleData` run generated HTML/log and `CONVERSION_RESULT`.
- Debug EXE sample-data run passed.
- Release EXE sample-data run passed.
- HTML smoke checks passed: title, Fireflies.ai, Exact-Conversation, resize handle, HTML encoding.
- Log smoke checks passed: `exported=3`, `itemReadFailures=0`.

## Fix: BackgroundWorker runspace failure

Reported error: `There is no Runspace available to run scripts in this thread. You can provide one in the DefaultRunspace property...`

Root cause:
- The version `1.0.4.0` progress-bar GUI ran the conversion from a .NET `System.ComponentModel.BackgroundWorker` worker thread.
- That thread did not have a PowerShell runspace in the PS2EXE-hosted app.
- PowerShell script blocks/functions such as `Invoke-EmbeddedConversion` therefore could not execute on that thread.

Implementation:
- Removed `BackgroundWorker`, `RunWorkerAsync`, and worker-thread script-block conversion.
- Added `Set-ConversionProgress` helper that updates the progress bar/status label, refreshes the form, and calls `[System.Windows.Forms.Application]::DoEvents()`.
- Conversion now runs on the normal PowerShell/WinForms thread, where the active runspace exists.
- Preserved staged progress updates and automatic report opening.
- Rebuilt debug and release EXEs as version `1.0.5.0`.

Tradeoff:
- This avoids the runspace failure and is safer for PS2EXE/PowerShell script execution.
- During the deepest Outlook COM scanning work, progress is still stage-based rather than per-message; the app updates at safe stages without moving PowerShell conversion logic to a runspace-less thread.

Post-change verification:
- Launcher parser OK.
- Confirmed `BackgroundWorker`, `RunWorkerAsync`, and `Add_DoWork` references are removed.
- Confirmed progress helpers are present.
- Source launcher `-NoGui -UseSampleData` run generated HTML/log and `CONVERSION_RESULT`.
- Debug EXE sample-data run passed.
- Release EXE sample-data run passed.
- HTML smoke checks passed: title, Fireflies.ai, Exact-Conversation, resize handle, HTML encoding.
- Log smoke checks passed: `exported=3`, `itemReadFailures=0`.

## Change: Remove built-in sample data option from GUI

Requested change: remove the `Use built-in sample data` option.

Implementation:
- Removed the `Use built-in sample data (safe test mode; no Outlook/PST required)` checkbox from the WinForms GUI.
- Removed GUI references to `$sampleCheck`.
- GUI conversion now always requires an existing PST file and always calls Outlook availability checks with `-UseSampleData:$false`.
- The hidden command-line `-NoGui -UseSampleData` option remains in the launcher for automated smoke tests without real PST/Outlook dependencies.
- Rebuilt debug and release EXEs as version `1.0.6.0`.

Build note:
- The first release rebuild attempt failed because the old release EXE was still running and locked. The running converter process was terminated, then the release EXE rebuilt successfully.

Post-change verification:
- Launcher parser OK.
- Confirmed GUI sample option terms are removed: `sampleCheck`, `Use built-in sample data`, `safe test mode`, and `enable sample-data mode`.
- Confirmed command-line sample-data support remains for testing.
- Source launcher `-NoGui -UseSampleData` run generated HTML/log and `CONVERSION_RESULT`.
- Debug EXE sample-data run passed.
- Release EXE sample-data run passed.
- HTML smoke checks passed: title, Fireflies.ai, Exact-Conversation, resize handle, HTML encoding.
- Log smoke checks passed: `exported=3`, `itemReadFailures=0`.

## Change: Low-overhead collection-phase progress every 100 messages

Requested change: implement periodic progress callbacks inside the message collection loop, updating every 100 messages with the least performance impact practical.

Implementation:
- Added `Write-ConversionProgress` to the core converter.
- During `Read-OutlookFolder`, after every `$LogEvery` collected message-like items, the converter emits one machine-readable stdout line:
  `CONVERSION_PROGRESS|ItemsAttempted=...|ItemsExported=...|FoldersScanned=...|ItemReadFailures=...|ElapsedSeconds=...|RatePerMinute=...|FolderPath=...`
- The default `$LogEvery` remains 100, so GUI progress updates occur every 100 collected messages.
- The progress line is stdout-only and avoids extra UI work per message. Existing periodic file logging remains at the same 100-message cadence.
- Updated `Invoke-EmbeddedConversion` in the launcher to read child process stdout line-by-line instead of waiting for the process to exit before reading all output.
- Added an `-OnProgress` callback in the GUI path. The callback updates the status label with exported count, attempted count, folders scanned, elapsed time, and messages/minute rate.
- The progress bar advances from the collection stage toward the write stage as 100-message batches arrive, capped before completion because no exact total is available without a costly pre-scan.
- Re-embedded the updated core converter into the launcher.
- Rebuilt debug and release EXEs as version `1.0.7.0`.

Performance rationale:
- No pre-scan was added.
- No per-message UI updates were added.
- UI updates happen only once per 100 collected messages.
- The child converter emits one compact progress line per 100 messages, which is low overhead compared with Outlook COM PST traversal.

Post-change verification:
- Core parser OK.
- Launcher parser OK.
- Confirmed references for `CONVERSION_PROGRESS`, `Write-ConversionProgress`, `OnProgress`, line-by-line stdout reading, and collection status updates.
- Source launcher `-NoGui -UseSampleData` run generated HTML/log and `CONVERSION_RESULT`.
- Debug EXE sample-data run passed.
- Release EXE sample-data run passed.
- HTML smoke checks passed: title, Fireflies.ai, Exact-Conversation, resize handle, HTML encoding.
- Log smoke checks passed: `exported=3`, `itemReadFailures=0`.

Limitations:
- The built-in sample data contains only three messages and does not naturally exercise the every-100-message progress line. Real PSTs with at least 100 collected message-like items will exercise the GUI callback.
- The progress remains an estimate/stage indicator rather than a true percentage of total PST work because the final nested PST item total is not known without pre-scanning.

## Fix: Keep GUI responsive during message collection

Reported issue: after conversion started, the progress bar looked no different and Windows showed the app as `Not Responding`.

Root cause:
- Version `1.0.7.0` avoided PowerShell background runspace errors, but the GUI still called a blocking child-process output loop on the WinForms UI thread.
- While that loop waited for stdout lines, the WinForms message pump could not process paint/input messages reliably.
- Windows therefore marked the app as unresponsive during long PST collection phases.

Implementation:
- Added `Start-EmbeddedConversionProcess` for GUI runs. It starts the embedded core converter as a child process and immediately returns to the UI.
- The GUI no longer blocks on `ReadLine()` or waits for the child conversion process in the button-click handler.
- Added a WinForms `Timer` that ticks once per second while conversion is running.
- The timer checks the child process with `HasExited`, polls the log file with `Get-Content -Tail 25`, and updates the status/progress label from log checkpoints.
- When `Collected N message-like items so far` appears, the GUI displays collected count, elapsed time, and messages/minute.
- Before the first 100-message checkpoint, the GUI still updates elapsed time and current folder when available.
- On process exit, the timer performs success/failure handling, enables buttons, opens the report, and cleans up the temporary embedded core folder.
- Rebuilt debug and release EXEs as version `1.0.8.0`.

Performance/responsiveness rationale:
- The converter remains a separate PowerShell process, so Outlook COM/PST work does not run on the GUI thread.
- The GUI polls a small tail of the log once per second; no per-message UI updates or PST pre-scan are used.
- This should keep Windows from showing `Not Responding` while still providing progress during long conversions.

Post-change verification:
- Core parser OK.
- Launcher parser OK.
- Confirmed nonblocking GUI references: `Start-EmbeddedConversionProcess`, `Windows.Forms.Timer`, `activeConversion`, `Get-Content -Tail`, `HasExited`, and `RedirectStandardOutput = $false`.
- Source launcher `-NoGui -UseSampleData` run generated HTML/log and `CONVERSION_RESULT`.
- Debug EXE sample-data run passed.
- Release EXE sample-data run passed.
- HTML smoke checks passed: title, Fireflies.ai, Exact-Conversation, resize handle, HTML encoding.
- Log smoke checks passed: `exported=3`, `itemReadFailures=0`.

Limitations:
- Automated sample-data tests verify the converter/report path but do not fully simulate a long interactive PST conversion. Real PST testing is still the best validation for responsiveness under long Outlook COM workloads.

## Change: Monotonic/slower progress bar movement

Reported issue: the progress bar repeatedly advanced quickly and then moved backward.

Root cause:
- The timer fallback used a repeating pulse formula based on elapsed seconds modulo 10.
- That made the bar cycle from roughly 25 to 34 and then back to 25.
- The first collected-message checkpoint could also compute a lower progress value than the current pulse value.

Implementation:
- Added `$script:lastProgressValue`.
- Updated `Set-ConversionProgress` so progress is monotonic within a conversion run: requested values lower than the last progress value are ignored.
- Added `Reset-ConversionProgress` to reset the monotonic state only at the start of a new conversion.
- Replaced the repeating modulo pulse with a slow elapsed-time ramp: one progress point every 30 seconds, capped at 35 before message counts are available.
- Slowed collection-count bar advancement from one point per 100 collected messages to one point per 250 collected messages, while still updating the status text as checkpoints arrive.
- Rebuilt debug and release EXEs as version `1.0.9.0`.

Post-change verification:
- Core parser OK.
- Launcher parser OK.
- Confirmed monotonic/slow-progress references: `lastProgressValue`, `Reset-ConversionProgress`, `Ceiling($script:lastCollectedCount / 250)`, and `Floor($elapsedSeconds / 30)`.
- Source launcher `-NoGui -UseSampleData` run generated HTML/log and `CONVERSION_RESULT`.
- Debug EXE sample-data run passed.
- Release EXE sample-data run passed.
- HTML smoke checks passed: title, Fireflies.ai, Exact-Conversation, resize handle, HTML encoding.
- Log smoke checks passed: `exported=3`, `itemReadFailures=0`.

## Change: Custom EXE icon

Requested change: give the `.exe` file a cool thumbnail/icon.

Implementation:
- Created a custom app icon with a blue/purple rounded-square background, document, PST/database cylinder, Teams-style chat bubbles, HTML label, and check mark.
- Saved source PNG and multi-size Windows `.ico` under `assets/`.
- Rebuilt both debug and release EXEs with PS2EXE `-IconFile`.
- Rebuilt as version `1.0.10.0`.

Assets:
- `assets/PurviewTeamsPstToHtmlConverter_icon_1024.png`
- `assets/PurviewTeamsPstToHtmlConverter_icon.ico`
- Extracted verification preview: `assets/PurviewTeamsPstToHtmlConverter_extracted_icon.png`

Post-change verification:
- `System.Drawing.Icon.ExtractAssociatedIcon()` successfully extracted the custom icon from the release EXE.
- Release EXE file version verified as `1.0.10.0`.
- Debug EXE sample-data run passed.
- Release EXE sample-data run passed.
- HTML smoke checks passed: title, Fireflies.ai, Exact-Conversation, resize handle, HTML encoding.
- Log smoke checks passed: `exported=3`, `itemReadFailures=0`.

## Change: PST email username-based output filename

Requested change: name the created HTML file according to the user name of the PST file followed by `Teams Messages`. User clarified that the user name is the part of the email address before the `@` sign in the PST filename.

Implementation:
- Added `Get-SafeOutputBaseNameFromPstPath` in the GUI launcher.
- When the PST filename contains `@`, the app uses only the substring before `@`.
- The GUI sets output paths to `<username> Teams Messages.html` and `<username> Teams Messages.log` in the default output folder.
- Output/log paths update when the user selects or types a PST path, unless the user has manually chosen a custom output/log path via the Save dialogs.
- If a PST filename does not contain `@`, the app falls back to the PST filename without extension.
- Rebuilt debug and release EXEs as version `1.0.11.0`.

Examples:
- `jane.doe@company.com.pst` -> `jane.doe Teams Messages.html`
- `john_smith@perfectionlearning.com Teams Export.pst` -> `john_smith Teams Messages.html`
- `NoEmailName.pst` -> `NoEmailName Teams Messages.html`

Post-change verification:
- Launcher parser OK.
- Confirmed source references: `IndexOf('@')`, `Substring(0, $atIndex)`, and `Teams Messages`.
- Validated representative filename transformations in PowerShell.
- Source launcher `-NoGui -UseSampleData` run generated HTML/log and `CONVERSION_RESULT`.
- Release EXE version verified as `1.0.11.0`.
- Debug EXE sample-data run passed.
- Release EXE sample-data run passed.
- HTML smoke checks passed: title, Fireflies.ai, Exact-Conversation, resize handle, HTML encoding.
- Log smoke checks passed: `exported=3`, `itemReadFailures=0`.

## Change: Default GUI save location is Downloads

Requested change: make the default save locations for Output HTML and Log file the user's Downloads folder.

Implementation:
- Updated `Get-DefaultOutputDirectory` in the launcher to return `Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads'` when it exists.
- Fallback is the user profile folder if Downloads does not exist.
- Removed the previous preference for the Hermes Output folder as the GUI default save location.
- PST-derived output naming remains unchanged; only the default folder changed.
- Rebuilt debug and release EXEs as version `1.0.12.0`.

Example default paths after selecting `jane.doe@company.com.pst`:
- `C:\Users\pbush\Downloads\jane.doe Teams Messages.html`
- `C:\Users\pbush\Downloads\jane.doe Teams Messages.log`

Post-change verification:
- Launcher parser OK.
- Confirmed `Get-DefaultOutputDirectory` source now prefers Downloads.
- Confirmed `C:\Users\pbush\Downloads` exists on this machine.
- Source launcher `-NoGui -UseSampleData` run generated HTML/log and `CONVERSION_RESULT`.
- Release EXE version verified as `1.0.12.0`.
- Debug EXE sample-data run passed.
- Release EXE sample-data run passed.
- HTML smoke checks passed: title, Fireflies.ai, Exact-Conversation, resize handle, HTML encoding.
- Log smoke checks passed: `exported=3`, `itemReadFailures=0`.

## Change: Clear prerequisite messages for PowerShell 7 and Outlook

Requested change: if the computer running the EXE does not have PowerShell 7 or Outlook installed, show a message saying the missing program needs to be installed.

Implementation:
- Added `Resolve-PowerShell7Path`.
- Removed direct reliance on `(Get-Command pwsh).Source`, which caused the reported `The property 'Source' cannot be found on this object` error on some machines.
- The resolver checks `Source`, `Path`, and `Definition`, then standard PowerShell 7 install paths.
- If PowerShell 7 is not found, it throws: `PowerShell 7 is required for this converter, but it was not found on this computer. Install PowerShell 7 from Microsoft, then run the converter again.`
- Updated Outlook prerequisite message to: `Microsoft Outlook is required for this converter, but Outlook is not installed or is not registered on this computer. Install and configure Microsoft Outlook, then run the converter again.`
- GUI startup checks both prerequisites and shows a `Required program missing` message box before the main form proceeds.
- The Convert button also checks PowerShell 7 and Outlook before launching conversion.
- Rebuilt debug and release EXEs as version `1.0.13.0`.

Post-change verification:
- Launcher parser OK.
- Confirmed source references for `Resolve-PowerShell7Path`, `PowerShell 7 is required`, `Microsoft Outlook is required`, and `Checking required programs`.
- Source launcher `-NoGui -UseSampleData` run generated HTML/log and `CONVERSION_RESULT`.
- Release EXE version verified as `1.0.13.0`.
- Debug EXE sample-data run passed.
- Release EXE sample-data run passed.
- HTML smoke checks passed: title, Fireflies.ai, Exact-Conversation, resize handle, HTML encoding.
- Log smoke checks passed: `exported=3`, `itemReadFailures=0`.

## Change: Handoff hardening pass

Requested change: review the handoff solutions, look for gaps, debug, harden, and optimize.

Implementation:
- Added core output-path safety validation so the HTML report and log file cannot target the same file, cannot point at a folder, and must have valid parent paths.
- Added GUI-side output/log validation before conversion starts, with clear messages for blank paths, same-file paths, wrong extensions, folder paths, or invalid parent paths.
- Added a form-closing guard during active conversion. If the user closes the GUI while a conversion is running, the app asks for confirmation; cancel keeps the conversion running, confirm stops the child converter and cleans its temp embedded-core directory.
- Hardened generated report JavaScript so filter-column resize persistence uses safe localStorage wrappers. If a browser blocks localStorage for local files, filtering and report startup should continue instead of aborting the script.
- Hardened hidden `-NoGui` error handling so packaged validation failures return process exit code `1` and write a clear stderr message instead of only emitting an error while returning exit code `0`.
- Re-embedded the updated core converter into the launcher.
- Rebuilt debug and release EXEs as version `1.0.14.0`.
- Updated README version and removed stale GUI sample-data instructions.

Post-change verification:
- Core parser OK.
- Launcher parser OK.
- Source launcher `-NoGui -UseSampleData` run passed.
- Source same-output/log-path guard failed as expected with the clear `must be different files` message.
- Debug EXE version verified as `1.0.14.0`.
- Release EXE version verified as `1.0.14.0`.
- Debug EXE sample-data run passed.
- Release EXE sample-data run passed.
- Packaged Debug EXE same-output/log-path guard returned exit code `1` and stderr contained `must be different files`.
- HTML smoke checks passed: title, Fireflies.ai, Exact-Conversation, resize handle, HTML encoding, safe localStorage wrapper.
- Log smoke checks passed: `exported=3`, `itemReadFailures=0`.

## Change: Port participant-match improvements from standalone script

Requested change: apply the same recent improvement from `C:\Users\pbush\Downloads\Convert-PurviewTeamsPstToHtml_1.0.ps1` to the application's underlying PowerShell converter.

Implementation:
- Compared the improved Downloads script against the app core converter.
- Ported the report filter improvement from a single `Exact-Conversation` checkbox to a `Participant match` dropdown with four modes:
  - `Messages from selected people`
  - `Conversations involving selected people`
  - `Exact selected people only`
  - `Exact selected people, ignoring Other/IDs`
- Added `data-person-participants` and `data-other-participants` attributes to each rendered conversation so the browser-side filter can ignore bots, meeting artifacts, and system IDs when requested.
- Preserved the previous hardening from version `1.0.14.0`, including safe localStorage wrappers, output/log path guards, and packaged `-NoGui` nonzero error exits.
- Re-embedded the updated core converter into the launcher.
- Rebuilt debug and release EXEs as version `1.0.15.0`.

Post-change verification:
- Core parser OK.
- Launcher parser OK.
- Source launcher `-NoGui -UseSampleData` run passed.
- Debug EXE version verified as `1.0.15.0`.
- Release EXE version verified as `1.0.15.0`.
- Debug EXE sample-data run passed.
- Release EXE sample-data run passed.
- HTML smoke checks passed: title, Fireflies.ai, participant match dropdown, all four participant modes, `data-person-participants`, `data-other-participants`, old `exactConversationOnly` checkbox removed, resize handle, safe HTML encoding, safe localStorage wrapper.
- Log smoke checks passed: `exported=3`, `itemReadFailures=0`.
- Packaged Debug EXE same-output/log-path guard still returned exit code `1` and stderr contained `must be different files`.


## Change: Add date filtering and conversation sorting

Requested change: try adding abilities to sort and filter by date in the generated HTML report.

Implementation:
- Added Start date and End date controls to the report filter panel.
- Added a Sort conversations dropdown with Newest first and Oldest first options.
- Added per-message `data-date` and `data-time` attributes based on each message SortTime.
- Added per-conversation `data-sort-time` based on the latest message in that conversation.
- Updated browser-side filtering so date range filters combine with participant and text filters.
- Updated browser-side sorting so conversation cards reorder in the report without regenerating the HTML.
- Re-embedded the updated core converter into the launcher.
- Rebuilt debug and release EXEs as version `1.0.17.0`.

Post-change verification:
- Core parser OK.
- Launcher parser OK.
- Source launcher `-NoGui -UseSampleData` run passed.
- Debug EXE version verified as `1.0.17.0`.
- Release EXE version verified as `1.0.17.0`.
- Debug EXE sample-data run passed.
- Release EXE sample-data run passed.
- HTML smoke checks passed: Start date, End date, Sort conversations, Newest first, Oldest first, `data-date`, `data-time`, `data-sort-time`, `messageInDateRange`, `sortConversations`, participant-match mode, safe HTML encoding.
- Log smoke checks passed: `exported=3`, `itemReadFailures=0`.


## Change: Outlook busy/RPC retry and progress warning fix

Observed failure: GUI reported `Converter exited with code 1. See log file for details.` The latest real PST log `C:\Users\pbush\Downloads\LArtley Teams Messages.log` showed the root cause: `Fatal error: Call was rejected by callee. (0x80010001 (RPC_E_CALL_REJECTED))` before the PST attach step.

Implementation:
- Added `Test-OutlookRpcRejected` and `Invoke-OutlookComOperation` helpers.
- Wrapped startup-sensitive Outlook COM operations with retry/clear-message handling: creating `Outlook.Application`, opening the MAPI namespace, attaching the PST, and locating the attached PST store.
- If Outlook keeps rejecting automation calls after retries, the converter now logs a clearer fatal message telling the user to close Outlook completely, wait a few seconds, and run again.
- Improved GUI failure reporting: when the child converter exits nonzero, the message box includes the latest `Fatal error:` log line when available instead of only the generic exit-code message.
- Fixed progress-path sanitization to avoid regex replacement. Prior real PST logs had repeated warnings: `The regular expression pattern \ is not valid.` The code now uses literal string replacement for backslashes and pipe characters.
- Re-embedded the updated core converter into the launcher.
- Rebuilt debug and release EXEs as version `1.0.18.0`.

Post-change verification:
- Core parser OK.
- Launcher parser OK.
- Direct core `-UseSampleData -LogEvery 1` run passed with no regex warning.
- Debug EXE version verified as `1.0.18.0`.
- Release EXE version verified as `1.0.18.0`.
- Debug EXE sample-data run passed.
- Release EXE sample-data run passed.
- Same-output/log-path guard still returned exit code `1` and stderr contained `must be different files`.
- HTML smoke checks passed: report title, date/sort controls, participant match mode, safe HTML encoding.
- Log smoke checks passed: `exported=3`, `itemReadFailures=0`, no regex warning.


## Change: Move date controls above conversations

Requested change: move the date filter and sort controls from the left filter column and put them above the conversations.

Implementation:
- Removed Start date, End date, and Sort conversations controls from the left filter panel.
- Added a `conversation-toolbar` section at the top of the conversation pane, before the first conversation card.
- Kept the same control IDs (`startDateFilter`, `endDateFilter`, `sortOrder`) so existing JavaScript date filtering and sorting behavior continues to work.
- Added toolbar styling and mobile stacking behavior.
- Re-embedded the updated core converter into the launcher.
- Rebuilt debug and release EXEs as version `1.0.19.0`.

Post-change verification:
- Core parser OK.
- Launcher parser OK.
- Source launcher `-NoGui -UseSampleData` run passed.
- Debug EXE version verified as `1.0.19.0`.
- Release EXE version verified as `1.0.19.0`.
- Debug EXE sample-data run passed.
- Release EXE sample-data run passed.
- HTML placement checks passed: toolbar present, toolbar before first conversation, Start date/End date/Sort conversations in toolbar, those controls absent from left filter panel.
- Existing behavior checks passed: participant match mode present, date filtering/sorting JS present, `exported=3`, no regex warning.


## Change: Add GUI credit line

Requested change: add a small line in the bottom-right corner of the GUI that says `By Patrick Bush`.

Implementation:
- Added a small dim-gray WinForms label with text `By Patrick Bush` near the bottom-right of the launcher window.
- Anchored the label to bottom-right so it stays in that corner when the window is resized.
- Rebuilt debug and release EXEs as version `1.0.20.0`.

Post-change verification:
- Core parser OK.
- Launcher parser OK.
- Source checks passed: credit text present, bottom-right anchor present, small font present.
- Debug EXE version verified as `1.0.20.0`.
- Release EXE version verified as `1.0.20.0`.
- Debug EXE sample-data run passed.
- Release EXE sample-data run passed.
- Existing generated-report checks passed: date toolbar still present, `exported=3`, no regex warning.


## Change: Narrow top report rows

Requested change: make the top three horizontal report rows slightly narrower: the blue banner, the row that starts with the PST file name, and the date filter/sort row.

Implementation:
- Added CSS variable `--top-row-inset: 18px`.
- Applied the inset to `.hero`, making the blue banner slightly narrower and centered.
- Applied the inset to `.summary-grid`, making the PST/generated/messages summary row slightly narrower and centered.
- Applied the inset to `.conversation-toolbar`, making the Start date / End date / Sort conversations row slightly narrower and centered above the conversations.
- Re-embedded the updated core converter into the launcher.
- Rebuilt debug and release EXEs as version `1.0.21.0`.

Post-change verification:
- Core parser OK.
- Launcher parser OK.
- Source launcher `-NoGui -UseSampleData` run passed.
- Debug EXE version verified as `1.0.21.0`.
- Release EXE version verified as `1.0.21.0`.
- Debug EXE sample-data run passed.
- Release EXE sample-data run passed.
- HTML CSS checks passed for `--top-row-inset`, narrowed `.hero`, narrowed `.summary-grid`, and narrowed `.conversation-toolbar`.
- Existing placement/behavior checks passed: toolbar remains above first conversation, date/sort controls still present, participant match mode still present, GUI credit remains embedded, `exported=3`, no regex warning.


## Change: Restore full row width and reduce top row heights

Requested correction: undo the previous horizontal narrowing; the desired change was to slightly reduce the height of the top rows, not their horizontal width.

Implementation:
- Removed the `--top-row-inset` CSS variable and all related `width: calc(...)` horizontal narrowing rules.
- Restored full horizontal width for the blue hero banner, the PST/generated/messages summary row, and the date/sort toolbar.
- Slightly reduced vertical height by changing:
  - `.hero` padding from `18px 22px` to `14px 22px`.
  - `.summary-card` padding from `10px 12px` to `7px 12px`.
  - `.summary-grid` vertical margin from `12px` to `10px`.
  - `.conversation-toolbar` padding from `12px 14px` to `8px 14px` and bottom margin from `14px` to `12px`.
- Re-embedded the updated core converter into the launcher.
- Rebuilt debug and release EXEs as version `1.0.22.0`.

Post-change verification:
- Core parser OK.
- Launcher parser OK.
- Source launcher `-NoGui -UseSampleData` run passed.
- Debug EXE version verified as `1.0.22.0`.
- Release EXE version verified as `1.0.22.0`.
- Debug EXE sample-data run passed.
- Release EXE sample-data run passed.
- HTML CSS checks passed: horizontal inset removed, hero/summary/toolbar restored to full width, reduced vertical paddings/margins present.
- Existing behavior checks passed: toolbar remains above first conversation, date/sort controls still present, participant match mode still present, GUI credit remains embedded, `exported=3`, no regex warning.


## Change: Add HTML blue-banner credit line

Requested change: add a small line of text to the bottom-right corner of the blue banner at the top of the generated HTML report reading `By Patrick Bush`.

Implementation:
- Added `.hero-credit` CSS with absolute bottom-right positioning inside the blue hero banner.
- Set `.hero` to `position: relative` so the credit anchors to the banner.
- Added `<div class='hero-credit'>By Patrick Bush</div>` inside the generated report hero header.
- Re-embedded the updated core converter into the launcher.
- Rebuilt debug and release EXEs as version `1.0.23.0`.

Post-change verification:
- Core parser OK.
- Launcher parser OK.
- Source launcher `-NoGui -UseSampleData` run passed.
- Debug EXE version verified as `1.0.23.0`.
- Release EXE version verified as `1.0.23.0`.
- Debug EXE sample-data run passed.
- Release EXE sample-data run passed.
- HTML checks passed: banner credit present, credit inside hero header, `.hero-credit` bottom-right CSS present, `.hero` relative positioning present.
- Existing behavior checks passed: date/sort controls still present above conversations, participant match mode still present, GUI credit remains embedded, `exported=3`, no regex warning.


## Change: Reduce conversation subject header and date toolbar height/text

Requested change: slightly reduce the height of the subject box above each conversation, slightly reduce the text size in the date sort/filter toolbar, and slightly reduce the toolbar height.

Implementation:
- Reduced conversation subject/header height by lowering `.conversation-header` padding and gap.
- Slightly reduced conversation subject text and participant/count text sizes.
- Split date toolbar input/select styling from the left-panel controls so only the toolbar text/height changed.
- Reduced date toolbar input/select padding and font size.
- Reduced date toolbar grid gap, padding, field gap, label size, and bottom margin.
- Re-embedded the updated core converter into the launcher.
- Rebuilt debug and release EXEs as version `1.0.24.0`.

Post-change verification:
- Core parser OK.
- Launcher parser OK.
- Source launcher `-NoGui -UseSampleData` run passed.
- Debug EXE version verified as `1.0.24.0`.
- Release EXE version verified as `1.0.24.0`.
- Debug EXE sample-data run passed.
- Release EXE sample-data run passed.
- HTML CSS checks passed: smaller toolbar inputs/selects, shorter toolbar, smaller toolbar labels, shorter conversation header/subject box, smaller subject/participant/count text.
- Existing behavior checks passed: date/sort controls still present above conversations, blue-banner credit remains present, participant match mode still present, `exported=3`, no regex warning.


## Change: Make GUI progress bar psychologically honest

Requested change: improve GUI progress bar accuracy in the psychological sense, without a significant conversion slowdown.

Implementation:
- Avoided a PST pre-scan, so conversion time is not meaningfully increased.
- Kept the existing asynchronous child-process + WinForms timer pattern.
- Added lightweight converter log markers for post-read stages:
  - `Preparing HTML report data.`
  - `Writing HTML report: 0 of N conversations.`
  - `HTML report progress: X of N conversations.`
  - `Finalizing HTML report.`
- Added throttled report-writing progress every 50 conversation groups and at completion.
- Changed GUI progress mapping from the old collected-count formula that jumped toward ~80% early to conservative stage ranges:
  - initial/read ramp: around 10–30%.
  - collected-message reading progress: capped around 58%.
  - finished reading/preparing report: 60–62%.
  - writing report: 66%, then 68–91% based on conversation groups written.
  - finalizing/cleanup: 92–96%.
  - complete: 100% on successful child-process exit.
- Re-embedded the updated core converter into the launcher.
- Rebuilt debug and release EXEs as version `1.0.25.0`.

Post-change verification:
- Core parser OK.
- Launcher parser OK.
- Source launcher `-NoGui -UseSampleData` run passed.
- Debug EXE version verified as `1.0.25.0`.
- Release EXE version verified as `1.0.25.0`.
- Debug EXE sample-data run passed.
- Release EXE sample-data run passed.
- Log marker checks passed for preparing report, writing report start, writing progress, and finalizing report.
- GUI source checks passed for conservative collected cap, lower initial ramp, and staged report progress mapping.
- Existing behavior checks passed: date/sort controls still present, blue-banner credit remains present, participant match mode still present, `exported=3`, no regex warning.


## Change: Fix stale-log progress jump regression

Reported issue after the psychological progress-bar pass: the progress bar immediately jumped past 90% and showed cleanup.

Root cause:
- GUI progress is driven by tailing the converter log.
- If the selected log file already contained final-stage lines from a previous run (`HTML report written...`, `Detaching PST...`, etc.), the timer could read those stale lines before the new child process wrote fresh log entries.
- That made the GUI think it was already in cleanup/final stages.

Implementation:
- GUI now removes an existing log file immediately before starting the child converter process.
- Converter core now clears the log file before writing the first log entry for every run.
- Initial GUI progress before child launch was lowered from 25% to 12% (`Starting PST read...`) so the bar no longer jumps high before real progress exists.
- Re-embedded the updated core converter into the launcher.
- Rebuilt debug and release EXEs as version `1.0.26.0`.

Post-change verification:
- Core parser OK.
- Launcher parser OK.
- Stale-log regression test passed for source launcher, debug EXE, and release EXE: each test pre-populated the target log with stale cleanup/final lines, ran sample conversion, and verified stale content was gone.
- Debug EXE version verified as `1.0.26.0`.
- Release EXE version verified as `1.0.26.0`.
- Progress marker checks still passed: preparing report, writing report start, writing progress, and finalizing report.
- GUI source checks passed: removes old log, core clears log, starts PST read at 12%, conservative progress mapping remains.
- Existing behavior checks passed: date/sort controls still present, blue-banner credit remains present, `exported=3`, no regex warning.
