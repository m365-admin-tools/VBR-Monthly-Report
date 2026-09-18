# VBR-Monthly-Report
Generates a monthly backup session summary for Veeam Backup &amp; Replication. It collects every backup, backup copy, replication, and agent job session that finished in a given calendar month and produces a formatted HTML report with an executive summary, a per-job rollup, a list of sessions that need attention, and the full session detail table.

# Veeam Monthly Backup Report (PowerShell)

Written for MSPs and backup administrators who have to produce the same client-facing or management-facing report at the start of every month.

Please visit https://m365admintools.com for more IT engineering tools

<!-- Add a screenshot of the HTML report here, then uncomment:
![Monthly report](docs/images/monthly-report.png)
-->

## Requirements

| Item | Requirement |
|---|---|
| Veeam Backup & Replication | v12 or newer. The script uses the `Veeam.Backup.PowerShell` module and does not fall back to the legacy snap-in |
| Where to run it | On the VBR server, or on a machine with the VBR console installed |
| Rights | Run elevated, with an account holding a Veeam Backup role that can read session history |
| Session history | The reporting month must still be inside the session history retention configured in VBR. Older months return nothing |

## Quick start

```powershell
# Previous calendar month, HTML written to C:\Reports
.\VBR-Monthly-Report.ps1

# A specific month, with the session detail also written to CSV
.\VBR-Monthly-Report.ps1 -Month 6 -Year 2026 -OutputPath D:\Reports -ExportCsv

# A remote VBR server
.\VBR-Monthly-Report.ps1 -Server vbr01.contoso.local
```

If the script is blocked on first run:

```powershell
Unblock-File .\VBR-Monthly-Report.ps1
```

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-Month` | int 1-12 | Previous calendar month | Target month |
| `-Year` | int | Year of the previous calendar month | Target year |
| `-OutputPath` | string | `C:\Reports` | Output folder. Created if it does not exist |
| `-ExportCsv` | switch | Off | Also write the session detail table to CSV |
| `-Server` | string | `localhost` | VBR server to connect to |

With no parameters the script reports on the previous calendar month, which is what makes it suitable for an unattended run on the first of the month.

## What the report contains

**Summary cards.** Total sessions, success rate, success count, warning count, failure count, and total GB transferred.

**Per-job rollup.** One row per job name with job type, number of runs, success, warning, and failure counts, success rate, average duration in minutes, and total GB transferred. This is the table most clients read.

**Sessions requiring attention.** Every session that did not end in Success, most recent first. When the month was clean, the section states that instead of showing an empty table.

**All sessions.** The full detail table: job name, job type, start, end, duration in minutes, result, GB processed, and GB transferred. Result cells are colour coded green, amber, and red.

## Output files

```
VBR-Monthly-Report-2026-06.html
VBR-Monthly-Sessions-2026-06.csv     (only with -ExportCsv)
```

The HTML file is self-contained with inline styling, so it can be emailed as a single attachment or printed to PDF from the browser. Re-running the same month overwrites the existing file, which keeps a scheduled run from filling the folder with duplicates.

## Running it monthly

Create a scheduled task on the VBR server that runs on the first day of each month. Because the script defaults to the previous month, no date parameters are needed.

```
Program:   powershell.exe
Arguments: -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\VBR-Monthly-Report.ps1" -ExportCsv
```

Set the task to run whether the user is logged on or not, with highest privileges, using an account that holds a Veeam Backup role.

## Scope

| Included | Not included |
|---|---|
| Backup jobs | SureBackup sessions |
| Backup copy jobs | Tape job sessions |
| Replication jobs | File share and NAS backup sessions |
| Agent jobs for physical servers and workstations | Configuration backup sessions |

Sessions are selected by end time, so a job that starts on the last night of the month and finishes after midnight is counted in the following month.

## Known limitations

- **Agent sessions report zero data volume.** Processed and transferred GB are recorded as 0 for agent jobs, so those jobs appear in the session counts and success rate but contribute nothing to the GB transferred figure. Treat the data volume as covering hypervisor-based jobs only.
- **The script disconnects from the VBR server when it finishes.** If it is run from a console session that was already connected, that connection is closed too. This matters only for interactive use.
- **No email delivery.** The report is written to disk. Sending it is left to the scheduler, a follow-up script, or your RMM.
- **A point-in-time view of history.** The report summarizes what the VBR configuration database recorded. A high success rate is not evidence that a restore works, and the report is not a substitute for a tested restore.

## Troubleshooting

**`Could not load Veeam.Backup.PowerShell. Confirm VBR v12+ console is installed.`**

The module was not found. Install the VBR console on this machine, or run the script on the VBR server. Confirm with:

```powershell
Get-Module -ListAvailable Veeam.Backup.PowerShell
```

**`No sessions found for <month>. Check retention on session history.`**

Either no jobs ran in that month, or the session history has aged out. Session history retention is set in the VBR console under Options, History. Increase it if you need to report on older months, and note that the change is not retroactive.

**Connection fails against a remote server**

The script calls `Connect-VBRServer -Server <name>` with no credential parameter, so it connects as the calling account. Run it as an account that holds a Veeam Backup role on the target server, or run it locally on the VBR server.

**Success rate looks lower than expected**

Every session counts equally, including retries. A job that failed once and succeeded on retry contributes one failure and one success. Check the per-job rollup rather than the headline rate when investigating.

## Related

- Free Microsoft 365, Active Directory, and Veeam tools at [m365admintools.com](https://m365admintools.com)

## Author

Charles Arconi, [m365admintools.com](https://m365admintools.com)

Not affiliated with, endorsed by, or supported by Veeam Software. Veeam is a trademark of Veeam Software Group GmbH.

## License

MIT. See [LICENSE](LICENSE).
