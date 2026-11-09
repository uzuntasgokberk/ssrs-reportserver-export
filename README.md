# SSRS and Power BI Report Server Export

[![Validate](https://github.com/uzuntasgokberk/ssrs-reportserver-export/actions/workflows/validate.yml/badge.svg)](https://github.com/uzuntasgokberk/ssrs-reportserver-export/actions/workflows/validate.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
![PowerShell 5.1 | 7](https://img.shields.io/badge/PowerShell-5.1%20%7C%207-blue)

Back up the content of a SQL Server Reporting Services (SSRS) or Power BI Report Server as files you can open, diff and keep in Git. It's a single PowerShell script that uses the report server's own REST API. No module and no extra tool to install.

| Report server | What is exported |
|---|---|
| SSRS 2017 or later | **Paginated content only:** reports (`.rdl`), shared data sources (`.rds`), shared datasets (`.rsd`) and report parts (`.rsc`). SSRS cannot host Power BI reports, so there is never a `.pbix` |
| Power BI Report Server | The same, **plus Power BI reports** (`.pbix`) |
| Power BI service / Microsoft Fabric | Not covered here: use [fabric-powerbi-export](https://github.com/uzuntasgokberk/fabric-powerbi-export) |

| On the report server | On disk |
|---|---|
| Report (paginated) | `.rdl` file, byte for byte as stored on the server |
| Shared dataset | `.rsd` file, byte for byte |
| Report part | `.rsc` file, byte for byte |
| Shared data source | `.rds` file (the format Visual Studio report projects use), **without passwords** |
| Power BI report (Power BI Report Server only) | `.pbix` file, byte for byte |

## Features

- **Same folder tree as the server.** `/Finance/Monthly/Invoices` becomes `Finance\Monthly\Invoices.rdl`.
- **Pick what you need.** Use `-Folder`, `-Name` and `-Exclude` (wildcards), `-Type`, or a dry run with `-List`.
- **Read-only.** Only GET requests are sent; nothing on the report server is changed.
- **No passwords in the export.** The report server never returns stored passwords. A password written into a data source connection string is masked, and one written into an embedded connection string of a report is reported.
- **Safe overwrite.** A file is replaced only after its item was downloaded completely, so a failed run never destroys the last good export.
- **Nothing is deleted.** After a full run, local files whose item no longer exists on the server are listed as warnings.
- **Built for unattended runs.** It retries network errors, writes a log and a CSV for every run, and returns exit codes that schedulers can check.
- **Windows PowerShell 5.1 and PowerShell 7.** Windows sign-in (Kerberos / NTLM) works in both, also over http.

## Output

```text
export\
├── _logs\                          (log + CSV of every run)
├── Data Sources\
│   └── Sales DW.rds
├── Datasets\
│   └── Regions.rsd
├── Report Parts\
│   └── DateFilter.rsc
├── Finance\
│   ├── Invoices.rdl
│   └── Monthly\
│       └── Revenue by Region.rdl
└── Sales\
    ├── Sales Dashboard.pbix        (Power BI Report Server only)
    └── ...
```

- Characters that Windows does not allow in file names (`\ / : * ? " < > |`) are replaced with `_`, on every platform.
- If two items still end up with the same file name, or their names differ only in case, they do not overwrite each other. Each gets a short ID suffix: `Invoices (1a2b3c4d).rdl`.

## Requirements

| Requirement | Details |
|---|---|
| Report server | **SSRS 2017 or later**, or **Power BI Report Server** (both have the REST API v2.0). Tested with SSRS 2019 and Power BI Report Server 1.26 |
| PowerShell | Windows PowerShell 5.1 or PowerShell 7+ |
| Permissions | The **Content Manager** or **Publisher** role on the items you export. The **Browser** role can view reports but cannot download their files; with it you get HTTP 403 |
| Sign-in | The current Windows account by default, or `-Credential` |

Tested on Windows with PowerShell 5.1 and 7, signed in with the current Windows account. Signing in with `-Credential` (only a rejected sign-in was tested), Linux and macOS have not been tested yet.

## Quick start

```powershell
git clone https://github.com/uzuntasgokberk/ssrs-reportserver-export.git
cd ssrs-reportserver-export

.\Export-SsrsContent.ps1 -ServerUrl "http://reportserver/Reports" -List    # dry run: what would be exported
.\Export-SsrsContent.ps1 -ServerUrl "http://reportserver/Reports"          # export to .\export
```

Use the web portal address you see in the browser. A URL copied from the browser (`.../Reports/browse/`), the web service address (`.../ReportServer`) or just the host also work.

If PowerShell refuses to run the script ("running scripts is disabled on this system"), you do not need an administrator window. Allow local scripts once for your user:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
Unblock-File .\Export-SsrsContent.ps1      # only needed if the file was downloaded as a ZIP
```

or start it with `powershell.exe -ExecutionPolicy Bypass -File .\Export-SsrsContent.ps1 ...`.

## Usage

| Parameter | Description | Default |
|---|---|---|
| `-ServerUrl` | Web portal address, e.g. `http://reportserver/Reports` | Required |
| `-Folder` | Only these server folders, subfolders included. Folder name or path; wildcards allowed. A folder name without `/` matches at any depth | All |
| `-Name` | Only items with these names (wildcards allowed) | All |
| `-Exclude` | Skip items **or folders** with these names (wildcards allowed). A folder name without `/` matches at any depth | None |
| `-Type` | `Report`, `DataSource`, `DataSet`, `ReportPart`, `PowerBIReport` (one or more). SSRS has no Power BI reports, so there `PowerBIReport` finds nothing | All |
| `-Credential` | Account to sign in with | Current Windows account |
| `-List` | List only, do not download | Off |
| `-ExportRoot` | Output folder | `.\export` |

```powershell
$Server = "http://reportserver/Reports"

# Everything into a separate folder (for example your backup repository)
.\Export-SsrsContent.ps1 -ServerUrl $Server -ExportRoot "D:\ssrs-backup"

# One folder with its subfolders, without the archive
.\Export-SsrsContent.ps1 -ServerUrl $Server -Folder "Finance" -Exclude "Archive"

# Only shared data sources and datasets
.\Export-SsrsContent.ps1 -ServerUrl $Server -Type DataSource, DataSet

# Power BI Report Server: only the Power BI reports, or everything except them
.\Export-SsrsContent.ps1 -ServerUrl $Server -Type PowerBIReport
.\Export-SsrsContent.ps1 -ServerUrl $Server -Type Report, DataSource, DataSet, ReportPart

# One report (copy the name from the -List output)
.\Export-SsrsContent.ps1 -ServerUrl $Server -Name "Revenue by Region"

# Another account (e.g. from a machine that is not in the domain)
.\Export-SsrsContent.ps1 -ServerUrl "https://reports.contoso.com/Reports" -Credential (Get-Credential)
```

A filtered run (`-Folder`, `-Name`, `-Exclude`, `-Type`) updates only the selected files and leaves all other files alone.

With `powershell.exe -File ...`, a comma-separated list (`-Folder "A","B"`) arrives as a single value. Pass several values from inside a PowerShell window instead.

### Exit codes

| Code | Meaning |
|---|---|
| `0` | All selected items succeeded |
| `1` | At least one item failed (details in the log/CSV), or nothing matched the filter |
| `2` | Could not start: invalid `-ServerUrl`, or the report server could not be reached |

## How it works

1. **Connection test.** `GET /api/v2.0/System` returns the product and version, which are written to the log together with the account in use.
2. **List.** `GET /api/v2.0/CatalogItems` returns every item on the server in one request. The filters are applied locally.
3. **Download.** Each item is downloaded into memory:
   - Reports, datasets, report parts and Power BI reports come from `CatalogItems({id})/Content/$value`, unchanged.
   - Shared data sources come from `DataSources({id})`. They are written as `.rds` with the extension, the connection string and the credential mode (integrated / prompt). The server ID is used as `DataSourceID`, so the file stays the same between runs.
4. **Save.** Only a completely downloaded item is written to disk.
5. **Stale file check.** After a full (unfiltered) run, local `.rdl` / `.rds` / `.rsd` / `.rsc` / `.pbix` files without an item on the server are listed. They are never deleted.
6. **Summary.** The summary is printed, and a log and a CSV are written to `_logs`.

Requests are sent with `HttpWebRequest` rather than `Invoke-RestMethod`, because PowerShell 7 refuses to send Windows credentials over plain http. This is also why the script behaves the same in 5.1 and 7.

## Using the exported files

- **Report Builder** opens `.rdl` files. If a report uses a shared data source on the server (`/Data Sources/...`), connect Report Builder to the server for the preview: **Connect** at the bottom left → `http://reportserver/ReportServer`.
- **Visual Studio** (Reporting Services Projects extension): create a Report Server Project and add the files with **Add Existing Item** (`.rds` → Shared Data Sources, `.rsd` → Shared Datasets, `.rdl` → Reports). The `.rds` files contain no password; enter the credentials in the data source properties.
- **Power BI Desktop optimized for Power BI Report Server** opens `.pbix` files. Use the release that matches your report server.
- **Restoring to a server** changes the server, and the script never does that. Use **Upload** in the web portal or **Deploy** in Visual Studio. Afterwards, bind the shared data sources on the report's **Manage → Data sources** page, and enter stored credentials again. For a Power BI report, also enter its data source credentials and create its scheduled refresh plans again.

## Logs and troubleshooting

Every run writes `_logs/export_YYYYMMDD_HHMMSS.log` and a `.csv` with one row per item. CSV columns: server path, type, ID, status, downloaded bytes, size on the server, duration, local file, message.

Each log line shows:
- the position (`[12/243]`)
- the type and server path
- the duration
- the shared data sources and datasets the report uses

| Message | Cause | Fix |
|---|---|---|
| `HTTP 401` | The report server did not accept the account | Check the network / VPN. From a machine outside the domain, pass `-Credential`; with `-Credential`, check the user name (`DOMAIN\user`) and the password |
| `HTTP 403` | No permission on the item | Ask for the Content Manager or Publisher role on that folder |
| `HTTP 404` on the connection test | Wrong `-ServerUrl`, or the server is older than SSRS 2017 (no REST API v2.0) | Use the web portal address from the browser |
| `Network error` | The report server could not be reached | Each request is tried 3 times, 15 s apart; check the network / VPN |
| `The response is not JSON` | The address is not the API (e.g. a sign-in page) | Pass the web portal address as `-ServerUrl` |
| `Could not write to disk` | The file is open in another program | Close Report Builder / Visual Studio and run again |
| `WARNING: An embedded connection string contains a password` | A plain-text password is written inside a report or dataset | Review the file before committing it, and remove the password on the server |
| `WARNING: The file is larger than 100 MB` | A `.pbix` file is bigger than GitHub accepts | Keep such files with [Git LFS](https://git-lfs.com), or leave them out of Git (see [Keeping exports in Git](#keeping-exports-in-git)) |
| `Not on the server any more` | A local file has no item on the server | If the item was really deleted, delete the file (`git rm "<file>"`) |
| `Unsupported type, skipped` | A type this script does not export (see [Limitations](#limitations)) | Informational |

The CSV can contain the user names of stored credentials. Review it before you share it.

## Unattended runs

The script signs in with the account it runs under. Run the scheduled task as a service account (or a gMSA) that has the Content Manager role, and no password has to be stored anywhere.

Example for Windows Task Scheduler:

```text
Program:   powershell.exe
Arguments: -NoProfile -ExecutionPolicy Bypass -File "C:\tools\ssrs-reportserver-export\Export-SsrsContent.ps1" -ServerUrl "http://reportserver/Reports" -ExportRoot "D:\ssrs-backup"
```

A non-zero exit code marks the task run as failed.

If you must use `-Credential` in a scheduled task, store it encrypted for the task's account with `Get-Credential | Export-Clixml` and load it with `Import-Clixml`. Never put passwords in the command line or in a repository.

## Keeping exports in Git

Copy the two files from [`templates/export-repo`](templates/export-repo) to the root of your export folder:

- **`.gitignore`:** Leaves out the logs, plus database and encryption key backups in case they end up there by mistake.
- **`.gitattributes`:** Normalizes line endings. Report Builder and Visual Studio save CRLF, while the server returns whatever was uploaded; Git stores both the same way, so a diff shows only real changes, never whole files. After an export Git may print `LF will be replaced by CRLF`; that is only information.

For later backups, run the export again, then `git add .`, `git commit` and `git push`; Git only sends the changed files.

`.pbix` files (Power BI Report Server) are binary: a diff shows only that the file changed, and Git keeps a full copy of every changed version. GitHub rejects files over 100 MB; the script warns about them. Keep large `.pbix` files with [Git LFS](https://git-lfs.com), or leave all of them out with the commented `*.pbix` line in the template `.gitignore`.

Use a **private** repository: the files contain server and database names, queries and connection strings, and a `.pbix` with imported data contains that data.

## What this export does not cover

The export contains the item definitions. Everything else lives in the report server databases:
- subscriptions and schedules
- scheduled refresh plans of Power BI reports
- folder and item permissions
- cache, snapshot and history settings
- report settings changed in the web portal (e.g. parameter defaults, the data source a report is bound to)
- stored credentials

For a complete disaster-recovery backup, also back up:

| What | How |
|---|---|
| `ReportServer` database | Regular SQL Server backups (full, plus log backups in the full recovery model). If someone already backs it up, use `COPY_ONLY` for extra backups so you don't break their backup chain |
| `ReportServerTempDB` database | Optional; it can be recreated. Backing it up makes a restore simpler |
| Encryption key | Once, and again after the service account changes: Report Server Configuration Manager → **Encryption Keys** → **Backup**, or `rskeymgmt -e -f <file>.snk -p <password>` on the report server. Without it, the stored credentials in a restored database cannot be decrypted. Never use `rskeymgmt -s` (re-creates the key) or `-d` (deletes encrypted content) for a backup |
| Configuration files | `rsreportserver.config`, `rssrvpolicy.config` and `web.config` of the report server |

Keep database and key backups out of Git, in a secured location. Keep the key password in a password manager.

## Limitations

- **Unsupported types:** linked reports, resources (images, PDF, Excel), Excel workbooks, KPIs and mobile reports are not exported; they are listed as `Unsupported type` in the log.
- **Power BI reports stay `.pbix`.** The report server stores only the `.pbix` file, so there is no PBIP / TMDL text version to diff. For that, see [fabric-powerbi-export](https://github.com/uzuntasgokberk/fabric-powerbi-export) (Power BI service only).
- **Deleted or renamed items** are not removed from an existing export; the script only warns about them.
- **Personal folders:** "My Reports" content is exported only if it is visible to the account in the catalog.
- **No data**, with one exception: the export contains definitions only, but a `.pbix` with imported data contains that data.

## Related projects

- [fabric-powerbi-export](https://github.com/uzuntasgokberk/fabric-powerbi-export): the same idea for the Power BI service / Microsoft Fabric (PBIP, TMDL and RDL), using the Fabric CLI.

## Contributing

Issues are welcome; pull requests are by invitation. [CONTRIBUTING.md](CONTRIBUTING.md) describes the process and the checks to run. When reporting a bug, include the relevant lines from the log file and remove server, folder or item names you don't want to share. Report security problems privately, never in an issue: [SECURITY.md](SECURITY.md).

## License

[MIT](LICENSE) © uzuntasgokberk

This project is not affiliated with or endorsed by Microsoft. SQL Server, SQL Server Reporting Services and Power BI are trademarks of the Microsoft group of companies.
