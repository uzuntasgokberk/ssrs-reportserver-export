# Changelog

All notable changes to this project are documented in this file.
The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project adheres to [Semantic Versioning](https://semver.org/).

## [1.0.0] - 2026-11-02

### Added

- Export of reports (`.rdl`), shared datasets (`.rsd`) and report parts (`.rsc`) byte for byte as stored on the report server.
- Export of shared data sources as standard `.rds` files built from their settings, without passwords.
- Export of Power BI reports (`.pbix`) from Power BI Report Server, byte for byte as stored on the server (`-Type PowerBIReport`); a warning for files larger than 100 MB.
- The folder tree of the report server is kept on disk.
- Filters: `-Folder`, `-Name` and `-Exclude` (wildcards; `-Exclude` also skips whole folders), `-Type`, `-List` (dry run).
- Sign-in with the current Windows account, or with `-Credential`.
- Read-only: only GET requests are sent to the report server.
- Retries on network errors, timeouts and HTTP 503.
- Safe overwrite: a file is replaced only after its item was downloaded completely.
- Items that differ only after invalid characters are replaced get a short ID suffix instead of overwriting each other.
- Local files whose item no longer exists on the server are reported after a full run, never deleted.
- Passwords in data source connection strings are masked; plain-text passwords in embedded connection strings are reported.
- Log file and CSV summary per run; exit codes for schedulers.
- Templates (`.gitignore`, `.gitattributes`) for keeping exports in Git; `.gitattributes` normalizes line endings, so files saved with LF or CRLF show only real changes.
