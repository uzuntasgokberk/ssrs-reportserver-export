# Security policy

## Report a vulnerability privately

Use GitHub's private vulnerability reporting: the repository's **Security** tab, then **Report a vulnerability**
([direct link](https://github.com/uzuntasgokberk/ssrs-reportserver-export/security/advisories/new)). Do not open a
public issue, discussion or pull request for a security problem.

Include:

- the script version (the latest entry in `CHANGELOG.md`, or the commit you downloaded);
- the PowerShell version (`$PSVersionTable.PSVersion`), the report server product and version (the first lines of the
  log) and whether you signed in with the current Windows account or with `-Credential`;
- the smallest command line that reproduces it, with a fake server URL and fake folder and item names, and what you
  expected.

**Never send secrets or data**, in a report or anywhere else: no passwords, connection strings or exported files, and
no log files you have not reviewed (`_logs` holds the server URL, the account name and folder and item names).
Reproduce with fake values; if a credential was exposed, rotate it first.

This is a personal project maintained on a best-effort basis. Fixes go to the latest `main` only.

## What counts

- The script sends a request other than GET, or changes anything on the report server.
- A password ends up unmasked in an exported `.rds` file, or a password from a data source or from `-Credential`
  appears in the console output, the log file or the CSV.
- It writes, replaces or deletes a file outside the export folder, or deletes any local file.

## What is not a vulnerability

- Reports and datasets are exported byte for byte. A plain-text password written into an embedded connection string is
  in the exported file; the script warns about it (`WARNING: An embedded connection string contains a password`).
  Remove it on the server.
- A `.pbix` with imported data contains that data.
- Over `http`, the downloaded files cross the network unencrypted; use an `https` address on networks you don't trust.
- Problems in SQL Server Reporting Services or Power BI Report Server belong to Microsoft: report them through the
  [Microsoft Security Response Center](https://msrc.microsoft.com).
