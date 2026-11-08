#requires -Version 5.1

<#
.SYNOPSIS
    Exports the content of a SQL Server Reporting Services (SSRS) or Power BI Report Server,
    using the REST API v2.0: reports, shared data sources, shared datasets and report parts to
    .rdl / .rds / .rsd / .rsc files, and Power BI reports (Power BI Report Server only) to .pbix.

.DESCRIPTION
    The folder tree of the report server is kept as it is:

      Report        -> <ExportRoot>\<Folder>\<Report>.rdl
      DataSource    -> <ExportRoot>\<Folder>\<DataSource>.rds
      DataSet       -> <ExportRoot>\<Folder>\<DataSet>.rsd
      ReportPart    -> <ExportRoot>\<Folder>\<ReportPart>.rsc
      PowerBIReport -> <ExportRoot>\<Folder>\<Report>.pbix   (Power BI Report Server only)

    Read-only: the script sends GET requests only; nothing on the report server is changed.
    By default it signs in with the current Windows account (Kerberos / NTLM); use -Credential
    for another account.

    .rdl, .rsd, .rsc and .pbix files are saved byte for byte as stored on the server. .rds files are
    built from the data source settings. The report server never returns stored passwords, so
    passwords are never part of the export.

    Each item is downloaded into memory first; the file on disk is replaced only after the
    download succeeded, so a failed run never destroys the last good export. Local files whose
    item no longer exists on the server are reported, never deleted.

.PARAMETER ServerUrl
    Web portal address, e.g. "http://reportserver/Reports". A URL copied from the browser
    (".../Reports/browse/"), the web service address (".../ReportServer") or just the host
    ("http://reportserver") also work. Named instances keep their suffix ("/Reports_SQL2019").

.PARAMETER Folder
    Only export items in these server folders, subfolders included. Folder name or path;
    wildcards are allowed, e.g. "Finance", "Finance/Monthly", "Sales*".

.PARAMETER Name
    Only export items with these names. Wildcards are allowed, e.g. "Invoice*".

.PARAMETER Exclude
    Skip items or whole folders with these names. Wildcards are allowed.

.PARAMETER Type
    Only export these item types: Report, DataSource, DataSet, ReportPart, PowerBIReport.
    Default: all five. SSRS has no Power BI reports; there PowerBIReport simply finds nothing.

.PARAMETER Credential
    Account to sign in with. Default: the current Windows account.

.PARAMETER List
    Only list the matching items (type, folder, name, size). Nothing is downloaded.

.PARAMETER ExportRoot
    Output folder. Default: "export" under the current directory.

.EXAMPLE
    .\Export-SsrsContent.ps1 -ServerUrl "http://reportserver/Reports" -List
    Lists every item on the server without downloading anything.

.EXAMPLE
    .\Export-SsrsContent.ps1 -ServerUrl "http://reportserver/Reports" -ExportRoot "D:\ssrs-backup"
    Exports everything into a separate folder (for example your backup repository).

.EXAMPLE
    .\Export-SsrsContent.ps1 -ServerUrl "http://reportserver/Reports" -Folder "Finance" -Exclude "Archive"
    Exports the Finance folder and its subfolders, except any folder named "Archive".

.EXAMPLE
    .\Export-SsrsContent.ps1 -ServerUrl "http://reportserver/Reports" -Type DataSource, DataSet
    Exports only the shared data sources and shared datasets.

.EXAMPLE
    .\Export-SsrsContent.ps1 -ServerUrl "http://reportserver/Reports" -Type PowerBIReport
    Exports only the Power BI reports of a Power BI Report Server as .pbix files.

.EXAMPLE
    .\Export-SsrsContent.ps1 -ServerUrl "https://reports.contoso.com/Reports" -Credential (Get-Credential)
    Signs in with another account.

.LINK
    https://learn.microsoft.com/sql/reporting-services/developer/rest-api
#>

# PositionalBinding = $false: an unnamed or extra argument fails loudly instead of being
# silently bound to the next parameter (e.g. -Name "A","B" with powershell -File).
[CmdletBinding(PositionalBinding = $false)]
param(
    [Parameter(Mandatory)]
    [string]$ServerUrl,

    [string[]]$Folder,

    [string[]]$Name,

    [string[]]$Exclude,

    [ValidateSet("Report", "DataSource", "DataSet", "ReportPart", "PowerBIReport")]
    [string[]]$Type = @("Report", "DataSource", "DataSet", "ReportPart", "PowerBIReport"),

    [System.Management.Automation.PSCredential]$Credential,

    [switch]$List,

    [string]$ExportRoot = (Join-Path (Get-Location).Path "export")
)

# StrictMode 1.0 catches undefined variables but returns $null (instead of
# throwing) when a JSON property is missing from an API response.
Set-StrictMode -Version 1.0
$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------

$RequestTimeoutSeconds = 120   # Give up on a single HTTP request after this many seconds.
$MaxRetries            = 3     # Attempts per request on a network error, timeout or HTTP 503.
$RetryWaitSeconds      = 15    # Wait between attempts.

# Accept what people copy from the browser: ".../Reports/browse/", ".../Reports",
# ".../ReportServer" (web service address) or just the host. Named instances keep their suffix.
$PortalUrl = $ServerUrl.Trim().TrimEnd("/")
if     ($PortalUrl -match '(?i)^(https?://[^/]+)/ReportServer([^/]*)') { $PortalUrl = "$($Matches[1])/Reports$($Matches[2])" }
elseif ($PortalUrl -match '(?i)^(https?://[^/]+/Reports[^/]*)')       { $PortalUrl = $Matches[1] }
elseif ($PortalUrl -match '(?i)^https?://[^/]+$')                      { $PortalUrl = "$PortalUrl/Reports" }
else {
    Write-Host "-ServerUrl has an unexpected format: '$ServerUrl' (example: http://reportserver/Reports)" -ForegroundColor Red
    exit 2
}
$ApiBase = "$PortalUrl/api/v2.0"

$Separator  = [string][System.IO.Path]::DirectorySeparatorChar
$ExportRoot = [System.IO.Path]::GetFullPath($ExportRoot)
if ($ExportRoot -ne [System.IO.Path]::GetPathRoot($ExportRoot)) { $ExportRoot = $ExportRoot.TrimEnd([char[]]@('\', '/')) }
$LogFolder  = Join-Path $ExportRoot "_logs"
$RunStart   = Get-Date
$LogFile    = Join-Path $LogFolder ("export_{0:yyyyMMdd_HHmmss}.log" -f $RunStart)
$CsvFile    = [System.IO.Path]::ChangeExtension($LogFile, ".csv")

# -Type name -> catalog type and file extension (the extensions Visual Studio / Report Builder /
# Power BI Desktop use). The order is also the export order: data sources and datasets first,
# reports last.
$TypeInfo = [ordered]@{
    DataSource    = @{ ApiType = "DataSource";    Extension = ".rds" }
    DataSet       = @{ ApiType = "DataSet";       Extension = ".rsd" }
    ReportPart    = @{ ApiType = "Component";     Extension = ".rsc" }
    Report        = @{ ApiType = "Report";        Extension = ".rdl" }
    PowerBIReport = @{ ApiType = "PowerBIReport"; Extension = ".pbix" }   # Power BI Report Server only.
}

# GitHub rejects files larger than 100 MB; a bigger .pbix needs Git LFS if the export is kept in Git.
$LargeFileBytes = 100MB
$ApiTypeToName = @{}
$TypeOrder     = @{}
foreach ($Key in $TypeInfo.Keys) {
    $ApiTypeToName[$TypeInfo[$Key].ApiType] = $Key
    $TypeOrder[$Key] = $TypeOrder.Count
}

# "Password=..." / "Pwd=..." inside a connection string, with its value.
$PasswordPattern = '(?i)\b(password|pwd)\s*=\s*("[^"]*"|''[^'']*''|[^;]*)'
# The same, inside an embedded <ConnectString> of a report or dataset definition.
$EmbeddedPasswordPattern = '(?is)<ConnectString>[^<]*\b(password|pwd)\s*='

$Utf8NoBom = New-Object System.Text.UTF8Encoding $false
$Results   = New-Object System.Collections.Generic.List[object]

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

function Write-ExportLog([string]$Message, [string]$Level = "INFO") {
    $Line   = "[{0:yyyy-MM-dd HH:mm:ss}][{1}] {2}" -f (Get-Date), $Level, $Message
    $Colors = @{ INFO = "Cyan"; SUCCESS = "Green"; WARNING = "Yellow"; ERROR = "Red" }
    Write-Host $Line -ForegroundColor $Colors[$Level]

    # Logging must never stop the export (e.g. antivirus briefly locking the log file).
    try   { [System.IO.File]::AppendAllText($LogFile, $Line + [Environment]::NewLine, $Utf8NoBom) }
    catch { Write-Host "  (could not write to the log file: $($_.Exception.Message))" -ForegroundColor DarkGray }
}

function Add-Result($Entry, [string]$Status, [int]$Bytes, [string]$Message, [double]$Seconds) {
    $Results.Add([PSCustomObject]@{
        DateTime   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        ServerPath = $Entry.Path
        ItemType   = $Entry.Type
        ItemId     = $Entry.Id
        Status     = $Status
        Bytes      = $Bytes
        ServerSize = $Entry.Size
        Seconds    = [Math]::Round($Seconds, 1)
        OutputPath = $Entry.Target
        Message    = $Message
    })
}

# ---------------------------------------------------------------------------
# File helpers (.NET APIs, because cmdlets treat [ ] in names as wildcards)
# ---------------------------------------------------------------------------

# Windows rules on every platform, so an export can be copied or cloned anywhere.
$InvalidNameChars = [char[]]'\/:*?"<>|' + [char[]](0..31)

function Get-SafeFileName([string]$Name) {
    foreach ($Char in $InvalidNameChars) { $Name = $Name.Replace([string]$Char, "_") }

    $Name = $Name.Trim().TrimEnd(".")
    if (-not $Name) { return "_Unnamed" }
    if ($Name -match '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\..*)?$') { return "_$Name" }
    return $Name
}

function Save-ItemFile([string]$Path, [byte[]]$Bytes) {
    # Called only after the download succeeded. Retries because antivirus/indexer can briefly lock files.
    for ($Attempt = 1; ; $Attempt++) {
        try {
            [void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($Path))
            [System.IO.File]::WriteAllBytes($Path, $Bytes)
            return
        }
        catch {
            if ($Attempt -ge 5) {
                throw "Could not write to disk: $($_.Exception.Message) Hint: the file may be open in another program (Report Builder / Visual Studio)."
            }
            Start-Sleep -Seconds 1
        }
    }
}

# ---------------------------------------------------------------------------
# Report server REST API (GET only)
# ---------------------------------------------------------------------------

function Get-HttpHint([int]$StatusCode) {
    switch ($StatusCode) {
        401 { return "Authentication failed: the report server does not accept the account (check the network / VPN, or pass -Credential)." }
        403 { return "Access denied: downloading a definition needs the 'Content Manager' or 'Publisher' role on this item ('Browser' is not enough)." }
        404 { return "Not found: the item may have been deleted or moved, or -ServerUrl is wrong. The REST API v2.0 needs SSRS 2017 or later, or Power BI Report Server." }
    }
    if ($StatusCode -ge 500) { return "Report server error: try again later; if it persists, check the report server logs." }
    return ""
}

function Get-ServerError($Response) {
    # The REST API returns errors as {"error":{"code":"...","message":"..."}}; show that message.
    try {
        $Reader = New-Object System.IO.StreamReader($Response.GetResponseStream(), $Utf8NoBom)
        $Text   = $Reader.ReadToEnd()
        $Reader.Close()
        return [string]($Text | ConvertFrom-Json).error.message
    }
    catch { return "" }
    finally { $Response.Close() }
}

function Invoke-SsrsRequest([string]$Url) {
    # Returns the response body as bytes. GET is the only request this script ever sends.
    # HttpWebRequest is used instead of Invoke-RestMethod: PowerShell 7 refuses
    # -UseDefaultCredentials over http, this works the same in 5.1 and 7.
    if ($Url -notmatch '^https?://') { $Url = "$ApiBase/$Url" }

    for ($Attempt = 1; ; $Attempt++) {
        $Request = [System.Net.WebRequest]::Create($Url)
        $Request.Method           = "GET"
        $Request.Timeout          = $RequestTimeoutSeconds * 1000
        $Request.ReadWriteTimeout = $RequestTimeoutSeconds * 1000
        if ($Credential) { $Request.Credentials = $Credential.GetNetworkCredential() }
        else             { $Request.UseDefaultCredentials = $true }   # Current Windows account (Kerberos / NTLM).

        $Response = $null
        try {
            $Response = $Request.GetResponse()
            $Memory   = New-Object System.IO.MemoryStream
            $Response.GetResponseStream().CopyTo($Memory)
            return ,$Memory.ToArray()   # The comma keeps the byte array in one piece.
        }
        catch {
            $WebError = $_.Exception
            while ($WebError -and $WebError -isnot [System.Net.WebException]) { $WebError = $WebError.InnerException }
            if (-not $WebError) { throw }

            $Status = 0
            if ($WebError.Response) { $Status = [int]$WebError.Response.StatusCode }

            # No response at all (network down, DNS, timeout) or "server busy": try again.
            if (($Status -eq 0 -or $Status -eq 503) -and $Attempt -lt $MaxRetries) {
                Write-ExportLog "Request failed ($($WebError.Message)) | retrying in $RetryWaitSeconds s ($Attempt/$MaxRetries): $Url" "WARNING"
                Start-Sleep -Seconds $RetryWaitSeconds
                continue
            }

            if ($Status -eq 0) {
                throw "Network error: the report server could not be reached (check the network / VPN connection). Detail: $($WebError.Message)"
            }
            throw ((@("HTTP $Status", (Get-ServerError $WebError.Response), (Get-HttpHint $Status)) | Where-Object { $_ }) -join " | ")
        }
        finally {
            if ($Response) { $Response.Close() }
        }
    }
}

function Invoke-SsrsJson([string]$Url) {
    $Text = $Utf8NoBom.GetString((Invoke-SsrsRequest $Url)).TrimStart([char]0xFEFF)
    try   { return $Text | ConvertFrom-Json }
    catch { throw "The response is not JSON ($Url): $($Text.Substring(0, [Math]::Min(300, $Text.Length)))" }
}

function Get-CatalogItem {
    # Every item on the server (all folders, recursive). Follows paging if the server uses it.
    $Items = New-Object System.Collections.Generic.List[object]
    $Url   = "CatalogItems?`$select=Id,Name,Path,Type,Size"
    while ($Url) {
        $Page = Invoke-SsrsJson $Url
        foreach ($Item in @($Page.value)) { $Items.Add($Item) }
        $Url = $Page.'@odata.nextLink'
    }
    return $Items
}

# ---------------------------------------------------------------------------
# Filters
# ---------------------------------------------------------------------------

function Test-NameMatch([string]$Text, [string[]]$Patterns) {
    # Exact match (ignoring surrounding spaces) or wildcard match.
    foreach ($Pattern in $Patterns) {
        if ($Text.Trim() -eq $Pattern.Trim() -or $Text.Trim() -like $Pattern.Trim()) { return $true }
    }
    return $false
}

function Test-FolderMatch([string]$ServerFolder, [string[]]$Patterns) {
    # "/Finance/Monthly" matches "Finance", "Finance/Monthly", "Fin*" and "/Finance/" (subfolders included).
    $Parent = $ServerFolder.Trim("/")
    foreach ($Pattern in $Patterns) {
        $Pattern = $Pattern.Trim().Trim("/")
        if (-not $Pattern) { continue }
        if ($Parent -eq $Pattern -or $Parent -like $Pattern -or $Parent -like "$Pattern/*") { return $true }
        # A folder name without "/" matches at any depth, e.g. -Exclude "Archive" also skips "/Finance/Archive".
        if ($Pattern -notmatch "/" -and (@($Parent.Split("/") | Where-Object { Test-NameMatch $_ @($Pattern) }).Count -gt 0)) { return $true }
    }
    return $false
}

function Test-Selected($Entry, [string[]]$Types, [string[]]$Folders, [string[]]$Names, [string[]]$Excludes) {
    # Applies -Type, -Folder, -Exclude (item or folder name) and -Name.
    if ($Types -notcontains $Entry.Type) { return $false }
    if ($Folders -and -not (Test-FolderMatch $Entry.ServerFolder $Folders)) { return $false }
    if ($Excludes -and ((Test-NameMatch $Entry.Name $Excludes) -or (Test-FolderMatch $Entry.ServerFolder $Excludes))) { return $false }
    if (-not $Names) { return $true }
    return (Test-NameMatch $Entry.Name $Names)
}

# ---------------------------------------------------------------------------
# Item builders: return @{ Bytes; Info; Warn }
# ---------------------------------------------------------------------------

function Get-ContentInfo($Entry, [byte[]]$Bytes) {
    # Short summary for the log: which shared data sources / datasets the definition uses.
    $Info = @("$($TypeInfo[$Entry.Type].Extension) downloaded ({0:N0} bytes)" -f $Bytes.Length)
    $Warn = @()

    # A .pbix is a binary (ZIP) file: there are no references to read, only its size to check.
    if ($Entry.Type -eq "PowerBIReport" -and $Bytes.Length -gt $LargeFileBytes) {
        $Warn += "The file is larger than 100 MB; GitHub rejects such files, use Git LFS if you keep the export in Git."
    }
    if ($Entry.Type -eq "ReportPart" -or $Entry.Type -eq "PowerBIReport") { return @{ Info = $Info; Warn = $Warn } }

    $Text     = $Utf8NoBom.GetString($Bytes)
    $Sources  = @([regex]::Matches($Text, '<DataSourceReference>\s*([^<]*?)\s*</DataSourceReference>') | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
    $DataSets = @([regex]::Matches($Text, '<SharedDataSetReference>\s*([^<]*?)\s*</SharedDataSetReference>') | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
    $Embedded = [regex]::Matches($Text, '<ConnectionProperties>').Count

    if ($Sources)  { $Info += "Data source: $($Sources -join ', ')" }
    if ($Embedded) { $Info += "Embedded data sources: $Embedded" }
    if ($DataSets) { $Info += "Shared dataset: $($DataSets -join ', ')" }

    # Only reported; the file is kept byte for byte as on the server.
    if ($Text -match $EmbeddedPasswordPattern) {
        $Warn += "An embedded connection string contains a password (file not changed; review it before committing)."
    }
    return @{ Info = $Info; Warn = $Warn }
}

function Build-Content($Entry) {
    # Reports, shared datasets, report parts and Power BI reports: the stored file, unchanged.
    $Bytes = Invoke-SsrsRequest "CatalogItems($($Entry.Id))/Content/`$value"
    if ($Bytes.Length -eq 0) { throw "The server returned empty content (0 bytes)." }

    $Content = Get-ContentInfo $Entry $Bytes
    return @{ Bytes = $Bytes; Info = $Content.Info; Warn = $Content.Warn }
}

function Build-DataSourceFile($Entry) {
    # Shared data source: the report server has no file for it, so the standard .rds XML (the
    # format Visual Studio report projects use) is built from its settings.
    $Source        = Invoke-SsrsJson "DataSources($($Entry.Id))"
    $Retrieval     = [string]$Source.CredentialRetrieval      # prompt / store / integrated / none
    $ConnectString = [string]$Source.ConnectionString
    $Warn          = @()

    if ($ConnectString -match $PasswordPattern) {
        $ConnectString = [regex]::Replace($ConnectString, $PasswordPattern, '$1=********')
        $Warn += "The password in the connection string was masked with ********."
    }
    if ($Source.IsEnabled -eq $false) { $Warn += "The data source is disabled on the server." }

    $Settings          = New-Object System.Xml.XmlWriterSettings
    $Settings.Indent   = $true
    $Settings.Encoding = $Utf8NoBom
    $Memory            = New-Object System.IO.MemoryStream
    $Writer            = [System.Xml.XmlWriter]::Create($Memory, $Settings)

    $Writer.WriteStartDocument()
    $Writer.WriteComment((" Exported from the report server: {0} | CredentialRetrieval: {1} | Passwords are never exported; enter them again after restoring. " -f $Entry.Path, $Retrieval).Replace("--", "- -"))
    $Writer.WriteStartElement("RptDataSource")
    $Writer.WriteAttributeString("xmlns", "xsi", $null, "http://www.w3.org/2001/XMLSchema-instance")
    $Writer.WriteAttributeString("xmlns", "xsd", $null, "http://www.w3.org/2001/XMLSchema")
    $Writer.WriteAttributeString("Name", $Entry.Name)
    $Writer.WriteStartElement("ConnectionProperties")
    $Writer.WriteElementString("Extension", [string]$Source.DataSourceType)
    $Writer.WriteElementString("ConnectString", $ConnectString)
    if ($Retrieval -eq "integrated") { $Writer.WriteElementString("IntegratedSecurity", "true") }
    if ($Retrieval -eq "prompt") {
        $PromptText = [string]$Source.CredentialsByUser.DisplayText
        if (-not $PromptText) { $PromptText = "Enter a user name and password to access the data source:" }
        $Writer.WriteElementString("Prompt", $PromptText)
    }
    $Writer.WriteEndElement()
    $Writer.WriteElementString("DataSourceID", $Entry.Id)   # Server ID, so the file does not change between runs.
    $Writer.WriteEndElement()
    $Writer.WriteEndDocument()
    $Writer.Close()
    $Bytes = $Memory.ToArray()

    # The user name of stored credentials goes to the log/CSV only, never into the .rds file.
    $CredentialText = "Credentials: $Retrieval"
    if ($Retrieval -eq "store" -and $Source.CredentialsInServer.UserName) { $CredentialText += " (user: $($Source.CredentialsInServer.UserName))" }

    $Info = @((".rds created ({0:N0} bytes)" -f $Bytes.Length), "Type: $($Source.DataSourceType)", $CredentialText)
    if ($Retrieval -eq "store") { $Info += "Password is not in the export; enter it again when restoring" }
    return @{ Bytes = $Bytes; Info = $Info; Warn = $Warn }
}

function Export-Entry($Entry, [string]$Progress) {
    $Timer = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        if ($Entry.Type -eq "DataSource") { $Output = Build-DataSourceFile $Entry }
        else                              { $Output = Build-Content $Entry }

        Save-ItemFile $Entry.Target $Output.Bytes

        $Message = $Output.Info -join " | "
        $Level   = "SUCCESS"
        if ($Output.Warn) {
            $Message = "WARNING: $($Output.Warn -join ' ') | $Message"
            $Level   = "WARNING"
        }
        Write-ExportLog ("{0} {1} | {2} | {3:N1} s | {4}" -f $Progress, $Entry.Type, $Entry.Path, $Timer.Elapsed.TotalSeconds, $Message) $Level
        Add-Result $Entry "Success" $Output.Bytes.Length $Message $Timer.Elapsed.TotalSeconds
    }
    catch {
        $ErrorText = $_.Exception.Message
        Write-ExportLog ("{0} {1} | {2} | FAILED | id={3} | {4}" -f $Progress, $Entry.Type, $Entry.Path, $Entry.Id, $ErrorText) "ERROR"
        Add-Result $Entry "Failed" 0 $ErrorText $Timer.Elapsed.TotalSeconds
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

$Account = if ($Credential) { "$($Credential.UserName) (-Credential)" }
           elseif ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT) { [System.Security.Principal.WindowsIdentity]::GetCurrent().Name }
           else { "$([Environment]::UserName) (default credentials)" }

[void][System.IO.Directory]::CreateDirectory($LogFolder)
Write-ExportLog "Export started (read-only: only GET requests are sent to the report server)."
Write-ExportLog "Server        : $PortalUrl"
Write-ExportLog "Target folder : $ExportRoot"
Write-ExportLog "PowerShell    : $($PSVersionTable.PSVersion)"
Write-ExportLog "Account       : $Account"

$FilterText = "Type = $($Type -join ', ')"
if ($Folder)  { $FilterText += " | Folder = $($Folder -join ', ')" }
if ($Name)    { $FilterText += " | Name = $($Name -join ', ')" }
if ($Exclude) { $FilterText += " | Exclude = $($Exclude -join ', ')" }
if ($List)    { $FilterText += " | List only" }
Write-ExportLog "Filter        : $FilterText"

try {
    $System = Invoke-SsrsJson "System"
    Write-ExportLog "Report server : $($System.ProductName) $($System.ProductVersion) | $($System.ReportServerAbsoluteUrl)"
    $Catalog = @(Get-CatalogItem)
}
catch {
    Write-ExportLog "Could not connect to the report server: $($_.Exception.Message)" "ERROR"
    Write-ExportLog "Check: does $PortalUrl open in the browser? Is the network / VPN connected? Is -ServerUrl correct?" "ERROR"
    exit 2
}

# One entry per exportable item, with its local file path.
$Entries = New-Object System.Collections.Generic.List[object]
foreach ($Item in $Catalog) {
    $TypeName = $ApiTypeToName[[string]$Item.Type]
    if (-not $TypeName) { continue }

    $Segments = @(([string]$Item.Path).Split("/") | Where-Object { $_ })
    $Parents  = @()
    if ($Segments.Count -gt 1) { $Parents = @($Segments[0..($Segments.Count - 2)]) }

    $Entries.Add([PSCustomObject]@{
        Id           = [string]$Item.Id
        Name         = [string]$Item.Name
        Path         = [string]$Item.Path
        Type         = $TypeName
        Size         = $Item.Size
        ServerFolder = "/" + ($Parents -join "/")
        LocalFolder  = (@($Parents | ForEach-Object { Get-SafeFileName $_ }) -join $Separator)
        SafeName     = Get-SafeFileName ([string]$Item.Name)
        Target       = $null
    })
}

# Names are unique per folder on the server, but two names can become the same file name once
# invalid characters are replaced (or differ only in case). Those files get a short ID suffix
# instead of overwriting each other. All server items are compared (not only the selected ones),
# so file names do not depend on the filter.
$NameCount = @{}
foreach ($Entry in $Entries) {
    $Key = ("{0}|{1}|{2}" -f $Entry.Type, $Entry.LocalFolder, $Entry.SafeName).ToLowerInvariant()
    $NameCount[$Key] = 1 + [int]$NameCount[$Key]
}
foreach ($Entry in $Entries) {
    $Key      = ("{0}|{1}|{2}" -f $Entry.Type, $Entry.LocalFolder, $Entry.SafeName).ToLowerInvariant()
    $FileName = $Entry.SafeName
    if ($NameCount[$Key] -gt 1) { $FileName = "$FileName ($($Entry.Id.Substring(0, 8)))" }
    $Entry.Target = [System.IO.Path]::Combine($ExportRoot, $Entry.LocalFolder, $FileName + $TypeInfo[$Entry.Type].Extension)
}

$Counts = foreach ($TypeName in $TypeInfo.Keys) { "{0} {1}" -f @($Entries | Where-Object Type -eq $TypeName).Count, $TypeName }
Write-ExportLog ("On the server : {0} | {1} folders | {2} items in total" -f ($Counts -join ", "), @($Catalog | Where-Object Type -eq "Folder").Count, $Catalog.Count)

$Unsupported = @($Catalog | Where-Object { $_.Type -ne "Folder" -and -not $ApiTypeToName[[string]$_.Type] })
foreach ($Item in $Unsupported) {
    Write-ExportLog "Unsupported type, skipped: $($Item.Type) | $($Item.Path)" "WARNING"
}

$Selected = @($Entries |
    Where-Object { Test-Selected $_ $Type $Folder $Name $Exclude } |
    Sort-Object @{ Expression = { $TypeOrder[$_.Type] } }, Path)

$Counts = foreach ($TypeName in $TypeInfo.Keys) { "{0} {1}" -f @($Selected | Where-Object Type -eq $TypeName).Count, $TypeName }
Write-ExportLog ("Selected      : {0} | {1} in total" -f ($Counts -join ", "), $Selected.Count)

if ($Selected.Count -eq 0) {
    Write-ExportLog "No item matches the filter. To see the names, run the script with -List." "ERROR"
    exit 1
}

if ($List) {
    Write-Host ($Selected | Format-Table @{ n = "Type"; e = { $_.Type } }, @{ n = "Folder"; e = { $_.ServerFolder } },
        @{ n = "Name"; e = { $_.Name } }, @{ n = "Size"; e = { $_.Size } } -AutoSize | Out-String -Width 250)
    exit 0
}

$Index = 0
foreach ($Entry in $Selected) {
    $Index++
    Export-Entry $Entry ("[{0}/{1}]" -f $Index, $Selected.Count)
}

# After a full run, list local files whose item is no longer on the server (deleted, moved or
# renamed). They are never deleted automatically. Skipped for filtered runs, where most items
# are intentionally not part of the run.
$Stale = @()
if (-not $Folder -and -not $Name -and -not $Exclude -and $Type.Count -eq $TypeInfo.Count) {
    $Expected = New-Object "System.Collections.Generic.HashSet[string]" ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($Entry in $Entries) { [void]$Expected.Add($Entry.Target) }
    $Extensions = @($TypeInfo.Values | ForEach-Object { $_.Extension })

    $Stale = @([System.IO.Directory]::EnumerateFiles($ExportRoot, "*", [System.IO.SearchOption]::AllDirectories) |
        Where-Object { $Extensions -contains [System.IO.Path]::GetExtension($_) -and -not $Expected.Contains($_) })

    foreach ($File in $Stale) {
        Write-ExportLog "Not on the server any more (deleted, moved or renamed). The file was NOT deleted; delete it yourself if needed: $File" "WARNING"
    }
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

if ($Results.Count -gt 0) {
    # UTF-8 with BOM so Excel shows non-ASCII characters correctly ("UTF8" has no BOM in PowerShell 7).
    $CsvEncoding = if ($PSVersionTable.PSVersion.Major -ge 6) { "utf8BOM" } else { "UTF8" }
    $Results | Export-Csv -LiteralPath $CsvFile -NoTypeInformation -Encoding $CsvEncoding
}

$Failed   = @($Results | Where-Object Status -eq "Failed")
$Warnings = @($Results | Where-Object { $_.Status -eq "Success" -and $_.Message -like "WARNING:*" })

Write-ExportLog "============================================================"
Write-ExportLog ("Done. Succeeded: {0} (with warnings: {1}), Failed: {2}, Duration: {3:hh\:mm\:ss}" -f
    ($Results.Count - $Failed.Count), $Warnings.Count, $Failed.Count, ((Get-Date) - $RunStart))

if ($Stale.Count -gt 0) {
    Write-ExportLog "Local files not on the server: $($Stale.Count) (not deleted; see the warnings above)." "WARNING"
}
foreach ($Row in $Failed) {
    Write-ExportLog "  FAILED | $($Row.ServerPath) [$($Row.ItemType)] | $($Row.Message)" "ERROR"
}

Write-ExportLog "Log: $LogFile"
if ($Results.Count -gt 0) { Write-ExportLog "CSV: $CsvFile" }

# Non-zero exit code so scheduled tasks can detect failures.
if ($Failed.Count -gt 0) { exit 1 }
