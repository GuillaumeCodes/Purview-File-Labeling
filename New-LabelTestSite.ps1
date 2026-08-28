#Requires -Version 5.1

<#
.SYNOPSIS
    Creates a SharePoint Online test site with a document library, a nested folder
    hierarchy, and real Office files that Microsoft Purview can label. The amount of
    test content defaults to a small set and is adjustable.
#>

[CmdletBinding()]
param(
    [string]$TenantRootUrl = '',
    [string]$TenantHint = '',
    [string]$SiteName = '',
    [string]$LibraryName = '',
    # The defaults make a small tree that is quick to provision and quick to label; raise them for a busier library.
    [ValidateRange(0, 10)][int]$TopLevelFolders = 2,
    [ValidateRange(1, 5)][int]$FolderDepth = 3,
    [ValidateRange(0, 3)][int]$SubfoldersPerFolder = 1,
    # Drawn per folder from an inclusive range, so the library is not uniform. A single number pins the count instead.
    [ValidatePattern('^\s*\d+\s*(-\s*\d+\s*)?$')][string]$FilesPerFolder = '1-4',
    [ValidatePattern('^\s*\d+\s*(-\s*\d+\s*)?$')][string]$SensitiveFilesPerFolder = '0-2',
    [string]$LogFolder = '',
    [string]$ClientId = '',
    [string]$ApplicationName = 'PnP PowerShell - Label Test Site',
    [ValidateSet('Browser', 'DeviceCode')]
    [string]$AuthMode = 'Browser',
    [switch]$RegisterApp,
    [switch]$KeepApp,
    [switch]$AcceptDefaults,
    [switch]$PreflightOnly,
    [switch]$NoRelaunch,
    # Internal: set only when this script re-invokes itself to clean up an application.
    [string]$RemoveApplicationClientId = '',
    [string]$RemoveApplicationTenantId = '',
    [ValidateSet('', 'Production', 'USGovernmentHigh', 'Germany', 'China')]
    [string]$RemoveApplicationEnvironment = '',
    [string]$RemoveApplicationResultPath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:LibraryTitle = 'Label Test Library'
$script:LibraryUrl = 'LabelTestLibrary'
# Shared by log names and the generated site and application names.
$script:TimestampFormat = 'yyyyMMdd-HHmmss'
$script:LogPath = ''
$script:LogBuffer = [System.Collections.Generic.List[string]]::new()
# Only an application this run registered may ever be cleaned up.
$script:CreatedApplication = $null
# Set only when the operator asks for the labeling utility to run once provisioning finishes.
$script:LaunchLabelingPath = ''

# Built from the requested counts, parents first, so each parent exists before its child.
$script:FolderTree = @()
# One name pool per depth level, so a generated tree still reads like a real library.
$script:FolderNamesByDepth = @(
    @('Finance', 'HR', 'Legal', 'Operations', 'Marketing', 'Sales', 'Facilities', 'Research', 'Support', 'Projects'),
    @('Reports', 'Policies', 'Contracts', 'Budgets', 'Plans', 'Records'),
    @('Q1', 'Q2', 'Q3', 'Q4', 'Annual', 'Monthly'),
    @('Archive', 'Drafts', 'Final', 'Working', 'Review', 'Shared'),
    @('2023', '2024', '2025', 'Signed', 'Superseded', 'Reference')
)
# Fabricated records shaped like the public sample sets at dlptest.com, so Purview classifiers have something to match.
# Each value pairs a valid pattern with a keyword from that SIT's documented supporting list, because most SITs need both within 300 characters.
# The card numbers pass the Luhn test but are not the reserved numbers that the credit card SIT deliberately ignores.
$script:SensitiveSampleRecords = @(
    @(
        'Customer record - payroll onboarding'
        'Full name: Jane A. Doe'
        'Home address: 4120 Maple Avenue, Redmond, WA 98052'
        'Date of birth: 11 March 1974'
        'Social Security Number (SSN): 539-95-4188'
        'Credit card number (Visa): 4539 1488 0343 6467'
        'Card expiration date: 04/2028'
        'Card verification value (CVV2): 731'
        'Email: jane.doe@contoso.com'
        'Phone: +1 (425) 555-0142'
    ),
    @(
        'Customer record - account verification'
        'Full name: Miguel R. Santos'
        'Home address: 88 Belmont Street, Boston, MA 02116'
        'Date of birth: 02 September 1986'
        'Social Security Number (SSN): 412-33-7291'
        'Bank account number: 493042798132'
        'ABA routing number: 121000358'
        'IBAN: GB82WEST12345698765432'
        'Email: miguel.santos@contoso.com'
    ),
    @(
        'Customer record - card dispute'
        'Full name: Priya N. Raman'
        'Billing address: 1701 Harbor Drive, Seattle, WA 98101'
        'Credit card number (Mastercard): 5412 7412 3456 7899'
        'Card expiration date: 11/2027'
        'Card verification value (CVV2): 204'
        'Credit card number (American Express): 3765 010234 56786'
        'Social Security Number (SSN): 205-64-8813'
        'Phone: +1 (206) 555-0173'
    ),
    @(
        'Customer record - identity check'
        'Full name: Thomas L. Berger'
        'Home address: 2200 Alameda Street, San Jose, CA 95112'
        'Date of birth: 24 December 1969'
        'US passport number: 340025519'
        'Driver''s license number (California DL): I1234562'
        'Individual Taxpayer Identification Number (ITIN): 912-73-4567'
        'Credit card number (Discover): 6011 2345 6789 0123'
        'Card expiration date: 08/2029'
        'Email: thomas.berger@contoso.com'
    )
)
# Test data is disposable, so a shape that would take hours to provision is refused instead of started.
$script:MaximumTestFolders = 250
$script:MaximumTestFiles = 1000

function Add-LogEntry {
    <# .SYNOPSIS Records one action in the run log, buffering until the log file exists. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Message,
        [ValidateSet('Info', 'Good', 'Warn', 'Error')][string]$Severity = 'Info'
    )

    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Severity.ToUpperInvariant(), ($Message -replace '[\r\n]+', ' ')
    $script:LogBuffer.Add($line)
    if ([string]::IsNullOrWhiteSpace($script:LogPath)) { return }
    try { Add-Content -LiteralPath $script:LogPath -Value $line -Encoding utf8 -ErrorAction Stop }
    catch {
        $script:LogPath = ''
        Write-Host "  Warning: the run log could not be written, so logging continues in memory only. $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

function Write-Banner {
    <# .SYNOPSIS Prints a boxed notice whose rules always span the widest line inside it. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Title,
        [string[]]$Body = @()
    )

    $lines = @($Title) + @($Body)
    $longest = ($lines | Measure-Object -Property Length -Maximum).Maximum
    # Capped to the window, because a rule that wraps looks worse than one that is short.
    $limit = 100
    try { $limit = [math]::Max(40, $Host.UI.RawUI.WindowSize.Width - 4) }
    catch { Write-Verbose 'The console width is unavailable, so a default banner width is used.' }
    $rule = '  ' + ('=' * [math]::Min($longest + 1, $limit))

    Write-Host $rule -ForegroundColor Cyan
    Write-Host "   $Title" -ForegroundColor Cyan
    foreach ($line in $Body) {
        if ([string]::IsNullOrEmpty($line)) { Write-Host '' } else { Write-Host "   $line" -ForegroundColor Gray }
    }
    Write-Host $rule -ForegroundColor Cyan
}

function Write-Step {
    <# .SYNOPSIS Writes one readable progress line to the console and the run log. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Info', 'Good', 'Warn', 'Error')][string]$Severity = 'Info'
    )

    $color = switch ($Severity) {
        'Good' { 'Green' }
        'Warn' { 'Yellow' }
        'Error' { 'Red' }
        default { 'Gray' }
    }
    $tag = switch ($Severity) {
        'Warn' { 'Warning: ' }
        'Error' { 'ERROR: ' }
        default { '' }
    }
    Write-Host "  $tag$Message" -ForegroundColor $color
    Add-LogEntry -Message $Message -Severity $Severity
}

function Start-RunLog {
    <# .SYNOPSIS Creates the timestamped action log and flushes everything recorded so far. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Folder)

    try {
        if (-not (Test-Path -LiteralPath $Folder -PathType Container)) {
            $null = New-Item -ItemType Directory -Path $Folder -Force -ErrorAction Stop
        }
        $candidatePath = Join-Path $Folder ("New-LabelTestSite-{0}.log" -f (Get-Date -Format $script:TimestampFormat))
        Set-Content -LiteralPath $candidatePath -Value $script:LogBuffer -Encoding utf8 -ErrorAction Stop
        $script:LogPath = $candidatePath
        Write-Step -Severity Good -Message "Logging every action to $candidatePath"
        return $candidatePath
    }
    catch {
        Write-Step -Severity Warn -Message "Could not create a log file in '$Folder', so this run is logged to the console only. $($_.Exception.Message)"
        return ''
    }
}

function Read-Setting {
    <# .SYNOPSIS Proposes a value and lets the operator accept it with Enter or type a replacement. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Default,
        [switch]$UseDefault
    )

    if ($UseDefault -or -not (Test-InteractiveHost)) {
        Write-Step -Message "${Prompt}: $Default (default)"
        return $Default
    }

    $value = Read-Host "  $Prompt [$Default]"
    if ([string]::IsNullOrWhiteSpace($value)) { $value = $Default }
    $value = $value.Trim()
    Add-LogEntry -Message "${Prompt}: $value"
    return $value
}

function Read-CountSetting {
    <# .SYNOPSIS Proposes a count and accepts only a whole number inside the allowed range. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [Parameter(Mandatory)][int]$Default,
        [Parameter(Mandatory)][int]$Minimum,
        [Parameter(Mandatory)][int]$Maximum,
        [switch]$UseDefault
    )

    while ($true) {
        $value = Read-Setting -Prompt "$Prompt ($Minimum-$Maximum)" -Default ([string]$Default) -UseDefault:$UseDefault
        $parsed = 0
        if ([int]::TryParse($value, [ref]$parsed) -and $parsed -ge $Minimum -and $parsed -le $Maximum) { return $parsed }
        Write-Step -Severity Warn -Message "Enter a whole number between $Minimum and $Maximum."
        # A non-interactive host cannot correct itself, and the proposed value is always inside the range.
        if ($UseDefault -or -not (Test-InteractiveHost)) { return $Default }
    }
}

function ConvertTo-CountRange {
    <# .SYNOPSIS Reads "3" or "1-4" as an inclusive range, or returns nothing when the text cannot be one. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Text,
        [Parameter(Mandatory)][int]$Minimum,
        [Parameter(Mandatory)][int]$Maximum
    )

    $parsed = [regex]::Match($Text.Trim(), '^(\d+)\s*(?:-\s*(\d+))?$')
    if (-not $parsed.Success) { return $null }
    $low = [int]$parsed.Groups[1].Value
    $high = if ($parsed.Groups[2].Success) { [int]$parsed.Groups[2].Value } else { $low }
    if ($low -gt $high) { $low, $high = $high, $low }
    if ($low -lt $Minimum -or $high -gt $Maximum) { return $null }
    return [pscustomobject]@{ Minimum = $low; Maximum = $high }
}

function Read-RangeSetting {
    <# .SYNOPSIS Proposes a count or a range like 1-4 and accepts only a value inside the allowed bounds. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [Parameter(Mandatory)][string]$Default,
        [Parameter(Mandatory)][int]$Minimum,
        [Parameter(Mandatory)][int]$Maximum,
        [switch]$UseDefault
    )

    while ($true) {
        $value = Read-Setting -Prompt "$Prompt, one number or a range within $Minimum-$Maximum" -Default $Default -UseDefault:$UseDefault
        $range = ConvertTo-CountRange -Text $value -Minimum $Minimum -Maximum $Maximum
        if ($null -ne $range) { return $range }
        # A supplied value is wrong rather than mistyped, so it is reported instead of asked for again.
        if ($UseDefault -or -not (Test-InteractiveHost)) { throw "'$value' is not a whole number, or a range like 1-4, between $Minimum and $Maximum." }
        Write-Step -Severity Warn -Message "Enter a whole number, or a range like 1-4, between $Minimum and $Maximum."
        # Proposing the full range next keeps Enter from repeating a value that was just refused.
        $Default = "$Minimum-$Maximum"
    }
}

function Test-InteractiveHost {
    <# .SYNOPSIS Reports whether this host can prompt, so unattended runs fail fast instead of hanging. #>
    [CmdletBinding()]
    param()

    if (-not [Environment]::UserInteractive) { return $false }
    # PowerShell accepts any unambiguous prefix, so -noni and -noninteractive must both count as suppressing prompts.
    return -not @([Environment]::GetCommandLineArgs() -match '(?i)^-noni').Count
}

function Read-RequiredValue {
    <# .SYNOPSIS Prompts until the operator supplies a non-empty value. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [string]$Default = ''
    )

    while ($true) {
        $suffix = if ($Default) { " [$Default]" } else { '' }
        $value = Read-Host "  $Prompt$suffix"
        if ([string]::IsNullOrWhiteSpace($value) -and $Default) { return $Default }
        if (-not [string]::IsNullOrWhiteSpace($value)) { return $value.Trim() }
        Write-Step -Severity Warn -Message 'A value is required.'
    }
}

function Get-WebExceptionResponse {
    <# .SYNOPSIS Returns the HTTP response carried by a failed web request, or null. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord)

    $responseProperty = $ErrorRecord.Exception.PSObject.Properties['Response']
    if (-not $responseProperty) { return $null }
    return $responseProperty.Value
}

function Get-WebExceptionStatusCode {
    <# .SYNOPSIS Returns the HTTP status code of a failed web request, or 0 when there is none. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord)

    $response = Get-WebExceptionResponse -ErrorRecord $ErrorRecord
    if ($null -eq $response) { return 0 }
    $statusProperty = $response.PSObject.Properties['StatusCode']
    if (-not $statusProperty -or $null -eq $statusProperty.Value) { return 0 }
    try { return [int]$statusProperty.Value }
    catch { return 0 }
}

function Get-RetryAfterDelay {
    <# .SYNOPSIS Returns the server-requested retry delay, which SharePoint sends when it throttles. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord)

    $response = Get-WebExceptionResponse -ErrorRecord $ErrorRecord
    if ($null -eq $response) { return 0 }
    $headersProperty = $response.PSObject.Properties['Headers']
    if (-not $headersProperty -or $null -eq $headersProperty.Value) { return 0 }

    $headerValue = ''
    try { $headerValue = [string]($headersProperty.Value['Retry-After'] | Select-Object -First 1) }
    catch { Write-Verbose "Could not read the Retry-After header: $($_.Exception.Message)" }

    $seconds = 0
    if ([int]::TryParse($headerValue, [ref]$seconds) -and $seconds -gt 0) { return [math]::Min($seconds, 120) }
    return 0
}

function Test-TransientFailure {
    <# .SYNOPSIS Decides whether a failure is worth retrying rather than reporting to the operator. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord)

    $statusCode = Get-WebExceptionStatusCode -ErrorRecord $ErrorRecord
    if ($statusCode -gt 0) { return $statusCode -in 408, 425, 429, 500, 502, 503, 504 }
    # A host name that does not exist is a configuration error, so it is deliberately not retried.
    return [string]$ErrorRecord.Exception.Message -match '(?i)timed out|timeout|throttl|too many requests|temporarily|service unavailable|connection was closed|actively refused|transport connection|please try again'
}

function Invoke-WithRetry {
    <# .SYNOPSIS Runs an operation, retrying only transient failures with backoff and server-requested delays. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$Operation,
        [Parameter(Mandatory)][string]$Description,
        [int]$MaximumAttempts = 5,
        [int]$InitialDelaySeconds = 3
    )

    $delaySeconds = $InitialDelaySeconds
    for ($attempt = 1; $attempt -le $MaximumAttempts; $attempt++) {
        try { return & $Operation }
        catch {
            if ($attempt -ge $MaximumAttempts -or -not (Test-TransientFailure -ErrorRecord $_)) { throw }
            $retryAfterSeconds = Get-RetryAfterDelay -ErrorRecord $_
            $waitSeconds = if ($retryAfterSeconds -gt 0) { $retryAfterSeconds } else { $delaySeconds }
            Write-Step -Severity Warn -Message "$Description did not succeed on attempt $attempt of $MaximumAttempts. Retrying in $waitSeconds seconds. $($_.Exception.Message)"
            Start-Sleep -Seconds $waitSeconds
            $delaySeconds = [math]::Min($delaySeconds * 2, 60)
        }
    }
}

function Show-DetectedValuesBlock {
    <# .SYNOPSIS Prints all detected tenant-related values in a copy-friendly block. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TenantRootUrl,
        [Parameter(Mandatory)][string]$TenantAdminUrl,
        [Parameter(Mandatory)][string]$RecommendedTenantHint,
        [string]$TenantId = '',
        [string]$TenantDomain = ''
    )

    Write-Host ''
    Write-Host '  Detected values (copy/paste ready)' -ForegroundColor Cyan
    Write-Host "    SharePoint root URL         : $TenantRootUrl" -ForegroundColor Gray
    Write-Host "    SharePoint admin URL        : $TenantAdminUrl" -ForegroundColor Gray
    Write-Host "    Recommended tenant hint     : $RecommendedTenantHint" -ForegroundColor Gray
    if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
        Write-Host "    Tenant ID                   : $TenantId" -ForegroundColor Gray
    }
    if (-not [string]::IsNullOrWhiteSpace($TenantDomain)) {
        Write-Host "    Tenant domain               : $TenantDomain" -ForegroundColor Gray
    }
    Write-Host ''
}

function ConvertTo-UrlSlug {
    <# .SYNOPSIS Converts a display name into a SharePoint-safe URL segment. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)

    $slug = ($Name -replace '[^A-Za-z0-9]+', '-').Trim('-')
    if ([string]::IsNullOrWhiteSpace($slug)) { throw "Site name '$Name' contains no usable URL characters." }
    return $slug
}

function Get-TestFolderName {
    <# .SYNOPSIS Picks the name for one folder from the pool for its depth. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int]$Depth,
        [Parameter(Mandatory)][int]$Index
    )

    # Wrapped in @() because a one-name fallback pool would otherwise arrive as a bare string.
    $pool = @(if ($Depth -le $script:FolderNamesByDepth.Count) { $script:FolderNamesByDepth[$Depth - 1] } else { "Level$Depth" })
    $name = $pool[($Index - 1) % $pool.Count]
    # Past the end of the pool the names repeat, so a suffix keeps siblings unique.
    $round = [math]::Floor(($Index - 1) / $pool.Count)
    if ($round -gt 0) { $name = '{0}-{1}' -f $name, ($round + 1) }
    return $name
}

function Add-TestFolderBranch {
    <# .SYNOPSIS Appends one folder and everything beneath it, parent before child. #>
    [CmdletBinding()]
    param(
        # The list starts out empty, which a mandatory parameter otherwise rejects.
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[string]]$Paths,
        [Parameter(Mandatory)][AllowEmptyString()][string]$ParentPath,
        [Parameter(Mandatory)][int]$Depth,
        [Parameter(Mandatory)][int]$MaximumDepth,
        [Parameter(Mandatory)][int]$Index,
        [Parameter(Mandatory)][int]$SubfolderCount
    )

    $name = Get-TestFolderName -Depth $Depth -Index $Index
    $path = if ($ParentPath) { "$ParentPath/$name" } else { $name }
    $Paths.Add($path)
    if ($Depth -ge $MaximumDepth) { return }
    for ($child = 1; $child -le $SubfolderCount; $child++) {
        Add-TestFolderBranch -Paths $Paths -ParentPath $path -Depth ($Depth + 1) -MaximumDepth $MaximumDepth -Index $child -SubfolderCount $SubfolderCount
    }
}

function New-TestFolderTree {
    <# .SYNOPSIS Builds the library-relative folder paths to create, in creation order. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateRange(0, 100)][int]$TopLevelCount,
        [Parameter(Mandatory)][ValidateRange(1, 10)][int]$Depth,
        [Parameter(Mandatory)][ValidateRange(0, 20)][int]$SubfolderCount
    )

    $paths = [System.Collections.Generic.List[string]]::new()
    for ($index = 1; $index -le $TopLevelCount; $index++) {
        Add-TestFolderBranch -Paths $paths -ParentPath '' -Depth 1 -MaximumDepth $Depth -Index $index -SubfolderCount $SubfolderCount
    }
    return @($paths)
}

function New-TestFilePlan {
    <# .SYNOPSIS Draws how many files, and how many of them sensitive, each folder gets. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int]$FolderCount,
        [Parameter(Mandatory)][psobject]$FileRange,
        [Parameter(Mandatory)][psobject]$SensitiveRange
    )

    $plan = [System.Collections.Generic.List[psobject]]::new()
    for ($index = 0; $index -lt $FolderCount; $index++) {
        $files = Get-Random -Minimum $FileRange.Minimum -Maximum ($FileRange.Maximum + 1)
        $sensitive = Get-Random -Minimum $SensitiveRange.Minimum -Maximum ($SensitiveRange.Maximum + 1)
        $plan.Add([pscustomobject]@{
                FileCount      = $files
                SensitiveCount = [math]::Min($sensitive, $files)
            })
    }
    return @($plan)
}

function Add-ZipTextEntry {
    <# .SYNOPSIS Adds one UTF-8 XML part to an open OOXML package. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.IO.Compression.ZipArchive]$Archive,
        [Parameter(Mandatory)][string]$EntryName,
        [Parameter(Mandatory)][string]$Content
    )

    $entry = $Archive.CreateEntry($EntryName)
    $stream = $entry.Open()
    try {
        $writer = [System.IO.StreamWriter]::new($stream, [System.Text.UTF8Encoding]::new($false))
        try { $writer.Write($Content) } finally { $writer.Dispose() }
    }
    finally { $stream.Dispose() }
}

function Get-TestFileContent {
    <# .SYNOPSIS Builds the body of one test file, seeded with fabricated sensitive data when asked. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FolderLabel,
        [Parameter(Mandatory)][int]$FileNumber,
        [switch]$IncludeSensitiveData
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add('Contoso Ltd - Microsoft Purview sensitivity label test file')
    $lines.Add("Library folder: $FolderLabel    Item: $FileNumber")
    $lines.Add("Created by New-LabelTestSite.ps1 on $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
    $lines.Add('')

    if (-not $IncludeSensitiveData) {
        $lines.Add('This file holds ordinary business text and nothing a classifier should match. Use it to confirm that a label applied by hand or by policy also reaches unremarkable content, and that automatic policies leave it alone.')
        $lines.Add('')
        $lines.Add('Quarterly review notes')
        $lines.Add('1. Confirm the reporting calendar with each regional lead.')
        $lines.Add('2. Publish the consolidated figures to the finance workspace.')
        $lines.Add('3. Archive the previous quarter working files.')
        $lines.Add('4. Book the follow-up session for the operations team.')
        $lines.Add('')
        $lines.Add('Status: open. Owner: the review team. No customer or personal data is stored in this file.')
        return @($lines)
    }

    $records = @($script:SensitiveSampleRecords)
    $record = @($records[($FileNumber - 1) % $records.Count])
    $lines.Add('CONFIDENTIAL - fabricated personal and payment data for classifier testing.')
    $lines.Add('Every value below is invented for testing, in the spirit of the public sample sets at dlptest.com. None of it identifies a real person, account, or payment card.')
    $lines.Add('It contains a social security number, credit card number, bank account number, passport number, driver''s license number, and taxpayer identification number, each written next to its keyword so Purview sensitive information types can match both the pattern and its supporting evidence.')
    $lines.Add('')
    foreach ($recordLine in $record) { $lines.Add($recordLine) }
    $lines.Add('')
    $lines.Add('Handling: restricted customer file. Do not distribute outside the review team.')
    return @($lines)
}

function New-TestOfficeFile {
    <# .SYNOPSIS Writes a minimal but valid .docx or .xlsx that Purview can read and label. #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][ValidateSet('docx', 'xlsx')][string]$Format,
        # Blank lines are deliberate spacing, which a mandatory parameter otherwise rejects.
        [Parameter(Mandatory)][AllowEmptyString()][string[]]$Line
    )

    if (-not $PSCmdlet.ShouldProcess($Path, "Create test $Format")) { return }

    $encodedLines = @($Line | ForEach-Object { [System.Security.SecurityElement]::Escape([string]$_) })
    $relationshipsNs = 'http://schemas.openxmlformats.org/package/2006/relationships'
    $officeDocumentType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'

    $stream = [System.IO.File]::Open($Path, 'Create', 'Write', 'None')
    try {
        $archive = [System.IO.Compression.ZipArchive]::new($stream, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            if ($Format -eq 'docx') {
                Add-ZipTextEntry -Archive $archive -EntryName '[Content_Types].xml' -Content (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
                    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
                    '<Default Extension="xml" ContentType="application/xml"/>' +
                    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
                    '</Types>')
                Add-ZipTextEntry -Archive $archive -EntryName '_rels/.rels' -Content (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    "<Relationships xmlns=`"$relationshipsNs`">" +
                    "<Relationship Id=`"rId1`" Type=`"$officeDocumentType`" Target=`"word/document.xml`"/>" +
                    '</Relationships>')
                Add-ZipTextEntry -Archive $archive -EntryName 'word/document.xml' -Content (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
                    '<w:body>' +
                    (($encodedLines | ForEach-Object { "<w:p><w:r><w:t xml:space=`"preserve`">$_</w:t></w:r></w:p>" }) -join '') +
                    '</w:body>' +
                    '</w:document>')
            }
            else {
                Add-ZipTextEntry -Archive $archive -EntryName '[Content_Types].xml' -Content (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
                    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
                    '<Default Extension="xml" ContentType="application/xml"/>' +
                    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
                    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
                    '</Types>')
                Add-ZipTextEntry -Archive $archive -EntryName '_rels/.rels' -Content (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    "<Relationships xmlns=`"$relationshipsNs`">" +
                    "<Relationship Id=`"rId1`" Type=`"$officeDocumentType`" Target=`"xl/workbook.xml`"/>" +
                    '</Relationships>')
                Add-ZipTextEntry -Archive $archive -EntryName 'xl/workbook.xml' -Content (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
                    '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>')
                Add-ZipTextEntry -Archive $archive -EntryName 'xl/_rels/workbook.xml.rels' -Content (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    "<Relationships xmlns=`"$relationshipsNs`">" +
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
                    '</Relationships>')
                $rows = for ($rowNumber = 1; $rowNumber -le $encodedLines.Count; $rowNumber++) {
                    "<row r=`"$rowNumber`"><c r=`"A$rowNumber`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$($encodedLines[$rowNumber - 1])</t></is></c></row>"
                }
                Add-ZipTextEntry -Archive $archive -EntryName 'xl/worksheets/sheet1.xml' -Content (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
                    "<sheetData>$($rows -join '')</sheetData>" +
                    '</worksheet>')
            }
        }
        finally { $archive.Dispose() }
    }
    finally { $stream.Dispose() }
}

function Get-PowerShellHostVersion {
    <# .SYNOPSIS Reads a pwsh.exe version from file metadata, falling back to running it. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $versionText = ''
    try { $versionText = [string](Get-Item -LiteralPath $Path -ErrorAction Stop).VersionInfo.ProductVersion }
    catch { Write-Verbose "Could not read version metadata from ${Path}: $($_.Exception.Message)" }

    $versionMatch = [regex]::Match($versionText, '^\d+\.\d+(\.\d+)?')
    if (-not $versionMatch.Success) {
        # Store execution aliases carry no file metadata, so ask the executable itself.
        try { $versionText = [string](& $Path -NoLogo -NoProfile -Command '$PSVersionTable.PSVersion.ToString()' 2>$null) }
        catch { Write-Verbose "Could not run ${Path}: $($_.Exception.Message)" }
        $versionMatch = [regex]::Match($versionText, '^\d+\.\d+(\.\d+)?')
    }
    if (-not $versionMatch.Success) { return $null }
    try { return [version]$versionMatch.Value }
    catch { return $null }
}

function Get-PowerShell7Path {
    <# .SYNOPSIS Returns the newest installed PowerShell 7 host that meets a minimum version. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][version]$MinimumVersion)

    $candidatePaths = [System.Collections.Generic.List[string]]::new()
    foreach ($command in @(Get-Command pwsh.exe -CommandType Application -ErrorAction SilentlyContinue)) {
        $candidatePaths.Add([string]$command.Source)
    }
    foreach ($root in $env:ProgramFiles, ${env:ProgramFiles(x86)}) {
        if ([string]::IsNullOrWhiteSpace($root)) { continue }
        $candidatePaths.Add((Join-Path $root 'PowerShell\7\pwsh.exe'))
    }
    if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
        $candidatePaths.Add((Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\pwsh.exe'))
    }

    $bestPath = ''
    $bestVersion = [version]'0.0'
    foreach ($candidatePath in $candidatePaths) {
        if ([string]::IsNullOrWhiteSpace($candidatePath) -or -not (Test-Path -LiteralPath $candidatePath -PathType Leaf)) { continue }
        $candidateVersion = Get-PowerShellHostVersion -Path $candidatePath
        if ($null -eq $candidateVersion -or $candidateVersion -lt $MinimumVersion -or $candidateVersion -le $bestVersion) { continue }
        $bestPath = $candidatePath
        $bestVersion = $candidateVersion
    }
    return $bestPath
}

function Get-RelaunchArgumentList {
    <# .SYNOPSIS Builds the pwsh argument list that reruns this script with the same parameters. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][hashtable]$BoundParameters
    )

    if (-not (Test-Path -LiteralPath $ScriptPath -PathType Leaf)) {
        throw "Cannot restart in PowerShell 7 because the script path '$ScriptPath' was not found."
    }

    $arguments = [System.Collections.Generic.List[string]]::new()
    $arguments.Add('-NoLogo')
    $arguments.Add('-File')
    $arguments.Add($ScriptPath)
    $arguments.Add('-NoRelaunch')
    foreach ($name in $BoundParameters.Keys) {
        if ($name -eq 'NoRelaunch') { continue }
        $value = $BoundParameters[$name]
        if ($value -is [switch]) {
            if ($value.IsPresent) { $arguments.Add("-$name") }
            continue
        }
        $arguments.Add("-$name")
        $arguments.Add([string]$value)
    }
    return @($arguments)
}

function Install-PnPPowerShellModule {
    <# .SYNOPSIS Installs PnP.PowerShell for the current user, choosing a version the host can load. #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param()

    if (-not $PSCmdlet.ShouldProcess('PnP.PowerShell', 'Install module for the current user')) { return }

    # The PowerShell Gallery requires TLS 1.2, which Windows PowerShell does not always negotiate by default.
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    }
    catch { Write-Verbose "Could not raise the TLS version: $($_.Exception.Message)" }

    try { $null = Get-PackageProvider -Name NuGet -ForceBootstrap -ErrorAction Stop }
    catch { Write-Verbose "NuGet provider bootstrap skipped: $($_.Exception.Message)" }

    $parameters = @{
        Name = 'PnP.PowerShell'
        Scope = 'CurrentUser'
        Force = $true
        AllowClobber = $true
        ErrorAction = 'Stop'
    }
    # PnP.PowerShell 1.12.0 is the last release for Windows PowerShell; 2.12.0 supports PowerShell 7.2 and 7.3.
    if ($PSVersionTable.PSVersion -lt [version]'7.2.0') {
        $parameters.RequiredVersion = '1.12.0'
        Write-Step -Message 'Host is older than PowerShell 7.2, so PnP.PowerShell 1.12.0 will be installed.'
    }
    elseif ($PSVersionTable.PSVersion -lt [version]'7.4.0') {
        $parameters.RequiredVersion = '2.12.0'
        Write-Step -Message 'Host is older than PowerShell 7.4, so PnP.PowerShell 2.12.0 will be installed.'
    }
    Install-Module @parameters
}

function Test-PnPPrerequisite {
    <# .SYNOPSIS Ensures PnP.PowerShell is installed and imported, installing it when missing. #>
    [CmdletBinding()]
    param()

    try {
        $maximumVersion = if ($PSVersionTable.PSVersion -lt [version]'7.2.0') {
            [version]'1.12.0'
        }
        elseif ($PSVersionTable.PSVersion -lt [version]'7.4.0') {
            [version]'2.12.0'
        }
        else {
            $null
        }
        $availableModules = Get-Module -ListAvailable -Name PnP.PowerShell -ErrorAction Stop
        # Plain -ListAvailable reports only the newest version, so an older copy that can shadow it stays invisible.
        $allVersions = @(Get-Module -ListAvailable -Name PnP.PowerShell -All -ErrorAction SilentlyContinue |
                Where-Object { $_.Version -and [version]$_.Version -ne [version]'0.0.0.0' } |
                Select-Object -ExpandProperty Version -Unique | Sort-Object -Descending)
        if ($allVersions.Count -gt 1) {
            Write-Step -Severity Warn -Message "PnP.PowerShell $($allVersions[0]) is installed alongside $(($allVersions | Select-Object -Skip 1) -join ', '). Whichever version a session loads first wins, because PnP assemblies cannot be unloaded, so remove the older copies with: Uninstall-Module PnP.PowerShell -RequiredVersion <version> -Force"
        }
        if ($null -ne $maximumVersion) {
            $availableModules = $availableModules | Where-Object Version -LE $maximumVersion
        }
        $module = $availableModules |
            Sort-Object Version -Descending | Select-Object -First 1
        if (-not $module) {
            Write-Step -Message 'PnP.PowerShell is not installed. Installing it for the current user...'
            Install-PnPPowerShellModule -Confirm:$false
            $availableModules = Get-Module -ListAvailable -Name PnP.PowerShell -ErrorAction Stop
            if ($null -ne $maximumVersion) {
                $availableModules = $availableModules | Where-Object Version -LE $maximumVersion
            }
            $module = $availableModules |
                Sort-Object Version -Descending | Select-Object -First 1
            if (-not $module) { throw 'PnP.PowerShell could not be found after the installation attempt.' }
            Write-Step -Severity Good -Message "Installed PnP.PowerShell $($module.Version)."
        }
        Import-Module PnP.PowerShell -RequiredVersion $module.Version -ErrorAction Stop
        foreach ($commandName in 'Connect-PnPOnline', 'Disconnect-PnPOnline', 'Get-PnPConnection', 'Get-PnPWeb', 'Get-PnPTenant', 'New-PnPSite', 'New-PnPList', 'Add-PnPFolder', 'Add-PnPFile') {
            if (-not (Get-Command $commandName -ErrorAction SilentlyContinue)) {
                throw "PnP.PowerShell $($module.Version) does not provide required command $commandName."
            }
        }
        $connectCommand = Get-Command Connect-PnPOnline -ErrorAction Stop
        foreach ($parameterName in 'Url', 'ClientId', 'Interactive', 'DeviceLogin', 'Tenant') {
            if (-not $connectCommand.Parameters.ContainsKey($parameterName)) {
                throw "Connect-PnPOnline in PnP.PowerShell $($module.Version) does not support required parameter -$parameterName."
            }
        }
        Write-Step -Severity Good -Message "PnP.PowerShell $($module.Version) is ready."
        return $true
    }
    catch {
        Write-Step -Severity Error -Message $_.Exception.Message
        Write-Step -Message 'Install manually with: Install-Module PnP.PowerShell -Scope CurrentUser -Force'
        Write-Step -Message 'On Windows PowerShell 5.1 use: Install-Module PnP.PowerShell -RequiredVersion 1.12.0 -Force'
        return $false
    }
}

function ConvertTo-GuidString {
    <# .SYNOPSIS Returns a normalized GUID string, or an empty string when the value is not a GUID. #>
    [CmdletBinding()]
    param(
        [AllowNull()][object]$Value
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '' }
    $parsedValue = [guid]::Empty
    if (-not [guid]::TryParse(([string]$Value).Trim(), [ref]$parsedValue) -or $parsedValue -eq [guid]::Empty) {
        return ''
    }
    return $parsedValue.ToString()
}

function Get-TenantClientIdEnvironmentVariableName {
    <# .SYNOPSIS Builds the per-tenant environment-variable name used to remember a generated client ID. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$TenantId)

    $normalizedTenantId = ConvertTo-GuidString -Value $TenantId
    if ([string]::IsNullOrWhiteSpace($normalizedTenantId)) { throw 'A valid resolved tenant ID is required.' }
    return 'LABEL_TEST_SITE_CLIENT_ID_' + $normalizedTenantId.Replace('-', '').ToUpperInvariant()
}

function Get-PnPClientIdCandidate {
    <# .SYNOPSIS Returns distinct Entra application IDs to validate for the resolved tenant. #>
    [CmdletBinding()]
    param(
        [string]$ExplicitClientId = '',
        [Parameter(Mandatory)][string]$TenantId
    )

    $tenantVariableName = Get-TenantClientIdEnvironmentVariableName -TenantId $TenantId
    $candidates = [System.Collections.Generic.List[object]]::new()
    if (-not [string]::IsNullOrWhiteSpace($ExplicitClientId)) {
        $candidates.Add([pscustomobject]@{Source = 'the -ClientId parameter'; Value = $ExplicitClientId; IsExplicit = $true})
    }
    else {
        foreach ($variableName in $tenantVariableName, 'ENTRAID_APP_ID', 'ENTRAID_CLIENT_ID', 'AZURE_CLIENT_ID') {
            $processValue = [Environment]::GetEnvironmentVariable($variableName, [EnvironmentVariableTarget]::Process)
            $candidates.Add([pscustomobject]@{Source = "process environment variable $variableName"; Value = $processValue; IsExplicit = $false})
            try {
                $userValue = [Environment]::GetEnvironmentVariable($variableName, [EnvironmentVariableTarget]::User)
                $candidates.Add([pscustomobject]@{Source = "user environment variable $variableName"; Value = $userValue; IsExplicit = $false})
            }
            catch { Write-Verbose "Could not read user environment variable ${variableName}: $($_.Exception.Message)" }
        }
    }

    $resolvedCandidates = [System.Collections.Generic.List[string]]::new()
    foreach ($item in $candidates) {
        $candidate = [string]$item.Value
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }

        $candidate = $candidate.Trim()
        $normalizedClientId = ConvertTo-GuidString -Value $candidate
        if ([string]::IsNullOrWhiteSpace($normalizedClientId)) {
            if ($item.IsExplicit) {
                throw "The PnP Entra application ID from $($item.Source) is not a valid GUID: '$candidate'."
            }
            Write-Step -Severity Warn -Message "Ignoring the invalid PnP Entra application ID in $($item.Source)."
            continue
        }
        if ($resolvedCandidates.Contains($normalizedClientId)) { continue }
        Write-Step -Message "Found a PnP Entra application ID in $($item.Source): $normalizedClientId"
        $resolvedCandidates.Add($normalizedClientId)
    }

    return @($resolvedCandidates)
}

function Resolve-SharePointRootUrl {
    <# .SYNOPSIS Normalizes a SharePoint URL, tenant domain, or bare alias into root and admin URLs. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Url)

    $value = $Url.Trim().Trim('"', "'")
    if ([string]::IsNullOrWhiteSpace($value)) { throw 'A SharePoint URL, tenant domain, or tenant alias is required.' }
    $parsedTenantId = [guid]::Empty
    if ([guid]::TryParse($value, [ref]$parsedTenantId)) {
        throw "'$Url' is a tenant ID, which does not identify a SharePoint host. Supply the SharePoint URL, the <tenant>.onmicrosoft.com domain, or the tenant alias."
    }

    if ($value -match '^[A-Za-z][A-Za-z0-9+.\-]*://') {
        $uri = $null
        if (-not [uri]::TryCreate($value, [UriKind]::Absolute, [ref]$uri)) { throw "'$Url' is not a valid URL." }
        if (-not [string]::IsNullOrWhiteSpace($uri.UserInfo)) { throw "SharePoint tenant URL '$Url' must not contain credentials." }
        $value = $uri.Host
    }

    $value = $value.Split('/')[0].Split('?')[0].Split('#')[0].Split(':')[0]
    if ($value.Contains('@')) { $value = $value.Split('@')[-1] }
    $value = $value.Trim().TrimEnd('.').ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($value)) { throw "Could not determine a tenant host from '$Url'." }

    $sharePointMatch = [regex]::Match($value, '^(?<alias>[a-z0-9][a-z0-9-]*?)(?:-admin|-my)?\.sharepoint\.(?<cloud>com|us|de|cn)$')
    $initialDomainMatch = [regex]::Match($value, '^(?<alias>[a-z0-9][a-z0-9-]*)\.(?<domain>partner\.onmschina\.cn|onmicrosoft\.com|onmicrosoft\.us|onmicrosoft\.de)$')
    if ($sharePointMatch.Success) {
        $tenantAlias = $sharePointMatch.Groups['alias'].Value
        $cloudSuffix = $sharePointMatch.Groups['cloud'].Value
    }
    elseif ($initialDomainMatch.Success) {
        $tenantAlias = $initialDomainMatch.Groups['alias'].Value
        $cloudSuffix = switch ($initialDomainMatch.Groups['domain'].Value) {
            'onmicrosoft.us' { 'us' }
            'onmicrosoft.de' { 'de' }
            'partner.onmschina.cn' { 'cn' }
            default { 'com' }
        }
    }
    elseif ($value -match '\.sharepoint(-[a-z0-9]+)?\.') {
        throw "SharePoint host '$value' is not in a cloud this helper supports (commercial, US Government, Germany, or 21Vianet)."
    }
    elseif ($value -notmatch '\.') {
        if ($value -notmatch '^[a-z0-9][a-z0-9-]*$') { throw "'$Url' is not a usable SharePoint tenant alias." }
        $tenantAlias = $value
        $cloudSuffix = 'com'
        Write-Step -Message "Treating '$value' as the tenant alias, so https://$value.sharepoint.com will be verified."
    }
    else {
        # A custom verified domain does not reveal the SharePoint alias, so infer it and let discovery confirm.
        $tenantAlias = $value.Split('.')[0]
        $cloudSuffix = 'com'
        if ($tenantAlias -notmatch '^[a-z0-9][a-z0-9-]*$') { throw "Could not derive a SharePoint tenant alias from '$Url'." }
        Write-Step -Severity Warn -Message "'$value' is not a SharePoint or Microsoft 365 tenant domain, so the alias '$tenantAlias' was inferred. Pass the SharePoint URL if that is wrong."
    }

    $rootHost = "$tenantAlias.sharepoint.$cloudSuffix"
    return [pscustomobject]@{
        RootUrl = "https://$rootHost"
        AdminUrl = "https://$tenantAlias-admin.sharepoint.$cloudSuffix"
        Host = $rootHost
    }
}

function Test-DnsHostResolution {
    <# .SYNOPSIS Reports whether a host name resolves through DNS. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$HostName)

    try { return ([System.Net.Dns]::GetHostAddresses($HostName)).Count -gt 0 }
    catch {
        Write-Verbose "DNS lookup for '$HostName' failed: $($_.Exception.Message)"
        return $false
    }
}

function Get-CloudLoginAuthority {
    <# .SYNOPSIS Returns the sign-in authority that matches a SharePoint host's cloud. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$HostName)

    switch ($HostName.Substring($HostName.LastIndexOf('.') + 1)) {
        'us' { 'https://login.microsoftonline.us' }
        'de' { 'https://login.microsoftonline.de' }
        'cn' { 'https://login.partner.microsoftonline.cn' }
        default { 'https://login.microsoftonline.com' }
    }
}

function Get-EntraTenantIdForDomain {
    <# .SYNOPSIS Best-effort tenant lookup used only to explain an unreachable SharePoint address. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Domain,
        [Parameter(Mandatory)][string]$LoginAuthority
    )

    if ([string]::IsNullOrWhiteSpace($Domain)) { return '' }
    try {
        $metadataUrl = "$LoginAuthority/$([uri]::EscapeDataString($Domain))/v2.0/.well-known/openid-configuration"
        $metadata = Invoke-RestMethod -Uri $metadataUrl -Method Get -TimeoutSec 20 -ErrorAction Stop
        foreach ($segment in ([uri][string]$metadata.issuer).AbsolutePath.Trim('/').Split('/')) {
            $tenantId = ConvertTo-GuidString -Value $segment
            if (-not [string]::IsNullOrWhiteSpace($tenantId)) { return $tenantId }
        }
    }
    catch { Write-Verbose "Could not resolve tenant '$Domain': $($_.Exception.Message)" }
    return ''
}

function Test-SharePointHostReachable {
    <# .SYNOPSIS Confirms a SharePoint host exists and explains precisely when it does not. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$HostName,
        [Parameter(Mandatory)][string]$SuppliedValue
    )

    if (Test-DnsHostResolution -HostName $HostName) { return $true }

    $loginAuthority = Get-CloudLoginAuthority -HostName $HostName
    $controlHost = ([uri]$loginAuthority).Host
    if (-not (Test-DnsHostResolution -HostName $controlHost)) {
        # Proxy-only networks cannot resolve anything locally, so do not block on the pre-check.
        Write-Step -Severity Warn -Message "This machine cannot resolve $controlHost either, so '$HostName' was not pre-checked."
        return $true
    }

    Write-Step -Severity Error -Message "'$HostName' does not exist in DNS, so it is not this tenant's SharePoint address."
    $tenantDomain = ConvertTo-TenantHint -Value $SuppliedValue
    $tenantId = Get-EntraTenantIdForDomain -Domain $tenantDomain -LoginAuthority $loginAuthority
    if (-not [string]::IsNullOrWhiteSpace($tenantId)) {
        Write-Step -Message "The Microsoft 365 tenant '$tenantDomain' does exist (tenant $tenantId), so only the SharePoint address is wrong."
    }
    Write-Step -Message 'A SharePoint address does not always match the onmicrosoft.com prefix, and some tenants have no SharePoint site provisioned.'
    Write-Step -Message 'Open any SharePoint site or the SharePoint admin center in a browser and supply the address shown there.'
    return $false
}

function Resolve-PnPAzureEnvironment {
    <# .SYNOPSIS Maps a verified SharePoint cloud and login authority to the matching PnP environment. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RootUrl,
        [Parameter(Mandatory)][string]$LoginAuthority
    )

    $sharePointHost = ([uri]$RootUrl).Host.ToLowerInvariant()
    $authorityHost = ([uri]$LoginAuthority).Host.ToLowerInvariant()
    $cloudSuffix = $sharePointHost.Substring($sharePointHost.LastIndexOf('.') + 1)
    $environment = switch ("$cloudSuffix|$authorityHost") {
        'com|login.microsoftonline.com' { 'Production' }
        'us|login.microsoftonline.us' { 'USGovernmentHigh' }
        'de|login.microsoftonline.de' { 'Germany' }
        'cn|login.chinacloudapi.cn' { 'China' }
        'cn|login.partner.microsoftonline.cn' { 'China' }
        default { '' }
    }
    if ([string]::IsNullOrWhiteSpace($environment)) {
        throw "SharePoint host '$sharePointHost' and login authority '$authorityHost' do not identify the same supported Microsoft cloud."
    }
    return $environment
}

function Get-SharePointTenantChallenge {
    <# .SYNOPSIS Reads SharePoint's anonymous authentication challenge to discover the tenant realm and login authority. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$RootUrl)

    $response = $null
    try {
        $response = Invoke-WebRequest -Uri "$($RootUrl.TrimEnd('/'))/_vti_bin/client.svc/" -Method Get -Headers @{ Authorization = 'Bearer' } -UseBasicParsing -MaximumRedirection 0 -TimeoutSec 20 -ErrorAction Stop
    }
    catch {
        $response = Get-WebExceptionResponse -ErrorRecord $_
        if ($null -eq $response) {
            throw "Could not reach SharePoint at '$RootUrl': $($_.Exception.Message)"
        }
    }

    $authenticationHeader = ''
    if ($null -ne $response -and $null -ne $response.Headers) {
        try {
            # PowerShell 7 returns HttpResponseHeaders, which has no string indexer; Windows PowerShell returns WebHeaderCollection, which has no TryGetValues.
            if ($response.Headers.PSObject.Methods['TryGetValues']) {
                $values = $null
                if ($response.Headers.TryGetValues('WWW-Authenticate', [ref]$values)) { $authenticationHeader = ($values -join ',') }
            }
            else {
                $authenticationHeader = [string]$response.Headers['WWW-Authenticate']
            }
        }
        catch { Write-Verbose "Could not read the SharePoint authentication header: $($_.Exception.Message)" }
        if ([string]::IsNullOrWhiteSpace($authenticationHeader)) {
            $headerProperty = $response.Headers.PSObject.Properties['WwwAuthenticate']
            if ($headerProperty) { $authenticationHeader = [string]$headerProperty.Value }
        }
    }

    $tenantId = ''
    $loginAuthority = ''
    if (-not [string]::IsNullOrWhiteSpace($authenticationHeader)) {
        $realmMatch = [regex]::Match($authenticationHeader, '(?i)\brealm\s*=\s*"?(?<value>[^",\s]+)')
        if ($realmMatch.Success) { $tenantId = ConvertTo-GuidString -Value $realmMatch.Groups['value'].Value }

        $authorityMatch = [regex]::Match($authenticationHeader, '(?i)authorization_uri\s*=\s*"(?<value>[^"]+)"')
        if ($authorityMatch.Success) {
            $authorityUri = $null
            if ([uri]::TryCreate($authorityMatch.Groups['value'].Value, [UriKind]::Absolute, [ref]$authorityUri)) {
                $loginAuthority = $authorityUri.GetLeftPart([UriPartial]::Authority)
            }
        }
    }

    return [pscustomobject]@{
        TenantId = $tenantId
        LoginAuthority = $loginAuthority
    }
}

function Resolve-EntraTenantDetail {
    <# .SYNOPSIS Resolves a SharePoint tenant to its actual Entra tenant ID and OAuth endpoints. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RootUrl,
        [string]$PreferredTenant = ''
    )

    $challenge = Invoke-WithRetry -Description 'SharePoint tenant discovery' -Operation { Get-SharePointTenantChallenge -RootUrl $RootUrl }
    if ([string]::IsNullOrWhiteSpace($challenge.TenantId)) {
        throw "SharePoint did not return a tenant realm for '$RootUrl'. Refusing to infer a tenant before sign-in."
    }
    if ([string]::IsNullOrWhiteSpace($challenge.LoginAuthority)) {
        throw "SharePoint did not return a login authority for '$RootUrl'. Refusing to guess an identity cloud."
    }
    $tenantCandidate = if (-not [string]::IsNullOrWhiteSpace($PreferredTenant)) {
        ConvertTo-TenantHint -Value $PreferredTenant
    }
    else {
        $challenge.TenantId
    }
    if ([string]::IsNullOrWhiteSpace($tenantCandidate)) { throw 'Could not derive an Entra tenant from the SharePoint URL.' }

    $loginAuthority = $challenge.LoginAuthority.TrimEnd('/')
    $azureEnvironment = Resolve-PnPAzureEnvironment -RootUrl $RootUrl -LoginAuthority $loginAuthority
    $encodedTenant = [uri]::EscapeDataString($tenantCandidate)
    $metadataUrl = "$loginAuthority/$encodedTenant/v2.0/.well-known/openid-configuration"
    try {
        $metadata = Invoke-WithRetry -Description 'Entra tenant metadata lookup' -Operation {
            Invoke-RestMethod -Uri $metadataUrl -Method Get -TimeoutSec 20 -ErrorAction Stop
        }
    }
    catch {
        throw "Could not resolve Entra tenant '$tenantCandidate' through $loginAuthority. Verify -TenantRootUrl and -TenantHint. $($_.Exception.Message)"
    }

    $resolvedTenantId = ''
    foreach ($propertyName in 'issuer', 'authorization_endpoint', 'token_endpoint') {
        $property = $metadata.PSObject.Properties[$propertyName]
        if (-not $property -or [string]::IsNullOrWhiteSpace([string]$property.Value)) { continue }
        $endpointUri = $null
        if (-not [uri]::TryCreate([string]$property.Value, [UriKind]::Absolute, [ref]$endpointUri)) { continue }
        foreach ($segment in $endpointUri.AbsolutePath.Trim('/').Split('/')) {
            $resolvedTenantId = ConvertTo-GuidString -Value $segment
            if (-not [string]::IsNullOrWhiteSpace($resolvedTenantId)) { break }
        }
        if (-not [string]::IsNullOrWhiteSpace($resolvedTenantId)) { break }
    }
    if ([string]::IsNullOrWhiteSpace($resolvedTenantId)) {
        throw "OpenID metadata for '$tenantCandidate' did not contain a concrete tenant ID."
    }
    if (-not [string]::IsNullOrWhiteSpace($challenge.TenantId) -and $resolvedTenantId -ne $challenge.TenantId) {
        throw "Tenant hint '$tenantCandidate' resolves to $resolvedTenantId, but SharePoint reports tenant $($challenge.TenantId). Refusing a cross-tenant sign-in."
    }

    $deviceEndpointProperty = $metadata.PSObject.Properties['device_authorization_endpoint']
    if (-not $deviceEndpointProperty -or [string]::IsNullOrWhiteSpace([string]$deviceEndpointProperty.Value)) {
        throw "OpenID metadata for tenant $resolvedTenantId does not advertise a device authorization endpoint."
    }

    return [pscustomobject]@{
        TenantId = $resolvedTenantId
        TenantHint = $tenantCandidate
        LoginAuthority = $loginAuthority
        AzureEnvironment = $azureEnvironment
        DeviceAuthorizationEndpoint = [string]$deviceEndpointProperty.Value
    }
}

function Test-PnPClientApplication {
    <# .SYNOPSIS Validates that a public-client application can request SharePoint access in the resolved tenant. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][object]$TenantMetadata,
        [Parameter(Mandatory)][string]$SharePointRootUrl
    )

    $normalizedClientId = ConvertTo-GuidString -Value $ClientId
    if ([string]::IsNullOrWhiteSpace($normalizedClientId)) { throw "Client ID '$ClientId' is not a valid GUID." }

    $body = @{
        client_id = $normalizedClientId
        scope = "openid offline_access $($SharePointRootUrl.TrimEnd('/'))/AllSites.FullControl"
    }
    try {
        $null = Invoke-WithRetry -Description 'Entra application preflight' -Operation {
            Invoke-RestMethod -Uri $TenantMetadata.DeviceAuthorizationEndpoint -Method Post -Body $body -ContentType 'application/x-www-form-urlencoded' -TimeoutSec 20 -ErrorAction Stop
        }
        return [pscustomobject]@{
            Status = 'Ready'
            Message = "Application $normalizedClientId is available as a public client in tenant $($TenantMetadata.TenantId)."
        }
    }
    catch {
        $rawMessage = if (-not [string]::IsNullOrWhiteSpace([string]$_.ErrorDetails.Message)) {
            [string]$_.ErrorDetails.Message
        }
        else {
            [string]$_.Exception.Message
        }
        $errorPayload = $null
        try { $errorPayload = $rawMessage | ConvertFrom-Json -ErrorAction Stop }
        catch { Write-Verbose "The Entra error response was not JSON: $($_.Exception.Message)" }

        $errorDescription = if ($null -ne $errorPayload -and $errorPayload.PSObject.Properties['error_description']) {
            [string]$errorPayload.error_description
        }
        else {
            $rawMessage
        }
        $errorName = if ($null -ne $errorPayload -and $errorPayload.PSObject.Properties['error']) {
            [string]$errorPayload.error
        }
        else {
            ''
        }
        $errorCodes = if ($null -ne $errorPayload -and $errorPayload.PSObject.Properties['error_codes']) {
            @($errorPayload.error_codes)
        }
        else {
            @()
        }

        $status = if ($errorCodes -contains 700016 -or $errorDescription -match 'AADSTS700016') {
            'Missing'
        }
        elseif ($errorName -in 'invalid_client', 'unauthorized_client', 'invalid_scope') {
            'Misconfigured'
        }
        else {
            'Error'
        }
        return [pscustomobject]@{
            Status = $status
            Message = $errorDescription
        }
    }
}

function Get-PnPInteractiveAppRegistrationCommand {
    <# .SYNOPSIS Finds the current or legacy PnP cmdlet that registers an interactive Entra application. #>
    [CmdletBinding()]
    param()

    foreach ($commandName in 'Register-PnPEntraIDAppForInteractiveLogin', 'Register-PnPAzureADAppForInteractiveLogin') {
        $command = Get-Command $commandName -ErrorAction SilentlyContinue
        if ($command) { return $command }
    }
    return $null
}

function Register-PnPInteractiveApplication {
    <# .SYNOPSIS Registers a tenant-owned public client with only the delegated SharePoint permission this helper needs. #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ApplicationName,
        [ValidateSet('Interactive', 'DeviceCode')][string]$AuthenticationMode = 'Interactive',
        [Parameter(Mandatory)][string]$AzureEnvironment
    )

    $registrationCommand = Get-PnPInteractiveAppRegistrationCommand
    if (-not $registrationCommand) {
        throw 'This PnP.PowerShell version cannot register an interactive Entra app automatically. Run the script in PowerShell 7.2 or later, or pass an existing tenant app with -ClientId.'
    }
    if (-not $PSCmdlet.ShouldProcess("tenant $TenantId", "Register Entra application '$ApplicationName'")) { return '' }

    $parameters = @{
        ApplicationName = $ApplicationName
        Tenant = $TenantId
        ErrorAction = 'Stop'
    }
    if ($registrationCommand.Parameters.ContainsKey('SharePointDelegatePermissions')) {
        $parameters.SharePointDelegatePermissions = @('AllSites.FullControl')
    }
    elseif ($registrationCommand.Parameters.ContainsKey('Scopes')) {
        $parameters.Scopes = @('AllSites.FullControl')
    }
    else {
        throw "$($registrationCommand.Name) cannot accept an explicit SharePoint permission set. Registration was not attempted."
    }
    if ($AuthenticationMode -eq 'DeviceCode' -and $registrationCommand.Parameters.ContainsKey('DeviceLogin')) {
        $parameters.DeviceLogin = $true
    }
    if ($registrationCommand.Parameters.ContainsKey('AzureEnvironment')) {
        $parameters.AzureEnvironment = $AzureEnvironment
    }
    elseif ($AzureEnvironment -ne 'Production') {
        throw "$($registrationCommand.Name) cannot target the resolved PnP environment '$AzureEnvironment'. Registration was not attempted."
    }

    Write-Step -Message "Registering '$ApplicationName' in tenant $TenantId with delegated AllSites.FullControl..."
    $registrationOutput = @()
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            $registrationOutput = @(& $registrationCommand @parameters)
            break
        }
        catch {
            if ($attempt -eq 3 -or [string]$_.Exception.Message -notmatch '(?i)already exists') { throw }
            # This application is disposable and removed on exit, so a name collision just takes the next free name.
            $parameters.ApplicationName = "$ApplicationName $(Get-Date -Format $script:TimestampFormat)"
            Write-Step -Severity Warn -Message "An application named '$ApplicationName' already exists, so this run registers '$($parameters.ApplicationName)' instead."
        }
    }
    $clientIdPropertyNames = @('AzureAppId/ClientId', 'ClientId', 'AppId', 'ApplicationId')
    foreach ($outputItem in $registrationOutput) {
        if ($null -eq $outputItem) { continue }
        foreach ($propertyName in $clientIdPropertyNames) {
            $property = $outputItem.PSObject.Properties[$propertyName]
            if (-not $property) { continue }
            $registeredClientId = ConvertTo-GuidString -Value $property.Value
            if (-not [string]::IsNullOrWhiteSpace($registeredClientId)) { return $registeredClientId }
        }
    }

    $guidPattern = '(?i)(?<![0-9a-f])(?<value>[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12})(?![0-9a-f])'
    foreach ($outputItem in $registrationOutput) {
        foreach ($match in [regex]::Matches([string]$outputItem, $guidPattern)) {
            $registeredClientId = ConvertTo-GuidString -Value $match.Groups['value'].Value
            if (-not [string]::IsNullOrWhiteSpace($registeredClientId)) { return $registeredClientId }
        }
    }
    throw "$($registrationCommand.Name) completed but did not return a recognizable application client ID."
}

function Save-TenantPnPClientId {
    <# .SYNOPSIS Remembers a non-secret application client ID under a tenant-specific environment key. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId
    )

    $variableName = Get-TenantClientIdEnvironmentVariableName -TenantId $TenantId
    $normalizedClientId = ConvertTo-GuidString -Value $ClientId
    if ([string]::IsNullOrWhiteSpace($normalizedClientId)) { throw 'Cannot save an invalid application client ID.' }

    [Environment]::SetEnvironmentVariable($variableName, $normalizedClientId, [EnvironmentVariableTarget]::Process)
    try {
        [Environment]::SetEnvironmentVariable($variableName, $normalizedClientId, [EnvironmentVariableTarget]::User)
        Write-Step -Severity Good -Message "Saved the non-secret client ID in user environment variable $variableName."
    }
    catch {
        Write-Step -Severity Warn -Message "Could not persist $variableName for later sessions: $($_.Exception.Message)"
    }
}

function Clear-TenantPnPClientId {
    <# .SYNOPSIS Forgets a remembered client ID so later runs never reference a removed application. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$TenantId)

    $variableName = Get-TenantClientIdEnvironmentVariableName -TenantId $TenantId
    [Environment]::SetEnvironmentVariable($variableName, $null, [EnvironmentVariableTarget]::Process)
    try {
        [Environment]::SetEnvironmentVariable($variableName, $null, [EnvironmentVariableTarget]::User)
        Write-Step -Message "Cleared $variableName so later runs do not reuse the removed application."
    }
    catch {
        Write-Step -Severity Warn -Message "Could not clear $variableName. Remove it manually: $($_.Exception.Message)"
    }
}

function Clear-LegacyTenantContextPreference {
    <# .SYNOPSIS Removes cross-tenant defaults persisted by earlier versions of either script. #>
    [CmdletBinding()]
    param()

    $removed = [System.Collections.Generic.List[string]]::new()
    foreach ($name in 'LABEL_TEST_SITE_TENANT_URL', 'PURVIEW_FILE_LABELING_SITE_URL', 'PURVIEW_FILE_LABELING_LIBRARY') {
        $targets = if ($name -eq 'LABEL_TEST_SITE_TENANT_URL') {
            @([EnvironmentVariableTarget]::Process, [EnvironmentVariableTarget]::User)
        }
        else {
            @([EnvironmentVariableTarget]::User)
        }
        foreach ($target in $targets) {
            try {
                $value = [Environment]::GetEnvironmentVariable($name, $target)
                if ([string]::IsNullOrWhiteSpace($value)) { continue }
                [Environment]::SetEnvironmentVariable($name, $null, $target)
                $removed.Add("$name ($target)")
            }
            catch { Write-Verbose "Could not clear ${name} from ${target}: $($_.Exception.Message)" }
        }
    }
    if ($removed.Count -gt 0) {
        Write-Step -Message "Removed persisted cross-tenant defaults: $($removed -join ', ')."
    }
}

function Resolve-PnPClientForTenant {
    <# .SYNOPSIS Resolves, validates, or registers the PnP public client used for this tenant. #>
    [CmdletBinding()]
    param(
        [string]$ExplicitClientId = '',
        [Parameter(Mandatory)][object]$TenantMetadata,
        [Parameter(Mandatory)][string]$SharePointRootUrl,
        [Parameter(Mandatory)][string]$ApplicationName,
        [ValidateSet('Interactive', 'DeviceCode')][string]$AuthenticationMode = 'Interactive',
        [switch]$RegisterIfNeeded
    )

    $isExplicitClientId = -not [string]::IsNullOrWhiteSpace($ExplicitClientId)
    foreach ($candidate in @(Get-PnPClientIdCandidate -ExplicitClientId $ExplicitClientId -TenantId $TenantMetadata.TenantId)) {
        Write-Step -Message "Checking application $candidate in tenant $($TenantMetadata.TenantId)..."
        $applicationCheck = Test-PnPClientApplication -ClientId $candidate -TenantMetadata $TenantMetadata -SharePointRootUrl $SharePointRootUrl
        if ($applicationCheck.Status -eq 'Ready') {
            Write-Step -Severity Good -Message $applicationCheck.Message
            return $candidate
        }
        if ($applicationCheck.Status -eq 'Error') {
            throw "The Entra application preflight could not complete. $($applicationCheck.Message)"
        }
        if ($isExplicitClientId) {
            throw "Application $candidate supplied with -ClientId is $($applicationCheck.Status.ToLowerInvariant()) in tenant $($TenantMetadata.TenantId). $($applicationCheck.Message)"
        }
        Write-Step -Severity Warn -Message "Application $candidate is $($applicationCheck.Status.ToLowerInvariant()) in this tenant. $($applicationCheck.Message)"
    }

    if (-not $RegisterIfNeeded) {
        if (-not (Test-InteractiveHost)) {
            throw "No usable Entra application was found in tenant $($TenantMetadata.TenantId) and this host cannot prompt. Re-run with -RegisterApp to register one, or with -ClientId to use an existing application."
        }
        Write-Host ''
        Write-Step -Severity Warn -Message "No usable Entra application was found in tenant $($TenantMetadata.TenantId)."
        Write-Step -Message "Registering '$ApplicationName' requests only delegated SharePoint AllSites.FullControl."
        Write-Step -Message 'Answer no to stop, then re-run with -ClientId to use an existing application.'
        $answer = (Read-Host '  Register this application now? [Y/n]').Trim()
        if (-not [string]::IsNullOrWhiteSpace($answer) -and $answer -notmatch '^(?i)y(es)?$') {
            throw 'Stopped before authentication because no validated tenant application was available.'
        }
    }

    $registeredClientId = Register-PnPInteractiveApplication -TenantId $TenantMetadata.TenantId -ApplicationName $ApplicationName -AuthenticationMode $AuthenticationMode -AzureEnvironment $TenantMetadata.AzureEnvironment -Confirm:$false
    if ([string]::IsNullOrWhiteSpace($registeredClientId)) { throw 'Entra application registration was cancelled.' }
    $script:CreatedApplication = [pscustomobject]@{
        ClientId = $registeredClientId
        DisplayName = $ApplicationName
        TenantId = $TenantMetadata.TenantId
        AzureEnvironment = $TenantMetadata.AzureEnvironment
        AuthMode = if ($AuthenticationMode -eq 'DeviceCode') { 'DeviceCode' } else { 'Browser' }
    }
    Save-TenantPnPClientId -TenantId $TenantMetadata.TenantId -ClientId $registeredClientId

    # A new registration is not usable until Entra finishes replicating it.
    $applicationCheck = $null
    for ($attempt = 1; $attempt -le 10; $attempt++) {
        $applicationCheck = Test-PnPClientApplication -ClientId $registeredClientId -TenantMetadata $TenantMetadata -SharePointRootUrl $SharePointRootUrl
        if ($applicationCheck.Status -eq 'Ready') { break }
        if ($attempt -lt 10) {
            Write-Step -Message "Waiting for application $registeredClientId to become usable (attempt $attempt)..."
            Start-Sleep -Seconds 6
        }
    }
    if ($applicationCheck.Status -ne 'Ready') {
        throw "The new application did not become usable in tenant $($TenantMetadata.TenantId). $($applicationCheck.Message)"
    }
    Write-Step -Severity Good -Message "Registered and validated application $registeredClientId."
    return $registeredClientId
}

function ConvertTo-TenantDomainFromHost {
    <# .SYNOPSIS Maps a SharePoint host to its cloud's default tenant domain, or returns the host unchanged. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$HostName)

    $normalizedHost = $HostName.Trim().ToLowerInvariant()
    $hostMatch = [regex]::Match($normalizedHost, '^(?<alias>.+?)(?:-admin|-my)?\.sharepoint\.(?<cloud>com|us|de|cn)$')
    if (-not $hostMatch.Success) { return $normalizedHost }

    # Each sovereign cloud publishes tenants under its own default domain suffix.
    $domainSuffix = switch ($hostMatch.Groups['cloud'].Value) {
        'us' { 'onmicrosoft.us' }
        'de' { 'onmicrosoft.de' }
        'cn' { 'partner.onmschina.cn' }
        default { 'onmicrosoft.com' }
    }
    return "$($hostMatch.Groups['alias'].Value).$domainSuffix"
}

function Get-TenantHintFromUrl {
    <# .SYNOPSIS Converts a SharePoint URL into the Entra tenant domain or tenant ID expected by PnP device-code auth. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Url
    )

    try {
        $uri = [System.Uri]$Url
        $hostname = $uri.Host
        if (-not [string]::IsNullOrWhiteSpace($hostname)) {
            return ConvertTo-TenantDomainFromHost -HostName $hostname
        }
    }
    catch {
        Write-Verbose "Could not parse the host from '$Url': $($_.Exception.Message)"
    }

    $value = $Url.TrimEnd('/')
    $value = $value -replace '^https?://', ''
    if ($value.Contains('/')) { $value = $value.Split('/')[0] }
    if ([string]::IsNullOrWhiteSpace($value)) { throw "Could not determine the tenant identifier from '$Url'." }
    return ConvertTo-TenantDomainFromHost -HostName $value
}

function ConvertTo-TenantHint {
    <# .SYNOPSIS Normalizes a user-supplied value to an Entra tenant domain or GUID, not a SharePoint host. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Value
    )

    $trimmed = $Value.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) { return '' }

    $parsedTenantId = [guid]::Empty
    if ([guid]::TryParse($trimmed, [ref]$parsedTenantId) -and $parsedTenantId -ne [guid]::Empty) {
        return $parsedTenantId.ToString()
    }

    $candidate = $trimmed.ToLowerInvariant()
    if ($candidate.StartsWith('http://') -or $candidate.StartsWith('https://')) {
        return Get-TenantHintFromUrl -Url $candidate
    }

    if ($candidate -match '^https?://') {
        return Get-TenantHintFromUrl -Url $candidate
    }

    $candidate = $candidate -replace '^https?://', ''
    if ($candidate.Contains('/')) { $candidate = $candidate.Split('/')[0] }
    return ConvertTo-TenantDomainFromHost -HostName $candidate
}

function Disconnect-PnPCurrentSession {
    <# .SYNOPSIS Closes the current PnP connection without clearing unrelated persisted login caches. #>
    [CmdletBinding()]
    param()

    if (-not (Get-Command Disconnect-PnPOnline -ErrorAction SilentlyContinue)) { return }
    try {
        Disconnect-PnPOnline -ErrorAction Stop
        Write-Verbose 'Disconnected the current PnP session.'
    }
    catch {
        Write-Verbose "No active PnP session required cleanup: $($_.Exception.Message)"
    }
}

function Connect-GraphForApplicationCleanup {
    <# .SYNOPSIS Opens a Graph sign-in scoped to application management, used only to clean up this run's app. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$AzureEnvironment,
        [ValidateSet('Interactive', 'DeviceCode')][string]$AuthenticationMode = 'Interactive'
    )

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication -ErrorAction SilentlyContinue)) {
        throw 'Microsoft.Graph.Authentication is not installed, so the application cannot be changed from here. Install it with: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    $connectCommand = Get-Command Connect-MgGraph -ErrorAction Stop

    $parameters = @{
        TenantId = $TenantId
        Scopes = @('Application.ReadWrite.All')
        ErrorAction = 'Stop'
    }
    if ($connectCommand.Parameters.ContainsKey('NoWelcome')) { $parameters.NoWelcome = $true }
    if ($AuthenticationMode -eq 'DeviceCode' -and $connectCommand.Parameters.ContainsKey('UseDeviceAuthentication')) {
        $parameters.UseDeviceAuthentication = $true
    }
    if ($connectCommand.Parameters.ContainsKey('Environment')) {
        $graphEnvironment = switch ($AzureEnvironment) {
            'USGovernmentHigh' { 'USGov' }
            'China' { 'China' }
            'Germany' { 'Germany' }
            default { 'Global' }
        }
        $parameters.Environment = $graphEnvironment
    }
    elseif ($AzureEnvironment -ne 'Production') {
        throw "Connect-MgGraph cannot target the resolved cloud '$AzureEnvironment'."
    }

    Write-Step -Message 'Signing in once more to remove the Entra application the provisioning run created.'
    $null = Connect-MgGraph @parameters
}

function Disconnect-GraphCleanupSession {
    <# .SYNOPSIS Closes the short-lived Graph session opened for application cleanup. #>
    [CmdletBinding()]
    param()

    if (-not (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue)) { return }
    try { $null = Disconnect-MgGraph -ErrorAction Stop }
    catch { Write-Verbose "No Graph cleanup session required closing: $($_.Exception.Message)" }
}

function Get-GraphDirectoryObjectId {
    <# .SYNOPSIS Returns the directory object ID of an application or its service principal, or an empty string. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('applications', 'servicePrincipals')][string]$Collection,
        [Parameter(Mandatory)][string]$ClientId
    )

    $response = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/$Collection`?`$filter=appId eq '$ClientId'" -ErrorAction Stop
    foreach ($item in @($response.value)) {
        if ($null -ne $item -and $item.ContainsKey('id')) { return [string]$item.id }
    }
    return ''
}

function Clear-EntraApplicationAccess {
    <# .SYNOPSIS Revokes consent, removes requested permissions, and disables sign-in for an application. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ClientId)

    $strippedActions = [System.Collections.Generic.List[string]]::new()
    $servicePrincipalId = Get-GraphDirectoryObjectId -Collection servicePrincipals -ClientId $ClientId
    if (-not [string]::IsNullOrWhiteSpace($servicePrincipalId)) {
        $grants = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/oauth2PermissionGrants?`$filter=clientId eq '$servicePrincipalId'" -ErrorAction Stop
        $revokedCount = 0
        foreach ($grant in @($grants.value)) {
            if ($null -eq $grant -or -not $grant.ContainsKey('id')) { continue }
            $null = Invoke-MgGraphRequest -Method DELETE -Uri "/v1.0/oauth2PermissionGrants/$($grant.id)" -ErrorAction Stop
            $revokedCount++
        }
        if ($revokedCount -gt 0) { $strippedActions.Add("revoked $revokedCount delegated permission grant(s)") }

        $null = Invoke-MgGraphRequest -Method PATCH -Uri "/v1.0/servicePrincipals/$servicePrincipalId" -Body @{ accountEnabled = $false } -ErrorAction Stop
        $strippedActions.Add('disabled sign-in on the service principal')
    }

    $applicationId = Get-GraphDirectoryObjectId -Collection applications -ClientId $ClientId
    if (-not [string]::IsNullOrWhiteSpace($applicationId)) {
        $body = @{
            requiredResourceAccess = @()
            publicClient = @{ redirectUris = @() }
            web = @{ redirectUris = @() }
            spa = @{ redirectUris = @() }
        }
        $null = Invoke-MgGraphRequest -Method PATCH -Uri "/v1.0/applications/$applicationId" -Body $body -ErrorAction Stop
        $strippedActions.Add('removed every requested API permission and redirect URI')
    }

    return [pscustomobject]@{
        ApplicationId = $applicationId
        ServicePrincipalId = $servicePrincipalId
        Actions = @($strippedActions)
    }
}

function Write-ManualApplicationRemovalNote {
    <# .SYNOPSIS Logs exactly what is left behind and how to remove it by hand. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Application,
        [Parameter(Mandatory)][string]$Reason,
        [Parameter(Mandatory)][bool]$AccessWasStripped
    )

    Write-Step -Severity Warn -Message "The Entra application created by this run could not be deleted. $Reason"
    Write-Step -Severity Warn -Message "Left behind: '$($Application.DisplayName)', client ID $($Application.ClientId), tenant $($Application.TenantId)."
    if ($AccessWasStripped) {
        Write-Step -Message 'Its consent, API permissions, redirect URIs, and sign-in were removed, so it can no longer be used to reach any data.'
    }
    else {
        Write-Step -Severity Warn -Message 'Its permissions could not be stripped either, so it still holds delegated SharePoint AllSites.FullControl.'
    }
    Write-Step -Message 'Remove it in the Entra admin center under App registrations, or run:'
    Write-Step -Message "  Connect-MgGraph -TenantId $($Application.TenantId) -Scopes Application.ReadWrite.All"
    Write-Step -Message "  Get-MgApplication -Filter `"appId eq '$($Application.ClientId)'`" | ForEach-Object { Remove-MgApplication -ApplicationId `$_.Id }"
}

function Invoke-ApplicationCleanupWorker {
    <# .SYNOPSIS Strips and deletes one application, then reports the outcome through a result file. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$AzureEnvironment,
        [Parameter(Mandatory)][string]$ResultPath,
        [ValidateSet('Interactive', 'DeviceCode')][string]$AuthenticationMode = 'Interactive'
    )

    $deleted = $false
    $strippedActions = @()
    $errorMessage = ''
    try {
        Connect-GraphForApplicationCleanup -TenantId $TenantId -AzureEnvironment $AzureEnvironment -AuthenticationMode $AuthenticationMode
        try {
            $stripResult = Clear-EntraApplicationAccess -ClientId $ClientId
            $strippedActions = @($stripResult.Actions)
            foreach ($action in $strippedActions) { Write-Step -Severity Good -Message "Stripped access: $action." }
            if ([string]::IsNullOrWhiteSpace($stripResult.ApplicationId)) {
                throw "Application $ClientId was not found in tenant $TenantId."
            }
            $null = Invoke-MgGraphRequest -Method DELETE -Uri "/v1.0/applications/$($stripResult.ApplicationId)" -ErrorAction Stop
            $deleted = $true
        }
        finally { Disconnect-GraphCleanupSession }
    }
    catch { $errorMessage = $_.Exception.Message }

    $payload = [pscustomobject]@{
        Deleted = $deleted
        StrippedActions = $strippedActions
        ErrorMessage = $errorMessage
    }
    try { $payload | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $ResultPath -Encoding utf8 -ErrorAction Stop }
    catch { Write-Step -Severity Error -Message "Could not write the cleanup result to '$ResultPath'. $($_.Exception.Message)" }

    if ($deleted) { return 0 }
    if ($strippedActions.Count -gt 0) { return 2 }
    return 3
}

function Get-CurrentPowerShellHostPath {
    <# .SYNOPSIS Returns an executable able to run the cleanup, preferring the host this run already uses. #>
    [CmdletBinding()]
    param()

    if ($PSVersionTable.PSVersion -ge [version]'7.2.0') {
        try {
            $currentPath = [string](Get-Process -Id $PID -ErrorAction Stop).Path
            if (-not [string]::IsNullOrWhiteSpace($currentPath) -and (Test-Path -LiteralPath $currentPath -PathType Leaf)) { return $currentPath }
        }
        catch { Write-Verbose "Could not read this host's executable path: $($_.Exception.Message)" }
    }
    return Get-PowerShell7Path -MinimumVersion ([version]'7.2.0')
}

function Remove-EntraApplicationInSeparateProcess {
    <# .SYNOPSIS Runs the Graph cleanup in a fresh process, because PnP.PowerShell and Microsoft.Graph cannot share one MSAL assembly. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Application,
        [Parameter(Mandatory)][string]$ScriptPath
    )

    if ([string]::IsNullOrWhiteSpace($ScriptPath) -or -not (Test-Path -LiteralPath $ScriptPath -PathType Leaf)) {
        throw "The cleanup could not locate this script at '$ScriptPath'."
    }
    $hostPath = Get-CurrentPowerShellHostPath
    if ([string]::IsNullOrWhiteSpace($hostPath)) {
        throw 'No PowerShell 7.2 or later host was found to run the cleanup in a separate process.'
    }

    $resultPath = Join-Path ([System.IO.Path]::GetTempPath()) ('LabelTestSiteCleanup-' + [guid]::NewGuid().ToString('N') + '.json')
    try {
        Write-Step -Message 'Running the removal in a separate process, because PnP and Microsoft Graph cannot share one identity assembly.'
        # Out-Host keeps the child's sign-in prompts visible; without it they become this function's return value.
        & $hostPath -NoLogo -NoProfile -File $ScriptPath `
            -RemoveApplicationClientId $Application.ClientId `
            -RemoveApplicationTenantId $Application.TenantId `
            -RemoveApplicationEnvironment $Application.AzureEnvironment `
            -RemoveApplicationResultPath $resultPath `
            -AuthMode $Application.AuthMode | Out-Host

        if (-not (Test-Path -LiteralPath $resultPath -PathType Leaf)) {
            throw "The cleanup process exited with code $LASTEXITCODE without reporting a result."
        }
        $payload = Get-Content -LiteralPath $resultPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        return [pscustomobject]@{
            Deleted = [bool](Get-PayloadValue -InputObject $payload -Name 'Deleted')
            StrippedActions = @((Get-PayloadValue -InputObject $payload -Name 'StrippedActions') | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
            ErrorMessage = [string](Get-PayloadValue -InputObject $payload -Name 'ErrorMessage')
        }
    }
    finally { Remove-Item -LiteralPath $resultPath -Force -ErrorAction SilentlyContinue }
}

function Get-PayloadValue {
    <# .SYNOPSIS Reads one property from a deserialized result, returning null when it is absent. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowNull()][object]$InputObject,
        [Parameter(Mandatory)][string]$Name
    )

    if ($null -eq $InputObject) { return $null }
    $property = $InputObject.PSObject.Properties[$Name]
    if (-not $property) { return $null }
    return $property.Value
}

function Remove-CreatedEntraApplication {
    <# .SYNOPSIS Deletes the app this run registered, stripping its access first so a failed delete still leaves it harmless. #>
    [CmdletBinding()]
    param(
        [AllowEmptyString()][string]$ScriptPath = '',
        [switch]$Keep
    )

    $application = $script:CreatedApplication
    if ($null -eq $application) { return }
    $script:CreatedApplication = $null

    if ($Keep) {
        Write-Step -Severity Warn -Message "-KeepApp was supplied, so '$($application.DisplayName)' (client ID $($application.ClientId)) stays in tenant $($application.TenantId)."
        return
    }

    Write-Host ''
    Write-Step -Message "Cleaning up the Entra application this run created: $($application.ClientId)"
    if (Get-Command Remove-PnPEntraIDApp -ErrorAction SilentlyContinue) {
        try {
            Remove-PnPEntraIDApp -Identity $application.ClientId -Force -ErrorAction Stop
            Write-Step -Severity Good -Message "Deleted application $($application.ClientId) using the existing PnP session."
            Clear-TenantPnPClientId -TenantId $application.TenantId
            return
        }
        catch {
            # The SharePoint-scoped connection has no Graph rights, so fall through to the separate cleanup process.
            Write-Verbose "Remove-PnPEntraIDApp could not delete the application: $($_.Exception.Message)"
        }
    }

    $accessWasStripped = $false
    try {
        $cleanupResult = Remove-EntraApplicationInSeparateProcess -Application $application -ScriptPath $ScriptPath
        $accessWasStripped = $cleanupResult.StrippedActions.Count -gt 0
        if ($cleanupResult.Deleted) {
            Write-Step -Severity Good -Message "Deleted application '$($application.DisplayName)' ($($application.ClientId)) from tenant $($application.TenantId)."
            Clear-TenantPnPClientId -TenantId $application.TenantId
            return
        }
        $reason = if ([string]::IsNullOrWhiteSpace($cleanupResult.ErrorMessage)) { 'The cleanup process reported no reason.' } else { $cleanupResult.ErrorMessage }
        throw $reason
    }
    catch {
        Write-ManualApplicationRemovalNote -Application $application -Reason $_.Exception.Message -AccessWasStripped $accessWasStripped
        if ($accessWasStripped) {
            # A stripped application is useless, so never let a later run pick it up again.
            Clear-TenantPnPClientId -TenantId $application.TenantId
        }
        else {
            Write-Step -Message 'The saved client ID is kept so a later run reuses this application instead of creating another one.'
        }
    }
}

function Read-SignInRecovery {
    <# .SYNOPSIS Offers a way out when sign-in fails, so a wrong application does not end the run. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$ClientId,
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Message
    )

    Write-Host ''
    if ($Message -match '(?i)AADSTS700016|was not found in the directory') {
        Write-Step -Severity Warn -Message "Application $ClientId does not exist in tenant $TenantId. It was almost certainly registered in a different tenant, so pick option 2 below."
    }
    elseif ($Message -match '(?i)passkey|work profile|AADSTS50011|AADSTS65004') {
        Write-Step -Severity Warn -Message 'The browser could not complete that credential. Device-code sign-in lets you choose Microsoft Authenticator instead.'
    }
    Write-Host '  Sign-in did not complete. Choose what to do next.' -ForegroundColor Cyan
    Write-Host '    1. Retry with device-code sign-in (lets you pick Microsoft Authenticator)'
    Write-Host '    2. Forget the remembered application and register a new one for this tenant'
    Write-Host '    3. Enter a different application (client) ID'
    Write-Host '    4. Stop'

    while ($true) {
        $answer = ([string](Read-Host '  Selection [1]')).Trim()
        if ([string]::IsNullOrWhiteSpace($answer)) { $answer = '1' }
        switch ($answer) {
            '1' { return 'DeviceCode' }
            '2' { return 'Register' }
            '3' { return 'Different' }
            '4' { return 'Stop' }
        }
        Write-Step -Severity Warn -Message 'Enter 1, 2, 3, or 4.'
    }
}

function Save-LabelingHandoff {
    <# .SYNOPSIS Hands the created site and library to a labeling utility launched by this process. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [Parameter(Mandatory)][string]$LibraryTitle,
        [AllowEmptyString()][string]$ClientId = '',
        [AllowEmptyString()][string]$TenantId = ''
    )

    $values = [ordered]@{
        PURVIEW_FILE_LABELING_SITE_URL = $SiteUrl
        PURVIEW_FILE_LABELING_LIBRARY = $LibraryTitle
    }
    # Invoke-PurviewFileLabeling.ps1 looks the application up per tenant, so the name has to match the form it reads.
    $parsedTenant = [guid]::Empty
    if (-not [string]::IsNullOrWhiteSpace($ClientId) -and [guid]::TryParse($TenantId, [ref]$parsedTenant) -and $parsedTenant -ne [guid]::Empty) {
        $values['PURVIEW_FILE_LABELING_CLIENT_ID_' + $parsedTenant.ToString('N').ToUpperInvariant()] = $ClientId
    }
    foreach ($name in $values.Keys) {
        [Environment]::SetEnvironmentVariable($name, $values[$name], [EnvironmentVariableTarget]::Process)
    }
    [Environment]::SetEnvironmentVariable('PURVIEW_FILE_LABELING_PROCESS_HANDOFF', '1', [EnvironmentVariableTarget]::Process)
    if ($values.Count -gt 2) {
        Write-Step -Message 'Prepared a process-only handoff of this site, library, and application if Invoke-PurviewFileLabeling.ps1 starts now.'
    }
    else {
        Write-Step -Message 'Prepared a process-only handoff of this site and library if Invoke-PurviewFileLabeling.ps1 starts now.'
    }
}

function Connect-PnPOnlineChecked {
    <# .SYNOPSIS Connects to SharePoint using the validated tenant and tenant-owned Entra application. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Url,
        [string]$ClientId = '',
        [ValidateSet('Interactive', 'DeviceCode')][string]$AuthenticationMode = 'Interactive',
        [string]$Tenant = '',
        [Parameter(Mandatory)][string]$AzureEnvironment
    )

    $parameters = @{
        Url = $Url
        ClientId = $ClientId
        ErrorAction = 'Stop'
    }

    if ([string]::IsNullOrWhiteSpace($ClientId)) {
        throw 'A tenant-owned Entra application client ID is required before PnP authentication can start.'
    }

    Disconnect-PnPCurrentSession

    if ($AuthenticationMode -eq 'DeviceCode') {
        if ([string]::IsNullOrWhiteSpace($Tenant)) { throw 'Device-code authentication requires the resolved tenant ID.' }
        $parameters.DeviceLogin = $true
        $parameters.Tenant = $Tenant
        Write-Step -Message 'Using device-code sign-in with the tenant-owned Entra app client ID.'
    }
    else {
        $parameters.Interactive = $true
        # Without an explicit tenant, MSAL reuses a cached account from whichever tenant was signed in last.
        if (-not [string]::IsNullOrWhiteSpace($Tenant)) { $parameters.Tenant = $Tenant }
        Write-Step -Message 'Using interactive sign-in with the tenant-owned Entra app client ID.'
        Write-Host ''
        Write-Host '  Opening a browser to sign in. This waits here until you finish there, so' -ForegroundColor Cyan
        Write-Host '  check for a window behind this one if nothing seems to happen. If it never' -ForegroundColor Cyan
        Write-Host '  appears, press Ctrl+C and start again with -AuthMode DeviceCode.' -ForegroundColor Gray
        Write-Host ''
    }

    $connectCommand = Get-Command Connect-PnPOnline -ErrorAction Stop
    if ($connectCommand.Parameters.ContainsKey('AzureEnvironment')) {
        $parameters.AzureEnvironment = $AzureEnvironment
    }
    elseif ($AzureEnvironment -ne 'Production') {
        throw "Connect-PnPOnline cannot target the resolved PnP environment '$AzureEnvironment'."
    }
    if ($connectCommand.Parameters.ContainsKey('ValidateConnection')) { $parameters.ValidateConnection = $true }
    Connect-PnPOnline @parameters
}

function Test-PnPAdminConnection {
    <# .SYNOPSIS Confirms that PnP connected to the expected admin center with the expected app and tenant access. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ExpectedAdminUrl,
        [Parameter(Mandatory)][string]$ExpectedTenantId,
        [Parameter(Mandatory)][string]$ExpectedClientId
    )

    $connection = Get-PnPConnection -ErrorAction Stop
    if ($null -eq $connection) { throw 'Connect-PnPOnline returned without creating an active connection.' }

    $expectedHost = ([uri]$ExpectedAdminUrl).Host
    $connectionUrlProperty = $connection.PSObject.Properties['Url']
    if ($connectionUrlProperty -and -not [string]::IsNullOrWhiteSpace([string]$connectionUrlProperty.Value)) {
        $actualHost = ([uri][string]$connectionUrlProperty.Value).Host
        if ($actualHost -ne $expectedHost) {
            throw "PnP connected to '$actualHost' instead of expected admin host '$expectedHost'."
        }
    }

    $connectionClientProperty = $connection.PSObject.Properties['ClientId']
    if ($connectionClientProperty) {
        $actualClientId = ConvertTo-GuidString -Value $connectionClientProperty.Value
        if (-not [string]::IsNullOrWhiteSpace($actualClientId) -and $actualClientId -ne $ExpectedClientId) {
            throw "PnP used application $actualClientId instead of validated application $ExpectedClientId."
        }
    }

    $web = Get-PnPWeb -ErrorAction Stop
    $webUrlProperty = $web.PSObject.Properties['Url']
    if ($webUrlProperty -and -not [string]::IsNullOrWhiteSpace([string]$webUrlProperty.Value)) {
        $webHost = ([uri][string]$webUrlProperty.Value).Host
        if ($webHost -ne $expectedHost) { throw "Authenticated web '$webHost' does not match expected admin host '$expectedHost'." }
    }

    $tenant = Get-PnPTenant -ErrorAction Stop
    $tenantIdProperty = $tenant.PSObject.Properties['TenantId']
    if ($tenantIdProperty) {
        $actualTenantId = ConvertTo-GuidString -Value $tenantIdProperty.Value
        if (-not [string]::IsNullOrWhiteSpace($actualTenantId) -and $actualTenantId -ne $ExpectedTenantId) {
            throw "Authenticated tenant $actualTenantId does not match resolved tenant $ExpectedTenantId."
        }
    }

    Write-Step -Severity Good -Message 'Authenticated SharePoint admin-center and tenant access checks passed.'
    return $tenant
}

# ZipArchive lives in System.IO.Compression, which Windows PowerShell does not load with its FileSystem companion.
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# This mode is entered only by the cleanup child process, which must never load PnP.PowerShell.
if (-not [string]::IsNullOrWhiteSpace($RemoveApplicationClientId)) {
    foreach ($requiredValue in $RemoveApplicationTenantId, $RemoveApplicationEnvironment, $RemoveApplicationResultPath) {
        if ([string]::IsNullOrWhiteSpace($requiredValue)) {
            throw '-RemoveApplicationClientId also requires -RemoveApplicationTenantId, -RemoveApplicationEnvironment, and -RemoveApplicationResultPath.'
        }
    }
    $cleanupAuthenticationMode = if ($AuthMode -eq 'DeviceCode') { 'DeviceCode' } else { 'Interactive' }
    $cleanupExitCode = @(Invoke-ApplicationCleanupWorker -ClientId $RemoveApplicationClientId -TenantId $RemoveApplicationTenantId -AzureEnvironment $RemoveApplicationEnvironment -ResultPath $RemoveApplicationResultPath -AuthenticationMode $cleanupAuthenticationMode) | Select-Object -Last 1
    exit ([int]$cleanupExitCode)
}

# PnP.PowerShell can only register an Entra app on PowerShell 7.2 or later.
if (-not $NoRelaunch -and $PSVersionTable.PSVersion -lt [version]'7.2.0') {
    $powerShell7Path = Get-PowerShell7Path -MinimumVersion ([version]'7.2.0')
    if (-not [string]::IsNullOrWhiteSpace($powerShell7Path)) {
        Write-Host ''
        Write-Step -Message "Windows PowerShell cannot register an Entra app, so this run restarts in $powerShell7Path."
        $relaunchArguments = Get-RelaunchArgumentList -ScriptPath $PSCommandPath -BoundParameters $PSBoundParameters
        & $powerShell7Path @relaunchArguments
        exit $(if ($null -eq $LASTEXITCODE) { 0 } else { [int]$LASTEXITCODE })
    }
    Write-Host ''
    Write-Step -Severity Warn -Message 'PowerShell 7 was not found, so automatic Entra app registration is unavailable in this host.'
    Write-Step -Message 'Install it with: winget install --id Microsoft.PowerShell --source winget'
    Write-Step -Message 'Or pass an existing tenant application with -ClientId.'
}

Write-Host ''
Write-Host '  SharePoint Online label test-data provisioner' -ForegroundColor Cyan
Write-Host '  Creates a new site, a document library, nested folders, and labelable files.' -ForegroundColor Gray
Write-Host '  Some files hold fabricated personal and payment data so classifiers have something to detect.' -ForegroundColor Gray
Write-Host '  This helper is optional and only for validation. It is not required for normal file labeling.' -ForegroundColor Gray
Write-Host ''

$defaultLogFolder = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { $PSScriptRoot } else { (Get-Location).Path }
$resolvedLogFolder = if (-not [string]::IsNullOrWhiteSpace($LogFolder)) {
    $LogFolder.Trim()
}
else {
    Read-Setting -Prompt 'Folder for the run log' -Default $defaultLogFolder -UseDefault:$AcceptDefaults
}
$null = Start-RunLog -Folder $resolvedLogFolder
Add-LogEntry -Message "Host: PowerShell $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition)); user: $env:USERNAME; machine: $env:COMPUTERNAME"
Clear-LegacyTenantContextPreference

try {
    if (-not (Test-PnPPrerequisite)) { throw 'PnP.PowerShell prerequisite validation failed.' }

    Write-Host ''
    $candidateInput = if (-not [string]::IsNullOrWhiteSpace($TenantRootUrl)) {
        $TenantRootUrl
    }
    else {
        ''
    }

    $resolvedUrls = $null
    while ($null -eq $resolvedUrls) {
        if ([string]::IsNullOrWhiteSpace($candidateInput)) {
            if (-not (Test-InteractiveHost)) {
                throw 'The SharePoint tenant is unknown and this host cannot prompt. Pass -TenantRootUrl explicitly.'
            }
            Write-Step -Message 'Enter the tenant once. Everything else is resolved and verified automatically.'
            $candidateInput = Read-RequiredValue -Prompt 'SharePoint URL, tenant domain, or alias (for example contoso.sharepoint.com, contoso.onmicrosoft.com, or contoso)'
        }
        $candidateUrls = $null
        try { $candidateUrls = Resolve-SharePointRootUrl -Url $candidateInput }
        catch {
            if (-not (Test-InteractiveHost)) { throw }
            Write-Step -Severity Error -Message $_.Exception.Message
        }

        if ($null -ne $candidateUrls -and (Test-SharePointHostReachable -HostName $candidateUrls.Host -SuppliedValue $candidateInput)) {
            if ($candidateInput.Trim().TrimEnd('/') -ne $candidateUrls.RootUrl) {
                Write-Step -Message "Normalized the tenant to $($candidateUrls.RootUrl)"
            }
            $resolvedUrls = $candidateUrls
            break
        }

        if (-not (Test-InteractiveHost)) { throw "Could not use '$candidateInput' as a SharePoint tenant address." }
        $candidateInput = ''
    }

    $tenantRootUrl = $resolvedUrls.RootUrl
    $tenantAdminUrl = $resolvedUrls.AdminUrl

    $preferredTenant = if ([string]::IsNullOrWhiteSpace($TenantHint)) { '' } else { $TenantHint.Trim() }
    Write-Step -Message "Resolving the Entra tenant behind $tenantRootUrl..."
    $tenantMetadata = Resolve-EntraTenantDetail -RootUrl $tenantRootUrl -PreferredTenant $preferredTenant
    $tenantDomain = if ([string]::IsNullOrWhiteSpace((ConvertTo-GuidString -Value $tenantMetadata.TenantHint))) {
        $tenantMetadata.TenantHint
    }
    else {
        ''
    }
    Show-DetectedValuesBlock -TenantRootUrl $tenantRootUrl -TenantAdminUrl $tenantAdminUrl -RecommendedTenantHint $tenantMetadata.TenantId -TenantId $tenantMetadata.TenantId -TenantDomain $tenantDomain
    Write-Step -Severity Good -Message "SharePoint and OpenID metadata agree on tenant $($tenantMetadata.TenantId)."

    $authenticationMode = if ($AuthMode -eq 'DeviceCode') { 'DeviceCode' } else { 'Interactive' }
    if ([string]::IsNullOrWhiteSpace($ApplicationName)) { throw 'ApplicationName cannot be empty.' }

    # A wrong or foreign application must be recoverable, so sign-in retries with a different choice rather than ending the run.
    $clientId = ''
    $explicitClientId = $ClientId
    $registerIfNeeded = [bool]$RegisterApp
    $signedIn = $false
    while (-not $signedIn) {
        try {
            $clientId = Resolve-PnPClientForTenant -ExplicitClientId $explicitClientId -TenantMetadata $tenantMetadata -SharePointRootUrl $tenantRootUrl -ApplicationName $ApplicationName.Trim() -AuthenticationMode $authenticationMode -RegisterIfNeeded:$registerIfNeeded
            Write-Step -Message "Signing in to SharePoint Online via $tenantAdminUrl ..."
            Connect-PnPOnlineChecked -Url $tenantAdminUrl -ClientId $clientId -AuthenticationMode $authenticationMode -Tenant $tenantMetadata.TenantId -AzureEnvironment $tenantMetadata.AzureEnvironment
            $null = Test-PnPAdminConnection -ExpectedAdminUrl $tenantAdminUrl -ExpectedTenantId $tenantMetadata.TenantId -ExpectedClientId $clientId
            $signedIn = $true
        }
        catch {
            $failure = $_.Exception.Message
            Write-Step -Severity Error -Message $failure
            if (-not (Test-InteractiveHost)) { throw }
            Disconnect-PnPCurrentSession
            switch (Read-SignInRecovery -ClientId $clientId -TenantId $tenantMetadata.TenantId -Message $failure) {
                'DeviceCode' {
                    $authenticationMode = 'DeviceCode'
                    Write-Step -Message 'Switching to device-code sign-in for the rest of this run.'
                }
                'Register' {
                    Clear-TenantPnPClientId -TenantId $tenantMetadata.TenantId
                    $explicitClientId = ''
                    $registerIfNeeded = $true
                }
                'Different' {
                    $explicitClientId = Read-RequiredValue -Prompt 'Application (client) ID to use instead'
                    $registerIfNeeded = $false
                }
                default { throw 'Stopped at the sign-in step at your request.' }
            }
        }
    }
    if ($PreflightOnly) {
        Write-Host ''
        Write-Host '  Preflight complete' -ForegroundColor Cyan
        Write-Step -Severity Good -Message "Tenant       : $($tenantMetadata.TenantId)"
        Write-Step -Severity Good -Message "Client ID    : $clientId"
        Write-Step -Severity Good -Message "Admin center : $tenantAdminUrl"
        Write-Step -Message 'No SharePoint test site was created.'
        if ($null -ne $script:CreatedApplication -and -not $KeepApp) {
            Write-Step -Message 'The application registered for this preflight is removed on exit. Add -KeepApp to reuse it in the next run.'
        }
        return
    }

    $defaultSiteName = if (-not [string]::IsNullOrWhiteSpace($SiteName)) {
        $SiteName.Trim()
    }
    else {
        "Label Test $(Get-Date -Format $script:TimestampFormat)"
    }
    $defaultLibraryName = if (-not [string]::IsNullOrWhiteSpace($LibraryName)) { $LibraryName.Trim() } else { $script:LibraryTitle }

    Write-Host ''
    Write-Host '  Proposed values. Press Enter to accept each one, or type a replacement.' -ForegroundColor Cyan
    $siteName = Read-Setting -Prompt 'Site name' -Default $defaultSiteName -UseDefault:$AcceptDefaults
    if ([string]::IsNullOrWhiteSpace($siteName)) { throw 'A site name is required.' }
    $script:LibraryTitle = Read-Setting -Prompt 'Document library name' -Default $defaultLibraryName -UseDefault:$AcceptDefaults
    if ([string]::IsNullOrWhiteSpace($script:LibraryTitle)) { throw 'A document library name is required.' }
    $script:LibraryUrl = ConvertTo-UrlSlug -Name $script:LibraryTitle

    Write-Host ''
    Write-Host '  How much test content to create. The proposed amounts are enough to try labeling.' -ForegroundColor Cyan
    Write-Host '  Every extra level multiplies the folder count by the number of subfolders, so the totals grow quickly.' -ForegroundColor Gray
    Write-Host '  File counts are drawn per folder from the range you give, so the library is not uniform.' -ForegroundColor Gray
    $topLevelFolders = $TopLevelFolders
    $folderDepth = $FolderDepth
    $subfoldersPerFolder = $SubfoldersPerFolder
    $fileRangeText = $FilesPerFolder
    $sensitiveRangeText = $SensitiveFilesPerFolder
    $filePlan = @()
    $plannedFileCount = 0
    $plannedSensitiveCount = 0
    while ($true) {
        # Each answer becomes the next proposal, so an oversized shape only needs the one number changed.
        $topLevelFolders = Read-CountSetting -Prompt 'Top-level folders' -Default $topLevelFolders -Minimum 0 -Maximum 10 -UseDefault:$AcceptDefaults
        $folderDepth = Read-CountSetting -Prompt 'Folder levels, counting the top one' -Default $folderDepth -Minimum 1 -Maximum 5 -UseDefault:$AcceptDefaults
        $subfoldersPerFolder = Read-CountSetting -Prompt 'Subfolders inside each folder' -Default $subfoldersPerFolder -Minimum 0 -Maximum 3 -UseDefault:$AcceptDefaults
        $fileRange = Read-RangeSetting -Prompt 'Files in each folder, including the library root' -Default $fileRangeText -Minimum 0 -Maximum 10 -UseDefault:$AcceptDefaults
        $sensitiveRange = Read-RangeSetting -Prompt 'Of those, files holding fabricated sensitive data' -Default $sensitiveRangeText -Minimum 0 -Maximum 10 -UseDefault:$AcceptDefaults
        $fileRangeText = '{0}-{1}' -f $fileRange.Minimum, $fileRange.Maximum
        $sensitiveRangeText = '{0}-{1}' -f $sensitiveRange.Minimum, $sensitiveRange.Maximum

        # Wrapped in @() so that asking for no folders still yields an empty array rather than nothing.
        $script:FolderTree = @(New-TestFolderTree -TopLevelCount $topLevelFolders -Depth $folderDepth -SubfolderCount $subfoldersPerFolder)
        # The library root holds files too, so it counts as one more folder.
        $filePlan = @(New-TestFilePlan -FolderCount ($script:FolderTree.Count + 1) -FileRange $fileRange -SensitiveRange $sensitiveRange)
        $plannedFileCount = ($filePlan | Measure-Object -Property FileCount -Sum).Sum
        $plannedSensitiveCount = ($filePlan | Measure-Object -Property SensitiveCount -Sum).Sum
        if ($script:FolderTree.Count -le $script:MaximumTestFolders -and $plannedFileCount -le $script:MaximumTestFiles) { break }

        $refusal = "$topLevelFolders top-level folders, $folderDepth levels and $subfoldersPerFolder subfolders each come to $($script:FolderTree.Count) folders, which at $fileRangeText files per folder is $plannedFileCount files. That is above the limit of $script:MaximumTestFolders folders and $script:MaximumTestFiles files for disposable test data."
        if ($AcceptDefaults -or -not (Test-InteractiveHost)) { throw "$refusal Choose smaller amounts." }
        Write-Step -Severity Warn -Message "$refusal Lowering the levels or the subfolders per folder shrinks it fastest."
        Write-Host ''
    }

    Write-Host ''
    Write-Step -Message "Library         : $script:LibraryTitle"
    Write-Step -Message "Test content    : $($script:FolderTree.Count) folders, $plannedFileCount files, $plannedSensitiveCount of them with fabricated sensitive data"
    Write-Host ''

    $siteUrl = ''
    for ($nameAttempt = 1; $nameAttempt -le 5 -and [string]::IsNullOrWhiteSpace($siteUrl); $nameAttempt++) {
        $candidateName = if ($nameAttempt -eq 1) { $siteName } else { "$siteName $nameAttempt" }
        $candidateUrl = "$tenantRootUrl/sites/$(ConvertTo-UrlSlug -Name $candidateName)"
        Write-Step -Message "Creating site '$candidateName' at $candidateUrl ..."
        try {
            $null = Invoke-WithRetry -Description 'Site creation' -Operation {
                New-PnPSite -Type TeamSiteWithoutMicrosoft365Group -Title $candidateName -Url $candidateUrl -ErrorAction Stop
            }
            $siteName = $candidateName
            $siteUrl = $candidateUrl
        }
        catch {
            if ($_.Exception.Message -notmatch '(?i)already (exists|in use)|not available|is taken|in use by') { throw }
            Write-Step -Severity Warn -Message "$candidateUrl is already in use. Trying the next available name..."
        }
    }
    if ([string]::IsNullOrWhiteSpace($siteUrl)) { throw "Could not find an unused site URL under $tenantRootUrl/sites for '$siteName'." }
    Write-Step -Severity Good -Message "Site created: $siteUrl"

    # A new site collection is not always immediately connectable.
    $connected = $false
    for ($attempt = 1; $attempt -le 10 -and -not $connected; $attempt++) {
        try {
            Connect-PnPOnlineChecked -Url $siteUrl -ClientId $clientId -AuthenticationMode $authenticationMode -Tenant $tenantMetadata.TenantId -AzureEnvironment $tenantMetadata.AzureEnvironment
            $null = Get-PnPWeb -ErrorAction Stop
            $connected = $true
        }
        catch {
            Write-Step -Message "Waiting for the new site to finish provisioning (attempt $attempt)..."
            Start-Sleep -Seconds 15
        }
    }
    if (-not $connected) { throw "The new site did not become available: $siteUrl" }
    Write-Step -Severity Good -Message 'Connected to the new site.'

    Write-Step -Message "Creating document library '$script:LibraryTitle'..."
    $null = Invoke-WithRetry -Description 'Document library creation' -Operation {
        New-PnPList -Title $script:LibraryTitle -Url $script:LibraryUrl -Template DocumentLibrary -OnQuickLaunch -ErrorAction Stop
    }
    Write-Step -Severity Good -Message 'Library created.'

    foreach ($relativePath in $script:FolderTree) {
        $segments = $relativePath.Split('/')
        $parentRelative = if ($segments.Count -gt 1) { ($segments[0..($segments.Count - 2)]) -join '/' } else { '' }
        $parent = if ($parentRelative) { "$script:LibraryUrl/$parentRelative" } else { $script:LibraryUrl }
        $folderName = $segments[-1]
        $null = Invoke-WithRetry -Description "Folder creation for $relativePath" -Operation {
            Add-PnPFolder -Name $folderName -Folder $parent -ErrorAction Stop
        }
        Write-Step -Message "Created folder: $relativePath"
    }

    $stagingRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("LabelTestFiles-" + [guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $stagingRoot -Force
    $uploadedCount = 0
    $sensitiveUploadedCount = 0
    $formatCounts = [ordered]@{ docx = 0; xlsx = 0 }

    try {
        $targets = @('') + $script:FolderTree
        $formats = @('docx', 'xlsx')
        for ($folderIndex = 0; $folderIndex -lt $targets.Count; $folderIndex++) {
            $relativePath = $targets[$folderIndex]
            $folderFileCount = $filePlan[$folderIndex].FileCount
            $folderSensitiveCount = $filePlan[$folderIndex].SensitiveCount
            $folderPath = if ([string]::IsNullOrEmpty($relativePath)) { $script:LibraryUrl } else { "$script:LibraryUrl/$relativePath" }
            $label = if ([string]::IsNullOrEmpty($relativePath)) { 'root' } else { $relativePath }
            $namePart = ($label -replace '[^A-Za-z0-9]+', '-').Trim('-')

            for ($fileNumber = 1; $fileNumber -le $folderFileCount; $fileNumber++) {
                # Offset by folder so both formats carry sensitive data even where only one file per folder does.
                $format = $formats[($fileNumber - 1 + $folderIndex) % $formats.Count]
                $isSensitive = $fileNumber -le $folderSensitiveCount
                $namePrefix = if ($isSensitive) { 'SensitiveDoc' } else { 'TestDoc' }
                $fileName = "$namePrefix-$namePart-$fileNumber.$format"
                $localPath = Join-Path $stagingRoot $fileName
                $content = Get-TestFileContent -FolderLabel $label -FileNumber $fileNumber -IncludeSensitiveData:$isSensitive
                New-TestOfficeFile -Path $localPath -Format $format -Line $content -Confirm:$false
                $null = Invoke-WithRetry -Description "Upload of $fileName" -Operation {
                    Add-PnPFile -Path $localPath -Folder $folderPath -ErrorAction Stop
                }
                $uploadedCount++
                $formatCounts[$format]++
                if ($isSensitive) { $sensitiveUploadedCount++ }
            }
            if ($folderFileCount -gt 0) {
                Write-Step -Message "Uploaded $folderFileCount $(if ($folderFileCount -eq 1) { 'file' } else { 'files' }), $folderSensitiveCount with sensitive data, to: $label"
            }
            else {
                Write-Step -Message "Left empty: $label"
            }
        }
    }
    finally {
        Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
    }

    $treeDepth = if ($script:FolderTree.Count -gt 0) { $folderDepth } else { 0 }
    $formatSummary = if ($uploadedCount -eq 0) { '' } else { " ($($formatCounts['docx']) .docx, $($formatCounts['xlsx']) .xlsx)" }
    Write-Host ''
    Write-Host '  Test site ready' -ForegroundColor Cyan
    Write-Step -Severity Good -Message "Site      : $siteUrl"
    Write-Step -Severity Good -Message "Library   : $siteUrl/$script:LibraryUrl"
    Write-Step -Severity Good -Message "Folders   : $($script:FolderTree.Count) ($treeDepth levels deep)"
    Write-Step -Severity Good -Message "Files     : $uploadedCount$formatSummary"
    if ($sensitiveUploadedCount -gt 0) {
        Write-Step -Severity Good -Message "Sensitive : $sensitiveUploadedCount named SensitiveDoc-*, holding fabricated SSNs, card numbers, and bank details next to the keywords those classifiers look for"
    }
    Write-Host ''
    Write-Step -Message 'Every file starts with no sensitivity label, so a first labeling run reports them all as changeable. To exercise the rule that protects an existing higher label, label a few files by hand and run it again.'
    if ($sensitiveUploadedCount -gt 0) {
        Write-Step -Message 'The SensitiveDoc-* files give auto-labeling, DLP, and content explorer something to detect, while the TestDoc-* files are deliberately plain. Classification is not instant, so allow time for the service to crawl the new library.'
    }
    Write-Host ''
    # Handing the application over only helps when it still exists after this run has cleaned up.
    $applicationSurvives = ($null -eq $script:CreatedApplication) -or $KeepApp
    $handoffClientId = if ($applicationSurvives) { $clientId } else { '' }
    Save-LabelingHandoff -SiteUrl $siteUrl -LibraryTitle $script:LibraryTitle -ClientId $handoffClientId -TenantId $tenantMetadata.TenantId

    $labelingScript = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) { '' } else { Join-Path $PSScriptRoot 'Invoke-PurviewFileLabeling.ps1' }
    if (-not [string]::IsNullOrWhiteSpace($labelingScript) -and (Test-Path -LiteralPath $labelingScript -PathType Leaf)) {
        Write-Step -Message "Invoke-PurviewFileLabeling.ps1 is in this folder and can start against the new library straight away."
        if ($AcceptDefaults -or -not (Test-InteractiveHost)) {
            Write-Step -Message 'Running unattended, so it was not started. Run it yourself to label this library.'
        }
        else {
            $answer = ([string](Read-Host '  Start Invoke-PurviewFileLabeling.ps1 against this library now? [Y/n]')).Trim()
            if ([string]::IsNullOrWhiteSpace($answer) -or $answer -match '^(?i)y(es)?$') {
                # Deferred to the end of the script so the Entra cleanup in the finally block runs first.
                $script:LaunchLabelingPath = $labelingScript
                Write-Step -Message 'It will start once this run has cleaned up its own application.'
            }
        }
    }
    else {
        Write-Step -Message 'To label this library with Invoke-PurviewFileLabeling.ps1, choose the SharePoint Online source and supply:'
        Write-Step -Message "  Site URL : $siteUrl"
        Write-Step -Message "  Library  : $script:LibraryTitle"
    }
    if (-not $applicationSurvives) {
        Write-Step -Message 'That script registers its own application, because the one created here is SharePoint-only and is removed when this run ends. Add -KeepApp to hand this one over instead and save a sign-in.'
    }
    Write-Step -Message 'To label through the local/UNC/SharePoint Server file path source instead, sync this library with the OneDrive client and point the script at the local synced folder.'
    Write-Host ''
}
catch {
    Write-Host ''
    Write-Step -Severity Error -Message $_.Exception.Message
    Write-Step -Message 'No shared or embedded application identity is used. Re-run with -PreflightOnly to validate the URL, tenant, app, sign-in, and SharePoint admin access without creating a test site.'
    Write-Host ''
    throw
}
finally {
    Remove-CreatedEntraApplication -ScriptPath $PSCommandPath -Keep:$KeepApp
    Disconnect-PnPCurrentSession
    if (-not [string]::IsNullOrWhiteSpace($script:LogPath)) {
        Write-Host ''
        Write-Host "  Action log: $script:LogPath" -ForegroundColor Cyan
        Write-Host ''
    }
}

# Started here, not inside the try block, so the Entra cleanup above has already finished.
if (-not [string]::IsNullOrWhiteSpace($script:LaunchLabelingPath)) {
    Write-Banner -Title 'Starting Invoke-PurviewFileLabeling.ps1 against the new library' -Body @(
        'The site and library are available to this child run only, so press',
        'Enter at those prompts to accept them.'
    )
    Write-Host ''
    # The source is already known, so it is carried over rather than asked again.
    $labelingArguments = @('-InitialSource', 'SharePoint')
    if ($AuthMode -eq 'DeviceCode') { $labelingArguments += '-DeviceLogin' }
    & $script:LaunchLabelingPath @labelingArguments
}
