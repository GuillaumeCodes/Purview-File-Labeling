#Requires -Version 5.1

[CmdletBinding()]
# Both switches are read inside functions, which the analyser does not follow out of the param block.
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'NoRelaunch', Justification = 'Read by the PowerShell 7 handover check.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'ModuleRestarted', Justification = 'Read by the module-update restart guard.')]
param(
    [switch]$NoRelaunch,
    # Signs in with a device code instead of the embedded browser, which avoids the passkey prompt.
    [switch]$DeviceLogin,
    # Ignores every value remembered from previous runs, without deleting any of them.
    [switch]$Fresh,
    # Internal: carries the already-answered source across a restart, so it is never asked twice.
    [ValidateSet('', 'Local', 'SharePoint')]
    [string]$InitialSource = '',
    # Internal: set when this script restarted itself after updating a module, so it never loops on a failed update.
    [switch]$ModuleRestarted,
    # Internal: set only when this script restarts itself in PowerShell 7.
    [switch]$Restarted
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:UtilityVersion = '1.1.0'
# Shared by log and report names, retired module folders, and generated display names.
$script:TimestampFormat = 'yyyyMMdd-HHmmss'
# Floors, not pins: the utility probes for the commands and parameters it needs, and warns when a
# version is old enough that those probes are likely to start failing.
$script:ModuleExpectation = @(
    [pscustomobject]@{ Name = 'PnP.PowerShell'; Minimum = [version]'2.12.0'; Purpose = 'the SharePoint Online source' }
    [pscustomobject]@{ Name = 'ExchangeOnlineManagement'; Minimum = [version]'3.0.0'; Purpose = 'reading the sensitivity labels published in your tenant' }
    [pscustomobject]@{ Name = 'Microsoft.Graph.Authentication'; Minimum = [version]'2.0.0'; Purpose = 'the explicit administrator-consent fallback when Azure CLI cannot complete it' }
    [pscustomobject]@{ Name = 'Az.Accounts'; Minimum = [version]'2.12.0'; Purpose = 'linking Azure billing for metered writes' }
    [pscustomobject]@{ Name = 'Az.Resources'; Minimum = [version]'6.0.0'; Purpose = 'linking Azure billing for metered writes' }
)

$script:RunStarted = Get-Date
$script:ScriptPath = $PSCommandPath
$script:RelaunchCompleted = $false
$script:RelaunchHostPath = ''
$script:RelaunchReason = ''
$script:RelaunchIsModuleUpdate = $false
$script:LogPath = $null
$script:ReportPath = $null
$script:ComplianceSessionOpened = $false
$script:SharePointSessionOpened = $false
$script:GraphSessionOpened = $false
$script:AzureCliSessionOpened = $false
$script:AzureCliAccount = ''
$script:AzureCliIsolatedConfigDirectory = ''
$script:AzureCliPreviousConfigDirectory = $null
$script:AzurePowerShellSessionOpened = $false
$script:CachedLabels = $null
$script:EmptyInputCount = 0
$script:Results = [System.Collections.Generic.List[object]]::new()
$script:Failures = [System.Collections.Generic.List[object]]::new()
$script:RejectedClientIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
# Anything registered during this run may still be replicating, so it is never treated as stale.
$script:SessionSavedClientIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
# Everything the Purview Information Protection client can label on disk, per its supported file types.
$script:ClientFileExtensions = @(
    '.doc', '.docm', '.docx', '.dot', '.dotm', '.dotx',
    '.xls', '.xlsb', '.xlsm', '.xlsx', '.xltm', '.xltx',
    '.ppt', '.pps', '.ppsm', '.ppsx', '.potm', '.potx', '.pptm', '.pptx',
    '.vsd', '.vsdm', '.vsdx', '.vss', '.vssm', '.vssx', '.vst', '.vstm', '.vstx', '.vdw', '.pdf'
)
# SharePoint applies labels to a much shorter list than the client does, and .pdf needs the tenant PDF opt-in.
$script:SharePointFileExtensions = @('.docx', '.docm', '.xlsx', '.xlsm', '.xlsb', '.pptx', '.ppsx', '.pdf')
$script:ModernFileExtensions = @('.docx', '.xlsx', '.pptx', '.pdf')
$script:FilesSinceCollection = 0
$script:UseDeviceCode = [bool]$DeviceLogin
$script:ForceFreshSignIn = $false
$script:SetupInterrupted = $false
$script:Source = ''
$script:IgnoreRemembered = [bool]$Fresh
$script:ConnectedAppOnly = $false
$script:LastBillingCorrelationIds = @()
$script:LastBillingWasServiceFault = $false
$script:LastBillingWasProviderFault = $false
$script:LastBillingWasClientFault = $false
$script:LastBillingPreviewAttempted = $false
$script:LastBillingFailureSignature = ''
$script:LastBillingPreflightFailed = $false
$script:LastBillingResourceStatus = $null
$script:LastSharePointTenantLookupStatus = ''
# Every certificate this run generates is tracked, so nothing it created outlives its use.
$script:CreatedCertificateThumbprints = [System.Collections.Generic.List[string]]::new()
# Windows PowerShell writes a BOM for 'utf8' and PowerShell 7 does not, yet Excel needs one to read non-ASCII paths correctly.
$script:CsvEncoding = if ($PSVersionTable.PSEdition -eq 'Core') { 'utf8BOM' } else { 'utf8' }

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

function Write-RunLog {
    <# .SYNOPSIS Writes one structured record to the run log and a readable line to the console. #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')][string]$Severity,
        [Parameter(Mandatory)][string]$Action,
        [Parameter(Mandatory)][string]$Result,
        [string]$FilePath = '',
        [switch]$NoConsole
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = '[{0}] [{1}] Path="{2}" Action="{3}" Result="{4}"' -f `
        $timestamp, $Severity, $FilePath, $Action, ($Result -replace '[\r\n]+', ' ')
    $color = switch ($Severity) {
        'ERROR' { 'Red' }
        'WARN' { 'Yellow' }
        'SUCCESS' { 'Green' }
        default { 'Gray' }
    }

    if (-not $NoConsole) {
        $tag = switch ($Severity) {
            'ERROR' { 'ERROR: ' }
            'WARN' { 'Warning: ' }
            default { '' }
        }
        $text = if ([string]::IsNullOrWhiteSpace($FilePath)) { "$tag$Result" } else { "$tag$Result  ($FilePath)" }
        Write-Host "  $text" -ForegroundColor $color
    }
    if ($script:LogPath -and $PSCmdlet.ShouldProcess($script:LogPath, 'Append log entry')) {
        try {
            Add-Content -LiteralPath $script:LogPath -Value $line -Encoding utf8 -ErrorAction Stop
        }
        catch {
            Write-Host "  ERROR: could not write the log file. $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

function Read-OperatorInput {
    <# .SYNOPSIS Reads one typed line, refusing to continue once redirected input has run out. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Prompt)

    $value = Read-Host $Prompt
    if (-not [string]::IsNullOrWhiteSpace($value)) {
        $script:EmptyInputCount = 0
        return $value
    }
    # Exhausted redirected input returns empty forever, which would drive every menu at its default in a tight loop.
    if ([Console]::IsInputRedirected) {
        $script:EmptyInputCount++
        if ($script:EmptyInputCount -ge 5) {
            throw 'Standard input is exhausted. This utility is interactive, so run it in a console rather than piping input into it.'
        }
    }
    return $value
}

function Read-MenuChoice {
    <# .SYNOPSIS Displays a numbered menu and returns the selected value. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][System.Collections.IDictionary]$Options,
        [string]$Default
    )

    while ($true) {
        Write-RunLog -Severity INFO -Action 'Menu' -Result $Title -NoConsole
        Write-Host ''
        Write-Host "  $Title" -ForegroundColor Cyan
        foreach ($key in $Options.Keys) {
            Write-RunLog -Severity INFO -Action 'Menu option' -Result "[$key] $($Options[$key])" -NoConsole
            Write-Host ('    {0}. {1}' -f $key, $Options[$key])
        }
        $suffix = if ($Default) { " [$Default]" } else { '' }
        # A stray space around the digit would otherwise miss every key and silently redisplay the menu.
        $choice = ([string](Read-OperatorInput -Prompt "  Selection$suffix")).Trim()
        if ([string]::IsNullOrWhiteSpace($choice) -and $Default) { $choice = $Default }
        if ($Options.Contains($choice)) { return $choice }
        Write-RunLog -Severity WARN -Action 'Validate menu choice' -Result 'Enter one of the displayed option numbers.'
    }
}

function Read-PhaseAction {
    <# .SYNOPSIS Offers the standard action menu after a phase. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Phase)

    $choice = Read-MenuChoice -Title "$Phase complete. What next?" -Options ([ordered]@{
            '1' = 'Continue'
            '2' = 'Change a setting and retry'
            '3' = 'Skip this phase or item'
            '4' = 'Return to the main menu'
            '5' = 'Exit cleanly'
        }) -Default '1'
    return @{'1' = 'Continue'; '2' = 'Change'; '3' = 'Skip'; '4' = 'Main'; '5' = 'Exit'}[$choice]
}

function Add-RunFailure {
    <# .SYNOPSIS Records a non-terminating run failure. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$FilePath,
        [Parameter(Mandatory)][string]$Action,
        [Parameter(Mandatory)][string]$Reason
    )

    $script:Failures.Add([pscustomobject]@{
            Timestamp = Get-Date
            FilePath = $FilePath
            Action = $Action
            Reason = $Reason
        })
    Write-RunLog -Severity ERROR -FilePath $FilePath -Action $Action -Result $Reason
}

function Get-ErrorText {
    <# .SYNOPSIS Flattens an error into one string, because PnP and Graph bury the useful text in inner exceptions. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord)

    $parts = [System.Collections.Generic.List[string]]::new()
    if ($null -ne $ErrorRecord.ErrorDetails -and -not [string]::IsNullOrWhiteSpace($ErrorRecord.ErrorDetails.Message)) {
        $parts.Add([string]$ErrorRecord.ErrorDetails.Message)
    }
    $exception = $ErrorRecord.Exception
    for ($depth = 0; $null -ne $exception -and $depth -lt 8; $depth++) {
        if (-not [string]::IsNullOrWhiteSpace($exception.Message)) { $parts.Add([string]$exception.Message) }
        $exception = $exception.InnerException
    }
    return ((@($parts) | Select-Object -Unique) -join ' ')
}

function Test-TransientFailure {
    <# .SYNOPSIS Reports whether an error is worth retrying, which throttling on a large library makes common. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Message)

    return $Message -match '(?i)\b(429|502|503|504)\b|throttl|too many requests|timed out|timeout|temporarily unavailable|service unavailable|connection was closed|transport connection|please try again'
}

function Test-AzureServiceFailure {
    <# .SYNOPSIS Reports whether an ARM failure is Azure's own fault, so it is worth retrying and must not be reported as a misconfiguration. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Message)

    return $Message -match '(?i)internalservererror|internal server error|\b(500|502|503|504)\b|servicebusy|service unavailable|gatewaytimeout|temporarily unavailable|please retry|try again later'
}

function Test-AzureProviderImplementationFailure {
    <# .SYNOPSIS Identifies the Microsoft.GraphServices OpenTelemetry type-load error that request retries cannot repair. #>
    [CmdletBinding()]
    param([AllowEmptyString()][string]$Message)

    return ($Message -match '(?is)(TryCreateLogger|OpenTelemetry(?:\.Logs\.)?LoggerProviderSdk).*cannot\W*reduce\W*access')
}

function Get-FailureSignature {
    <# .SYNOPSIS Reduces an error to a comparable form, because one fault reported twice differs in correlation ID and escaping. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Message)

    $text = [regex]::Replace($Message, '\\u(?<code>[0-9a-fA-F]{4})', { param($match) [string][char][Convert]::ToInt32($match.Groups['code'].Value, 16) })
    if (Test-AzureProviderImplementationFailure -Message $text) {
        return 'microsoftgraphservicesopentelemetrytrycreateloggercannotreduceaccess'
    }
    $text = [regex]::Replace($text, '(?i)(correlationid|request-id|client-request-id|date)\s*:\s*\S*', '')
    $text = [regex]::Replace($text, '(?i)[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}', '')
    return ([regex]::Replace($text, '[^a-zA-Z]', '')).ToLowerInvariant()
}

function Get-AzureFailureDiagnostic {
    <# .SYNOPSIS Extracts Azure error codes and support identifiers without retaining credentials or response bodies. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Message)

    $codes = [System.Collections.Generic.List[string]]::new()
    foreach ($entry in [regex]::Matches($Message, '(?i)"code"\s*:\s*"(?<value>[^"]+)"')) {
        $codes.Add($entry.Groups['value'].Value)
    }

    $identifiers = [System.Collections.Generic.List[string]]::new()
    $identifierPattern = '(?i)(?<kind>tracking\s+id|correlation\s*id|request-id|client-request-id)\s*(?:is|[:=])\s*[''"]?(?<value>[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12})'
    foreach ($entry in [regex]::Matches($Message, $identifierPattern)) {
        $kind = ($entry.Groups['kind'].Value -replace '[\s-]', '').ToLowerInvariant()
        $label = switch ($kind) {
            'trackingid' { 'TrackingId' }
            'correlationid' { 'CorrelationId' }
            'clientrequestid' { 'ClientRequestId' }
            default { 'RequestId' }
        }
        $identifiers.Add("$label=$($entry.Groups['value'].Value)")
    }

    return [pscustomobject]@{
        Codes = @($codes | Select-Object -Unique)
        Identifiers = @($identifiers | Select-Object -Unique)
    }
}

function Invoke-WithTransientRetry {
    <# .SYNOPSIS Runs an operation, retrying only throttling and transport failures with exponential backoff. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$Operation,
        [Parameter(Mandatory)][string]$Description,
        [int]$MaximumAttempts = 4
    )

    $delaySeconds = 5
    for ($attempt = 1; $attempt -le $MaximumAttempts; $attempt++) {
        try { return & $Operation }
        catch {
            if ($attempt -ge $MaximumAttempts -or -not (Test-TransientFailure -Message (Get-ErrorText -ErrorRecord $_))) { throw }
            Write-RunLog -Severity WARN -Action $Description -Result "Attempt $attempt of $MaximumAttempts hit a transient failure. Retrying in $delaySeconds seconds. $($_.Exception.Message)"
            Start-Sleep -Seconds $delaySeconds
            $delaySeconds = [math]::Min($delaySeconds * 2, 60)
        }
    }
}

function Initialize-RunArtifact {
    <# .SYNOPSIS Creates timestamped log and report paths. #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
    param([Parameter(Mandatory)][string]$Folder)

    try {
        if (-not (Test-Path -LiteralPath $Folder -PathType Container -ErrorAction Stop)) {
            if ($PSCmdlet.ShouldProcess($Folder, 'Create logging folder')) {
                $null = New-Item -ItemType Directory -Path $Folder -Force -ErrorAction Stop
            }
        }
        $stamp = Get-Date -Format $script:TimestampFormat
        $script:LogPath = Join-Path $Folder "Invoke-PurviewFileLabeling-$stamp.log"
        $script:ReportPath = Join-Path $Folder "Invoke-PurviewFileLabeling-$stamp.csv"
        if ($PSCmdlet.ShouldProcess($script:LogPath, 'Create log file')) {
            $null = New-Item -ItemType File -Path $script:LogPath -Force -ErrorAction Stop
        }
        Write-RunLog -Severity INFO -Action 'Initialize run artifacts' -Result "Log: $($script:LogPath); report: $($script:ReportPath)"
        return $true
    }
    catch {
        $script:LogPath = $null
        Write-RunLog -Severity ERROR -FilePath $Folder -Action 'Initialize run artifacts' -Result $_.Exception.Message
        return $false
    }
}

function Connect-LabelDiscoveryService {
    <# .SYNOPSIS Offers an interactive Security and Compliance PowerShell connection. #>
    [CmdletBinding()]
    param()

    try {
        $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement -ErrorAction Stop |
            Sort-Object Version -Descending | Select-Object -First 1
        if (-not $module) {
            if (Request-ModuleInstall -Name 'ExchangeOnlineManagement' -Purpose 'to read the tenant sensitivity labels') {
                $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement -ErrorAction Stop |
                    Sort-Object Version -Descending | Select-Object -First 1
            }
            if (-not $module) {
                throw 'ExchangeOnlineManagement is not installed. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser'
            }
        }
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        if (-not (Get-Command Connect-IPPSSession -ErrorAction Stop)) {
            throw 'Connect-IPPSSession is unavailable after importing ExchangeOnlineManagement.'
        }

        $choice = Read-MenuChoice -Title 'Tenant label discovery requires Security and Compliance PowerShell. Open an interactive Microsoft 365 sign-in?' -Options ([ordered]@{
                '1' = 'Sign in now in the browser'
                '2' = 'Return without signing in'
            }) -Default '1'
        if ($choice -ne '1') { return $false }

        Write-Host ''
        Write-Host '  Sign in as an account that can read the tenant sensitivity labels.' -ForegroundColor Gray
        Write-Host '  Type the account here to skip the account picker, or press Enter to' -ForegroundColor Gray
        Write-Host '  let the browser ask which account to use.' -ForegroundColor Gray
        $userPrincipalName = Read-OperatorInput -Prompt '  Admin sign-in address, such as admin@contoso.onmicrosoft.com (Enter to choose in the browser)'
        $parameters = @{ErrorAction = 'Stop'}
        if (-not [string]::IsNullOrWhiteSpace($userPrincipalName)) {
            $parameters.UserPrincipalName = $userPrincipalName.Trim()
        }
        # The module's own banner is a fixed-width block that does not match this console's output.
        if ((Get-Command Connect-IPPSSession).Parameters.ContainsKey('ShowBanner')) { $parameters.ShowBanner = $false }
        Connect-IPPSSession @parameters
        $script:ComplianceSessionOpened = $true
        Write-RunLog -Severity SUCCESS -Action 'Connect label discovery service' -Result 'Connected to Security and Compliance PowerShell.'
        return $true
    }
    catch {
        Add-RunFailure -FilePath '' -Action 'Connect label discovery service' -Reason $_.Exception.Message
        Write-RunLog -Severity INFO -Action 'Label discovery guidance' -Result 'Install with: Install-Module ExchangeOnlineManagement -Scope CurrentUser. The account also needs Purview permissions to run Get-Label. Then choose re-check.'
        return $false
    }
}

function Get-TenantSensitivityLabel {
    <# .SYNOPSIS Retrieves file-capable sensitivity labels and their GUIDs from the tenant. #>
    [CmdletBinding()]
    param()

    try {
        # A second run in the same session would otherwise re-query, and can prompt again if the session lapsed.
        if ($null -ne $script:CachedLabels -and $script:ComplianceSessionOpened) {
            Write-RunLog -Severity SUCCESS -Action 'Discover tenant labels' -Result "Reusing the $($script:CachedLabels.Count) labels already discovered in this session."
            return $script:CachedLabels
        }
        if (-not (Get-Command Get-Label -ErrorAction SilentlyContinue)) {
            if (-not (Connect-LabelDiscoveryService)) { return $null }
        }
        if (-not (Get-Command Get-Label -ErrorAction Stop)) {
            throw 'Get-Label is unavailable. Verify the Security and Compliance PowerShell connection and Purview role assignments.'
        }

        $tenantLabels = @(Get-Label -ErrorAction Stop)
        if ($tenantLabels.Count -eq 0) {
            throw 'Get-Label returned no sensitivity labels.'
        }

        $normalized = [System.Collections.Generic.List[object]]::new()
        foreach ($tenantLabel in $tenantLabels) {
            $guidProperty = $tenantLabel.PSObject.Properties['Guid']
            $priorityProperty = $tenantLabel.PSObject.Properties['Priority']
            $contentTypeProperty = $tenantLabel.PSObject.Properties['ContentType']
            $displayNameProperty = $tenantLabel.PSObject.Properties['DisplayName']
            $nameProperty = $tenantLabel.PSObject.Properties['Name']
            $parentIdProperty = $tenantLabel.PSObject.Properties['ParentId']
            $disabledProperty = $tenantLabel.PSObject.Properties['Disabled']
            $guid = [guid]::Empty
            $priority = 0

            if (-not $guidProperty -or -not [guid]::TryParse([string]$guidProperty.Value, [ref]$guid) -or
                $guid -eq [guid]::Empty -or -not $priorityProperty -or
                -not [int]::TryParse([string]$priorityProperty.Value, [ref]$priority)) {
                Write-RunLog -Severity WARN -Action 'Normalize discovered label' -Result 'Skipped a tenant label with a missing GUID or priority.'
                continue
            }

            $name = if ($displayNameProperty -and -not [string]::IsNullOrWhiteSpace([string]$displayNameProperty.Value)) {
                [string]$displayNameProperty.Value
            }
            elseif ($nameProperty) { [string]$nameProperty.Value }
            else { [string]$guid }
            $normalized.Add([pscustomobject]@{
                    Id = [string]$guid
                    Name = $name
                    Priority = $priority
                    ParentId = if ($parentIdProperty) { [string]$parentIdProperty.Value } else { '' }
                    ContentType = if ($contentTypeProperty) { [string]$contentTypeProperty.Value } else { '' }
                    Disabled = $disabledProperty -and [bool]$disabledProperty.Value
                })
        }

        $parentIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($label in $normalized) {
            if (-not [string]::IsNullOrWhiteSpace($label.ParentId) -and $label.ParentId -ne [guid]::Empty.ToString()) {
                $null = $parentIds.Add($label.ParentId)
            }
        }

        $labelById = @{}
        foreach ($label in $normalized) { $labelById[$label.Id] = $label }

        # A projection rather than an in-place rename, so re-running discovery cannot prefix a parent name twice.
        $usableLabels = @($normalized | Where-Object {
                -not $_.Disabled -and $_.ContentType -match 'File' -and -not $parentIds.Contains($_.Id)
            } | ForEach-Object {
                $parent = if ([string]::IsNullOrWhiteSpace($_.ParentId)) { $null } else { $labelById[$_.ParentId] }
                [pscustomobject]@{
                    Id = $_.Id
                    Name = if ($parent) { "$($parent.Name) \ $($_.Name)" } else { $_.Name }
                    Priority = $_.Priority
                    ParentId = $_.ParentId
                    ContentType = $_.ContentType
                    Disabled = $_.Disabled
                }
            } | Sort-Object Priority, Name)
        if ($usableLabels.Count -eq 0) {
            throw 'No enabled, file-capable leaf sensitivity labels were returned. Check label scopes, publication, and Purview permissions.'
        }

        Write-RunLog -Severity SUCCESS -Action 'Discover tenant labels' -Result "Found $($usableLabels.Count) file-capable labels with GUIDs and priorities."
        $script:CachedLabels = $usableLabels
        return $usableLabels
    }
    catch {
        Add-RunFailure -FilePath '' -Action 'Discover tenant labels' -Reason $_.Exception.Message
        return $null
    }
}

function Disconnect-LabelDiscoveryService {
    <# .SYNOPSIS Closes a Security and Compliance session opened by this utility. #>
    [CmdletBinding()]
    param()

    if (-not $script:ComplianceSessionOpened) { return }
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
        $script:ComplianceSessionOpened = $false
        $script:CachedLabels = $null
        Write-RunLog -Severity INFO -Action 'Disconnect label discovery service' -Result 'Security and Compliance PowerShell session closed.'
    }
    catch {
        Add-RunFailure -FilePath '' -Action 'Disconnect label discovery service' -Reason $_.Exception.Message
    }
}

function Install-SharePointModule {
    <# .SYNOPSIS Checks and offers to install PnP.PowerShell for the SharePoint Online source. #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param()

    $moduleName = 'PnP.PowerShell'
    if (Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue) {
        Write-RunLog -Severity SUCCESS -Action 'Check SharePoint module' -Result 'PnP.PowerShell is installed.'
        return $true
    }

    Write-Host ''
    Write-Host '  The SharePoint Online source requires PnP.PowerShell to enumerate libraries,' -ForegroundColor Gray
    Write-Host '  read labels, and call Graph assignSensitivityLabel. Azure and Graph setup' -ForegroundColor Gray
    Write-Host '  modules are needed only if you choose to enable metered writes later.' -ForegroundColor Gray
    $choice = Read-MenuChoice -Title 'Install PnP.PowerShell now?' -Options ([ordered]@{
            '1' = 'Yes, install it from the PowerShell Gallery'
            '2' = 'No, I will install it myself'
        }) -Default '1'
    if ($choice -ne '1') { return $false }
    if (-not $PSCmdlet.ShouldProcess($moduleName, 'Install from PowerShell Gallery')) { return $false }

    Write-RunLog -Severity INFO -Action 'Install module' -Result 'Installing PnP.PowerShell from PowerShell Gallery. This can take a minute.'
    if (Request-ModuleInstall -Name $moduleName -Purpose 'for the SharePoint Online source' -Confirm:$false) {
        Write-RunLog -Severity SUCCESS -Action 'Install module' -Result 'PnP.PowerShell is installed.'
        return $true
    }
    Write-RunLog -Severity WARN -Action 'Install module' -Result 'PnP.PowerShell installation was declined or failed.'
    return $false
}

function Install-PurviewClient {
    <# .SYNOPSIS Installs the Purview Information Protection client, which is what actually writes a label to a file. #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param()

    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-RunLog -Severity WARN -Action 'Install prerequisite' -Result 'winget is unavailable here, so the client cannot be installed from this utility. Get it from https://www.microsoft.com/download/details.aspx?id=105945 and start the utility again.'
        return $false
    }

    Write-Host ''
    Write-Host '  The Purview Information Protection client is what writes a label to a file. It' -ForegroundColor Gray
    Write-Host '  needs no Azure subscription, no billing link, and no metered API, so labeling' -ForegroundColor Gray
    Write-Host '  local files and OneDrive-synced libraries costs nothing per file.' -ForegroundColor Gray
    $choice = Read-MenuChoice -Title 'Install it now?' -Options ([ordered]@{
            '1' = 'Yes, install it with winget (default)'
            '2' = 'No, I will install it myself'
        }) -Default '1'
    if ($choice -ne '1') { return $false }
    if (-not $PSCmdlet.ShouldProcess('Microsoft.PurviewInformationProtection', 'Install with winget')) { return $false }

    Write-RunLog -Severity INFO -Action 'Install prerequisite' -Result 'Installing with winget. This takes a few minutes, and Windows may ask for administrator approval.'
    try {
        & winget install --id Microsoft.PurviewInformationProtection --exact --accept-package-agreements --accept-source-agreements
        if ($LASTEXITCODE -ne 0) { throw "winget exited with code $LASTEXITCODE." }
    }
    catch {
        Add-RunFailure -FilePath '' -Action 'Install prerequisite' -Reason (Get-ErrorText -ErrorRecord $_)
        return $false
    }

    Write-RunLog -Severity SUCCESS -Action 'Install prerequisite' -Result 'The Purview Information Protection client is installed.'
    # Its module is installed machine-wide, and this session scanned the module paths at startup.
    return (Request-SelfRestart -Reason 'The Purview Information Protection client was installed, so the utility restarts to load it.')
}

function Test-PurviewPrerequisite {
    <# .SYNOPSIS Imports and verifies the Purview Information Protection cmdlets. #>
    [CmdletBinding()]
    param()

    try {
        $module = Get-Module -ListAvailable -Name PurviewInformationProtection -ErrorAction Stop |
            Sort-Object Version -Descending | Select-Object -First 1
        if (-not $module) {
            if (Install-PurviewClient -Confirm:$false) { return $false }
            throw 'PurviewInformationProtection is not installed. Install the Microsoft Purview Information Protection client, then run: Import-Module PurviewInformationProtection'
        }
        try {
            Import-Module PurviewInformationProtection -ErrorAction Stop
        }
        catch {
            if ($PSVersionTable.PSEdition -ne 'Core') { throw }
            # The client ships .NET Framework assemblies, so PowerShell 7 loads it through a Windows PowerShell compatibility session.
            Import-Module PurviewInformationProtection -UseWindowsPowerShell -ErrorAction Stop
        }
        foreach ($commandName in 'Get-FileStatus', 'Set-FileLabel') {
            if (-not (Get-Command $commandName -ErrorAction Stop)) {
                throw "Required command $commandName is unavailable after importing the module."
            }
        }
        Write-RunLog -Severity SUCCESS -Action 'Check prerequisite' -Result "PurviewInformationProtection $($module.Version) is ready."
        return $true
    }
    catch {
        Add-RunFailure -FilePath '' -Action 'Check prerequisite' -Reason $_.Exception.Message
        Write-RunLog -Severity INFO -Action 'Prerequisite guidance' -Result 'Install the Microsoft Purview Information Protection client/module, import it, and run Set-Authentication. Windows PowerShell 5.1 is the most reliable host for this module. Then select re-check.'
        return $false
    }
}

function Test-TargetPath {
    <# .SYNOPSIS Verifies that a local or UNC directory can be reached. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Path)

    try {
        if (-not (Test-Path -LiteralPath $Path -PathType Container -ErrorAction Stop)) {
            throw 'Path is unreachable or is not a directory.'
        }
        $null = Get-Item -LiteralPath $Path -ErrorAction Stop
        Write-RunLog -Severity SUCCESS -FilePath $Path -Action 'Validate target path' -Result 'Path is reachable.'
        return $true
    }
    catch {
        $reason = if ($_.Exception.Message -match 'denied|unauthorized') { "Access denied. $($_.Exception.Message)" } else { $_.Exception.Message }
        Add-RunFailure -FilePath $Path -Action 'Validate target path' -Reason $reason
        return $false
    }
}

function Get-ObjectPropertyValue {
    <# .SYNOPSIS Returns the first matching property value from an object, or an empty string. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowNull()][object]$InputObject,
        [Parameter(Mandatory)][string[]]$Names
    )

    if ($null -eq $InputObject) { return '' }
    # The indexer is a keyed lookup; piping the property collection rescans it for every candidate name on every file.
    $properties = $InputObject.PSObject.Properties
    foreach ($name in $Names) {
        $property = $properties[$name]
        if ($property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) { return [string]$property.Value }
    }
    return ''
}

function Get-RawObjectPropertyValue {
    <# .SYNOPSIS Returns a property without converting nested objects or arrays to text. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowNull()][object]$InputObject,
        [Parameter(Mandatory)][string[]]$Names
    )

    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Collections.IDictionary]) {
        foreach ($name in $Names) {
            if ($InputObject.Contains($name)) { return $InputObject[$name] }
        }
        return $null
    }
    foreach ($name in $Names) {
        $property = $InputObject.PSObject.Properties[$name]
        if ($property) { return $property.Value }
    }
    return $null
}

function Request-SelfRestart {
    <# .SYNOPSIS Arranges a restart in a new process, which is the only way to pick up a different version of a loaded module. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Reason)

    if ($ModuleRestarted) {
        Write-RunLog -Severity WARN -Action 'Restart utility' -Result 'This run already restarted once for a module update, so it will not do so again. Close this window and start the utility again.'
        return $false
    }
    if ([string]::IsNullOrWhiteSpace($script:ScriptPath) -or -not (Test-Path -LiteralPath $script:ScriptPath -PathType Leaf)) {
        Write-RunLog -Severity WARN -Action 'Restart utility' -Result 'This utility could not locate its own script file, so it cannot restart itself. Close this window and start it again.'
        return $false
    }

    $hostPath = ''
    try {
        $currentPath = [string](Get-Process -Id $PID -ErrorAction Stop).Path
        if (-not [string]::IsNullOrWhiteSpace($currentPath) -and (Test-Path -LiteralPath $currentPath -PathType Leaf)) { $hostPath = $currentPath }
    }
    catch { Write-Verbose "Could not read this host's executable path: $($_.Exception.Message)" }
    if ([string]::IsNullOrWhiteSpace($hostPath)) { $hostPath = Get-PowerShell7Path -MinimumVersion ([version]'7.2.0') }
    if ([string]::IsNullOrWhiteSpace($hostPath)) {
        Write-RunLog -Severity WARN -Action 'Restart utility' -Result 'No PowerShell host could be located to restart with. Close this window and start the utility again.'
        return $false
    }

    $script:RelaunchHostPath = $hostPath
    $script:RelaunchReason = $Reason
    $script:RelaunchIsModuleUpdate = $true
    $script:RelaunchCompleted = $true
    return $true
}

function Get-InstalledPnPVersion {
    <# .SYNOPSIS Returns every installed PnP.PowerShell version, which plain -ListAvailable hides by showing only the newest. #>
    [CmdletBinding()]
    param()

    try {
        $versions = [System.Collections.Generic.List[version]]::new()
        foreach ($module in @(Get-Module -ListAvailable -Name PnP.PowerShell -All -ErrorAction Stop)) {
            if ($null -eq $module.Version -or [version]$module.Version -eq [version]'0.0.0.0') { continue }
            # -All also returns folders whose name is not a version, and PowerShell never loads those.
            $leaf = Split-Path -Leaf $module.ModuleBase
            $folderVersion = [version]'0.0'
            if ($leaf -ne 'PnP.PowerShell' -and -not [version]::TryParse($leaf, [ref]$folderVersion)) { continue }
            if (-not $versions.Contains([version]$module.Version)) { $versions.Add([version]$module.Version) }
        }
        return @($versions | Sort-Object -Descending)
    }
    catch {
        Write-Verbose "Could not enumerate PnP.PowerShell versions: $($_.Exception.Message)"
        return @()
    }
}

function Get-LoadedAssemblyPath {
    <# .SYNOPSIS Returns assemblies that the current PowerShell process loaded from a module folder. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Folder)

    try {
        $trimCharacters = [char[]]@([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
        $folderPrefix = [System.IO.Path]::GetFullPath($Folder).TrimEnd($trimCharacters) + [System.IO.Path]::DirectorySeparatorChar
    }
    catch { return @() }

    $paths = [System.Collections.Generic.List[string]]::new()
    foreach ($assembly in [AppDomain]::CurrentDomain.GetAssemblies()) {
        try {
            $location = [string]$assembly.Location
            if ([string]::IsNullOrWhiteSpace($location)) { continue }
            $fullPath = [System.IO.Path]::GetFullPath($location)
            if ($fullPath.StartsWith($folderPrefix, [System.StringComparison]::OrdinalIgnoreCase)) { $paths.Add($fullPath) }
        }
        catch { Write-Verbose "Could not inspect one loaded assembly: $($_.Exception.Message)" }
    }
    return @($paths | Sort-Object -Unique)
}

function Show-PnPHoldingProcess {
    <# .SYNOPSIS Names the PowerShell processes holding a module folder open, because a loaded assembly cannot be deleted. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Folder)

    $currentAssemblies = @(Get-LoadedAssemblyPath -Folder $Folder)
    if ($currentAssemblies.Count -gt 0) {
        Write-Host ''
        Write-Host '  This PowerShell process has already loaded files from that version:' -ForegroundColor Yellow
        foreach ($path in @($currentAssemblies | Select-Object -First 5)) { Write-Host "    $path" -ForegroundColor DarkGray }
        if ($currentAssemblies.Count -gt 5) { Write-Host "    ... and $($currentAssemblies.Count - 5) more" -ForegroundColor DarkGray }
        Write-Host '  PowerShell cannot unload those assemblies. Rename the version folder now, then' -ForegroundColor Gray
        Write-Host '  close this PowerShell process before using the newer PnP version.' -ForegroundColor Gray
    }

    $candidates = [System.Collections.Generic.List[object]]::new()
    foreach ($process in @(Get-Process -Name 'pwsh', 'powershell', 'powershell_ise' -ErrorAction SilentlyContinue)) {
        if ($process.Id -eq $PID) { continue }
        try {
            if (@($process.Modules | Where-Object { [string]$_.FileName -like "$Folder*" }).Count -gt 0) { $candidates.Add($process) }
        }
        catch {
            # A process running elevated or as another user cannot be inspected, so it stays a suspect.
            $candidates.Add($process)
        }
    }
    if ($candidates.Count -eq 0 -and $currentAssemblies.Count -eq 0) {
        Write-Host ''
        Write-Host '  No PowerShell process could be confirmed as holding that DLL. Access denied can' -ForegroundColor Gray
        Write-Host '  also come from folder permissions, antivirus, or a process Windows will not let' -ForegroundColor Gray
        Write-Host '  this session inspect. Try renaming the version folder first.' -ForegroundColor Gray
        return
    }
    if ($candidates.Count -gt 0) {
        Write-Host ''
        Write-Host '  These PowerShell processes may also be holding that module open:' -ForegroundColor Yellow
        foreach ($process in $candidates) {
            $title = if ([string]::IsNullOrWhiteSpace($process.MainWindowTitle)) { '(no window - possibly an editor terminal)' } else { $process.MainWindowTitle }
            Write-Host ('    PID {0,-8} {1,-16} {2}' -f $process.Id, $process.ProcessName, $title)
        }
        Write-Host '  Close a confirmed holder before deleting the folder, or rename the folder now.' -ForegroundColor Gray
    }
}

function Disable-ModuleManifest {
    <# .SYNOPSIS Renames the manifests in a module version folder, which is what module discovery reads. #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory)][string]$Folder,
        [Parameter(Mandatory)][string]$Suffix
    )

    $renamed = [System.Collections.Generic.List[string]]::new()
    $candidates = @(Get-ChildItem -LiteralPath $Folder -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -in '.psd1', '.psm1' })
    foreach ($manifest in $candidates) {
        $target = '{0}.{1}' -f $manifest.Name, $Suffix
        if (-not $PSCmdlet.ShouldProcess($manifest.FullName, "Rename to $target")) { continue }
        try {
            Rename-Item -LiteralPath $manifest.FullName -NewName $target -ErrorAction Stop
            $renamed.Add($manifest.FullName)
        }
        catch { Write-Verbose "Could not rename $($manifest.FullName): $($_.Exception.Message)" }
    }
    return @($renamed)
}

function Disable-ModuleVersionFolder {
    <# .SYNOPSIS Stops PowerShell discovering a module version, without deleting any of its files. #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param([Parameter(Mandatory)][string]$Folder)

    $suffix = 'disabled-{0}' -f (Get-Date -Format $script:TimestampFormat)
    $newName = '{0}.{1}' -f (Split-Path -Leaf $Folder), $suffix
    if (-not $PSCmdlet.ShouldProcess($Folder, "Rename to $newName")) { return $false }
    try {
        # Renaming the parent does not delete its DLLs, so it often remains possible after deletion is denied.
        Rename-Item -LiteralPath $Folder -NewName $newName -ErrorAction Stop
        $renamed = Join-Path (Split-Path -Parent $Folder) $newName
        Write-RunLog -Severity SUCCESS -Action 'Check prerequisite' -Result "Renamed that version to '$newName'. Module discovery only accepts a folder named after a valid version, so new PowerShell processes no longer see it."
        Write-RunLog -Severity INFO -Action 'Check prerequisite' -Result "Delete it once nothing holds it open: Remove-Item '$renamed' -Recurse -Force"
        return $true
    }
    catch {
        $folderReason = Get-ErrorText -ErrorRecord $_
        # Windows maps a loaded assembly into memory, which blocks renaming the folder holding it.
        # A manifest is only ever read, and discovery needs one, so renaming it retires the version.
        $manifests = @(Disable-ModuleManifest -Folder $Folder -Suffix $suffix -Confirm:$false)
        if ($manifests.Count -gt 0) {
            Write-RunLog -Severity SUCCESS -Action 'Check prerequisite' -Result "That folder is in use and could not be renamed, so its $($manifests.Count) manifest file(s) were renamed instead. PowerShell finds a module through its manifest, so new processes no longer see this version."
            Write-RunLog -Severity INFO -Action 'Check prerequisite' -Result "Delete the folder once nothing holds it open: Remove-Item '$Folder' -Recurse -Force"
            return $true
        }
        Add-RunFailure -FilePath $Folder -Action 'Check prerequisite' -Reason $folderReason
        return $false
    }
}

function Test-PnPVersionHygiene {
    <# .SYNOPSIS Warns when an older PnP copy can shadow the newest, and offers to remove it. #>
    [CmdletBinding()]
    param()

    $installed = @(Get-InstalledPnPVersion)
    if ($installed.Count -le 1) { return }
    $newest = $installed[0]
    $older = @($installed | Where-Object { $_ -ne $newest })

    # PnP loads .NET assemblies that cannot be unloaded, so whichever version a session touches first wins for its lifetime.
    Write-RunLog -Severity WARN -Action 'Check prerequisite' -Result "PnP.PowerShell $newest is installed alongside $($older -join ', '). Whichever version a session loads first wins, so an older copy can silently disable newer features such as registering a confidential client."

    $loaded = Get-Module -Name PnP.PowerShell | Sort-Object Version -Descending | Select-Object -First 1
    if ($loaded) {
        Write-RunLog -Severity INFO -Action 'Check prerequisite' -Result "PnP.PowerShell $($loaded.Version) is already loaded here. A loaded DLL cannot be deleted from this process, so cleanup will prefer renaming its version folder; this process must then close before another PnP version is loaded."
    }

    $choice = Read-MenuChoice -Title 'Remove the older PnP.PowerShell copies now?' -Options ([ordered]@{
            '1' = "Yes, uninstall $($older -join ', ') and keep $newest"
            '2' = 'No, leave them installed'
        }) -Default '1'
    if ($choice -ne '1') { return }

    foreach ($version in $older) {
        try {
            Uninstall-Module -Name PnP.PowerShell -RequiredVersion $version -Force -ErrorAction Stop
            Write-RunLog -Severity SUCCESS -Action 'Check prerequisite' -Result "Removed PnP.PowerShell $version."
        }
        catch {
            $folders = @(Get-Module -ListAvailable -Name PnP.PowerShell -All -ErrorAction SilentlyContinue |
                    Where-Object { [string]$_.Version -eq [string]$version } |
                    Select-Object -ExpandProperty ModuleBase -Unique)
            Write-RunLog -Severity WARN -Action 'Check prerequisite' -Result "Uninstall-Module could not remove PnP.PowerShell ${version}: $(Get-ErrorText -ErrorRecord $_). It may belong to another PowerShell edition, have a loaded DLL, or be protected by folder permissions."
            foreach ($folder in $folders) {
                # Refuse to touch anything that is not a PnP.PowerShell version folder.
                if ($folder -notmatch '(?i)[\\/]PnP\.PowerShell[\\/]' -or -not (Test-Path -LiteralPath $folder -PathType Container)) { continue }
                $loadedAssemblyPaths = @(Get-LoadedAssemblyPath -Folder $folder)
                if ($loadedAssemblyPaths.Count -gt 0) {
                    Write-RunLog -Severity WARN -Action 'Check prerequisite' -Result "This PowerShell process has loaded $($loadedAssemblyPaths.Count) assembly file(s) from '$folder', so deleting that folder cannot succeed until this process exits. Renaming is still worth trying."
                }
                $deleteChoice = Read-MenuChoice -Title 'Remove that version another way?' -Options ([ordered]@{
                        '1' = "Rename $folder so PowerShell stops discovering it (recommended)"
                        '2' = 'Try to delete it now'
                        '3' = 'No, leave it in place'
                    }) -Default '1'
                if ($deleteChoice -eq '3') {
                    Write-RunLog -Severity INFO -Action 'Check prerequisite' -Result "Remove it later with: Remove-Item '$folder' -Recurse -Force"
                    continue
                }
                if ($deleteChoice -eq '1') {
                    if (Disable-ModuleVersionFolder -Folder $folder -Confirm:$false) {
                        if ($loadedAssemblyPaths.Count -gt 0) {
                            throw 'The old PnP.PowerShell folder is disabled, but this PowerShell process already loaded assemblies from it. Close this PowerShell process and run the utility again; no Windows sign-out or reboot is required.'
                        }
                    }
                    else { Show-PnPHoldingProcess -Folder $folder }
                    continue
                }
                try {
                    Remove-Item -LiteralPath $folder -Recurse -Force -ErrorAction Stop
                    Write-RunLog -Severity SUCCESS -Action 'Check prerequisite' -Result "Deleted $folder."
                }
                catch {
                    Add-RunFailure -FilePath $folder -Action 'Check prerequisite' -Reason (Get-ErrorText -ErrorRecord $_)
                    Show-PnPHoldingProcess -Folder $folder
                    $renameChoice = Read-MenuChoice -Title 'Rename that folder instead, so this version stops shadowing the newest one?' -Options ([ordered]@{
                            '1' = 'Yes, rename it now'
                            '2' = 'No, I will remove it myself later'
                        }) -Default '1'
                    if ($renameChoice -eq '1' -and (Disable-ModuleVersionFolder -Folder $folder -Confirm:$false)) {
                        if ($loadedAssemblyPaths.Count -gt 0) {
                            throw 'The old PnP.PowerShell folder is disabled, but this PowerShell process already loaded assemblies from it. Close this PowerShell process and run the utility again; no Windows sign-out or reboot is required.'
                        }
                        continue
                    }
                    Write-RunLog -Severity INFO -Action 'Check prerequisite' -Result "Then run: Remove-Item '$folder' -Recurse -Force"
                }
            }
        }
    }

    $remaining = @(Get-InstalledPnPVersion)
    if ($remaining.Count -le 1) {
        Write-RunLog -Severity SUCCESS -Action 'Check prerequisite' -Result "PnP.PowerShell $newest is now the only version PowerShell can discover."
        return
    }
    Write-RunLog -Severity WARN -Action 'Check prerequisite' -Result "PowerShell can still discover PnP.PowerShell $($remaining -join ', '), so whichever one a session loads first still wins."
}

function Test-SharePointPrerequisite {
    <# .SYNOPSIS Imports PnP.PowerShell and verifies the cloud labeling cmdlets this utility needs. #>
    [CmdletBinding()]
    param()

    try {
        if ($PSVersionTable.PSVersion -lt [version]'7.2.0') {
            throw "SharePoint Online labeling needs PnP.PowerShell, which requires PowerShell 7.2 or later. This host is PowerShell $($PSVersionTable.PSVersion). Start pwsh.exe and run this script again, or choose the local folder source."
        }
        Test-PnPVersionHygiene
        $module = Get-Module -ListAvailable -Name PnP.PowerShell -ErrorAction Stop |
            Sort-Object Version -Descending | Select-Object -First 1
        if (-not $module) {
            if (Request-ModuleInstall -Name 'PnP.PowerShell' -Purpose 'for the SharePoint Online source') {
                $module = Get-Module -ListAvailable -Name PnP.PowerShell -ErrorAction Stop |
                    Sort-Object Version -Descending | Select-Object -First 1
            }
            if (-not $module) {
                throw 'PnP.PowerShell is not installed. Run: Install-Module PnP.PowerShell -Scope CurrentUser -Force'
            }
        }
        Import-Module PnP.PowerShell -RequiredVersion $module.Version -ErrorAction Stop
        # PnP's assemblies cannot be unloaded, so an older version already in the session keeps winning command resolution.
        $loaded = Get-Module -Name PnP.PowerShell | Sort-Object Version -Descending | Select-Object -First 1
        if ($loaded -and [version]$loaded.Version -lt [version]$module.Version) {
            Write-RunLog -Severity WARN -Action 'Check prerequisite' -Result "PnP.PowerShell $($loaded.Version) was already loaded in this session, but $($module.Version) is installed. Start a new PowerShell window to pick up the newer one; some features are unavailable in $($loaded.Version)."
        }
        foreach ($commandName in 'Connect-PnPOnline', 'Disconnect-PnPOnline', 'Get-PnPProperty', 'Get-PnPFolderItem', 'Get-PnPFileSensitivityLabel', 'Add-PnPFileSensitivityLabel') {
            if (-not (Get-Command $commandName -ErrorAction SilentlyContinue)) {
                throw "PnP.PowerShell $($module.Version) does not provide required command $commandName. Update it with: Update-Module PnP.PowerShell"
            }
        }
        Write-RunLog -Severity SUCCESS -Action 'Check prerequisite' -Result "PnP.PowerShell $($module.Version) is ready."
        return $true
    }
    catch {
        Add-RunFailure -FilePath '' -Action 'Check prerequisite' -Reason $_.Exception.Message
        return $false
    }
}

function Test-ExistingSharePointConnection {
    <# .SYNOPSIS Reports whether the open PnP connection already serves this site with this application. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [Parameter(Mandatory)][string]$ClientId
    )

    # A forced fresh sign-in exists precisely to discard the cached credential, so never short-circuit it.
    if ($script:ForceFreshSignIn) { return $false }
    if (-not (Get-Command Get-PnPConnection -ErrorAction SilentlyContinue)) { return $false }
    try {
        $connection = Get-PnPConnection -ErrorAction Stop
        if ($null -eq $connection) { return $false }
        $currentUrl = Get-ObjectPropertyValue -InputObject $connection -Names 'Url'
        if ([string]::IsNullOrWhiteSpace($currentUrl)) { return $false }
        if ($currentUrl.TrimEnd('/') -ne $SiteUrl.TrimEnd('/')) { return $false }
        $currentClientId = Get-ObjectPropertyValue -InputObject $connection -Names 'ClientId'
        if (-not [string]::IsNullOrWhiteSpace($currentClientId) -and $currentClientId -ne $ClientId) { return $false }
        # A connection object can outlive its token, so prove it still works before trusting it.
        $null = Get-PnPWeb -ErrorAction Stop
        return $true
    }
    catch {
        Write-Verbose "No reusable SharePoint connection: $($_.Exception.Message)"
        return $false
    }
}

function Connect-SharePointSite {
    <# .SYNOPSIS Signs in interactively to one SharePoint site using a tenant-owned Entra application. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][AllowEmptyString()][string]$TenantId,
        [switch]$JustRegistered
    )

    # Only a brand new registration can still be replicating, so a remembered application fails fast instead of waiting.
    $maxAttempts = if ($JustRegistered) { 4 } else { 1 }
    if (Test-ExistingSharePointConnection -SiteUrl $SiteUrl -ClientId $ClientId) {
        $script:SharePointSessionOpened = $true
        $script:ConnectedAppOnly = $false
        Write-RunLog -Severity SUCCESS -FilePath $SiteUrl -Action 'Connect SharePoint site' -Result "Reusing the sign-in already open for this site, so no browser window is needed."
        return $true
    }
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        try {
            Disconnect-SharePointSession
            if ($JustRegistered -and $attempt -eq 1) {
                Write-Host ''
                Write-Host '  Giving Entra a few seconds to replicate the new registration...' -ForegroundColor Gray
                Start-Sleep -Seconds 10
            }
            $parameters = @{Url = $SiteUrl; ClientId = $ClientId; ErrorAction = 'Stop'}
            # Without an explicit tenant, MSAL reuses a cached account from whichever tenant was signed in last.
            if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $parameters.Tenant = $TenantId }
            if ($script:UseDeviceCode) {
                $parameters.DeviceLogin = $true
                Write-Host ''
                Write-Host '  Device-code sign-in: open the URL below in any browser and enter the code.' -ForegroundColor Cyan
                Write-Host '  That page lets you pick Microsoft Authenticator instead of a passkey.' -ForegroundColor Gray
            }
            else {
                $parameters.Interactive = $true
                Write-Host ''
                Write-Host '  Opening a browser to sign in. This waits here until you finish there,' -ForegroundColor Cyan
                Write-Host '  so check for a window behind this one if nothing seems to happen.' -ForegroundColor Cyan
                Write-Host '  If it never appears, or only offers a passkey you cannot use, press' -ForegroundColor Gray
                Write-Host '  Ctrl+C and start again with:  .\Invoke-PurviewFileLabeling.ps1 -DeviceLogin' -ForegroundColor Gray
            }
            # A cached token also caches the credential choice, so a fresh prompt is the only way back to the account picker.
            if ($script:ForceFreshSignIn) { $parameters.ForceAuthentication = $true }
            Connect-PnPOnline @parameters
            $script:SharePointSessionOpened = $true
            $script:ConnectedAppOnly = $false
            # The forced prompt has served its purpose, so later connections may reuse this sign-in again.
            $script:ForceFreshSignIn = $false
            $null = Get-PnPWeb -ErrorAction Stop
            Write-RunLog -Severity SUCCESS -FilePath $SiteUrl -Action 'Connect SharePoint site' -Result "Signed in with application $ClientId in tenant $TenantId."
            return $true
        }
        catch {
            $message = Get-ErrorText -ErrorRecord $_
            $isMissingApplication = $message -match '(?i)AADSTS700016|was not found in the directory|application with identifier'
            $isWrongTenantAccount = $message -match '(?i)AADSTS50020|AADSTS500011|user account .* does not exist in tenant|was not found in the directory'
            if ($attempt -lt $maxAttempts -and $isMissingApplication) {
                Write-RunLog -Severity WARN -Action 'Connect SharePoint site' -Result "Application $ClientId is not visible in the tenant yet. Waiting for Entra to replicate the new registration (attempt $attempt)."
                Start-Sleep -Seconds 10
                continue
            }
            Add-RunFailure -FilePath $SiteUrl -Action 'Connect SharePoint site' -Reason $message
            if ($isMissingApplication) {
                Block-LabelingClientId -ClientId $ClientId -TenantId $TenantId
                Write-RunLog -Severity INFO -Action 'SharePoint sign-in guidance' -Result "Application $ClientId does not exist in tenant $TenantId. It was most likely registered in a different tenant, so register a new one at the next prompt."
            }
            elseif ($isWrongTenantAccount) {
                Write-RunLog -Severity INFO -Action 'SharePoint sign-in guidance' -Result "The account you signed in with does not exist in tenant $TenantId. Choose a different account in the browser, or sign out of the previous tenant first."
            }
            else {
                Write-RunLog -Severity INFO -Action 'SharePoint sign-in guidance' -Result 'The application needs delegated SharePoint AllSites.Read plus Microsoft Graph Files.Read.All. Both are user-consentable, so an administrator is only needed if tenant policy restricts user consent.'
                Write-RunLog -Severity INFO -Action 'SharePoint sign-in guidance' -Result 'If the browser offered only a passkey, or showed an Android Work Profile message, choose device-code sign-in at the next prompt. That page lets you pick Microsoft Authenticator instead.'
            }
            return $false
        }
    }
    return $false
}

function Disconnect-SharePointSession {
    <# .SYNOPSIS Closes a SharePoint session opened by this utility. #>
    [CmdletBinding()]
    param()

    if (-not (Get-Command Disconnect-PnPOnline -ErrorAction SilentlyContinue)) {
        $script:SharePointSessionOpened = $false
        $script:ConnectedAppOnly = $false
        return
    }
    try {
        Disconnect-PnPOnline -ErrorAction Stop
        if ($script:SharePointSessionOpened) {
            Write-RunLog -Severity INFO -Action 'Disconnect SharePoint site' -Result 'SharePoint session closed.'
        }
    }
    catch { Write-Verbose "No active SharePoint session required cleanup: $($_.Exception.Message)" }
    finally {
        $script:SharePointSessionOpened = $false
        $script:ConnectedAppOnly = $false
    }
}

function Get-SharePointTargetFile {
    <# .SYNOPSIS Enumerates files in a SharePoint library or folder and filters them by extension. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [Parameter(Mandatory)][string]$FolderSiteRelativeUrl,
        [Parameter(Mandatory)][AllowEmptyString()][string]$LibraryTitle,
        [Parameter(Mandatory)][string[]]$Extensions,
        [Parameter(Mandatory)][bool]$Recurse
    )

    try {
        $parameters = @{ItemType = 'File'; ErrorAction = 'Stop'}
        # -List is the only form that copes with libraries over the 5,000 item list view threshold, and it always recurses.
        if ($Recurse -and -not [string]::IsNullOrWhiteSpace($LibraryTitle) -and $FolderSiteRelativeUrl -notmatch '/') {
            $parameters.List = $LibraryTitle
        }
        else {
            $parameters.FolderSiteRelativeUrl = $FolderSiteRelativeUrl
            if ($Recurse) { $parameters.Recursive = $true }
        }
        $siteHost = ([uri]$SiteUrl).GetLeftPart([UriPartial]::Authority)
        $extensionSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$Extensions, [System.StringComparer]::OrdinalIgnoreCase)

        $files = [System.Collections.Generic.List[object]]::new()
        foreach ($item in @(Invoke-WithTransientRetry -Description 'Enumerate SharePoint files' -Operation { Get-PnPFolderItem @parameters })) {
            $name = Get-ObjectPropertyValue -InputObject $item -Names 'Name'
            $serverRelativeUrl = Get-ObjectPropertyValue -InputObject $item -Names 'ServerRelativeUrl'
            if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($serverRelativeUrl)) { continue }
            $extension = [System.IO.Path]::GetExtension($name)
            if (-not $extensionSet.Contains($extension)) { continue }
            $files.Add([pscustomobject]@{
                    Name = $name
                    FullName = $serverRelativeUrl
                    AbsoluteUrl = "$siteHost$serverRelativeUrl"
                    Extension = $extension.ToLowerInvariant()
                })
        }
        Write-RunLog -Severity INFO -FilePath $FolderSiteRelativeUrl -Action 'Enumerate files' -Result "Found $($files.Count) matching files."
        return @($files)
    }
    catch {
        Add-RunFailure -FilePath $FolderSiteRelativeUrl -Action 'Enumerate files' -Reason $_.Exception.Message
        return $null
    }
}

function Get-SharePointFileLabel {
    <# .SYNOPSIS Reads the sensitivity label currently applied to a SharePoint file. #>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'File', Justification = 'Used inside the retry scriptblock, which the analyser does not follow.')]
    param([Parameter(Mandatory)][object]$File)

    # Not Get-PnPFileSensitivityLabelInfo: despite the name, that one is a tenant-admin CSOM call and refuses ordinary accounts.
    $info = @(Invoke-WithTransientRetry -Description 'Read SharePoint file label' -Operation {
            Get-PnPFileSensitivityLabel -Url $File.AbsoluteUrl -ErrorAction Stop
        }) | Select-Object -First 1
    if ($null -eq $info) { return $null }
    $labelId = Get-ObjectPropertyValue -InputObject $info -Names 'SensitivityLabelId', 'sensitivityLabelId', 'LabelId', 'labelId', 'Id', 'id'
    if ([string]::IsNullOrWhiteSpace($labelId)) { return $null }
    $parsedLabelId = [guid]::Empty
    if (-not [guid]::TryParse($labelId, [ref]$parsedLabelId) -or $parsedLabelId -eq [guid]::Empty) { return $null }
    $labelName = Get-ObjectPropertyValue -InputObject $info -Names 'DisplayName', 'displayName', 'Name', 'name'
    if ([string]::IsNullOrWhiteSpace($labelName)) { $labelName = $parsedLabelId.ToString() }
    return [pscustomobject]@{Id = $parsedLabelId.ToString(); Name = $labelName}
}

function Set-OneSharePointFileLabel {
    <# .SYNOPSIS Applies one sensitivity label to a SharePoint Online file through the metered Graph API, then confirms it. #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][object]$File,
        [Parameter(Mandatory)][guid]$LabelId
    )

    if (-not $PSCmdlet.ShouldProcess($File.FullName, "Apply sensitivity label $LabelId")) {
        return [pscustomobject]@{Status = 'Skipped'; Comment = 'ShouldProcess declined the operation.'}
    }
    # This PnP cmdlet calls driveItem/assignSensitivityLabel. The caller categorises any Graph failure.
    $null = Add-PnPFileSensitivityLabel -Identity $File.FullName -SensitivityLabelId $LabelId -AssignmentMethod Standard -ErrorAction Stop

    # Graph returns 202 Accepted, so a request that did not throw has not necessarily finished.
    for ($attempt = 1; $attempt -le 4; $attempt++) {
        Start-Sleep -Seconds ($attempt * 2)
        try {
            $applied = Get-SharePointFileLabel -File $File
            if ($applied -and $applied.Id -eq $LabelId.ToString()) {
                return [pscustomobject]@{Status = 'Success'; Comment = 'Label applied and confirmed on the file.'}
            }
        }
        catch { Write-Verbose "Could not read the label back for $($File.FullName): $($_.Exception.Message)" }
    }
    return [pscustomobject]@{
        Status = 'Unconfirmed'
        Comment = 'Graph accepted the request but the file does not report the label yet. assignSensitivityLabel is asynchronous, so it may still be processing; if it never appears, check that a confidential client is in use and that its Azure billing link exists.'
    }
}

function Get-TargetFile {
    <# .SYNOPSIS Enumerates target files using explicit extension filters. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string[]]$Extensions,
        [Parameter(Mandatory)][bool]$Recurse
    )

    try {
        $parameters = @{LiteralPath = $Path; File = $true; ErrorAction = 'Stop'}
        if ($Recurse) { $parameters.Recurse = $true }
        $extensionSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$Extensions, [System.StringComparer]::OrdinalIgnoreCase)
        $files = @(Get-ChildItem @parameters | Where-Object { $extensionSet.Contains($_.Extension) })
        Write-RunLog -Severity INFO -FilePath $Path -Action 'Enumerate files' -Result "Found $($files.Count) matching files."
        return $files
    }
    catch {
        $reason = if ($_.Exception.Message -match 'denied|unauthorized') { "Access denied. $($_.Exception.Message)" } else { $_.Exception.Message }
        Add-RunFailure -FilePath $Path -Action 'Enumerate files' -Reason $reason
        return $null
    }
}

function Test-PurviewAuthentication {
    <# .SYNOPSIS Probes authentication by reading one file's Purview status. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ProbeFile)

    try {
        $null = Get-FileStatus -Path $ProbeFile -ErrorAction Stop
        Write-RunLog -Severity SUCCESS -FilePath $ProbeFile -Action 'Verify authentication' -Result 'Get-FileStatus succeeded.'
        return $true
    }
    catch {
        Add-RunFailure -FilePath $ProbeFile -Action 'Verify authentication' -Reason $_.Exception.Message
        Write-RunLog -Severity INFO -Action 'Authentication guidance' -Result 'In PowerShell, run: Import-Module PurviewInformationProtection; Set-Authentication. Complete sign-in, then select re-check.'
        return $false
    }
}

function Test-LabelReadAccess {
    <# .SYNOPSIS Confirms the selected source can be read before any label is written. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$ProbeFile,
        [ValidateSet('Local', 'SharePoint')][string]$Source = 'Local'
    )

    if ($Source -ne 'SharePoint') { return Test-PurviewAuthentication -ProbeFile $ProbeFile.FullName }

    try {
        $null = Get-SharePointFileLabel -File $ProbeFile
        Write-RunLog -Severity SUCCESS -FilePath $ProbeFile.FullName -Action 'Verify authentication' -Result 'Read the sensitivity label through Microsoft Graph.'
        return $true
    }
    catch {
        Add-RunFailure -FilePath $ProbeFile.FullName -Action 'Verify authentication' -Reason $_.Exception.Message
        Write-RunLog -Severity INFO -Action 'Authentication guidance' -Result 'Reading a label needs delegated Microsoft Graph Files.Read.All at minimum. Browsing the library uses SharePoint permissions instead, so it can succeed while this step is refused.'
        Write-RunLog -Severity INFO -Action 'Authentication guidance' -Result 'To see which scopes this sign-in actually obtained, run: Get-PnPAccessToken -Scopes. If Files.Read.All is missing, sign out, register a fresh application, and accept the consent prompt in the browser.'
        return $false
    }
}

function Get-ExistingLabel {
    <# .SYNOPSIS Extracts the effective label identity and name from file status. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Status)

    if (-not $Status.IsLabeled) { return $null }
    if ($Status.SubLabelId -and ([string]$Status.SubLabelId -ne [guid]::Empty.ToString())) {
        return [pscustomobject]@{Id = [string]$Status.SubLabelId; Name = [string]$Status.SubLabelName}
    }
    return [pscustomobject]@{Id = [string]$Status.MainLabelId; Name = [string]$Status.MainLabelName}
}

function Install-GalleryModule {
    <# .SYNOPSIS Installs a PowerShell Gallery module for the current user, which needs no administrator. #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory)][string]$Name,
        [string]$RequiredVersion = ''
    )

    if (-not $PSCmdlet.ShouldProcess($Name, 'Install module for the current user')) { return $false }
    try {
        # The gallery requires TLS 1.2, which Windows PowerShell does not always negotiate by default.
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 }
        catch { Write-Verbose "Could not raise the TLS version: $($_.Exception.Message)" }
        try { $null = Get-PackageProvider -Name NuGet -ForceBootstrap -ErrorAction Stop }
        catch { Write-Verbose "NuGet provider bootstrap skipped: $($_.Exception.Message)" }

        $parameters = @{Name = $Name; Scope = 'CurrentUser'; Force = $true; AllowClobber = $true; ErrorAction = 'Stop'}
        if (-not [string]::IsNullOrWhiteSpace($RequiredVersion)) { $parameters.RequiredVersion = $RequiredVersion }
        Write-RunLog -Severity INFO -Action 'Install module' -Result "Installing $Name for the current user. This can take a minute."
        try {
            Install-Module @parameters
        }
        catch {
            # PowerShellGet blocks an upgrade when the signing publisher changes, which PnP.PowerShell did.
            if ((Get-ErrorText -ErrorRecord $_) -notmatch '(?i)SkipPublisherCheck') { throw }
            Write-Host ''
            Write-Host "  The installed $Name was signed by a different publisher than the new version," -ForegroundColor Yellow
            Write-Host '  so PowerShell refused the upgrade. PnP.PowerShell genuinely moved from Microsoft' -ForegroundColor Yellow
            Write-Host '  to the .NET Foundation, which produces exactly this message, but the check exists' -ForegroundColor Yellow
            Write-Host '  to catch a package that changed hands unexpectedly. Continue only if that' -ForegroundColor Yellow
            Write-Host '  explanation fits the module being installed.' -ForegroundColor Yellow
            $choice = Read-MenuChoice -Title "Install $Name anyway, skipping the publisher check?" -Options ([ordered]@{
                    '1' = 'No, stop so I can check first (default)'
                    '2' = 'Yes, install with -SkipPublisherCheck'
                }) -Default '1'
            if ($choice -ne '2') { throw }
            $parameters.SkipPublisherCheck = $true
            Install-Module @parameters
        }
        Write-RunLog -Severity SUCCESS -Action 'Install module' -Result "$Name is installed."
        $olderVersions = @(Get-InstalledPnPVersion)
        if ($Name -eq 'PnP.PowerShell' -and $olderVersions.Count -gt 1) {
            Write-RunLog -Severity WARN -Action 'Install module' -Result "PnP.PowerShell $($olderVersions -join ', ') are now all installed. The older copies still shadow the newest one, so remove them and start a new PowerShell window before using the new features."
        }
        return $true
    }
    catch {
        Add-RunFailure -FilePath '' -Action 'Install module' -Reason (Get-ErrorText -ErrorRecord $_)
        Write-RunLog -Severity INFO -Action 'Install module' -Result "Install it yourself with: Install-Module $Name -Scope CurrentUser -Force"
        return $false
    }
}

function Request-ModuleInstall {
    <# .SYNOPSIS Offers to install or update a module, so a prerequisite gap is recoverable without leaving the utility. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Purpose,
        [string]$RequiredVersion = ''
    )

    $isPresent = @(Get-Module -ListAvailable -Name $Name -ErrorAction SilentlyContinue).Count -gt 0
    if ($isPresent) {
        Write-RunLog -Severity WARN -Action 'Check prerequisite' -Result "$Name is required $Purpose, and the installed version is too old."
        $title = "Update $Name now?"
        $yes = "Yes, install a newer $Name from the PowerShell Gallery for my user account"
    }
    else {
        Write-RunLog -Severity WARN -Action 'Check prerequisite' -Result "$Name is required $Purpose, and it is not installed."
        $title = "Install $Name now?"
        $yes = "Yes, install $Name from the PowerShell Gallery for my user account"
    }
    $choice = Read-MenuChoice -Title $title -Options ([ordered]@{
            '1' = $yes
            '2' = 'No, I will do it myself'
        }) -Default '1'
    if ($choice -ne '1') { return $false }
    return (Install-GalleryModule -Name $Name -RequiredVersion $RequiredVersion -Confirm:$false)
}

function Get-ErrorCategory {
    <# .SYNOPSIS Maps a Purview or file-system error to an operator-facing category. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Message)

    # Each branch breaks, because switch -Regex otherwise runs every matching branch and returns them all.
    switch -Regex ($Message) {
        'payment ?required|graph-metered|premium API' { 'Graph metered API not enabled'; break }
        'not authenticated|sign.?in|access token|authentication' { 'Not authenticated'; break }
        'denied|unauthorized|0x80070005' { 'Access denied'; break }
        'used by another process|locked|sharing violation' { 'File locked or in use'; break }
        'not supported|unsupported|file type' { 'Unsupported file type'; break }
        default { 'Purview operation failed' }
    }
}

function Test-FileAvailable {
    <# .SYNOPSIS Checks whether a file can be opened exclusively before labeling. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    try {
        $stream = [System.IO.File]::Open($Path, 'Open', 'ReadWrite', 'None')
        $stream.Dispose()
        return [pscustomobject]@{Available = $true; Reason = ''}
    }
    catch {
        return [pscustomobject]@{Available = $false; Reason = $_.Exception.Message}
    }
}

function Clear-PurviewHandle {
    <# .SYNOPSIS Releases the native resources the Purview client leaves behind after labeling. #>
    [CmdletBinding()]
    param()

    $script:FilesSinceCollection = 0
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Set-OneFileLabel {
    <# .SYNOPSIS Applies one sensitivity label after ShouldProcess approval. #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][guid]$LabelId
    )

    if (-not $PSCmdlet.ShouldProcess($Path, "Apply sensitivity label $LabelId")) {
        return [pscustomobject]@{Status = 'Skipped'; Comment = 'ShouldProcess declined the operation.'}
    }
    try {
        return Set-FileLabel -Path $Path -LabelId $LabelId -PreserveFileDetails -ErrorAction Stop
    }
    finally {
        # A blocking full collection after every file dominates the runtime of a large batch, so sweep periodically instead.
        $script:FilesSinceCollection++
        if ($script:FilesSinceCollection -ge 50) { Clear-PurviewHandle }
    }
}

function Add-FileResult {
    <# .SYNOPSIS Adds one file outcome to the final report collection. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [string]$PreviousLabel = '',
        [string]$NewLabel = '',
        [Parameter(Mandatory)][string]$Outcome,
        [string]$Details = ''
    )

    $script:Results.Add([pscustomobject]@{
            FilePath = $FilePath
            PreviousLabel = $PreviousLabel
            NewLabel = $NewLabel
            Outcome = $Outcome
            Details = $Details
        })
}

function Invoke-FileProcessing {
    <# .SYNOPSIS Reads status and optionally labels each selected file. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$Files,
        [Parameter(Mandatory)][object]$TargetLabel,
        [Parameter(Mandatory)][object[]]$ConfiguredLabels,
        [Parameter(Mandatory)][bool]$DryRun,
        [ValidateSet('Local', 'SharePoint')][string]$Source = 'Local'
    )

    $priorityById = @{}
    $nameById = @{}
    foreach ($configuredLabel in $ConfiguredLabels) {
        $priorityById[[string]$configuredLabel.Id] = [int]$configuredLabel.Priority
        $nameById[[string]$configuredLabel.Id] = [string]$configuredLabel.Name
    }

    $skipAllErrors = $false
    $fileNumber = 0
    $activity = 'Processing {0} file(s)' -f $Files.Count
    foreach ($file in $Files) {
        $fileNumber++
        Write-Progress -Id 1 -Activity $activity -Status "$fileNumber of $($Files.Count): $($file.Name)" -PercentComplete ([math]::Min(100, [int](($fileNumber * 100) / $Files.Count)))
        $retry = $true
        while ($retry) {
            $retry = $false
            try {
                $existing = if ($Source -eq 'SharePoint') {
                    Get-SharePointFileLabel -File $file
                }
                else {
                    Get-ExistingLabel -Status (Get-FileStatus -Path $file.FullName -ErrorAction Stop)
                }
                # Graph reports only the label GUID, so the readable name comes from the tenant list.
                $previousName = if (-not $existing) { '' }
                elseif ($nameById.ContainsKey($existing.Id)) { $nameById[$existing.Id] }
                else { $existing.Name }

                if ($existing -and -not $priorityById.ContainsKey($existing.Id)) {
                    $details = 'Existing label is not in the priority config; skipped conservatively.'
                    Add-FileResult -FilePath $file.FullName -PreviousLabel $previousName -NewLabel $TargetLabel.Name -Outcome 'Skipped' -Details $details
                    Write-RunLog -Severity WARN -FilePath $file.FullName -Action 'Evaluate label priority' -Result $details
                    break
                }
                if ($existing -and $priorityById[$existing.Id] -ge [int]$TargetLabel.Priority) {
                    $details = 'File already has an equal or higher priority label.'
                    Add-FileResult -FilePath $file.FullName -PreviousLabel $previousName -NewLabel $TargetLabel.Name -Outcome 'Skipped' -Details $details
                    Write-RunLog -Severity WARN -FilePath $file.FullName -Action 'Evaluate label priority' -Result $details
                    break
                }
                if ($DryRun) {
                    Add-FileResult -FilePath $file.FullName -PreviousLabel $previousName -NewLabel $TargetLabel.Name -Outcome 'DryRun' -Details 'Would apply label.'
                    Write-RunLog -Severity INFO -FilePath $file.FullName -Action 'Dry run label' -Result "Would apply $($TargetLabel.Name)."
                    break
                }

                $result = if ($Source -eq 'SharePoint') {
                    Set-OneSharePointFileLabel -File $file -LabelId ([guid]$TargetLabel.Id) -Confirm:$false
                }
                else {
                    $availability = Test-FileAvailable -Path $file.FullName
                    if (-not $availability.Available) { throw $availability.Reason }
                    Set-OneFileLabel -Path $file.FullName -LabelId ([guid]$TargetLabel.Id) -Confirm:$false
                }
                $outcome = switch ([string]$result.Status) {
                    'Success' { 'Labeled' }
                    'Unconfirmed' { 'Unconfirmed' }
                    default { 'Skipped' }
                }
                $details = [string]$result.Comment
                Add-FileResult -FilePath $file.FullName -PreviousLabel $previousName -NewLabel $TargetLabel.Name -Outcome $outcome -Details $details
                $severity = if ($outcome -eq 'Labeled') { 'SUCCESS' } else { 'WARN' }
                Write-RunLog -Severity $severity -FilePath $file.FullName -Action 'Apply label' -Result "$outcome. $details"
            }
            catch {
                $message = Get-ErrorText -ErrorRecord $_
                $category = Get-ErrorCategory -Message $message
                $reason = "${category}: $message"
                Add-RunFailure -FilePath $file.FullName -Action 'Process file' -Reason $reason

                # A tenant-level billing gap fails identically on every remaining file, so asking per file is pointless.
                if ($category -eq 'Graph metered API not enabled') {
                    Add-FileResult -FilePath $file.FullName -NewLabel $TargetLabel.Name -Outcome 'Failed' -Details $reason
                    Write-RunLog -Severity ERROR -Action 'Apply label' -Result 'Microsoft Graph rejected the assignment with paymentRequired, so the run stopped instead of failing on every remaining file.'
                    Write-RunLog -Severity INFO -Action 'Metered API guidance' -Result 'assignSensitivityLabel is billed per call and accepted only from a confidential client. Choose "Enable SharePoint Online metered label writing" on the main menu to register one and link an Azure subscription. See https://aka.ms/graph-metered-overview'
                    Write-RunLog -Severity INFO -Action 'Metered API guidance' -Result 'A token issued before the billing link was created is still refused, so start the utility again after linking. At no extra cost: use a Purview auto-labeling policy, or sync the library with OneDrive and run the local source against the synced folder.'
                    return 'Main'
                }

                $choice = if ($skipAllErrors) {
                    '3'
                }
                else {
                    Read-MenuChoice -Title 'The file operation failed. Choose an action.' -Options ([ordered]@{
                            '1' = 'Retry this item'
                            '2' = 'Retry with different input (return to settings)'
                            '3' = 'Skip and continue'
                            '4' = 'Return to the main menu'
                            '5' = 'Skip this and every later failure without asking again'
                        }) -Default '3'
                }
                if ($choice -eq '1') { $retry = $true; continue }
                if ($choice -eq '5') { $skipAllErrors = $true }
                Add-FileResult -FilePath $file.FullName -NewLabel $TargetLabel.Name -Outcome 'Failed' -Details $reason
                if ($choice -eq '2') { return 'Change' }
                if ($choice -eq '4') { return 'Main' }
            }
        }
    }
    Write-Progress -Id 1 -Activity $activity -Completed
    return 'Complete'
}

function Export-RunReport {
    <# .SYNOPSIS Exports all collected file outcomes to CSV. #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
    param()

    if (-not $script:ReportPath) { return }
    try {
        if ($PSCmdlet.ShouldProcess($script:ReportPath, 'Export CSV report')) {
            $script:Results | Export-Csv -LiteralPath $script:ReportPath -NoTypeInformation -Encoding $script:CsvEncoding -ErrorAction Stop
            Write-RunLog -Severity SUCCESS -FilePath $script:ReportPath -Action 'Export report' -Result "Exported $($script:Results.Count) rows."
        }
    }
    catch {
        Add-RunFailure -FilePath $script:ReportPath -Action 'Export report' -Reason $_.Exception.Message
    }
}

function Show-RunSummary {
    <# .SYNOPSIS Prints totals and elapsed time for the current session. #>
    [CmdletBinding()]
    param()

    $labeled = @($script:Results | Where-Object Outcome -eq 'Labeled').Count
    $unconfirmed = @($script:Results | Where-Object Outcome -eq 'Unconfirmed').Count
    $skipped = @($script:Results | Where-Object Outcome -in 'Skipped', 'DryRun').Count
    $failed = @($script:Results | Where-Object Outcome -eq 'Failed').Count
    $elapsed = (Get-Date) - $script:RunStarted
    $summary = 'Total scanned: {0}; labeled: {1}; unconfirmed: {2}; skipped/dry-run: {3}; failed: {4}; tracked failures: {5}; elapsed: {6:hh\:mm\:ss}' -f `
        $script:Results.Count, $labeled, $unconfirmed, $skipped, $failed, $script:Failures.Count, $elapsed
    Write-RunLog -Severity INFO -Action 'Closing summary' -Result $summary -NoConsole
    Write-Host ''
    Write-Host "  Session summary  (utility $script:UtilityVersion)" -ForegroundColor Cyan
    Write-Host "    Files scanned    : $($script:Results.Count)"
    Write-Host "    Labeled          : $labeled"
    Write-Host "    Not confirmed    : $unconfirmed" -ForegroundColor $(if ($unconfirmed -gt 0) { 'Yellow' } else { 'Gray' })
    Write-Host "    Skipped/dry-run  : $skipped"
    Write-Host "    Failed           : $failed" -ForegroundColor $(if ($failed -gt 0) { 'Red' } else { 'Gray' })
    Write-Host "    Tracked failures : $($script:Failures.Count)"
    Write-Host ('    Elapsed          : {0:hh\:mm\:ss}' -f $elapsed)
    if ($script:ReportPath) { Write-Host "    Report           : $script:ReportPath" }
    if ($unconfirmed -gt 0) {
        Write-Host ''
        Write-Host '  Files counted as not confirmed were accepted by Microsoft Graph but do not' -ForegroundColor Yellow
        Write-Host '  report the label yet. assignSensitivityLabel is asynchronous, so re-scanning' -ForegroundColor Yellow
        Write-Host '  shortly may show them labeled. If they never change, confirm that a' -ForegroundColor Yellow
        Write-Host '  confidential client is configured and that its Azure billing link exists.' -ForegroundColor Yellow
    }
}

function Get-DependencyReport {
    <# .SYNOPSIS Reports the installed version of every module this utility can use. #>
    [CmdletBinding()]
    param()

    # One scan: -ListAvailable walks every module path on each call, which costs about a second.
    $names = @($script:ModuleExpectation | ForEach-Object { $_.Name } | Sort-Object -Unique)
    $newestByName = @{}
    foreach ($module in @(Get-Module -ListAvailable -Name $names -ErrorAction SilentlyContinue)) {
        if ($null -eq $module.Version) { continue }
        $name = [string]$module.Name
        $version = [version]$module.Version
        if (-not $newestByName.ContainsKey($name) -or $version -gt $newestByName[$name]) {
            $newestByName[$name] = $version
        }
    }

    return @(foreach ($expectation in $script:ModuleExpectation) {
            $version = if ($newestByName.ContainsKey($expectation.Name)) { $newestByName[$expectation.Name] } else { $null }
            [pscustomobject]@{
                Name = $expectation.Name
                Installed = $version
                Minimum = $expectation.Minimum
                Purpose = $expectation.Purpose
                State = if ($null -eq $version) { 'Missing' }
                elseif ($version -lt $expectation.Minimum) { 'TooOld' }
                else { 'Ready' }
            }
        })
}

function Test-DependencyDrift {
    <# .SYNOPSIS Records the dependency versions this run used and warns about combinations known to break. #>
    [CmdletBinding()]
    param()

    $report = @(Get-DependencyReport)
    foreach ($row in $report) {
        $state = switch ($row.State) {
            'Missing' { 'not installed' }
            'TooOld' { "$($row.Installed), older than the $($row.Minimum) this utility expects" }
            default { [string]$row.Installed }
        }
        Write-RunLog -Severity INFO -Action 'Check dependency' -Result "$($row.Name): $state" -NoConsole
    }
    foreach ($row in @($report | Where-Object { $_.State -eq 'TooOld' })) {
        Write-RunLog -Severity WARN -Action 'Check dependency' -Result "$($row.Name) $($row.Installed) is older than the $($row.Minimum) expected for $($row.Purpose), so some steps may fail in ways this utility cannot explain. Update it with: Update-Module $($row.Name) -Scope CurrentUser"
    }

    # Az and PnP ship different major versions of the same Microsoft.Extensions assemblies, and the first
    # one loaded wins for the life of the process. Az therefore runs in child processes, and finding it
    # loaded here means that isolation has been broken.
    $azLoaded = @(Get-Module -Name 'Az.*' -ErrorAction SilentlyContinue | ForEach-Object Name)
    $pnpLoaded = @(Get-Module -Name 'PnP.PowerShell' -ErrorAction SilentlyContinue)
    if ($azLoaded.Count -gt 0 -and $pnpLoaded.Count -gt 0) {
        Write-RunLog -Severity WARN -Action 'Check dependency' -Result "This session has loaded both PnP.PowerShell and $($azLoaded -join ', '), which ship incompatible Microsoft.Extensions assemblies. SharePoint sign-in can fail here with a missing 'get_Services' implementation. Start a new PowerShell window and run this utility again."
    }
}

function Write-RunEnvironment {
    <# .SYNOPSIS Stamps the log with what this run was, so a later report is diagnosable without guesswork. #>
    [CmdletBinding()]
    param()

    $edition = if ($PSVersionTable.PSEdition) { $PSVersionTable.PSEdition } else { 'Desktop' }
    Write-RunLog -Severity INFO -Action 'Record environment' -Result "Utility $script:UtilityVersion on PowerShell $($PSVersionTable.PSVersion) ($edition), $([System.Environment]::OSVersion.VersionString), process $PID." -NoConsole
    Write-RunLog -Severity INFO -Action 'Record environment' -Result "Script: $script:ScriptPath" -NoConsole
}

function Get-RememberedValue {
    <# .SYNOPSIS Reads a remembered non-secret value, such as the site the test-data helper just created. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)

    if ($script:IgnoreRemembered) { return '' }
    foreach ($target in [EnvironmentVariableTarget]::Process, [EnvironmentVariableTarget]::User) {
        try {
            $value = [Environment]::GetEnvironmentVariable($Name, $target)
            if (-not [string]::IsNullOrWhiteSpace($value)) { return $value.Trim() }
        }
        catch { Write-Verbose "Could not read ${Name}: $($_.Exception.Message)" }
    }
    return ''
}

function Save-RememberedValue {
    <# .SYNOPSIS Remembers a non-secret value so a later prompt can propose it. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { return }
    [Environment]::SetEnvironmentVariable($Name, $Value, [EnvironmentVariableTarget]::Process)
    try { [Environment]::SetEnvironmentVariable($Name, $Value, [EnvironmentVariableTarget]::User) }
    catch { Write-Verbose "Could not persist ${Name}: $($_.Exception.Message)" }
}

function Get-SessionValue {
    <# .SYNOPSIS Reads a process-only value handed off within the current console session. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)

    if ($script:IgnoreRemembered) { return '' }
    try {
        $value = [Environment]::GetEnvironmentVariable($Name, [EnvironmentVariableTarget]::Process)
        if (-not [string]::IsNullOrWhiteSpace($value)) { return $value.Trim() }
    }
    catch { Write-Verbose "Could not read ${Name}: $($_.Exception.Message)" }
    return ''
}

function Save-SessionValue {
    <# .SYNOPSIS Keeps tenant context in this process without persisting it for later runs. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { return }
    [Environment]::SetEnvironmentVariable($Name, $Value, [EnvironmentVariableTarget]::Process)
}

function Clear-LegacyTenantContextPreference {
    <# .SYNOPSIS Removes global site and library defaults persisted by earlier versions. #>
    [CmdletBinding()]
    param()

    $handoff = [Environment]::GetEnvironmentVariable('PURVIEW_FILE_LABELING_PROCESS_HANDOFF', [EnvironmentVariableTarget]::Process) -eq '1'
    [Environment]::SetEnvironmentVariable('PURVIEW_FILE_LABELING_PROCESS_HANDOFF', $null, [EnvironmentVariableTarget]::Process)
    foreach ($name in 'LABEL_TEST_SITE_TENANT_URL', 'PURVIEW_FILE_LABELING_SITE_URL', 'PURVIEW_FILE_LABELING_LIBRARY') {
        $targets = [System.Collections.Generic.List[EnvironmentVariableTarget]]::new()
        $targets.Add([EnvironmentVariableTarget]::User)
        if ($name -eq 'LABEL_TEST_SITE_TENANT_URL' -or -not $handoff) {
            $targets.Add([EnvironmentVariableTarget]::Process)
        }
        foreach ($target in $targets) {
            try {
                $value = [Environment]::GetEnvironmentVariable($name, $target)
                if ([string]::IsNullOrWhiteSpace($value)) { continue }
                [Environment]::SetEnvironmentVariable($name, $null, $target)
                Write-RunLog -Severity INFO -Action 'Scrub tenant context' -Result "Removed legacy $target default $name."
            }
            catch { Write-Verbose "Could not clear ${name} from ${target}: $($_.Exception.Message)" }
        }
    }
}

function Get-TenantScopedClientIdName {
    <# .SYNOPSIS Builds the per-tenant variable name, so an application from one tenant is never offered to another. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$TenantId)

    $parsed = [guid]::Empty
    if (-not [guid]::TryParse($TenantId, [ref]$parsed) -or $parsed -eq [guid]::Empty) { return '' }
    return 'PURVIEW_FILE_LABELING_CLIENT_ID_' + $parsed.ToString('N').ToUpperInvariant()
}

function Get-ConfiguredClientId {
    <# .SYNOPSIS Finds a remembered Entra application ID for one tenant, skipping any the tenant already rejected. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$TenantId)

    if ($script:IgnoreRemembered) { return $null }
    $tenantScopedName = Get-TenantScopedClientIdName -TenantId $TenantId
    $names = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($tenantScopedName)) { $names.Add($tenantScopedName) }
    # The legacy and third-party variables are tenant-agnostic, so they are offered only as unverified fallbacks.
    foreach ($name in 'PURVIEW_FILE_LABELING_CLIENT_ID', 'ENTRAID_APP_ID', 'ENTRAID_CLIENT_ID', 'AZURE_CLIENT_ID') { $names.Add($name) }

    foreach ($variableName in $names) {
        foreach ($target in [EnvironmentVariableTarget]::Process, [EnvironmentVariableTarget]::User) {
            try {
                $value = [Environment]::GetEnvironmentVariable($variableName, $target)
                if ([string]::IsNullOrWhiteSpace($value)) { continue }
                $value = $value.Trim()
                if ($script:RejectedClientIds.Contains($value)) { continue }
                return [pscustomobject]@{
                    ClientId = $value
                    VariableName = $variableName
                    Scope = [string]$target
                    IsTenantScoped = ($variableName -eq $tenantScopedName)
                }
            }
            catch { Write-Verbose "Could not read ${variableName}: $($_.Exception.Message)" }
        }
    }
    return $null
}

function Read-ValueWithDefault {
    <# .SYNOPSIS Proposes a value and lets the operator accept it with Enter or type a replacement. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Default
    )

    while ($true) {
        $suffix = if ([string]::IsNullOrWhiteSpace($Default)) { '' } else { " [$Default]" }
        $value = Read-OperatorInput -Prompt "  $Prompt$suffix"
        if ([string]::IsNullOrWhiteSpace($value)) { $value = $Default }
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $value = $value.Trim()
            Write-RunLog -Severity INFO -Action 'Read setting' -Result "${Prompt}: $value" -NoConsole
            return $value
        }
        Write-RunLog -Severity WARN -Action 'Validate setting' -Result 'A value is required.'
    }
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

function Restart-InPowerShell7 {
    <# .SYNOPSIS Decides whether to hand this run over to PowerShell 7 with source choice preserved. #>
    [CmdletBinding()]
    param(
        [ValidateSet('Local', 'SharePoint')][string]$Source = ''
    )

    if ([string]::IsNullOrWhiteSpace($script:ScriptPath) -or -not (Test-Path -LiteralPath $script:ScriptPath -PathType Leaf)) {
        Write-RunLog -Severity WARN -Action 'Restart in PowerShell 7' -Result 'This utility could not locate its own script file, so it cannot restart itself.'
        return $false
    }
    $hostPath = Get-PowerShell7Path -MinimumVersion ([version]'7.2.0')
    if ([string]::IsNullOrWhiteSpace($hostPath)) {
        Write-RunLog -Severity INFO -Action 'Restart in PowerShell 7' -Result 'PowerShell 7.2 or later is not installed. Install it with: winget install --id Microsoft.PowerShell --source winget'
        return $false
    }

        $choice = Read-MenuChoice -Title "Restart this utility in $hostPath so the SharePoint Online source can be used?" -Options ([ordered]@{
            '1' = 'Yes, restart now (your authentication will be reused)'
            '2' = 'No, stay here and choose the local/on-prem file path source'
        }) -Default '1'
    if ($choice -ne '1') { return $false }

    Write-RunLog -Severity SUCCESS -Action 'Restart in PowerShell 7' -Result "Closing this session and handing over to $hostPath." -NoConsole
    # The launch is deferred to script level, because a child process started inside a
    # function has its console output captured as that function's return value.
    $script:RelaunchHostPath = $hostPath
    $script:RelaunchCompleted = $true
    # Preserve the source choice and the reason for restart
    if (-not [string]::IsNullOrWhiteSpace($Source)) {
        $script:RelaunchReason = "source choice ($Source)"
    }
    return $true
}

function Test-SharePointSourceAvailable {
    <# .SYNOPSIS Reports up front whether this host can use the SharePoint Online source at all. #>
    [CmdletBinding()]
    param([version]$HostVersion = $PSVersionTable.PSVersion)

    if ($HostVersion -ge [version]'7.2.0') { return $true }
    Write-RunLog -Severity WARN -Action 'Select source' -Result "The SharePoint Online source needs PnP.PowerShell, which requires PowerShell 7.2 or later. This host is Windows PowerShell $HostVersion."
    if (-not $NoRelaunch -and (Restart-InPowerShell7 -Source 'SharePoint')) { return $false }

    $command = if ([string]::IsNullOrWhiteSpace($script:ScriptPath)) { 'pwsh -File .\Invoke-PurviewFileLabeling.ps1' } else { "pwsh -File `"$($script:ScriptPath)`"" }
    Write-RunLog -Severity INFO -Action 'Select source' -Result "To use it later, run: $command"
    Write-RunLog -Severity INFO -Action 'Select source' -Result 'The local/on-prem file path source works in this host with the Purview Information Protection client.'
    return $false
}

function Get-SharePointDocumentLibrary {
    <# .SYNOPSIS Lists the visible document libraries of the connected site with their site-relative URLs. #>
    [CmdletBinding()]
    param()

    try {
        $web = Get-PnPWeb -Includes ServerRelativeUrl -ErrorAction Stop
        $webRelativeUrl = ([string]$web.ServerRelativeUrl).TrimEnd('/')

        $libraries = [System.Collections.Generic.List[object]]::new()
        # RootFolder is loaded per library, because asking for it across every list also touches system lists the account cannot read.
        foreach ($list in @(Get-PnPList -ErrorAction Stop)) {
            if ([int]$list.BaseTemplate -ne 101 -or $list.Hidden) { continue }
            $serverRelativeUrl = ''
            try {
                $serverRelativeUrl = [string](Get-PnPProperty -ClientObject $list -Property RootFolder -ErrorAction Stop).ServerRelativeUrl
            }
            catch {
                Write-RunLog -Severity WARN -Action 'List document libraries' -Result "Skipped '$($list.Title)' because its folder could not be read. $($_.Exception.Message)"
                continue
            }
            if ([string]::IsNullOrWhiteSpace($serverRelativeUrl)) { continue }
            $siteRelativeUrl = $serverRelativeUrl
            if ($webRelativeUrl -and $siteRelativeUrl.StartsWith($webRelativeUrl, [System.StringComparison]::OrdinalIgnoreCase)) {
                $siteRelativeUrl = $siteRelativeUrl.Substring($webRelativeUrl.Length)
            }
            $libraries.Add([pscustomobject]@{
                    Title = [string]$list.Title
                    SiteRelativeUrl = $siteRelativeUrl.Trim('/')
                    ItemCount = [int]$list.ItemCount
                })
        }
        return @($libraries | Sort-Object Title)
    }
    catch {
        Add-RunFailure -FilePath '' -Action 'List document libraries' -Reason $_.Exception.Message
        Write-RunLog -Severity INFO -Action 'List document libraries guidance' -Result 'Access denied here usually means the account cannot open this site, or the sign-in did not obtain SharePoint AllSites.Read. That permission is user-consentable, so accept the consent prompt at sign-in; run Get-PnPAccessToken -Scopes to see what the session actually holds.'
        return $null
    }
}

function Get-SharePointTenantId {
    <# .SYNOPSIS Reads the tenant GUID from SharePoint, distinguishing a missing site from an omitted realm. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SiteUrl)

    $script:LastSharePointTenantLookupStatus = 'InvalidUrl'
    try {
        $siteUri = [uri]$SiteUrl
        if (-not $siteUri.IsAbsoluteUri -or $siteUri.Scheme -ne 'https') { return '' }
        $tenantRoot = $siteUri.GetLeftPart([System.UriPartial]::Authority)
    }
    catch { return '' }

    $probeRealm = {
        param([Parameter(Mandatory)][string]$Uri)

        $response = $null
        try {
            $response = Invoke-WebRequest -Uri $Uri -Method Get -Headers @{ Authorization = 'Bearer' } -UseBasicParsing -MaximumRedirection 0 -TimeoutSec 20 -ErrorAction Stop
        }
        catch {
            $responseProperty = $_.Exception.PSObject.Properties['Response']
            if ($responseProperty) { $response = $responseProperty.Value }
        }
        if ($null -eq $response) { return [pscustomobject]@{StatusCode = 0; TenantId = ''} }

        $statusCode = 0
        try { $statusCode = [int]$response.StatusCode }
        catch { Write-Verbose "Could not read the SharePoint response status: $($_.Exception.Message)" }

        $header = ''
        try {
            # PowerShell 7 returns HttpResponseHeaders, which has no string indexer; Windows PowerShell returns WebHeaderCollection, which has no TryGetValues.
            if ($response.Headers.PSObject.Methods['TryGetValues']) {
                $values = $null
                if ($response.Headers.TryGetValues('WWW-Authenticate', [ref]$values)) { $header = ($values -join ',') }
            }
            else {
                $header = [string]$response.Headers['WWW-Authenticate']
            }
        }
        catch { Write-Verbose "Could not read the SharePoint authentication header: $($_.Exception.Message)" }
        $realmMatch = [regex]::Match($header, '(?i)\brealm\s*=\s*"?(?<value>[^",\s]+)')
        $tenantId = [guid]::Empty
        if ($realmMatch.Success -and [guid]::TryParse($realmMatch.Groups['value'].Value, [ref]$tenantId) -and $tenantId -ne [guid]::Empty) {
            return [pscustomobject]@{StatusCode = $statusCode; TenantId = $tenantId.ToString()}
        }
        return [pscustomobject]@{StatusCode = $statusCode; TenantId = ''}
    }

    $siteEndpoint = "$($SiteUrl.TrimEnd('/'))/_vti_bin/client.svc/"
    $siteProbe = & $probeRealm -Uri $siteEndpoint
    if (-not [string]::IsNullOrWhiteSpace($siteProbe.TenantId)) {
        $script:LastSharePointTenantLookupStatus = 'Ready'
        return $siteProbe.TenantId
    }
    if ($siteProbe.StatusCode -eq 404) {
        $script:LastSharePointTenantLookupStatus = 'SiteNotFound'
        return ''
    }

    $rootEndpoint = "$tenantRoot/_vti_bin/client.svc/"
    if ([string]::Equals($siteEndpoint, $rootEndpoint, [System.StringComparison]::OrdinalIgnoreCase)) {
        $script:LastSharePointTenantLookupStatus = 'RealmMissing'
        return ''
    }
    $rootProbe = & $probeRealm -Uri $rootEndpoint
    if (-not [string]::IsNullOrWhiteSpace($rootProbe.TenantId)) {
        $script:LastSharePointTenantLookupStatus = 'ReadyFromTenantRoot'
        return $rootProbe.TenantId
    }
    $script:LastSharePointTenantLookupStatus = 'RealmMissing'
    return ''
}

function Save-LabelingClientId {
    <# .SYNOPSIS Remembers a non-secret application ID against the tenant it belongs to. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][AllowEmptyString()][string]$TenantId
    )

    $variableName = Get-TenantScopedClientIdName -TenantId $TenantId
    $null = $script:SessionSavedClientIds.Add($ClientId)
    if ([string]::IsNullOrWhiteSpace($variableName)) {
        Write-RunLog -Severity WARN -Action 'Save application ID' -Result 'The tenant could not be identified, so the application ID was not remembered. It still works for this run.'
        return
    }
    [Environment]::SetEnvironmentVariable($variableName, $ClientId, [EnvironmentVariableTarget]::Process)
    try {
        [Environment]::SetEnvironmentVariable($variableName, $ClientId, [EnvironmentVariableTarget]::User)
        Write-RunLog -Severity SUCCESS -Action 'Save application ID' -Result "Saved the non-secret client ID in user environment variable $variableName, which is specific to tenant $TenantId."
    }
    catch {
        Write-RunLog -Severity WARN -Action 'Save application ID' -Result "Could not persist ${variableName}: $($_.Exception.Message)"
    }
}

function Block-LabelingClientId {
    <# .SYNOPSIS Stops proposing an application ID the tenant does not recognise. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][AllowEmptyString()][string]$TenantId,
        # Set when a pre-flight check, rather than a failed sign-in, is the only evidence against the application.
        [switch]$SessionOnly
    )

    # The session list is what actually stops the prompt, because the ID may sit in a variable this utility must not modify.
    $null = $script:RejectedClientIds.Add($ClientId)
    if ($SessionOnly) {
        Write-RunLog -Severity INFO -Action 'Forget application ID' -Result "Application $ClientId will not be proposed again in this session. The saved value was kept, because a check made before signing in can also fail while a new registration is still replicating."
        return
    }
    $ownedNames = [System.Collections.Generic.List[string]]::new()
    $ownedNames.Add('PURVIEW_FILE_LABELING_CLIENT_ID')
    $tenantScopedName = Get-TenantScopedClientIdName -TenantId $TenantId
    if (-not [string]::IsNullOrWhiteSpace($tenantScopedName)) { $ownedNames.Add($tenantScopedName) }

    foreach ($variableName in @($ownedNames) + @('ENTRAID_APP_ID', 'ENTRAID_CLIENT_ID', 'AZURE_CLIENT_ID')) {
        foreach ($target in [EnvironmentVariableTarget]::Process, [EnvironmentVariableTarget]::User) {
            try {
                if ([Environment]::GetEnvironmentVariable($variableName, $target) -ne $ClientId) { continue }
                if (-not $ownedNames.Contains($variableName)) {
                    Write-RunLog -Severity WARN -Action 'Forget application ID' -Result "$variableName ($target scope) also holds $ClientId. Another tool owns that variable, so this utility left it alone; remove it yourself if nothing else needs it."
                    continue
                }
                [Environment]::SetEnvironmentVariable($variableName, $null, $target)
                Write-RunLog -Severity WARN -Action 'Forget application ID' -Result "Removed $ClientId from $variableName ($target scope)."
            }
            catch { Write-Verbose "Could not clear ${variableName}: $($_.Exception.Message)" }
        }
    }
    Write-RunLog -Severity INFO -Action 'Forget application ID' -Result "Application $ClientId will not be proposed again in this session."
}

function Resolve-ExistingApplicationId {
    <# .SYNOPSIS Tries to read back an application the tenant already has under this name. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DisplayName)

    if (-not (Get-Command Get-PnPEntraIDApp -ErrorAction SilentlyContinue)) { return '' }
    try {
        $existing = @(Get-PnPEntraIDApp -Identity $DisplayName -ErrorAction Stop) | Select-Object -First 1
        return Get-ObjectPropertyValue -InputObject $existing -Names 'AppId', 'ClientId', 'Id'
    }
    catch {
        # Reading the directory needs Graph Application.Read.All, which a SharePoint sign-in does not carry.
        Write-Verbose "Could not look up '$DisplayName': $($_.Exception.Message)"
        return ''
    }
}

function Read-DuplicateApplicationChoice {
    <# .SYNOPSIS Offers a way forward when the tenant already has an application with this name. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DisplayName)

    Write-RunLog -Severity WARN -Action 'Register application' -Result "The tenant already has an application named '$DisplayName', and reading it back needs Graph Application.Read.All, which this sign-in does not have."
    $choice = Read-MenuChoice -Title 'How should that be resolved?' -Options ([ordered]@{
            '1' = 'Register under a new, unique name'
            '2' = 'Enter the existing application ID (copy it from the Entra admin center)'
            '3' = 'Cancel'
        }) -Default '1'
    if ($choice -eq '3') { return $null }
    if ($choice -eq '2') {
        Write-Host ''
        Write-Host '  Entra admin center > Identity > Applications > App registrations, then copy' -ForegroundColor Gray
        Write-Host "  the Application (client) ID of '$DisplayName'." -ForegroundColor Gray
        $typed = Read-ValueWithDefault -Prompt 'Application (client) ID' -Default ''
        $parsed = [guid]::Empty
        if ([guid]::TryParse($typed, [ref]$parsed) -and $parsed -ne [guid]::Empty) {
            return [pscustomobject]@{Action = 'Existing'; Value = $parsed.ToString()}
        }
        Write-RunLog -Severity WARN -Action 'Validate application ID' -Result "'$typed' is not a valid GUID."
        return $null
    }
    return [pscustomobject]@{Action = 'Rename'; Value = "$DisplayName $(Get-Date -Format $script:TimestampFormat)"}
}

function Invoke-GraphAction {
    <# .SYNOPSIS Runs one Microsoft Graph operation in a separate process, because its assemblies break PnP in a shared one. #>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'resultLine', Justification = 'Assigned inside a ForEach-Object block, which runs in this scope.')]
    param(
        [Parameter(Mandatory)][ValidateSet('Check', 'Consent', 'CreateApp', 'RemoveApp', 'SignOut')][string]$Action,
        [Parameter(Mandatory)][AllowEmptyString()][string]$TenantId,
        [hashtable]$Arguments = @{},
        [string[]]$Scopes = @('Application.ReadWrite.All'),
        # Set when the call is only worth making if a sign-in already exists, so nothing prompts unexpectedly.
        [switch]$NoPrompt
    )

    $hostPath = ''
    try {
        $current = [string](Get-Process -Id $PID -ErrorAction Stop).Path
        if (-not [string]::IsNullOrWhiteSpace($current) -and $current -match '(?i)pwsh(\.exe)?$') { $hostPath = $current }
    }
    catch { Write-Verbose "Could not read this host's path: $($_.Exception.Message)" }
    if ([string]::IsNullOrWhiteSpace($hostPath)) { $hostPath = Get-PowerShell7Path -MinimumVersion ([version]'7.2.0') }
    if ([string]::IsNullOrWhiteSpace($hostPath)) {
        return [pscustomobject]@{Status = 'Error'; Message = 'PowerShell 7 is needed to talk to Microsoft Graph separately from PnP.'; Value = ''}
    }

    $workerPath = Join-Path ([System.IO.Path]::GetTempPath()) ('PurviewFileLabelingGraph-' + $PID + '-' + [guid]::NewGuid().ToString('N') + '.ps1')
    $worker = @'
param([string]$TenantId, [string]$Action, [string]$ArgumentJson, [string]$ScopeList, [switch]$NoPrompt, [switch]$UseDeviceCode)
$ErrorActionPreference = 'Stop'
$result = [ordered]@{ Status = 'Error'; Message = ''; Value = ''; Count = 0; OpenedSession = $false }
# The status line alone says nothing; Graph puts the reason in ErrorDetails.
function Get-Detail { param($Record)
    $text = "$($Record.Exception.Message)"
    if ($Record.ErrorDetails -and $Record.ErrorDetails.Message) { $text = "$($Record.ErrorDetails.Message) $text" }
    return $text
}
try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    $arguments = if ([string]::IsNullOrWhiteSpace($ArgumentJson)) { @{} } else { ConvertFrom-Json $ArgumentJson -AsHashtable }
    $context = Get-MgContext -ErrorAction SilentlyContinue
    $sameTenant = $context -and ([string]::IsNullOrWhiteSpace($TenantId) -or [string]$context.TenantId -eq $TenantId)
    $requestedScopes = @(($ScopeList -split ',') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $missingScopes = @($requestedScopes | Where-Object { $null -eq $context -or @($context.Scopes) -notcontains $_ })
    if ((-not $sameTenant -or $missingScopes.Count -gt 0) -and $Action -ne 'SignOut') {
        if ($NoPrompt) {
            $result.Status = 'NoSession'
            throw 'No Microsoft Graph sign-in with the required tenant and scopes exists yet.'
        }
        $connect = @{ Scopes = $requestedScopes; NoWelcome = $true; ErrorAction = 'Stop' }
        if ((Get-Command Connect-MgGraph).Parameters.ContainsKey('ContextScope')) { $connect.ContextScope = 'CurrentUser' }
        if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $connect.TenantId = $TenantId }
        if ($UseDeviceCode) { $connect.UseDeviceCode = $true }
        Connect-MgGraph @connect
        $result.OpenedSession = $true
    }

    function Get-Value { param($Object, [string]$Name)
        if ($null -eq $Object) { return $null }
        if ($Object -is [System.Collections.IDictionary]) { if ($Object.Contains($Name)) { return $Object[$Name] } return $null }
        $p = $Object.PSObject.Properties[$Name]; if ($p) { return $p.Value } return $null
    }
    function Get-Principal { param([string]$AppId)
        try { return Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals(appId='$AppId')" -ErrorAction Stop }
        catch { return $null }
    }

    switch ($Action) {
        'Check' {
            $collection = [string]$arguments['Collection']
            if ([string]::IsNullOrWhiteSpace($collection)) { $collection = 'applications' }
            try {
                $null = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/$collection(appId='$($arguments['ClientId'])')" -ErrorAction Stop
                $result.Status = 'Present'
            }
            catch {
                $detail = Get-Detail $_
                if ($detail -match '(?i)Request_ResourceNotFound|ResourceNotFound|does not exist|not found') { $result.Status = 'Missing' }
                else { $result.Status = 'Unknown'; $result.Message = $detail }
            }
        }
        'CreateApp' {
            $required = @()
            foreach ($pair in @(@{ R = '00000003-0000-0000-c000-000000000000'; S = 'Files.Read.All' }, @{ R = '00000003-0000-0ff1-ce00-000000000000'; S = 'AllSites.Read' })) {
                $sp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals(appId='$($pair.R)')" -ErrorAction Stop
                $id = ''
                foreach ($scope in @(Get-Value $sp 'oauth2PermissionScopes')) {
                    if ([string](Get-Value $scope 'value') -eq $pair.S) { $id = [string](Get-Value $scope 'id'); break }
                }
                if ([string]::IsNullOrWhiteSpace($id)) { throw "The resource $($pair.R) does not publish a delegated permission called '$($pair.S)'." }
                $required += @{ resourceAppId = $pair.R; resourceAccess = @(@{ id = $id; type = 'Scope' }) }
            }
            $body = @{
                displayName = [string]$arguments['DisplayName']
                signInAudience = 'AzureADMyOrg'
                isFallbackPublicClient = $true
                publicClient = @{ redirectUris = @('http://localhost') }
                requiredResourceAccess = $required
            } | ConvertTo-Json -Depth 8
            $created = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/applications' -Body $body -ContentType 'application/json' -ErrorAction Stop
            $appId = [string](Get-Value $created 'appId')
            if ([string]::IsNullOrWhiteSpace($appId)) { throw 'Graph created the application but returned no application ID.' }
            try { $null = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals' -Body (@{ appId = $appId } | ConvertTo-Json) -ContentType 'application/json' -ErrorAction Stop }
            catch { if ((Get-Detail $_) -notmatch '(?i)already exists|samekeyvalue') { throw } }
            $result.Status = 'Created'; $result.Value = $appId
        }
        'Consent' {
            $clientId = [string]$arguments['ClientId']
            $application = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/applications(appId='$clientId')" -ErrorAction Stop
            $applicationObjectId = [string](Get-Value $application 'id')
            if ([string]::IsNullOrWhiteSpace($applicationObjectId)) { throw 'The application registration has no directory object ID.' }

            # Register-PnPEntraIDApp can create an ownerless registration. Microsoft requires the
            # billing operator to own the application, so make this administrator an explicit owner.
            $operator = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/me' -ErrorAction Stop
            $operatorId = [string](Get-Value $operator 'id')
            if ([string]::IsNullOrWhiteSpace($operatorId)) { throw 'Microsoft Graph did not return the signed-in administrator object ID.' }
            $ownerResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/applications/$applicationObjectId/owners" -ErrorAction Stop
            $ownerIds = @(@(Get-Value $ownerResponse 'value') | ForEach-Object { [string](Get-Value $_ 'id') })
            if ($ownerIds -notcontains $operatorId) {
                $ownerBody = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$operatorId" } | ConvertTo-Json
                try {
                    $null = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/applications/$applicationObjectId/owners/`$ref" -Body $ownerBody -ContentType 'application/json' -ErrorAction Stop
                }
                catch {
                    if ((Get-Detail $_) -notmatch '(?i)already exist|added object references') { throw }
                }
                $ownerResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/applications/$applicationObjectId/owners" -ErrorAction Stop
                $ownerIds = @(@(Get-Value $ownerResponse 'value') | ForEach-Object { [string](Get-Value $_ 'id') })
                if ($ownerIds -notcontains $operatorId) { throw "Application owner assignment for object $operatorId could not be verified." }
            }
            $operatorName = [string](Get-Value $operator 'userPrincipalName')
            if ([string]::IsNullOrWhiteSpace($operatorName)) { $operatorName = [string](Get-Value $operator 'displayName') }
            if ([string]::IsNullOrWhiteSpace($operatorName)) { $operatorName = $operatorId }

            $principal = Get-Principal -AppId $clientId
            $principalId = [string](Get-Value $principal 'id')
            if ([string]::IsNullOrWhiteSpace($principalId)) {
                try {
                    $made = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals' -Body (@{ appId = $clientId } | ConvertTo-Json) -ContentType 'application/json' -ErrorAction Stop
                    $principalId = [string](Get-Value $made 'id')
                }
                catch {
                    # A principal that already exists but could not be read a moment ago is not an error.
                    if ((Get-Detail $_) -notmatch '(?i)already exists|samekeyvalue') { throw }
                    $principal = Get-Principal -AppId $clientId
                    $principalId = [string](Get-Value $principal 'id')
                }
            }
            if ([string]::IsNullOrWhiteSpace($principalId)) { throw 'The service principal could not be created or read back.' }
            $granted = 0
            $requested = 0
            foreach ($resource in @(Get-Value $application 'requiredResourceAccess')) {
                $resourceAppId = [string](Get-Value $resource 'resourceAppId')
                $roles = @(@(Get-Value $resource 'resourceAccess') | Where-Object { [string](Get-Value $_ 'type') -eq 'Role' })
                if ($roles.Count -eq 0 -or [string]::IsNullOrWhiteSpace($resourceAppId)) { continue }
                $requested += $roles.Count
                $resourcePrincipal = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals(appId='$resourceAppId')" -ErrorAction Stop
                foreach ($role in $roles) {
                    $body = @{ principalId = $principalId; resourceId = [string](Get-Value $resourcePrincipal 'id'); appRoleId = [string](Get-Value $role 'id') } | ConvertTo-Json
                    try { $null = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$principalId/appRoleAssignments" -Body $body -ContentType 'application/json' -ErrorAction Stop; $granted++ }
                    catch { if ((Get-Detail $_) -match '(?i)already exists|Permission being assigned') { $granted++ } else { throw } }
                }
            }
            $result.Status = 'Granted'; $result.Count = $granted; $result.Value = $operatorName
            if ($requested -eq 0) { $result.Message = 'The registration requests no application permissions, so there was nothing to consent to.' }
        }
        'RemoveApp' {
            $clientId = [string]$arguments['ClientId']
            $stripped = $false
            $principal = Get-Principal -AppId $clientId
            $principalId = [string](Get-Value $principal 'id')
            if (-not [string]::IsNullOrWhiteSpace($principalId)) {
                $null = Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$principalId" -Body (@{ accountEnabled = $false } | ConvertTo-Json) -ContentType 'application/json' -ErrorAction Stop
                $stripped = $true
            }
            $application = $null
            try { $application = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/applications(appId='$clientId')" -ErrorAction Stop } catch { }
            $objectId = [string](Get-Value $application 'id')
            if (-not [string]::IsNullOrWhiteSpace($objectId)) {
                $null = Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/applications/$objectId" -Body (@{ requiredResourceAccess = @() } | ConvertTo-Json -Depth 4) -ContentType 'application/json' -ErrorAction Stop
                $null = Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/applications/$objectId" -ErrorAction Stop
                $result.Status = 'Deleted'
            }
            else { $result.Status = if ($stripped) { 'Disabled' } else { 'Missing' } }
        }
        'SignOut' {
            if ($context) { $null = Disconnect-MgGraph -ErrorAction SilentlyContinue -WarningAction SilentlyContinue }
            $result.Status = 'SignedOut'
        }
    }
}
catch {
    if ($result.Status -eq 'Error') { $result.Message = Get-Detail $_ }
    if ($result.Status -eq 'NoSession') { $result.Message = 'No Microsoft Graph sign-in exists yet.' }
}
'RESULT:' + ($result | ConvertTo-Json -Compress -Depth 6)
'@

    try {
        Set-Content -LiteralPath $workerPath -Value $worker -Encoding utf8 -ErrorAction Stop
        $argumentJson = if ($Arguments.Count -gt 0) { $Arguments | ConvertTo-Json -Compress -Depth 6 } else { '' }
        $workerArguments = @('-NoLogo', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $workerPath,
            '-TenantId', $TenantId, '-Action', $Action, '-ArgumentJson', $argumentJson, '-ScopeList', ($Scopes -join ','))
        if ($NoPrompt) { $workerArguments += '-NoPrompt' }
        if ($script:UseDeviceCode) { $workerArguments += '-UseDeviceCode' }

        $resultLine = ''
        $other = [System.Collections.Generic.List[string]]::new()
        & $hostPath @workerArguments 2>&1 | ForEach-Object {
            $text = "$_"
            if ($text -like 'RESULT:*') { $resultLine = $text; return }
            if ($text -like 'PROGRESS:*') { Write-Host ('    ' + $text.Substring(9)) -ForegroundColor DarkGray; return }
            $other.Add($text)
        }
        if ([string]::IsNullOrWhiteSpace($resultLine)) {
            return [pscustomobject]@{Status = 'Error'; Message = (@($other) -join ' '); Value = ''; Count = 0; OpenedSession = $false}
        }
        $response = $resultLine.Substring(7) | ConvertFrom-Json
        if ($null -ne $response -and $null -ne $response.PSObject.Properties['OpenedSession'] -and [bool]$response.OpenedSession) {
            $script:GraphSessionOpened = $true
        }
        return $response
    }
    catch {
        return [pscustomobject]@{Status = 'Error'; Message = (Get-ErrorText -ErrorRecord $_); Value = ''; Count = 0; OpenedSession = $false}
    }
    finally { Remove-Item -LiteralPath $workerPath -Force -ErrorAction SilentlyContinue }
}

function New-LabelingApplicationViaGraph {
    <# .SYNOPSIS Creates the read-only sign-in application through Graph, which is visible and cannot stall on a hidden window. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ApplicationName
    )

    Write-RunLog -Severity INFO -Action 'Register application' -Result 'Creating it through Microsoft Graph, in a separate process so its assemblies cannot disturb PnP.'
    $response = Invoke-GraphAction -Action 'CreateApp' -TenantId $TenantId -Arguments @{DisplayName = $ApplicationName} -Scopes @('Application.ReadWrite.All')
    if ($response.Status -eq 'Created' -and -not [string]::IsNullOrWhiteSpace($response.Value)) {
        Write-RunLog -Severity SUCCESS -Action 'Register application' -Result "Created application $($response.Value) with delegated SharePoint AllSites.Read and Graph Files.Read.All."
        return [string]$response.Value
    }
    Write-RunLog -Severity WARN -Action 'Register application' -Result "Creating the application through Microsoft Graph did not work: $($response.Message)"
    return ''
}

function Show-RegistrationSignInNotice {
    <# .SYNOPSIS Says what to expect, because the registration cmdlet opens its own sign-in and otherwise looks like a hang. #>
    [CmdletBinding()]
    param([switch]$DeviceCode)

    Write-Host ''
    if ($DeviceCode) {
        Write-Host '  A device code appears below. Open the address it shows, enter the code, and' -ForegroundColor Cyan
        Write-Host '  sign in as an account allowed to create app registrations.' -ForegroundColor Cyan
    }
    else {
        Write-Host '  A browser window is opening for this registration. Windows often places it' -ForegroundColor Cyan
        Write-Host '  behind this one, so check the taskbar if nothing seems to happen.' -ForegroundColor Cyan
    }
    Write-Host '  You are asked to sign in, then to consent to the permissions being requested.' -ForegroundColor Gray
    Write-Host '  Registration then waits for Entra to finish, which usually takes under a minute' -ForegroundColor Gray
    Write-Host '  and can take longer. Nothing more is needed here until it reports back.' -ForegroundColor Gray
    Write-Host ''
}

function New-LabelingApplication {
    <# .SYNOPSIS Registers an Entra application carrying exactly the permissions this utility needs. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [AllowEmptyString()][string]$TenantId = '',
        [string]$ApplicationName = 'PnP PowerShell - Purview File Labeling'
    )

    try {
        $tenantId = if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $TenantId } else { Get-SharePointTenantId -SiteUrl $SiteUrl }
        if ([string]::IsNullOrWhiteSpace($tenantId)) {
            if ($script:LastSharePointTenantLookupStatus -eq 'SiteNotFound') {
                throw "SharePoint returned 404 for '$SiteUrl'. The site may have been deleted or the URL may be incorrect."
            }
            throw "SharePoint at '$SiteUrl' did not return a tenant realm, so the application cannot be registered against a verified tenant."
        }

        Write-RunLog -Severity INFO -Action 'Register application' -Result "Registering '$ApplicationName' in tenant $tenantId with delegated SharePoint AllSites.Read and Microsoft Graph Files.Read.All."
        Write-RunLog -Severity INFO -Action 'Register application' -Result 'Sign in as an account allowed to create app registrations. Both permissions are user-consentable, so an administrator is only needed if tenant policy restricts user consent.'

        # Graph is tried first because it reports progress here, rather than waiting on a window this utility cannot see.
        $graphClientId = New-LabelingApplicationViaGraph -TenantId $tenantId -ApplicationName $ApplicationName
        if (-not [string]::IsNullOrWhiteSpace($graphClientId)) {
            Save-LabelingClientId -ClientId $graphClientId -TenantId $tenantId
            return $graphClientId
        }

        $command = Get-Command Register-PnPEntraIDAppForInteractiveLogin -ErrorAction SilentlyContinue
        if (-not $command) {
            throw 'Microsoft Graph could not create the application, and this PnP.PowerShell version cannot register one either. Update it with: Update-Module PnP.PowerShell'
        }
        foreach ($parameterName in 'SharePointDelegatePermissions', 'GraphDelegatePermissions', 'Tenant') {
            if (-not $command.Parameters.ContainsKey($parameterName)) {
                throw "Register-PnPEntraIDAppForInteractiveLogin does not support -$parameterName, so the permission set cannot be controlled. Registration was not attempted."
            }
        }
        Write-RunLog -Severity INFO -Action 'Register application' -Result 'Falling back to PnP.PowerShell to register it instead.'

        $registrationName = $ApplicationName
        $output = @()
        $useDeviceLogin = $script:UseDeviceCode -and $command.Parameters.ContainsKey('DeviceLogin')
        Show-RegistrationSignInNotice -DeviceCode:$useDeviceLogin
        while ($true) {
            try {
                $parameters = @{
                    ApplicationName = $registrationName
                    Tenant = $tenantId
                    SharePointDelegatePermissions = 'AllSites.Read'
                    GraphDelegatePermissions = 'Files.Read.All'
                    ErrorAction = 'Stop'
                }
                if ($useDeviceLogin) { $parameters.DeviceLogin = $true }
                $output = @(Register-PnPEntraIDAppForInteractiveLogin @parameters)
                break
            }
            catch {
                if ((Get-ErrorText -ErrorRecord $_) -notmatch '(?i)already exists') { throw }
                $reusableId = Resolve-ExistingApplicationId -DisplayName $registrationName
                if (-not [string]::IsNullOrWhiteSpace($reusableId)) {
                    Write-RunLog -Severity SUCCESS -Action 'Register application' -Result "Reusing the existing application '$registrationName', $reusableId."
                    Save-LabelingClientId -ClientId $reusableId -TenantId $tenantId
                    return $reusableId
                }
                $decision = Read-DuplicateApplicationChoice -DisplayName $registrationName
                if ($null -eq $decision) { throw "Registration stopped because '$registrationName' already exists in tenant $tenantId." }
                if ($decision.Action -eq 'Existing') {
                    Save-LabelingClientId -ClientId $decision.Value -TenantId $tenantId
                    return $decision.Value
                }
                $registrationName = $decision.Value
                Write-RunLog -Severity INFO -Action 'Register application' -Result "Retrying registration as '$registrationName'."
            }
        }

        foreach ($item in $output) {
            if ($null -eq $item) { continue }
            foreach ($propertyName in 'AzureAppId/ClientId', 'ClientId', 'AppId', 'ApplicationId') {
                $value = Get-ObjectPropertyValue -InputObject $item -Names @($propertyName)
                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    Write-RunLog -Severity SUCCESS -Action 'Register application' -Result "Registered application $value. Remove it in the Entra admin center when this utility is no longer needed."
                    Save-LabelingClientId -ClientId $value -TenantId $tenantId
                    return $value
                }
            }
        }
        $guidMatch = [regex]::Match([string]$output, '(?i)(?<![0-9a-f])[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}(?![0-9a-f])')
        if ($guidMatch.Success) {
            Write-RunLog -Severity SUCCESS -Action 'Register application' -Result "Registered application $($guidMatch.Value). Remove it in the Entra admin center when this utility is no longer needed."
            Save-LabelingClientId -ClientId $guidMatch.Value -TenantId $tenantId
            return $guidMatch.Value
        }
        throw 'Registration completed but did not return a recognizable application client ID.'
    }
    catch {
        Add-RunFailure -FilePath '' -Action 'Register application' -Reason $_.Exception.Message
        return ''
    }
}

function Get-SharePointLoginAuthority {
    <# .SYNOPSIS Returns the sign-in authority matching a SharePoint host's cloud. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SiteUrl)

    $sharePointHost = ''
    try { $sharePointHost = ([uri]$SiteUrl).Host.ToLowerInvariant() }
    catch { return 'https://login.microsoftonline.com' }
    switch ($sharePointHost.Substring($sharePointHost.LastIndexOf('.') + 1)) {
        'us' { 'https://login.microsoftonline.us' }
        'de' { 'https://login.microsoftonline.de' }
        'cn' { 'https://login.partner.microsoftonline.cn' }
        default { 'https://login.microsoftonline.com' }
    }
}

function Test-ApplicationPresence {
    <# .SYNOPSIS Asks Entra whether an application exists in a tenant, which needs no sign-in and no permission. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][AllowEmptyString()][string]$TenantId,
        [Parameter(Mandatory)][string]$SiteUrl
    )

    if ([string]::IsNullOrWhiteSpace($TenantId)) { return 'Unknown' }
    $authority = Get-SharePointLoginAuthority -SiteUrl $SiteUrl
    # The SharePoint resource is the tenant host; a site path is not a resource, and asking for one fails for the wrong reason.
    $resource = ''
    try { $resource = ([uri]$SiteUrl).GetLeftPart([UriPartial]::Authority) }
    catch { return 'Unknown' }
    # The device-code endpoint answers AADSTS700016 for an application the tenant does not have, before any consent is involved.
    $body = @{
        client_id = $ClientId
        scope = "openid offline_access $resource/AllSites.Read"
    }
    try {
        $null = Invoke-RestMethod -Uri "$authority/$TenantId/oauth2/v2.0/devicecode" -Method Post -Body $body `
            -ContentType 'application/x-www-form-urlencoded' -TimeoutSec 20 -ErrorAction Stop
        return 'Present'
    }
    catch {
        $raw = if ($null -ne $_.ErrorDetails -and -not [string]::IsNullOrWhiteSpace($_.ErrorDetails.Message)) {
            [string]$_.ErrorDetails.Message
        }
        else {
            Get-ErrorText -ErrorRecord $_
        }
        if ($raw -match '(?i)700016|was not found in the directory') { return 'Missing' }
        return 'Unknown'
    }
}

function Test-ApplicationInDirectory {
    <# .SYNOPSIS Asks the directory itself whether an object exists, without prompting for a sign-in that does not already exist. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][AllowEmptyString()][string]$TenantId,
        [ValidateSet('applications', 'servicePrincipals')][string]$Collection = 'applications'
    )

    $response = Invoke-GraphAction -Action 'Check' -TenantId $TenantId -NoPrompt `
        -Arguments @{ClientId = $ClientId; Collection = $Collection} -Scopes @('Application.Read.All')
    if ($response.Status -in 'Present', 'Missing') { return [string]$response.Status }
    return 'Unknown'
}

function Read-SharePointApplication {
    <# .SYNOPSIS Chooses, or registers, the Entra application used to sign in to one specific tenant. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [Parameter(Mandatory)][AllowEmptyString()][string]$TenantId
    )

    while ($true) {
        $saved = Get-ConfiguredClientId -TenantId $TenantId
        if ($saved) {
            # A remembered ID is worth nothing if the tenant no longer has that application, so check before proposing it.
            $state = Test-ApplicationPresence -ClientId $saved.ClientId -TenantId $TenantId -SiteUrl $SiteUrl
            if ($state -eq 'Missing') {
                $directory = Test-ApplicationInDirectory -ClientId $saved.ClientId -TenantId $TenantId
                if ($directory -eq 'Present') {
                    # The registration is there but has no service principal yet, and an ordinary sign-in creates one.
                    Write-RunLog -Severity WARN -Action 'Check application' -Result "Application $($saved.ClientId) exists in tenant $TenantId but has never been consented to. Signing in with it should complete that."
                }
                elseif ($directory -eq 'Missing' -or -not $script:SessionSavedClientIds.Contains($saved.ClientId)) {
                    $evidence = if ($directory -eq 'Missing') { 'The directory confirms it does not exist' } else { 'Tenant {0} does not recognise it' -f $TenantId }
                    Write-RunLog -Severity WARN -Action 'Check application' -Result "$evidence, so application $($saved.ClientId) was forgotten and will not be proposed again."
                    Block-LabelingClientId -ClientId $saved.ClientId -TenantId $TenantId
                    continue
                }
                else {
                    Write-RunLog -Severity WARN -Action 'Check application' -Result "Application $($saved.ClientId) was registered moments ago and is not visible yet, so it was kept but is not offered this time."
                    Block-LabelingClientId -ClientId $saved.ClientId -TenantId $TenantId -SessionOnly
                    continue
                }
            }
            if ($state -eq 'Present') {
                Write-RunLog -Severity SUCCESS -Action 'Check application' -Result "Application $($saved.ClientId) exists in tenant $TenantId."
            }
        }
        $options = [ordered]@{}
        if ($saved) {
            $origin = if ($saved.IsTenantScoped) {
                "remembered for tenant $TenantId"
            }
            else {
                "from $($saved.VariableName), which is not tenant-specific and may belong to another tenant"
            }
            $options['1'] = "Use application $($saved.ClientId)  ($origin)"
            $options['2'] = "Forget $($saved.ClientId)"
        }
        $options[[string]($options.Count + 1)] = 'Register a new application for this tenant'
        $options[[string]($options.Count + 1)] = 'Enter a different application ID'
        $options[[string]($options.Count + 1)] = 'Return to the main menu'
        # An application saved for another tenant is a likely mismatch, so registering is the safer default there.
        $default = if ($saved -and $saved.IsTenantScoped) { '1' } elseif ($saved) { '3' } else { '1' }
        $choice = $options[(Read-MenuChoice -Title "Choose the Entra application for signing in to tenant $TenantId." -Options $options -Default $default)]

        if ($choice -eq 'Return to the main menu') { return $null }

        if ($choice -like 'Use application*') {
            return [pscustomobject]@{ClientId = $saved.ClientId; JustRegistered = $false}
        }
        if ($choice -like 'Forget *') {
            Block-LabelingClientId -ClientId $saved.ClientId -TenantId $TenantId
            continue
        }
        if ($choice -eq 'Register a new application for this tenant') {
            $newClientId = New-LabelingApplication -SiteUrl $SiteUrl -TenantId $TenantId
            if (-not [string]::IsNullOrWhiteSpace($newClientId)) {
                return [pscustomobject]@{ClientId = $newClientId; JustRegistered = $true}
            }
            continue
        }
        Write-Host ''
        Write-Host '  Copy the Application (client) ID from the Entra admin center, under' -ForegroundColor Gray
        Write-Host '  Identity > Applications > App registrations. It looks like a GUID.' -ForegroundColor Gray
        $typedClientId = Read-ValueWithDefault -Prompt 'Application (client) ID, for example 00000000-1111-2222-3333-444444444444' -Default ''
        $parsedClientId = [guid]::Empty
        if ([guid]::TryParse($typedClientId, [ref]$parsedClientId) -and $parsedClientId -ne [guid]::Empty) {
            return [pscustomobject]@{ClientId = $parsedClientId.ToString(); JustRegistered = $false}
        }
        Write-RunLog -Severity WARN -Action 'Validate application ID' -Result "'$typedClientId' is not a valid GUID."
    }
}

function Read-SharePointTarget {
    <# .SYNOPSIS Signs in to a site, then lets the operator choose a library and an optional subfolder. #>
    [CmdletBinding()]
    param([switch]$SkipSubfolder)

    while ($true) {
        Write-Host ''
        Write-Host '  Type back at any prompt below to return to the main menu.' -ForegroundColor DarkGray
        $siteUrl = Read-ValueWithDefault -Prompt 'Site URL (for example https://contoso.sharepoint.com/sites/LabelTest)' -Default (Get-SessionValue -Name 'PURVIEW_FILE_LABELING_SITE_URL')
        if ($siteUrl -eq 'back') { return $null }
        if ($siteUrl -notmatch '^https://[^/]+\.') {
            Write-RunLog -Severity WARN -Action 'Validate site URL' -Result 'Enter the full site URL, starting with https://, or type back to return to the main menu.'
            continue
        }
        $siteUrl = $siteUrl.TrimEnd('/')
        if (-not (Test-SharePointPrerequisite)) { return $null }

        # The tenant comes from SharePoint itself, so nothing remembered from a previous tenant can be reused here.
        $tenantId = Get-SharePointTenantId -SiteUrl $siteUrl
        if ([string]::IsNullOrWhiteSpace($tenantId)) {
            $reason = if ($script:LastSharePointTenantLookupStatus -eq 'SiteNotFound') {
                'SharePoint returned 404 for this site. It may have been deleted or the URL may be incorrect.'
            }
            else { 'SharePoint did not return a tenant realm for this site, so the sign-in tenant could not be verified.' }
            Add-RunFailure -FilePath $siteUrl -Action 'Resolve tenant' -Reason $reason
            if ((Read-MenuChoice -Title 'Try a different site URL?' -Options ([ordered]@{'1' = 'Yes'; '2' = 'No, return to the main menu'}) -Default '1') -eq '1') { continue }
            return $null
        }
        Write-RunLog -Severity INFO -Action 'Resolve tenant' -Result "SharePoint reports tenant $tenantId for $siteUrl."
        # The URL answered a real SharePoint tenant challenge, so it is worth proposing again even if sign-in later fails.
        Save-SessionValue -Name 'PURVIEW_FILE_LABELING_SITE_URL' -Value $siteUrl

        # Sign-in retries stay on this site, so a failed attempt never re-asks for the URL.
        $connected = $false
        $chooseAnotherSite = $false
        $application = $null
        while (-not $connected) {
            # A saved confidential client is preferred, because it is the only sign-in that can also write labels.
            $confidential = Test-ConfidentialClientReady -TenantId $tenantId
            if ($confidential -and (Test-ApplicationPresence -ClientId $confidential.ClientId -TenantId $tenantId -SiteUrl $siteUrl) -eq 'Missing') {
                # Registering creates the application; only consent creates the service principal that app-only sign-in needs.
                Write-RunLog -Severity WARN -Action 'Check confidential client' -Result "Confidential client $($confidential.ClientId) is not usable in tenant $tenantId, which almost always means administrator consent has not been granted."
                $consentChoice = Read-MenuChoice -Title 'Grant administrator consent now, so this run can apply labels?' -Options ([ordered]@{
                        '1' = 'Yes, sign in as an administrator and grant it from here'
                        '2' = 'No, continue read-only for this run'
                    }) -Default '1'
                if ($consentChoice -eq '1' -and (Grant-LabelingAdminConsent -ClientId $confidential.ClientId -TenantId $tenantId -Confirm:$false)) {
                    Write-RunLog -Severity INFO -Action 'Check confidential client' -Result 'Consent granted. Waiting a few seconds for it to replicate before signing in.'
                    Start-Sleep -Seconds 10
                }
                else {
                    # Consent may have opened a Graph session, which makes an authoritative check free at this point.
                    if ((Test-ApplicationInDirectory -ClientId $confidential.ClientId -TenantId $tenantId) -eq 'Missing') {
                        Write-RunLog -Severity WARN -Action 'Check confidential client' -Result "The directory confirms application $($confidential.ClientId) no longer exists, so the saved confidential client was forgotten. Choose 'Enable SharePoint Online metered label writing' to register a new one."
                        Clear-ConfidentialClientConfig
                    }
                    Write-RunLog -Severity WARN -Action 'Check confidential client' -Result 'Falling back to the read-only delegated sign-in for this run.'
                    $confidential = $null
                }
            }
            if ($confidential) {
                $clientId = $confidential.ClientId
                $connected = Connect-SharePointAppOnly -SiteUrl $siteUrl -Config $confidential
            }
            else {
                if ($null -eq $application) {
                    $application = Read-SharePointApplication -SiteUrl $siteUrl -TenantId $tenantId
                    if ($null -eq $application) { return $null }
                }
                $clientId = $application.ClientId
                $connected = Connect-SharePointSite -SiteUrl $siteUrl -ClientId $clientId -TenantId $tenantId -JustRegistered:$application.JustRegistered
            }
            if ($connected) { break }

            $retryChoice = Read-MenuChoice -Title "Sign-in to $siteUrl failed. What next?" -Options ([ordered]@{
                    '1' = 'Retry with device-code sign-in (lets you choose Microsoft Authenticator)'
                    '2' = 'Retry in the browser, forcing a fresh sign-in and the account picker'
                    '3' = 'Choose or register a different application'
                    '4' = 'Retry in the browser as before'
                    '5' = 'Enter a different site URL'
                    '6' = 'Return to the main menu'
                }) -Default '1'
            if ($retryChoice -eq '6') { return $null }
            if ($retryChoice -eq '5') { $chooseAnotherSite = $true; break }
            $script:UseDeviceCode = ($retryChoice -eq '1')
            $script:ForceFreshSignIn = ($retryChoice -eq '2')
            # Only a deliberate choice reopens the application prompt; every other retry keeps the same application.
            if ($retryChoice -eq '3') { $application = $null }
        }
        if ($chooseAnotherSite) { continue }

        $libraries = @(Get-SharePointDocumentLibrary)
        if ($libraries.Count -eq 0) {
            Write-RunLog -Severity WARN -Action 'List document libraries' -Result 'No document library on this site is available to this account.'
            return $null
        }
        $libraryOptions = [ordered]@{}
        for ($index = 0; $index -lt $libraries.Count; $index++) {
            $library = $libraries[$index]
            $libraryOptions[[string]($index + 1)] = '{0}  ({1} items, /{2})' -f $library.Title, $library.ItemCount, $library.SiteRelativeUrl
        }
        $rememberedLibrary = Get-SessionValue -Name 'PURVIEW_FILE_LABELING_LIBRARY'
        $libraryDefault = '1'
        for ($index = 0; $index -lt $libraries.Count; $index++) {
            if ($libraries[$index].Title -eq $rememberedLibrary) { $libraryDefault = [string]($index + 1); break }
        }
        $libraryChoice = Read-MenuChoice -Title 'Choose the document library to scan.' -Options $libraryOptions -Default $libraryDefault
        $selectedLibrary = $libraries[[int]$libraryChoice - 1]
        Save-SessionValue -Name 'PURVIEW_FILE_LABELING_LIBRARY' -Value $selectedLibrary.Title

        $subfolder = ''
        if (-not $SkipSubfolder) {
            $subfolder = (Read-ValueWithDefault -Prompt "Subfolder inside '$($selectedLibrary.Title)', or . for the whole library" -Default '.').Trim('/', '\')
            if ($subfolder -eq '.') { $subfolder = '' }
        }
        $targetPath = if ($subfolder) { "$($selectedLibrary.SiteRelativeUrl)/$($subfolder -replace '\\', '/')" } else { $selectedLibrary.SiteRelativeUrl }

        return [pscustomobject]@{
            SiteUrl = $siteUrl
            TenantId = $tenantId
            ClientId = $clientId
            LibraryTitle = $selectedLibrary.Title
            LibraryUrl = $selectedLibrary.SiteRelativeUrl
            Subfolder = $subfolder
            TargetPath = $targetPath
        }
    }
}

function Read-FileSource {
    <# .SYNOPSIS Selects either file-path labeling through the Purview client or native SharePoint Online labeling through Graph. #>
    [CmdletBinding()]
    param()

    while ($true) {
        $sourceChoice = Read-MenuChoice -Title 'Where are the files?' -Options ([ordered]@{
                '1' = 'Local/UNC path or mounted SharePoint Server library (Purview client)'
                '2' = 'SharePoint Online document library (metered Graph API for Apply)'
            }) -Default '1'

        if ($sourceChoice -ne '2') {
            # File-path source selected. This includes SharePoint Server content exposed through a mounted path.
            if (-not (Get-Module -ListAvailable -Name PurviewInformationProtection -ErrorAction SilentlyContinue)) {
                if (Install-PurviewClient -Confirm:$false) {
                    if ($script:RelaunchCompleted) { return 'Local' }
                }
            }
            return 'Local'
        }

        if (-not (Test-SharePointSourceAvailable)) {
            if ($script:RelaunchCompleted) { return '' }
            continue
        }
        if (-not (Install-SharePointModule -Confirm:$false)) {
            Write-RunLog -Severity WARN -Action 'Select source' -Result 'SharePoint Online requires PnP.PowerShell. The other source remains available.'
            continue
        }
        return 'SharePoint'
    }
}

function Read-RunMode {
    <# .SYNOPSIS Chooses dry run or apply, refusing SharePoint Online writes unless the metered API is configured. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('Local', 'SharePoint')][string]$Source,
        [AllowEmptyString()][string]$TenantId = ''
    )

    if ($Source -eq 'SharePoint') {
        $confidential = Test-ConfidentialClientReady -TenantId $TenantId
        # Apply needs the connection this run actually holds to be app-only, not merely a configured application.
        if ($confidential -and $script:ConnectedAppOnly) {
            Write-Host ''
            Write-Host "  Confidential client $($confidential.ClientId) is configured, so this source can" -ForegroundColor Yellow
            Write-Host '  call Graph assignSensitivityLabel. Each file is a metered API call billed to' -ForegroundColor Yellow
            Write-Host '  the linked Azure subscription. Reading and dry runs remain unmetered.' -ForegroundColor Yellow
            Write-RunLog -Severity INFO -Action 'Choose run mode' -Result "Apply is available through confidential client $($confidential.ClientId) and the metered Graph assignSensitivityLabel API."
            return (Read-MenuChoice -Title 'Choose run mode.' -Options ([ordered]@{
                        '1' = 'Dry run (default, no labels changed, nothing billed)'
                        '2' = 'Apply labels (metered per file)'
                    }) -Default '1') -eq '1'
        }

        Write-Host ''
        Write-Host '  This SharePoint Online connection is survey-only, so no label is written and' -ForegroundColor Yellow
        Write-Host '  nothing is billed. SPO writes use Graph assignSensitivityLabel, an advanced' -ForegroundColor Yellow
        Write-Host '  metered and protected API that accepts confidential clients only.' -ForegroundColor Yellow
        Write-Host ''
        Write-Host '  To enable it, register a certificate-based application with application' -ForegroundColor Gray
        Write-Host '  Files.ReadWrite.All, grant it administrator consent, then associate the' -ForegroundColor Gray
        Write-Host '  application with an Azure subscription by creating a' -ForegroundColor Gray
        Write-Host '  Microsoft.GraphServices/accounts resource. This utility guides that setup.' -ForegroundColor Gray
        Write-Host ''
        Write-Host '  Local paths, UNC shares, and file access to SharePoint Server use the separate' -ForegroundColor Cyan
        Write-Host '  Purview client path with Set-FileLabel; that path does not call this API.' -ForegroundColor Cyan
        Write-RunLog -Severity WARN -Action 'Choose run mode' -Result 'SharePoint Online source forced to dry run. Its write path is the metered, protected, confidential-client-only Graph assignSensitivityLabel API.'

        $setupChoice = Read-MenuChoice -Title 'What next?' -Options ([ordered]@{
                '1' = 'Continue with this dry run (default)'
                '2' = 'Switch to the local/UNC/SharePoint Server file path source'
                '3' = 'Register a confidential client and link Azure billing now'
            }) -Default '1'
        if ($setupChoice -eq '2') {
            $script:Source = 'Local'
            $script:SetupInterrupted = $true
            return $true
        }
        if ($setupChoice -eq '3') {
            $null = Invoke-MeteredSetup
            Write-Host ''
            if ($script:SetupInterrupted) {
                Write-Host '  Returning to the main menu because setup replaced this run''s SharePoint' -ForegroundColor Yellow
                Write-Host '  sign-in. Start a new run to connect as the application.' -ForegroundColor Yellow
            }
            else {
                Write-Host '  Setup did not complete, so this run continues as a dry run.' -ForegroundColor Yellow
            }
        }
        return $true
    }
    return (Read-MenuChoice -Title 'Choose run mode.' -Options ([ordered]@{
                '1' = 'Dry run (default, no labels changed)'
                '2' = 'Apply labels'
            }) -Default '1') -eq '1'
}

function Get-SyncedLibraryFolder {
    <# .SYNOPSIS Lists the SharePoint and OneDrive libraries the OneDrive client already syncs to this machine. #>
    [CmdletBinding()]
    param()

    $found = [System.Collections.Generic.List[object]]::new()
    $accountsKey = 'HKCU:\Software\Microsoft\OneDrive\Accounts'
    if (-not (Test-Path -LiteralPath $accountsKey)) { return @() }
    try {
        foreach ($account in @(Get-ChildItem -LiteralPath $accountsKey -ErrorAction Stop)) {
            $tenantsKey = Join-Path $account.PSPath 'Tenants'
            if (-not (Test-Path -LiteralPath $tenantsKey)) { continue }
            foreach ($tenant in @(Get-ChildItem -LiteralPath $tenantsKey -ErrorAction SilentlyContinue)) {
                # The client records each mounted library as a property whose name is its local path.
                foreach ($path in @((Get-Item -LiteralPath $tenant.PSPath -ErrorAction SilentlyContinue).Property)) {
                    if ([string]::IsNullOrWhiteSpace($path)) { continue }
                    if (-not (Test-Path -LiteralPath $path -PathType Container -ErrorAction SilentlyContinue)) { continue }
                    $found.Add([pscustomobject]@{Tenant = [string]$tenant.PSChildName; Path = [string]$path})
                }
            }
        }
    }
    catch { Write-Verbose "Could not read the OneDrive sync list: $($_.Exception.Message)" }
    return @($found | Sort-Object Path -Unique)
}

function Read-LocalFolder {
    <# .SYNOPSIS Chooses the folder to scan, offering any library OneDrive already syncs. #>
    [CmdletBinding()]
    param()

    $synced = @(Get-SyncedLibraryFolder)
    if ($synced.Count -eq 0) {
        Write-Host ''
        Write-Host '  Type the folder that holds the files to label, for example' -ForegroundColor Gray
        Write-Host '  C:\Data\Contracts or \\fileserver\Finance\Reports.' -ForegroundColor Gray
        return Read-ValueWithDefault -Prompt 'Folder to scan' -Default ''
    }

    Write-Host ''
    Write-Host '  OneDrive already syncs the libraries below. Labeling a synced copy uses the' -ForegroundColor Gray
    Write-Host '  Purview client, so it needs no metered API and no Azure billing; the client' -ForegroundColor Gray
    Write-Host '  uploads each change back to SharePoint.' -ForegroundColor Gray
    $options = [ordered]@{}
    for ($index = 0; $index -lt $synced.Count; $index++) {
        $options[[string]($index + 1)] = '{0}  ({1})' -f $synced[$index].Path, $synced[$index].Tenant
    }
    $options[[string]($synced.Count + 1)] = 'Type a different folder or UNC path'
    $choice = [int](Read-MenuChoice -Title 'Which folder holds the files to label?' -Options $options -Default '1')
    if ($choice -gt $synced.Count) { return Read-ValueWithDefault -Prompt 'Folder to scan' -Default '' }
    return [string]$synced[$choice - 1].Path
}

function Read-RunSetting {
    <# .SYNOPSIS Collects target, filter, recursion, label, logging, and mode settings. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$Labels,
        [Parameter(Mandatory)][ValidateSet('Local', 'SharePoint')][string]$Source
    )

    $targetPath = ''
    $siteUrl = ''
    $tenantId = ''
    $clientId = ''
    $libraryTitle = ''
    $libraryUrl = ''
    if ($Source -eq 'SharePoint') {
        $target = Read-SharePointTarget
        if ($null -eq $target) { return $null }
        $siteUrl = $target.SiteUrl
        $tenantId = $target.TenantId
        $clientId = $target.ClientId
        $libraryTitle = $target.LibraryTitle
        $libraryUrl = $target.LibraryUrl
        $targetPath = $target.TargetPath
    }
    else {
        $targetPath = Read-LocalFolder
    }

    $includeSubfolders = (Read-MenuChoice -Title 'Include subfolders?' -Options ([ordered]@{'1' = 'No'; '2' = 'Yes'}) -Default '1') -eq '2'
    $extensions = Read-FileExtension -Source $Source

    $labelOptions = [ordered]@{}
    for ($index = 0; $index -lt $Labels.Count; $index++) {
        $label = $Labels[$index]
        $labelOptions[[string]($index + 1)] = "$($label.Name)  (priority $($label.Priority))"
    }
    $labelChoice = Read-MenuChoice -Title 'Choose the target sensitivity label.' -Options $labelOptions
    $targetLabel = $Labels[[int]$labelChoice - 1]
    Write-RunLog -Severity INFO -Action 'Select sensitivity label' -Result "Selected $($targetLabel.Name); GUID $($targetLabel.Id); priority $($targetLabel.Priority)."
    $dryRun = Read-RunMode -Source $Source -TenantId $tenantId
    if ($script:SetupInterrupted) {
        $script:SetupInterrupted = $false
        return $null
    }
    $logFolder = Read-ValueWithDefault -Prompt 'Log/report folder' -Default $PSScriptRoot

    return [pscustomobject]@{
        Source = $Source
        SiteUrl = $siteUrl
        ClientId = $clientId
        LibraryTitle = $libraryTitle
        LibraryUrl = $libraryUrl
        TargetPath = $targetPath
        IncludeSubfolders = $includeSubfolders
        Extensions = $extensions
        TargetLabel = $targetLabel
        DryRun = $dryRun
        LogFolder = $logFolder
    }
}

function Read-NextSharePointFolder {
    <# .SYNOPSIS Asks for another folder in the library just scanned, keeping every other setting. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Settings)

    $subfolder = (Read-ValueWithDefault -Prompt "Subfolder inside '$($Settings.LibraryTitle)', or . for the whole library" -Default '.').Trim('/', '\')
    if ($subfolder -eq '.') { $subfolder = '' }
    $targetPath = if ($subfolder) { "$($Settings.LibraryUrl)/$($subfolder -replace '\\', '/')" } else { $Settings.LibraryUrl }

    $next = $Settings.PSObject.Copy()
    $next.TargetPath = $targetPath
    Write-RunLog -Severity INFO -Action 'Select folder' -Result "Continuing in the same library with /$targetPath, keeping the label, filters, and mode already chosen."
    return $next
}

function Show-SettingsSummary {
    <# .SYNOPSIS Logs the complete selection summary. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Settings)

    $mode = if ($Settings.DryRun) { 'DRY RUN (no labels changed)' } else { 'APPLY (files will be changed)' }
    Write-RunLog -Severity INFO -Action 'Selection summary' -Result ("Source: {0}; site: {1}; target: {2}; label: {3}; priority: {4}; subfolders: {5}; filters: {6}; mode: {7}; output: {8}" -f `
            $Settings.Source, $Settings.SiteUrl, $Settings.TargetPath, $Settings.TargetLabel.Name, $Settings.TargetLabel.Priority,
            $Settings.IncludeSubfolders, ($Settings.Extensions -join ', '),
            $(if ($Settings.DryRun) { 'DRY RUN' } else { 'APPLY' }), $Settings.LogFolder) -NoConsole
    Write-Host ''
    Write-Host '  Selection summary' -ForegroundColor Cyan
    Write-Host "    Source        : $($Settings.Source)"
    if ($Settings.Source -eq 'SharePoint') {
        Write-Host "    Site          : $($Settings.SiteUrl)"
        Write-Host "    Application   : $($Settings.ClientId)"
        Write-Host "    Library       : $($Settings.LibraryTitle)"
        Write-Host "    Scanned path  : /$($Settings.TargetPath)"
    }
    else {
        Write-Host "    Target folder : $($Settings.TargetPath)"
    }
    Write-Host "    Subfolders    : $($Settings.IncludeSubfolders)"
    Write-Host "    Extensions    : $($Settings.Extensions -join ', ')"
    Write-Host "    Label         : $($Settings.TargetLabel.Name)  (priority $($Settings.TargetLabel.Priority))"
    Write-Host "    Mode          : $mode" -ForegroundColor $(if ($Settings.DryRun) { 'Green' } else { 'Yellow' })
    Write-Host "    Log/report    : $($Settings.LogFolder)"
}

function ConvertTo-ExtensionList {
    <# .SYNOPSIS Normalizes a comma-separated list into lowercase extensions that each start with a dot. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    return @($Text.Split(',', [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object {
            $value = $_.Trim().ToLowerInvariant()
            if (-not $value.StartsWith('.')) { $value = ".$value" }
            $value
        } | Select-Object -Unique)
}

function Read-FileExtension {
    <# .SYNOPSIS Offers the file sets the chosen source can actually label, or a typed list. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][ValidateSet('Local', 'SharePoint')][string]$Source)

    if ($Source -eq 'SharePoint') {
        $supported = @($script:SharePointFileExtensions)
        $wideChoice = 'Every format SharePoint can label: ' + ($script:SharePointFileExtensions -join ' ')
        Write-Host ''
        Write-Host '  SharePoint labels a shorter list of formats than the Purview client does.' -ForegroundColor Gray
        Write-Host '  Legacy (.doc .xls .ppt), template, and Visio files cannot be labeled there,' -ForegroundColor Gray
        Write-Host '  and .pdf works only where the tenant enabled sensitivity labels for PDF.' -ForegroundColor Gray
    }
    else {
        $supported = @($script:ClientFileExtensions)
        $wideChoice = 'All Office documents, Visio drawings, and PDF, including legacy and macro-enabled formats'
    }

    while ($true) {
        $choice = Read-MenuChoice -Title 'Which files should be scanned?' -Options ([ordered]@{
                '1' = $wideChoice
                '2' = 'Current Office formats only: .docx .xlsx .pptx .pdf'
                '3' = 'Type my own list'
            }) -Default '1'
        if ($choice -eq '1') { return @($supported) }
        if ($choice -eq '2') { return @($script:ModernFileExtensions) }

        $extensions = ConvertTo-ExtensionList -Text (Read-ValueWithDefault -Prompt 'File extensions, comma-separated' -Default ($script:ModernFileExtensions -join ','))
        if ($extensions.Count -eq 0) {
            Write-RunLog -Severity WARN -Action 'Validate file filters' -Result 'At least one extension is required.'
            continue
        }
        # Scanning an unlabelable type is allowed, because only the write fails, but on SharePoint that write still costs.
        $unsupported = @($extensions | Where-Object { $supported -notcontains $_ })
        if ($unsupported.Count -gt 0) {
            $consequence = if ($Source -eq 'SharePoint') {
                'SharePoint refuses to label these, and in apply mode each refusal is still a metered API call'
            }
            else {
                'the Purview client cannot label these'
            }
            Write-RunLog -Severity WARN -Action 'Validate file filters' -Result "$($unsupported -join ' '): $consequence. They are still listed by the scan."
        }
        return @($extensions)
    }
}

function Resolve-PlanLabel {
    <# .SYNOPSIS Matches one CSV label cell against the tenant labels by GUID, full name, or leaf name. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Value,
        [Parameter(Mandatory)][object[]]$Labels
    )

    $text = $Value.Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    foreach ($label in $Labels) { if ($label.Id -eq $text) { return $label } }
    foreach ($label in $Labels) { if ($label.Name -eq $text) { return $label } }
    # A parented label reads "Parent \ Child", so accept the child on its own only when
    # that leaf name identifies exactly one tenant label. Guessing here could apply the
    # wrong protection when two parents contain a sublabel with the same display name.
    $leafMatches = @($Labels | Where-Object { ($_.Name -split ' \\ ')[-1] -eq $text })
    if ($leafMatches.Count -eq 1) { return $leafMatches[0] }
    return $null
}

function Test-LabelingPlanFolder {
    <# .SYNOPSIS Rejects a CSV folder that could escape the root selected for the batch run. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Folder)

    $text = $Folder.Trim()
    if ([string]::IsNullOrWhiteSpace($text) -or $text -eq '.') { return $true }
    # Reject UNC/rooted paths, drive-qualified paths (including C:relative), and URI schemes.
    if ($text -match '^(?:[\\/]|[A-Za-z]:|[A-Za-z][A-Za-z0-9+.-]*:)') { return $false }
    # Join-Path and SharePoint both resolve parent segments, so even a middle ".." can leave the root.
    return @($text -split '[\\/]+' | Where-Object { $_ -eq '..' }).Count -eq 0
}

function Import-LabelingPlan {
    <# .SYNOPSIS Reads folder-to-label assignments from CSV and resolves each label against the tenant. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][object[]]$Labels,
        [Parameter(Mandatory)][string[]]$DefaultExtensions
    )

    try {
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf -ErrorAction Stop)) {
            throw "No file was found at '$Path'."
        }
        $rows = @(Import-Csv -LiteralPath $Path -ErrorAction Stop)
        if ($rows.Count -eq 0) { throw 'The CSV file has a header but no data rows.' }

        $columns = @($rows[0].PSObject.Properties.Name)
        foreach ($required in 'Folder', 'Label') {
            if ($columns -notcontains $required) {
                throw "The CSV file needs a '$required' column. Columns found: $($columns -join ', ')."
            }
        }

        $plan = [System.Collections.Generic.List[object]]::new()
        $lineNumber = 1
        foreach ($row in $rows) {
            $lineNumber++
            $labelText = [string](Get-ObjectPropertyValue -InputObject $row -Names 'Label')
            $label = Resolve-PlanLabel -Value $labelText -Labels $Labels
            if ($null -eq $label) {
                Add-RunFailure -FilePath $Path -Action 'Read labeling plan' -Reason "Line ${lineNumber}: label '$labelText' does not uniquely match a file-capable tenant label, so the row was dropped. Use its full parented name or GUID."
                continue
            }

            $folderText = ([string](Get-ObjectPropertyValue -InputObject $row -Names 'Folder')).Trim()
            if (-not (Test-LabelingPlanFolder -Folder $folderText)) {
                Add-RunFailure -FilePath $Path -Action 'Read labeling plan' -Reason "Line ${lineNumber}: folder '$folderText' is rooted, drive-qualified, a URI, or contains a parent '..' segment, so it could escape the selected batch root and was dropped."
                continue
            }

            $recurseText = ([string](Get-ObjectPropertyValue -InputObject $row -Names 'Recurse')).Trim()
            $recurse = $true
            if (-not [string]::IsNullOrWhiteSpace($recurseText)) {
                if ($recurseText -match '^(?i)(true|yes|y|1)$') { $recurse = $true }
                elseif ($recurseText -match '^(?i)(false|no|n|0)$') { $recurse = $false }
                else {
                    Add-RunFailure -FilePath $Path -Action 'Read labeling plan' -Reason "Line ${lineNumber}: Recurse value '$recurseText' is neither true nor false, so the row was dropped."
                    continue
                }
            }

            $extensionText = [string](Get-ObjectPropertyValue -InputObject $row -Names 'Extensions')
            $extensions = @($DefaultExtensions)
            $parsed = @(ConvertTo-ExtensionList -Text $extensionText)
            if ($parsed.Count -gt 0) { $extensions = $parsed }

            $plan.Add([pscustomobject]@{
                    Folder = $folderText.Trim('/', '\')
                    Label = $label
                    Recurse = $recurse
                    Extensions = @($extensions)
                })
        }

        if ($plan.Count -eq 0) { throw 'Every row was dropped, so there is nothing to process.' }
        Write-RunLog -Severity SUCCESS -FilePath $Path -Action 'Read labeling plan' -Result "Loaded $($plan.Count) folder assignments."
        return @($plan)
    }
    catch {
        Add-RunFailure -FilePath $Path -Action 'Read labeling plan' -Reason $_.Exception.Message
        return $null
    }
}

function Show-LabelingPlan {
    <# .SYNOPSIS Prints the folder-to-label assignments a batch run will process. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$Rows,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Root
    )

    Write-Host ''
    Write-Host '  Labeling plan' -ForegroundColor Cyan
    Write-Host "    Root: $Root"
    $index = 0
    foreach ($row in $Rows) {
        $index++
        $folderText = if ([string]::IsNullOrWhiteSpace($row.Folder)) { '(the root itself)' } else { $row.Folder }
        Write-Host ('    {0}. {1}' -f $index, $folderText)
        Write-Host "         label      : $($row.Label.Name)  (priority $($row.Label.Priority))"
        Write-Host "         subfolders : $($row.Recurse)"
        Write-Host "         extensions : $($row.Extensions -join ', ')"
        Write-RunLog -Severity INFO -Action 'Labeling plan' -Result "$folderText -> $($row.Label.Name); subfolders $($row.Recurse); $($row.Extensions -join ' ')" -NoConsole
    }
}

function Join-LabelingPath {
    <# .SYNOPSIS Combines the run root with one plan folder using the separator the source expects. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Folder,
        [Parameter(Mandatory)][ValidateSet('Local', 'SharePoint')][string]$Source
    )

    if ([string]::IsNullOrWhiteSpace($Folder) -or $Folder -eq '.') { return $Root }
    if ($Source -eq 'SharePoint') { return "$($Root.TrimEnd('/'))/$($Folder -replace '\\', '/')" }
    return (Join-Path $Root ($Folder -replace '/', '\'))
}

function Invoke-BatchRun {
    <# .SYNOPSIS Applies a different label to each folder listed in a CSV, in a single pass. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][ValidateSet('Local', 'SharePoint')][string]$Source)

    $source = $Source
    $labels = @(Get-TenantSensitivityLabel | Where-Object { $null -ne $_ })
    if ($labels.Count -eq 0) { return 'Main' }

    $siteUrl = ''
    $tenantId = ''
    $libraryTitle = ''
    $root = ''
    if ($source -eq 'SharePoint') {
        $target = Read-SharePointTarget -SkipSubfolder
        if ($null -eq $target) { return 'Main' }
        $siteUrl = $target.SiteUrl
        $tenantId = $target.TenantId
        $libraryTitle = $target.LibraryTitle
        $root = $target.TargetPath
    }
    else {
        if (-not (Test-PurviewPrerequisite)) { return 'Main' }
        Write-Host ''
        Write-Host '  Every folder in the CSV is read relative to this root.' -ForegroundColor Gray
        $root = Read-ValueWithDefault -Prompt 'Root folder, for example C:\Data' -Default ''
        if (-not (Test-TargetPath -Path $root)) { return 'Main' }
    }

    Write-Host ''
    Write-Host '  Rows that leave Extensions empty fall back to this choice.' -ForegroundColor Gray
    $defaultExtensions = Read-FileExtension -Source $Source

    Write-Host ''
    Write-Host '  The CSV needs a Folder column and a Label column. Recurse and Extensions' -ForegroundColor Gray
    Write-Host '  are optional. Folder is relative to the root shown above, and a label may' -ForegroundColor Gray
    Write-Host '  be its display name or its GUID. For example:' -ForegroundColor Gray
    Write-Host ''
    Write-Host '    Folder,Label,Recurse,Extensions' -ForegroundColor DarkGray
    Write-Host '    Contracts,Confidential,true,".docx,.pdf"' -ForegroundColor DarkGray
    Write-Host '    Public,General \ Anyone (unrestricted),false,' -ForegroundColor DarkGray
    Write-Host '    ,Confidential,false,' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '  The last row leaves Folder empty, which means the root itself.' -ForegroundColor Gray
    $csvPath = Read-ValueWithDefault -Prompt 'Path to the plan CSV' -Default ''

    $plan = Import-LabelingPlan -Path $csvPath -Labels $labels -DefaultExtensions $defaultExtensions
    if ($null -eq $plan) { return 'Main' }
    Show-LabelingPlan -Rows $plan -Root $root

    $action = Read-PhaseAction -Phase 'Plan review'
    if ($action -eq 'Exit') { return 'Exit' }
    if ($action -ne 'Continue') { return 'Main' }

    $dryRun = Read-RunMode -Source $source -TenantId $tenantId
    if ($script:SetupInterrupted) {
        $script:SetupInterrupted = $false
        return 'Main'
    }
    $logFolder = Read-ValueWithDefault -Prompt 'Log/report folder' -Default $PSScriptRoot
    if (-not (Initialize-RunArtifact -Folder $logFolder -Confirm:$false)) { return 'Main' }

    if (-not $dryRun) {
        $confirm = Read-MenuChoice -Title "WRITE OPERATION: Apply this plan to $($plan.Count) folders?" -Options ([ordered]@{
                '1' = 'No, return to the main menu (default)'
                '2' = 'Yes, apply labels'
            }) -Default '1'
        if ($confirm -ne '2') { return 'Main' }
    }

    foreach ($row in $plan) {
        $targetPath = Join-LabelingPath -Root $root -Folder $row.Folder -Source $source
        Write-Host ''
        Write-Host "  Scanning $targetPath" -ForegroundColor Cyan
        Write-Host "    Label      : $($row.Label.Name)"
        Write-Host "    Subfolders : $($row.Recurse)"
        Write-Host "    Extensions : $($row.Extensions -join ', ')"

        $files = if ($source -eq 'SharePoint') {
            Get-SharePointTargetFile -SiteUrl $siteUrl -FolderSiteRelativeUrl $targetPath -LibraryTitle $libraryTitle -Extensions $row.Extensions -Recurse $row.Recurse
        }
        else {
            Get-TargetFile -Path $targetPath -Extensions $row.Extensions -Recurse $row.Recurse
        }
        if ($null -eq $files) { continue }
        $files = @($files)
        if ($files.Count -eq 0) { continue }

        $result = Invoke-FileProcessing -Files $files -TargetLabel $row.Label -ConfiguredLabels $labels -DryRun $dryRun -Source $source
        Write-Progress -Id 1 -Activity 'Processing files' -Completed
        if ($result -eq 'Main') {
            Write-RunLog -Severity WARN -Action 'Batch run' -Result 'Stopped early. The folders after this one were not processed.'
            break
        }
    }

    Export-RunReport -Confirm:$false
    Show-RunSummary
    return 'Main'
}

function Get-ConfidentialClientConfig {
    <# .SYNOPSIS Returns the certificate-based application this utility registered, or null. #>
    [CmdletBinding()]
    param()

    if ($script:IgnoreRemembered) { return $null }
    $found = @{}
    foreach ($name in 'PURVIEW_FILE_LABELING_CC_CLIENT_ID', 'PURVIEW_FILE_LABELING_CC_TENANT_ID', 'PURVIEW_FILE_LABELING_CC_THUMBPRINT') {
        $value = ''
        foreach ($target in [EnvironmentVariableTarget]::Process, [EnvironmentVariableTarget]::User) {
            try {
                $candidate = [Environment]::GetEnvironmentVariable($name, $target)
                if (-not [string]::IsNullOrWhiteSpace($candidate)) { $value = $candidate.Trim(); break }
            }
            catch { Write-Verbose "Could not read ${name}: $($_.Exception.Message)" }
        }
        if ([string]::IsNullOrWhiteSpace($value)) { return $null }
        $found[$name] = $value
    }
    return [pscustomobject]@{
        ClientId = $found['PURVIEW_FILE_LABELING_CC_CLIENT_ID']
        TenantId = $found['PURVIEW_FILE_LABELING_CC_TENANT_ID']
        Thumbprint = $found['PURVIEW_FILE_LABELING_CC_THUMBPRINT']
    }
}

function Save-ConfidentialClientConfig {
    <# .SYNOPSIS Remembers the non-secret identifiers of the certificate-based application. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$Thumbprint
    )

    $values = [ordered]@{
        PURVIEW_FILE_LABELING_CC_CLIENT_ID = $ClientId
        PURVIEW_FILE_LABELING_CC_TENANT_ID = $TenantId
        PURVIEW_FILE_LABELING_CC_THUMBPRINT = $Thumbprint
    }
    foreach ($name in $values.Keys) {
        [Environment]::SetEnvironmentVariable($name, $values[$name], [EnvironmentVariableTarget]::Process)
        try { [Environment]::SetEnvironmentVariable($name, $values[$name], [EnvironmentVariableTarget]::User) }
        catch { Write-RunLog -Severity WARN -Action 'Save confidential client' -Result "Could not persist ${name}: $($_.Exception.Message)" }
    }
    # Only the private key can authenticate, and it never leaves the user certificate store, so none of these values are secret.
    Write-RunLog -Severity SUCCESS -Action 'Save confidential client' -Result "Remembered application $ClientId with certificate $Thumbprint."
}

function Clear-ConfidentialClientConfig {
    <# .SYNOPSIS Forgets the certificate-based application so later runs fall back to survey-only mode. #>
    [CmdletBinding()]
    param()

    # Read before clearing: once the thumbprint is gone the certificate can no longer be identified.
    $config = Get-ConfidentialClientConfig
    $thumbprint = if ($null -ne $config) { [string]$config.Thumbprint } else { '' }
    foreach ($name in 'PURVIEW_FILE_LABELING_CC_CLIENT_ID', 'PURVIEW_FILE_LABELING_CC_TENANT_ID', 'PURVIEW_FILE_LABELING_CC_THUMBPRINT') {
        foreach ($target in [EnvironmentVariableTarget]::Process, [EnvironmentVariableTarget]::User) {
            try { [Environment]::SetEnvironmentVariable($name, $null, $target) }
            catch { Write-Verbose "Could not clear ${name}: $($_.Exception.Message)" }
        }
    }
    $note = if (Remove-LabelingCertificate -Thumbprint $thumbprint -Confirm:$false) {
        'Its certificate was deleted from your personal certificate store as well.'
    }
    else { 'No certificate was left behind.' }
    Write-RunLog -Severity INFO -Action 'Forget confidential client' -Result "The saved confidential client was removed. $note"
}

function Test-SigningCertificateAvailable {
    <# .SYNOPSIS Reports whether the signing certificate is still installed for this user. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Thumbprint)

    try { return (Test-Path -LiteralPath "Cert:\CurrentUser\My\$Thumbprint") }
    catch {
        Write-Verbose "Could not read the certificate store: $($_.Exception.Message)"
        return $false
    }
}

function Test-ConfidentialClientReady {
    <# .SYNOPSIS Returns the saved confidential client only when it belongs to this tenant and its certificate is usable. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$TenantId)

    $config = Get-ConfidentialClientConfig
    if ($null -eq $config) { return $null }
    if (-not [string]::IsNullOrWhiteSpace($TenantId) -and $config.TenantId -ne $TenantId) {
        Write-RunLog -Severity INFO -Action 'Check confidential client' -Result "The saved confidential client belongs to tenant $($config.TenantId), not $TenantId, so it was not used. Run the setup again while connected to this tenant to enable writing here."
        return $null
    }
    if (Test-SigningCertificateAvailable -Thumbprint $config.Thumbprint) { return $config }
    Write-RunLog -Severity WARN -Action 'Check confidential client' -Result "Certificate $($config.Thumbprint) is no longer in this user's certificate store, so application $($config.ClientId) cannot sign in. Run the setup again."
    return $null
}

function Connect-SharePointAppOnly {
    <# .SYNOPSIS Signs in to one site as the application itself, which is what the metered write path requires. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [Parameter(Mandatory)][object]$Config
    )

    # A registration is not usable until Entra replicates it, and app-only sign-in fails outright until then.
    if (Test-ExistingSharePointConnection -SiteUrl $SiteUrl -ClientId $Config.ClientId) {
        $script:SharePointSessionOpened = $true
        $script:ConnectedAppOnly = $true
        Write-RunLog -Severity SUCCESS -FilePath $SiteUrl -Action 'Connect SharePoint site' -Result 'Reusing the app-only sign-in already open for this site.'
        return $true
    }
    for ($attempt = 1; $attempt -le 4; $attempt++) {
        try {
            Disconnect-SharePointSession
            Connect-PnPOnline -Url $SiteUrl -ClientId $Config.ClientId -Tenant $Config.TenantId -Thumbprint $Config.Thumbprint -ErrorAction Stop
            $script:SharePointSessionOpened = $true
            $script:ConnectedAppOnly = $true
            $script:ForceFreshSignIn = $false
            $null = Get-PnPWeb -ErrorAction Stop
            Write-RunLog -Severity SUCCESS -FilePath $SiteUrl -Action 'Connect SharePoint site' -Result "Signed in app-only as confidential client $($Config.ClientId)."
            return $true
        }
        catch {
            # Connect-PnPOnline can succeed before the validation read fails. Do not let that
            # partial connection advertise a usable app-only session to the rest of the run.
            $script:SharePointSessionOpened = $false
            $script:ConnectedAppOnly = $false
            $message = Get-ErrorText -ErrorRecord $_
            $isMissingApplication = $message -match '(?i)AADSTS700016|was not found in the directory|application with identifier'
            if ($attempt -lt 4 -and $isMissingApplication) {
                Write-RunLog -Severity WARN -Action 'Connect SharePoint site' -Result "Application $($Config.ClientId) is not visible in tenant $($Config.TenantId) yet. Waiting for Entra to replicate it (attempt $attempt of 4)."
                Start-Sleep -Seconds 15
                continue
            }
            Add-RunFailure -FilePath $SiteUrl -Action 'Connect SharePoint site' -Reason $message
            if ($isMissingApplication) {
                Write-RunLog -Severity INFO -Action 'App-only sign-in guidance' -Result 'For an app-only sign-in this almost always means administrator consent was never granted. Registering the application creates the application object, but the service principal that app-only authentication needs is only created when an administrator consents to its permissions.'
                Write-RunLog -Severity INFO -Action 'App-only sign-in guidance' -Result 'Choose "Enable SharePoint Online metered label writing" on the main menu, keep the existing application, and accept the offer to grant consent.'
            }
            else {
                Write-RunLog -Severity INFO -Action 'App-only sign-in guidance' -Result 'App-only access also needs the site to be reachable by the application. Re-run the confidential client setup from the main menu if consent was never granted.'
            }
            return $false
        }
    }
    return $false
}

function Get-LabelingCertificate {
    <# .SYNOPSIS Finds certificates generated for one application, which the registration names after it. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ApplicationName)

    try {
        return @(Get-ChildItem Cert:\CurrentUser\My -ErrorAction Stop |
                Where-Object { $_.Subject -like "CN=$ApplicationName*" })
    }
    catch {
        Write-Verbose "Could not read the certificate store: $($_.Exception.Message)"
        return @()
    }
}

function Remove-LabelingCertificate {
    <# .SYNOPSIS Deletes a certificate this utility generated. #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Thumbprint)

    if ([string]::IsNullOrWhiteSpace($Thumbprint)) { return $false }
    $path = "Cert:\CurrentUser\My\$Thumbprint"
    if (-not (Test-Path -LiteralPath $path)) { return $false }
    if (-not $PSCmdlet.ShouldProcess($Thumbprint, 'Remove certificate from the personal store')) { return $false }
    try {
        Remove-Item -LiteralPath $path -Force -ErrorAction Stop
        Write-RunLog -Severity INFO -Action 'Remove certificate' -Result "Removed certificate $Thumbprint from your personal certificate store."
        return $true
    }
    catch {
        Write-RunLog -Severity WARN -Action 'Remove certificate' -Result "Could not remove certificate ${Thumbprint}: $(Get-ErrorText -ErrorRecord $_)"
        return $false
    }
}

function Clear-UnusedLabelingCertificate {
    <# .SYNOPSIS Removes every certificate this run generated except the one the saved client still needs. #>
    [CmdletBinding()]
    param()

    if ($script:CreatedCertificateThumbprints.Count -eq 0) { return }
    $keep = ''
    $config = Get-ConfidentialClientConfig
    if ($null -ne $config) { $keep = [string]$config.Thumbprint }
    foreach ($thumbprint in @($script:CreatedCertificateThumbprints | Select-Object -Unique)) {
        if ($thumbprint -eq $keep) { continue }
        $null = Remove-LabelingCertificate -Thumbprint $thumbprint -Confirm:$false
    }
    $script:CreatedCertificateThumbprints.Clear()
}

function Clear-LabelingTemporaryArtifact {
    <# .SYNOPSIS Removes temporary workers and certificate exports whose owning run has ended. #>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingEmptyCatchBlock', '', Justification = 'A missing owner process is the expected cleanup condition.')]
    param([switch]$IncludeCurrentProcess)

    $temporaryRoot = [System.IO.Path]::GetTempPath()
    $legacyCutoff = (Get-Date).AddHours(-1)
    $removed = 0
    $namePattern = '^(?:PurviewFileLabelingGraph|PurviewFileLabelingAzureCli|PurviewFileLabelingAzure|PurviewFileLabelingCert|graphservices)-(?:(?<pid>\d+)-)?[0-9a-f]{32}(?:\.(?:ps1|json))?$'
    foreach ($candidate in @(Get-ChildItem -LiteralPath $temporaryRoot -Force -ErrorAction SilentlyContinue)) {
        $nameMatch = [regex]::Match($candidate.Name, $namePattern)
        if (-not $nameMatch.Success) { continue }
        $ownerText = [string]$nameMatch.Groups['pid'].Value
        if ([string]::IsNullOrWhiteSpace($ownerText)) {
            # Legacy names carry no owner PID, so age is the only safe way to avoid another live run.
            if ($candidate.LastWriteTime -gt $legacyCutoff) { continue }
        }
        else {
            $ownerPid = 0
            if (-not [int]::TryParse($ownerText, [ref]$ownerPid)) { continue }
            if ($ownerPid -eq $PID -and $IncludeCurrentProcess) {
                # The caller is the final cleanup for this run.
            }
            else {
                try {
                    $null = Get-Process -Id $ownerPid -ErrorAction Stop
                    continue
                }
                catch { Write-Verbose "Temporary artifact owner process $ownerPid is no longer running." }
            }
        }
        try {
            Remove-Item -LiteralPath $candidate.FullName -Recurse -Force -ErrorAction Stop
            $removed++
        }
        catch {
            Write-RunLog -Severity WARN -Action 'Remove temporary artifact' -Result "Could not remove $($candidate.FullName): $(Get-ErrorText -ErrorRecord $_)"
        }
    }
    if ($removed -gt 0) {
        Write-RunLog -Severity INFO -Action 'Remove temporary artifact' -Result "Removed $removed temporary worker, Azure CLI profile, or certificate-export artifact(s) left by completed runs."
    }
}

function Clear-OrphanedLabelingCertificate {
    <# .SYNOPSIS Removes certificates from earlier runs that ended before they could clean up after themselves. #>
    [CmdletBinding()]
    param()

    # In-memory tracking dies with the process, so a run that crashes or is closed strands its certificate.
    # Only the store itself can reveal those, and only the saved client's certificate is still needed.
    $keep = ''
    $config = Get-ConfidentialClientConfig
    if ($null -ne $config) { $keep = [string]$config.Thumbprint }

    $orphans = @()
    try {
        $orphans = @(Get-ChildItem Cert:\CurrentUser\My -ErrorAction Stop |
                Where-Object { $_.Subject -like '*Purview File Labeling*' -and $_.Thumbprint -ne $keep })
    }
    catch {
        Write-Verbose "Could not read the certificate store: $($_.Exception.Message)"
        return
    }
    if ($orphans.Count -eq 0) { return }

    Write-RunLog -Severity INFO -Action 'Remove certificate' -Result "Found $($orphans.Count) certificate(s) in your personal store from earlier runs that did not finish. They cannot authenticate anything, because the applications they belonged to were never saved, so they are being removed."
    foreach ($certificate in $orphans) {
        $null = Remove-LabelingCertificate -Thumbprint $certificate.Thumbprint -Confirm:$false
    }
}

function New-ConfidentialLabelingApplication {
    <# .SYNOPSIS Registers a certificate-based application, the only client shape the metered API accepts. #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ApplicationName
    )

    $command = Get-Command Register-PnPEntraIDApp -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $command) {
        throw 'This PnP.PowerShell version cannot register a certificate-based application. Update it with: Update-Module PnP.PowerShell'
    }
    # An auto-discovered command can report an incomplete parameter list until its module is actually loaded.
    if (-not $command.Parameters.ContainsKey('GraphApplicationPermissions')) {
        try { Import-Module PnP.PowerShell -ErrorAction Stop }
        catch { Write-Verbose "Could not import PnP.PowerShell: $($_.Exception.Message)" }
        $command = Get-Command Register-PnPEntraIDApp -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $command) { throw 'Register-PnPEntraIDApp disappeared after importing PnP.PowerShell.' }
    }
    $missing = @(@('GraphApplicationPermissions', 'SharePointApplicationPermissions', 'Store', 'Tenant') |
            Where-Object { -not $command.Parameters.ContainsKey($_) })
    if ($missing.Count -gt 0) {
        $loadedVersion = [string]$command.Version
        $moduleBase = if ($command.Module) { [string]$command.Module.ModuleBase } else { 'unknown location' }
        $newest = Get-Module -ListAvailable -Name PnP.PowerShell -ErrorAction SilentlyContinue |
            Sort-Object Version -Descending | Select-Object -First 1
        # PnP loads .NET assemblies that cannot be unloaded, so a newer version cannot replace an older one in place.
        if ($newest -and [version]$newest.Version -gt [version]$loadedVersion) {
            throw "PnP.PowerShell $loadedVersion is loaded in this session, but $($newest.Version) is installed and does support the application permission parameters. PnP loads .NET assemblies that cannot be swapped once loaded, so close this PowerShell window, open a new one, and run this utility again."
        }

        Write-RunLog -Severity WARN -Action 'Register confidential client' -Result "PnP.PowerShell $loadedVersion, loaded from $moduleBase, cannot set application permissions: -$($missing -join ', -') are unavailable. A newer PnP.PowerShell is needed."
        if (Request-ModuleInstall -Name 'PnP.PowerShell' -Purpose 'to register a confidential client') {
            if (Request-SelfRestart -Reason 'PnP.PowerShell was updated, and a module already loaded cannot be replaced in a running session.') {
                throw 'PnP.PowerShell has been updated. This utility restarts itself now so the new version is loaded.'
            }
            throw 'PnP.PowerShell has been updated. Close this PowerShell window, open a new one, and run this utility again.'
        }
        $available = (@($command.Parameters.Keys) | Sort-Object) -join ', '
        throw "Register-PnPEntraIDApp in PnP.PowerShell $loadedVersion cannot set application permissions. Update it with 'Update-Module PnP.PowerShell -Scope CurrentUser -Force', then start a new PowerShell session. Parameters it offers: $available"
    }
    if (-not $PSCmdlet.ShouldProcess("tenant $TenantId", "Register confidential client '$ApplicationName'")) { return $null }

    Write-RunLog -Severity INFO -Action 'Register confidential client' -Result "Registering '$ApplicationName' with Graph Files.ReadWrite.All and SharePoint Sites.Read.All, both application permissions."
    Write-RunLog -Severity INFO -Action 'Register confidential client' -Result 'Application permissions always need administrator consent, so sign in as a Global Administrator or Privileged Role Administrator and accept the consent prompt.'
    # The generated key is installed into the user certificate store, so the exported copies PnP also writes are
    # redundant. They are directed to a temporary folder and deleted, rather than left in Downloads.
    # Registering signs in again, which replaces any SharePoint connection the caller was using.
    $script:SetupInterrupted = $true
    $certificateOutPath = Join-Path ([System.IO.Path]::GetTempPath()) ('PurviewFileLabelingCert-' + $PID + '-' + [guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $certificateOutPath -Force -ErrorAction Stop
    $registrationName = $ApplicationName
    $registrationStarted = Get-Date
    $output = @()
    $useDeviceLogin = $script:UseDeviceCode -and $command.Parameters.ContainsKey('DeviceLogin')
    Show-RegistrationSignInNotice -DeviceCode:$useDeviceLogin
    try {
        while ($true) {
            try {
                $parameters = @{
                    ApplicationName = $registrationName
                    Tenant = $TenantId
                    GraphApplicationPermissions = 'Files.ReadWrite.All'
                    SharePointApplicationPermissions = 'Sites.Read.All'
                    Store = 'CurrentUser'
                    OutPath = $certificateOutPath
                    ErrorAction = 'Stop'
                }
                if ($useDeviceLogin) { $parameters.DeviceLogin = $true }
                $output = @(Register-PnPEntraIDApp @parameters)
                break
            }
            catch {
                if ((Get-ErrorText -ErrorRecord $_) -notmatch '(?i)already exists') { throw }
                $decision = Read-DuplicateApplicationChoice -DisplayName $registrationName
                if ($null -eq $decision) { throw "Registration stopped because '$registrationName' already exists in tenant $TenantId." }
                if ($decision.Action -eq 'Existing') {
                    throw "Application $($decision.Value) already exists. This utility cannot reuse it, because it cannot confirm the certificate it was created with. Delete it in the Entra admin center, or register under a different name."
                }
                $registrationName = $decision.Value
                Write-RunLog -Severity INFO -Action 'Register confidential client' -Result "Retrying registration as '$registrationName'."
            }
        }
    }
    finally {
        $exported = @(Get-ChildItem -LiteralPath $certificateOutPath -File -ErrorAction SilentlyContinue)
        Remove-Item -LiteralPath $certificateOutPath -Recurse -Force -ErrorAction SilentlyContinue
        if ($exported.Count -gt 0) {
            Write-RunLog -Severity INFO -Action 'Register confidential client' -Result "Deleted $($exported.Count) exported certificate file(s). The private key stays in Cert:\CurrentUser\My, protected by Windows, and can be exported from there if another machine ever needs it."
        }
        # Recorded so anything generated here is removed later unless it becomes the saved client's credential.
        # The time filter matters: a name can prefix-match a certificate this utility did not create.
        foreach ($certificate in @(Get-LabelingCertificate -ApplicationName $registrationName)) {
            if ($certificate.NotBefore -lt $registrationStarted.AddMinutes(-5)) { continue }
            if (-not $script:CreatedCertificateThumbprints.Contains([string]$certificate.Thumbprint)) {
                $script:CreatedCertificateThumbprints.Add([string]$certificate.Thumbprint)
            }
        }
    }

    $clientId = ''
    $thumbprint = ''
    foreach ($item in $output) {
        if ($null -eq $item) { continue }
        if ([string]::IsNullOrWhiteSpace($clientId)) {
            foreach ($propertyName in 'AzureAppId/ClientId', 'ClientId', 'AppId', 'ApplicationId') {
                $value = Get-ObjectPropertyValue -InputObject $item -Names @($propertyName)
                if (-not [string]::IsNullOrWhiteSpace($value)) { $clientId = $value; break }
            }
        }
        if ([string]::IsNullOrWhiteSpace($thumbprint)) {
            $thumbprint = Get-ObjectPropertyValue -InputObject $item -Names 'Certificate Thumbprint', 'CertificateThumbprint', 'Thumbprint'
        }
    }
    $text = [string]$output
    if ([string]::IsNullOrWhiteSpace($clientId)) {
        $match = [regex]::Match($text, '(?i)(?<![0-9a-f])[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}(?![0-9a-f])')
        if ($match.Success) { $clientId = $match.Value }
    }
    if ([string]::IsNullOrWhiteSpace($thumbprint)) {
        $match = [regex]::Match($text, '(?i)(?<![0-9a-f])[0-9a-f]{40}(?![0-9a-f])')
        if ($match.Success) { $thumbprint = $match.Value }
    }
    if ([string]::IsNullOrWhiteSpace($clientId) -or [string]::IsNullOrWhiteSpace($thumbprint)) {
        throw 'Registration completed but did not report both an application ID and a certificate thumbprint.'
    }

    return [pscustomobject]@{
        ClientId = $clientId.Trim()
        TenantId = $TenantId
        Thumbprint = $thumbprint.Trim().Replace(' ', '').ToUpperInvariant()
    }
}

function Disable-LabelingApplication {
    <# .SYNOPSIS Strips an application's permissions and disables its sign-in, then deletes it when that is allowed. #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][AllowEmptyString()][string]$TenantId
    )

    if (-not $PSCmdlet.ShouldProcess($ClientId, 'Disable and remove the Entra application')) { return $false }

    $response = Invoke-GraphAction -Action 'RemoveApp' -TenantId $TenantId -Arguments @{ClientId = $ClientId} -Scopes @('Application.ReadWrite.All')
    switch ([string]$response.Status) {
        'Deleted' {
            Write-RunLog -Severity SUCCESS -Action 'Disable application' -Result "Deleted application $ClientId from tenant $TenantId."
            return $true
        }
        'Disabled' {
            Write-ManualRemovalNote -ClientId $ClientId -TenantId $TenantId -Stripped $true
            return $true
        }
        'Missing' {
            Write-RunLog -Severity WARN -Action 'Disable application' -Result "Application $ClientId was not found in tenant $TenantId, so nothing was changed."
            return $false
        }
        default {
            Add-RunFailure -FilePath '' -Action 'Disable application' -Reason ([string]$response.Message)
            Write-ManualRemovalNote -ClientId $ClientId -TenantId $TenantId -Stripped $false
            return $false
        }
    }
}

function Write-ManualRemovalNote {
    <# .SYNOPSIS Says exactly what is left behind and how to remove it by hand. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][AllowEmptyString()][string]$TenantId,
        [Parameter(Mandatory)][bool]$Stripped
    )

    if ($Stripped) {
        Write-RunLog -Severity WARN -Action 'Disable application' -Result "Application $ClientId could not be deleted, but its permissions and sign-in were removed, so it can no longer reach any data."
    }
    else {
        Write-RunLog -Severity WARN -Action 'Disable application' -Result "Application $ClientId still holds its permissions. Removing it needs Microsoft Graph Application.ReadWrite.All, which this sign-in does not have."
    }
    Write-Host ''
    Write-Host '  Remove it in the Entra admin center under App registrations, or run:' -ForegroundColor Gray
    Write-Host "    Connect-MgGraph -TenantId $TenantId -Scopes Application.ReadWrite.All" -ForegroundColor DarkGray
    Write-Host "    Get-MgApplication -Filter `"appId eq '$ClientId'`" | ForEach-Object { Remove-MgApplication -ApplicationId `$_.Id }" -ForegroundColor DarkGray
    Write-Host ''
}

function Show-AdminConsentInstruction {
    <# .SYNOPSIS Describes the portal path to consent, used only when granting it from here was declined or failed. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ClientId)

    # The /adminconsent endpoint needs a redirect_uri that is registered on the application, and a
    # certificate-only registration has none, so linking to it would land on an error page.
    Write-Host ''
    Write-Host '  Grant it in the Entra admin center (https://entra.microsoft.com):' -ForegroundColor Gray
    Write-Host '    Identity > Applications > App registrations > All applications' -ForegroundColor DarkGray
    Write-Host "    open the application with ID $ClientId" -ForegroundColor DarkGray
    Write-Host '    Owners > Add owners > select the account that will run Azure billing setup' -ForegroundColor DarkGray
    Write-Host '    API permissions > Grant admin consent for <your tenant>' -ForegroundColor DarkGray
    Write-Host ''
}

function Get-GraphErrorText {
    <# .SYNOPSIS Reduces a Graph failure to its code and message, because the raw text carries the whole HTTP response. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord)

    $full = Get-ErrorText -ErrorRecord $ErrorRecord
    $code = [regex]::Match($full, '"code"\s*:\s*"(?<value>[^"]+)"')
    $message = [regex]::Match($full, '"message"\s*:\s*"(?<value>[^"]+)"')
    if ($code.Success -and $message.Success) { return '{0}: {1}' -f $code.Groups['value'].Value, $message.Groups['value'].Value }
    if ($full.Length -gt 400) { return $full.Substring(0, 400) + '...' }
    return $full
}

function Get-AzureCliAccountContext {
    <# .SYNOPSIS Returns the active Azure CLI user without exposing its cached access tokens. #>
    [CmdletBinding()]
    param()

    try {
        $output = & az account show --output json --only-show-errors 2>$null
        if ($LASTEXITCODE -ne 0) { return $null }
        $accounts = @(ConvertFrom-AzureCliJson -Output $output)
        if ($accounts.Count -eq 0) { return $null }
        $account = $accounts[0]
        $user = Get-RawObjectPropertyValue -InputObject $account -Names 'user'
        return [pscustomobject]@{
            TenantId = [string](Get-ObjectPropertyValue -InputObject $account -Names 'tenantId', 'homeTenantId')
            SubscriptionId = [string](Get-ObjectPropertyValue -InputObject $account -Names 'id')
            UserName = [string](Get-ObjectPropertyValue -InputObject $user -Names 'name')
            UserType = [string](Get-ObjectPropertyValue -InputObject $user -Names 'type')
        }
    }
    catch {
        Write-Verbose "Could not read the active Azure CLI account: $($_.Exception.Message)"
        return $null
    }
}

function Get-AzureCliApplicationConsentState {
    <# .SYNOPSIS Verifies the application, service principal, and every requested application-role assignment. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ClientId,
        [AllowNull()][object]$Application = $null
    )

    try {
        if ($null -eq $Application) {
            $applicationOutput = & az ad app show --id $ClientId --output json --only-show-errors 2>&1
            if ($LASTEXITCODE -ne 0) {
                $message = (($applicationOutput -join ' ') -replace '\s+', ' ').Trim()
                $status = if ($message -match '(?i)does not exist|not found|ResourceNotFound|Request_ResourceNotFound') { 'ApplicationMissing' } else { 'Unknown' }
                return [pscustomobject]@{Status = $status; Count = 0; Message = "Azure CLI could not read application ${ClientId}: $message"}
            }
            $applications = @(ConvertFrom-AzureCliJson -Output $applicationOutput)
            if ($applications.Count -eq 0) {
                return [pscustomobject]@{Status = 'Unknown'; Count = 0; Message = "Azure CLI returned no application for $ClientId."}
            }
            $Application = $applications[0]
        }

        $requestedRoles = [System.Collections.Generic.List[object]]::new()
        foreach ($resource in @(Get-RawObjectPropertyValue -InputObject $Application -Names 'requiredResourceAccess')) {
            $resourceAppId = [string](Get-ObjectPropertyValue -InputObject $resource -Names 'resourceAppId')
            foreach ($access in @(Get-RawObjectPropertyValue -InputObject $resource -Names 'resourceAccess')) {
                if ([string](Get-ObjectPropertyValue -InputObject $access -Names 'type') -ne 'Role') { continue }
                $requestedRoles.Add([pscustomobject]@{
                        ResourceAppId = $resourceAppId
                        AppRoleId = [string](Get-ObjectPropertyValue -InputObject $access -Names 'id')
                    })
            }
        }
        if ($requestedRoles.Count -eq 0) {
            return [pscustomobject]@{Status = 'ApplicationInvalid'; Count = 0; Message = 'The registration requests no application permissions.'}
        }

        $principalOutput = & az ad sp show --id $ClientId --output json --only-show-errors 2>&1
        if ($LASTEXITCODE -ne 0) {
            $message = (($principalOutput -join ' ') -replace '\s+', ' ').Trim()
            $status = if ($message -match '(?i)does not exist|not found|ResourceNotFound|Request_ResourceNotFound') { 'ConsentMissing' } else { 'Unknown' }
            return [pscustomobject]@{Status = $status; Count = 0; Message = "Azure CLI could not read the application service principal: $message"}
        }
        $principals = @(ConvertFrom-AzureCliJson -Output $principalOutput)
        $principalId = if ($principals.Count -eq 0) { '' } else { [string](Get-ObjectPropertyValue -InputObject $principals[0] -Names 'id') }
        if ([string]::IsNullOrWhiteSpace($principalId)) {
            return [pscustomobject]@{Status = 'Unknown'; Count = 0; Message = 'The application service principal has no directory object ID.'}
        }

        $assignmentUrl = "https://graph.microsoft.com/v1.0/servicePrincipals/$principalId/appRoleAssignments?`$select=appRoleId,resourceId"
        $assignmentOutput = & az rest --method get --url $assignmentUrl --output json --only-show-errors 2>&1
        if ($LASTEXITCODE -ne 0) {
            return [pscustomobject]@{Status = 'Unknown'; Count = 0; Message = "Azure CLI could not verify application-role assignments: $(($assignmentOutput -join ' ').Trim())"}
        }
        $assignmentResponses = @(ConvertFrom-AzureCliJson -Output $assignmentOutput)
        $assignments = if ($assignmentResponses.Count -eq 0) { @() } else {
            @(Get-RawObjectPropertyValue -InputObject $assignmentResponses[0] -Names 'value')
        }
        $resourcePrincipalIds = @{}
        $missingRoles = [System.Collections.Generic.List[string]]::new()
        foreach ($role in $requestedRoles) {
            if (-not $resourcePrincipalIds.ContainsKey($role.ResourceAppId)) {
                $resourceOutput = & az ad sp show --id $role.ResourceAppId --query id --output tsv --only-show-errors 2>&1
                if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace("$resourceOutput")) {
                    return [pscustomobject]@{Status = 'Unknown'; Count = 0; Message = "Azure CLI could not read resource service principal $($role.ResourceAppId): $(($resourceOutput -join ' ').Trim())"}
                }
                $resourcePrincipalIds[$role.ResourceAppId] = "$resourceOutput".Trim()
            }
            $resourcePrincipalId = [string]$resourcePrincipalIds[$role.ResourceAppId]
            $matched = @($assignments | Where-Object {
                    [string]::Equals([string](Get-ObjectPropertyValue -InputObject $_ -Names 'appRoleId'), $role.AppRoleId, [System.StringComparison]::OrdinalIgnoreCase) -and
                    [string]::Equals([string](Get-ObjectPropertyValue -InputObject $_ -Names 'resourceId'), $resourcePrincipalId, [System.StringComparison]::OrdinalIgnoreCase)
                }).Count -gt 0
            if (-not $matched) { $missingRoles.Add("$($role.ResourceAppId)/$($role.AppRoleId)") }
        }
        if ($missingRoles.Count -gt 0) {
            return [pscustomobject]@{Status = 'ConsentMissing'; Count = $requestedRoles.Count - $missingRoles.Count; Message = "These requested application permissions are not assigned: $($missingRoles -join ', ')."}
        }
        return [pscustomobject]@{Status = 'Ready'; Count = $requestedRoles.Count; Message = "$($requestedRoles.Count) requested application permission(s) are assigned."}
    }
    catch {
        return [pscustomobject]@{Status = 'Unknown'; Count = 0; Message = "Azure CLI could not verify application consent for ${ClientId}: $(Get-ErrorText -ErrorRecord $_)"}
    }
}

function Start-IsolatedAzureCliProfile {
    <# .SYNOPSIS Gives utility-owned Azure authentication a temporary profile, preserving the user's normal CLI state. #>
    [CmdletBinding()]
    param()

    if (-not [string]::IsNullOrWhiteSpace($script:AzureCliIsolatedConfigDirectory)) { return $true }
    $directory = Join-Path ([System.IO.Path]::GetTempPath()) ('PurviewFileLabelingAzureCli-' + $PID + '-' + [guid]::NewGuid().ToString('N'))
    try {
        $null = New-Item -ItemType Directory -Path $directory -Force -ErrorAction Stop
        $script:AzureCliPreviousConfigDirectory = [Environment]::GetEnvironmentVariable('AZURE_CONFIG_DIR', [EnvironmentVariableTarget]::Process)
        [Environment]::SetEnvironmentVariable('AZURE_CONFIG_DIR', $directory, [EnvironmentVariableTarget]::Process)
        $script:AzureCliIsolatedConfigDirectory = $directory
        Write-RunLog -Severity INFO -Action 'Prepare Azure CLI' -NoConsole -Result "Using temporary Azure CLI profile $directory so this run cannot change or sign out a pre-existing CLI account."
        return $true
    }
    catch {
        $reason = Get-ErrorText -ErrorRecord $_
        try {
            [Environment]::SetEnvironmentVariable('AZURE_CONFIG_DIR', $script:AzureCliPreviousConfigDirectory, [EnvironmentVariableTarget]::Process)
            if (Test-Path -LiteralPath $directory) { Remove-Item -LiteralPath $directory -Recurse -Force -ErrorAction Stop }
        }
        catch { Write-Verbose "Could not roll back temporary Azure CLI profile ${directory}: $($_.Exception.Message)" }
        $script:AzureCliIsolatedConfigDirectory = ''
        $script:AzureCliPreviousConfigDirectory = $null
        Write-RunLog -Severity WARN -Action 'Prepare Azure CLI' -Result "Could not create an isolated Azure CLI profile: $reason"
        return $false
    }
}

function Clear-IsolatedAzureCliProfile {
    <# .SYNOPSIS Restores the previous Azure CLI profile and removes only this run's temporary credentials. #>
    [CmdletBinding()]
    param()

    $directory = $script:AzureCliIsolatedConfigDirectory
    if ([string]::IsNullOrWhiteSpace($directory)) { return }
    try {
        [Environment]::SetEnvironmentVariable('AZURE_CONFIG_DIR', $script:AzureCliPreviousConfigDirectory, [EnvironmentVariableTarget]::Process)
        if (Test-Path -LiteralPath $directory) {
            Remove-Item -LiteralPath $directory -Recurse -Force -ErrorAction Stop
        }
        Write-RunLog -Severity INFO -Action 'Disconnect Azure CLI' -Result 'Removed this run''s temporary Azure CLI profile and restored the pre-existing CLI profile unchanged.'
    }
    catch {
        Write-RunLog -Severity WARN -Action 'Disconnect Azure CLI' -Result "Could not remove temporary Azure CLI profile ${directory}: $(Get-ErrorText -ErrorRecord $_)"
    }
    finally {
        $script:AzureCliIsolatedConfigDirectory = ''
        $script:AzureCliPreviousConfigDirectory = $null
        $script:AzureCliSessionOpened = $false
        $script:AzureCliAccount = ''
    }
}

function Grant-LabelingAdminConsentViaAzureCli {
    <# .SYNOPSIS Uses the active Azure CLI identity for application ownership and administrator consent. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$TenantId
    )

    $result = [ordered]@{Status = 'Error'; Message = ''; Value = ''; Count = 0}
    try {
        $context = Get-AzureCliAccountContext
        if ($null -eq $context -or
            -not [string]::Equals($context.TenantId, $TenantId, [System.StringComparison]::OrdinalIgnoreCase)) {
            $result.Status = 'NoSession'
            $result.Message = "Azure CLI is not signed in to application tenant $TenantId."
            return [pscustomobject]$result
        }
        if (-not [string]::Equals($context.UserType, 'user', [System.StringComparison]::OrdinalIgnoreCase)) {
            $result.Message = 'Azure CLI is signed in as a service principal or managed identity. Administrator consent and direct application ownership require a user account.'
            return [pscustomobject]$result
        }

        $applicationOutput = & az ad app show --id $ClientId --output json --only-show-errors 2>&1
        if ($LASTEXITCODE -ne 0) { throw "Azure CLI could not read application ${ClientId}: $(($applicationOutput -join ' ').Trim())" }
        $applications = @(ConvertFrom-AzureCliJson -Output $applicationOutput)
        if ($applications.Count -eq 0) { throw "Azure CLI returned no application for $ClientId." }
        $application = $applications[0]

        $requestedRoles = [System.Collections.Generic.List[object]]::new()
        foreach ($resource in @(Get-RawObjectPropertyValue -InputObject $application -Names 'requiredResourceAccess')) {
            $resourceAppId = [string](Get-ObjectPropertyValue -InputObject $resource -Names 'resourceAppId')
            foreach ($access in @(Get-RawObjectPropertyValue -InputObject $resource -Names 'resourceAccess')) {
                if ([string](Get-ObjectPropertyValue -InputObject $access -Names 'type') -ne 'Role') { continue }
                $requestedRoles.Add([pscustomobject]@{
                        ResourceAppId = $resourceAppId
                        AppRoleId = [string](Get-ObjectPropertyValue -InputObject $access -Names 'id')
                    })
            }
        }
        if ($requestedRoles.Count -eq 0) { throw 'The registration requests no application permissions, so there is nothing to consent to.' }

        $operatorOutput = & az ad signed-in-user show --output json --only-show-errors 2>&1
        if ($LASTEXITCODE -ne 0) { throw "Azure CLI could not identify its signed-in user: $(($operatorOutput -join ' ').Trim())" }
        $operators = @(ConvertFrom-AzureCliJson -Output $operatorOutput)
        if ($operators.Count -eq 0) { throw 'Azure CLI returned no signed-in user.' }
        $operator = $operators[0]
        $operatorId = [string](Get-ObjectPropertyValue -InputObject $operator -Names 'id')
        if ([string]::IsNullOrWhiteSpace($operatorId)) { throw 'Azure CLI returned no object ID for its signed-in user.' }

        $ownerOutput = & az ad app owner list --id $ClientId --output json --only-show-errors 2>&1
        if ($LASTEXITCODE -ne 0) { throw "Azure CLI could not list application owners: $(($ownerOutput -join ' ').Trim())" }
        $ownerIds = @((ConvertFrom-AzureCliJson -Output $ownerOutput) | ForEach-Object {
                [string](Get-ObjectPropertyValue -InputObject $_ -Names 'id')
            })
        if ($ownerIds -notcontains $operatorId) {
            $ownerAddOutput = & az ad app owner add --id $ClientId --owner-object-id $operatorId --only-show-errors 2>&1
            if ($LASTEXITCODE -ne 0) { throw "Azure CLI could not assign the signed-in user as application owner: $(($ownerAddOutput -join ' ').Trim())" }
            $ownerOutput = & az ad app owner list --id $ClientId --output json --only-show-errors 2>&1
            if ($LASTEXITCODE -ne 0) { throw "Azure CLI could not verify application ownership: $(($ownerOutput -join ' ').Trim())" }
            $ownerIds = @((ConvertFrom-AzureCliJson -Output $ownerOutput) | ForEach-Object {
                    [string](Get-ObjectPropertyValue -InputObject $_ -Names 'id')
                })
            if ($ownerIds -notcontains $operatorId) { throw "Application owner assignment for object $operatorId could not be verified." }
        }

        $principalOutput = & az ad sp show --id $ClientId --output json --only-show-errors 2>&1
        if ($LASTEXITCODE -ne 0) {
            $principalOutput = & az ad sp create --id $ClientId --output json --only-show-errors 2>&1
            if ($LASTEXITCODE -ne 0) { throw "Azure CLI could not create the service principal: $(($principalOutput -join ' ').Trim())" }
        }
        $principals = @(ConvertFrom-AzureCliJson -Output $principalOutput)
        if ($principals.Count -eq 0) { throw 'Azure CLI returned no service principal.' }
        $principalId = [string](Get-ObjectPropertyValue -InputObject $principals[0] -Names 'id')
        if ([string]::IsNullOrWhiteSpace($principalId)) { throw 'The service principal has no directory object ID.' }

        $consentOutput = & az ad app permission admin-consent --id $ClientId --only-show-errors 2>&1
        if ($LASTEXITCODE -ne 0) { throw "Azure CLI could not grant administrator consent: $(($consentOutput -join ' ').Trim())" }

        $consentState = $null
        for ($attempt = 1; $attempt -le 3; $attempt++) {
            $consentState = Get-AzureCliApplicationConsentState -ClientId $ClientId -Application $application
            if ($consentState.Status -eq 'Ready' -or $consentState.Status -ne 'ConsentMissing' -or $attempt -eq 3) { break }
            Write-RunLog -Severity INFO -Action 'Grant admin consent' -NoConsole -Result "Consent succeeded but its role assignments are not visible yet; read-only verification attempt $attempt of 3 will be repeated."
            Start-Sleep -Seconds 3
        }
        if ($consentState.Status -ne 'Ready') { throw "Administrator consent could not be verified: $($consentState.Message)" }

        $operatorName = [string](Get-ObjectPropertyValue -InputObject $operator -Names 'userPrincipalName', 'displayName')
        if ([string]::IsNullOrWhiteSpace($operatorName)) { $operatorName = $context.UserName }
        if ([string]::IsNullOrWhiteSpace($operatorName)) { $operatorName = $operatorId }
        $result.Status = 'Granted'
        $result.Value = $operatorName
        $result.Count = $consentState.Count
    }
    catch { $result.Message = Get-ErrorText -ErrorRecord $_ }
    return [pscustomobject]$result
}

function Disconnect-LabelingGraph {
    <# .SYNOPSIS Signs the Microsoft Graph session out, in the separate process that owns it. #>
    [CmdletBinding()]
    param()

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication -ErrorAction SilentlyContinue)) { return }
    $response = Invoke-GraphAction -Action 'SignOut' -TenantId '' -NoPrompt
    if ($response.Status -eq 'SignedOut') {
        $script:GraphSessionOpened = $false
        Write-RunLog -Severity INFO -Action 'Disconnect Microsoft Graph' -Result 'Microsoft Graph session signed out.'
    }
}

function Grant-LabelingAdminConsent {
    <# .SYNOPSIS Creates the service principal and grants the application permissions the registration asks for. #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$TenantId,
        [switch]$PreferAzureCli
    )

    if (-not $PSCmdlet.ShouldProcess("application $ClientId", 'Create the service principal and grant its application permissions')) { return $false }

    if ($PreferAzureCli -and (Get-Command az -CommandType Application -ErrorAction SilentlyContinue)) {
        $tool = [pscustomobject]@{Kind = 'AzureCli'; Detail = [string](Get-Command az -CommandType Application | Select-Object -First 1).Source}
        $context = Get-AzureCliAccountContext
        if ($null -eq $context -or
            -not [string]::Equals($context.TenantId, $TenantId, [System.StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals($context.UserType, 'user', [System.StringComparison]::OrdinalIgnoreCase)) {
            Write-RunLog -Severity INFO -Action 'Grant admin consent' -Result 'One Azure CLI sign-in will be reused for administrator consent, application ownership, prerequisite checks, and billing.'
            if (-not (Connect-AzureCommandLine -Tool $tool -TenantId $TenantId)) {
                Write-RunLog -Severity WARN -Action 'Grant admin consent' -Result 'Azure CLI sign-in did not complete, so consent and billing were not attempted. No second sign-in was opened.'
                return $false
            }
        }
        $context = Get-AzureCliAccountContext
        if ($null -eq $context -or
            -not [string]::Equals($context.TenantId, $TenantId, [System.StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals($context.UserType, 'user', [System.StringComparison]::OrdinalIgnoreCase)) {
            Write-RunLog -Severity WARN -Action 'Grant admin consent' -Result "Azure CLI has no user session in application tenant $TenantId, so consent and billing were not attempted."
            return $false
        }
        $cliResponse = Grant-LabelingAdminConsentViaAzureCli -ClientId $ClientId -TenantId $TenantId
        if ($cliResponse.Status -eq 'Granted') {
            Write-RunLog -Severity SUCCESS -Action 'Grant admin consent' -Result "Consent is granted through the reused Azure CLI account: $($cliResponse.Count) application permission(s) are assigned to $ClientId in tenant $TenantId, and $($cliResponse.Value) is verified as an application owner."
            return $true
        }
        Write-RunLog -Severity WARN -Action 'Grant admin consent' -Result "The reused Azure CLI account could not complete consent: $($cliResponse.Message)"
        $fallback = Read-MenuChoice -Title 'Open a separate Microsoft Graph administrator sign-in as a fallback?' -Options ([ordered]@{
                '1' = 'No, stop here and keep the current Azure CLI session (default)'
                '2' = 'Yes, open one separate Microsoft Graph sign-in'
            }) -Default '1'
        if ($fallback -ne '2') {
            Show-AdminConsentInstruction -ClientId $ClientId
            return $false
        }
        Write-RunLog -Severity INFO -Action 'Grant admin consent' -Result 'Opening the Microsoft Graph fallback because it was explicitly selected. Use the same account as Azure CLI so application ownership and billing stay aligned.'
    }

    Write-RunLog -Severity INFO -Action 'Grant admin consent' -Result 'Signing in to Microsoft Graph, in a separate process. Use an account that can administer app registrations; this account is also assigned as the application owner for Azure billing setup.'
    if (-not $script:UseDeviceCode) {
        Write-Host ''
        Write-Host '  A sign-in window opens now. Windows may place it behind this one, so check' -ForegroundColor Cyan
        Write-Host '  the taskbar if nothing seems to happen. Start again with -DeviceLogin to' -ForegroundColor Gray
        Write-Host '  sign in with a code instead.' -ForegroundColor Gray
    }
    $response = Invoke-GraphAction -Action 'Consent' -TenantId $TenantId -Arguments @{ClientId = $ClientId} `
        -Scopes @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All', 'User.Read')
    if ($response.Status -eq 'Granted') {
        if ([int]$response.Count -eq 0) {
            Write-RunLog -Severity WARN -Action 'Grant admin consent' -Result "The service principal exists, but no application permission was assigned. $($response.Message)"
            return $false
        }
        Write-RunLog -Severity SUCCESS -Action 'Grant admin consent' -Result "Consent is granted: $($response.Count) application permission(s) are assigned to $ClientId in tenant $TenantId, and $($response.Value) is verified as an application owner. Use that same account for Azure billing setup unless another billing operator is added as an owner."
        return $true
    }
    # Consent instructions would be nonsense for an application the directory does not have.
    if ("$($response.Message)" -match '(?i)Request_ResourceNotFound|does not exist') {
        Write-RunLog -Severity WARN -Action 'Grant admin consent' -Result "Application $ClientId does not exist in tenant $TenantId, so there is nothing to consent to. It has to be registered again."
        return $false
    }
    Add-RunFailure -FilePath '' -Action 'Grant admin consent' -Reason ([string]$response.Message)
    Show-AdminConsentInstruction -ClientId $ClientId
    return $false
}

function Get-AzureResourceTool {
    <# .SYNOPSIS Finds a tool able to create the Azure billing resource, preferring the CLI Microsoft documents. #>
    [CmdletBinding()]
    param()

    $azureCli = Get-Command az -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($azureCli) { return [pscustomobject]@{Kind = 'AzureCli'; Detail = [string]$azureCli.Source} }
    $module = Get-Module -ListAvailable -Name Az.Resources -ErrorAction SilentlyContinue |
        Sort-Object Version -Descending | Select-Object -First 1
    if ($module) { return [pscustomobject]@{Kind = 'AzPowerShell'; Detail = "Az.Resources $($module.Version)"} }
    return $null
}

function Write-AzureCliBillingDiagnostic {
    <# .SYNOPSIS Records the non-secret Azure CLI and billing request context needed for support. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Tool,
        [Parameter(Mandatory)][string]$SubscriptionId,
        [Parameter(Mandatory)][string]$ResourceGroup,
        [Parameter(Mandatory)][string]$ResourceName,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][AllowEmptyString()][string]$TenantId
    )

    $cliVersion = '<unavailable>'
    $extensionVersion = '<not installed yet>'
    try {
        $versionOutput = & az version --output json --only-show-errors 2>&1
        if ($LASTEXITCODE -eq 0) {
            $version = @((ConvertFrom-AzureCliJson -Output $versionOutput) | Select-Object -First 1)
            if ($version.Count -gt 0) {
                $reportedCliVersion = Get-ObjectPropertyValue -InputObject $version[0] -Names 'azure-cli'
                if (-not [string]::IsNullOrWhiteSpace($reportedCliVersion)) { $cliVersion = [string]$reportedCliVersion }
                $extensions = Get-ObjectPropertyValue -InputObject $version[0] -Names 'extensions'
                $reportedExtensionVersion = Get-ObjectPropertyValue -InputObject $extensions -Names 'graphservices'
                if (-not [string]::IsNullOrWhiteSpace($reportedExtensionVersion)) { $extensionVersion = [string]$reportedExtensionVersion }
            }
        }
        else { $cliVersion = "<version command exited $LASTEXITCODE>" }
    }
    catch { $cliVersion = "<version command failed: $($_.Exception.Message)>" }

    $applicationTenant = if ([string]::IsNullOrWhiteSpace($TenantId)) { '<not supplied>' } else { $TenantId }
    Write-RunLog -Severity INFO -Action 'Record Azure billing diagnostics' -NoConsole -Result `
        "Tool=Azure CLI $cliVersion; executable=$($Tool.Detail); graphservices extension=$extensionVersion; subscription=$SubscriptionId; application tenant=$applicationTenant; application=$ClientId; resource=$ResourceGroup/$ResourceName."
}

function Get-AzureBillingActivityEvidence {
    <# .SYNOPSIS Reads the failed ARM event for a billing request so the provider operation can be identified. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SubscriptionId,
        [Parameter(Mandatory)][string]$ResourceGroup,
        [Parameter(Mandatory)][string]$ResourceName
    )

    $resourceId = "/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroup/providers/Microsoft.GraphServices/accounts/$ResourceName"
    $lookupCommand = "az monitor activity-log list --resource-id '$resourceId' --subscription $SubscriptionId --offset 6h --status Failed --max-events 20 --select correlationId eventTimestamp operationId operationName resourceId status subStatus --output json --only-show-errors"
    try {
        $output = & az monitor activity-log list --resource-id $resourceId --subscription $SubscriptionId `
            --offset 6h --status Failed --max-events 20 `
            --select correlationId eventTimestamp operationId operationName resourceId status subStatus `
            --output json --only-show-errors 2>&1
        if ($LASTEXITCODE -ne 0) {
            $message = (("$output" -replace '\s+', ' ').Trim())
            Write-RunLog -Severity WARN -Action 'Read Azure billing activity' -NoConsole -Result `
                "The read-only Activity Log query failed: $message Command for support diagnostics: $lookupCommand"
            return [pscustomobject]@{CorrelationIds = @(); EventCount = 0; LookupCommand = $lookupCommand}
        }

        $events = @((ConvertFrom-AzureCliJson -Output $output) | Where-Object {
                [string]::Equals([string](Get-ObjectPropertyValue -InputObject $_ -Names 'resourceId'), $resourceId, [System.StringComparison]::OrdinalIgnoreCase)
            } | Sort-Object { [string](Get-ObjectPropertyValue -InputObject $_ -Names 'eventTimestamp') } -Descending)
        if ($events.Count -eq 0) {
            Write-RunLog -Severity INFO -Action 'Read Azure billing activity' -NoConsole -Result `
                "No matching failed event is visible yet. Azure Activity Log ingestion can lag. Read-only lookup command: $lookupCommand"
            return [pscustomobject]@{CorrelationIds = @(); EventCount = 0; LookupCommand = $lookupCommand}
        }

        $correlationIds = [System.Collections.Generic.List[string]]::new()
        foreach ($activityEvent in $events) {
            $correlationId = [string](Get-ObjectPropertyValue -InputObject $activityEvent -Names 'correlationId')
            $operationId = [string](Get-ObjectPropertyValue -InputObject $activityEvent -Names 'operationId')
            $timestamp = [string](Get-ObjectPropertyValue -InputObject $activityEvent -Names 'eventTimestamp')
            $operation = Get-ObjectPropertyValue -InputObject $activityEvent -Names 'operationName'
            $operationText = [string](Get-ObjectPropertyValue -InputObject $operation -Names 'localizedValue', 'value')
            $status = Get-ObjectPropertyValue -InputObject $activityEvent -Names 'status'
            $statusText = [string](Get-ObjectPropertyValue -InputObject $status -Names 'localizedValue', 'value')
            $subStatus = Get-ObjectPropertyValue -InputObject $activityEvent -Names 'subStatus'
            $subStatusText = [string](Get-ObjectPropertyValue -InputObject $subStatus -Names 'localizedValue', 'value')
            if (-not [string]::IsNullOrWhiteSpace($correlationId)) { $correlationIds.Add($correlationId) }
            Write-RunLog -Severity INFO -Action 'Read Azure billing activity' -NoConsole -Result `
                "Failed event timestamp=$timestamp; correlationId=$correlationId; operationId=$operationId; operation=$operationText; status=$statusText; subStatus=$subStatusText; resource=$resourceId."
        }
        return [pscustomobject]@{
            CorrelationIds = @($correlationIds | Select-Object -Unique)
            EventCount = $events.Count
            LookupCommand = $lookupCommand
        }
    }
    catch {
        Write-RunLog -Severity WARN -Action 'Read Azure billing activity' -NoConsole -Result `
            "The read-only Activity Log query could not be processed: $(Get-ErrorText -ErrorRecord $_) Command for support diagnostics: $lookupCommand"
        return [pscustomobject]@{CorrelationIds = @(); EventCount = 0; LookupCommand = $lookupCommand}
    }
}

function Install-AzureCommandLine {
    <# .SYNOPSIS Installs the Azure CLI, which is self-contained and so cannot hit the module assembly clash. #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param()

    $existing = Get-Command az -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($existing) { return [pscustomobject]@{Kind = 'AzureCli'; Detail = [string]$existing.Source} }
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-RunLog -Severity WARN -Action 'Install prerequisite' -Result 'winget is unavailable here, so the Azure CLI cannot be installed from this utility. Get it from https://aka.ms/installazurecliwindows'
        return $null
    }

    Write-Host ''
    Write-Host '  The Azure CLI is a self-contained application. It does not load the PowerShell' -ForegroundColor Gray
    Write-Host '  modules that failed, so the clash between them cannot reach it, and it is the tool' -ForegroundColor Gray
    Write-Host '  Microsoft documents for this step.' -ForegroundColor Gray
    $choice = Read-MenuChoice -Title 'Install the Azure CLI and finish the link with it?' -Options ([ordered]@{
            '1' = 'Yes, install it (default)'
            '2' = 'No'
        }) -Default '1'
    if ($choice -ne '1') { return $null }
    if (-not $PSCmdlet.ShouldProcess('Microsoft.AzureCLI', 'Install with winget')) { return $null }

    Write-RunLog -Severity INFO -Action 'Install prerequisite' -Result 'Installing the Azure CLI with winget. This takes a few minutes, and Windows may ask for administrator approval.'
    try {
        & winget install --id Microsoft.AzureCLI --exact --accept-package-agreements --accept-source-agreements
        if ($LASTEXITCODE -ne 0) { throw "winget exited with code $LASTEXITCODE." }
    }
    catch {
        Add-RunFailure -FilePath '' -Action 'Install prerequisite' -Reason (Get-ErrorText -ErrorRecord $_)
        return $null
    }

    # The installer edits PATH for new processes only, so this session has to pick it up itself.
    $env:Path = (@([Environment]::GetEnvironmentVariable('Path', 'Machine'), [Environment]::GetEnvironmentVariable('Path', 'User')) |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ';'
    $installed = Get-Command az -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $installed) {
        Write-RunLog -Severity WARN -Action 'Install prerequisite' -Result 'The Azure CLI installed, but this session cannot see it yet. Start the utility again and it will be used automatically.'
        return $null
    }
    Write-RunLog -Severity SUCCESS -Action 'Install prerequisite' -Result "Azure CLI is ready at $($installed.Source)."
    return [pscustomobject]@{Kind = 'AzureCli'; Detail = [string]$installed.Source}
}

function Show-MeteredBillingCommand {
    <# .SYNOPSIS Prints the documented commands so the link can be created by hand or by another operator. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][AllowEmptyString()][string]$SubscriptionId,
        [Parameter(Mandatory)][AllowEmptyString()][string]$ResourceGroup,
        [Parameter(Mandatory)][AllowEmptyString()][string]$ResourceName
    )

    $subscriptionText = if ([string]::IsNullOrWhiteSpace($SubscriptionId)) { '<subscription-id>' } else { $SubscriptionId }
    $groupText = if ([string]::IsNullOrWhiteSpace($ResourceGroup)) { '<resource-group>' } else { $ResourceGroup }
    $nameText = if ([string]::IsNullOrWhiteSpace($ResourceName)) { '<billing-resource-name>' } else { $ResourceName }
    Write-Host ''
    Write-Host '  Run these in Azure Cloud Shell (https://shell.azure.com) or a local Azure CLI:' -ForegroundColor Cyan
    Write-Host "    az extension add --name graphservices --allow-preview true --upgrade --yes" -ForegroundColor DarkGray
    Write-Host "    az account set --subscription $subscriptionText" -ForegroundColor DarkGray
    Write-Host "    az provider register --namespace Microsoft.GraphServices --subscription $subscriptionText --wait" -ForegroundColor DarkGray
    Write-Host "    az graph-services account create --resource-group $groupText --resource-name $nameText --subscription $subscriptionText --location global --app-id $ClientId" -ForegroundColor DarkGray
    Write-Host "    az resource list --resource-type Microsoft.GraphServices/accounts --subscription $subscriptionText" -ForegroundColor DarkGray
    Write-Host "    az resource show --resource-group $groupText --name $nameText --resource-type Microsoft.GraphServices/accounts --subscription $subscriptionText" -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '  The resource group must already exist, you need Contributor on the subscription,' -ForegroundColor Gray
    Write-Host '  and owner rights on the application registration.' -ForegroundColor Gray
}

function Get-MeteredBillingResourceState {
    <# .SYNOPSIS Validates that an ARM billing resource is provisioned for the intended Entra application. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowNull()][object]$Resource,
        [Parameter(Mandatory)][string]$ClientId
    )

    if ($null -eq $Resource) {
        return [pscustomobject]@{Status = 'Missing'; Message = 'The billing resource was not found.'}
    }
    $propertiesProperty = $Resource.PSObject.Properties['properties']
    $properties = if ($propertiesProperty) { $propertiesProperty.Value } else { $null }
    $actualClientId = Get-ObjectPropertyValue -InputObject $properties -Names 'appId', 'AppId'
    $billingPlanId = Get-ObjectPropertyValue -InputObject $properties -Names 'billingPlanId', 'BillingPlanId'
    $resourceType = Get-ObjectPropertyValue -InputObject $Resource -Names 'type', 'Type'
    $location = Get-ObjectPropertyValue -InputObject $Resource -Names 'location', 'Location'
    $provisioningState = Get-ObjectPropertyValue -InputObject $properties -Names 'provisioningState', 'ProvisioningState'
    if ([string]::IsNullOrWhiteSpace($provisioningState)) {
        $provisioningState = Get-ObjectPropertyValue -InputObject $Resource -Names 'provisioningState', 'ProvisioningState'
    }
    if ([string]::IsNullOrWhiteSpace($actualClientId)) {
        return [pscustomobject]@{Status = 'NotReady'; Message = 'The billing resource does not report an application ID.'}
    }
    if (-not [string]::Equals($actualClientId, $ClientId, [System.StringComparison]::OrdinalIgnoreCase)) {
        return [pscustomobject]@{Status = 'WrongApplication'; Message = "The billing resource is linked to application $actualClientId, not $ClientId."}
    }
    if ([string]::IsNullOrWhiteSpace($billingPlanId)) {
        return [pscustomobject]@{Status = 'NotReady'; Message = 'The billing resource does not report a billing plan ID.'}
    }
    if (-not [string]::Equals($resourceType, 'Microsoft.GraphServices/accounts', [System.StringComparison]::OrdinalIgnoreCase)) {
        return [pscustomobject]@{Status = 'NotReady'; Message = "Azure returned resource type '$resourceType', not Microsoft.GraphServices/accounts."}
    }
    if (-not [string]::Equals($location, 'global', [System.StringComparison]::OrdinalIgnoreCase)) {
        return [pscustomobject]@{Status = 'NotReady'; Message = "The billing resource reports location '$location', not Global."}
    }
    if ([string]::IsNullOrWhiteSpace($provisioningState)) {
        return [pscustomobject]@{Status = 'NotReady'; Message = 'The billing resource does not report a provisioning state.'}
    }
    if ($provisioningState -ne 'Succeeded') {
        return [pscustomobject]@{Status = 'NotReady'; Message = "The billing resource reports state '$provisioningState' rather than Succeeded."}
    }
    return [pscustomobject]@{Status = 'Ready'; Message = "The billing resource is linked to application $ClientId, has billing plan $billingPlanId, and is provisioned."}
}

function Test-MeteredBillingResource {
    <# .SYNOPSIS Reports whether the billing resource is provisioned for the intended Entra application. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SubscriptionId,
        [Parameter(Mandatory)][string]$ResourceGroup,
        [Parameter(Mandatory)][string]$ResourceName,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][object]$Tool
    )

    try {
        if ($Tool.Kind -eq 'AzureCli') {
            # Microsoft documents list for the resource state, then show for the linked application ID.
            $listOutput = & az resource list --resource-type 'Microsoft.GraphServices/accounts' `
                --subscription $SubscriptionId --output json --only-show-errors 2>&1
            if ($LASTEXITCODE -ne 0) {
                $message = ("$listOutput" -replace '\s+', ' ').Trim()
                $script:LastBillingResourceStatus = [pscustomobject]@{Status = 'LookupFailed'; Message = $message}
                return $false
            }
            $match = @((ConvertFrom-AzureCliJson -Output $listOutput) | Where-Object {
                    [string]::Equals((Get-ObjectPropertyValue -InputObject $_ -Names 'name', 'Name'), $ResourceName, [System.StringComparison]::OrdinalIgnoreCase) -and
                    [string]::Equals((Get-ObjectPropertyValue -InputObject $_ -Names 'resourceGroup', 'ResourceGroup'), $ResourceGroup, [System.StringComparison]::OrdinalIgnoreCase)
                } | Select-Object -First 1)
            if ($match.Count -eq 0) {
                $script:LastBillingResourceStatus = [pscustomobject]@{Status = 'Missing'; Message = 'The billing resource was not found in the selected subscription.'}
                return $false
            }
            $output = & az resource show --resource-group $ResourceGroup --name $ResourceName `
                --resource-type 'Microsoft.GraphServices/accounts' --subscription $SubscriptionId --output json --only-show-errors 2>&1
            if ($LASTEXITCODE -ne 0) {
                $message = ("$output" -replace '\s+', ' ').Trim()
                $status = if ($message -match '(?i)ResourceNotFound|could not be found|not found') { 'Missing' } else { 'LookupFailed' }
                $script:LastBillingResourceStatus = [pscustomobject]@{Status = $status; Message = $message}
                return $false
            }
            $resource = ((@($output) -join [Environment]::NewLine) | ConvertFrom-Json -ErrorAction Stop)
            $resource | Add-Member -NotePropertyName provisioningState -NotePropertyValue (Get-ObjectPropertyValue -InputObject $match[0] -Names 'provisioningState', 'ProvisioningState') -Force
            $script:LastBillingResourceStatus = Get-MeteredBillingResourceState -Resource $resource -ClientId $ClientId
            if ($script:LastBillingResourceStatus.Status -ne 'Ready') {
                Write-RunLog -Severity WARN -Action 'Check metered billing' -Result $script:LastBillingResourceStatus.Message
            }
            return ($script:LastBillingResourceStatus.Status -eq 'Ready')
        }
        $script:LastBillingResourceStatus = [pscustomobject]@{Status = 'LookupFailed'; Message = 'The Azure PowerShell fallback cannot verify the linked application ID.'}
        return $false
    }
    catch {
        $script:LastBillingResourceStatus = [pscustomobject]@{Status = 'LookupFailed'; Message = (Get-ErrorText -ErrorRecord $_)}
        Write-Verbose "Billing resource lookup reported: $($script:LastBillingResourceStatus.Message)"
        return $false
    }
}

function Test-MeteredBillingPreflight {
    <# .SYNOPSIS Checks the Azure conditions that can be proved before creating the billing association. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Tool,
        [Parameter(Mandatory)][string]$SubscriptionId,
        [Parameter(Mandatory)][string]$ResourceGroup,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][AllowEmptyString()][string]$TenantId
    )

    if ($Tool.Kind -ne 'AzureCli') {
        return [pscustomobject]@{Ready = $false; FailureKind = 'Prerequisite'; ApiVersions = @(); Message = 'This preflight must run inside the isolated Az worker for the Azure PowerShell route.'}
    }

    $problems = [System.Collections.Generic.List[string]]::new()
    $providerApiVersions = @()
    $parsedClientId = [guid]::Empty
    if (-not [guid]::TryParse($ClientId, [ref]$parsedClientId) -or $parsedClientId -eq [guid]::Empty) {
        $problems.Add("Application $ClientId is not a valid client ID.")
    }

    try {
        $cloudName = & az cloud show --query name --output tsv --only-show-errors 2>&1
        if ($LASTEXITCODE -ne 0) {
            $problems.Add("Azure could not read the active cloud: $(("$cloudName" -replace '\s+', ' ').Trim())")
        }
        else {
            $cloudName = "$cloudName".Trim()
            Write-RunLog -Severity INFO -Action 'Record Azure billing diagnostics' -NoConsole -Result "Active Azure cloud=$cloudName."
            if ($cloudName -ne 'AzureCloud') {
                $problems.Add("Azure CLI is set to '$cloudName'. Metered Graph APIs require the global AzureCloud environment.")
            }
        }

        $accountOutput = & az account show --subscription $SubscriptionId --output json --only-show-errors 2>&1
        if ($LASTEXITCODE -ne 0) {
            $problems.Add("Azure could not read subscription ${SubscriptionId}: $(("$accountOutput" -replace '\s+', ' ').Trim())")
        }
        else {
            $account = @((ConvertFrom-AzureCliJson -Output $accountOutput) | Select-Object -First 1)
            if ($account.Count -eq 0) {
                $problems.Add("Azure did not return details for subscription $SubscriptionId.")
            }
            else {
                $accountTenantId = Get-ObjectPropertyValue -InputObject $account[0] -Names 'tenantId', 'homeTenantId'
                $accountState = Get-ObjectPropertyValue -InputObject $account[0] -Names 'state'
                Write-RunLog -Severity INFO -Action 'Record Azure billing diagnostics' -NoConsole -Result `
                    "Subscription state=$accountState; subscription tenant=$accountTenantId."
                if (-not [string]::IsNullOrWhiteSpace($accountState) -and $accountState -ne 'Enabled') {
                    $problems.Add("Subscription $SubscriptionId is in state '$accountState', not Enabled.")
                }
                if (-not [string]::IsNullOrWhiteSpace($TenantId) -and -not [string]::IsNullOrWhiteSpace($accountTenantId) -and
                    -not [string]::Equals($accountTenantId, $TenantId, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $problems.Add("Subscription $SubscriptionId belongs to tenant $accountTenantId, not application tenant $TenantId.")
                }
            }
        }

        $operatorOutput = & az ad signed-in-user show --query id --output tsv --only-show-errors 2>&1
        if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace("$operatorOutput")) {
            $operatorId = "$operatorOutput".Trim()
            $ownerOutput = & az ad app owner list --id $ClientId --query '[].id' --output tsv --only-show-errors 2>&1
            if ($LASTEXITCODE -eq 0) {
                $ownerIds = @(("$ownerOutput" -split '\r?\n') | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                if ($ownerIds -contains $operatorId) {
                    Write-RunLog -Severity INFO -Action 'Record Azure billing diagnostics' -NoConsole -Result 'Application owner check=succeeded for the signed-in Azure CLI user.'
                }
                else {
                    $problems.Add("The signed-in Azure CLI user (object ID $operatorId) is not listed as an owner of application $ClientId. Azure subscription ownership and Entra application ownership are separate permissions; subscription Owner does not grant application ownership. Sign Azure CLI in as an application owner that also has Contributor or Owner on subscription $SubscriptionId, or ask a current application owner or authorized Entra administrator to run: az ad app owner add --id $ClientId --owner-object-id $operatorId. Then rerun option 3 and keep the existing application, certificate, resource group, and billing-resource name; do not register another application.")
                }
            }
            else {
                Write-RunLog -Severity WARN -Action 'Check application owner' -Result "Azure CLI could not list owners for application ${ClientId}: $(("$ownerOutput" -replace '\s+', ' ').Trim()). The documented create command will remain the authoritative check."
            }

            $subscriptionScope = "/subscriptions/$SubscriptionId"
            $roleOutput = & az role assignment list --assignee-object-id $operatorId --include-groups --include-inherited `
                --scope $subscriptionScope --fill-principal-name false --output json --only-show-errors 2>&1
            if ($LASTEXITCODE -eq 0) {
                $assignments = @(ConvertFrom-AzureCliJson -Output $roleOutput)
                $hasContributorAccess = @($assignments | Where-Object {
                        $roleName = [string](Get-ObjectPropertyValue -InputObject $_ -Names 'roleDefinitionName')
                        $roleId = [string](Get-ObjectPropertyValue -InputObject $_ -Names 'roleDefinitionId')
                        $roleName -in 'Contributor', 'Owner' -or
                            $roleId -match '(?i)/(b24988ac-6180-42a0-ab88-20f7382dd24c|8e3af657-a8ff-443c-a75c-2fe8c4bcb635)$'
                    }).Count -gt 0
                if ($hasContributorAccess) {
                    Write-RunLog -Severity INFO -Action 'Record Azure billing diagnostics' -NoConsole -Result 'Subscription role check=succeeded with Contributor or Owner access, including inherited and group assignments.'
                }
                else {
                    $problems.Add("The signed-in Azure CLI user has no Contributor or Owner role at subscription $SubscriptionId, including inherited and group assignments. Microsoft requires Contributor access for the billing association.")
                }
            }
            else {
                Write-RunLog -Severity WARN -Action 'Check subscription role' -Result "Azure CLI could not list effective role assignments at ${subscriptionScope}: $(("$roleOutput" -replace '\s+', ' ').Trim()). The documented create command will remain the authoritative check."
            }
        }
        else {
            Write-RunLog -Severity WARN -Action 'Check application owner' -Result "Azure CLI could not identify a signed-in user, so application ownership could not be verified independently. The documented create command will remain the authoritative check. $(("$operatorOutput" -replace '\s+', ' ').Trim())"
        }

        $groupOutput = & az group show --name $ResourceGroup --subscription $SubscriptionId `
            --query '{location:location,provisioningState:properties.provisioningState}' --output json --only-show-errors 2>&1
        if ($LASTEXITCODE -ne 0) {
            $problems.Add("Resource group '$ResourceGroup' cannot be read in subscription ${SubscriptionId}: $(("$groupOutput" -replace '\s+', ' ').Trim())")
        }
        else {
            $group = @((ConvertFrom-AzureCliJson -Output $groupOutput) | Select-Object -First 1)
            if ($group.Count -eq 0) {
                $problems.Add("Azure returned no readable details for resource group '$ResourceGroup'.")
            }
            else {
                $groupLocation = [string](Get-ObjectPropertyValue -InputObject $group[0] -Names 'location')
                $groupState = [string](Get-ObjectPropertyValue -InputObject $group[0] -Names 'provisioningState')
                Write-RunLog -Severity INFO -Action 'Record Azure billing diagnostics' -NoConsole -Result `
                    "Resource group=$ResourceGroup; location=$groupLocation; provisioning state=$groupState."
                if ($groupState -ne 'Succeeded') {
                    $reportedState = if ([string]::IsNullOrWhiteSpace($groupState)) { '<not reported>' } else { $groupState }
                    $problems.Add("Resource group '$ResourceGroup' is in provisioning state '$reportedState', not Succeeded.")
                }
            }
        }

        $providerOutput = & az provider show --namespace Microsoft.GraphServices --subscription $SubscriptionId `
            --output json --only-show-errors 2>&1
        if ($LASTEXITCODE -ne 0) {
            $problems.Add("Azure could not read Microsoft.GraphServices registration: $(($providerOutput -replace '\s+', ' ').Trim())")
        }
        else {
            $provider = @((ConvertFrom-AzureCliJson -Output $providerOutput) | Select-Object -First 1)
            if ($provider.Count -eq 0) {
                $problems.Add('Azure returned no readable details for Microsoft.GraphServices registration.')
            }
            else {
                $providerState = [string](Get-ObjectPropertyValue -InputObject $provider[0] -Names 'registrationState')
                $accountResourceType = @(@(Get-RawObjectPropertyValue -InputObject $provider[0] -Names 'resourceTypes') | Where-Object {
                        [string](Get-ObjectPropertyValue -InputObject $_ -Names 'resourceType') -eq 'accounts'
                    } | Select-Object -First 1)
                $providerApiVersions = @(if ($accountResourceType.Count -gt 0) {
                    (Get-RawObjectPropertyValue -InputObject $accountResourceType[0] -Names 'apiVersions') |
                            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                    })
                $apiVersionText = if ($providerApiVersions.Count -gt 0) { $providerApiVersions -join ', ' } else { '<not advertised>' }
                Write-RunLog -Severity INFO -Action 'Record Azure billing diagnostics' -NoConsole -Result `
                    "Microsoft.GraphServices registration=$providerState; accounts API versions=$apiVersionText."
                if ($providerState -ne 'Registered') {
                    $problems.Add("Microsoft.GraphServices is '$providerState', not Registered.")
                }
            }
        }

    }
    catch {
        $problems.Add((Get-ErrorText -ErrorRecord $_))
    }

    if ($problems.Count -gt 0) {
        $message = $problems -join ' '
        $failureKind = if (Test-AzureProviderImplementationFailure -Message $message) { 'ProviderFault' } else { 'Prerequisite' }
        return [pscustomobject]@{Ready = $false; FailureKind = $failureKind; ApiVersions = @($providerApiVersions); Message = $message}
    }
    return [pscustomobject]@{Ready = $true; FailureKind = 'None'; ApiVersions = @($providerApiVersions); Message = "Global Azure, subscription $SubscriptionId, resource group '$ResourceGroup', Microsoft.GraphServices, application ownership, and subscription Contributor access are ready for the documented billing association command."}
}

function Invoke-AzureAction {
    <# .SYNOPSIS Runs Azure work in a separate process, because Az ships Microsoft.Extensions assemblies that break PnP. #>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'resultLine', Justification = 'Assigned inside a ForEach-Object block, which runs in this scope.')]
    param(
        [Parameter(Mandatory)][ValidateSet('ListSubscriptions', 'ListResourceGroups', 'CreateResourceGroup', 'SignIn', 'SignOut', 'BillingLink', 'BillingExists')][string]$Action,
        [hashtable]$Arguments = @{},
        [AllowEmptyString()][string]$TenantId = ''
    )

    $hostPath = ''
    try {
        $current = [string](Get-Process -Id $PID -ErrorAction Stop).Path
        if (-not [string]::IsNullOrWhiteSpace($current) -and $current -match '(?i)pwsh(\.exe)?$') { $hostPath = $current }
    }
    catch { Write-Verbose "Could not read this host's path: $($_.Exception.Message)" }
    if ([string]::IsNullOrWhiteSpace($hostPath)) { $hostPath = Get-PowerShell7Path -MinimumVersion ([version]'7.2.0') }
    if ([string]::IsNullOrWhiteSpace($hostPath)) {
        return [pscustomobject]@{Status = 'Error'; Message = 'PowerShell 7 is needed to talk to Azure separately from PnP.'; Value = @(); Correlation = @()}
    }

    $workerPath = Join-Path ([System.IO.Path]::GetTempPath()) ('PurviewFileLabelingAzure-' + $PID + '-' + [guid]::NewGuid().ToString('N') + '.ps1')
    $worker = @'
param([string]$Action, [string]$ArgumentJson, [string]$TenantId)
$ErrorActionPreference = 'Stop'
$result = [ordered]@{ Status = 'Error'; Message = ''; Value = @(); Correlation = @(); Signature = '' }
function Get-Signature { param([string]$Text)
    $t = [regex]::Replace($Text, '\\u(?<c>[0-9a-fA-F]{4})', { param($m) [string][char][Convert]::ToInt32($m.Groups['c'].Value, 16) })
    $t = [regex]::Replace($t, '(?i)(correlationid|request-id|client-request-id|date)\s*:\s*\S*', '')
    $t = [regex]::Replace($t, '(?i)[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}', '')
    # A deployment reports the provider's fault wrapped in its own words, which is the same fault.
    $t = [regex]::Replace($t, "(?i)the deployment '[^']*' failed with error\(s\)\.?", '')
    $t = [regex]::Replace($t, '(?i)showing \d+ out of \d+ error\(s\)\.?', '')
    $t = [regex]::Replace($t, '(?i)status (message|code)\s*:', '')
    # The error code appears as a prefix on one route and a suffix on another, so it is removed entirely.
    $code = ''
    $bracketed = [regex]::Match($t, '(?i)\(\s*code\s*:\s*(?<c>[A-Za-z]+)\s*\)')
    if ($bracketed.Success) { $code = $bracketed.Groups['c'].Value }
    else {
        $prefixed = [regex]::Match($t, '(?i)^\s*(?<c>[A-Za-z]+Error)\s*:')
        if ($prefixed.Success) { $code = $prefixed.Groups['c'].Value }
    }
    if (-not [string]::IsNullOrWhiteSpace($code)) { $t = [regex]::Replace($t, '(?i)' + [regex]::Escape($code), '') }
    $t = [regex]::Replace($t, '(?i)\(\s*code\s*:?\s*\)', '')
    return ([regex]::Replace($t, '[^a-zA-Z]', '')).ToLowerInvariant()
}
function Add-Signature { param($Set, [string]$Signature)
    if ([string]::IsNullOrWhiteSpace($Signature)) { return }
    foreach ($known in @($Set)) {
        # One route's text can contain another's verbatim, which still means a single underlying fault.
        if ($known -eq $Signature -or $known.Contains($Signature) -or $Signature.Contains($known)) { return }
    }
    $null = $Set.Add($Signature)
}
function Import-AzResourcesModule {
    if ($script:AzResourcesVersion) { Import-Module Az.Resources -RequiredVersion $script:AzResourcesVersion -ErrorAction Stop }
    else { Import-Module Az.Resources -ErrorAction Stop }
}
function Get-Value { param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) { if ($Object.Contains($Name)) { return $Object[$Name] } return $null }
    $property = $Object.PSObject.Properties[$Name]
    if ($property) { return $property.Value }
    return $null
}
function Get-BillingResource { param([string]$ResourceGroup, [string]$ResourceName)
    try {
        $found = @(Get-AzResource -ResourceType 'Microsoft.GraphServices/accounts' -ResourceGroupName $ResourceGroup `
                -Name $ResourceName -ExpandProperties -ErrorAction SilentlyContinue | Select-Object -First 1)
        if ($found.Count -eq 0) { return $null }
        $resource = $found[0]
        $properties = Get-Value $resource 'Properties'
        $state = Get-Value $properties 'provisioningState'
        if ([string]::IsNullOrWhiteSpace([string]$state)) { $state = Get-Value $resource 'ProvisioningState' }
        return [ordered]@{
            type = [string](Get-Value $resource 'ResourceType')
            location = [string](Get-Value $resource 'Location')
            properties = [ordered]@{
                appId = [string](Get-Value $properties 'appId')
                billingPlanId = [string](Get-Value $properties 'billingPlanId')
                provisioningState = [string]$state
            }
        }
    }
    catch { return $null }
}
try {
    $arguments = if ([string]::IsNullOrWhiteSpace($ArgumentJson)) { @{} } else { ConvertFrom-Json $ArgumentJson -AsHashtable }
    # Az.Resources records the exact Az.Accounts build it was compiled against. Importing Az.Accounts
    # first loads the newest instead, and a mismatched pair fails later inside OpenTelemetry with a
    # type-load error that names nothing useful.
    $resourcesModule = Get-Module -ListAvailable -Name Az.Resources -ErrorAction SilentlyContinue |
        Sort-Object Version -Descending | Select-Object -First 1
    $script:AzResourcesVersion = ''
    $accountsVersion = ''
    if ($resourcesModule) {
        $script:AzResourcesVersion = [string]$resourcesModule.Version
        foreach ($required in @($resourcesModule.RequiredModules)) {
            if ($required.Name -eq 'Az.Accounts' -and $required.Version) { $accountsVersion = [string]$required.Version }
        }
    }
    $accountsImported = $false
    if ($accountsVersion) {
        try {
            Import-Module Az.Accounts -RequiredVersion $accountsVersion -ErrorAction Stop
            $accountsImported = $true
        }
        catch { "PROGRESS:Az.Accounts $accountsVersion is not installed, so the newest is used instead." }
    }
    if (-not $accountsImported) { Import-Module Az.Accounts -ErrorAction Stop }
    $context = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $context -and $Action -notin 'SignIn', 'SignOut') { $result.Status = 'NoSession'; throw 'Not signed in to Azure.' }

    switch ($Action) {
        'SignIn' {
            if ($context) { $result.Status = 'Reused' }
            else {
                $p = @{ ErrorAction = 'Stop' }
                if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $p.Tenant = $TenantId }
                $null = Connect-AzAccount @p
                $result.Status = 'SignedIn'
            }
        }
        'SignOut' {
            if ($context) { $null = Disconnect-AzAccount -AzureContext $context -Scope CurrentUser -Confirm:$false -ErrorAction SilentlyContinue }
            $result.Status = 'SignedOut'
        }
        'ListSubscriptions' {
            $result.Value = @(Get-AzSubscription -ErrorAction Stop | ForEach-Object {
                @{ Name = [string]$_.Name; Id = [string]$_.Id; TenantId = [string]$_.TenantId } })
            $result.Status = 'Listed'
        }
        'ListResourceGroups' {
            $null = Set-AzContext -Subscription $arguments['SubscriptionId'] -ErrorAction Stop
            Import-AzResourcesModule
            $groups = @(Get-AzResourceGroup -ErrorAction Stop)
            $result.Value = @($groups | ForEach-Object {
                    @{ Name = [string]$_.ResourceGroupName; Location = [string]$_.Location; State = [string]$_.ProvisioningState }
                })
            $result.Status = 'Listed'
        }
        'CreateResourceGroup' {
            $null = Set-AzContext -Subscription $arguments['SubscriptionId'] -ErrorAction Stop
            Import-AzResourcesModule
            $null = New-AzResourceGroup -Name $arguments['Name'] -Location $arguments['Location'] -Force -ErrorAction Stop
            $result.Status = 'Created'
        }
        'BillingExists' {
            $null = Set-AzContext -Subscription $arguments['SubscriptionId'] -ErrorAction Stop
            Import-AzResourcesModule
            $existing = Get-BillingResource -ResourceGroup $arguments['ResourceGroup'] -ResourceName $arguments['ResourceName']
            if ($existing) { $result.Value = @($existing) }
            $result.Status = if ($existing) { 'Present' } else { 'Missing' }
        }
        'BillingLink' {
            $subscriptionId = [string]$arguments['SubscriptionId']
            $resourceGroup = [string]$arguments['ResourceGroup']
            $resourceName = [string]$arguments['ResourceName']
            $clientId = [string]$arguments['ClientId']
            $selected = Set-AzContext -Subscription $subscriptionId -ErrorAction Stop
            Import-AzResourcesModule

            $environmentName = [string]$selected.Environment.Name
            if ($environmentName -ne 'AzureCloud') {
                $result.Status = 'PreflightFailed'
                $result.Message = "Azure is set to '$environmentName'. Metered Graph APIs require the global AzureCloud environment."
                break
            }
            $subscription = @(Get-AzSubscription -SubscriptionId $subscriptionId -ErrorAction Stop | Select-Object -First 1)
            if ($subscription.Count -eq 0) {
                $result.Status = 'PreflightFailed'
                $result.Message = "Azure could not read subscription $subscriptionId."
                break
            }
            if ([string]$subscription[0].State -ne 'Enabled') {
                $result.Status = 'PreflightFailed'
                $result.Message = "Subscription $subscriptionId is in state '$($subscription[0].State)', not Enabled."
                break
            }
            if (-not [string]::IsNullOrWhiteSpace($TenantId) -and -not [string]::IsNullOrWhiteSpace([string]$subscription[0].TenantId) -and
                -not [string]::Equals([string]$subscription[0].TenantId, $TenantId, [System.StringComparison]::OrdinalIgnoreCase)) {
                $result.Status = 'PreflightFailed'
                $result.Message = "Subscription $subscriptionId belongs to tenant $($subscription[0].TenantId), not application tenant $TenantId."
                break
            }
            $group = @()
            try { $group = @(Get-AzResourceGroup -Name $resourceGroup -ErrorAction Stop | Select-Object -First 1) }
            catch {
                $result.Status = 'PreflightFailed'
                $result.Message = "Resource group '$resourceGroup' cannot be read in subscription ${subscriptionId}: $($_.Exception.Message)"
                break
            }
            $groupState = if ($group.Count -eq 0) { '' } else { [string]$group[0].ProvisioningState }
            if ($groupState -ne 'Succeeded') {
                $reportedState = if ([string]::IsNullOrWhiteSpace($groupState)) { '<not reported>' } else { $groupState }
                $result.Status = 'PreflightFailed'
                $result.Message = "Resource group '$resourceGroup' is in provisioning state '$reportedState', not Succeeded."
                break
            }

            function Get-Existing { return Get-BillingResource -ResourceGroup $resourceGroup -ResourceName $resourceName }
            'PROGRESS:Checking whether the billing resource already exists.'
            $existing = Get-Existing
            if ($existing) { $result.Value = @($existing); $result.Status = 'Exists'; throw 'done' }

            # An unregistered namespace is simply absent, so absence is the signal to register.
            $registered = $false
            for ($wait = 1; $wait -le 6; $wait++) {
                $state = ''
                try {
                    $p = @(Get-AzResourceProvider -ProviderNamespace Microsoft.GraphServices -ErrorAction Stop)
                    if ($p.Count -gt 0) { $state = [string]$p[0].RegistrationState }
                }
                catch { }
                if ($state -eq 'Registered') { $registered = $true; break }
                if ($wait -eq 1) {
                    'PROGRESS:Registering the Microsoft.GraphServices provider on the subscription.'
                    try { $null = Register-AzResourceProvider -ProviderNamespace Microsoft.GraphServices -ErrorAction Stop } catch { }
                }
                "PROGRESS:Waiting for the provider to report itself registered ($wait of 6)."
                Start-Sleep -Seconds 10
            }
            if (-not $registered) {
                $result.Status = 'PreflightFailed'
                $result.Message = 'Microsoft.GraphServices did not report Registered after provider registration.'
                break
            }
            $result.Message = 'Global Azure, subscription, resource group, and Microsoft.GraphServices are ready.'

            $versions = @()
            try {
                foreach ($p in @(Get-AzResourceProvider -ProviderNamespace Microsoft.GraphServices -ErrorAction Stop)) {
                    foreach ($t in @($p.ResourceTypes)) {
                        if ([string]$t.ResourceTypeName -ne 'accounts') { continue }
                        $versions = @($t.ApiVersions)
                    }
                }
            }
            catch { }
            $usable = @(@($versions) | Where-Object { "$_" -match '^\d{4}-\d{2}-\d{2}' } | ForEach-Object { "$_" })
            if ($usable.Count -eq 0) { $usable = @('2023-04-13', '2022-09-22-preview') }
            $ordered = @(@($usable | Where-Object { $_ -notmatch '(?i)preview' } | Sort-Object -Descending -Unique)) +
                       @(@($usable | Where-Object { $_ -match '(?i)preview' } | Sort-Object -Descending -Unique))

            $routes = @()
            $routes += @{ Kind = 'Cmdlet'; Api = $ordered[0] }
            $routes += @{ Kind = 'Rest'; Api = $ordered[0] }
            $routes += @{ Kind = 'Template'; Api = $ordered[0] }
            # Each API version reaches a different service code path, so a fault on one is not a fault on all.
            foreach ($v in @($ordered | Select-Object -Skip 1 -First 2)) {
                $routes += @{ Kind = 'Cmdlet'; Api = $v }
                $routes += @{ Kind = 'Rest'; Api = $v }
                $routes += @{ Kind = 'Template'; Api = $v }
            }

            $signatures = [System.Collections.Generic.HashSet[string]]::new()
            $kinds = [System.Collections.Generic.HashSet[string]]::new()
            $apis = [System.Collections.Generic.HashSet[string]]::new()
            $correlation = [System.Collections.Generic.List[string]]::new()
            $created = $false
            $last = ''
            "PROGRESS:Provider API version(s) available: $($ordered -join ', ')."
            foreach ($route in $routes) {
                if ($created) { break }
                "PROGRESS:Trying a $($route.Kind) request on API version $($route.Api)."
                try {
                    if ($route.Kind -eq 'Rest') {
                        $token = Get-AzAccessToken -ResourceUrl 'https://management.azure.com/' -ErrorAction Stop
                        $secret = $token.Token
                        $plain = ''
                        if ($secret -is [System.Security.SecureString]) {
                            $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secret)
                            try { $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr) } finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) }
                        }
                        else { $plain = [string]$secret }
                        $uri = 'https://management.azure.com/subscriptions/{0}/resourceGroups/{1}/providers/Microsoft.GraphServices/accounts/{2}?api-version={3}' -f $subscriptionId, $resourceGroup, $resourceName, $route.Api
                        $body = @{ location = 'global'; properties = @{ appId = $clientId } } | ConvertTo-Json -Depth 4
                        try { $null = Invoke-RestMethod -Method Put -Uri $uri -Headers @{ Authorization = "Bearer $plain" } -Body $body -ContentType 'application/json' -TimeoutSec 60 -ErrorAction Stop }
                        finally { $plain = '' }
                    }
                    elseif ($route.Kind -eq 'Template') {
                        $template = @{
                            '$schema' = 'https://schema.management.azure.com/schemas/2019-04-01/deploymentTemplate.json#'
                            contentVersion = '1.0.0.0'
                            resources = @(@{ type = 'Microsoft.GraphServices/accounts'; apiVersion = $route.Api; name = $resourceName; location = 'global'; properties = @{ appId = $clientId } })
                        }
                        $null = New-AzResourceGroupDeployment -ResourceGroupName $resourceGroup -TemplateObject $template -Name ('purview-file-labeling-billing-' + (Get-Date -Format 'yyyyMMddHHmmss')) -ErrorAction Stop
                    }
                    else {
                        $null = New-AzResource -ResourceType 'Microsoft.GraphServices/accounts' -ResourceGroupName $resourceGroup -Name $resourceName -Location global -Properties @{ appId = $clientId } -ApiVersion $route.Api -Force -ErrorAction Stop
                    }
                    $existing = Get-Existing
                    if ($existing) { $result.Value = @($existing); $created = $true }
                    else { $last = 'The create request completed, but Azure did not return the billing resource during read-back.'; break }
                }
                catch {
                    $last = "$($_.Exception.Message)"
                    if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $last = "$($_.ErrorDetails.Message) $last" }
                    $null = Add-Signature -Set $signatures -Signature (Get-Signature -Text $last)
                    $null = $kinds.Add([string]$route.Kind)
                    $null = $apis.Add([string]$route.Api)
                    foreach ($m in [regex]::Matches($last, '(?i)CorrelationId:\s*(?<id>[0-9a-f-]{36})')) { $null = $correlation.Add($m.Groups['id'].Value) }
                    $existing = Get-Existing
                    if ($existing) { $result.Value = @($existing); $created = $true; break }

                    $reason = ($last -replace '\s+', ' ').Trim()
                    if ($reason.Length -gt 220) { $reason = $reason.Substring(0, 220) + '...' }
                    "PROGRESS:That route was refused: $reason"

                    if ($last -match '(?is)(TryCreateLogger|OpenTelemetry(?:\.Logs\.)?LoggerProviderSdk).*cannot\W*reduce\W*access') {
                        $result.Status = 'ProviderFault'
                        $result.Message = $last
                        'PROGRESS:Microsoft.GraphServices returned its known provider implementation fault. No other create route will be attempted.'
                        break
                    }
                    # A .NET loader error is this process failing to load the Azure module, not Azure
                    # refusing the request. Every route uses that module, so retrying cannot help.
                    if ($last -match '(?i)cannot reduce access|typeload|could not load file or assembly|does not have an implementation|methodaccess|ambiguous match') {
                        $result.Status = 'ClientFault'
                        $result.Message = $last
                        'PROGRESS:That failure came from loading the Azure module here, not from Azure.'
                        break
                    }
                    if ($kinds.Count -ge 3 -and $apis.Count -ge 2 -and $signatures.Count -eq 1) {
                        'PROGRESS:Every route on every API version returned the same fault, so the provider itself is refusing. Stopping.'
                        break
                    }
                    Start-Sleep -Seconds 5
                }
            }
            $result.Correlation = @($correlation | Select-Object -Unique)
            $result.Signature = if ($signatures.Count -gt 0) { @($signatures)[0] } else { '' }
            if ($created) { $result.Status = 'Created' }
            elseif ($result.Status -notin 'ClientFault', 'ProviderFault') { $result.Status = 'Failed'; $result.Message = $last }
        }
    }
}
catch {
    if ($_.Exception.Message -ne 'done') {
        if ($result.Status -eq 'Error') {
            $result.Message = "$($_.Exception.Message)"
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $result.Message = "$($_.ErrorDetails.Message) $($result.Message)" }
        }
    }
}
'RESULT:' + ($result | ConvertTo-Json -Compress -Depth 8)
'@

    try {
        Set-Content -LiteralPath $workerPath -Value $worker -Encoding utf8 -ErrorAction Stop
        $argumentJson = if ($Arguments.Count -gt 0) { $Arguments | ConvertTo-Json -Compress -Depth 6 } else { '' }
        $resultLine = ''
        $other = [System.Collections.Generic.List[string]]::new()
        # Streamed rather than captured, so a minute of provider registration does not look like a hang.
        & $hostPath '-NoLogo' '-NoProfile' '-ExecutionPolicy' 'Bypass' '-File' $workerPath `
            '-Action' $Action '-ArgumentJson' $argumentJson '-TenantId' $TenantId 2>&1 |
            ForEach-Object {
                $text = "$_"
                if ($text -like 'RESULT:*') { $resultLine = $text; return }
                if ($text -like 'PROGRESS:*') { Write-Host ('    ' + $text.Substring(9)) -ForegroundColor DarkGray; return }
                $other.Add($text)
            }
        if ([string]::IsNullOrWhiteSpace($resultLine)) {
            return [pscustomobject]@{Status = 'Error'; Message = (@($other) -join ' '); Value = @(); Correlation = @(); Signature = ''}
        }
        return ($resultLine.Substring(7) | ConvertFrom-Json)
    }
    catch {
        return [pscustomobject]@{Status = 'Error'; Message = (Get-ErrorText -ErrorRecord $_); Value = @(); Correlation = @(); Signature = ''}
    }
    finally { Remove-Item -LiteralPath $workerPath -Force -ErrorAction SilentlyContinue }
}

function Save-PendingBillingLink {
    <# .SYNOPSIS Remembers an unfinished billing link so later runs can complete it without being asked. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][AllowEmptyString()][string]$TenantId,
        [Parameter(Mandatory)][string]$SubscriptionId,
        [Parameter(Mandatory)][string]$ResourceGroup,
        [Parameter(Mandatory)][string]$ResourceName,
        [AllowEmptyString()][string]$Signature = '',
        [int]$FailureCount = 0
    )

    # None of these identify anything secret; they are the same values the Azure portal shows.
    Save-RememberedValue -Name 'PURVIEW_FILE_LABELING_PENDING_BILLING_CLIENT_ID' -Value $ClientId
    Save-RememberedValue -Name 'PURVIEW_FILE_LABELING_PENDING_BILLING_TENANT_ID' -Value $TenantId
    Save-RememberedValue -Name 'PURVIEW_FILE_LABELING_PENDING_BILLING_SUBSCRIPTION' -Value $SubscriptionId
    Save-RememberedValue -Name 'PURVIEW_FILE_LABELING_PENDING_BILLING_GROUP' -Value $ResourceGroup
    Save-RememberedValue -Name 'PURVIEW_FILE_LABELING_PENDING_BILLING_NAME' -Value $ResourceName
    # The signature and count are what let a later run tell a passing outage from a permanent refusal.
    if (-not [string]::IsNullOrWhiteSpace($Signature)) {
        Save-RememberedValue -Name 'PURVIEW_FILE_LABELING_PENDING_BILLING_SIGNATURE' -Value $Signature.Substring(0, [math]::Min(120, $Signature.Length))
    }
    if ($FailureCount -gt 0) { Save-RememberedValue -Name 'PURVIEW_FILE_LABELING_PENDING_BILLING_FAILURES' -Value ([string]$FailureCount) }
}

function Get-PendingBillingLink {
    <# .SYNOPSIS Returns an unfinished billing link, or null. #>
    [CmdletBinding()]
    param()

    $clientId = Get-RememberedValue -Name 'PURVIEW_FILE_LABELING_PENDING_BILLING_CLIENT_ID'
    $subscriptionId = Get-RememberedValue -Name 'PURVIEW_FILE_LABELING_PENDING_BILLING_SUBSCRIPTION'
    $resourceGroup = Get-RememberedValue -Name 'PURVIEW_FILE_LABELING_PENDING_BILLING_GROUP'
    $resourceName = Get-RememberedValue -Name 'PURVIEW_FILE_LABELING_PENDING_BILLING_NAME'
    if ([string]::IsNullOrWhiteSpace($clientId) -or [string]::IsNullOrWhiteSpace($subscriptionId) -or
        [string]::IsNullOrWhiteSpace($resourceGroup) -or [string]::IsNullOrWhiteSpace($resourceName)) {
        return $null
    }
    $failureText = Get-RememberedValue -Name 'PURVIEW_FILE_LABELING_PENDING_BILLING_FAILURES'
    $failureCount = 0
    if (-not [int]::TryParse($failureText, [ref]$failureCount)) { $failureCount = 0 }
    return [pscustomobject]@{
        ClientId = $clientId
        TenantId = Get-RememberedValue -Name 'PURVIEW_FILE_LABELING_PENDING_BILLING_TENANT_ID'
        SubscriptionId = $subscriptionId
        ResourceGroup = $resourceGroup
        ResourceName = $resourceName
        Signature = Get-RememberedValue -Name 'PURVIEW_FILE_LABELING_PENDING_BILLING_SIGNATURE'
        FailureCount = $failureCount
    }
}

function Clear-PendingBillingLink {
    <# .SYNOPSIS Forgets an unfinished billing link once it has been created. #>
    [CmdletBinding()]
    param()

    foreach ($name in 'PURVIEW_FILE_LABELING_PENDING_BILLING_CLIENT_ID', 'PURVIEW_FILE_LABELING_PENDING_BILLING_TENANT_ID',
        'PURVIEW_FILE_LABELING_PENDING_BILLING_SUBSCRIPTION', 'PURVIEW_FILE_LABELING_PENDING_BILLING_GROUP', 'PURVIEW_FILE_LABELING_PENDING_BILLING_NAME',
        'PURVIEW_FILE_LABELING_PENDING_BILLING_SIGNATURE', 'PURVIEW_FILE_LABELING_PENDING_BILLING_FAILURES') {
        foreach ($target in [EnvironmentVariableTarget]::Process, [EnvironmentVariableTarget]::User) {
            try { [Environment]::SetEnvironmentVariable($name, $null, $target) }
            catch { Write-Verbose "Could not clear ${name}: $($_.Exception.Message)" }
        }
    }
}

function Repair-AzureModuleVersion {
    <# .SYNOPSIS Removes older Az copies, whose assemblies stop the newest ones from loading. #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param()

    $stale = [System.Collections.Generic.List[object]]::new()
    foreach ($name in 'Az.Accounts', 'Az.Resources') {
        $versions = @(Get-Module -ListAvailable -Name $name -ErrorAction SilentlyContinue | Sort-Object Version -Descending)
        foreach ($old in @($versions | Select-Object -Skip 1)) {
            $stale.Add([pscustomobject]@{Name = $name; Version = [string]$old.Version})
        }
    }
    if ($stale.Count -eq 0) {
        Write-RunLog -Severity INFO -Action 'Check dependency' -Result 'Only one version of each Az module is installed, so a version mismatch is not what failed.'
        return $false
    }

    Write-Host ''
    Write-Host '  More than one version of these modules is installed:' -ForegroundColor Yellow
    foreach ($item in $stale) { Write-Host "    $($item.Name) $($item.Version)" -ForegroundColor DarkGray }
    Write-Host '  Az.Resources is built against one exact Az.Accounts build. When another version' -ForegroundColor Gray
    Write-Host '  loads beside it, their shared assemblies disagree and the type fails to load, which' -ForegroundColor Gray
    Write-Host '  is the OpenTelemetry error above. Removing the older copies resolves it.' -ForegroundColor Gray

    $choice = Read-MenuChoice -Title 'Remove the older copies, keeping the newest of each?' -Options ([ordered]@{
            '1' = 'Yes, uninstall them (default)'
            '2' = 'No, leave them installed'
        }) -Default '1'
    if ($choice -ne '1') { return $false }
    if (-not $PSCmdlet.ShouldProcess('Az.Accounts and Az.Resources', 'Uninstall the older versions')) { return $false }

    $removed = 0
    foreach ($item in $stale) {
        try {
            Uninstall-Module -Name $item.Name -RequiredVersion $item.Version -Force -ErrorAction Stop
            Write-RunLog -Severity SUCCESS -Action 'Check dependency' -Result "Removed $($item.Name) $($item.Version)."
            $removed++
        }
        catch {
            Write-RunLog -Severity WARN -Action 'Check dependency' -Result "Could not remove $($item.Name) $($item.Version): $(Get-ErrorText -ErrorRecord $_). It may have been installed for all users, which needs an elevated session."
        }
    }
    return ($removed -gt 0)
}

function Invoke-BillingLinkInCloudShell {
    <# .SYNOPSIS Completes the billing link in Azure Cloud Shell, which nothing on this machine can block. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][AllowEmptyString()][string]$SubscriptionId,
        [Parameter(Mandatory)][AllowEmptyString()][string]$ResourceGroup,
        [Parameter(Mandatory)][AllowEmptyString()][string]$ResourceName,
        [Parameter(Mandatory)][AllowEmptyString()][string]$TenantId
    )

    Write-Host ''
    Write-Host '  Azure Cloud Shell runs in your browser, signed in as you, with the Azure CLI already' -ForegroundColor Gray
    Write-Host '  installed. This machine takes no part, so nothing here can refuse it. This is also the' -ForegroundColor Gray
    Write-Host '  route Microsoft documents for creating this resource.' -ForegroundColor Gray

    $command = 'az extension add --name graphservices --allow-preview true --upgrade --yes --only-show-errors; ' +
        "az account set --subscription $SubscriptionId; " +
        "az provider register --namespace Microsoft.GraphServices --subscription $SubscriptionId --wait --only-show-errors; " +
        "az graph-services account create --resource-group $ResourceGroup --resource-name $ResourceName --subscription $SubscriptionId --location global --app-id $ClientId; " +
        "az resource list --resource-type Microsoft.GraphServices/accounts --subscription $SubscriptionId --output json --only-show-errors; " +
        "az resource show --resource-group $ResourceGroup --name $ResourceName --resource-type Microsoft.GraphServices/accounts --subscription $SubscriptionId --output json --only-show-errors"

    $copied = $false
    try { Set-Clipboard -Value $command -ErrorAction Stop; $copied = $true }
    catch { Write-Verbose "The clipboard is unavailable: $($_.Exception.Message)" }
    if ($copied) {
        Write-Host ''
        Write-Host '  The whole command is on your clipboard. Paste it into Cloud Shell and press Enter.' -ForegroundColor Green
    }
    else {
        Show-MeteredBillingCommand -ClientId $ClientId -SubscriptionId $SubscriptionId -ResourceGroup $ResourceGroup -ResourceName $ResourceName
    }

    $open = Read-MenuChoice -Title 'Open Azure Cloud Shell now?' -Options ([ordered]@{
            '1' = 'Yes, open it in the browser (default)'
            '2' = 'No, I will open it myself'
        }) -Default '1'
    if ($open -eq '1') {
        try { Start-Process 'https://shell.azure.com' -ErrorAction Stop }
        catch { Write-RunLog -Severity WARN -Action 'Link metered billing' -Result "A browser could not be opened: $(Get-ErrorText -ErrorRecord $_). Go to https://shell.azure.com" }
    }

    while ($true) {
        $choice = Read-MenuChoice -Title 'Once the command has finished, confirm the link.' -Options ([ordered]@{
                '1' = 'It finished, check now (default)'
                '2' = 'Stop and return without the link'
            }) -Default '1'
        if ($choice -ne '1') { return $false }

        $check = Invoke-AzureModuleAction -TenantId $TenantId -Action 'BillingExists' -Arguments @{
            SubscriptionId = $SubscriptionId
            ResourceGroup = $ResourceGroup
            ResourceName = $ResourceName
            ClientId = $ClientId
        }
        if ($check.Status -eq 'Present') {
            $state = Get-MeteredBillingResourceState -Resource (@($check.Value) | Select-Object -First 1) -ClientId $ClientId
            if ($state.Status -ne 'Ready') {
                Write-RunLog -Severity WARN -Action 'Link metered billing' -Result "Azure found the billing resource, but it does not satisfy the documented success response: $($state.Message)"
                return $false
            }
            Write-RunLog -Severity SUCCESS -Action 'Link metered billing' -Result "Billing resource '$ResourceName' satisfies the documented success response, so application $ClientId is billed to subscription $SubscriptionId."
            $script:LastBillingWasServiceFault = $false
            Clear-PendingBillingLink
            return $true
        }
        if ($check.Status -eq 'Missing') {
            Write-RunLog -Severity WARN -Action 'Link metered billing' -Result 'Azure still reports no billing resource by that name. Read the Cloud Shell output for the reason before checking again.'
            continue
        }

        # The same local fault that sent us to Cloud Shell can also block this check.
        Write-RunLog -Severity WARN -Action 'Link metered billing' -Result "The link could not be confirmed from this machine: $([string]$check.Message)"
        $accept = Read-MenuChoice -Title 'Cloud Shell is the authority here. What did it report?' -Options ([ordered]@{
                '1' = 'Nothing conclusive yet, let me check again (default)'
                '2' = "Show returned Global, appId $ClientId, a billingPlanId, and Succeeded"
            }) -Default '1'
        if ($accept -eq '2') {
            Write-RunLog -Severity INFO -Action 'Link metered billing' -Result 'Accepted on the complete documented Cloud Shell response: Global location, expected application ID, billing plan ID, and Succeeded provisioning state.'
            $script:LastBillingWasServiceFault = $false
            Clear-PendingBillingLink
            return $true
        }
    }
}

function Set-MeteredBillingLink {
    <# .SYNOPSIS Creates the Microsoft.GraphServices/accounts resource that bills this application's metered calls. #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$SubscriptionId,
        [Parameter(Mandatory)][string]$ResourceGroup,
        [Parameter(Mandatory)][string]$ResourceName,
        [Parameter(Mandatory)][object]$Tool,
        [AllowEmptyString()][string]$TenantId = '',
        # Used by the automatic retries, which must not re-print the guidance or wait between routes.
        [switch]$Quiet
    )

    if (-not $PSCmdlet.ShouldProcess("subscription $SubscriptionId", "Create billing resource '$ResourceName' for application $ClientId")) { return $false }

    $script:LastBillingPreflightFailed = $false
    $script:LastBillingWasProviderFault = $false
    $script:LastBillingWasServiceFault = $false
    $script:LastBillingWasClientFault = $false
    $script:LastBillingPreviewAttempted = $false
    $script:LastBillingFailureSignature = ''
    $script:LastBillingCorrelationIds = @()
    try {
        if ($Tool.Kind -eq 'AzureCli') {
            Write-RunLog -Severity INFO -Action 'Link metered billing' -Result 'Using the Azure CLI. A browser sign-in may appear if the CLI is not already signed in.'
            Write-AzureCliBillingDiagnostic -Tool $Tool -SubscriptionId $SubscriptionId -ResourceGroup $ResourceGroup `
                -ResourceName $ResourceName -ClientId $ClientId -TenantId $TenantId
            # Every command receives the subscription explicitly, so this utility never changes the
            # default subscription of an Azure CLI session that was already open.
            if (Test-MeteredBillingResource -SubscriptionId $SubscriptionId -ResourceGroup $ResourceGroup -ResourceName $ResourceName -ClientId $ClientId -Tool $Tool) {
                Write-RunLog -Severity SUCCESS -Action 'Link metered billing' -Result "Billing resource '$ResourceName' already exists in $ResourceGroup, so nothing needed creating."
                $script:LastBillingWasServiceFault = $false
                $script:LastBillingWasProviderFault = $false
                Clear-PendingBillingLink
                return $true
            }
            if ($script:LastBillingResourceStatus.Status -ne 'Missing') {
                $script:LastBillingPreflightFailed = $true
                Add-RunFailure -FilePath '' -Action 'Preflight metered billing' -Reason $script:LastBillingResourceStatus.Message
                return $false
            }
            Write-RunLog -Severity INFO -Action 'Link metered billing' -Result 'Registering the Microsoft.GraphServices resource provider on that subscription. This can take a minute the first time.'
            $providerOutput = & az provider register --namespace Microsoft.GraphServices --subscription $SubscriptionId --wait --only-show-errors 2>&1
            if ($LASTEXITCODE -ne 0) { throw "az provider register failed: $(($providerOutput -join ' ').Trim())" }
            $preflight = Test-MeteredBillingPreflight -Tool $Tool -SubscriptionId $SubscriptionId -ResourceGroup $ResourceGroup `
                -ClientId $ClientId -TenantId $TenantId
            if (-not $preflight.Ready) {
                Add-RunFailure -FilePath '' -Action 'Preflight metered billing' -Reason $preflight.Message
                $diagnostic = Get-AzureFailureDiagnostic -Message $preflight.Message
                $codeText = if ($diagnostic.Codes.Count -gt 0) { $diagnostic.Codes -join ' > ' } else { '<none parsed>' }
                $identifierText = if ($diagnostic.Identifiers.Count -gt 0) { $diagnostic.Identifiers -join ', ' } else { '<none parsed>' }
                Write-RunLog -Severity INFO -Action 'Record Azure billing diagnostics' -NoConsole -Result `
                    "Preflight classification=$($preflight.FailureKind); error codes=$codeText; support identifiers=$identifierText."
                if ($preflight.FailureKind -eq 'ProviderFault') {
                    $script:LastBillingWasProviderFault = $true
                    $script:LastBillingWasServiceFault = $true
                    $script:LastBillingFailureSignature = Get-FailureSignature -Message $preflight.Message
                }
                else {
                    $script:LastBillingPreflightFailed = $true
                }
                return $false
            }
            Write-RunLog -Severity SUCCESS -Action 'Preflight metered billing' -Result $preflight.Message
            # The graph-services command group ships in the graphservices extension package.
            $extensionOutput = & az extension add --name graphservices --allow-preview true --upgrade --yes --only-show-errors 2>&1
            if ($LASTEXITCODE -ne 0) { throw "az extension add failed: $(($extensionOutput -join ' ').Trim())" }
            $extensionVersionOutput = & az extension show --name graphservices --query version --output tsv --only-show-errors 2>&1
            $extensionVersion = if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace("$extensionVersionOutput")) {
                "$extensionVersionOutput".Trim()
            }
            else { '<installed; version unavailable>' }
            Write-RunLog -Severity INFO -Action 'Record Azure billing diagnostics' -NoConsole -Result "graphservices extension after installation=$extensionVersion."
            $output = & az graph-services account create --resource-group $ResourceGroup --resource-name $ResourceName `
                --subscription $SubscriptionId --location global --app-id $ClientId 2>&1
            $created = $false
            if ($LASTEXITCODE -eq 0) {
                $created = Test-MeteredBillingResource -SubscriptionId $SubscriptionId -ResourceGroup $ResourceGroup `
                    -ResourceName $ResourceName -ClientId $ClientId -Tool $Tool
                if (-not $created) {
                    throw "az graph-services account create completed, but $($script:LastBillingResourceStatus.Message)"
                }
            }
            else {
                $primaryError = (("$output" -replace '\s+', ' ').Trim())
                if (Test-MeteredBillingResource -SubscriptionId $SubscriptionId -ResourceGroup $ResourceGroup -ResourceName $ResourceName -ClientId $ClientId -Tool $Tool) {
                    $created = $true
                    Write-RunLog -Severity SUCCESS -Action 'Link metered billing' -Result 'Azure returned an error after creating the resource, but the documented list/show verification confirmed the billing link.'
                }
                elseif (-not $Quiet -and (Test-AzureProviderImplementationFailure -Message $primaryError) -and
                    @($preflight.ApiVersions) -contains '2022-09-22-preview') {
                    Write-Host ''
                    Write-Host '  The GA API failed inside Microsoft.GraphServices, but this subscription also' -ForegroundColor Yellow
                    Write-Host '  advertises Microsoft''s published 2022-09-22-preview contract. It has the same' -ForegroundColor Gray
                    Write-Host '  resource ID and application property, but may reach the provider''s earlier path.' -ForegroundColor Gray
                    Write-Host '  Trying it sends one PUT to this same billing resource; it does not create a' -ForegroundColor Gray
                    Write-Host '  second application, certificate, resource group, or billing-resource name.' -ForegroundColor Gray
                    $previewChoice = Read-MenuChoice -Title 'Try that provider-advertised preview API once?' -Options ([ordered]@{
                            '1' = 'Yes, try the same billing resource with 2022-09-22-preview'
                            '2' = 'No, keep the pending link without another create request (default)'
                        }) -Default '2'
                    if ($previewChoice -ne '1') { throw "az graph-services account create failed: $primaryError" }

                    $previewApiVersion = '2022-09-22-preview'
                    $previewProperties = @{appId = $ClientId} | ConvertTo-Json -Compress
                    Write-RunLog -Severity INFO -Action 'Link metered billing' -Result "Trying one provider-advertised $previewApiVersion PUT against the same Microsoft.GraphServices/accounts resource after the GA provider fault."
                    $script:LastBillingPreviewAttempted = $true
                    $previewOutput = & az resource create --resource-group $ResourceGroup --name $ResourceName `
                        --resource-type 'Microsoft.GraphServices/accounts' --api-version $previewApiVersion `
                        --location global --properties $previewProperties --subscription $SubscriptionId `
                        --output json --only-show-errors 2>&1
                    $previewExitCode = $LASTEXITCODE
                    if (Test-MeteredBillingResource -SubscriptionId $SubscriptionId -ResourceGroup $ResourceGroup `
                            -ResourceName $ResourceName -ClientId $ClientId -Tool $Tool) {
                        $created = $true
                        Write-RunLog -Severity SUCCESS -Action 'Link metered billing' -Result "The $previewApiVersion request produced the complete verified billing resource after the GA provider path failed."
                    }
                    else {
                        $previewError = (("$previewOutput" -replace '\s+', ' ').Trim())
                        if ($previewExitCode -eq 0) {
                            $previewError = "The preview request completed, but $($script:LastBillingResourceStatus.Message)"
                        }
                        elseif ([string]::IsNullOrWhiteSpace($previewError)) {
                            $previewError = "Azure CLI exited with code $previewExitCode and returned no error text."
                        }
                        throw "GA create failed: $primaryError Preview $previewApiVersion create also failed: $previewError"
                    }
                }
                else {
                    throw "az graph-services account create failed: $primaryError"
                }
            }
        }
        else {
            if (-not (Get-Module -ListAvailable -Name Az.Resources -ErrorAction SilentlyContinue)) { throw 'Az.Resources is not available.' }
            Write-RunLog -Severity INFO -Action 'Link metered billing' -Result 'Working in a separate Azure process. Registering the provider can take a minute, and progress appears below.'
            # Everything Az does happens in a separate process, because its assemblies otherwise break PnP in this one.
            $response = Invoke-AzureModuleAction -TenantId $TenantId -Action 'BillingLink' -Arguments @{
                SubscriptionId = $SubscriptionId
                ResourceGroup = $ResourceGroup
                ResourceName = $ResourceName
                ClientId = $ClientId
            }
            $script:LastBillingCorrelationIds = @(@($response.Correlation) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            if ($response.Status -eq 'Exists') {
                $state = Get-MeteredBillingResourceState -Resource (@($response.Value) | Select-Object -First 1) -ClientId $ClientId
                if ($state.Status -ne 'Ready') {
                    $script:LastBillingPreflightFailed = $true
                    Add-RunFailure -FilePath '' -Action 'Preflight metered billing' -Reason $state.Message
                    return $false
                }
                Write-RunLog -Severity SUCCESS -Action 'Link metered billing' -Result "Billing resource '$ResourceName' already exists in $ResourceGroup, so nothing needed creating."
                $script:LastBillingWasServiceFault = $false
                $script:LastBillingWasProviderFault = $false
                Clear-PendingBillingLink
                return $true
            }
            if ($response.Status -eq 'PreflightFailed') {
                $script:LastBillingPreflightFailed = $true
                Add-RunFailure -FilePath '' -Action 'Preflight metered billing' -Reason ([string]$response.Message)
                return $false
            }
            if ($response.Status -eq 'ProviderFault') {
                $script:LastBillingWasProviderFault = $true
                $script:LastBillingWasServiceFault = $true
                $script:LastBillingFailureSignature = Get-FailureSignature -Message ([string]$response.Message)
                Add-RunFailure -FilePath '' -Action 'Link metered billing' -Reason ([string]$response.Message)
                return $false
            }
            if ($response.Status -eq 'ClientFault') {
                Write-RunLog -Severity WARN -Action 'Link metered billing' -Result 'The Azure module failed to load in its own process, so no route reached Azure at all. That is a fault on this machine, not a rejection by Azure.'
                $script:LastBillingWasClientFault = $true
                $script:LastBillingWasServiceFault = $false
                $script:LastBillingFailureSignature = Get-FailureSignature -Message ([string]$response.Message)
                # Mismatched Az versions are the usual cause, and removing them is worth one more attempt.
                if (-not $Quiet -and (Repair-AzureModuleVersion -Confirm:$false)) {
                    Write-RunLog -Severity INFO -Action 'Link metered billing' -Result 'The mismatched copies were removed, so the billing link is being tried once more.'
                    $retry = Invoke-AzureModuleAction -TenantId $TenantId -Action 'BillingLink' -Arguments @{
                        SubscriptionId = $SubscriptionId
                        ResourceGroup = $ResourceGroup
                        ResourceName = $ResourceName
                        ClientId = $ClientId
                    }
                    if ($retry.Status -eq 'Created' -or $retry.Status -eq 'Exists') {
                        $retryState = Get-MeteredBillingResourceState -Resource (@($retry.Value) | Select-Object -First 1) -ClientId $ClientId
                        if ($retryState.Status -eq 'Ready') {
                            Write-RunLog -Severity SUCCESS -Action 'Link metered billing' -Result "Application $ClientId is now billed to subscription $SubscriptionId through resource '$ResourceName'."
                            $script:LastBillingWasClientFault = $false
                            Clear-PendingBillingLink
                            return $true
                        }
                        Write-RunLog -Severity WARN -Action 'Link metered billing' -Result "The resource still does not satisfy the documented success response: $($retryState.Message)"
                    }
                    Write-RunLog -Severity WARN -Action 'Link metered billing' -Result "It still failed after the cleanup: $([string]$retry.Message)"
                }
                # The CLI is a separate application, so it cannot inherit the fault that just occurred.
                if (-not $Quiet) {
                    $cli = Install-AzureCommandLine -Confirm:$false
                    if ($null -ne $cli) {
                        $null = & az account show --only-show-errors 2>&1
                        if ($LASTEXITCODE -ne 0 -and -not (Connect-AzureCommandLine -Tool $cli -TenantId $TenantId)) { return $false }
                        Write-RunLog -Severity INFO -Action 'Link metered billing' -Result 'Finishing the link with the Azure CLI instead.'
                        return (Set-MeteredBillingLink -ClientId $ClientId -SubscriptionId $SubscriptionId -ResourceGroup $ResourceGroup `
                                -ResourceName $ResourceName -Tool $cli -TenantId $TenantId -Confirm:$false)
                    }
                }
                return $false
            }
            if ($response.Status -ne 'Created') { throw ([string]$response.Message) }
            $state = Get-MeteredBillingResourceState -Resource (@($response.Value) | Select-Object -First 1) -ClientId $ClientId
            if ($state.Status -ne 'Ready') { throw "Az PowerShell created a resource that does not satisfy the documented success response: $($state.Message)" }
        }
        Write-RunLog -Severity SUCCESS -Action 'Link metered billing' -Result "Application $ClientId is now billed to subscription $SubscriptionId through resource '$ResourceName'."
        $script:LastBillingWasServiceFault = $false
        $script:LastBillingWasProviderFault = $false
        Clear-PendingBillingLink
        return $true
    }
    catch {
        $text = Get-GraphErrorText -ErrorRecord $_
        # A server-side failure can still have created the resource, so confirm before reporting a failure.
        if (Test-MeteredBillingResource -SubscriptionId $SubscriptionId -ResourceGroup $ResourceGroup -ResourceName $ResourceName -ClientId $ClientId -Tool $Tool) {
            Write-RunLog -Severity SUCCESS -Action 'Link metered billing' -Result "Azure reported an error, but billing resource '$ResourceName' exists, so application $ClientId is linked."
            $script:LastBillingWasServiceFault = $false
            $script:LastBillingWasProviderFault = $false
            Clear-PendingBillingLink
            return $true
        }
        $script:LastBillingWasServiceFault = [bool](Test-AzureServiceFailure -Message $text)
        $script:LastBillingWasProviderFault = [bool](Test-AzureProviderImplementationFailure -Message $text)
        $script:LastBillingWasClientFault = $false
        $script:LastBillingFailureSignature = Get-FailureSignature -Message $text
        $diagnostic = Get-AzureFailureDiagnostic -Message $text
        $directIds = @($diagnostic.Identifiers | ForEach-Object {
                $match = [regex]::Match($_, '=(?<value>[0-9a-f-]{36})$')
                if ($match.Success) { $match.Groups['value'].Value }
            })
        $activityEvidence = $null
        if ($Tool.Kind -eq 'AzureCli' -and $script:LastBillingWasServiceFault) {
            $activityEvidence = Get-AzureBillingActivityEvidence -SubscriptionId $SubscriptionId `
                -ResourceGroup $ResourceGroup -ResourceName $ResourceName
        }
        $activityIds = if ($null -ne $activityEvidence) { @($activityEvidence.CorrelationIds) } else { @() }
        $allCorrelationIds = @($directIds) + @($activityIds)
        $script:LastBillingCorrelationIds = @($allCorrelationIds |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
        $codeText = if ($diagnostic.Codes.Count -gt 0) { $diagnostic.Codes -join ' > ' } else { '<none parsed>' }
        $identifierText = if ($script:LastBillingCorrelationIds.Count -gt 0) { $script:LastBillingCorrelationIds -join ', ' } else { '<none available yet>' }
        Write-RunLog -Severity INFO -Action 'Record Azure billing diagnostics' -NoConsole -Result `
            "Create classification=$(if ($script:LastBillingWasProviderFault) { 'ProviderFault' } elseif ($script:LastBillingWasServiceFault) { 'ServiceFault' } else { 'CallerFault' }); preview recovery attempted=$($script:LastBillingPreviewAttempted); error codes=$codeText; ARM correlation IDs=$identifierText."
        if ($Quiet) { return $false }
        Add-RunFailure -FilePath '' -Action 'Link metered billing' -Reason $text
        if ($script:LastBillingWasProviderFault) {
            Write-Host ''
            Write-Host '  Microsoft.GraphServices returned an OpenTelemetry type-load failure.' -ForegroundColor Yellow
            if ($script:LastBillingPreviewAttempted) {
                Write-Host '  Both the GA request and the explicitly selected preview request reached that' -ForegroundColor Gray
                Write-Host '  provider without producing a complete billing resource. No further mutation' -ForegroundColor Gray
                Write-Host '  was sent.' -ForegroundColor Gray
            }
            else {
                Write-Host '  No preview recovery request was sent because it was not advertised, was' -ForegroundColor Gray
                Write-Host '  declined, or this was an unattended retry.' -ForegroundColor Gray
            }
            Write-Host '  The pending link is retained, but' -ForegroundColor Gray
            Write-Host '  startup will not retry this exact failure automatically.' -ForegroundColor Gray
        }
        elseif ($script:LastBillingWasServiceFault) {
            Write-Host ''
            Write-Host '  Azure returned a server-side error while creating the billing resource.' -ForegroundColor Yellow
            Write-Host '  Preflight completed, but it cannot prove Microsoft.GraphServices create backend' -ForegroundColor Gray
            Write-Host '  health. The full error is in the run log, with correlation IDs below if present.' -ForegroundColor Gray
        }
        if ($script:LastBillingWasServiceFault -and $script:LastBillingCorrelationIds.Count -gt 0) {
            Write-Host ''
            Write-Host '  ARM correlation IDs for the failed provider operation:' -ForegroundColor Gray
            foreach ($id in $script:LastBillingCorrelationIds) { Write-Host "    $id" -ForegroundColor DarkGray }
        }
        elseif ($script:LastBillingWasServiceFault -and $null -ne $activityEvidence) {
            Write-Host ''
            Write-Host '  Azure Activity Log has not exposed a correlation ID yet. It can take a few' -ForegroundColor Gray
            Write-Host '  minutes to ingest the failed event. This command only reads that evidence:' -ForegroundColor Gray
            Write-Host "    $($activityEvidence.LookupCommand)" -ForegroundColor DarkGray
        }
        return $false
    }
}

function Test-MeteredPrerequisite {
    <# .SYNOPSIS Verifies one supported route for consenting the metered application. #>
    [CmdletBinding()]
    param()

    if (Get-Command az -CommandType Application -ErrorAction SilentlyContinue) {
        Write-RunLog -Severity INFO -Action 'Check prerequisite' -Result 'Azure CLI is available, so one tenant-pinned CLI sign-in can be reused for administrator consent and metered billing.'
        return $true
    }

    $module = 'Microsoft.Graph.Authentication'
    $purpose = 'as the administrator-consent fallback when Azure CLI is unavailable'
    $installed = Get-Module -ListAvailable -Name $module -ErrorAction SilentlyContinue
    if (-not $installed -and (Request-ModuleInstall -Name $module -Purpose $purpose)) {
        $installed = Get-Module -ListAvailable -Name $module -ErrorAction SilentlyContinue
    }
    if (-not $installed) {
        Write-RunLog -Severity WARN -Action 'Check prerequisite' -Result "$module is required $purpose, but installation was declined or failed."
        return $false
    }
    if (-not (Get-Module -ListAvailable -Name Az.Resources -ErrorAction SilentlyContinue)) {
        Write-RunLog -Severity INFO -Action 'Check prerequisite' -Result 'Azure CLI or Az.Resources will be offered when the billing association is created; neither is required before application registration.'
    }
    return $true
}

function Invoke-MeteredSetup {
    <# .SYNOPSIS Guides the one-time setup that lets this utility write SharePoint labels. #>
    [CmdletBinding()]
    param()

    if (-not (Test-MeteredPrerequisite)) {
        Write-RunLog -Severity ERROR -Action 'Metered setup' -Result 'Required PowerShell modules are not installed. Complete installation and try again.'
        return $false
    }

    Write-Host ''
    Write-Host '  Enable SharePoint Online metered label writing' -ForegroundColor Cyan
    Write-Host '  Writing a label to SharePoint calls the Graph assignSensitivityLabel API.' -ForegroundColor Gray
    Write-Host '  That API is metered, and Microsoft accepts it only from a confidential client,' -ForegroundColor Gray
    Write-Host '  which means an application that authenticates with a certificate rather than' -ForegroundColor Gray
    Write-Host '  an interactive user. This sets one up in three steps:' -ForegroundColor Gray
    Write-Host ''
    Write-Host '    1. Register an application with a certificate in your personal store.' -ForegroundColor Gray
    Write-Host '    2. Grant it administrator consent, which creates its service principal.' -ForegroundColor Gray
    Write-Host '    3. Link it to an Azure subscription, which is then billed per API call.' -ForegroundColor Gray
    Write-Host ''
    Write-Host '  You need: an administrator to consent, an Azure subscription in the same tenant,' -ForegroundColor Yellow
    Write-Host '  and Contributor rights on it. Metered APIs are unavailable in national clouds,' -ForegroundColor Yellow
    Write-Host '  including GCC.' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '  Microsoft classes assignSensitivityLabel as a protected API, which means it' -ForegroundColor Yellow
    Write-Host '  needs validation beyond permissions and consent. That validation is the Azure' -ForegroundColor Yellow
    Write-Host '  billing link in step 3, so finishing this setup is all that is asked of you.' -ForegroundColor Yellow
    Write-Host '  A token issued before that link exists is still refused, so start the utility' -ForegroundColor Yellow
    Write-Host '  again once setup finishes.' -ForegroundColor Yellow

    $existing = Get-ConfidentialClientConfig
    $replaceOldClientId = ''
    $replaceOldTenantId = ''
    if ($existing) {
        $certificateState = if (Test-SigningCertificateAvailable -Thumbprint $existing.Thumbprint) { 'present' } else { 'MISSING from your certificate store' }
        Write-Host ''
        Write-Host "  A confidential client is already saved: $($existing.ClientId)" -ForegroundColor Green
        Write-Host "  Certificate $($existing.Thumbprint) is $certificateState." -ForegroundColor Gray
        $choice = Read-MenuChoice -Title 'What should happen to it?' -Options ([ordered]@{
                '1' = 'Keep it and only (re)link Azure billing'
                '2' = 'Grant administrator consent and ensure the billing operator is an app owner'
                '3' = 'Replace it with a newly registered application'
                '4' = 'Forget it, so this utility returns to survey-only mode'
                '5' = 'Return to the main menu'
            }) -Default '1'
        if ($choice -eq '5') { return 'Main' }
        if ($choice -eq '2') {
            $null = Grant-LabelingAdminConsent -ClientId $existing.ClientId -TenantId $existing.TenantId -PreferAzureCli -Confirm:$false
            return 'Main'
        }
        if ($choice -eq '4') {
            $removeChoice = Read-MenuChoice -Title "Also remove application $($existing.ClientId) from Entra ID?" -Options ([ordered]@{
                    '1' = 'Yes, strip its permissions, disable sign-in, and delete it if allowed'
                    '2' = 'No, only forget it here (it keeps its permissions in the tenant)'
                }) -Default '1'
            if ($removeChoice -eq '1') {
                $null = Disable-LabelingApplication -ClientId $existing.ClientId -TenantId $existing.TenantId -Confirm:$false
            }
            else {
                Write-RunLog -Severity WARN -Action 'Forget confidential client' -Result "Application $($existing.ClientId) stays in tenant $($existing.TenantId) holding application-level Files.ReadWrite.All. Remove it in the Entra admin center when it is no longer needed."
            }
            Clear-ConfidentialClientConfig
            return 'Main'
        }
        if ($choice -eq '1') {
            $null = Read-MeteredBillingSetting -ClientId $existing.ClientId -TenantId $existing.TenantId -Thumbprint $existing.Thumbprint
            return 'Main'
        }
        $replaceChoice = Read-MenuChoice -Title "Remove the old application $($existing.ClientId) from Entra ID once the new one works?" -Options ([ordered]@{
                '1' = 'Yes, remove it after the replacement is registered'
                '2' = 'No, leave it in the tenant'
            }) -Default '1'
        if ($replaceChoice -eq '1') {
            # Deleting first would leave nothing behind if registration then failed, which is how a saved client goes stale.
            $replaceOldClientId = $existing.ClientId
            $replaceOldTenantId = $existing.TenantId
        }
        else {
            Write-RunLog -Severity WARN -Action 'Replace confidential client' -Result "Application $($existing.ClientId) stays in the tenant and keeps its permissions. Delete it in the Entra admin center once the replacement works."
        }
    }

    if (-not (Test-SharePointPrerequisite)) { return 'Main' }

    Write-Host ''
    Write-Host '  The tenant is read from SharePoint itself, so no GUID has to be typed.' -ForegroundColor Gray
    $siteUrl = Read-ValueWithDefault -Prompt 'Any site URL in the tenant (for example https://contoso.sharepoint.com/sites/LabelTest)' -Default (Get-SessionValue -Name 'PURVIEW_FILE_LABELING_SITE_URL')
    if ($siteUrl -notmatch '^https://[^/]+\.') {
        Write-RunLog -Severity WARN -Action 'Validate site URL' -Result 'Enter the full site URL, starting with https://.'
        return 'Main'
    }
    $siteUrl = $siteUrl.TrimEnd('/')
    Save-SessionValue -Name 'PURVIEW_FILE_LABELING_SITE_URL' -Value $siteUrl
    $tenantId = Get-SharePointTenantId -SiteUrl $siteUrl
    if ([string]::IsNullOrWhiteSpace($tenantId)) {
        $reason = if ($script:LastSharePointTenantLookupStatus -eq 'SiteNotFound') {
            'SharePoint returned 404 for this site. It may have been deleted or the URL may be incorrect, so no application was registered.'
        }
        else { 'SharePoint did not return a tenant realm, so no application was registered.' }
        Add-RunFailure -FilePath $siteUrl -Action 'Resolve tenant' -Reason $reason
        return 'Main'
    }
    Write-RunLog -Severity SUCCESS -Action 'Resolve tenant' -Result "SharePoint reports tenant $tenantId."

    $applicationName = Read-ValueWithDefault -Prompt 'Application display name' -Default 'Purview File Labeling - Confidential Client'
    $confirm = Read-MenuChoice -Title "Register '$applicationName' in tenant $tenantId with application-level Graph Files.ReadWrite.All?" -Options ([ordered]@{
            '1' = 'No, return to the main menu (default)'
            '2' = 'Yes, register it now'
        }) -Default '1'
    if ($confirm -ne '2') { return 'Main' }

    $application = $null
    try { $application = New-ConfidentialLabelingApplication -TenantId $tenantId -ApplicationName $applicationName -Confirm:$false }
    catch {
        Add-RunFailure -FilePath '' -Action 'Register confidential client' -Reason (Get-ErrorText -ErrorRecord $_)
        return 'Main'
    }
    if ($null -eq $application) { return 'Main' }

    Save-ConfidentialClientConfig -ClientId $application.ClientId -TenantId $application.TenantId -Thumbprint $application.Thumbprint
    if (-not [string]::IsNullOrWhiteSpace($replaceOldClientId)) {
        Write-RunLog -Severity INFO -Action 'Replace confidential client' -Result "The replacement is registered and saved, so the previous application $replaceOldClientId is being removed now."
        $null = Disable-LabelingApplication -ClientId $replaceOldClientId -TenantId $replaceOldTenantId -Confirm:$false
    }
    Write-Host ''
    Write-Host "  Application : $($application.ClientId)" -ForegroundColor Green
    Write-Host "  Certificate : $($application.Thumbprint)  (Cert:\CurrentUser\My)" -ForegroundColor Green
    Write-Host ''
    Write-Host '  Registering created the application, but app-only sign-in also needs a service' -ForegroundColor Yellow
    Write-Host '  principal and administrator consent. The next step also makes its signed-in' -ForegroundColor Yellow
    Write-Host '  administrator an explicit application owner, which Azure billing requires.' -ForegroundColor Yellow
    Write-Host '  Until consent is complete, sign-in fails with AADSTS700016 and this utility stays read-only.' -ForegroundColor Yellow
    $consentChoice = Read-MenuChoice -Title 'Grant administrator consent now?' -Options ([ordered]@{
            '1' = 'Yes, sign in as an administrator and grant it from here'
            '2' = 'No, I will grant it in the Entra admin center'
        }) -Default '1'
    if ($consentChoice -eq '1') {
            if (-not (Grant-LabelingAdminConsent -ClientId $application.ClientId -TenantId $application.TenantId -PreferAzureCli -Confirm:$false)) {
                Write-RunLog -Severity WARN -Action 'Metered setup' -Result 'Administrator consent was not completed, so Azure billing was not attempted. The application and certificate are retained for option 3 to finish later.'
                return 'Main'
            }
    }
    else {
        Show-AdminConsentInstruction -ClientId $application.ClientId
            Write-RunLog -Severity INFO -Action 'Metered setup' -Result 'Azure billing was not attempted because administrator consent was deferred. Finish the portal step, then choose option 3 and keep the existing confidential client.'
            return 'Main'
    }

    if (Read-MeteredBillingSetting -ClientId $application.ClientId -TenantId $application.TenantId -Thumbprint $application.Thumbprint) { return 'Main' }

    if ($script:LastBillingWasProviderFault) {
        Write-Host ''
        Write-Host '  The application and certificate are being kept because Microsoft.GraphServices,' -ForegroundColor Yellow
        Write-Host '  not this registration, failed the documented billing command. Registering another' -ForegroundColor Yellow
        Write-Host '  application would repeat the same provider failure. The exact pending link is saved.' -ForegroundColor Yellow
        Write-RunLog -Severity WARN -Action 'Keep confidential client' -Result "Application $($application.ClientId), its certificate, and the pending billing link were retained after the Microsoft.GraphServices provider fault. No rollback or re-registration is needed."
        return 'Main'
    }

    # A confidential client without a billing link cannot write anything, so leaving it behind only accumulates clutter.
    Write-Host ''
    Write-Host '  Billing was not linked, so this application cannot apply labels. Keeping it' -ForegroundColor Yellow
    Write-Host '  only makes sense if you expect the link to succeed on a later run.' -ForegroundColor Yellow
    $rollbackChoice = Read-MenuChoice -Title 'Undo this setup?' -Options ([ordered]@{
            '1' = 'Yes, delete the application and its certificate, and forget the pending link'
            '2' = 'No, keep it so a later run can finish the billing link'
        }) -Default '1'
    if ($rollbackChoice -eq '1') {
        $null = Disable-LabelingApplication -ClientId $application.ClientId -TenantId $application.TenantId -Confirm:$false
        $null = Remove-LabelingCertificate -Thumbprint $application.Thumbprint -Confirm:$false
        Clear-ConfidentialClientConfig
        Clear-PendingBillingLink
        Write-RunLog -Severity SUCCESS -Action 'Undo setup' -Result 'The application, its certificate, and the pending billing link were all removed. Nothing from this attempt is left behind.'
    }
    return 'Main'
}

function Connect-AzureCommandLine {
    <# .SYNOPSIS Signs the Azure tooling in, so subscriptions can be listed instead of typed. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Tool,
        [Parameter(Mandatory)][AllowEmptyString()][string]$TenantId
    )

    try {
        if ($Tool.Kind -eq 'AzureCli') {
            $existing = Get-AzureCliAccountContext
            if ($null -ne $existing -and
                ([string]::IsNullOrWhiteSpace($TenantId) -or
                    [string]::Equals($existing.TenantId, $TenantId, [System.StringComparison]::OrdinalIgnoreCase)) -and
                [string]::Equals($existing.UserType, 'user', [System.StringComparison]::OrdinalIgnoreCase)) {
                Write-RunLog -Severity SUCCESS -Action 'Sign in to Azure' -Result "Reusing the Azure CLI session for $($existing.UserName) in tenant $($existing.TenantId)."
                return $true
            }
            if (-not (Start-IsolatedAzureCliProfile)) { return $false }
            $configOutput = & az config set core.login_experience_v2=off --only-show-errors 2>&1
            if ($LASTEXITCODE -ne 0) {
                Write-RunLog -Severity WARN -Action 'Prepare Azure CLI' -Result "The optional Azure CLI subscription-selector setting could not be disabled: $(($configOutput -join ' ').Trim())"
            }
            Write-RunLog -Severity INFO -Action 'Sign in to Azure' -Result 'Opening an Azure CLI sign-in. Use one account that is an owner of the application registration and has Contributor or Owner access to the subscription being billed.'
            $arguments = [System.Collections.Generic.List[string]]::new()
            $arguments.Add('login')
            if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $arguments.Add('--tenant'); $arguments.Add($TenantId) }
            $arguments.Add('--allow-no-subscriptions')
            if ($script:UseDeviceCode) { $arguments.Add('--use-device-code') }
            $arguments.Add('--output'); $arguments.Add('none')
            & az @arguments
            $signedIn = ($LASTEXITCODE -eq 0)
            if ($signedIn) {
                $account = Get-AzureCliAccountContext
                if ($null -eq $account -or
                    (-not [string]::IsNullOrWhiteSpace($TenantId) -and
                        -not [string]::Equals($account.TenantId, $TenantId, [System.StringComparison]::OrdinalIgnoreCase)) -or
                    -not [string]::Equals($account.UserType, 'user', [System.StringComparison]::OrdinalIgnoreCase)) {
                    Write-RunLog -Severity WARN -Action 'Sign in to Azure' -Result "Azure CLI did not expose an active account in application tenant $TenantId after sign-in."
                    Clear-IsolatedAzureCliProfile
                    return $false
                }
                $script:AzureCliSessionOpened = $true
                $script:AzureCliAccount = $account.UserName
            }
            else { Clear-IsolatedAzureCliProfile }
            return $signedIn
        }
        Write-RunLog -Severity INFO -Action 'Sign in to Azure' -Result 'Opening an Azure sign-in in a separate process. It may appear behind this window.'
        $response = Invoke-AzureAction -Action 'SignIn' -TenantId $TenantId
        if ($response.Status -eq 'SignedIn') { $script:AzurePowerShellSessionOpened = $true }
        return ($response.Status -in 'SignedIn', 'Reused')
    }
    catch {
        if (-not [string]::IsNullOrWhiteSpace($script:AzureCliIsolatedConfigDirectory)) {
            Clear-IsolatedAzureCliProfile
        }
        Add-RunFailure -FilePath '' -Action 'Sign in to Azure' -Reason (Get-ErrorText -ErrorRecord $_)
        return $false
    }
}

function Invoke-AzureModuleAction {
    <# .SYNOPSIS Convenience wrapper that keeps every Az module call inside the separate process. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Action,
        [hashtable]$Arguments = @{},
        [AllowEmptyString()][string]$TenantId = ''
    )

    $response = Invoke-AzureAction -Action $Action -Arguments $Arguments -TenantId $TenantId
    if ($response.Status -eq 'NoSession') {
        Write-RunLog -Severity INFO -Action 'Sign in to Azure' -Result 'Azure is not signed in yet. A sign-in window opens now; it may be behind this one.'
        $signIn = Invoke-AzureAction -Action 'SignIn' -TenantId $TenantId
        if ($signIn.Status -notin 'SignedIn', 'Reused') {
            Add-RunFailure -FilePath '' -Action 'Sign in to Azure' -Reason ([string]$signIn.Message)
            return $response
        }
        if ($signIn.Status -eq 'SignedIn') { $script:AzurePowerShellSessionOpened = $true }
        $response = Invoke-AzureAction -Action $Action -Arguments $Arguments -TenantId $TenantId
    }
    return $response
}

function Disconnect-AzureSession {
    <# .SYNOPSIS Signs out only Azure sessions this utility opened during the current run. #>
    [CmdletBinding()]
    param()

    if (-not [string]::IsNullOrWhiteSpace($script:AzureCliIsolatedConfigDirectory)) { Clear-IsolatedAzureCliProfile }
    if ($script:AzurePowerShellSessionOpened) {
        try {
            $response = Invoke-AzureAction -Action 'SignOut'
            if ($response.Status -ne 'SignedOut') { throw ([string]$response.Message) }
            Write-RunLog -Severity INFO -Action 'Disconnect Azure PowerShell' -Result 'The Azure PowerShell context opened by this utility was signed out.'
        }
        catch {
            Write-RunLog -Severity WARN -Action 'Disconnect Azure PowerShell' -Result (Get-ErrorText -ErrorRecord $_)
        }
        finally { $script:AzurePowerShellSessionOpened = $false }
    }
}

function ConvertFrom-AzureCliJson {
    <# .SYNOPSIS Parses Azure CLI output, which arrives as separate lines that Windows PowerShell will not pipe into ConvertFrom-Json. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowNull()][object]$Output)

    $text = (@($Output) -join [Environment]::NewLine).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return @() }
    try { return @($text | ConvertFrom-Json -ErrorAction Stop) }
    catch {
        Write-Verbose "Azure CLI output was not JSON: $($_.Exception.Message)"
        return @()
    }
}

function Get-AzureSubscriptionList {
    <# .SYNOPSIS Lists the Azure subscriptions the signed-in account can reach. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Tool)

    if ($Tool.Kind -eq 'AzureCli') {
        try {
            $raw = & az account list --all --output json 2>$null
            if ($LASTEXITCODE -ne 0) { return @() }
            return @(ConvertFrom-AzureCliJson -Output $raw | ForEach-Object {
                    [pscustomobject]@{
                        Name = Get-ObjectPropertyValue -InputObject $_ -Names 'name'
                        Id = Get-ObjectPropertyValue -InputObject $_ -Names 'id'
                        TenantId = Get-ObjectPropertyValue -InputObject $_ -Names 'tenantId', 'homeTenantId'
                    }
                } | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Id) })
        }
        catch {
            Write-Verbose "Could not list Azure subscriptions: $($_.Exception.Message)"
            return @()
        }
    }
    $response = Invoke-AzureAction -Action 'ListSubscriptions'
    if ($response.Status -ne 'Listed') { return @() }
    return @(@($response.Value) | ForEach-Object {
            [pscustomobject]@{Name = [string]$_.Name; Id = [string]$_.Id; TenantId = [string]$_.TenantId}
        } | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Id) })
}

function Get-AzureResourceGroupDetail {
    <# .SYNOPSIS Returns the resource groups of one subscription with their regions. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Tool,
        [Parameter(Mandatory)][string]$SubscriptionId
    )

    if ($Tool.Kind -eq 'AzureCli') {
        try {
            $raw = & az group list --subscription $SubscriptionId --output json 2>$null
            if ($LASTEXITCODE -ne 0) { return @() }
            return @(ConvertFrom-AzureCliJson -Output $raw | ForEach-Object {
                    $propertiesProperty = $_.PSObject.Properties['properties']
                    $properties = if ($null -eq $propertiesProperty) { $null } else { $propertiesProperty.Value }
                    [pscustomobject]@{
                        Name = Get-ObjectPropertyValue -InputObject $_ -Names 'name'
                        Location = Get-ObjectPropertyValue -InputObject $_ -Names 'location'
                        State = if ($null -eq $properties) { '' } else { Get-ObjectPropertyValue -InputObject $properties -Names 'provisioningState' }
                    }
                } | Where-Object {
                    -not [string]::IsNullOrWhiteSpace($_.Name) -and
                    [string]::Equals([string]$_.State, 'Succeeded', [System.StringComparison]::OrdinalIgnoreCase)
                })
        }
        catch {
            Write-Verbose "Could not list resource groups: $($_.Exception.Message)"
            return @()
        }
    }
    $response = Invoke-AzureModuleAction -Action 'ListResourceGroups' -Arguments @{SubscriptionId = $SubscriptionId}
    if ($response.Status -ne 'Listed') { return @() }
    return @(@($response.Value) | ForEach-Object {
            [pscustomobject]@{Name = [string]$_.Name; Location = [string]$_.Location; State = [string]$_.State}
        } | Where-Object {
            -not [string]::IsNullOrWhiteSpace($_.Name) -and
            [string]::Equals([string]$_.State, 'Succeeded', [System.StringComparison]::OrdinalIgnoreCase)
        })
}

function Get-AzureResourceGroupList {
    <# .SYNOPSIS Lists the resource group names of one subscription. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Tool,
        [Parameter(Mandatory)][string]$SubscriptionId
    )

    return @(Get-AzureResourceGroupDetail -Tool $Tool -SubscriptionId $SubscriptionId | ForEach-Object { $_.Name })
}

function Get-AzureDefaultLocation {
    <# .SYNOPSIS Suggests a region by reusing one the subscription already uses. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Tool,
        [Parameter(Mandatory)][string]$SubscriptionId
    )

    $first = @(Get-AzureResourceGroupDetail -Tool $Tool -SubscriptionId $SubscriptionId) | Select-Object -First 1
    if ($first -and -not [string]::IsNullOrWhiteSpace($first.Location)) { return [string]$first.Location }
    return 'eastus'
}

function Select-AzureSubscription {
    <# .SYNOPSIS Presents the discovered subscriptions as a menu, falling back to a typed ID. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Tool,
        [Parameter(Mandatory)][AllowEmptyString()][string]$TenantId
    )

    if ($Tool.Kind -eq 'AzureCli' -and -not (Connect-AzureCommandLine -Tool $Tool -TenantId $TenantId)) { return '' }
    $subscriptions = @(Get-AzureSubscriptionList -Tool $Tool)
    if ($subscriptions.Count -eq 0 -and $Tool.Kind -ne 'AzureCli') {
        Write-RunLog -Severity INFO -Action 'List subscriptions' -Result 'No Azure subscription was visible, which usually means the Azure tooling is not signed in yet.'
        if ((Read-MenuChoice -Title 'Sign in to Azure now so the subscriptions can be listed?' -Options ([ordered]@{'1' = 'Yes, sign in'; '2' = 'No, I will type the subscription ID'}) -Default '1') -eq '1') {
            if (Connect-AzureCommandLine -Tool $Tool -TenantId $TenantId) {
                $subscriptions = @(Get-AzureSubscriptionList -Tool $Tool)
            }
        }
    }
    if ($subscriptions.Count -eq 0) { return '' }

    $wrongTenant = @($subscriptions | Where-Object {
            -not [string]::IsNullOrWhiteSpace($TenantId) -and -not [string]::IsNullOrWhiteSpace($_.TenantId) -and
            -not [string]::Equals($_.TenantId, $TenantId, [System.StringComparison]::OrdinalIgnoreCase)
        })
    if ($wrongTenant.Count -gt 0) {
        Write-RunLog -Severity WARN -Action 'List subscriptions' -Result "Excluded $($wrongTenant.Count) subscription(s) from a different tenant. Metered Graph billing requires the subscription and application to be in the same tenant."
        $subscriptions = @($subscriptions | Where-Object { $wrongTenant.Id -notcontains $_.Id })
    }
    if ($subscriptions.Count -eq 0) {
        Write-RunLog -Severity WARN -Action 'List subscriptions' -Result "No subscription in tenant $TenantId is available for metered Graph billing."
        return ''
    }

    $options = [ordered]@{}
    for ($index = 0; $index -lt $subscriptions.Count; $index++) {
        $subscription = $subscriptions[$index]
        $options[[string]($index + 1)] = '{0}  ({1})' -f $subscription.Name, $subscription.Id
    }
    $options[[string]($subscriptions.Count + 1)] = 'Type a subscription ID instead'
    $choice = Read-MenuChoice -Title 'Choose the Azure subscription that should be billed.' -Options $options -Default '1'
    if ([int]$choice -gt $subscriptions.Count) { return '' }
    return [string]$subscriptions[[int]$choice - 1].Id
}

function New-AzureResourceGroup {
    <# .SYNOPSIS Creates a resource group, so a missing one never sends the operator to the portal. #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory)][object]$Tool,
        [Parameter(Mandatory)][string]$SubscriptionId,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Location
    )

    if (-not $PSCmdlet.ShouldProcess("subscription $SubscriptionId", "Create resource group '$Name' in $Location")) { return $false }
    if ($Tool.Kind -eq 'AzureCli') {
        $null = & az group create --name $Name --location $Location --subscription $SubscriptionId --only-show-errors 2>&1
        if ($LASTEXITCODE -ne 0) {
            Add-RunFailure -FilePath '' -Action 'Create resource group' -Reason "az group create failed for '$Name'."
            return $false
        }
    }
    else {
        $response = Invoke-AzureModuleAction -Action 'CreateResourceGroup' -Arguments @{SubscriptionId = $SubscriptionId; Name = $Name; Location = $Location}
        if ($response.Status -ne 'Created') {
            Add-RunFailure -FilePath '' -Action 'Create resource group' -Reason ([string]$response.Message)
            return $false
        }
    }
    Write-RunLog -Severity SUCCESS -Action 'Create resource group' -Result "Created resource group '$Name' in $Location."
    return $true
}

function Read-NewAzureResourceGroup {
    <# .SYNOPSIS Collects a name and region, then creates the resource group. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Tool,
        [Parameter(Mandatory)][string]$SubscriptionId,
        [Parameter(Mandatory)][AllowEmptyString()][string]$DefaultLocation
    )

    $name = Read-ValueWithDefault -Prompt 'Name for the new resource group' -Default 'purview-file-labeling-rg'
    $location = Read-ValueWithDefault -Prompt 'Azure region for it' -Default $(if ([string]::IsNullOrWhiteSpace($DefaultLocation)) { 'eastus' } else { $DefaultLocation })
    if (-not (New-AzureResourceGroup -Tool $Tool -SubscriptionId $SubscriptionId -Name $name -Location $location -Confirm:$false)) { return '' }
    return $name
}

function Test-AzureResourceGroupReady {
    <# .SYNOPSIS Revalidates a selected resource group because ARM list results can lag deletion. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Tool,
        [Parameter(Mandatory)][string]$SubscriptionId,
        [Parameter(Mandatory)][string]$Name
    )

    if ($Tool.Kind -eq 'AzureCli') {
        try {
            $raw = & az group show --name $Name --subscription $SubscriptionId --output json --only-show-errors 2>&1
            if ($LASTEXITCODE -ne 0) {
                $reason = (("$raw" -replace '\s+', ' ').Trim())
                if ([string]::IsNullOrWhiteSpace($reason)) { $reason = 'Azure returned no details.' }
                Write-RunLog -Severity WARN -Action 'Validate resource group' -Result "Azure no longer confirms resource group '$Name' as existing in subscription ${SubscriptionId}: $reason"
                return $false
            }
            $group = @(ConvertFrom-AzureCliJson -Output $raw) | Select-Object -First 1
            $propertiesProperty = if ($null -eq $group) { $null } else { $group.PSObject.Properties['properties'] }
            $properties = if ($null -eq $propertiesProperty) { $null } else { $propertiesProperty.Value }
            $state = if ($null -eq $properties) { '' } else { [string](Get-ObjectPropertyValue -InputObject $properties -Names 'provisioningState') }
        }
        catch {
            Write-RunLog -Severity WARN -Action 'Validate resource group' -Result "Azure could not revalidate resource group '$Name': $(Get-ErrorText -ErrorRecord $_)"
            return $false
        }
    }
    else {
        $group = @(Get-AzureResourceGroupDetail -Tool $Tool -SubscriptionId $SubscriptionId |
                Where-Object { [string]::Equals([string]$_.Name, $Name, [System.StringComparison]::OrdinalIgnoreCase) } |
                Select-Object -First 1)
        if ($group.Count -eq 0) {
            Write-RunLog -Severity WARN -Action 'Validate resource group' -Result "Azure no longer reports resource group '$Name' in subscription $SubscriptionId."
            return $false
        }
        $state = [string]$group[0].State
    }

    if ([string]::Equals($state, 'Succeeded', [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    $reportedState = if ([string]::IsNullOrWhiteSpace($state)) { '<not reported>' } else { $state }
    Write-RunLog -Severity WARN -Action 'Validate resource group' -Result "Resource group '$Name' is no longer ready; Azure reports provisioning state '$reportedState'."
    return $false
}

function Select-AzureResourceGroup {
    <# .SYNOPSIS Presents the discovered resource groups as a menu, and can create a new one. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Tool,
        [Parameter(Mandatory)][string]$SubscriptionId
    )

    $groups = @(Get-AzureResourceGroupList -Tool $Tool -SubscriptionId $SubscriptionId)
    $defaultLocation = Get-AzureDefaultLocation -Tool $Tool -SubscriptionId $SubscriptionId
    while ($true) {
        if ($groups.Count -eq 0) {
            Write-RunLog -Severity WARN -Action 'List resource groups' -Result 'That subscription has no ready resource group this account can see, so one has to be created.'
            return Read-NewAzureResourceGroup -Tool $Tool -SubscriptionId $SubscriptionId -DefaultLocation $defaultLocation
        }
        $options = [ordered]@{}
        for ($index = 0; $index -lt $groups.Count; $index++) { $options[[string]($index + 1)] = $groups[$index] }
        $options[[string]($groups.Count + 1)] = 'Create a new resource group'
        $options[[string]($groups.Count + 2)] = 'Type an existing resource group name instead'
        $choice = [int](Read-MenuChoice -Title 'Choose the resource group to hold the billing resource.' -Options $options -Default '1')
        if ($choice -eq $groups.Count + 1) { return Read-NewAzureResourceGroup -Tool $Tool -SubscriptionId $SubscriptionId -DefaultLocation $defaultLocation }
        $selected = if ($choice -gt $groups.Count) {
            Read-ValueWithDefault -Prompt 'Existing Azure resource group name' -Default ''
        }
        else { [string]$groups[$choice - 1] }
        if ([string]::IsNullOrWhiteSpace($selected)) { continue }
        if (Test-AzureResourceGroupReady -Tool $Tool -SubscriptionId $SubscriptionId -Name $selected) { return $selected }
        Write-RunLog -Severity WARN -Action 'Select resource group' -Result "'$selected' was removed from this menu because Azure could not confirm that it still exists and is ready. Choose another group or create a new one."
        $groups = @($groups | Where-Object { -not [string]::Equals([string]$_, $selected, [System.StringComparison]::OrdinalIgnoreCase) })
    }
}

function Wait-ForBillingLink {
    <# .SYNOPSIS Decides what to do about a billing link that failed, once repeating it here is pointless. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Pending,
        [Parameter(Mandatory)][object]$Tool,
        [int]$FailureCount = 1
    )

    # Nothing here waits on a timer: every route was already tried and refused the same way, so a
    # later attempt through the same code returns the same answer.
    Write-Host ''
    if ($script:LastBillingWasProviderFault) {
        Write-Host '  The documented create command reached Microsoft.GraphServices, which failed while' -ForegroundColor Yellow
        Write-Host '  loading its OpenTelemetry assemblies. This is not a problem' -ForegroundColor Yellow
        Write-Host '  with your subscription, tenant, resource group, or authorization.' -ForegroundColor Yellow
        Write-Host ''
        Write-Host '  The exact billing link is saved, but startup will not retry this provider fault.' -ForegroundColor Gray
        if ($script:LastBillingPreviewAttempted) {
            Write-Host '  The provider-advertised preview recovery was also tried once and did not verify.' -ForegroundColor Gray
        }
        else {
            Write-Host '  Choose option 3 again to use the same saved values and accept the provider-' -ForegroundColor Gray
            Write-Host '  advertised preview recovery if it is offered.' -ForegroundColor Gray
        }
        Write-RunLog -Severity WARN -Action 'Link metered billing' -Result 'The pending link was retained without an automatic retry because Microsoft.GraphServices returned an OpenTelemetry type-load failure.'
        return $false
    }
    if ($script:LastBillingWasClientFault) {
        Write-Host '  The Azure module could not load in its own process on this machine, so the request' -ForegroundColor Yellow
        Write-Host '  never reached Azure at all. Retrying it here would fail the same way.' -ForegroundColor Yellow
    }
    elseif ($Tool.Kind -eq 'AzureCli') {
        Write-Host '  The documented Azure CLI create command was refused, and read-only list/show' -ForegroundColor Yellow
        Write-Host '  verification found no completed billing link. No alternate create route was sent.' -ForegroundColor Yellow
    }
    else {
        Write-Host '  Every route this utility has was refused: the Az cmdlet, a direct request, and a' -ForegroundColor Yellow
        Write-Host '  template deployment, on each provider API version. Repeating them here cannot give' -ForegroundColor Yellow
        Write-Host '  a different answer, so nothing is retried on a timer.' -ForegroundColor Yellow
    }
    if ($FailureCount -ge 2) {
        Write-Host ''
        Write-Host "  The same fault has now come back $FailureCount times, in separate runs, so it is not a" -ForegroundColor Yellow
        Write-Host '  passing outage.' -ForegroundColor Yellow
    }
    Write-Host ''
    Write-Host '  Two things are still worth trying, and neither depends on this machine. Cloud Shell' -ForegroundColor Gray
    Write-Host '  is the route Microsoft documents, and runs in the browser with the Azure CLI already' -ForegroundColor Gray
    Write-Host '  installed. A different subscription only helps when it is enabled, in the same' -ForegroundColor Gray
    Write-Host '  tenant as the application, and your account has the required Contributor access.' -ForegroundColor Gray

    $choice = Read-MenuChoice -Title 'How should that be handled?' -Options ([ordered]@{
            '1' = 'Create it in Azure Cloud Shell, in the browser (default)'
            '2' = 'Try a different Azure subscription'
            '3' = 'Stop here; retry it quietly on later runs in case Azure recovers'
            '4' = 'Stop, and forget the pending link entirely'
        }) -Default '1'
    if ($choice -eq '1') {
        return (Invoke-BillingLinkInCloudShell -ClientId $Pending.ClientId -SubscriptionId $Pending.SubscriptionId `
                -ResourceGroup $Pending.ResourceGroup -ResourceName $Pending.ResourceName -TenantId $Pending.TenantId)
    }
    if ($choice -eq '2') { return (Invoke-BillingLinkOnOtherSubscription -Pending $Pending -Tool $Tool) }
    if ($choice -eq '4') {
        Clear-PendingBillingLink
        Write-RunLog -Severity INFO -Action 'Link metered billing' -Result 'The pending billing link was forgotten. Surveying still works, and option 3 can set it up again at any time.'
        return $false
    }
    Write-RunLog -Severity INFO -Action 'Link metered billing' -Result 'The link stays saved and is retried quietly the next time this utility starts.'
    return $false
}

function Invoke-BillingLinkOnOtherSubscription {
    <# .SYNOPSIS Repeats the billing link against a different subscription, which is the only remaining variable. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Pending,
        [Parameter(Mandatory)][object]$Tool
    )

    $others = @(Get-AzureSubscriptionList -Tool $Tool | Where-Object {
            $_.Id -ne $Pending.SubscriptionId -and
            ([string]::IsNullOrWhiteSpace($_.TenantId) -or [string]::Equals($_.TenantId, $Pending.TenantId, [System.StringComparison]::OrdinalIgnoreCase))
        })
    if ($others.Count -eq 0) {
        Write-RunLog -Severity WARN -Action 'Link metered billing' -Result 'That is the only Azure subscription this account can see in the application tenant, so there is no other one to try. Gain access to another enabled subscription in this tenant with Contributor rights, then choose option 3 again.'
        return $false
    }
    $options = [ordered]@{}
    for ($index = 0; $index -lt $others.Count; $index++) {
        $options[[string]($index + 1)] = '{0}  ({1})' -f $others[$index].Name, $others[$index].Id
    }
    $options[[string]($others.Count + 1)] = 'Cancel'
    $choice = [int](Read-MenuChoice -Title 'Which subscription should be billed instead?' -Options $options -Default '1')
    if ($choice -gt $others.Count) { return $false }
    $subscriptionId = [string]$others[$choice - 1].Id

    $resourceGroup = Select-AzureResourceGroup -Tool $Tool -SubscriptionId $subscriptionId
    if ([string]::IsNullOrWhiteSpace($resourceGroup)) { return $false }
    if (Set-MeteredBillingLink -ClientId $Pending.ClientId -SubscriptionId $subscriptionId `
            -ResourceGroup $resourceGroup -ResourceName $Pending.ResourceName -Tool $Tool `
            -TenantId $Pending.TenantId -Confirm:$false) {
        return $true
    }
    if ($script:LastBillingPreflightFailed) {
        Write-RunLog -Severity WARN -Action 'Link metered billing' -Result 'The documented create command was not attempted because the selected subscription did not pass billing preflight.'
        return $false
    }
    Save-PendingBillingLink -ClientId $Pending.ClientId -TenantId $Pending.TenantId -SubscriptionId $subscriptionId `
        -ResourceGroup $resourceGroup -ResourceName $Pending.ResourceName -Signature $script:LastBillingFailureSignature -FailureCount 1
    Write-RunLog -Severity WARN -Action 'Link metered billing' -Result 'That subscription was refused in the same way, which points at the provider rather than the subscription.'
    return $false
}

function Resume-PendingBillingLink {
    <# .SYNOPSIS Finishes a billing link left over from an earlier run, without asking. #>
    [CmdletBinding()]
    param()

    $pending = Get-PendingBillingLink
    if ($null -eq $pending) { return }
    if (Test-AzureProviderImplementationFailure -Message $pending.Signature) {
        Write-RunLog -Severity WARN -Action 'Link metered billing' -Result 'The pending link ended with a Microsoft.GraphServices OpenTelemetry type-load failure, so it is not retried automatically. Main-menu option 3 can reuse the saved values and offer the provider-advertised preview recovery interactively.'
        return
    }
    $tool = Get-AzureResourceTool
    if ($null -eq $tool) { return }
    if ($tool.Kind -eq 'AzureCli') {
        $context = Get-AzureCliAccountContext
        if ($null -eq $context -or
            -not [string]::Equals($context.TenantId, $pending.TenantId, [System.StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals($context.UserType, 'user', [System.StringComparison]::OrdinalIgnoreCase)) {
            Write-RunLog -Severity INFO -Action 'Link metered billing' -Result 'The pending billing link was not retried automatically because Azure CLI has no matching user session in the application tenant. Main-menu option 3 can open one tenant-pinned sign-in without changing the existing CLI profile.'
            return
        }
    }

    Write-RunLog -Severity INFO -Action 'Link metered billing' -Result "A billing link for application $($pending.ClientId) was left unfinished, so it is being retried now."
    if (Set-MeteredBillingLink -ClientId $pending.ClientId -SubscriptionId $pending.SubscriptionId `
            -ResourceGroup $pending.ResourceGroup -ResourceName $pending.ResourceName -Tool $tool `
            -TenantId $pending.TenantId -Quiet -Confirm:$false) {
        Write-Host ''
        Write-Host '  The Azure billing link left over from an earlier run has now completed,' -ForegroundColor Green
        Write-Host '  so applying labels in SharePoint is available again.' -ForegroundColor Green
        return
    }
    if ($script:LastBillingPreflightFailed) {
        Write-RunLog -Severity WARN -Action 'Link metered billing' -Result 'The saved billing link still fails prerequisite preflight, so it is not retried automatically. Correct the reported prerequisite, then choose option 3 to run preflight again.'
        return
    }
    # Each unattended retry that fails is counted, so the advice hardens rather than repeating itself.
    Save-PendingBillingLink -ClientId $pending.ClientId -TenantId $pending.TenantId -SubscriptionId $pending.SubscriptionId `
        -ResourceGroup $pending.ResourceGroup -ResourceName $pending.ResourceName `
        -Signature $script:LastBillingFailureSignature -FailureCount ($pending.FailureCount + 1)
    Write-RunLog -Severity WARN -Action 'Link metered billing' -Result "Azure refused it again, for the $($pending.FailureCount + 1) time. It stays saved, and option 3 now offers a different subscription, which is the only remaining variable."
}

function Read-MeteredBillingSetting {
    <# .SYNOPSIS Collects the Azure details and creates the billing link, or prints the commands to run elsewhere. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][AllowEmptyString()][string]$TenantId,
        [AllowEmptyString()][string]$Thumbprint = ''
    )

    if (-not [string]::IsNullOrWhiteSpace($Thumbprint) -and -not (Test-SigningCertificateAvailable -Thumbprint $Thumbprint)) {
        Write-RunLog -Severity WARN -Action 'Link metered billing' -Result "Certificate $Thumbprint is missing from this user's certificate store, so application $ClientId cannot authenticate. The billing command was not attempted. Replace the confidential client or restore its private-key certificate first."
        return $false
    }

    $tool = Get-AzureResourceTool
    if ($null -eq $tool) {
        # Offered before the modules, because it is what Microsoft documents and it cannot clash with PnP.
        $tool = Install-AzureCommandLine -Confirm:$false
    }
    if ($null -eq $tool) {
        # Az.Resources pulls Az.Accounts with it, and neither needs an administrator at CurrentUser scope.
        if (Request-ModuleInstall -Name 'Az.Resources' -Purpose 'to create the Azure billing resource') {
            $tool = Get-AzureResourceTool
        }
    }
    if ($null -eq $tool) {
        Write-RunLog -Severity WARN -Action 'Link metered billing' -Result 'No Azure tooling is available, so the billing resource cannot be created from here.'
        Write-RunLog -Severity INFO -Action 'Link metered billing' -Result 'Install one with: Install-Module Az.Resources -Scope CurrentUser    or    winget install --id Microsoft.AzureCLI'
        Show-MeteredBillingCommand -ClientId $ClientId -SubscriptionId '' -ResourceGroup '' -ResourceName ''
        return $false
    }
    Write-RunLog -Severity INFO -Action 'Link metered billing' -Result "Using $($tool.Kind) ($($tool.Detail))."

    if ($tool.Kind -eq 'AzureCli' -and -not (Connect-AzureCommandLine -Tool $tool -TenantId $TenantId)) {
        Write-RunLog -Severity WARN -Action 'Link metered billing' -Result 'A tenant-pinned Azure CLI user sign-in is required before application and billing prerequisites can be checked.'
        return $false
    }

    # Azure resolves the application during creation, and the client cannot write unless every requested role is consented.
    $consentState = if ($tool.Kind -eq 'AzureCli') {
        Get-AzureCliApplicationConsentState -ClientId $ClientId
    }
    else {
        $status = Test-ApplicationInDirectory -ClientId $ClientId -TenantId $TenantId -Collection 'servicePrincipals'
        [pscustomobject]@{Status = if ($status -eq 'Present') { 'Ready' } elseif ($status -eq 'Missing') { 'ConsentMissing' } else { 'Unknown' }; Count = 0; Message = "Microsoft Graph service-principal check returned $status."}
    }
    if ($consentState.Status -in 'Unknown', 'ApplicationMissing', 'ApplicationInvalid') {
        Write-RunLog -Severity WARN -Action 'Link metered billing' -Result "Application prerequisite check failed, so the billing command was not attempted: $($consentState.Message)"
        return $false
    }
    if ($consentState.Status -eq 'ConsentMissing') {
        Write-RunLog -Severity WARN -Action 'Link metered billing' -Result "Application $ClientId does not have complete administrator consent in tenant ${TenantId}: $($consentState.Message)"
        $consentFirst = Read-MenuChoice -Title 'Grant administrator consent first?' -Options ([ordered]@{
                '1' = 'Yes, grant it now, then link billing'
                '2' = 'No, return to the main menu'
            }) -Default '1'
        if ($consentFirst -ne '1') { return $false }
        if (-not (Grant-LabelingAdminConsent -ClientId $ClientId -TenantId $TenantId -PreferAzureCli -Confirm:$false)) { return $false }
            $consentState = if ($tool.Kind -eq 'AzureCli') {
                Get-AzureCliApplicationConsentState -ClientId $ClientId
            }
            else {
                Write-RunLog -Severity INFO -Action 'Link metered billing' -Result 'Waiting a few seconds for the new service principal to replicate.'
                Start-Sleep -Seconds 15
                $status = Test-ApplicationInDirectory -ClientId $ClientId -TenantId $TenantId -Collection 'servicePrincipals'
                [pscustomobject]@{Status = if ($status -eq 'Present') { 'Ready' } elseif ($status -eq 'Missing') { 'ConsentMissing' } else { 'Unknown' }; Count = 0; Message = "Microsoft Graph service-principal check returned $status."}
            }
            if ($consentState.Status -ne 'Ready') {
                Write-RunLog -Severity WARN -Action 'Link metered billing' -Result "Application $ClientId still does not have verifiable administrator consent in tenant $TenantId after consent: $($consentState.Message) The billing command was not attempted."
            return $false
        }
    }

    $choice = Read-MenuChoice -Title 'Link this application to an Azure subscription now?' -Options ([ordered]@{
            '1' = 'Yes, look up my subscriptions and create the billing resource'
            '2' = 'No, just show me the commands to run elsewhere'
        }) -Default '1'
    if ($choice -ne '1') {
        Show-MeteredBillingCommand -ClientId $ClientId -SubscriptionId '' -ResourceGroup '' -ResourceName ''
        return $false
    }

    $subscriptionId = Select-AzureSubscription -Tool $tool -TenantId $TenantId
    if ([string]::IsNullOrWhiteSpace($subscriptionId)) {
        $subscriptionId = Read-ValueWithDefault -Prompt 'Azure subscription ID that should receive the charges' -Default ''
    }
    $parsedSubscription = [guid]::Empty
    if (-not [guid]::TryParse($subscriptionId, [ref]$parsedSubscription) -or $parsedSubscription -eq [guid]::Empty) {
        Write-RunLog -Severity WARN -Action 'Validate subscription' -Result "'$subscriptionId' is not a valid subscription GUID."
        return $false
    }
    $knownSubscription = @(Get-AzureSubscriptionList -Tool $tool | Where-Object { $_.Id -eq $parsedSubscription.ToString() } | Select-Object -First 1)
    if ($knownSubscription.Count -eq 0) {
        Write-RunLog -Severity WARN -Action 'Validate subscription' -Result "Subscription $($parsedSubscription.ToString()) could not be verified from the Azure account. Sign in with an account that can access it, then run billing setup again."
        return $false
    }
    if (-not [string]::IsNullOrWhiteSpace($TenantId) -and -not [string]::IsNullOrWhiteSpace($knownSubscription[0].TenantId) -and
        -not [string]::Equals($knownSubscription[0].TenantId, $TenantId, [System.StringComparison]::OrdinalIgnoreCase)) {
        Write-RunLog -Severity WARN -Action 'Validate subscription' -Result "Subscription $($parsedSubscription.ToString()) belongs to tenant $($knownSubscription[0].TenantId), not application tenant $TenantId."
        return $false
    }
    $resourceGroup = Select-AzureResourceGroup -Tool $tool -SubscriptionId $parsedSubscription.ToString()
    if ([string]::IsNullOrWhiteSpace($resourceGroup)) { return $false }
    $resourceName = Read-ValueWithDefault -Prompt 'Name for the billing resource' -Default 'purview-file-labeling-billing'

    if (-not (Set-MeteredBillingLink -ClientId $ClientId -SubscriptionId $parsedSubscription.ToString() -ResourceGroup $resourceGroup -ResourceName $resourceName -Tool $tool -TenantId $TenantId -Confirm:$false)) {
        if ($script:LastBillingPreflightFailed) {
            Write-RunLog -Severity WARN -Action 'Link metered billing' -Result 'The documented billing association command was not attempted because preflight found a subscription, tenant, resource-group, provider, or authorization problem.'
            return $false
        }
        # Saved whatever the cause, so the choice that follows never has to ask for these details again.
        $previous = Get-PendingBillingLink
        $failures = 1
        if ($null -ne $previous -and $previous.Signature -eq $script:LastBillingFailureSignature.Substring(0, [math]::Min(120, $script:LastBillingFailureSignature.Length))) {
            $failures = $previous.FailureCount + 1
        }
        Save-PendingBillingLink -ClientId $ClientId -TenantId $TenantId -SubscriptionId $parsedSubscription.ToString() `
            -ResourceGroup $resourceGroup -ResourceName $resourceName -Signature $script:LastBillingFailureSignature -FailureCount $failures
        Write-RunLog -Severity INFO -Action 'Link metered billing' -Result 'The link has been saved exactly as configured, so it can be completed without asking anything again.'
        $pending = Get-PendingBillingLink
        if ($null -eq $pending) { return $false }
        return (Wait-ForBillingLink -Pending $pending -Tool $tool -FailureCount $failures)
    }
    Write-Host ''
    Write-Host '  Billing is linked. Applying labels is now offered for the SharePoint Online source.' -ForegroundColor Green
    Write-Host '  A token issued before the link may still be refused, so if the first apply run' -ForegroundColor Gray
    Write-Host '  reports paymentRequired, start the utility again to obtain a fresh token.' -ForegroundColor Gray
    return $true
}

function Get-LabelingRememberedVariable {
    <# .SYNOPSIS Lists the environment variables this utility owns and currently has set. #>
    [CmdletBinding()]
    param()

    $owned = [System.Collections.Generic.List[object]]::new()
    foreach ($target in [EnvironmentVariableTarget]::Process, [EnvironmentVariableTarget]::User) {
        $all = $null
        try { $all = [Environment]::GetEnvironmentVariables($target) }
        catch { Write-Verbose "Could not read the $target environment: $($_.Exception.Message)"; continue }
        foreach ($name in @($all.Keys)) {
            if ([string]$name -notmatch '^(PURVIEW_FILE_LABELING_|LABEL_TEST_SITE_)') { continue }
            # Clearing a variable can leave an empty entry behind, which is nothing to report or clear again.
            if ([string]::IsNullOrWhiteSpace([string]$all[$name])) { continue }
            $owned.Add([pscustomobject]@{
                    Name = [string]$name
                    Scope = [string]$target
                    Value = [string]$all[$name]
                })
        }
    }
    return @($owned)
}

function Clear-LabelingRememberedValue {
    <# .SYNOPSIS Forgets everything this utility remembers, so no earlier run supplies a default. #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param()

    $owned = @(Get-LabelingRememberedVariable)
    if ($owned.Count -eq 0) {
        Write-RunLog -Severity INFO -Action 'Forget settings' -Result 'Nothing is remembered, so there is nothing to clear.'
    }
    else {
        Write-Host ''
        Write-Host '  Remembered by this utility' -ForegroundColor Cyan
        foreach ($item in ($owned | Sort-Object Name, Scope)) {
            Write-Host ('    {0,-48} {1,-8} {2}' -f $item.Name, $item.Scope, $item.Value)
        }
    }

    # These belong to other tools, so they are reported but never modified.
    $foreign = [System.Collections.Generic.List[string]]::new()
    foreach ($name in 'ENTRAID_APP_ID', 'ENTRAID_CLIENT_ID', 'AZURE_CLIENT_ID') {
        foreach ($target in [EnvironmentVariableTarget]::Process, [EnvironmentVariableTarget]::User) {
            try {
                $value = [Environment]::GetEnvironmentVariable($name, $target)
                if (-not [string]::IsNullOrWhiteSpace($value)) { $foreign.Add("$name ($target) = $value") }
            }
            catch { Write-Verbose "Could not read ${name}: $($_.Exception.Message)" }
        }
    }
    if ($foreign.Count -gt 0) {
        Write-Host ''
        Write-Host '  Set by other tools, and left untouched' -ForegroundColor Yellow
        foreach ($item in $foreign) { Write-Host "    $item" -ForegroundColor Yellow }
        Write-Host '  These can still be offered as application IDs, flagged as possibly belonging' -ForegroundColor Gray
        Write-Host '  to another tenant. Remove them yourself if nothing else needs them.' -ForegroundColor Gray
    }

    if ($owned.Count -eq 0) { return }
    $choice = Read-MenuChoice -Title 'Clear everything listed as remembered by this utility?' -Options ([ordered]@{
            '1' = 'No, keep them (default)'
            '2' = 'Yes, clear them all'
        }) -Default '1'
    if ($choice -ne '2') { return }
    if (-not $PSCmdlet.ShouldProcess('remembered settings', 'Clear')) { return }

    foreach ($item in $owned) {
        try {
            [Environment]::SetEnvironmentVariable($item.Name, $null, [EnvironmentVariableTarget]$item.Scope)
            Write-RunLog -Severity SUCCESS -Action 'Forget settings' -Result "Cleared $($item.Name) ($($item.Scope) scope)."
        }
        catch {
            Add-RunFailure -FilePath '' -Action 'Forget settings' -Reason "Could not clear $($item.Name): $(Get-ErrorText -ErrorRecord $_)"
        }
    }
    $script:RejectedClientIds.Clear()
    Disconnect-LabelingGraph
    Write-RunLog -Severity INFO -Action 'Forget settings' -Result 'The certificate of any confidential client is still in your personal certificate store, and the applications themselves still exist in Entra ID.'
}

function Invoke-GuidedRun {
    <# .SYNOPSIS Runs one complete guided scan or label operation. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][ValidateSet('Local', 'SharePoint')][string]$Source)

    $source = $Source
    while ($true) {
        $labels = @(Get-TenantSensitivityLabel | Where-Object { $null -ne $_ })
        if ($labels.Count -gt 0) { break }
        $action = Read-PhaseAction -Phase 'Tenant label discovery'
        if ($action -eq 'Exit') { return 'Exit' }
        if ($action -eq 'Main' -or $action -eq 'Skip') { return 'Main' }
    }
    $action = Read-PhaseAction -Phase 'Tenant label discovery'
    if ($action -eq 'Exit') { return 'Exit' }
    if ($action -eq 'Main' -or $action -eq 'Skip') { return 'Main' }

    $pendingSettings = $null
    while ($true) {
        # Set when the operator chose another folder in the same library, so nothing is asked twice.
        $reusing = $null -ne $pendingSettings
        if ($reusing) {
            $settings = $pendingSettings
            $pendingSettings = $null
        }
        else {
            $settings = Read-RunSetting -Labels $labels -Source $source
            if ($null -eq $settings) { return 'Main' }
        }
        if (-not $reusing -and -not (Initialize-RunArtifact -Folder $settings.LogFolder -Confirm:$false)) {
            $action = Read-PhaseAction -Phase 'Logging setup'
            if ($action -eq 'Exit') { return 'Exit' }
            if ($action -eq 'Main' -or $action -eq 'Skip') { return 'Main' }
            continue
        }
        Show-SettingsSummary -Settings $settings
        $action = Read-PhaseAction -Phase 'Settings'
        if ($action -eq 'Exit') { return 'Exit' }
        if ($action -eq 'Main' -or $action -eq 'Skip') { return 'Main' }
        if ($action -eq 'Change') { continue }

        Write-Host ''
        Write-Host '  Scanning for files' -ForegroundColor Cyan
        Write-Host "    Looking in : $($settings.TargetPath)"
        Write-Host "    Subfolders : $($settings.IncludeSubfolders)"
        Write-Host "    Extensions : $($settings.Extensions -join ', ')"
        if ($settings.Source -eq 'SharePoint') {
            $files = Get-SharePointTargetFile -SiteUrl $settings.SiteUrl -FolderSiteRelativeUrl $settings.TargetPath -LibraryTitle $settings.LibraryTitle -Extensions $settings.Extensions -Recurse $settings.IncludeSubfolders
        }
        else {
            if (-not (Test-PurviewPrerequisite)) {
                # Installing the client schedules a restart, so there is nothing more to ask here.
                if ($script:RelaunchCompleted) { return 'Exit' }
                $action = Read-PhaseAction -Phase 'Prerequisite check'
                if ($action -eq 'Exit') { return 'Exit' }
                if ($action -eq 'Main' -or $action -eq 'Skip') { return 'Main' }
                continue
            }
            if (-not (Test-TargetPath -Path $settings.TargetPath)) {
                $action = Read-PhaseAction -Phase 'Path validation'
                if ($action -eq 'Exit') { return 'Exit' }
                if ($action -eq 'Main' -or $action -eq 'Skip') { return 'Main' }
                continue
            }
            $files = Get-TargetFile -Path $settings.TargetPath -Extensions $settings.Extensions -Recurse $settings.IncludeSubfolders
        }
        if ($null -eq $files) { continue }
        $files = @($files)
        if ($files.Count -eq 0) {
            Write-RunLog -Severity WARN -FilePath $settings.TargetPath -Action 'Enumerate files' -Result 'Nothing matched. Check the folder name, the extension list, and whether subfolders need to be included.'
            $action = Read-PhaseAction -Phase 'Enumeration'
            if ($action -eq 'Exit') { return 'Exit' }
            if ($action -eq 'Main' -or $action -eq 'Skip') { return 'Main' }
            continue
        }
        $action = Read-PhaseAction -Phase 'Enumeration'
        if ($action -eq 'Exit') { return 'Exit' }
        if ($action -eq 'Main' -or $action -eq 'Skip') { return 'Main' }
        if ($action -eq 'Change') { continue }

        while (-not (Test-LabelReadAccess -ProbeFile $files[0] -Source $settings.Source)) {
            $action = Read-PhaseAction -Phase 'Authentication check'
            if ($action -eq 'Exit') { return 'Exit' }
            if ($action -eq 'Main' -or $action -eq 'Skip') { return 'Main' }
            if ($action -eq 'Change') { break }
        }
        if ($action -eq 'Change') { continue }

        if (-not $settings.DryRun) {
            Show-SettingsSummary -Settings $settings
            $confirm = Read-MenuChoice -Title 'WRITE OPERATION: Apply the selected label to eligible files?' -Options ([ordered]@{
                    '1' = 'No, return to settings (default)'
                    '2' = 'Yes, apply labels'
                }) -Default '1'
            if ($confirm -ne '2') { continue }
        }

        $processingResult = Invoke-FileProcessing -Files $files -TargetLabel $settings.TargetLabel -ConfiguredLabels $labels -DryRun $settings.DryRun -Source $settings.Source
        # Processing has several early exits, and a bar left uncompleted stays on screen over every later prompt.
        Write-Progress -Id 1 -Activity 'Processing files' -Completed
        Export-RunReport -Confirm:$false
        Show-RunSummary
        if ($processingResult -eq 'Change') { continue }
        if ($processingResult -eq 'Main') { return 'Main' }

        # The site, library, sign-in, label, and filters are all still valid, so another folder
        # costs one question rather than a restart.
        if ($settings.Source -eq 'SharePoint' -and -not [string]::IsNullOrWhiteSpace($settings.LibraryUrl)) {
            $another = Read-MenuChoice -Title "Scan another folder in '$($settings.LibraryTitle)' without starting over?" -Options ([ordered]@{
                    '1' = 'No, continue (default)'
                    '2' = 'Yes, choose another folder in this library'
                }) -Default '1'
            if ($another -eq '2') {
                $pendingSettings = Read-NextSharePointFolder -Settings $settings
                continue
            }
        }

        $action = Read-PhaseAction -Phase 'File processing and reporting'
        if ($action -eq 'Exit') { return 'Exit' }
        if ($action -eq 'Change') { continue }
        return 'Main'
    }
}

Write-Host ''
Write-Host '  Microsoft Purview File Labeling Utility' -ForegroundColor Cyan
# Printed because running an old extracted copy by mistake is otherwise invisible.
$script:BuildStamp = ''
try {
    if (-not [string]::IsNullOrWhiteSpace($script:ScriptPath) -and (Test-Path -LiteralPath $script:ScriptPath -PathType Leaf)) {
        $script:BuildStamp = (Get-Item -LiteralPath $script:ScriptPath -ErrorAction Stop).LastWriteTime.ToString('yyyy-MM-dd HH:mm')
    }
}
catch { Write-Verbose "Could not read this script's timestamp: $($_.Exception.Message)" }
$buildText = if ($script:BuildStamp) { "version $script:UtilityVersion, file dated $script:BuildStamp" } else { "version $script:UtilityVersion" }
Write-Host "  $buildText" -ForegroundColor DarkGray
Write-Host "  $script:ScriptPath" -ForegroundColor DarkGray
Write-Host '  Dry run is the default. No file is changed unless Apply mode is selected and confirmed.' -ForegroundColor Gray
if ($script:IgnoreRemembered) {
    Write-Host '  -Fresh: every value remembered from a previous run is ignored, and nothing is deleted.' -ForegroundColor Yellow
}
if ($Restarted) {
    Write-Host "  Restarted in PowerShell $($PSVersionTable.PSVersion). The SharePoint Online source is available here." -ForegroundColor Green
}
elseif ($PSVersionTable.PSVersion -lt [version]'7.2.0') {
    Write-Host '  This host labels files on local/UNC paths and mounted SharePoint Server libraries.' -ForegroundColor Yellow
    Write-Host '  Choosing SharePoint Online offers to restart in PowerShell 7.' -ForegroundColor Yellow
}
Write-RunLog -Severity INFO -Action 'Start utility' -Result 'Microsoft Purview file labeling utility started. No write occurs unless Apply mode is selected and confirmed.' -NoConsole
Write-RunEnvironment
Test-DependencyDrift
# A killed process cannot execute finally, so recover its temporary workers and certificate exports
# before this run prompts, signs in, or makes a network request.
Clear-LabelingTemporaryArtifact
Clear-LegacyTenantContextPreference
$exitRequested = $false
try {
    # Asked once, before the menu, so an unreachable source costs no sign-in and the question is not repeated for every run.
    $script:Source = if (-not [string]::IsNullOrWhiteSpace($InitialSource)) {
        Write-RunLog -Severity INFO -Action 'Select source' -Result "Keeping the $InitialSource source chosen before the restart."
        $InitialSource
    }
    else {
        Read-FileSource
    }
    if ([string]::IsNullOrWhiteSpace($script:Source)) { $exitRequested = $true }

    # An earlier run may have been stopped by an Azure outage part-way through the billing step.
    Resume-PendingBillingLink

    # A run that crashed or was closed could not clean up after itself, so its leftovers are swept here.
    Clear-OrphanedLabelingCertificate

    while (-not $exitRequested) {
        $sourceLabel = if ($script:Source -eq 'SharePoint') { 'SharePoint Online (metered API for Apply)' } else { 'local/UNC/SharePoint Server path (Purview client)' }
        $mainChoice = Read-MenuChoice -Title "Main menu  (files: $sourceLabel)" -Options ([ordered]@{
                '1' = "Start a guided run (one label for one folder's contents)"
                '2' = "Start a batch run from a CSV (a different label for each folder's contents)"
                '3' = 'Enable SharePoint Online metered label writing (certificate and Azure billing)'
                '4' = "Change where the files are (currently $sourceLabel)"
                '5' = 'Forget settings remembered from previous runs (and sign out of Microsoft Graph)'
                '6' = 'Show current session summary'
                '7' = 'Exit cleanly'
            }) -Default '1'
        switch ($mainChoice) {
            '1' {
                $result = Invoke-GuidedRun -Source $script:Source
                if ($result -eq 'Exit') { $exitRequested = $true }
            }
            '2' {
                $result = Invoke-BatchRun -Source $script:Source
                if ($result -eq 'Exit') { $exitRequested = $true }
            }
            '3' {
                $result = Invoke-MeteredSetup
                if ($result -eq 'Exit') { $exitRequested = $true }
            }
            '4' {
                $changedSource = Read-FileSource
                if ([string]::IsNullOrWhiteSpace($changedSource)) { $exitRequested = $true }
                else { $script:Source = $changedSource }
            }
            '5' { Clear-LabelingRememberedValue -Confirm:$false }
            '6' { Show-RunSummary }
            '7' { $exitRequested = $true }
        }
        # A restart replaces this process, so the menu must not be shown again.
        if ($script:RelaunchCompleted) { $exitRequested = $true }
    }
    if (-not $script:RelaunchCompleted) {
        Write-RunLog -Severity SUCCESS -Action 'Exit utility' -Result 'The operator requested a clean exit.'
    }
}
finally {
    Write-Progress -Id 1 -Activity 'Processing files' -Completed
    if ($script:FilesSinceCollection -gt 0) { Clear-PurviewHandle }
    if (-not $script:RelaunchCompleted) {
        Export-RunReport -Confirm:$false
        Show-RunSummary
    }
    Disconnect-LabelDiscoveryService
    Disconnect-SharePointSession
    if ($script:GraphSessionOpened) { Disconnect-LabelingGraph }
    Disconnect-AzureSession
    # Nothing this run generated is left behind unless the saved client still needs it.
    Clear-UnusedLabelingCertificate
    Clear-LabelingTemporaryArtifact -IncludeCurrentProcess
    # Swept again by name, because a certificate abandoned earlier in this run was never recorded
    # here, and a restart is not an ending so it keeps its work.
    if (-not $script:RelaunchCompleted) { Clear-OrphanedLabelingCertificate }
}

# Started here, not inside a function, so the restarted run writes straight to this console.
if (-not [string]::IsNullOrWhiteSpace($script:RelaunchHostPath)) {
    $heading = if ($script:RelaunchIsModuleUpdate) { 'Restarting the utility' } else { 'Restarting in PowerShell 7' }
    $notice = @()
    if (-not [string]::IsNullOrWhiteSpace($script:RelaunchReason)) { $notice += $script:RelaunchReason }
    $notice += $script:RelaunchHostPath
    $notice += ''
    $notice += 'The utility starts again below, in this same window. The source you'
    $notice += 'already chose is carried over, so it is not asked again; the main menu'
    $notice += 'reappears ready to continue.'
    Write-Host ''
    Write-Banner -Title $heading -Body $notice

    $relaunchArguments = @('-NoLogo', '-NoProfile', '-File', $script:ScriptPath, '-NoRelaunch', '-Restarted')
    if ($script:UseDeviceCode) { $relaunchArguments += '-DeviceLogin' }
    if ($script:IgnoreRemembered) { $relaunchArguments += '-Fresh' }
    if ($script:RelaunchIsModuleUpdate) { $relaunchArguments += '-ModuleRestarted' }
    # A PowerShell 7 handover happens before the source is stored, and it only ever happens for SharePoint.
    $carrySource = if (-not [string]::IsNullOrWhiteSpace($script:Source)) { $script:Source } else { 'SharePoint' }
    $relaunchArguments += @('-InitialSource', $carrySource)
    & $script:RelaunchHostPath @relaunchArguments

    Write-Host ''
    Write-Host '  The PowerShell 7 session ended, so this one is finished too.' -ForegroundColor Cyan
    Write-Host ''
}
