<#
.SYNOPSIS
Imports NSG rule intent from an Excel workbook, compares it with a target Azure NSG, and optionally applies the required changes.

.DESCRIPTION
This script validates workbook structure and content, loads the current NSG rules from Azure,
builds an execution plan, writes a checkpoint file, and optionally applies Create and Update
operations to the target NSG while skipping shadowed rules.

The script is designed for Windows PowerShell 5.1 and Azure CLI based execution.

High-level flow:
1. Validate prerequisites unless -SkipPreflight is used.
2. Read workbook rules from the configured worksheets.
3. Normalize and validate workbook values.
4. Load live NSG rules from Azure.
5. Build a plan containing NoChange, Create, Update, SkipApply, and Conflict actions.
6. Save a checkpoint JSON file for the current run.
7. If -Apply is provided, execute Create and Update actions that are not shadowed.

.PARAMETER WorkbookPath
Path to the Excel workbook that contains the desired NSG rules.

.PARAMETER ResourceGroupName
Azure resource group that contains the target NSG.

.PARAMETER NsgName
Name of the target Network Security Group.

.PARAMETER SheetNames
Workbook sheet names to import. Defaults to CoreRules and AppRules.

.PARAMETER Direction
Direction applied to imported rules. Allowed values are Inbound and Outbound.

.PARAMETER CheckpointPath
Base path used for checkpoint output.

.PARAMETER CheckpointMode
Controls how checkpoint files are written.

Auto:
- Dry-run writes only the stable CheckpointPath file.
- Apply writes both a timestamped run file and the stable CheckpointPath file.

Overwrite:
- Always writes only the stable CheckpointPath file.

Timestamped:
- Always writes only a timestamped run file.

Both:
- Writes both a timestamped run file and the stable CheckpointPath file.

.PARAMETER Apply
When provided, executes Create and Update actions against Azure.
Rules detected as shadowed by a higher-precedence desired rule are logged and skipped.
Without this switch, the script runs in dry-run mode.

.PARAMETER FailOnShadowing
Stops execution if overlap warnings are detected during rule validation.

.PARAMETER SkipPreflight
Skips dependency checks, Azure login verification, and NSG existence validation.

.PARAMETER PassThru
Returns the execution/checkpoint objects to the pipeline.

.PARAMETER RetryCount
Maximum retry attempts for Azure CLI calls.

.PARAMETER RetryDelaySeconds
Delay between Azure CLI retry attempts.

.PARAMETER Description
Managed description applied to new and updated rules.
Expected format: RITMxxxx - NYxxxx - mm/dd/yyyy - Create|Update

.EXAMPLE
.\v5apply_nsg_rules.ps1 `
    -WorkbookPath .\dummyData.xlsx `
    -ResourceGroupName "NSG-test-script" `
    -NsgName "nsg-ci-automation-app" `
    -Description "RITM12345 - NY5678 - 03/20/2026 - Update"

Runs a dry-run and updates the checkpoint file without changing Azure resources.

.EXAMPLE
.\v5apply_nsg_rules.ps1 `
    -WorkbookPath .\dummyData.xlsx `
    -ResourceGroupName "NSG-test-script" `
    -NsgName "nsg-ci-automation-app" `
    -Description "RITM12345 - NY5678 - 03/20/2026 - Update" `
    -Apply `
    -CheckpointMode Both

Applies Create and Update actions that are not shadowed and stores both a stable checkpoint file and a timestamped archive.

.EXAMPLE
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\v5apply_nsg_rules.ps1 `
    -WorkbookPath .\dummyData.xlsx `
    -ResourceGroupName "NSG-test-script" `
    -NsgName "nsg-ci-automation-app" `
    -Description "RITM12345 - NY5678 - 03/20/2026 - Update" `
    -SkipPreflight `
    -PassThru

Shows a PowerShell 5.1-friendly execution pattern when local execution policy blocks unsigned scripts.

.NOTES
Dependencies:
- Azure CLI (`az`)
- ImportExcel PowerShell module

The workbook must contain the required headers expected by Get-WorkbookHeaderMap.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ResourceGroupName,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$NsgName,

    [string[]]$SheetNames = @('CoreRules', 'AppRules'),

    [ValidateSet('Inbound', 'Outbound')]
    [string]$Direction = 'Inbound',

    [string]$CheckpointPath = './nsg-apply-checkpoint.json',

    [ValidateSet('Auto', 'Overwrite', 'Timestamped', 'Both')]
    [string]$CheckpointMode = 'Auto',

    [switch]$Apply,
    [switch]$FailOnShadowing,
    [switch]$SkipPreflight,
    [switch]$PassThru,

    [ValidateRange(1, 20)]
    [int]$RetryCount = 3,

    [ValidateRange(1, 300)]
    [int]$RetryDelaySeconds = 5,

    [Parameter(Mandatory = $true)]
    [ValidateLength(1, 140)]
    [string]$Description
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:AddressTokenCache = @{}
$script:PortTokenCache = @{}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO',

        [Parameter()]
        [object]$Color = $null
    )

    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts][$Level] $Message"

    # Determine color to use. If caller provided a Color parameter, try to parse it
    # to a ConsoleColor (supports either enum value or string). Otherwise pick a
    # sensible default per log level. This keeps behaviour consistent on PS 5.1.
    $colorToUse = $null
    if ($PSBoundParameters.ContainsKey('Color') -and $null -ne $Color -and -not [string]::IsNullOrWhiteSpace([string]$Color)) {
        try {
            $colorToUse = [Enum]::Parse([System.ConsoleColor], [string]$Color, $true)
        }
        catch {
            $colorToUse = [System.ConsoleColor]::Gray
        }
    }
    else {
        switch ($Level) {
            'INFO'  { $colorToUse = [System.ConsoleColor]::Gray }
            'WARN'  { $colorToUse = [System.ConsoleColor]::Yellow }
            'ERROR' { $colorToUse = [System.ConsoleColor]::Red }
            'DEBUG' { $colorToUse = [System.ConsoleColor]::Gray }
            default { $colorToUse = [System.ConsoleColor]::Gray }
        }
    }

    if ($Level -eq 'DEBUG') {
        Write-Verbose $line
    }
    else {
        Write-Host $line -ForegroundColor $colorToUse
    }
}

function Assert-Dependency {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CommandName,

        [string]$InstallHint = ''
    )

    if (-not (Get-Command $CommandName -ErrorAction SilentlyContinue)) {
        $hint = if ($InstallHint) { " $InstallHint" } else { '' }
        throw "Required command '$CommandName' was not found.$hint"
    }
}

function Assert-Module {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModuleName,

        [string]$InstallHint = ''
    )

    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        $hint = if ($InstallHint) { " $InstallHint" } else { '' }
        throw "Required PowerShell module '$ModuleName' was not found.$hint"
    }
}

function Invoke-WithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,

        [int]$MaxAttempts = 3,
        [int]$DelaySeconds = 5,
        [string]$OperationName = 'operation'
    )

    $attempt = 0
    do {
        $attempt++
        try {
            return & $ScriptBlock
        }
        catch {
            if ($attempt -ge $MaxAttempts) {
                throw "Failed $OperationName after $attempt attempt(s). Last error: $($_.Exception.Message)"
            }

            Write-Log -Level WARN -Message "Attempt $attempt/$MaxAttempts failed for $OperationName. Retrying in $DelaySeconds second(s). Error: $($_.Exception.Message)"
            Start-Sleep -Seconds $DelaySeconds
        }
    } while ($true)
}

function Invoke-AzCommandText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments,

        [int]$RetryCount = 3,
        [int]$RetryDelaySeconds = 5,
        [string]$OperationName = 'az command'
    )

    return Invoke-WithRetry -MaxAttempts $RetryCount -DelaySeconds $RetryDelaySeconds -OperationName $OperationName -ScriptBlock {
        $output = & az @Arguments 2>&1
        $exitCode = $LASTEXITCODE
        $text = ($output | Out-String).Trim()

        if ($exitCode -ne 0) {
            throw "Azure CLI failed with exit code $exitCode. Output: $text"
        }

        return $text
    }
}

function Invoke-AzCliJson {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments,

        [int]$RetryCount = 3,
        [int]$RetryDelaySeconds = 5,
        [string]$OperationName = 'az command'
    )

    $jsonText = Invoke-AzCommandText -Arguments $Arguments -RetryCount $RetryCount -RetryDelaySeconds $RetryDelaySeconds -OperationName $OperationName
    if ([string]::IsNullOrWhiteSpace($jsonText)) {
        return $null
    }

    return $jsonText | ConvertFrom-Json
}

function Convert-RuleToCheckpointView {
    [CmdletBinding()]
    param(
        [AllowNull()][object]$Rule
    )

    if ($null -eq $Rule) {
        return $null
    }

    return [pscustomobject]@{
        Name                       = $Rule.Name
        Priority                   = $Rule.Priority
        Direction                  = $Rule.Direction
        Access                     = $Rule.Access
        Protocol                   = $Rule.Protocol
        SourcePortRanges           = @($Rule.SourcePortRanges)
        SourceAddressPrefixes      = @($Rule.SourceAddressPrefixes)
        DestinationAddressPrefixes = @($Rule.DestinationAddressPrefixes)
        DestinationPortRanges      = @($Rule.DestinationPortRanges)
        Description                = $Rule.Description
    }
}

function Compare-RuleForCheckpoint {
    [CmdletBinding()]
    param(
        [AllowNull()][object]$Before,
        [AllowNull()][object]$After
    )

    $changes = New-Object System.Collections.Generic.List[object]

    $fields = @(
        'Name',
        'Priority',
        'Direction',
        'Access',
        'Protocol',
        'SourcePortRanges',
        'SourceAddressPrefixes',
        'DestinationAddressPrefixes',
        'DestinationPortRanges',
        'Description'
    )

    foreach ($field in $fields) {
        if ($null -ne $Before) {
            $beforeValue = $Before.$field
        }
        else {
            $beforeValue = $null
        }

        if ($null -ne $After) {
            $afterValue = $After.$field
        }
        else {
            $afterValue = $null
        }

        if ($beforeValue -is [System.Array]) {
            $beforeValue = @($beforeValue)
        }

        if ($afterValue -is [System.Array]) {
            $afterValue = @($afterValue)
        }

        if ($null -eq $beforeValue) {
            $beforeJson = 'null'
        }
        else {
            $beforeJson = ConvertTo-Json -InputObject $beforeValue -Compress -Depth 10
        }

        if ($null -eq $afterValue) {
            $afterJson = 'null'
        }
        else {
            $afterJson = ConvertTo-Json -InputObject $afterValue -Compress -Depth 10
        }

        if ($beforeJson -ne $afterJson) {
            $changes.Add([pscustomobject]@{
                Field  = $field
                Before = $beforeValue
                After  = $afterValue
            })
        }
    }

    return $changes.ToArray()
}

function Get-WorkbookHeaderMap {
    [CmdletBinding()]
    param()

    return [ordered]@{
        RuleName           = 'Rule Name'
        RuleNumber         = 'Rule Number'
        Protocol           = 'Destination Protocol'
        SourceAddress      = 'Source IP adress / Subnet / Range IP'
        DestinationAddress = 'Destionation IP adress / Subnet / Range IP'
        DestinationPorts   = 'Destination Port or Service'
    }
}

function Assert-WorkbookSchema {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string[]]$SheetNames
    )

    $package = Open-ExcelPackage -Path $Path
    try {
        $requiredHeaders = @((Get-WorkbookHeaderMap).Values)

        foreach ($sheetName in $SheetNames) {
            $worksheet = $package.Workbook.Worksheets[$sheetName]
            if (-not $worksheet) {
                throw "Worksheet '$sheetName' was not found in workbook '$Path'."
            }

            if (-not $worksheet.Dimension) {
                throw "Worksheet '$sheetName' is empty and does not contain the required header row."
            }

            $headers = @()
            for ($column = 1; $column -le $worksheet.Dimension.End.Column; $column++) {
                $headers += [string]$worksheet.Cells[1, $column].Text
            }

            $missing = @($requiredHeaders | Where-Object { $_ -notin $headers })
            if ($missing.Count -gt 0) {
                throw "Worksheet '$sheetName' is missing required header(s): $($missing -join ', ')"
            }
        }
    }
    finally {
        if ($null -ne $package) {
            $package.Dispose()
        }
    }
}

function Convert-AnyToken {
    [CmdletBinding()]
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return '*' }

    $text = [string]$Value
    $text = $text.Trim()

    if ([string]::IsNullOrWhiteSpace($text)) { return '*' }
    if ($text -in @('*', 'Any', 'ANY', 'any')) { return '*' }

    return $text
}

function Split-NormalizedList {
    [CmdletBinding()]
    param(
        [AllowNull()][object]$Value,
        [string[]]$Delimiters = @(';', ','),
        [switch]$KeepAnyAsWildcard
    )

    $raw = Convert-AnyToken -Value $Value
    if ($raw -eq '*') {
        if ($KeepAnyAsWildcard) { return @('*') }
        return @()
    }

    $raw = $raw -replace "`r`n|`n|`r", ';'

    $pattern = '[' + (($Delimiters | ForEach-Object { [Regex]::Escape($_) }) -join '') + ']'
    $tokens = @(
        $raw -split $pattern |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    if ($tokens.Count -eq 0) {
        return @('*')
    }

    $normalized = foreach ($token in $tokens) {
        if ($token -in @('Any', 'ANY', 'any', '*')) { '*' } else { $token }
    }

    if ($normalized -contains '*') { return @('*') }
    return @($normalized | Select-Object -Unique)
}

function Test-ValidIpv4 {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Value)

    $ip = $null
    return [System.Net.IPAddress]::TryParse($Value, [ref]$ip) -and $ip.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork
}

function Test-ValidCidr {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Value)

    if ($Value -notmatch '^([^/]+)/([0-9]{1,2})$') { return $false }

    $ipPart = $Matches[1]
    $mask = [int]$Matches[2]

    if (-not (Test-ValidIpv4 -Value $ipPart)) { return $false }
    return $mask -ge 0 -and $mask -le 32
}

function Convert-Ipv4ToUInt32 {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Value)

    $bytes = [System.Net.IPAddress]::Parse($Value).GetAddressBytes()
    [Array]::Reverse($bytes)
    return [BitConverter]::ToUInt32($bytes, 0)
}

function Test-ValidIpv4Range {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Value)

    $candidate = ($Value -replace '\s*-\s*', '-').Trim()

    if ($candidate -notmatch '^([0-9]{1,3}(?:\.[0-9]{1,3}){3})-([0-9]{1,3}(?:\.[0-9]{1,3}){3})$') {
        return $false
    }

    $startIp = $Matches[1]
    $endIp = $Matches[2]

    if (-not (Test-ValidIpv4 -Value $startIp)) { return $false }
    if (-not (Test-ValidIpv4 -Value $endIp)) { return $false }

    return (Convert-Ipv4ToUInt32 -Value $startIp) -le (Convert-Ipv4ToUInt32 -Value $endIp)
}

function Get-Ipv4Span {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Value)

    $candidate = ($Value -replace '\s*-\s*', '-').Trim()

    if (Test-ValidIpv4 -Value $candidate) {
        $n = Convert-Ipv4ToUInt32 -Value $candidate
        return [pscustomobject]@{ Start = $n; End = $n }
    }

    if (Test-ValidCidr -Value $candidate) {
        $null = $candidate -match '^([^/]+)/([0-9]{1,2})$'
        $ipPart = $Matches[1]
        $maskBits = [int]$Matches[2]

        $binary = [Convert]::ToString((Convert-Ipv4ToUInt32 -Value $ipPart), 2).PadLeft(32, '0')
        $prefix = if ($maskBits -eq 0) { '' } else { $binary.Substring(0, $maskBits) }

        $start = [Convert]::ToUInt32($prefix.PadRight(32, '0'), 2)
        $end = [Convert]::ToUInt32($prefix.PadRight(32, '1'), 2)

        return [pscustomobject]@{ Start = $start; End = $end }
    }

    if (Test-ValidIpv4Range -Value $candidate) {
        $null = $candidate -match '^([0-9]{1,3}(?:\.[0-9]{1,3}){3})-([0-9]{1,3}(?:\.[0-9]{1,3}){3})$'
        return [pscustomobject]@{
            Start = Convert-Ipv4ToUInt32 -Value $Matches[1]
            End   = Convert-Ipv4ToUInt32 -Value $Matches[2]
        }
    }

    throw "Cannot create IPv4 span from '$Value'."
}

function Test-ValidServiceTag {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Value)

    $allowed = @(
        'VirtualNetwork',
        'AzureLoadBalancer',
        'Internet'
    )

    return $Value -in $allowed
}

function Test-ValidIpOrCidrOrRangeOrWildcardOrServiceTag {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Value)

    if ($Value -eq '*') { return $true }
    if (Test-ValidIpv4 -Value $Value) { return $true }
    if (Test-ValidCidr -Value $Value) { return $true }
    if (Test-ValidIpv4Range -Value $Value) { return $true }
    if (Test-ValidServiceTag -Value $Value) { return $true }

    return $false
}

function Normalize-AddressPrefixes {
    [CmdletBinding()]
    param([AllowNull()][object]$Value)

    $tokens = Split-NormalizedList -Value $Value -KeepAnyAsWildcard

    $normalized = foreach ($token in $tokens) {
        $candidate = ($token -replace '\s*-\s*', '-').Trim()

        if (-not (Test-ValidIpOrCidrOrRangeOrWildcardOrServiceTag -Value $candidate)) {
            throw "Invalid address token '$token'. Allowed values: explicit '*', IPv4, CIDR, IPv4 range, or supported service tags."
        }

        $candidate
    }

    if ($normalized -contains '*') { return @('*') }
    return @($normalized | Sort-Object -Unique)
}

function Normalize-Protocol {
    [CmdletBinding()]
    param([AllowNull()][object]$Value)

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        throw "Destination Protocol is required. Use '*' explicitly if you mean Any."
    }

    switch (($Value.ToString().Trim()).ToUpperInvariant()) {
        'TCP'  { return 'Tcp' }
        'UDP'  { return 'Udp' }
        'ICMP' { return 'Icmp' }
        'ESP'  { return 'Esp' }
        'AH'   { return 'Ah' }
        '*'    { return '*' }
        'ANY'  { return '*' }
        default { throw "Invalid protocol '$Value'. Allowed values: TCP, UDP, ICMP, ESP, AH, *" }
    }
}

function Normalize-Ports {
    [CmdletBinding()]
    param([AllowNull()][object]$Value)

    $tokens = @(Split-NormalizedList -Value $Value -KeepAnyAsWildcard)

    $ranges = @()

    foreach ($token in $tokens) {
        if ($token -eq '*') {
            return @('*')
        }

        $candidate = ($token -replace '\s*-\s*', '-').Trim()

        if ($candidate -match '^(\d+)-(\d+)$') {
            $start = [int]$Matches[1]
            $end = [int]$Matches[2]
        }
        elseif ($candidate -match '^\d+$') {
            $start = [int]$candidate
            $end = [int]$candidate
        }
        else {
            throw "Invalid port token '$token'."
        }

        if ($start -lt 0 -or $start -gt 65535 -or $end -lt 0 -or $end -gt 65535) {
            throw "Invalid port '$token'. Allowed range is 0-65535."
        }

        if ($start -gt $end) {
            throw "Invalid port range '$token'."
        }

        $ranges += [pscustomobject]@{
            Start = $start
            End   = $end
            Raw   = $token
        }
    }

    # IMPORTANT: ensure array
    $ranges = @($ranges | Sort-Object Start)

    # Detect overlaps
    for ($i = 1; $i -lt $ranges.Count; $i++) {
        $prev = $ranges[$i - 1]
        $curr = $ranges[$i]

        if ($curr.Start -le $prev.End) {
            throw "Invalid port definition in workbook. Overlap detected: '$($curr.Raw)' overlaps with '$($prev.Raw)'."
        }
    }

    # Return normalized
    $result = foreach ($r in $ranges) {
        if ($r.Start -eq $r.End) {
            [string]$r.Start
        }
        else {
            "$($r.Start)-$($r.End)"
        }
    }

    return @($result)
}

function Normalize-RuleName {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Value)

    $name = $Value.Trim()

    if ([string]::IsNullOrWhiteSpace($name)) {
        throw 'Rule name cannot be empty.'
    }

    if ($name.Length -gt 80) {
        throw "Rule name '$name' is longer than 80 characters."
    }

    if ($name -notmatch '^[A-Za-z0-9](?:[A-Za-z0-9._-]{0,78}[A-Za-z0-9_])?$') {
        throw "Invalid rule name '$name'. Allowed: 1-80 chars, alphanumerics, underscore, period, hyphen; must start with alphanumeric and end with alphanumeric or underscore."
    }

    return $name
}

function New-ManagedRuleDescription {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$HumanDescription
    )

    $description = $HumanDescription.Trim()

    if ([string]::IsNullOrWhiteSpace($description)) {
        throw "Introduce Description as: RITMxxxx - NYxxxx - Date(mm/dd/yyyy) - Action(Create/Update)"
    }

    $pattern = '^RITM\d+\s-\sNY\d+\s-\s\d{2}/\d{2}/\d{4}\s-\s(Create|Update)$'
    if ($description -notmatch $pattern) {
        throw "Description format invalid. Expected: RITMxxxx - NYxxxx - mm/dd/yyyy - Create|Update"
    }

    if ($description.Length -gt 140) {
        throw "Description exceeds Azure 140 character limit. Reduce -Description length."
    }

    return $description
}

function New-RuleFingerprint {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][pscustomobject]$Rule)

    $payload = [ordered]@{
        Direction                  = $Rule.Direction
        Access                     = $Rule.Access
        Protocol                   = $Rule.Protocol
        Priority                   = $Rule.Priority
        SourcePortRanges           = @($Rule.SourcePortRanges | Sort-Object)
        SourceAddressPrefixes      = @($Rule.SourceAddressPrefixes | Sort-Object)
        DestinationAddressPrefixes = @($Rule.DestinationAddressPrefixes | Sort-Object)
        DestinationPortRanges      = @($Rule.DestinationPortRanges | Sort-Object)
    }

    return ($payload | ConvertTo-Json -Depth 10 -Compress)
}

function Get-RowPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Row,
        [Parameter(Mandatory = $true)][string[]]$CandidateNames
    )

    foreach ($name in $CandidateNames) {
        if ($Row.PSObject.Properties.Name -contains $name) {
            return $Row.$name
        }
    }

    return $null
}

function Get-RequiredRowValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Row,
        [Parameter(Mandatory = $true)][string[]]$CandidateNames,
        [Parameter(Mandatory = $true)][string]$FieldLabel
    )

    $value = Get-RowPropertyValue -Row $Row -CandidateNames $CandidateNames
    if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) {
        throw "$FieldLabel is required. Use '*' explicitly if you mean Any."
    }

    return $value
}

function Test-IsBlankWorksheetRow {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][pscustomobject]$Row)

    $headers = Get-WorkbookHeaderMap
    foreach ($header in $headers.Values) {
        $value = Get-RowPropertyValue -Row $Row -CandidateNames @($header)
        if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
            return $false
        }
    }

    return $true
}

function Convert-WorksheetRowToRule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Row,
        [Parameter(Mandatory = $true)][string]$SheetName,
        [Parameter(Mandatory = $true)][int]$RowNumber,
        [Parameter(Mandatory = $true)][string]$Direction,
        [Parameter(Mandatory = $true)][string]$HumanDescription
    )

    $headers = Get-WorkbookHeaderMap

    $priorityRaw = Get-RequiredRowValue -Row $Row -CandidateNames @($headers.RuleNumber) -FieldLabel $headers.RuleNumber
    $ruleNameRaw = Get-RequiredRowValue -Row $Row -CandidateNames @($headers.RuleName) -FieldLabel $headers.RuleName
    $protocolRaw = Get-RequiredRowValue -Row $Row -CandidateNames @($headers.Protocol) -FieldLabel $headers.Protocol
    $sourceIpRaw = Get-RequiredRowValue -Row $Row -CandidateNames @($headers.SourceAddress) -FieldLabel $headers.SourceAddress
    $destinationIpRaw = Get-RequiredRowValue -Row $Row -CandidateNames @($headers.DestinationAddress) -FieldLabel $headers.DestinationAddress
    $destinationPortRaw = Get-RequiredRowValue -Row $Row -CandidateNames @($headers.DestinationPorts) -FieldLabel $headers.DestinationPorts

    $priority = 0
    if (-not [int]::TryParse([string]$priorityRaw, [ref]$priority)) {
        throw "Validation error in sheet '$SheetName', row $RowNumber, column '$($headers.RuleNumber)'. Value: '$priorityRaw'. Rule Number must be an integer."
    }

    if ($priority -lt 100 -or $priority -gt 4096) {
        throw "Validation error in sheet '$SheetName', row $RowNumber, column '$($headers.RuleNumber)'. Value: '$priorityRaw'. Priority $priority is outside Azure NSG allowed range 100-4096."
    }

    try {
        $ruleName = Normalize-RuleName -Value ([string]$ruleNameRaw)
    }
    catch {
        throw "Validation error in sheet '$SheetName', row $RowNumber, column '$($headers.RuleName)'. Value: '$ruleNameRaw'. $($_.Exception.Message)"
    }

    try {
        $description = New-ManagedRuleDescription -HumanDescription $HumanDescription
    }
    catch {
        throw "Validation error for script parameter '-Description'. Value: '$HumanDescription'. $($_.Exception.Message)"
    }

    try {
        $protocol = Normalize-Protocol -Value $protocolRaw
    }
    catch {
        throw "Validation error in sheet '$SheetName', row $RowNumber, rule '$ruleName', column '$($headers.Protocol)'. Value: '$protocolRaw'. $($_.Exception.Message)"
    }

    try {
        $sourcePrefixes = Normalize-AddressPrefixes -Value $sourceIpRaw
    }
    catch {
        throw "Validation error in sheet '$SheetName', row $RowNumber, rule '$ruleName', column '$($headers.SourceAddress)'. Value: '$sourceIpRaw'. $($_.Exception.Message)"
    }

    try {
        $destinationPrefixes = Normalize-AddressPrefixes -Value $destinationIpRaw
    }
    catch {
        throw "Validation error in sheet '$SheetName', row $RowNumber, rule '$ruleName', column '$($headers.DestinationAddress)'. Value: '$destinationIpRaw'. $($_.Exception.Message)"
    }

    try {
        $destinationPorts = Normalize-Ports -Value $destinationPortRaw
    }
    catch {
        throw "Validation error in sheet '$SheetName', row $RowNumber, rule '$ruleName', column '$($headers.DestinationPorts)'. Value: '$destinationPortRaw'. $($_.Exception.Message)"
    }

    $rule = [pscustomobject]@{
        SourceSheet                 = $SheetName
        OriginalRow                 = $RowNumber
        Priority                    = $priority
        Name                        = $ruleName
        Direction                   = $Direction
        Access                      = 'Allow'
        Protocol                    = $protocol
        SourcePortRanges            = @('*')
        SourceAddressPrefixes       = $sourcePrefixes
        DestinationAddressPrefixes  = $destinationPrefixes
        DestinationPortRanges       = $destinationPorts
        Description                 = $description
        Fingerprint                 = $null
    }

    $rule.Fingerprint = New-RuleFingerprint -Rule $rule
    return $rule
}

function Import-DesiredRules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string[]]$SheetNames,
        [Parameter(Mandatory = $true)][string]$Direction,
        [Parameter(Mandatory = $true)][string]$HumanDescription
    )

    Assert-WorkbookSchema -Path $Path -SheetNames $SheetNames

    $allRules = New-Object System.Collections.Generic.List[object]

    foreach ($sheet in $SheetNames) {
        $rows = @(Import-Excel -Path $Path -WorksheetName $sheet)
        $rowIndex = 1

        foreach ($row in $rows) {
            $rowIndex++

            if (Test-IsBlankWorksheetRow -Row $row) {
                continue
            }

            $rule = Convert-WorksheetRowToRule -Row $row -SheetName $sheet -RowNumber $rowIndex -Direction $Direction -HumanDescription $HumanDescription
            $allRules.Add($rule)
        }
    }

    if ($allRules.Count -eq 0) {
        throw "No rules were found in workbook '$Path' for sheet(s): $($SheetNames -join ', ')"
    }

    return @($allRules | Sort-Object Direction, Priority, Name)
}

function Test-DuplicateProperty {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Items,
        [Parameter(Mandatory = $true)][string]$PropertyName
    )

    return @(
        $Items |
        Group-Object -Property $PropertyName |
        Where-Object { $_.Count -gt 1 }
    )
}

function Get-LiveNsgRules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ResourceGroupName,
        [Parameter(Mandatory = $true)][string]$NsgName,
        [int]$RetryCount = 3,
        [int]$RetryDelaySeconds = 5
    )

    $parsed = Invoke-AzCliJson -Arguments @(
        'network', 'nsg', 'rule', 'list',
        '--resource-group', $ResourceGroupName,
        '--nsg-name', $NsgName,
        '--output', 'json',
        '--only-show-errors'
    ) -RetryCount $RetryCount -RetryDelaySeconds $RetryDelaySeconds -OperationName 'az network nsg rule list'

    if ($null -eq $parsed) {
        return @()
    }

    $liveRules = @($parsed)

    $normalized = foreach ($rule in $liveRules) {
        if (-not $rule) { continue }

        $props = $rule.PSObject.Properties.Name

        $sourcePrefixes = @()
        if (($props -contains 'sourceAddressPrefixes') -and $rule.sourceAddressPrefixes) {
            $sourcePrefixes = @($rule.sourceAddressPrefixes)
        }
        elseif (($props -contains 'sourceAddressPrefix') -and $rule.sourceAddressPrefix) {
            $sourcePrefixes = @($rule.sourceAddressPrefix)
        }
        else {
            $sourcePrefixes = @('*')
        }

        $destinationPrefixes = @()
        if (($props -contains 'destinationAddressPrefixes') -and $rule.destinationAddressPrefixes) {
            $destinationPrefixes = @($rule.destinationAddressPrefixes)
        }
        elseif (($props -contains 'destinationAddressPrefix') -and $rule.destinationAddressPrefix) {
            $destinationPrefixes = @($rule.destinationAddressPrefix)
        }
        else {
            $destinationPrefixes = @('*')
        }

        $sourcePorts = @()
        if (($props -contains 'sourcePortRanges') -and $rule.sourcePortRanges) {
            $sourcePorts = @($rule.sourcePortRanges)
        }
        elseif (($props -contains 'sourcePortRange') -and $rule.sourcePortRange) {
            $sourcePorts = @($rule.sourcePortRange)
        }
        else {
            $sourcePorts = @('*')
        }

        $destinationPorts = @()
        if (($props -contains 'destinationPortRanges') -and $rule.destinationPortRanges) {
            $destinationPorts = @($rule.destinationPortRanges)
        }
        elseif (($props -contains 'destinationPortRange') -and $rule.destinationPortRange) {
            $destinationPorts = @($rule.destinationPortRange)
        }
        else {
            $destinationPorts = @('*')
        }

        $description = ''
        if (($props -contains 'description') -and $rule.description) {
            $description = [string]$rule.description
        }

        $liveRule = [pscustomobject]@{
            Name                       = if (($props -contains 'name') -and $rule.name) { [string]$rule.name } else { '' }
            Priority                   = if (($props -contains 'priority') -and $null -ne $rule.priority) { [int]$rule.priority } else { 0 }
            Direction                  = if (($props -contains 'direction') -and $rule.direction) { [string]$rule.direction } else { '' }
            Access                     = if (($props -contains 'access') -and $rule.access) { [string]$rule.access } else { '' }
            Protocol                   = if (($props -contains 'protocol') -and $rule.protocol) { [string]$rule.protocol } else { '' }
            SourcePortRanges           = @($sourcePorts | ForEach-Object { Convert-AnyToken $_ } | Sort-Object -Unique)
            SourceAddressPrefixes      = @($sourcePrefixes | ForEach-Object { Convert-AnyToken $_ } | Sort-Object -Unique)
            DestinationAddressPrefixes = @($destinationPrefixes | ForEach-Object { Convert-AnyToken $_ } | Sort-Object -Unique)
            DestinationPortRanges      = @($destinationPorts | ForEach-Object { Convert-AnyToken $_ } | Sort-Object -Unique)
            Description                = $description
            Fingerprint                = $null
        }

        $liveRule.Fingerprint = New-RuleFingerprint -Rule $liveRule
        $liveRule
    }

    return ,@($normalized | Sort-Object Direction, Priority, Name)
}

function Parse-AddressToken {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Token)

    $candidate = ($Token -replace '\s*-\s*', '-').Trim()

    if ($script:AddressTokenCache.ContainsKey($candidate)) {
        return $script:AddressTokenCache[$candidate]
    }

    if ($candidate -eq '*') {
        $parsed = [pscustomobject]@{
            Kind  = 'Any'
            Raw   = $candidate
            Start = [uint32]0
            End   = [uint32]::MaxValue
        }
    }
    elseif (Test-ValidServiceTag -Value $candidate) {
        $parsed = [pscustomobject]@{
            Kind  = 'ServiceTag'
            Raw   = $candidate
            Start = $null
            End   = $null
        }
    }
    else {
        $span = Get-Ipv4Span -Value $candidate

        $kind = if (Test-ValidIpv4 -Value $candidate) {
            'Ip'
        }
        elseif (Test-ValidCidr -Value $candidate) {
            'Cidr'
        }
        elseif (Test-ValidIpv4Range -Value $candidate) {
            'Range'
        }
        else {
            throw "Unsupported address token '$candidate' for comparison."
        }

        $parsed = [pscustomobject]@{
            Kind  = $kind
            Raw   = $candidate
            Start = $span.Start
            End   = $span.End
        }
    }

    $script:AddressTokenCache[$candidate] = $parsed
    return $parsed
}

function Test-AddressTokenCovers {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Previous,
        [Parameter(Mandatory = $true)][pscustomobject]$Current
    )

    if ($Previous.Kind -eq 'Any') {
        return $true
    }

    if ($Previous.Kind -eq 'ServiceTag' -or $Current.Kind -eq 'ServiceTag') {
        return ($Previous.Kind -eq 'ServiceTag' -and $Current.Kind -eq 'ServiceTag' -and $Previous.Raw -ieq $Current.Raw)
    }

    return ($Previous.Start -le $Current.Start -and $Previous.End -ge $Current.End)
}

function Test-AddressSetCovered {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string[]]$PreviousTokens,
        [Parameter(Mandatory = $true)][string[]]$CurrentTokens
    )

    foreach ($currentToken in $CurrentTokens) {
        $currentParsed = Parse-AddressToken -Token $currentToken
        $covered = $false

        foreach ($previousToken in $PreviousTokens) {
            $previousParsed = Parse-AddressToken -Token $previousToken
            if (Test-AddressTokenCovers -Previous $previousParsed -Current $currentParsed) {
                $covered = $true
                break
            }
        }

        if (-not $covered) {
            return $false
        }
    }

    return $true
}

function Parse-PortToken {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Token)

    $candidate = ($Token -replace '\s*-\s*', '-').Trim()

    if ($script:PortTokenCache.ContainsKey($candidate)) {
        return $script:PortTokenCache[$candidate]
    }

    if ($candidate -eq '*') {
        $parsed = [pscustomobject]@{
            Kind  = 'Any'
            Raw   = $candidate
            Start = 0
            End   = 65535
        }
    }
    elseif ($candidate -match '^(\d+)-(\d+)$') {
        $parsed = [pscustomobject]@{
            Kind  = 'Range'
            Raw   = $candidate
            Start = [int]$Matches[1]
            End   = [int]$Matches[2]
        }
    }
    elseif ($candidate -match '^\d+$') {
        $parsed = [pscustomobject]@{
            Kind  = 'Port'
            Raw   = $candidate
            Start = [int]$candidate
            End   = [int]$candidate
        }
    }
    else {
        throw "Unsupported port token '$candidate' for comparison."
    }

    $script:PortTokenCache[$candidate] = $parsed
    return $parsed
}

function Test-PortTokenCovers {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Previous,
        [Parameter(Mandatory = $true)][pscustomobject]$Current
    )

    if ($Previous.Kind -eq 'Any') {
        return $true
    }

    return ($Previous.Start -le $Current.Start -and $Previous.End -ge $Current.End)
}

function Test-PortSetCovered {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string[]]$PreviousPorts,
        [Parameter(Mandatory = $true)][string[]]$CurrentPorts
    )

    foreach ($currentToken in $CurrentPorts) {
        $currentParsed = Parse-PortToken -Token $currentToken
        $covered = $false

        foreach ($previousToken in $PreviousPorts) {
            $previousParsed = Parse-PortToken -Token $previousToken
            if (Test-PortTokenCovers -Previous $previousParsed -Current $currentParsed) {
                $covered = $true
                break
            }
        }

        if (-not $covered) {
            return $false
        }
    }

    return $true
}

function Test-ProtocolCovers {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$PreviousProtocol,
        [Parameter(Mandatory = $true)][string]$CurrentProtocol
    )

    return ($PreviousProtocol -eq '*' -or $PreviousProtocol -eq $CurrentProtocol)
}

function Get-OverlapFindings {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Rules)

    $findings = @()
    $ordered = @($Rules | Sort-Object Direction, Priority, Name)

    for ($i = 0; $i -lt $ordered.Count; $i++) {
        $current = $ordered[$i]

        for ($j = 0; $j -lt $i; $j++) {
            $previous = $ordered[$j]

            if ($previous.Direction -ne $current.Direction) { continue }
            if (-not (Test-ProtocolCovers -PreviousProtocol $previous.Protocol -CurrentProtocol $current.Protocol)) { continue }
            if (-not (Test-AddressSetCovered -PreviousTokens $previous.SourceAddressPrefixes -CurrentTokens $current.SourceAddressPrefixes)) { continue }
            if (-not (Test-AddressSetCovered -PreviousTokens $previous.DestinationAddressPrefixes -CurrentTokens $current.DestinationAddressPrefixes)) { continue }
            if (-not (Test-PortSetCovered -PreviousPorts $previous.SourcePortRanges -CurrentPorts $current.SourcePortRanges)) { continue }
            if (-not (Test-PortSetCovered -PreviousPorts $previous.DestinationPortRanges -CurrentPorts $current.DestinationPortRanges)) { continue }

            $findings += [pscustomobject]@{
                RuleName          = $current.Name
                RulePriority      = $current.Priority
                ShadowedByName    = $previous.Name
                ShadowedByPriority= $previous.Priority
                ShadowedByAccess  = $previous.Access
                Message           = "Rule '$($current.Name)' may be fully covered by higher-precedence rule '$($previous.Name)' (priority $($previous.Priority), access $($previous.Access)). First-match NSG processing can make '$($current.Name)' ineffective for matching traffic. Rule will be not applied."
            }
            break
        }
    }

    return @($findings)
}

function Get-OverlapWarnings {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Rules)

    return @(Get-OverlapFindings -Rules $Rules | ForEach-Object { $_.Message })
}

function Test-DesiredRules {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Rules)

    $errors = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]

    $duplicatePriorities = Test-DuplicateProperty -Items $Rules -PropertyName 'Priority'
    foreach ($group in $duplicatePriorities) {
        $details = ($group.Group | ForEach-Object { "$($_.Name) [$($_.SourceSheet):$($_.OriginalRow)]" }) -join ', '
        $errors.Add("Duplicate priority $($group.Name): $details")
    }

    $duplicateNames = Test-DuplicateProperty -Items $Rules -PropertyName 'Name'
    foreach ($group in $duplicateNames) {
        $details = ($group.Group | ForEach-Object { "$($_.Name) [$($_.SourceSheet):$($_.OriginalRow)]" }) -join ', '
        $errors.Add("Duplicate rule name '$($group.Name)': $details")
    }

    $overlapFindings = @(Get-OverlapFindings -Rules $Rules)
    foreach ($finding in $overlapFindings) {
        $warnings.Add($finding.Message)
    }

    return [pscustomobject]@{
        IsValid          = ($errors.Count -eq 0)
        Errors           = @($errors)
        Warnings         = @($warnings)
        OverlapFindings  = $overlapFindings
    }
}

function New-UpdatedDescription {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$HumanDescription
    )

    return New-ManagedRuleDescription -HumanDescription $HumanDescription
}

function New-ApplyPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$DesiredRules,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$LiveRules,
        [AllowEmptyCollection()][object[]]$OverlapFindings = @()
    )

    $plan = New-Object System.Collections.Generic.List[object]
    $liveByName = @{}
    $liveByDirectionPriority = @{}
    $overlapByRuleName = @{}

    foreach ($finding in $OverlapFindings) {
        if (-not $overlapByRuleName.ContainsKey($finding.RuleName)) {
            $overlapByRuleName[$finding.RuleName] = $finding
        }
    }

    foreach ($item in $LiveRules) {
        $liveByName[$item.Name] = $item

        $priorityKey = '{0}|{1}' -f $item.Direction, $item.Priority
        if (-not $liveByDirectionPriority.ContainsKey($priorityKey)) {
            $liveByDirectionPriority[$priorityKey] = $item
        }
    }

    foreach ($desired in $DesiredRules) {
        $priorityKey = '{0}|{1}' -f $desired.Direction, $desired.Priority

        if ($liveByName.ContainsKey($desired.Name)) {
            $live = $liveByName[$desired.Name]

            if ($desired.Fingerprint -eq $live.Fingerprint) {
                $plan.Add([pscustomobject]@{
                    Action   = 'NoChange'
                    Name     = $desired.Name
                    Priority = $desired.Priority
                    Desired  = $desired
                    Live     = $live
                    Reason   = 'Live rule already matches desired state.'
                })
                continue
            }

            $plan.Add([pscustomobject]@{
                Action   = 'Update'
                Name     = $desired.Name
                Priority = $desired.Priority
                Desired  = $desired
                Live     = $live
                Reason   = if ($overlapByRuleName.ContainsKey($desired.Name)) { $overlapByRuleName[$desired.Name].Message } else { 'Live rule with same name differs from desired state and will be updated in place.' }
            })

            if ($overlapByRuleName.ContainsKey($desired.Name)) {
                $plan[$plan.Count - 1].Action = 'SkipApply'
            }
            continue
        }

        if ($liveByDirectionPriority.ContainsKey($priorityKey)) {
            $priorityOwner = $liveByDirectionPriority[$priorityKey]

            $plan.Add([pscustomobject]@{
                Action   = 'Conflict'
                Name     = $desired.Name
                Priority = $desired.Priority
                Desired  = $desired
                Live     = $priorityOwner
                Reason   = "A rule already exists at direction '$($desired.Direction)' priority '$($desired.Priority)' with a different name ('$($priorityOwner.Name)'). This looks like a rename or a priority collision."
            })
            continue
        }

        $plan.Add([pscustomobject]@{
            Action   = if ($overlapByRuleName.ContainsKey($desired.Name)) { 'SkipApply' } else { 'Create' }
            Name     = $desired.Name
            Priority = $desired.Priority
            Desired  = $desired
            Live     = $null
            Reason   = if ($overlapByRuleName.ContainsKey($desired.Name)) { $overlapByRuleName[$desired.Name].Message } else { 'Rule does not exist in target NSG.' }
        })
    }

    return @($plan | Sort-Object Priority, Name, Action)
}

function Save-Checkpoint {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Plan,
        [Parameter(Mandatory = $true)][object]$Metadata,
        [string]$LatestPath = ''
    )

    $payload = [ordered]@{
        Metadata   = $Metadata
        Plan       = @($Plan)
        SavedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
    }

    $parent = Split-Path -Path $Path -Parent
    if ($parent -and -not (Test-Path -Path $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }

    $payload | ConvertTo-Json -Depth 20 | Set-Content -Path $Path -Encoding UTF8

    if (-not [string]::IsNullOrWhiteSpace($LatestPath) -and $LatestPath -ne $Path) {
        $latestParent = Split-Path -Path $LatestPath -Parent
        if ($latestParent -and -not (Test-Path -Path $latestParent)) {
            New-Item -ItemType Directory -Path $latestParent -Force | Out-Null
        }

        $payload | ConvertTo-Json -Depth 20 | Set-Content -Path $LatestPath -Encoding UTF8
    }
}

function Resolve-CheckpointTargets {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$CheckpointPath,
        [Parameter(Mandatory = $true)][string]$CheckpointMode,
        [switch]$Apply
    )

    $effectiveMode = $CheckpointMode
    if ($effectiveMode -eq 'Auto') {
        if ($Apply) {
            $effectiveMode = 'Both'
        }
        else {
            $effectiveMode = 'Overwrite'
        }
    }

    $parent = Split-Path -Path $CheckpointPath -Parent
    $fileName = Split-Path -Path $CheckpointPath -Leaf
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    $extension = [System.IO.Path]::GetExtension($fileName)

    $timestampedFileName = '{0}-{1}{2}' -f $baseName, (Get-Date -Format 'yyyyMMdd-HHmmss'), $extension
    $timestampedPath = if ([string]::IsNullOrWhiteSpace($parent)) {
        $timestampedFileName
    }
    else {
        Join-Path -Path $parent -ChildPath $timestampedFileName
    }

    switch ($effectiveMode) {
        'Overwrite' {
            return [pscustomobject]@{
                EffectiveMode = $effectiveMode
                RunPath       = $CheckpointPath
                LatestPath    = ''
            }
        }
        'Timestamped' {
            return [pscustomobject]@{
                EffectiveMode = $effectiveMode
                RunPath       = $timestampedPath
                LatestPath    = ''
            }
        }
        'Both' {
            return [pscustomobject]@{
                EffectiveMode = $effectiveMode
                RunPath       = $timestampedPath
                LatestPath    = $CheckpointPath
            }
        }
        default {
            throw "Unsupported checkpoint mode '$effectiveMode'."
        }
    }
}

function Invoke-NsgRuleCreate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Rule,
        [Parameter(Mandatory = $true)][string]$ResourceGroupName,
        [Parameter(Mandatory = $true)][string]$NsgName,
        [int]$RetryCount = 3,
        [int]$RetryDelaySeconds = 5
    )

    $arguments = @(
        'network', 'nsg', 'rule', 'create',
        '--resource-group', $ResourceGroupName,
        '--nsg-name', $NsgName,
        '--name', $Rule.Name,
        '--priority', [string]$Rule.Priority,
        '--direction', $Rule.Direction,
        '--access', $Rule.Access,
        '--protocol', $Rule.Protocol,
        '--description', $Rule.Description,
        '--source-address-prefixes'
    )

    $arguments += $Rule.SourceAddressPrefixes
    $arguments += @('--source-port-ranges')
    $arguments += $Rule.SourcePortRanges
    $arguments += @('--destination-address-prefixes')
    $arguments += $Rule.DestinationAddressPrefixes
    $arguments += @('--destination-port-ranges')
    $arguments += $Rule.DestinationPortRanges
    $arguments += @('--output', 'json', '--only-show-errors')

    return Invoke-AzCliJson -Arguments $arguments -RetryCount $RetryCount -RetryDelaySeconds $RetryDelaySeconds -OperationName "az network nsg rule create $($Rule.Name)"
}

function Invoke-NsgRuleUpdate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Rule,
        [Parameter(Mandatory = $true)][string]$ResourceGroupName,
        [Parameter(Mandatory = $true)][string]$NsgName,
        [int]$RetryCount = 3,
        [int]$RetryDelaySeconds = 5
    )

    $arguments = @(
        'network', 'nsg', 'rule', 'update',
        '--resource-group', $ResourceGroupName,
        '--nsg-name', $NsgName,
        '--name', $Rule.Name,
        '--priority', [string]$Rule.Priority,
        '--direction', $Rule.Direction,
        '--access', $Rule.Access,
        '--protocol', $Rule.Protocol,
        '--description', $Rule.Description,
        '--source-address-prefixes'
    )

    $arguments += $Rule.SourceAddressPrefixes
    $arguments += @('--source-port-ranges')
    $arguments += $Rule.SourcePortRanges
    $arguments += @('--destination-address-prefixes')
    $arguments += $Rule.DestinationAddressPrefixes
    $arguments += @('--destination-port-ranges')
    $arguments += $Rule.DestinationPortRanges
    $arguments += @('--output', 'json', '--only-show-errors')

    return Invoke-AzCliJson -Arguments $arguments -RetryCount $RetryCount -RetryDelaySeconds $RetryDelaySeconds -OperationName "az network nsg rule update $($Rule.Name)"
}


function Invoke-NsgRuleWaitCreated {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RuleName,
        [Parameter(Mandatory = $true)][string]$ResourceGroupName,
        [Parameter(Mandatory = $true)][string]$NsgName,
        [int]$RetryCount = 3,
        [int]$RetryDelaySeconds = 5
    )

    [void](Invoke-AzCommandText -Arguments @(
        'network', 'nsg', 'rule', 'wait',
        '--resource-group', $ResourceGroupName,
        '--nsg-name', $NsgName,
        '--name', $RuleName,
        '--created',
        '--only-show-errors'
    ) -RetryCount $RetryCount -RetryDelaySeconds $RetryDelaySeconds -OperationName "az network nsg rule wait created $RuleName")
}

function Invoke-NsgRuleWaitUpdated {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RuleName,
        [Parameter(Mandatory = $true)][string]$ResourceGroupName,
        [Parameter(Mandatory = $true)][string]$NsgName,
        [int]$RetryCount = 3,
        [int]$RetryDelaySeconds = 5
    )

    [void](Invoke-AzCommandText -Arguments @(
        'network', 'nsg', 'rule', 'wait',
        '--resource-group', $ResourceGroupName,
        '--nsg-name', $NsgName,
        '--name', $RuleName,
        '--updated',
        '--only-show-errors'
    ) -RetryCount $RetryCount -RetryDelaySeconds $RetryDelaySeconds -OperationName "az network nsg rule wait updated $RuleName")
}

function Invoke-UpdateRule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$DesiredRule,
        [Parameter(Mandatory = $true)][pscustomobject]$LiveRule,
        [Parameter(Mandatory = $true)][string]$ResourceGroupName,
        [Parameter(Mandatory = $true)][string]$NsgName,
        [Parameter(Mandatory = $true)][string]$HumanDescription,
        [int]$RetryCount = 3,
        [int]$RetryDelaySeconds = 5
    )

    $updatedDescription = New-UpdatedDescription -HumanDescription $HumanDescription

    $updatedRule = [pscustomobject]@{
        SourceSheet                 = $DesiredRule.SourceSheet
        OriginalRow                 = $DesiredRule.OriginalRow
        Priority                    = $DesiredRule.Priority
        Name                        = $DesiredRule.Name
        Direction                   = $DesiredRule.Direction
        Access                      = $DesiredRule.Access
        Protocol                    = $DesiredRule.Protocol
        SourcePortRanges            = $DesiredRule.SourcePortRanges
        SourceAddressPrefixes       = $DesiredRule.SourceAddressPrefixes
        DestinationAddressPrefixes  = $DesiredRule.DestinationAddressPrefixes
        DestinationPortRanges       = $DesiredRule.DestinationPortRanges
        Description                 = $updatedDescription
        Fingerprint                 = $DesiredRule.Fingerprint
    }

    [void](Invoke-NsgRuleUpdate -Rule $updatedRule -ResourceGroupName $ResourceGroupName -NsgName $NsgName -RetryCount $RetryCount -RetryDelaySeconds $RetryDelaySeconds)
    Invoke-NsgRuleWaitUpdated -RuleName $updatedRule.Name -ResourceGroupName $ResourceGroupName -NsgName $NsgName -RetryCount $RetryCount -RetryDelaySeconds $RetryDelaySeconds
}

function Invoke-CreateRule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$DesiredRule,
        [Parameter(Mandatory = $true)][string]$ResourceGroupName,
        [Parameter(Mandatory = $true)][string]$NsgName,
        [int]$RetryCount = 3,
        [int]$RetryDelaySeconds = 5
    )

    [void](Invoke-NsgRuleCreate -Rule $DesiredRule -ResourceGroupName $ResourceGroupName -NsgName $NsgName -RetryCount $RetryCount -RetryDelaySeconds $RetryDelaySeconds)
    Invoke-NsgRuleWaitCreated -RuleName $DesiredRule.Name -ResourceGroupName $ResourceGroupName -NsgName $NsgName -RetryCount $RetryCount -RetryDelaySeconds $RetryDelaySeconds
}

function Invoke-ApplyPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Plan,
        [Parameter(Mandatory = $true)][string]$ResourceGroupName,
        [Parameter(Mandatory = $true)][string]$NsgName,
        [Parameter(Mandatory = $true)][string]$CheckpointPath,
        [string]$LatestCheckpointPath = '',
        [Parameter(Mandatory = $true)][string]$HumanDescription,
        [int]$RetryCount = 3,
        [int]$RetryDelaySeconds = 5,
        [switch]$Apply
    )

    $execution = foreach ($item in $Plan) {
        $beforeView = Convert-RuleToCheckpointView -Rule $item.Live
        $afterView  = Convert-RuleToCheckpointView -Rule $item.Desired

        $executionEntry = [ordered]@{
            Name           = $item.Name
            Priority       = $item.Priority
            PlannedAction  = $item.Action
            Status         = if ($item.Action -in @('NoChange', 'SkipApply')) { 'Skipped' } elseif ($item.Action -eq 'Conflict') { 'Blocked' } else { 'Pending' }
            Reason         = $item.Reason
            LastUpdatedUtc = (Get-Date).ToUniversalTime().ToString('o')
        }

        switch ($item.Action) {
            'Update' {
                $executionEntry.Before = $beforeView
                $executionEntry.After = $afterView
                $executionEntry.Changes = @(Compare-RuleForCheckpoint -Before $beforeView -After $afterView)
            }
            'Create' {
                $executionEntry.After = $afterView
                $executionEntry.Changes = @(Compare-RuleForCheckpoint -Before $null -After $afterView)
            }
            'SkipApply' {
                if ($null -ne $beforeView) {
                    $executionEntry.Before = $beforeView
                }
                $executionEntry.After = $afterView
                $executionEntry.Changes = @(Compare-RuleForCheckpoint -Before $beforeView -After $afterView)
            }
            'Conflict' {
                $executionEntry.Before = $beforeView
                $executionEntry.After = $afterView
                $executionEntry.Changes = @(Compare-RuleForCheckpoint -Before $beforeView -After $afterView)
            }
        }

        [pscustomobject]$executionEntry
    }

    $executionIndex = @{}
    foreach ($entry in $execution) {
        $key = '{0}|{1}|{2}' -f $entry.Name, $entry.Priority, $entry.PlannedAction
        $executionIndex[$key] = $entry
    }

    Save-Checkpoint -Path $CheckpointPath -LatestPath $LatestCheckpointPath -Plan $execution -Metadata ([ordered]@{
        ResourceGroupName = $ResourceGroupName
        NsgName           = $NsgName
        Mode              = if ($Apply) { 'Apply' } else { 'DryRun' }
        CheckpointPath    = $CheckpointPath
        LatestPath        = $LatestCheckpointPath
    })

    foreach ($item in $Plan) {
        $key = '{0}|{1}|{2}' -f $item.Name, $item.Priority, $item.Action
        $match = $executionIndex[$key]

        if ($item.Action -eq 'NoChange') {
            Write-Log -Message "NOCHANGE $($item.Name) priority=$($item.Priority)" -Level INFO -Color Gray
            continue
        }

        if ($item.Action -eq 'Conflict') {
            Write-Log -Message "CONFLICT $($item.Name) priority=$($item.Priority) : $($item.Reason)" -Level ERROR
            continue
        }

        if ($item.Action -eq 'SkipApply') {
            Write-Log -Message "SKIP_APPLY $($item.Name) priority=$($item.Priority)" -Level INFO -Color DarkYellow
            continue
        }

        $actionText = $item.Action.ToUpperInvariant()

        switch ($item.Action) {
            'Create' { $color = 'Green' }
            'Update' { $color = 'Yellow' }
            default  { $color = 'Gray' }
        }

        Write-Log -Message "$actionText $($item.Name) priority=$($item.Priority)" -Level INFO -Color $color

        if (-not $Apply) {
            continue
        }

        $target = "$ResourceGroupName/$NsgName/$($item.Name)"
        if (-not $PSCmdlet.ShouldProcess($target, $item.Action)) {
            Write-Log -Message "SKIPPED by ShouldProcess: $($item.Action) $target" -Level INFO
            continue
        }

        try {
            switch ($item.Action) {
                'Create' {
                    Invoke-CreateRule -DesiredRule $item.Desired -ResourceGroupName $ResourceGroupName -NsgName $NsgName -RetryCount $RetryCount -RetryDelaySeconds $RetryDelaySeconds
                }
                'Update' {
                    Invoke-UpdateRule -DesiredRule $item.Desired -LiveRule $item.Live -ResourceGroupName $ResourceGroupName -NsgName $NsgName -HumanDescription $HumanDescription -RetryCount $RetryCount -RetryDelaySeconds $RetryDelaySeconds
                }
                default {
                    throw "Unsupported planned action '$($item.Action)'"
                }
            }

            if ($null -ne $match) {
                $match.Status = 'Completed'
                $match.LastUpdatedUtc = (Get-Date).ToUniversalTime().ToString('o')
                Save-Checkpoint -Path $CheckpointPath -LatestPath $LatestCheckpointPath -Plan $execution -Metadata ([ordered]@{
                    ResourceGroupName = $ResourceGroupName
                    NsgName           = $NsgName
                    Mode              = 'Apply'
                    CheckpointPath    = $CheckpointPath
                    LatestPath        = $LatestCheckpointPath
                })
            }
        }
        catch {
            if ($null -ne $match) {
                $match.Status = 'Failed'
                $match.LastUpdatedUtc = (Get-Date).ToUniversalTime().ToString('o')
                Save-Checkpoint -Path $CheckpointPath -LatestPath $LatestCheckpointPath -Plan $execution -Metadata ([ordered]@{
                    ResourceGroupName = $ResourceGroupName
                    NsgName           = $NsgName
                    Mode              = 'Apply'
                    CheckpointPath    = $CheckpointPath
                    LatestPath        = $LatestCheckpointPath
                })
            }

            throw
        }
    }

    return @($execution)
}

function Show-PlanSummary {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Plan)

    $grouped = $Plan | Group-Object Action | Sort-Object Name
    Write-Host ''
    Write-Host 'PLAN SUMMARY'
    Write-Host '------------'
    foreach ($group in $grouped) {
        '{0,-10} : {1}' -f $group.Name, $group.Count | Write-Host
    }
    Write-Host ''
}

function Show-ValidationResult {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][pscustomobject]$Validation)

    if ($Validation.IsValid) {
        Write-Log -Message 'Validation PASSED.' -Level INFO
    }
    else {
        Write-Log -Message 'Validation FAILED.' -Level ERROR
    }

    foreach ($warning in $Validation.Warnings) {
        Write-Log -Message $warning -Level WARN
    }

    foreach ($erroro in $Validation.Errors) {
        Write-Log -Message $erroro -Level ERROR
    }
}

function Invoke-Preflight {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ResourceGroupName,
        [Parameter(Mandatory = $true)][string]$NsgName,
        [int]$RetryCount = 3,
        [int]$RetryDelaySeconds = 5
    )

    Assert-Dependency -CommandName 'az' -InstallHint "Install Azure CLI and ensure 'az' is on PATH."
    Assert-Module -ModuleName 'ImportExcel' -InstallHint "Install it with: Install-Module ImportExcel -Scope CurrentUser"

    [void](Invoke-AzCommandText -Arguments @(
        'account', 'show',
        '--output', 'none',
        '--only-show-errors'
    ) -RetryCount $RetryCount -RetryDelaySeconds $RetryDelaySeconds -OperationName 'az account show')

    [void](Invoke-AzCommandText -Arguments @(
        'network', 'nsg', 'show',
        '--resource-group', $ResourceGroupName,
        '--name', $NsgName,
        '--output', 'none',
        '--only-show-errors'
    ) -RetryCount $RetryCount -RetryDelaySeconds $RetryDelaySeconds -OperationName 'az network nsg show')
}

if (-not (Test-Path -Path $WorkbookPath -PathType Leaf)) {
    throw "Workbook not found: $WorkbookPath"
}

if (-not $SkipPreflight) {
    Invoke-Preflight -ResourceGroupName $ResourceGroupName -NsgName $NsgName -RetryCount $RetryCount -RetryDelaySeconds $RetryDelaySeconds
}

Write-Log -Message "Reading workbook '$WorkbookPath'" -Level INFO
$desiredRules = Import-DesiredRules -Path $WorkbookPath -SheetNames $SheetNames -Direction $Direction -HumanDescription $Description

$validation = Test-DesiredRules -Rules $desiredRules
Show-ValidationResult -Validation $validation

if (-not $validation.IsValid) {
    throw 'Input workbook validation failed. No Azure changes were made.'
}

if ($FailOnShadowing -and $validation.Warnings.Count -gt 0) {
    throw 'Overlap warnings detected and FailOnShadowing was specified. No Azure changes were made.'
}

Write-Log -Message "Loading current NSG rules from '$NsgName'" -Level INFO
$liveRules = Get-LiveNsgRules -ResourceGroupName $ResourceGroupName -NsgName $NsgName -RetryCount $RetryCount -RetryDelaySeconds $RetryDelaySeconds

$plan = New-ApplyPlan -DesiredRules $desiredRules -LiveRules $liveRules -OverlapFindings $validation.OverlapFindings
Show-PlanSummary -Plan $plan

$conflicts = @($plan | Where-Object { $_.Action -eq 'Conflict' })
if ($conflicts.Count -gt 0) {
    Write-Log -Message "One or more conflicts were detected. These rules will not be modified because they look like a rename or a priority collision." -Level WARN
}

$checkpointTargets = Resolve-CheckpointTargets -CheckpointPath $CheckpointPath -CheckpointMode $CheckpointMode -Apply:$Apply
Write-Log -Message "Checkpoint mode '$($checkpointTargets.EffectiveMode)' using run file '$($checkpointTargets.RunPath)'" -Level INFO
if (-not [string]::IsNullOrWhiteSpace($checkpointTargets.LatestPath)) {
    Write-Log -Message "Checkpoint latest file '$($checkpointTargets.LatestPath)' will also be updated." -Level INFO
}

$execution = Invoke-ApplyPlan -Plan $plan -ResourceGroupName $ResourceGroupName -NsgName $NsgName -CheckpointPath $checkpointTargets.RunPath -LatestCheckpointPath $checkpointTargets.LatestPath -HumanDescription $Description -RetryCount $RetryCount -RetryDelaySeconds $RetryDelaySeconds -Apply:$Apply

if ($PassThru) {
    $execution
}