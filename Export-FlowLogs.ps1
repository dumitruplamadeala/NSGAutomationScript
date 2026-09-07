<#
.SYNOPSIS
Exports inbound NSG flow-log traffic from Azure Log Analytics to CSV.

.DESCRIPTION
Executes a parameterized KQL query against an Azure Log Analytics workspace
using configurable time-based intervals. The script:

- Queries the workspace in manageable time chunks.
- Supports filtering by flow status, such as Denied or Allowed.
- Saves the generated query interval and execution logs.
- Stores each successful query result as an individual chunk CSV file.
- Re-aggregates all chunk results into a single server-specific CSV file.
- Prevents stale data from previous executions by using a timestamped output
  directory for each run.

The script requires the Azure CLI and the preview Azure CLI
"log-analytics" extension.

Required environment:

- Azure CLI version 2.90.0
- Azure CLI extension "log-analytics" version 1.0.0b1
- An authenticated Azure CLI session
- Log Analytics query permissions for the target workspace

Required installation:

    az extension add `
        --name log-analytics `
        --version 1.0.0b1 `
        --allow-preview true `
        --yes

.EXAMPLE
.\Export-FlowLogs.ps1 `
    -WorkspaceId "d099361a-dd21-4230-be80-fe8286679483" `
    -ServerName "mistroia" `
    -QueryTemplatePath ".\queries\flowlogs-inbound-query.kql" `
    -OutputDirectory ".\flowlogs" `
    -HistoryDays 180 `
    -ChunkDays 20 `
    -FlowStatus "Denied"

Exports denied inbound NSG flow-log traffic for "mistroia" over the previous
180 days. Log Analytics is queried in 20-day intervals, and the results are
aggregated into a final server-specific CSV file.

.NOTES
The Azure CLI extension is currently preview-only. It must be installed in
every environment where the script is executed, including Synopsys agents or
other CI/CD execution environments.

The script does not silently discard records. If the final aggregated result
exceeds the configured generator limit, the execution fails and requires the
input or generator configuration to be reviewed.

The Azure CLI must be authenticated before running the script, for example:

    az login
    az account set --subscription "<SUBSCRIPTION_ID>"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$WorkspaceId,
    [Parameter(Mandatory)] [string]$ServerName,
    [Parameter(Mandatory)] [string]$QueryTemplatePath,
    [Parameter(Mandatory)] [string]$OutputDirectory,
    [int]$HistoryDays = 180,
    [int]$ChunkDays = 14,
    [ValidateSet("Allowed", "Denied")]
    [string]$FlowStatus = "Denied",
    [int]$SafetyRowLimit = 450000,
    [int]$GeneratorRowLimit = 50000
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$scriptStart = Get-Date

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message" -ForegroundColor Cyan
}

function Assert-CommandExists {
    param([string]$CommandName)
    if (-not (Get-Command $CommandName -ErrorAction SilentlyContinue)) {
        throw "Required command was not found: $CommandName"
    }
}

function Convert-ToKqlDateTime {
    param([datetime]$DateTime)
    $DateTime.ToUniversalTime().ToString(
        "yyyy-MM-ddTHH:mm:ss.fffZ",
        [Globalization.CultureInfo]::InvariantCulture
    )
}

function Invoke-LogAnalyticsQuery {
    param(
        [string]$WorkspaceId,
        [string]$Query,
        [string]$LogFilePath
    )

    # Keep stdout and stderr separate. Only stdout is parsed as JSON.
    $tempDirectory = Split-Path -Parent $LogFilePath
    $stdoutPath = Join-Path $tempDirectory ("{0}.stdout" -f [guid]::NewGuid())
    $stderrPath = Join-Path $tempDirectory ("{0}.stderr" -f [guid]::NewGuid())

    try {
        # az.cmd on Windows can misparse embedded double quotes in a dynamic
        # argument. KQL accepts single-quoted string literals, so normalize
        # the query before passing it to Azure CLI.
        $queryArgument = $Query -replace '"', "'"
        $queryArgument = $queryArgument -replace "[\r\n]+", " "

        & az monitor log-analytics query `
            --workspace $WorkspaceId `
            --analytics-query $queryArgument `
            --output json `
            1> $stdoutPath `
            2> $stderrPath

        $exitCode = $LASTEXITCODE
        $jsonText = if (Test-Path -LiteralPath $stdoutPath) {
            Get-Content -LiteralPath $stdoutPath -Raw
        }
        else { "" }

        $stderrText = if (Test-Path -LiteralPath $stderrPath) {
            Get-Content -LiteralPath $stderrPath -Raw
        }
        else { "" }

        @(
            "=== STDOUT ==="
            $jsonText
            ""
            "=== STDERR ==="
            $stderrText
        ) | Out-File -LiteralPath $LogFilePath -Encoding utf8

        if ($exitCode -ne 0) {
            throw "Azure CLI query failed with exit code $exitCode. See: $LogFilePath"
        }

        if ([string]::IsNullOrWhiteSpace($jsonText)) {
            throw "Azure CLI returned an empty response. See: $LogFilePath"
        }

        try {
            return ($jsonText | ConvertFrom-Json)
        }
        catch {
            throw "Azure CLI returned invalid JSON. See: $LogFilePath"
        }
    }
    finally {
        Remove-Item -LiteralPath $stdoutPath, $stderrPath -Force -ErrorAction SilentlyContinue
    }
}

function Convert-LogAnalyticsTableToObjects {
    param($Response)

    if ($null -eq $Response) {
        return @()
    }

    $responseItems = @($Response)

    if ($responseItems.Count -eq 0) {
        return @()
    }

    $Response = $responseItems[0]

    if ($Response -is [psobject]) {
        $firstItemProperties = @($Response.PSObject.Properties.Name)

        if ($firstItemProperties -contains "SrcEff") {
            return @($responseItems | ForEach-Object {
                $_.PSObject.Properties.Remove("TableName")
                $_
            })
        }
    }

    if ($Response -isnot [psobject]) {
        return @()
    }

    $properties = @($Response.PSObject.Properties.Name)

    if (($properties -contains "error") -and $Response.error) {
        throw "Log Analytics returned an error: $($Response.error | ConvertTo-Json -Compress)"
    }

    if (($properties -notcontains "tables") -or -not $Response.tables -or $Response.tables.Count -eq 0) {
        return @()
    }

    $table = $Response.tables[0]
    $tableProperties = @($table.PSObject.Properties.Name)

    if (($tableProperties -notcontains "columns") -or -not $table.columns) {
        return @()
    }

    $columnNames = @($table.columns | ForEach-Object {
        if ($_ -is [string]) { $_ } else { $_.name }
    })

    if (($tableProperties -notcontains "rows") -or -not $table.rows) {
        return @()
    }

    foreach ($row in $table.rows) {
        $object = [ordered]@{}
        for ($i = 0; $i -lt $columnNames.Count; $i++) {
            $object[$columnNames[$i]] = $row[$i]
        }
        [pscustomobject]$object
    }
}

function Assert-ExpectedColumns {
    param([object[]]$Rows)

    if ($Rows.Count -eq 0) { return }

    $required = @(
        "SrcEff", "DestEff", "DestPort", "L4Protocol",
        "Server", "SrcPorts", "Requests"
    )
    $actual = @($Rows[0].PSObject.Properties.Name)

    foreach ($column in $required) {
        if ($actual -notcontains $column) {
            throw "Required KQL result column '$column' is missing. Returned columns: $($actual -join ', ')"
        }
    }
}

function Export-ObjectsToCsv {
    param([object[]]$Objects, [string]$Path)
    if ($Objects.Count -gt 0) {
        $Objects | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
    }
}

function Get-FlowKey {
    param($Row)
    [ordered]@{
        SrcEff = [string]$Row.SrcEff
        DestEff = [string]$Row.DestEff
        DestPort = [string]$Row.DestPort
        L4Protocol = [string]$Row.L4Protocol
        Server = [string]$Row.Server
        SrcPorts = [string]$Row.SrcPorts
    } | ConvertTo-Json -Compress
}

function Merge-FlowRows {
    param([string[]]$ChunkFiles, [string]$FinalCsvPath)

    Write-Step "Re-aggregating flow-log chunks"
    $aggregated = @{}

    foreach ($chunkFile in $ChunkFiles) {
        Write-Host "Reading: $chunkFile"
        foreach ($row in (Import-Csv -LiteralPath $chunkFile)) {
            $key = Get-FlowKey $row

            if (-not $aggregated.ContainsKey($key)) {
                $aggregated[$key] = [ordered]@{
                    SrcEff = $row.SrcEff; DestEff = $row.DestEff
                    DestPort = $row.DestPort; L4Protocol = $row.L4Protocol
                    Server = $row.Server; SrcPorts = $row.SrcPorts
                    Requests = [long]0
                }
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$row.Requests)) {
                $aggregated[$key].Requests += [long]$row.Requests
            }
        }
    }

    $result = @($aggregated.Values |
        ForEach-Object { [pscustomobject]$_ } |
        Sort-Object -Property Requests -Descending)

    $result | Export-Csv -LiteralPath $FinalCsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Final aggregated rows: $($result.Count)"
    return $result.Count
}

# Validation
Write-Step "Validating configuration"
Assert-CommandExists "az"

if (-not (Test-Path -LiteralPath $QueryTemplatePath -PathType Leaf)) {
    throw "KQL query template was not found: $QueryTemplatePath"
}
if ($HistoryDays -le 0 -or $ChunkDays -le 0) {
    throw "HistoryDays and ChunkDays must be greater than zero."
}
if ($SafetyRowLimit -le 0 -or $GeneratorRowLimit -le 0) {
    throw "SafetyRowLimit and GeneratorRowLimit must be greater than zero."
}
if ($SafetyRowLimit -ge 500000) {
    throw "SafetyRowLimit must be below Azure's 500,000-row limit."
}

$queryTemplate = Get-Content -LiteralPath $QueryTemplatePath -Raw
foreach ($placeholder in @("{{START_UTC}}", "{{END_UTC}}", "{{SERVER_FILTER}}", "{{FLOW_STATUS}}")) {
    if (-not $queryTemplate.Contains($placeholder)) {
        throw "KQL template must contain the placeholder $placeholder."
    }
}
if ($queryTemplate -match "(?i)\|\s*take\s+\d+") {
    throw "Remove the '| take ...' operator from the KQL template. Apply any final limit only after aggregation."
}

# Every execution gets an isolated directory, preventing stale chunks from being reused.
$runDirectory = Join-Path $OutputDirectory (Get-Date -Format "yyyyMMdd_HHmmss")
$chunkDirectory = Join-Path $runDirectory "chunks"
$logDirectory = Join-Path $runDirectory "logs"
New-Item -ItemType Directory -Path $chunkDirectory, $logDirectory -Force | Out-Null

# q41-mtax -> Q41. Change this to an explicit parameter if naming differs.
$serverFilter = $ServerName.Split("-")[0].ToUpperInvariant()
$exportEnd = [datetime]::UtcNow
$exportStart = $exportEnd.AddDays(-$HistoryDays)
$chunkFiles = New-Object System.Collections.Generic.List[string]
$currentStart = $exportStart
$chunkNumber = 0

Write-Step "Starting chunked Log Analytics export"
Write-Host "Server: $ServerName"
Write-Host "Server filter: $serverFilter"
Write-Host "Flow status: $FlowStatus"
Write-Host "Start UTC: $exportStart"
Write-Host "End UTC:   $exportEnd"
Write-Host "Chunk size: $ChunkDays day(s)"

while ($currentStart -lt $exportEnd) {
    $currentChunkDays = [double]$ChunkDays
    $chunkSucceeded = $false

    while (-not $chunkSucceeded) {
        $candidateEnd = $currentStart.AddDays($currentChunkDays)
        if ($candidateEnd -gt $exportEnd) { $candidateEnd = $exportEnd }

        $startText = Convert-ToKqlDateTime $currentStart
        $endText = Convert-ToKqlDateTime $candidateEnd
        $query = $queryTemplate.Replace("{{START_UTC}}", $startText).
            Replace("{{END_UTC}}", $endText).
            Replace("{{SERVER_FILTER}}", $serverFilter).
            Replace("{{FLOW_STATUS}}", $FlowStatus)

        $chunkNumber++
        $chunkLabel = "{0:D4}_{1}_{2}" -f $chunkNumber,
            $currentStart.ToString("yyyyMMddHHmmss"),
            $candidateEnd.ToString("yyyyMMddHHmmss")
        $chunkCsv = Join-Path $chunkDirectory "$chunkLabel.csv"
        $queryLog = Join-Path $logDirectory "$chunkLabel.log"

        Write-Host "Query interval: $startText to $endText ($currentChunkDays day(s))"
        $response = Invoke-LogAnalyticsQuery $WorkspaceId $query $queryLog
        $rows = @(Convert-LogAnalyticsTableToObjects $response)
        Assert-ExpectedColumns $rows

        if ($rows.Count -ge $SafetyRowLimit) {
            throw "Query interval returned at least $SafetyRowLimit rows: $startText to $endText. Retry with a smaller ChunkDays value."
        }

        if ($rows.Count -gt 0) {
            Export-ObjectsToCsv $rows $chunkCsv
            $chunkFiles.Add($chunkCsv)
        }

        $currentStart = $candidateEnd
        $chunkSucceeded = $true
    }
}

$finalCsv = Join-Path $runDirectory "$ServerName.csv"
if ($chunkFiles.Count -eq 0) {
    "SrcEff,DestEff,DestPort,L4Protocol,Server,SrcPorts,Requests" |
        Out-File -LiteralPath $finalCsv -Encoding utf8
    $finalRowCount = 0
}
else {
    $finalRowCount = Merge-FlowRows $chunkFiles.ToArray() $finalCsv
}

if ($finalRowCount -gt $GeneratorRowLimit) {
    throw "Final aggregated result contains $finalRowCount rows, exceeding the generator limit of $GeneratorRowLimit. No rows were discarded."
}

$elapsed = (Get-Date) - $scriptStart
Write-Step "Flow-log export completed"
Write-Host "Chunk files: $($chunkFiles.Count)"
Write-Host "Final CSV: $finalCsv"
Write-Host "Elapsed time: $elapsed"
Write-Output $finalCsv
