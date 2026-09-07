<#
.SYNOPSIS
Exports inbound NSG flow-log traffic from Azure Log Analytics to CSV.

.DESCRIPTION
Executes a parameterized KQL query against an Azure Log Analytics workspace
using configurable time-based intervals. The script:

- Queries the workspace in manageable time chunks.
- Automatically retries an oversized chunk with smaller day-based intervals.
- Supports filtering by flow status, such as Denied or Allowed.
- Saves the generated query interval and execution logs.
- Uses temporary chunk TSV files while the export is running.
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

The final CSV contains every aggregated flow row.

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
    [int]$MinimumChunkDays = 1,
    [ValidateSet("Allowed", "Denied")]
    [string]$FlowStatus = "Denied",
    [int]$SafetyRowLimit = 450000,
    [string]$PythonExecutable = ".\venv\Scripts\python.exe",
    [switch]$DisableQueryLogs
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
        [string]$OutputPath,
        [string]$LogFilePath
    )

    # Keep stdout and stderr separate. Python parses the TSV chunk later.
    $tempDirectory = [System.IO.Path]::GetTempPath()
    $stderrPath = Join-Path $tempDirectory ("{0}.stderr" -f [guid]::NewGuid())
    $logReference = if ([string]::IsNullOrWhiteSpace($LogFilePath)) {
        ""
    }
    else {
        " See: $LogFilePath"
    }

    try {
        # az.cmd on Windows can misparse embedded double quotes in a dynamic
        # argument. KQL accepts single-quoted string literals, so normalize
        # the query before passing it to Azure CLI.
        $queryArgument = $Query -replace '"', "'"
        $queryArgument = $queryArgument -replace "[\r\n]+", " "

        & az monitor log-analytics query `
            --workspace $WorkspaceId `
            --analytics-query $queryArgument `
            --output tsv `
            1> $OutputPath `
            2> $stderrPath

        $exitCode = $LASTEXITCODE
        $stderrText = if (Test-Path -LiteralPath $stderrPath) {
            Get-Content -LiteralPath $stderrPath -Raw
        }
        else { "" }

        if (-not [string]::IsNullOrWhiteSpace($LogFilePath)) {
            @(
                "=== STDOUT ==="
                (Get-Content -LiteralPath $OutputPath -Raw -ErrorAction SilentlyContinue)
                ""
                "=== STDERR ==="
                $stderrText
            ) | Out-File -LiteralPath $LogFilePath -Encoding utf8
        }

        if ($exitCode -ne 0) {
            throw "Azure CLI query failed with exit code $exitCode. $stderrText$logReference"
        }

        $rowCount = 0
        foreach ($line in [System.IO.File]::ReadLines($OutputPath)) {
            if ([string]::IsNullOrWhiteSpace($line)) {
                continue
            }

            # The CLI can emit the result-table name as a metadata line.
            if ($line -eq "PrimaryResult") {
                continue
            }

            $rowCount++
        }

        return $rowCount
    }
    finally {
        Remove-Item -LiteralPath $stderrPath -Force -ErrorAction SilentlyContinue
    }
}

function Test-RetryableQueryLimitError {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)

    $message = [string]$ErrorRecord.Exception.Message

    return $message -match '(?i)(500000|64\s*MB|too\s+many|too\s+large|result.{0,30}(limit|size)|response.{0,30}(limit|size)|maximum.{0,30}(row|record|size)|truncat|exceed.{0,20}(limit|size))'
}

# Validation
Write-Step "Validating configuration"
Assert-CommandExists "az"
Assert-CommandExists $PythonExecutable
$aggregationScript = Join-Path $PSScriptRoot "aggregate_flow_logs.py"

if (-not (Test-Path -LiteralPath $aggregationScript -PathType Leaf)) {
    throw "Python aggregation script was not found: $aggregationScript"
}

if (-not (Test-Path -LiteralPath $QueryTemplatePath -PathType Leaf)) {
    throw "KQL query template was not found: $QueryTemplatePath"
}
if ($HistoryDays -le 0 -or $ChunkDays -le 0 -or $MinimumChunkDays -le 0) {
    throw "HistoryDays, ChunkDays, and MinimumChunkDays must be greater than zero."
}
if ($MinimumChunkDays -gt $ChunkDays) {
    throw "MinimumChunkDays must be less than or equal to ChunkDays."
}
if ($SafetyRowLimit -le 0) {
    throw "SafetyRowLimit must be greater than zero."
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

# Every execution gets an isolated output directory. Intermediate chunks and the
# aggregation database are kept in a temporary staging directory instead.
$runDirectory = Join-Path $OutputDirectory (Get-Date -Format "yyyyMMdd_HHmmss")
$logDirectory = Join-Path $runDirectory "logs"
$stagingDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ("nsg-flow-export-{0}" -f [guid]::NewGuid())
$chunkDirectory = Join-Path $stagingDirectory "chunks"
$directoriesToCreate = @($runDirectory, $chunkDirectory)
if (-not $DisableQueryLogs) {
    $directoriesToCreate += $logDirectory
}
New-Item -ItemType Directory -Path $directoriesToCreate -Force | Out-Null

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
        $chunkTsv = Join-Path $chunkDirectory "$chunkLabel.tsv"
        $queryLog = if ($DisableQueryLogs) {
            $null
        }
        else {
            Join-Path $logDirectory "$chunkLabel.log"
        }

        Write-Host "Query interval: $startText to $endText ($currentChunkDays day(s))"

        try {
            $rowCount = Invoke-LogAnalyticsQuery $WorkspaceId $query $chunkTsv $queryLog
        }
        catch {
            if (-not (Test-RetryableQueryLimitError $_)) {
                throw
            }

            if ($currentChunkDays -le $MinimumChunkDays) {
                throw "The minimum interval still exceeded the Azure query result limit: $startText to $endText. See: $queryLog"
            }

            $chunkNumber--
            $currentChunkDays = [math]::Max(
                [math]::Floor($currentChunkDays / 2),
                [double]$MinimumChunkDays
            )

            Write-Warning "Azure returned a result-size/row-limit error. Retrying with $currentChunkDays day(s). See: $queryLog"
            continue
        }

        if ($rowCount -ge $SafetyRowLimit) {
            if ($currentChunkDays -le $MinimumChunkDays) {
                throw "The minimum interval still returned at least $SafetyRowLimit rows: $startText to $endText"
            }

            $chunkNumber--
            $currentChunkDays = [math]::Max(
                [math]::Floor($currentChunkDays / 2),
                [double]$MinimumChunkDays
            )

            Write-Warning "Interval returned $rowCount rows. Retrying with $currentChunkDays day(s)."
            continue
        }

        if ($rowCount -gt 0) {
            $chunkFiles.Add($chunkTsv)
        }
        else {
            Remove-Item -LiteralPath $chunkTsv -Force -ErrorAction SilentlyContinue
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
    $aggregationDatabase = Join-Path $stagingDirectory "flow-aggregation.sqlite"

    Write-Step "Aggregating flow-log chunks with Python"

    $pythonOutput = @(& $PythonExecutable $aggregationScript `
        --chunk-directory $chunkDirectory `
        --output-csv $finalCsv `
        --database $aggregationDatabase)

    $pythonOutput | ForEach-Object { Write-Host $_ }

    if ($LASTEXITCODE -ne 0) {
        throw "Python flow-log aggregation failed with exit code $LASTEXITCODE."
    }

    $countLine = $pythonOutput |
        Where-Object { $_ -match "^Final aggregated rows:\s+(\d+)" } |
        Select-Object -First 1

    if ($countLine -and $countLine -match "(\d+)$") {
        $finalRowCount = [int64]$Matches[1]
    }
    else {
        throw "Python aggregation completed but did not return the final row count."
    }
}

# Intermediate chunks and the SQLite aggregation database are implementation
# details and must not remain in the export output.
if (Test-Path -LiteralPath $stagingDirectory) {
    Remove-Item -LiteralPath $stagingDirectory -Recurse -Force
}

$elapsed = (Get-Date) - $scriptStart
Write-Step "Flow-log export completed"
Write-Host "Chunk files: $($chunkFiles.Count)"
Write-Host "Final CSV: $finalCsv"
Write-Host "Elapsed time: $elapsed"
Write-Output $finalCsv
