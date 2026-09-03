<#
.SYNOPSIS
Orchestrates existing NSG Flow 1 tools without reimplementing their logic.

.NOTES
Update CONFIGURATION before copying this script to the target repository.
The mandatory applier adapter must return an object with MatchesExistingNsg,
UnmanagedRules, and optionally Warnings. It receives applier output, log, run dir.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# =========================== CONFIGURATION ===========================
$RepositoryRoot = 'C:\path\to\repository'
$ServerName = 'q41-mtax'
$InputWorkbook = Join-Path $RepositoryRoot 'input\NSG-test-script__nsg-ci-automation-app_NetworkAccessRequest_v1.xlsx'
$FlowLogExports = @(
    (Join-Path $RepositoryRoot 'flowlogs\q41-mtax\export-01.csv'),
    (Join-Path $RepositoryRoot 'flowlogs\q41-mtax\export-02.csv')
)
$RunDirectory = Join-Path $RepositoryRoot (Join-Path "runs\$ServerName" (Get-Date -Format 'yyyyMMdd_HHmmss'))

$ApplierDirectory = Join-Path $RepositoryRoot 'nsg-rules-applier'
$GeneratorDirectory = Join-Path $RepositoryRoot 'nsg-rules-generator'
$MergerDirectory = Join-Path $RepositoryRoot 'nsg-rules-merger'
$ApplierScript = Join-Path $ApplierDirectory 'apply_nsg_rules.ps1'
$GeneratorScript = Join-Path $GeneratorDirectory 'script.py'
$MergerScript = Join-Path $MergerDirectory 'script.py'
$TemplateScript = Join-Path $GeneratorDirectory 'apply_template.py'

# The applier inserts '.reviewed' before the extension and writes beside its input.
$InputWorkbookDirectory = Split-Path -Path $InputWorkbook -Parent
$InputWorkbookBaseName = [System.IO.Path]::GetFileNameWithoutExtension($InputWorkbook)
$FirstReviewedWorkbook = Join-Path $InputWorkbookDirectory "$InputWorkbookBaseName.reviewed.xlsx"
$GeneratedServerWorkbook = Join-Path $GeneratorDirectory "output\$ServerName\$ServerName.xlsx"
$MergedWorkbookName = "$InputWorkbookBaseName.reviewed.merged.xlsx"
$MergedWorkbook = Join-Path $MergerDirectory $MergedWorkbookName
$TemplateInputName = $MergedWorkbookName
$TemplatedWorkbook = Join-Path $GeneratorDirectory "$InputWorkbookBaseName.reviewed.merged.templated.xlsx"
$TemplatedWorkbookBaseName = [System.IO.Path]::GetFileNameWithoutExtension($TemplatedWorkbook)
$FinalReviewedWorkbook = Join-Path $GeneratorDirectory "$TemplatedWorkbookBaseName.reviewed.xlsx"

$PowerShellExecutable = 'powershell.exe'
$PythonExecutable = 'python.exe'
$AllowedUnmanagedRules = @('CATCH_MISSED_TRAFFIC_UDP', 'CATCH_MISSED_TRAFFIC_TCP')

# REQUIRED: replace with client-applier report parsing. Return:
# [pscustomobject]@{ MatchesExistingNsg = $true; UnmanagedRules = @(); Warnings = @() }
$FirstApplierValidationAdapter = $null
# ======================================================================

$summary = [ordered]@{
    InputWorkbook = $null; ServerName = $ServerName; FlowLogExports = @();
    Steps = New-Object System.Collections.Generic.List[string]
    IntermediateOutputFiles = New-Object System.Collections.Generic.List[string]
    FinalOutputFile = $null; FirstApplierValidation = 'Not run'; FinalApplierValidation = 'Not run'
    UnmanagedRules = @(); Warnings = New-Object System.Collections.Generic.List[string]; Error = $null
}

function Write-Message { param([string]$Message, [string]$Level = 'INFO'); Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')][$Level] $Message" }
function Require-File { param([string]$Path, [string]$Label); if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "$Label was not found: $Path" }; [IO.Path]::GetFullPath($Path) }
function Require-Directory { param([string]$Path, [string]$Label); if (-not (Test-Path -LiteralPath $Path -PathType Container)) { throw "$Label was not found: $Path" }; [IO.Path]::GetFullPath($Path) }
function Copy-Artifact {
    param([string]$Source, [string]$DestinationDirectory)
    $sourcePath = Require-File $Source 'Expected output file'
    $destination = Join-Path $DestinationDirectory ([IO.Path]::GetFileName($sourcePath))
    Copy-Item -LiteralPath $sourcePath -Destination $destination -Force
    $summary.IntermediateOutputFiles.Add($destination)
    $destination
}
function Invoke-Step {
    param([string]$Name, [string]$WorkingDirectory, [string]$Executable, [string[]]$Arguments, [string]$LogPath)
    Write-Message "Step: $Name"
    Write-Message "Working directory: $WorkingDirectory"
    Push-Location -LiteralPath $WorkingDirectory
    try {
        & $Executable @Arguments 2>&1 | Tee-Object -FilePath $LogPath
        if ($LASTEXITCODE -ne 0) { throw "Step '$Name' failed with exit code $LASTEXITCODE. See: $LogPath" }
    }
    finally { Pop-Location }
    $summary.Steps.Add("${Name}: Completed")
}
function Validate-Applier {
    param([string]$Name, [string]$OutputPath, [string]$LogPath)
    if ($null -eq $FirstApplierValidationAdapter) { throw "$Name validation adapter is not configured. Set `$FirstApplierValidationAdapter to parse the actual applier output." }
    $result = & $FirstApplierValidationAdapter $OutputPath $LogPath $RunDirectory
    if ($null -eq $result -or $result.PSObject.Properties.Name -notcontains 'MatchesExistingNsg' -or $result.PSObject.Properties.Name -notcontains 'UnmanagedRules') { throw "$Name validation adapter must return MatchesExistingNsg and UnmanagedRules." }
    $unmanaged = @($result.UnmanagedRules | ForEach-Object { [string]$_ } | Sort-Object -Unique)
    $unexpected = @($unmanaged | Where-Object { $_ -notin $AllowedUnmanagedRules })
    foreach ($warning in @($result.Warnings)) { if ($warning) { $summary.Warnings.Add([string]$warning) } }
    if (-not [bool]$result.MatchesExistingNsg) { throw "$Name validation failed: workbook does not match the existing NSG." }
    if ($unexpected.Count -gt 0) { throw "$Name validation failed: unexpected unmanaged rule(s): $($unexpected -join ', ')." }
    [pscustomobject]@{ UnmanagedRules = $unmanaged }
}
function Show-Summary {
    Write-Host ''; Write-Host 'FLOW 1 SUMMARY'; Write-Host '=============='
    Write-Host "Input workbook: $($summary.InputWorkbook)"; Write-Host "Server name: $($summary.ServerName)"
    Write-Host 'Flow logs:'; foreach ($item in $summary.FlowLogExports) { Write-Host "  - $item" }
    Write-Host 'Steps:'; foreach ($item in $summary.Steps) { Write-Host "  - $item" }
    Write-Host 'Artifacts:'; foreach ($item in $summary.IntermediateOutputFiles) { Write-Host "  - $item" }
    Write-Host "Final output: $($summary.FinalOutputFile)"
    Write-Host "First applier validation: $($summary.FirstApplierValidation)"
    Write-Host "Final applier validation: $($summary.FinalApplierValidation)"
    Write-Host "Unmanaged rules: $($summary.UnmanagedRules -join ', ')"
    foreach ($item in $summary.Warnings) { Write-Host "WARNING: $item" }
    if ($summary.Error) { Write-Host "ERROR: $($summary.Error)" }
}

try {
    $RepositoryRoot = Require-Directory $RepositoryRoot 'Repository root'
    $InputWorkbook = Require-File $InputWorkbook 'Input workbook'
    $FlowLogExports = @($FlowLogExports | ForEach-Object { Require-File $_ 'Flow-log export' })
    if ($FlowLogExports.Count -eq 0) { throw 'At least one flow-log export is required.' }
    $ApplierDirectory = Require-Directory $ApplierDirectory 'Applier directory'
    $GeneratorDirectory = Require-Directory $GeneratorDirectory 'Generator directory'
    $MergerDirectory = Require-Directory $MergerDirectory 'Merger directory'
    $ApplierScript = Require-File $ApplierScript 'Applier script'; $GeneratorScript = Require-File $GeneratorScript 'Generator script'
    $MergerScript = Require-File $MergerScript 'Merger script'; $TemplateScript = Require-File $TemplateScript 'Template script'
    New-Item -ItemType Directory -Path $RunDirectory -Force | Out-Null; $RunDirectory = [IO.Path]::GetFullPath($RunDirectory)
    $artifactDirectory = Join-Path $RunDirectory 'artifacts'; $logDirectory = Join-Path $RunDirectory 'logs'
    New-Item -ItemType Directory -Path $artifactDirectory -Force | Out-Null; New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
    $summary.InputWorkbook = $InputWorkbook; $summary.FlowLogExports = @($FlowLogExports)
    Copy-Artifact $InputWorkbook $artifactDirectory | Out-Null; foreach ($item in $FlowLogExports) { Copy-Artifact $item $artifactDirectory | Out-Null }

    $log = Join-Path $logDirectory '01-applier-initial.log'
    Invoke-Step '1. Apply initial NSG workbook' $ApplierDirectory $PowerShellExecutable @('-NoProfile','-ExecutionPolicy','Bypass','-File',$ApplierScript,'-WorkbookPath',$InputWorkbook) $log
    Copy-Artifact $FirstReviewedWorkbook $artifactDirectory | Out-Null
    $validation = Validate-Applier 'First applier' $FirstReviewedWorkbook $log; $summary.FirstApplierValidation = 'Passed'; $summary.UnmanagedRules = @($validation.UnmanagedRules)

    $log = Join-Path $logDirectory '02-generator.log'
    Invoke-Step '2. Generate rules from flow logs' $GeneratorDirectory $PythonExecutable (@($GeneratorScript,$ServerName) + $FlowLogExports) $log
    Copy-Artifact $GeneratedServerWorkbook $artifactDirectory | Out-Null

    $log = Join-Path $logDirectory '03-merger.log'
    Invoke-Step '3. Merge reviewed and generated workbooks' $MergerDirectory $PythonExecutable @($MergerScript,$FirstReviewedWorkbook,$GeneratedServerWorkbook,$MergedWorkbook) $log
    Copy-Artifact $MergedWorkbook $artifactDirectory | Out-Null

    $templateInput = Join-Path $GeneratorDirectory $TemplateInputName
    Copy-Item -LiteralPath (Require-File $MergedWorkbook 'Merged workbook') -Destination $templateInput -Force
    Copy-Artifact $templateInput $artifactDirectory | Out-Null
    $log = Join-Path $logDirectory '04-template.log'
    Invoke-Step '4. Apply Excel template' $GeneratorDirectory $PythonExecutable @($TemplateScript,$TemplateInputName) $log
    Copy-Artifact $TemplatedWorkbook $artifactDirectory | Out-Null

    $log = Join-Path $logDirectory '05-applier-final.log'
    Invoke-Step '5. Apply templated NSG workbook' $ApplierDirectory $PowerShellExecutable @('-NoProfile','-ExecutionPolicy','Bypass','-File',$ApplierScript,'-WorkbookPath',$TemplatedWorkbook) $log
    $summary.FinalOutputFile = Copy-Artifact $FinalReviewedWorkbook $artifactDirectory
    $null = Validate-Applier 'Final applier' $FinalReviewedWorkbook $log; $summary.FinalApplierValidation = 'Passed'
    Show-Summary
}
catch { $summary.Error = $_.Exception.Message; Show-Summary; throw }
