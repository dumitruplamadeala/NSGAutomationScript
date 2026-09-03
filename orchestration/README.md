# Flow 1 Orchestrator

`run-flow1.ps1` runs the existing NSG tools for Flow 1 in a fixed sequence. It orchestrates the existing PowerShell and Python scripts; it does not reproduce their NSG, flow-log, merge, or Excel-template logic.

The script is a portable starting point for the client repository. Configure it on the client machine after confirming each existing tool's command-line interface and output location.

## Flow

1. Run the NSG applier with the input workbook.
2. Validate the first applier result through a configured adapter.
3. Run the flow-log generator for the configured server and flow-log exports.
4. Merge the reviewed NSG workbook with the generated server workbook.
5. Copy the merged workbook to the generator directory and apply the Excel template.
6. Run the NSG applier against the templated workbook.
7. Validate the final applier result through the same adapter.

The pipeline stops at the first failed command, missing input, missing expected output, or failed applier validation.

## Prerequisites

- Windows PowerShell 5.1 or PowerShell with compatible script behavior.
- `powershell.exe` available on `PATH`, or update `$PowerShellExecutable`.
- Python available on `PATH`, or update `$PythonExecutable`.
- All existing PowerShell/Python tool dependencies installed in the client repository environment.
- Access required by the NSG applier, including its Azure CLI/module prerequisites.
- Flow-log exports collected before the pipeline starts.

## Configuration

Open [run-flow1.ps1](run-flow1.ps1) and update the `CONFIGURATION` section. No command-line parameters are currently required.

| Setting | Purpose |
| --- | --- |
| `$RepositoryRoot` | Absolute path to the client repository root. |
| `$ServerName` | Server passed as the first generator argument. |
| `$InputWorkbook` | Initial NSG workbook. |
| `$FlowLogExports` | One or more prepared flow-log export files. |
| `$RunDirectory` | Unique output folder for this execution. Keep the timestamp component. |
| `$ApplierDirectory`, `$GeneratorDirectory`, `$MergerDirectory` | Tool working directories. |
| `$ApplierScript`, `$GeneratorScript`, `$MergerScript`, `$TemplateScript` | Existing tool entry points. |
| `$FirstReviewedWorkbook` | Derived applier output beside `$InputWorkbook`, using the applier's `.reviewed.xlsx` convention. |
| `$MergedWorkbook`, `$TemplateInputName`, `$TemplatedWorkbook` | Expected merger/template artifacts; verify the client tools' behavior. |
| `$FinalReviewedWorkbook` | Derived applier output beside the templated workbook, using `.reviewed.xlsx`. |
| `$AllowedUnmanagedRules` | The only unmanaged NSG rule names allowed after applier validation. |
| `$FirstApplierValidationAdapter` | Required repository-specific validation adapter. |

Example configuration for `q41-mtax`:

```powershell
$RepositoryRoot = 'C:\client\nsg-automation'
$ServerName = 'q41-mtax'

$InputWorkbook = Join-Path $RepositoryRoot 'input\NSG-test-script__nsg-ci-automation-app_NetworkAccessRequest_v1.xlsx'
$FlowLogExports = @(
    (Join-Path $RepositoryRoot 'flowlogs\q41-mtax\export-01.csv'),
    (Join-Path $RepositoryRoot 'flowlogs\q41-mtax\export-02.csv')
)
$RunDirectory = Join-Path $RepositoryRoot (Join-Path "runs\$ServerName" (Get-Date -Format 'yyyyMMdd_HHmmss'))
```

Paths are passed as argument arrays, so paths containing spaces are supported.

## Applier Validation Adapter

The orchestrator cannot infer whether the applier's workbook matches Azure or which unmanaged rules it found. Connect `$FirstApplierValidationAdapter` to the actual report, JSON, YAML, checkpoint, or pipeline object produced by the client applier.

The adapter receives three arguments:

```powershell
param($ApplierOutputPath, $ApplierLogPath, $RunDirectory)
```

It must return this shape:

```powershell
[pscustomobject]@{
    MatchesExistingNsg = $true
    UnmanagedRules = @(
        'CATCH_MISSED_TRAFFIC_UDP',
        'CATCH_MISSED_TRAFFIC_TCP'
    )
    Warnings = @()
}
```

The pipeline stops when:

- `MatchesExistingNsg` is `$false`.
- An unmanaged rule is not listed in `$AllowedUnmanagedRules`.
- The adapter returns an invalid object.

The same adapter runs after the initial and final applier invocation. Adapt the function if the client tools require different validation criteria for each stage.

## Run Outputs

The pipeline preserves original tool outputs in their existing locations. It also copies inputs and artifacts into an isolated run directory:

```text
<repository-root>/
  runs/
    q41-mtax/
      20260903_143000/
        artifacts/
          input-workbook.xlsx
          export-01.csv
          export-02.csv
          <input>.reviewed.xlsx
          q41-mtax.xlsx
          <input>.reviewed.merged.xlsx
          <input>.reviewed.merged.templated.xlsx
          <input>.reviewed.merged.templated.reviewed.xlsx
        logs/
          01-applier-initial.log
          02-generator.log
          03-merger.log
          04-template.log
          05-applier-final.log
```

The script creates a timestamped run directory, avoiding collisions with previous runs. `artifacts` are copies for diagnosis and audit; they do not change standalone output behavior.

## Usage

After updating the configuration and validation adapter:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\orchestration\run-flow1.ps1
```

Review the console's `FLOW 1 SUMMARY` and the generated `logs` and `artifacts` directories after every run.

## Client Verification Checklist

Before the first integrated execution, verify:

- The exact script names, particularly whether the merger is `script.py` or another file.
- The Python interpreter command and virtual environment requirements.
- Every tool's working-directory requirement.
- Exact reviewed, generated, merged, and templated workbook names and locations.
- Whether tools overwrite prior outputs and whether configured expected files should be cleared or uniquely named before a run.
- Whether the template tool modifies its input in place or writes a new output file.
- The source that authoritatively reports NSG match state and unmanaged rule names.
- The applier adapter against successful, mismatched, allowed-unmanaged, and unexpected-unmanaged examples.
- The flow-log file format and generator requirements.
- Write permissions for tool output folders and `<repository-root>/runs`.

## Testing Checklist

1. Run each existing tool manually using the same configured inputs and confirm expected files.
2. Configure a validation adapter using real applier output.
3. Run the orchestrator with one valid flow-log export.
4. Run with multiple exports and verify every path reaches the generator.
5. Use paths containing spaces for the repository, workbook, and export files.
6. Temporarily configure a missing input or output and confirm the pipeline stops at the correct step.
7. Return an unexpected unmanaged rule from the adapter and confirm the pipeline stops after Step 1.
8. Return `MatchesExistingNsg = $false` and confirm the pipeline stops after Step 1.
9. Confirm all five logs and all expected artifacts are copied to the run directory.
10. Confirm prior run directories remain unchanged.

## Failure Information To Collect

For integration support, provide:

- The console output from the failed run.
- The relevant file from `runs/<server>/<timestamp>/logs`.
- The command that successfully runs the affected client tool manually.
- The complete filename and directory of the actual output produced by that tool.
- The first 50 lines of the applier report, checkpoint, JSON, or YAML used for validation, with sensitive values redacted.
- The adapter code and the object it returns, if validation failed.
- Python traceback output and package/environment details for Python failures.

Do not include credentials, access tokens, secrets, or unrestricted production flow-log data.
