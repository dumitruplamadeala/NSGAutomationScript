# NSG Apply Script

## Purpose

`apply_nsg_rules.ps1` imports intended NSG rules from an Excel workbook, compares them with the current state of an Azure Network Security Group, produces a checkpoint JSON file, and optionally applies the required changes.

The script is built for Windows PowerShell 5.1 and uses Azure CLI plus the `ImportExcel` module.

## What the script does

For each run, the script:

1. Validates prerequisites 
2. Reads the workbook sheets configured through `-SheetNames`.
3. Validates and normalizes workbook values.
4. Loads the live rules from the target NSG.
5. Builds a plan with actions such as `NoChange`, `Create`, `Update`, `SkipApply`, and `Conflict`.
6. Writes a checkpoint JSON file describing the run.
7. Applies changes only when `-Apply` is provided.

## Prerequisites

- Windows PowerShell 5.1
- Azure CLI available on `PATH`
- `ImportExcel` PowerShell module installed
- Access to the Azure subscription and target NSG

Install `ImportExcel` if needed:

```powershell
Install-Module ImportExcel -Scope CurrentUser
```

## Expected workbook structure

By default, the script reads these worksheets:

- `CoreRules`
- `AppRules`

Required workbook headers:

- `Rule Name`
- `Rule Number`
- `Destination Protocol`
- `Source IP address / Subnet / Range IP`
- `Destination IP address / Subnet / Range IP`
- `Destination Port or Service`

## Main behavior

### Dry-run

If you do not pass `-Apply`, the script only:

- validates input
- loads live Azure state
- builds the plan
- writes checkpoint output
- logs what would happen

### Apply mode

If you pass `-Apply`, the script:

- creates missing rules
- updates existing rules with matching names
- skips shadowed rules detected during validation
- blocks conflict items
- updates the checkpoint file as execution progresses

## Checkpoint output

The checkpoint file records the run plan and execution status.

Checkpoint files are execution and audit records. They are not used to resume partial runs automatically.

Execution status values:

- `Unchanged`: planned action was `NoChange`
- `SkippedShadowed`: planned action was `SkipApply` due to overlap/shadowing detection
- `BlockedConflict`: planned action was `Conflict`
- `Pending`: planned `Create` or `Update` that was not executed yet (for example dry-run)
- `Completed`: planned `Create` or `Update` succeeded
- `Failed`: planned `Create` or `Update` failed during apply

Action-specific payload behavior:

- `NoChange`: minimal record only
- `Create`: includes `After` and `Changes`
- `Update`: includes `Before`, `After`, and `Changes`
- `SkipApply`: includes `After`, and includes `Before` when the skipped item would have been an update
- `Conflict`: includes `Before`, `After`, and `Changes`

## Checkpoint naming modes

Use `-CheckpointMode` to control how checkpoint files are written.

### `Auto`

Default behavior:

- dry-run: overwrite only the stable checkpoint path
- apply: write both a timestamped file and the stable checkpoint path

### `Overwrite`

Always writes only the stable path from `-CheckpointPath`.

### `Timestamped`

Always writes only a timestamped file based on `-CheckpointPath`.

Example:

- base path: `nsg-apply-checkpoint.json`
- run file: `nsg-apply-checkpoint-20260320-185541.json`

### `Both`

Writes:

- a timestamped run file
- the stable checkpoint file

This is useful when you want both history and a predictable latest file.

## Common commands

### Dry-run

```powershell
.\apply_nsg_rules.ps1 `
  -WorkbookPath .\dummyData.xlsx `
  -ResourceGroupName "NSG-test-script" `
  -NsgName "nsg-ci-automation-app" `
  -Description "RITM12345 - NY5678 - 03/20/2026 - Create"
```

### Dry-run with explicit checkpoint file

```powershell
.\apply_nsg_rules.ps1 `
  -WorkbookPath .\dummyData.xlsx `
  -ResourceGroupName "NSG-test-script" `
  -NsgName "nsg-ci-automation-app" `
  -Description "RITM12345 - NY5678 - 03/20/2026 - Update" `
  -CheckpointPath .\nsg-apply-checkpoint.json `
  -CheckpointMode Overwrite
```

### Apply with archived checkpoint and stable latest file

```powershell
.\apply_nsg_rules.ps1 `
  -WorkbookPath .\dummyData.xlsx `
  -ResourceGroupName "NSG-test-script" `
  -NsgName "nsg-ci-automation-app" `
  -Description "RITM12345 - NY5678 - 03/20/2026 - Update" `
  -Apply `
  -CheckpointMode Both
```

### Run when execution policy blocks unsigned scripts

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\apply_nsg_rules.ps1 `
  -WorkbookPath .\dummyData.xlsx `
  -ResourceGroupName "NSG-test-script" `
  -NsgName "nsg-ci-automation-app" `
  -Description "RITM12345 - NY5678 - 03/20/2026 - Create" `
  -SkipPreflight
```

## Important parameters

- `-WorkbookPath`: source Excel file
- `-ResourceGroupName`: target Azure resource group
- `-NsgName`: target NSG name
- `-SheetNames`: source worksheets, default `CoreRules`, `AppRules`
- `-Direction`: `Inbound` or `Outbound`
- `-CheckpointPath`: base checkpoint file path
- `-CheckpointMode`: `Auto`, `Overwrite`, `Timestamped`, `Both`
- `-Apply`: execute Azure changes
- `-FailOnShadowing`: fail if overlap warnings are detected
- `-SkipPreflight`: skip dependency and Azure checks
- `-PassThru`: emit execution objects to the pipeline
- `-RetryCount`: retry attempts for Azure CLI calls
- `-RetryDelaySeconds`: delay between retries
- `-Description`: managed rule description text

## Notes for operators

- `NoChange` means the live rule fingerprint already matches the desired state.
- `Conflict` means the same priority already exists with a different rule name.
- Overlap warnings do not stop execution unless `-FailOnShadowing` is used.
- Shadowed rules are logged as `SkippedShadowed` and are not sent to Azure.
- The script does not delete NSG rules.

## Troubleshooting

### Execution policy error

Use process-scoped bypass:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

### Azure CLI not found

Install Azure CLI and ensure `az` is available on `PATH`.

### ImportExcel module missing

```powershell
Install-Module ImportExcel -Scope CurrentUser
```

### Workbook validation errors

Check:

- worksheet names
- header spelling
- rule number range
- rule name format
- IP/CIDR/range values
- port range definitions
- description format