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

- `Action`
- `Rule Name`
- `Rule Number`
- `Destination Protocol`
- `Source IP address / Subnet / Range IP`
- `Destination IP address / Subnet / Range IP`
- `Destination Port or Service`

Allowed `Action` values are `Create`, `Update`, `Remove`, and `No-Change`.

### Pipeline filename convention

`-ResourceGroupName` and `-NsgName` can be omitted when the workbook filename follows:

```text
<resource-group>__<nsg-name>_NetworkAccessRequest_v<number[.number...]>.xlsx
```

Examples:

- `NSG-test-script__nsg-ci-automation-app_NetworkAccessRequest_v1.xlsx`
- `NSG-test-script__nsg-ci-automation-app_NetworkAccessRequest_v0.1.xlsx`
- `NSG-test-script__nsg-ci-automation-app_NetworkAccessRequest_v1.2.3.xlsx`

The double underscore separates the resource group from the NSG name. The version suffix must contain one or more dot-separated numeric components. It is validated as part of the filename but is not otherwise read or used by the script. Explicit `-ResourceGroupName` and `-NsgName` values take precedence over values in the filename.

`-Description` is also optional and defaults to `Generated for review only`.

## Main behavior

### Action reconciliation

The source workbook's `Action` value is retained as requested intent. The reviewed workbook updates it from the live NSG comparison:

| Live Azure state | Reviewed workbook `Action` |
| --- | --- |
| Rule is absent | `Create` |
| Matching Allow rule is identical | `No-Change` |
| Matching Allow rule differs | `Update` |
| Matching rule has Azure `Access` set to `Deny` | `Remove` |

`Remove` is reporting only. The script does not delete Azure rules, including when `-Apply` is used.

### Unmanaged NSG rules

A live Azure rule is reported as `Unmanaged` when its normalized direction and priority are not present in the workbook. Same-priority rules with a different name remain conflicts rather than unmanaged findings.

Unmanaged rules are:

- logged as warnings
- included in the approval YAML with their live name, priority, direction, access, and protocol
- included in checkpoint metadata under `UnmanagedNsgRules`

Unmanaged findings are informational and never cause an Azure change.

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
- `ReportedRemove`: a matching live Deny rule was mapped to workbook action `Remove`
- `SkippedShadowed`: planned action was `SkipApply` due to overlap/shadowing detection
- `BlockedConflict`: planned action was `Conflict`
- `Pending`: planned `Create` or `Update` that was not executed yet (for example dry-run)
- `Completed`: planned `Create` or `Update` succeeded
- `Failed`: planned `Create` or `Update` failed during apply

Action-specific payload behavior:

- `NoChange`: minimal record only
- `Create`: includes `After` and `Changes`
- `Update`: includes `Before`, `After`, and `Changes`
- `Remove`: includes `Before` and `Changes`; reporting only
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

### Pipeline-friendly dry-run

```powershell
.\apply_nsg_rules.ps1 `
  -WorkbookPath .\NSG-test-script__nsg-ci-automation-app_NetworkAccessRequest_v1.xlsx
```

This resolves the resource group and NSG from the filename and uses the default review-only description.

### Dry-run

```powershell
.\apply_nsg_rules.ps1 `
  -WorkbookPath .\mockData.xlsx `
  -ResourceGroupName "NSG-test-script" `
  -NsgName "nsg-ci-automation-app" `
  -Description "RITM12345 - NY5678 - 03/20/2026 - Create"
```

### Dry-run with explicit checkpoint file

```powershell
.\apply_nsg_rules.ps1 `
  -WorkbookPath .\mockData.xlsx `
  -ResourceGroupName "NSG-test-script" `
  -NsgName "nsg-ci-automation-app" `
  -Description "RITM12345 - NY5678 - 03/20/2026 - Update" `
  -CheckpointPath .\nsg-apply-checkpoint.json `
  -CheckpointMode Overwrite
```

### Apply with archived checkpoint and stable latest file

```powershell
.\apply_nsg_rules.ps1 `
  -WorkbookPath .\mockData.xlsx `
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
  -WorkbookPath .\mockData.xlsx `
  -ResourceGroupName "NSG-test-script" `
  -NsgName "nsg-ci-automation-app" `
  -Description "RITM12345 - NY5678 - 03/20/2026 - Create" `
```

## Important parameters

- `-WorkbookPath`: source Excel file
- `-ResourceGroupName`: target Azure resource group; optional with the pipeline filename convention
- `-NsgName`: target NSG name; optional with the pipeline filename convention
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
- `-Description`: managed rule description text; defaults to `Generated for review only`

## Notes for operators

- `NoChange` means the live rule fingerprint already matches the desired state.
- `Remove` means the matching Azure rule has `Access` set to `Deny`; it does not delete the rule.
- `Unmanaged` means no workbook rule has the same normalized direction and priority.
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