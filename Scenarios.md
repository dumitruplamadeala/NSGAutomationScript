# Demo Scenarios

- [ ] 1. Happy flow on empty NSG: dry-run first, review plan, then apply Create actions.
- [ ] 2. Happy flow on non-empty NSG: update an existing rule, review plan, apply, and show post-apply verification.
- [ ] 3. Wrong Excel sheet name or missing required header: fail fast during workbook schema validation and fix the Excel file.
- [ ] 4. Wrong field content caused by formatting mistakes: missing separator(comma) in `Source IP address / Subnet / Range IP`, `Destination IP address / Subnet / Range IP`, or `Destination Port or Service`; fix the Excel file.
- [ ] 5. Wrong IP format such as `10.10.10.10-12`: validation fails and the Excel file must be corrected.
- [ ] 6. Port overlap inside one rule, for example duplicate port or overlap with a range: validation fails and the Excel file must be corrected.
- [ ] 7. Shadowed rule: script raises a warning, marks the rule as `SkipShadowed`, and does not send that rule to Azure; implementation team should report it back to the workbook owner.
- [ ] 8. Conflict with live NSG state: same name or priority already used by another existing rule. Also mention that duplicate rule names inside the workbook are a validation error, not a live conflict.
- [ ] 9. Apply interrupted midway, for example terminal/network issue during apply: show checkpoint status, partial completion, and safe re-run behavior.