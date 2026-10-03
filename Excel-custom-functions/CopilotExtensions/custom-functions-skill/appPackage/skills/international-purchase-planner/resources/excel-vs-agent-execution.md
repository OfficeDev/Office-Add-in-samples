# Excel vs non-Excel execution guidance

## Inside Excel

- Use the workbook as the source of truth.
- Run `scripts/prepare-purchase-comparison.js` only for requests to convert complete vendor quotes or identify quotes near or over budget.
- Let the script perform all workbook changes.

## Outside Excel

- Do not run the script or claim that a workbook was changed.
- Explain that the skill must be invoked in Copilot in Excel with the workbook open.

## Quality rules

- Do not run the script for general purchasing, currency, or budgeting advice.
- Do not ask for missing workbook values.
- Report script errors without claiming partial success.
- Never claim that a workbook update succeeded unless the script completed.
