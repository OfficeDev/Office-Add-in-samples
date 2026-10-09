---
name: international-purchase-planner
description: |
  Use this skill in Excel when the user asks to compare complete vendor quotes
  in one reporting currency or identify quotes that are near or over budget.
metadata:
  category: Excel analysis
  version: 1.0.0
  tags: excel, office-js, custom-functions, currency, purchasing, budgeting
---

# International Purchase Planner

Convert complete vendor quotes with the latest exchange rates and classify each converted quote against its budget.

## Reference resources

Before running the script, consult:

- `resources/workbook-data-guardrails.md`
- `resources/excel-vs-agent-execution.md`

## Workflow

1. Confirm that the current context is Excel.
2. Confirm that the workbook contains a valid table named `Settings`.
3. Run `scripts/prepare-purchase-comparison.js`.
4. Report the quote table name, reporting currency, warning threshold, and number of processed rows.

## Workbook output

The script must use the first qualifying quote table in workbook order. It must add or update these calculated columns:

- `Exchange rate`
- `Converted quote`
- `Budget status`

The script must insert formulas, not static results. After calculation finishes, it must format numeric output, apply status-based conditional formatting, and autofit each output column to its header and calculated values. It must not create a summary or modify source columns.

## Copilot chat output

Report:

- the name of the processed table;
- the reporting currency;
- the warning threshold; and
- the number of processed quote rows.

If no table qualifies or `Settings` is invalid, report the exact error returned by the script.

## Common pitfalls to avoid

- Do not process any table other than the first qualifying table in workbook order.
- Do not infer or prompt for settings.
- Do not process incomplete or incorrectly typed quote data.
- Do not replace formulas with static values.
- Do not run the Office.js script outside Excel.
