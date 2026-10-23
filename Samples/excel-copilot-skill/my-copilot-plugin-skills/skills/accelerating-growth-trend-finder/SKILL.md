---
name: accelerating-growth-trend-finder
description: |
    Use this skill in Excel when the user asks to find rows in a workbook table that show an accelerating growth, increasingly larger increases, or numeric values that rise faster from left to right.
metadata:
    category: Excel analysis
    version: 1.0.0
    tags: excel, office-js, tables, trends, analysis
---

# Accelerating growth trend finder

Find rows in the first qualifying Excel table whose numeric values increase from left to right with increasingly larger increases.

## Reference resources

Before running the script, consult:

- `resources/workbook-data-guardrails.md`
- `resources/excel-vs-agent-execution.md`

Use these resources to confirm that the current workbook is the right source of truth, that the skill is running inside Excel, and that the user is asking for workbook analysis rather than general advice.

## Workflow

1. Confirm that the current context is Excel.
2. Run `scripts/find-accelerating-growth-trend-rows.js` when the user asks to identify rows with an accelerating growth trend.

## Workbook output

Let `scripts/find-accelerating-growth-trend-rows.js` create charts, but do not create new worksheets or formulas.

## Copilot chat output

1. For each chart that is created, report the table name and row that is the chart's source.
2. If there are no rows with an accelerating growth trend, report the problem. Use the exact error message that is returned by `scripts/find-accelerating-growth-trend-rows.js`. Do not reword it.
3. If there are any other problems, such as no qualifying tables, report the problem.

## Common pitfalls to avoid

- Do not inspect tables outside the first worksheet.
- Do not analyze a table unless it has at least 12 columns.
- Do not treat the header row as numeric data.
- Do not infer, coerce, or fill missing values.
- Do not run the Office.js script outside Excel.

