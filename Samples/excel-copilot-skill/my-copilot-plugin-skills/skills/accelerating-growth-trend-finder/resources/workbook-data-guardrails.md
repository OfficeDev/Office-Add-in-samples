# Workbook data guardrails

Use this reference before running the accelerating growth trend script.

## Data requirements

- Use only the first worksheet in the current workbook.
- Use the first table on that worksheet that has at least 12 columns.
- Ignore the table header row when checking whether the data is numeric.
- Require every table body cell to contain a finite number.
- Stop at the first table that satisfies the column-count and numeric-data requirements.

## Search rules

- Do not search other worksheets.
- Do not search loose ranges outside Excel tables.
- Do not rename sheets, tables, or headers.
- Do not create helper columns or formulas before running the script.

## Quality checks

- If no table qualifies, report that no qualifying table was found.
- If a table qualifies but no rows match, report that no rows in the qualifying table embody the trend.
- Treat blank cells, text, errors, booleans, and dates that aren't returned as numbers as nonnumeric data.
- If the user requests exponential trend data, report to the user that you will look for accelerating growth trend data, and that while all exponential trends are accelerating growth trends, the converse is not the case, so some of the trends you find may not be exponential.
