# Workbook data guardrails

## Settings table

The workbook must contain an Excel table named `Settings` with columns named `Setting` and `Value`. It must contain exactly one nonblank value for each required setting:

| Setting | Required value |
| --- | --- |
| `Reporting currency` | A three-letter currency code |
| `Warning threshold` | A number from 0 through 1 |

Do not create, repair, infer, or prompt for these settings. Stop if the table or either value is invalid.

## Qualifying quote table

Search worksheets and their tables in workbook order. Ignore the `Settings` table and use the first table that qualifies. Never prioritize the current selection.

A qualifying table:

- has at least one data row;
- has exactly one column for each required header: `Item`, `Vendor`, `Quote amount`, `Quote currency`, and `Budget limit`;
- has a nonblank string in every `Item` and `Vendor` cell;
- has a finite nonnegative number in every `Quote amount` cell;
- has a three-letter currency code in every `Quote currency` cell; and
- has a finite positive number in every `Budget limit` cell.

Header comparison is case-insensitive after trimming. A table with an incomplete or incorrectly typed required cell does not qualify. Continue searching for the next table; do not repair it or report individual rows.

## Change boundaries

- Add or update only `Exchange rate`, `Converted quote`, and `Budget status`.
- Do not rename or overwrite source columns.
- Stop if an output column name conflicts with a source column or duplicated header.
- Use only the latest exchange rate.
- Handle only one quote table per invocation.
- Do not create a summary.
