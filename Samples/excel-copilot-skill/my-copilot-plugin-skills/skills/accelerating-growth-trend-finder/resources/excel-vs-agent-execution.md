# Excel vs non-Excel execution guidance

## When the skill runs inside Excel

- Use the workbook as the source of truth.
- Follow the instructions in the **Workflow**, **Workbook output**, and **Copilot chat output** sections of the SKILL.md file.

## When the skill runs outside Excel

- Do not claim that workbook rows were analyzed.
- Explain that the skill can only be used in Copilot in Excel.
- Ask the user to open the workbook in Excel and invoke the skill there.

## When to use the scripts folder

- Run `scripts/find-accelerating-growth-trend-rows.js` only when the user asks to find table rows with accelerating, or increasingly larger left-to-right growth.
- Do not run the script for general explanations of accelerating growth growth.
- Do not run the script when the current environment cannot execute Excel Office.js APIs.
