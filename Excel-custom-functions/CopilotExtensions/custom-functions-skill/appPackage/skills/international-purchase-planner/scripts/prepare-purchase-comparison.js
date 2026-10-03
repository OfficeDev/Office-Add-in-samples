const REQUIRED_QUOTE_COLUMNS = [
  "Item",
  "Vendor",
  "Quote amount",
  "Quote currency",
  "Budget limit",
];
const OUTPUT_COLUMNS = ["Exchange rate", "Converted quote", "Budget status"];
const CURRENCY_CODE_PATTERN = /^[A-Z]{3}$/;
const CALCULATION_TIMEOUT_MS = 120000;
const CALCULATION_POLL_INTERVAL_MS = 250;

await Excel.run(async (context) => {
  const settings = await readSettings(context);
  const quoteTable = await findFirstQualifyingTable(context);

  if (!quoteTable) {
    throw new Error(
      "No qualifying quote table was found. Every required source cell must contain a value of the appropriate data type."
    );
  }

  await addMissingOutputColumns(context, quoteTable.table, quoteTable.columnNames);

  const exchangeRateColumn = quoteTable.table.columns.getItem(
    quoteTable.columnNames.get(normalizeHeader("Exchange rate"))
  );
  const convertedQuoteColumn = quoteTable.table.columns.getItem(
    quoteTable.columnNames.get(normalizeHeader("Converted quote"))
  );
  const budgetStatusColumn = quoteTable.table.columns.getItem(
    quoteTable.columnNames.get(normalizeHeader("Budget status"))
  );

  const exchangeRateRange = exchangeRateColumn.getDataBodyRange();
  const convertedQuoteRange = convertedQuoteColumn.getDataBodyRange();
  const budgetStatusRange = budgetStatusColumn.getDataBodyRange();

  exchangeRateRange.formulas = repeatFormula(
    quoteTable.rowCount,
    `=PURCHASEPLANNER.FXRATE([@[Quote currency]],"${settings.reportingCurrency}")`
  );
  convertedQuoteRange.formulas = repeatFormula(
    quoteTable.rowCount,
    "=[@[Quote amount]]*[@[Exchange rate]]"
  );
  budgetStatusRange.formulas = repeatFormula(
    quoteTable.rowCount,
    `=PURCHASEPLANNER.BUDGETSTATUS([@[Converted quote]],[@[Budget limit]],${settings.warningThreshold})`
  );

  exchangeRateRange.numberFormat = repeatFormat(quoteTable.rowCount, "0.0000");
  convertedQuoteRange.numberFormat = repeatFormat(
    quoteTable.rowCount,
    `[$${settings.reportingCurrency}]#,##0.00`
  );

  budgetStatusRange.conditionalFormats.clearAll();
  addStatusFormat(
    budgetStatusRange,
    "Within budget",
    "#E2F0D9",
    "#375623"
  );
  addStatusFormat(budgetStatusRange, "Near limit", "#FFF2CC", "#7F6000");
  addStatusFormat(budgetStatusRange, "Over budget", "#F4CCCC", "#9C0006");

  await context.sync();
  await waitForFinalValues(
    context,
    exchangeRateRange,
    convertedQuoteRange,
    budgetStatusRange
  );

  exchangeRateColumn.getRange().format.autofitColumns();
  convertedQuoteColumn.getRange().format.autofitColumns();
  budgetStatusColumn.getRange().format.autofitColumns();
  await context.sync();

  return JSON.stringify({
    table: quoteTable.table.name,
    reportingCurrency: settings.reportingCurrency,
    warningThreshold: settings.warningThreshold,
    processedRows: quoteTable.rowCount,
  });

  async function readSettings(excelContext) {
    const table = excelContext.workbook.tables.getItemOrNullObject("Settings");
    table.load(["isNullObject", "name"]);
    await excelContext.sync();

    if (table.isNullObject) {
      throw new Error('The workbook must contain a table named "Settings".');
    }

    const columns = table.columns;
    const rows = table.rows;
    columns.load("items/name");
    rows.load("count");
    await excelContext.sync();

    const columnMap = createColumnMap(columns.items);
    if (
      !columnMap.has(normalizeHeader("Setting")) ||
      !columnMap.has(normalizeHeader("Value")) ||
      rows.count === 0
    ) {
      throw new Error(
        'The "Settings" table must have nonempty "Setting" and "Value" columns.'
      );
    }

    const settingRange = table.columns
      .getItem(columnMap.get(normalizeHeader("Setting")))
      .getDataBodyRange();
    const valueRange = table.columns
      .getItem(columnMap.get(normalizeHeader("Value")))
      .getDataBodyRange();
    settingRange.load("values");
    valueRange.load("values");
    await excelContext.sync();

    const valuesBySetting = new Map();
    for (let rowIndex = 0; rowIndex < rows.count; rowIndex += 1) {
      const setting = settingRange.values[rowIndex][0];
      const value = valueRange.values[rowIndex][0];
      if (typeof setting !== "string" || setting.trim() === "" || isBlank(value)) {
        throw new Error('Every row in the "Settings" table must contain a setting and value.');
      }

      const key = setting.trim().toLowerCase();
      if (valuesBySetting.has(key)) {
        throw new Error(`The "Settings" table contains duplicate "${setting.trim()}" rows.`);
      }
      valuesBySetting.set(key, value);
    }

    const reportingValue = valuesBySetting.get("reporting currency");
    const reportingCurrency =
      typeof reportingValue === "string" ? reportingValue.trim().toUpperCase() : "";
    if (!CURRENCY_CODE_PATTERN.test(reportingCurrency)) {
      throw new Error(
        'The "Reporting currency" setting must be a three-letter currency code.'
      );
    }

    const warningThreshold = valuesBySetting.get("warning threshold");
    if (
      typeof warningThreshold !== "number" ||
      !Number.isFinite(warningThreshold) ||
      warningThreshold < 0 ||
      warningThreshold > 1
    ) {
      throw new Error(
        'The "Warning threshold" setting must be a number from zero through one.'
      );
    }

    return { reportingCurrency, warningThreshold };
  }

  async function findFirstQualifyingTable(excelContext) {
    const worksheets = excelContext.workbook.worksheets;
    worksheets.load("items/name");
    await excelContext.sync();

    for (const worksheet of worksheets.items) {
      const tables = worksheet.tables;
      tables.load("items/name");
      await excelContext.sync();

      for (const table of tables.items) {
        if (table.name.toLowerCase() === "settings") {
          continue;
        }

        const columns = table.columns;
        const rows = table.rows;
        columns.load("items/name");
        rows.load("count");
        await excelContext.sync();

        if (rows.count === 0) {
          continue;
        }

        const columnMap = createColumnMap(columns.items);
        if (!hasRequiredColumns(columnMap) || hasDuplicateRequiredHeaders(columns.items)) {
          continue;
        }

        const ranges = {};
        for (const requiredName of REQUIRED_QUOTE_COLUMNS) {
          const actualName = columnMap.get(normalizeHeader(requiredName));
          ranges[requiredName] = table.columns.getItem(actualName).getDataBodyRange();
          ranges[requiredName].load("values");
        }
        await excelContext.sync();

        if (hasValidQuoteData(ranges, rows.count)) {
          table.load("name");
          await excelContext.sync();
          return { table, columnNames: columnMap, rowCount: rows.count };
        }
      }
    }

    return null;
  }

  async function addMissingOutputColumns(excelContext, table, existingColumns) {
    for (const outputName of OUTPUT_COLUMNS) {
      const normalizedName = normalizeHeader(outputName);
      if (!existingColumns.has(normalizedName)) {
        table.columns.add(null, null, outputName);
        existingColumns.set(normalizedName, outputName);
      }
    }
    await excelContext.sync();
  }

  function hasValidQuoteData(ranges, rowCount) {
    for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
      const item = ranges["Item"].values[rowIndex][0];
      const vendor = ranges["Vendor"].values[rowIndex][0];
      const quoteAmount = ranges["Quote amount"].values[rowIndex][0];
      const quoteCurrency = ranges["Quote currency"].values[rowIndex][0];
      const budgetLimit = ranges["Budget limit"].values[rowIndex][0];

      if (
        typeof item !== "string" ||
        item.trim() === "" ||
        typeof vendor !== "string" ||
        vendor.trim() === "" ||
        typeof quoteAmount !== "number" ||
        !Number.isFinite(quoteAmount) ||
        quoteAmount < 0 ||
        typeof quoteCurrency !== "string" ||
        !CURRENCY_CODE_PATTERN.test(quoteCurrency.trim().toUpperCase()) ||
        typeof budgetLimit !== "number" ||
        !Number.isFinite(budgetLimit) ||
        budgetLimit <= 0
      ) {
        return false;
      }
    }
    return true;
  }

  function createColumnMap(columns) {
    const result = new Map();
    for (const column of columns) {
      const normalizedName = normalizeHeader(column.name);
      if (!result.has(normalizedName)) {
        result.set(normalizedName, column.name);
      }
    }
    return result;
  }

  function hasRequiredColumns(columnMap) {
    return REQUIRED_QUOTE_COLUMNS.every((name) => columnMap.has(normalizeHeader(name)));
  }

  function hasDuplicateRequiredHeaders(columns) {
    const counts = new Map();
    for (const column of columns) {
      const normalizedName = normalizeHeader(column.name);
      counts.set(normalizedName, (counts.get(normalizedName) || 0) + 1);
    }
    return [...REQUIRED_QUOTE_COLUMNS, ...OUTPUT_COLUMNS].some(
      (name) => (counts.get(normalizeHeader(name)) || 0) > 1
    );
  }

  function normalizeHeader(value) {
    return value.trim().toLowerCase();
  }

  function isBlank(value) {
    return value === null || value === undefined || (typeof value === "string" && value.trim() === "");
  }

  function repeatFormula(rowCount, formula) {
    return Array.from({ length: rowCount }, () => [formula]);
  }

  function repeatFormat(rowCount, format) {
    return Array.from({ length: rowCount }, () => [format]);
  }

  async function waitForFinalValues(
    excelContext,
    exchangeRates,
    convertedQuotes,
    budgetStatuses
  ) {
    const application = excelContext.workbook.application;
    application.calculate(Excel.CalculationType.full);
    await excelContext.sync();

    const deadline = Date.now() + CALCULATION_TIMEOUT_MS;
    while (Date.now() < deadline) {
      exchangeRates.load("values");
      convertedQuotes.load("values");
      budgetStatuses.load("values");
      await excelContext.sync();

      const ratesAreFinal = exchangeRates.values.every(
        ([value]) => typeof value === "number" && Number.isFinite(value) && value > 0
      );
      const quotesAreFinal = convertedQuotes.values.every(
        ([value]) => typeof value === "number" && Number.isFinite(value) && value >= 0
      );
      const statusesAreFinal = budgetStatuses.values.every(([value]) =>
        ["Within budget", "Near limit", "Over budget"].includes(value)
      );

      if (ratesAreFinal && quotesAreFinal && statusesAreFinal) {
        return;
      }
      await delay(CALCULATION_POLL_INTERVAL_MS);
    }

    throw new Error(
      "The purchase comparison formulas did not return final values within 120 seconds."
    );
  }

  function delay(milliseconds) {
    return new Promise((resolve) => setTimeout(resolve, milliseconds));
  }

  function addStatusFormat(range, status, fillColor, fontColor) {
    const format = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
    format.cellValue.rule = {
      formula1: `"${status}"`,
      operator: Excel.ConditionalCellValueOperator.equalTo,
    };
    format.cellValue.format.fill.color = fillColor;
    format.cellValue.format.font.color = fontColor;
  }
});
