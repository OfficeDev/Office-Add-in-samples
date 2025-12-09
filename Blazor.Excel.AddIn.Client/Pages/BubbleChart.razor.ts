/* Copyright(c) Microsoft. All rights reserved. Licensed under the MIT License. */
/* 
 * This bubble chart sample is extracted from the excellent helper tool called Script Lab.
 * Find this sample and more at https://aka.ms/getscriptlab.
 */
console.log("Loading BubbleChart.razor.js");

/**
 * Creates a sample table with product inventory data in Excel.
 * Deletes any existing "Sample" worksheet and creates a new one with sales data.
 * 
 * @returns A promise that resolves when the table is created
 */
export async function createTable(): Promise<void> {
  await Excel.run(async (context: Excel.RequestContext): Promise<void> => {
    context.workbook.worksheets.getItemOrNullObject("Sample").delete();
    const sheet: Excel.Worksheet = context.workbook.worksheets.add("Sample");

    const inventoryTable: Excel.Table = sheet.tables.add("A1:D1", true);
    inventoryTable.name = "Sales";
    inventoryTable.getHeaderRowRange().values = [["Product", "Inventory", "Price", "Current Market Share"]];

    inventoryTable.rows.add(undefined, [
      ["Calamansi", 2000, "$2.45", "10%"],
      ["Cara cara orange", 10000, "$2.12", "45%"],
      ["Limequat", 4000, "$0.70", "66%"],
      ["Meyer lemon", 100, "$2.65", "5%"],
      ["Pomelo", 4000, "$1.69", "14%"],
      ["Yuzu", 7500, "$3.23", "34%"]
    ]);

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    sheet.activate();
    await context.sync();
  });
}

/**
 * Creates a bubble chart from the Sales table data.
 * Each bubble represents a product with inventory, price, and market share.
 * 
 * The table must have the following structure:
 * - Column 0: Product (used for series names)
 * - Column 1: Inventory (X-axis values)
 * - Column 2: Price (Y-axis values)
 * - Column 3: Current Market Share (bubble sizes)
 * 
 * @returns A promise that resolves when the bubble chart is created
 */
export async function createBubbleChart(): Promise<void> {
  await Excel.run(async (context: Excel.RequestContext): Promise<void> => {
    /*
      The table has the following columns and data.
      Product, Inventory, Price, Current Market Share
      Calamansi, 2000, $2.45, 10%
      ...
 
      Each bubble represents a single row.
    */

    // Get the worksheet and table data.
    const sheet: Excel.Worksheet = context.workbook.worksheets.getItem("Sample");
    const table: Excel.Table = sheet.tables.getItem("Sales");
    const dataRange: Excel.Range = table.getDataBodyRange();

    // Get the table data without the row names.
    const valueRange: Excel.Range = dataRange.getOffsetRange(0, 1).getResizedRange(0, -1);

    // Create the chart.
    const bubbleChart: Excel.Chart = sheet.charts.add(Excel.ChartType.bubble, valueRange);
    bubbleChart.name = "Product Chart";

    // Remove the default series, since we want a unique series for each row.
    bubbleChart.series.getItemAt(0).delete();

    // Load the data necessary to make a chart series.
    dataRange.load(["rowCount", "values"]);
    await context.sync();

    // For each row, create a chart series (a bubble).
    const rowCount: number = dataRange.rowCount;
    const values: (string | number | boolean)[][] = dataRange.values as (string | number | boolean)[][];
    
    for (let i = 0; i < rowCount; i++) {
      const rowValues: (string | number | boolean)[] | undefined = values[i];
      if (!rowValues) {
        continue;
      }
      
      const productName: string | number | boolean | undefined = rowValues[0];
      const seriesName: string = productName !== undefined ? productName.toString() : "";
      const newSeries: Excel.ChartSeries = bubbleChart.series.add(seriesName, i);
      newSeries.setXAxisValues(dataRange.getCell(i, 1));
      newSeries.setValues(dataRange.getCell(i, 2));
      newSeries.setBubbleSizes(dataRange.getCell(i, 3));

      // Show the product name and market share percentage.
      newSeries.dataLabels.showSeriesName = true;
      newSeries.dataLabels.showBubbleSize = true;
      newSeries.dataLabels.showValue = false;
    }

    await context.sync();
  });
}
