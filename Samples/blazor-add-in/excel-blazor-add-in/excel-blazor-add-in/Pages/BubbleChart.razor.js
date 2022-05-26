/* Copyright(c) Microsoft. All rights reserved. Licensed under the MIT License. */
/* 
 * This sample is extracted from the excellent helper tool called Script Lab
 * Take a look at more samples provided there
 */

export async function createBubbleChart() {
    await Excel.run(async (context) => {
        /*
          The table is expected to look like this:
          Product, Inventory, Price, Current Market Share
          Calamansi, 2000, $2.45, 10%
          ...
    
          We want each bubble to represent a single row.
        */

        // Get the worksheet and table data.
        const sheet = context.workbook.worksheets.getItem("Sample");
        const table = sheet.tables.getItem("Sales");
        const dataRange = table.getDataBodyRange();

        // Get the table data without the row names.
        const valueRange = dataRange.getOffsetRange(0, 1).getResizedRange(0, -1);

        // Create the chart.
        const bubbleChart = sheet.charts.add(Excel.ChartType.bubble, valueRange);
        bubbleChart.name = "Product Chart";

        // Remove the default series, since we want a unique series for each row.
        bubbleChart.series.getItemAt(0).delete();

        // Load the data necessary to make a chart series.
        dataRange.load(["rowCount", "values"]);
        await context.sync();

        // For each row, create a chart series (a bubble).
        for (let i = 0; i < dataRange.rowCount; i++) {
            const newSeries = bubbleChart.series.add(dataRange.values[i][0], i);
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

export async function creatTable() {
    await Excel.run(async (context) => {
        context.workbook.worksheets.getItemOrNullObject("Sample").delete();
        const sheet = context.workbook.worksheets.add("Sample");

        let inventoryTable = sheet.tables.add("A1:D1", true);
        inventoryTable.name = "Sales";
        inventoryTable.getHeaderRowRange().values = [["Product", "Inventory", "Price", "Current Market Share"]];

        inventoryTable.rows.add(null, [
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

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
    try {
        await callback();
    } catch (error) {
        // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
        console.error(error);
    }
}
