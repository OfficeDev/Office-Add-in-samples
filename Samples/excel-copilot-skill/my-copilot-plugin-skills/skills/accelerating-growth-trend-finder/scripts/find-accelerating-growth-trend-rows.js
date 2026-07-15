await Excel.run(async (context) => {
    const firstWorksheet = context.workbook.worksheets.getFirst();

    const tables = firstWorksheet.tables;
    tables.load("items");
    await context.sync();

    const tableInfo = await findFirstQualifyingTable(tables.items);

    if (!tableInfo) {
        return;
    }

    const {
        table: matchingTable,
        tableRange: matchingTableRange,
        dataRange: matchingDataRange
    } = tableInfo;

     let createdCharts = [];

    // Review each row in the qualifying table's data body.
    for (let rowOffset = 0; rowOffset < matchingDataRange.values.length; rowOffset += 1) {
      const row = matchingDataRange.values[rowOffset];

      // If the row does not have an accelerating growth trend and there's another row, 
      // check the next row, otherwise end the row checking loop without creating a chart.
      if (!hasAcceleratingGrowthTrend(row)) {
        continue;
      }

      // The row has an accelerating growth trend, so create a chart for it.
      await createAndPositionChart(firstWorksheet, matchingTable, matchingTableRange, matchingDataRange, rowOffset);

      createdCharts.push({
            tableName: matchingTable.name,
            worksheetRowNumber: matchingDataRange.rowIndex + rowOffset + 1
      });
    }

    await context.sync();

    if (createdCharts.length === 0) {
       return "No rows with an accelerating growth trend were found.";
    }

    return;

    /**
     * Finds the first table in the provided array of tables that qualifies for charting.
     *
     * A qualifying table has:
     * - At least 12 columns.
     * - At least one data row.
     * - All numeric values in the data body.
     */
    async function findFirstQualifyingTable(tables) {
        for (const table of tables) {
            const tableRange = table.getRange();
            const rows = table.rows;

            table.load("name");
            tableRange.load(["columnCount", "rowIndex", "columnIndex", "rowCount"]);
            rows.load("count");

            await context.sync();

            if (tableRange.columnCount < 12 || rows.count === 0) {
                continue;
            }

            const dataRange = table.getDataBodyRange();

            dataRange.load([ "values", "rowIndex", "rowCount", "columnCount"]);

            await context.sync();

            const isEntirelyNumeric = dataRange.values.every((row) =>
                row.every(
                    (value) =>
                        typeof value === "number" &&
                        Number.isFinite(value)
                )
            );

            if (isEntirelyNumeric) {
                return { table,  tableRange, dataRange };
            }
        }

        return null;
    }


    /**
     * Determines whether a row has an accelerating growth trend.
     */
    function hasAcceleratingGrowthTrend(row) {
        let previousIncrease = null;

        for (let columnIndex = 1; columnIndex < row.length; columnIndex += 1) {
            const increase = row[columnIndex] - row[columnIndex - 1];

            if (increase <= 0 || (previousIncrease !== null && increase <= previousIncrease)) {
                return false;
            }
            previousIncrease = increase;
        }

        return true;
    }

    /**
     * Converts a zero-based column index to Excel column letters.
     *
     * Example:
     *   0 -> A
     *   1 -> B
     *   25 -> Z
     *   26 -> AA
     */
    function columnIndexToLetters(columnIndex) {
        let letters = "";
        let n = columnIndex + 1;

        while (n > 0) {
            const remainder = (n - 1) % 26;
            letters = String.fromCharCode(65 + remainder) + letters;
            n = Math.floor((n - 1) / 26);
        }

        return letters;
    }

    /**
     * Creates and positions a line chart for one matching table row.
     *
     * The chart is:
     *   - Based on the full table range.
     *   - Reduced to only the series for the matching row.
     *   - Positioned one column to the right of the table.
     *   - Aligned vertically with the matching data row.
     *   - Sized to 7 columns wide by 15 rows tall.
     */
    async function createAndPositionChart(worksheet, table, tableRange, dataRange, rowOffset) {

        // The worksheet row number, one-based, for the matching data row.
        const worksheetRowNumber = dataRange.rowIndex + rowOffset + 1;

        // Create a line chart from the full table range.
        // ChartSeriesBy.rows means each data row becomes a chart series,
        // while the table headers provide the X-axis category labels.
        const chart = worksheet.charts.add(
            Excel.ChartType.line,
            tableRange,
            Excel.ChartSeriesBy.rows
        );

        chart.series.load("count");
        await context.sync();

        // Remove every series except the one for the matching row.
        // The rowOffset corresponds to the matching data row's series.
        for (let seriesIndex = chart.series.count - 1; seriesIndex >= 0; seriesIndex -= 1) {
            if (seriesIndex !== rowOffset) {
            chart.series.getItemAt(seriesIndex).delete();
            }
        }

        // Chart title:
        //   <table name> - <worksheet row number>
        chart.title.text = `${table.name} - Row ${worksheetRowNumber}`;
        chart.title.visible = true;

        // The only chart elements should be axes and gridlines.
        // The title is also shown because you requested a chart title.
        chart.legend.visible = false;

        chart.axes.categoryAxis.visible = true;
        chart.axes.valueAxis.visible = true;

        chart.axes.valueAxis.majorGridlines.visible = true;
        chart.axes.valueAxis.minorGridlines.visible = false;

        chart.dataLabels.showValue = false;
        chart.dataLabels.showCategoryName = false;
        chart.dataLabels.showSeriesName = false;

        // Position the chart one column to the right of the table.
        const chartStartColumn = tableRange.columnIndex + tableRange.columnCount + 1;

        // Align the chart's top edge with the matching data row.
        const chartStartRow = dataRange.rowIndex + rowOffset;

        // Make the chart 7 columns wide and 15 rows tall.
        const chartEndColumn = chartStartColumn + 6;
        const chartEndRow = chartStartRow + 14;

        const startCell = `${columnIndexToLetters(chartStartColumn)}${chartStartRow + 1}`;
        const endCell = `${columnIndexToLetters(chartEndColumn)}${chartEndRow + 1}`;

        chart.setPosition(startCell, endCell);
    }
});




