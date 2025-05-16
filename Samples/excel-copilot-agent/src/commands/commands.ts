/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady(() => {});

async function insertSalesTemperatureCorrelation(xcolumn, ycolumn) {    
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");

    let dataRange = sheet.getRange("A40:A41");
    let chart = sheet.charts.add(Excel.ChartType.xyscatter, dataRange, Excel.ChartSeriesBy.auto);
    chart.name = "correlation1";
    chart.title.text = xcolumn + " and " + ycolumn + " correlation";
    chart.title.format.font.size = 12;
    chart.title.left = 8;
    chart.title.top = 8;
    chart.left = 600;
    chart.top = 250;
    chart.width = 300;
    chart.legend.visible = false;
    chart.setPosition("K16", "P30");

    let seriesCollection = chart.series;
    let series = seriesCollection.add(xcolumn + " and " + ycolumn);

    let rangeY;
    switch (ycolumn) {

        case "Date":
            rangeY = "A2:A33";
            break;
        case "Location":
            rangeY = "B2:B33";
            break;
        case "Temperature": 
            rangeY = "C2:C33";
            break;
        case "Leaflets":
            rangeY = "D2:D33";
            break;
        case "Price":
            rangeY = "E2:E33";
            break;
        case "Lemon Drink Sales":
            rangeY = "F2:F33";
            break;
        case "Orange Drink Sales":
            rangeY = "G2:G33";
            break;            
        case "Total Sales":
            rangeY = "H2:H33";
            break;
        default:
            return "I can't find a column named " + ycolumn;
    }

    let rangeX;
    switch (xcolumn) {

        case "Date":
            rangeX = "A2:A33";
            break;
        case "Location":
            rangeX = "B2:B33";
            break;
        case "Temperature": 
            rangeX = "C2:C33";
            break;
        case "Leaflets":
            rangeX = "D2:D33";
            break;
        case "Price":
            rangeX = "E2:E33";
            break;
        case "Lemon Drink Sales":
            rangeX = "F2:F33";
            chart.axes.valueAxis.displayUnit = Excel.ChartAxisDisplayUnit.thousands;
            break;
        case "Orange Drink Sales":
            rangeX = "G2:G33";
            chart.axes.valueAxis.displayUnit = Excel.ChartAxisDisplayUnit.thousands;
            break;            
        case "Total Sales":
            rangeX = "H2:H33";
            chart.axes.valueAxis.displayUnit = Excel.ChartAxisDisplayUnit.thousands;
            break;
        default:
            return "I can't find a column named " + xcolumn;
    }

    // if (ycolumn === "Temperature") {
    //   rangeY = "E2:E33";
    // } else {
    //     return "I can't find a column named " + ycolumn;
    // }
    
    let yValue = sheet.getRange(rangeY);
    series.setXAxisValues(yValue);

    // let rangeX;
    // if (xcolumn === "Total Sales") {
    //   rangeX = "H2:H33";
    //   chart.axes.valueAxis.displayUnit = Excel.ChartAxisDisplayUnit.thousands;
    // } else {
    //     return "I can't find a column named " + xcolumn;
    // }

    let value = sheet.getRange(rangeX);
    series.setValues(value);

    seriesCollection.getItemAt(0).delete();

    //chart.axes.valueAxis.displayUnit = Excel.ChartAxisDisplayUnit.thousands;

    // Show trendline.
    series.trendlines.add(Excel.ChartTrendlineType.exponential);
    let trendline = series.trendlines.getItem(0);
    trendline.showRSquared = true;

    // Highlight title.
    chart.title.getSubstring(21, 12).font.color = "#FF7F50";
    await context.sync();
  });
}

async function showSalesTemperatureCorrelation(message) {
    const {XAxisColumn: xcolumn, YAxisColumn: ycolumn} = JSON.parse(message);
    await insertSalesTemperatureCorrelation(xcolumn, ycolumn);
    return "Chart of " + ycolumn + " by " + xcolumn + " has been added.";
}

Office.actions.associate("ShowCorrelationChart", showSalesTemperatureCorrelation);
