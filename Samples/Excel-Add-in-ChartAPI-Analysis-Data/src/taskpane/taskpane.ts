/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("add-sample-data").onclick = addSampleData;
    document.getElementById("create-bar-chart").onclick = createBarChart;
    document.getElementById("reverse-plot-order").onclick = reversePlotOrder;
    document.getElementById("move-chart-title").onclick = moveChartTitle;
    document.getElementById("show-correlation-1").onclick = showCorrelation1;
    document.getElementById("highlight-top-10").onclick = hightlight1;
    document.getElementById("show-correlation-2").onclick = showCorrelation2;
    document.getElementById("highlight-top-5").onclick = hightlight2;
    document.getElementById("show-sales-trend").onclick = showSalesTrendline;
    document.getElementById("highlight-highest-sales").onclick = highlightSales;
  }
});

//Creates a bar chart based on the sample data
var index;
async function createBarChart() {
  try {
    await Excel.run(async (context) => {
      showStatus("Running", false);
      const sheet = context.workbook.worksheets.getItem("Sample");

      let dataRange = sheet.getRange("A36:C38");
      let chart = sheet.charts.add("BarClustered", dataRange, "auto");
      chart.name = "salesLocation";
      chart.title.text = "Sales in different locations";
      chart.setPosition("K2", "P15");
      chart.legend.position = "right";

      await context.sync();
      showStatus('Success for "Sales in different locations"', false);
    });
  } catch (error) {
    showStatus(error, true);
  }
}

//TODO:reverse plot order
async function reversePlotOrder() {
  try {
    await Excel.run(async (context) => {
      showStatus("Running", false);
      const sheet = context.workbook.worksheets.getItem("Sample");

      let chart = sheet.charts.getItem("salesLocation");

      chart.axes.categoryAxis.reversePlotOrder = true;
      chart.axes.valueAxis.displayUnit = "Thousands";

      await context.sync();
      showStatus('Success for "Reverse vertical axis order"', false);
    });
  } catch (error) {
    showStatus(error, true);
  }
}

async function moveChartTitle() {
  try {
    await Excel.run(async (context) => {
      showStatus("Running", false);
      const sheet = context.workbook.worksheets.getItem("Sample");

      let chart = sheet.charts.getItem("salesLocation");
      chart.legend.position = "Bottom";
      chart.legend.left = 0;
      chart.legend.top = 230;
      chart.legend.width = 80;
      chart.legend.height = 20;
      chart.title.left = 8;
      chart.title.top = 8;
      chart.title.setFormula("=Sample!L37");
      await context.sync();
      showStatus('Success for "Customize title and legend"', false);
    });
  } catch (error) {
    showStatus(error, true);
  }
}

async function showCorrelation1() {
  try {
    await Excel.run(async (context) => {
      showStatus("Running", false);
      const sheet = context.workbook.worksheets.getItem("Sample");

      let dataRange = sheet.getRange("A40:A41");
      let chart = sheet.charts.add("XYScatter", dataRange, "auto");
      chart.name = "correlation1";
      chart.title.text = "Correlation betweeen Sales and Temperature";
      chart.title.format.font.size = 12;
      chart.title.left = 8;
      chart.title.top = 8;
      chart.left = 600;
      chart.top = 250;
      chart.width = 300;
      chart.legend.visible = false;
      chart.setPosition("K16", "P30");

      let seriesCollection = chart.series;
      let series = seriesCollection.add("Sales and Temperature");

      let rangeX = "E2:E33";
      let xValue = sheet.getRange(rangeX);
      series.setXAxisValues(xValue);

      let rangeY = "H2:H33";
      let value = sheet.getRange(rangeY);
      series.setValues(value);

      seriesCollection.getItemAt(0).delete();

      chart.axes.valueAxis.displayUnit = "Thousands";

      //Show trendline
      series.trendlines.add("Exponential");
      let trendline = series.trendlines.getItem(0);
      trendline.displayRSquared = true;

      //High light title
      var font = chart.title.getSubstring(21, 5).font;
      var font2 = chart.title.getSubstring(31, 11).font;
      font.color = "#FF7F50";
      font2.color = "#FF7F50";
      await context.sync();
      showStatus('Success for "Sales/Temper correlation"', false);
    });
  } catch (error) {
    showStatus(error, true);
  }
}

async function hightlight1() {
  try {
    await Excel.run(async (context) => {
      showStatus("Running", false);
      const sheet = context.workbook.worksheets.getItem("Sample");
      let chart = sheet.charts.getItem("correlation1");

      let seriesCollection = chart.series;
      //grey out previous points
      let points = chart.series.getItemAt(0).points;
      points.load("count");
      await context.sync();
      let count = points.count;
      //grey out previous points
      for (let i = 0; i < count; i++) {
        let point = points.getItemAt(i);
        point.markerBackgroundColor = "grey";
        point.markerForegroundColor = "grey";
        await context.sync();
      }

      let series = seriesCollection.add("top5");

      let rangeX = "I2:I33";
      let xValue = sheet.getRange(rangeX);
      series.setXAxisValues(xValue);

      let rangeY = "H2:H33";
      let value = sheet.getRange(rangeY);
      series.setValues(value);

      await context.sync();
      showStatus('Success for "Highlight top 10"', false);
    });
  } catch (error) {
    showStatus(error, true);
  }
}
async function hightlight2() {
  try {
    await Excel.run(async (context) => {
      showStatus("Running", false);
      const sheet = context.workbook.worksheets.getItem("Sample");
      let chart = sheet.charts.getItem("correlation2");

      let seriesCollection = chart.series;
      let points = chart.series.getItemAt(0).points;
      points.load("count");
      await context.sync();
      let count = points.count;
      //grey out previous points
      for (let i = 0; i < count; i++) {
        let point = points.getItemAt(i);
        point.markerBackgroundColor = "grey";
        point.markerForegroundColor = "grey";
        await context.sync();
      }

      let series = seriesCollection.add("top5");

      let rangeX = "J2:J33";
      let xValue = sheet.getRange(rangeX);
      series.setXAxisValues(xValue);

      let rangeY = "H2:H33";
      let value = sheet.getRange(rangeY);
      series.setValues(value);

      await context.sync();
      showStatus('Success for "Highlight top 5"', false);
    });
  } catch (error) {
    showStatus(error, true);
  }
}

async function showCorrelation2() {
  try {
    await Excel.run(async (context) => {
      showStatus("Running", false);
      const sheet = context.workbook.worksheets.getItem("Sample");

      let dataRange = sheet.getRange("A40:A41");
      let chart = sheet.charts.add("XYScatter", dataRange, "auto");
      chart.name = "correlation2";
      chart.title.text = "Correlation betweeen Sales and Leaflets";
      chart.title.left = 8;
      chart.title.top = 8;
      chart.title.format.font.size = 12;
      chart.left = 900;
      chart.top = 250;
      chart.width = 300;
      chart.setPosition("Q2", "V15");

      chart.legend.visible = false;

      let seriesCollection = chart.series;
      let series = seriesCollection.add("Sales and Leaflets");

      let rangeX = "F2:F33";
      let xValue = sheet.getRange(rangeX);
      series.setXAxisValues(xValue);

      let rangeY = "H2:H33";
      let value = sheet.getRange(rangeY);
      series.setValues(value);

      seriesCollection.getItemAt(0).delete();

      chart.axes.valueAxis.displayUnit = "Thousands";

      //Show trendline
      series.trendlines.add("Exponential");
      let trendline = series.trendlines.getItem(0);
      trendline.displayRSquared = true;

      //High light title
      var font = chart.title.getSubstring(21, 5).font;
      var font2 = chart.title.getSubstring(31, 11).font;
      font.color = "#FF7F50";
      font2.color = "#FF7F50";

      await context.sync();
      showStatus('Success for "Sales/Leaflets correlation"', false);
    });
  } catch (error) {
    showStatus(error, true);
  }
}

//Add a ChartTitle to the first chart in the worksheet Sample
async function showSalesTrendline() {
  try {
    await Excel.run(async (context) => {
      showStatus("Running", false);
      const sheet = context.workbook.worksheets.getItem("Sample");
      let dataRange = sheet.getRange("H2:H33");
      let chart = sheet.charts.add("Line", dataRange, "auto");
      chart.activate();
      chart.name = "overallSales";
      chart.title.text = "Total Sales Trend in September";
      chart.left = 900;
      chart.top = 20;
      chart.width = 300;
      chart.setPosition("Q16", "V30");
      chart.legend.visible = false;

      let axis = chart.axes.categoryAxis;
      let yaxis = chart.axes.valueAxis;
      yaxis.displayUnit = "Thousands";
      let categoryNameRange = sheet.getRange("A2:A33");
      axis.setCategoryNames(categoryNameRange);

      let trendlines = chart.series.getItemAt(0).trendlines;
      trendlines.add("MovingAverage");
      let tre = trendlines.getItem(0);
      tre.movingAveragePeriod = 7;
      showStatus('Success for "Overall sales trend"', false);
    });
  } catch (error) {
    showStatus(error, true);
  }
}

async function clearDatalables() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Sample");
      let chart = sheet.charts.getItem("overallSales");
      let points = chart.series.getItemAt(0).points;
      points.load();
      await context.sync();
      let count = points.count;
      for (let i = 0; i < count; i++) {
        let point = points.getItemAt(i);
        point.hasDataLabel = false;
      }
      await context.sync();
    });
  } catch (error) {
    showStatus(error, true);
  }
}

//add a Axis Title to the Category Axis of the the first chart in the worksheet
async function highlightSales() {
  try {
    await Excel.run(async (context) => {
      showStatus("Running", false);
      const sheet = context.workbook.worksheets.getItem("Sample");
      let chart = sheet.charts.getItem("overallSales");
      let points = chart.series.getItemAt(0).points;
      points.load();
      await context.sync();
      let count = points.count;
      let max = 0;
      //let index = 0;
      for (let i = 0; i < count; i++) {
        let point = points.getItemAt(i);
        //point.hasDataLabel = false;
        point.load("value");
        await context.sync();
        if (point.value > max) {
          index = i;
          max = point.value;
        }
      }
      clearDatalables();
      highlight();
      await context.sync();
      showStatus('Success for "Highlight highest sales"', false);
    });
  } catch (error) {
    showStatus(error, true);
  }
}

async function highlight() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Sample");
      let chart = sheet.charts.getItem("overallSales");
      chart.load("series");
      await context.sync();
      let points = chart.series.getItemAt(0).points;
      let maxPoint = points.getItemAt(index);
      maxPoint.hasDataLabel = true;
      let datalabel = maxPoint.dataLabel;
      datalabel.showCategoryName = true;
      datalabel.showValue = true;
      datalabel.showLegendKey = true;

      maxPoint.markerStyle = "Diamond";
      await context.sync();
    });
  } catch (error) {
    showStatus(error, true);
  }
}

//add Sample Data
async function addSampleData() {
  try {
    await Excel.run(async (context) => {
      showStatus("Running", false);
      context.workbook.worksheets.getItemOrNullObject("Sample").delete();
      context.workbook.worksheets.add("Sample");
      const sheet = context.workbook.worksheets.getItem("Sample");

      let expensesTable = sheet.tables.add("A1:J1", true);
      expensesTable.name = "SalesTable";
      expensesTable.getHeaderRowRange().values = [
        [
          "Date",
          "Location",
          "Lemon",
          "Orange",
          "Temperature",
          "Leaflets",
          "Price",
          "Total Sales",
          "Top 5 Leaf",
          "Top 5 Temp",
        ],
      ];
      expensesTable.rows.add(null, [
        ["7/1/2016", "Park", 9409, 4489, 70, 90, 0.25, 0, 0, 0],
        ["7/2/2016", "Park", 9604, 4489, 72, 90, 0.25, 0, 0, 0],
        ["7/3/2016", "Park", 12100, 5929, 71, 104, 0.25, 0, 0, 0],
        ["7/4/2016", "Beach", 17956, 9801, 76, 98, 0.25, 0, 0, 0],
        ["7/5/2016", "Beach", 25281, 13924, 78, 135, 0.25, 0, 0, 0],
        ["7/6/2016", "Beach", 10609, 4761, 82, 90, 0.25, 0, 0, 0],
        ["7/6/2016", "Beach", 10609, 4761, 82, 90, 0.25, 0, 0, 0],
        ["7/7/2016", "Beach", 20449, 10201, 81, 135, 0.25, 0, 0, 0],
        ["7/8/2016", "Beach", 15129, 7396, 82, 113, 0.25, 0, 0, 0],
        ["7/9/2016", "Beach", 17956, 9025, 80, 126, 0.25, 0, 0, 0],
        ["7/10/2016", "Beach", 19600, 9604, 82, 131, 0.25, 0, 0, 0],
        ["7/11/2016", "Beach", 26244, 14400, 83, 135, 0.25, 0, 0, 0],
        ["7/12/2016", "Beach", 16900, 9025, 84, 99, 0.25, 0, 0, 0],
        ["7/13/2016", "Beach", 11881, 5625, 77, 99, 0.25, 0, 0, 0],
        ["7/14/2016", "Beach", 14884, 7225, 78, 113, 0.25, 0, 0, 0],
        ["7/15/2016", "Beach", 9604, 3844, 75, 108, 0.5, 0, 0, 0],
        ["7/16/2016", "Beach", 6561, 2500, 74, 90, 0.5, 0, 0, 0],
        ["7/17/2016", "Beach", 13225, 5776, 77, 126, 0.5, 0, 0, 0],
        ["7/18/2016", "Park", 17161, 8464, 81, 122, 0.5, 0, 0, 0],
        ["7/19/2016", "Park", 14884, 7225, 78, 113, 0.5, 0, 0, 0],
        ["7/20/2016", "Park", 5041, 1764, 70, 120, 0.5, 0, 0, 0],
        ["7/21/2016", "Park", 6889, 2500, 77, 90, 0.5, 0, 0, 0],
        ["7/22/2016", "Park", 12544, 5625, 80, 108, 0.5, 0, 0, 0],
        ["7/23/2016", "Park", 14400, 6724, 81, 117, 0.5, 0, 0, 0],
        ["7/24/2016", "Park", 14641, 6724, 82, 117, 0.5, 0, 0, 0],
        ["7/25/2016", "Park", 24336, 12769, 84, 135, 0.5, 0, 0, 0],
        ["7/26/2016", "Park", 30976, 16641, 83, 158, 0.35, 0, 0, 0],
        ["7/27/2016", "Park", 10816, 4624, 80, 99, 0.35, 0, 0, 0],
        ["7/28/2016", "Park", 9216, 3969, 82, 90, 0.35, 0, 0, 0],
        ["7/29/2016", "Park", 10000, 4356, 81, 95, 0.35, 0, 0, 0],
        ["7/30/2016", "Beach", 7744, 3249, 82, 81, 0.35, 0, 0, 0],
        ["7/31/2016", "Beach", 5776, 2209, 82, 68, 0.35, 0, 0, 0]
      ]);

      let totalSalesRange = sheet.getRange("H2:H33");
      let data = [];
      for (let i = 2; i < 34; i++) {
        let item = [];
        item.push("=C" + i.toString() + "+D" + i.toString());
        data.push(item);
      }
      totalSalesRange.formulas = data;
      totalSalesRange.format.autofitColumns();

      let top5leafRange = sheet.getRange("I2:I33");
      let data2 = [];
      for (let i = 2; i < 34; i++) {
        let item = [];
        item.push("=IF(RANK.EQ([@Temperature],[Temperature])<6,[@Temperature],NA())");
        data2.push(item);
      }
      top5leafRange.formulas = data2;
      top5leafRange.format.autofitColumns();

      let top5TempRange = sheet.getRange("J2:J33");
      let data3 = [];
      for (let i = 2; i < 34; i++) {
        let item = [];
        item.push("=IF(RANK.EQ([@Leaflets],[Leaflets])<6,[@Leaflets],NA())");
        data3.push(item);
      }
      top5TempRange.formulas = data3;
      top5TempRange.format.autofitColumns();

      const range1 = sheet.getRange("B36");
      range1.formulas = [["=C1"]];
      range1.format.autofitColumns();

      const range2 = sheet.getRange("C36");
      range2.formulas = [["=D1"]];
      range2.format.autofitColumns();

      const range3 = sheet.getRange("A37");
      range3.formulas = [["=B2"]];
      range3.format.autofitColumns();

      const range4 = sheet.getRange("A38");
      range4.formulas = [["=B5"]];
      range4.format.autofitColumns();

      const range11 = sheet.getRange("B37");
      range11.formulas = [['=SUMIF($B$2:$B$33,"=Park",$C$2:$C$33)']];
      range11.format.autofitColumns();

      const range12 = sheet.getRange("C37");
      range12.formulas = [['=SUMIF($B$2:$B$33,"=Park",$D$2:$D$33)']];
      range12.format.autofitColumns();

      const range21 = sheet.getRange("B38");
      range21.formulas = [['=SUMIF($B$2:$B$33,"=Beach",$C$2:$C$33)']];
      range21.format.autofitColumns();

      const range22 = sheet.getRange("C38");
      range22.formulas = [['=SUMIF($B$2:$B$33,"=Beach",$D$2:$D$33)']];
      range22.format.autofitColumns();

      const range33 = sheet.getRange("L36");
      range33.formulas = [["=MAX(H2:H33)"]];
      range33.format.autofitColumns();

      const range36 = sheet.getRange("L37");
      range36.formulas = [["=AVERAGE(H2:H33)"]];
      range36.format.autofitColumns();

      const range34 = sheet.getRange("M36");
      range34.formulas = [["=INDEX(A2:A33,MATCH(L36,H2:H33,0),0)"]];
      range34.format.autofitColumns();

      let dateRange = sheet.getRange("A2:A33");

      let formatdate = [];
      let formatdateitem = ["m/d"];
      for (let i = 0; i < 32; i++) {
        formatdate.push(formatdateitem);
      }
      dateRange.numberFormat = formatdate;

      let numRange1 = sheet.getRange("C2:C33");
      let numRange2 = sheet.getRange("D2:D33");
      let numRange3 = sheet.getRange("H2:H33");
      let formatnumber = [];
      let formatnumberitem = ["###,0"];
      for (let i = 0; i < 32; i++) {
        formatnumber.push(formatnumberitem);
      }
      numRange1.numberFormat = formatnumber;
      numRange2.numberFormat = formatnumber;
      numRange3.numberFormat = formatnumber;

      if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
      }
      sheet.activate();
      sheet.gridlines = false;

      sheet.onChanged.add(onDataChanged);
      await context.sync();
      showStatus('Success for "Add Sample data"', false);
    });
  } catch (error) {
    showStatus(error, true);
  }
}

async function onDataChanged(event) {
  try {
    await Excel.run(async (context) => {
      console.log("data changed");
      highlightSales();

      await context.sync();
    });
  } catch (error) {
    showStatus(error, true);
  }
}

function showStatus(message, isError) {
  let status = document.getElementById("status");
  // Clear previous content
  status.innerHTML = "";

  // Create the container div
  let statusCard = document.createElement("div");
  statusCard.className = `status-card ms-depth-4 ${isError ? "error-msg" : "success-msg"}`;

  // Create and append the first paragraph
  let p1 = document.createElement("p");
  p1.className = "ms-fontSize-24 ms-fontWeight-bold";
  p1.textContent = isError ? "An error occurred" : "";
  statusCard.appendChild(p1);

  // Create and append the second paragraph
  let p2 = document.createElement("p");
  p2.className = "ms-fontSize-16 ms-fontWeight-regular";
  p2.textContent = message;
  statusCard.appendChild(p2);

  // Append the status card to the status element
  status.appendChild(statusCard);
}