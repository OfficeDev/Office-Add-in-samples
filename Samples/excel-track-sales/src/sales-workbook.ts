/* global Excel */

import { HEADERS, METRICS, Sale, toTableRows, validateDate } from "./sales-data";

const SHEET_NAME = "Sales Tracker";
const TABLE_NAME = "SalesTrackerData";
const CHART_NAME = "SalesTrackerChart";

export const TABLE_STYLES = {
  blue: "TableStyleMedium2",
  green: "TableStyleMedium4",
  orange: "TableStyleMedium9",
};

export const CHART_TYPES = {
  line: "Line",
  column: "ColumnClustered",
  area: "Area",
} as const;

export const CHART_THEMES = {
  light: { background: "#FFFFFF", foreground: "#242424", series: "#0078D4" },
  dark: { background: "#243746", foreground: "#FFFFFF", series: "#6CB8FF" },
  blue: { background: "#005A9E", foreground: "#FFFFFF", series: "#FFFFFF" },
};

export interface ChartOptions {
  metric: string;
  type: keyof typeof CHART_TYPES;
  theme: keyof typeof CHART_THEMES;
}

function getTable(context: Excel.RequestContext): Excel.Table {
  return context.workbook.worksheets.getItem(SHEET_NAME).tables.getItem(TABLE_NAME);
}

function validateChartOptions(options: ChartOptions): void {
  if (
    !METRICS.includes(options.metric) ||
    !CHART_TYPES[options.type] ||
    !CHART_THEMES[options.theme]
  ) {
    throw new Error("Choose a valid chart metric, type, and theme.");
  }
}

async function configureChart(
  context: Excel.RequestContext,
  chart: Excel.Chart,
  table: Excel.Table,
  options: ChartOptions
): Promise<void> {
  validateChartOptions(options);
  chart.chartType = CHART_TYPES[options.type];
  chart.series.load("items");
  await context.sync();
  const extraSeries = chart.series.items.slice(1);
  if (extraSeries.length > 0) {
    extraSeries.reverse().forEach((series) => series.delete());
    chart.series.load("items");
    await context.sync();
  }
  const colors = CHART_THEMES[options.theme];
  chart.title.text = `${options.metric} by date`;
  chart.legend.visible = false;
  chart.plotVisibleOnly = false;
  chart.format.fill.setSolidColor(colors.background);
  chart.title.format.font.color = colors.foreground;
  chart.axes.categoryAxis.format.font.color = colors.foreground;
  chart.axes.valueAxis.format.font.color = colors.foreground;
  chart.axes.categoryAxis.categoryType = Excel.ChartAxisCategoryType.dateAxis;
  chart.axes.categoryAxis.numberFormat = "mmm d";
  const series = chart.series.items[0] || chart.series.add(options.metric);
  series.name = options.metric;
  series.setXAxisValues(table.columns.getItem("Date").getDataBodyRange());
  series.setValues(table.columns.getItem(options.metric).getDataBodyRange());
  series.format.line.color = colors.series;
  if (options.type === "line") {
    series.markerStyle = Excel.ChartMarkerStyle.circle;
    series.markerBackgroundColor = colors.series;
    series.markerForegroundColor = colors.series;
  } else {
    series.format.fill.setSolidColor(colors.series);
  }
}

export async function hasSalesTable(): Promise<boolean> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAME);
    await context.sync();
    if (sheet.isNullObject) {
      return false;
    }
    const table = sheet.tables.getItemOrNullObject(TABLE_NAME);
    await context.sync();
    return !table.isNullObject;
  });
}

export async function importSales(sales: Sale[], options: ChartOptions): Promise<void> {
  if (sales.length === 0) {
    throw new Error("No sales match this filter. The workbook has not been changed.");
  }
  validateChartOptions(options);
  await Excel.run(async (context) => {
    const dates = sales.map((sale) => {
      validateDate(sale.date);
      const [year, month, day] = sale.date.split("-").map(Number);
      const result = context.workbook.functions.date(year, month, day);
      result.load("value,error");
      return result;
    });
    let sheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAME);
    const namedTable = context.workbook.tables.getItemOrNullObject(TABLE_NAME);
    await context.sync();
    const rows = toTableRows(
      sales,
      dates.map((result) => {
        if (result.error) {
          throw new Error(`Could not convert a sales date to an Excel date: ${result.error}`);
        }
        return result.value;
      })
    );
    let table: Excel.Table;
    if (sheet.isNullObject) {
      if (!namedTable.isNullObject) {
        throw new Error(
          `A table named ${TABLE_NAME} already exists elsewhere. Rename it before importing.`
        );
      }
      sheet = context.workbook.worksheets.add(SHEET_NAME);
      sheet.getRange("A1:G1").values = [HEADERS];
      sheet.getRange(`A2:G${rows.length + 1}`).values = rows;
      table = sheet.tables.add(`A1:G${rows.length + 1}`, true);
      table.name = TABLE_NAME;
    } else {
      table = sheet.tables.getItemOrNullObject(TABLE_NAME);
      await context.sync();
      if (table.isNullObject) {
        throw new Error(
          `The "${SHEET_NAME}" worksheet is not owned by this add-in. Rename it before importing.`
        );
      }
      const range = table.getRange();
      range.load("address");
      table.load("showTotals");
      await context.sync();
      if (!/!\$?A\$?1:\$?G\$?\d+$/.test(range.address)) {
        throw new Error(
          "The sales table has been moved or its columns changed. Restore it to A1:G before importing."
        );
      }
      if (table.showTotals) {
        throw new Error("Turn off the sales table's total row before importing.");
      }
      table.getDataBodyRange().clear(Excel.ClearApplyTo.contents);
      table.resize(`A1:G${rows.length + 1}`);
      await context.sync();
      table.getHeaderRowRange().values = [HEADERS];
      table.getDataBodyRange().values = rows;
    }
    table.style = TABLE_STYLES.blue;
    table.getRange().columnHidden = false;
    table.columns.getItem("Date").getDataBodyRange().numberFormat = [["mmm d, yyyy"]];
    table.columns.getItem("Units").getDataBodyRange().numberFormat = [["0"]];
    sheet.getRange(`C2:G${rows.length + 1}`).numberFormat = [['"$"#,##0.00']];
    table.getRange().format.autofitColumns();
    sheet.activate();
    sheet.getRange("A1").select();
    const chart = sheet.charts.getItemOrNullObject(CHART_NAME);
    await context.sync();
    const salesChart = chart.isNullObject
      ? sheet.charts.add(
          CHART_TYPES[options.type],
          table.columns.getItem(options.metric).getDataBodyRange(),
          Excel.ChartSeriesBy.columns
        )
      : chart;
    salesChart.name = CHART_NAME;
    salesChart.setPosition("I2", "Q20");
    await configureChart(context, salesChart, table, options);
    await context.sync();
  });
}

export async function formatTable(
  style: keyof typeof TABLE_STYLES,
  visibleColumns: string[]
): Promise<void> {
  if (!TABLE_STYLES[style] || visibleColumns.some((column) => !HEADERS.includes(column))) {
    throw new Error("Choose valid table settings.");
  }
  await Excel.run(async (context) => {
    const table = getTable(context);
    table.style = TABLE_STYLES[style];
    HEADERS.forEach((header) => {
      table.columns.getItem(header).getRange().columnHidden = !visibleColumns.includes(header);
    });
    await context.sync();
  });
}

export async function formatChart(options: ChartOptions): Promise<void> {
  await Excel.run(async (context) => {
    const table = getTable(context);
    const chart = context.workbook.worksheets.getItem(SHEET_NAME).charts.getItem(CHART_NAME);
    await configureChart(context, chart, table, options);
    await context.sync();
  });
}
