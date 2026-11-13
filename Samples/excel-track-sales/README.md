---
title: Import and visualize sales data in Excel
page_type: sample
urlFragment: excel-sales-tracker
products:
  - office-excel
  - office
  - m365
languages:
  - typescript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: "11/12/2026 10:00:00 AM"
description: "Import bundled demo sales data into Excel tables and customize charts using the Excel JavaScript API."
---

# Import and visualize sales data in Excel

Explore sales trends for a product over a selected date range in Excel. This sample imports demo sales data into a worksheet, lets you customize the table, and charts a selected sales metric. It demonstrates:

- Importing bundled JSON demo data into an Excel table.
- Filtering sales by product and an inclusive start and end date range.
- Changing table styles and column visibility.
- Visualizing one selected sales metric in a line, clustered column, or area chart.
- Updating an existing table and chart and activating the result worksheet.

## Applies to

- Excel on Windows
- Excel on Mac
- Excel on the web

The sample requires **ExcelApi 1.9** and uses a unified manifest for Microsoft 365. To check API availability across supported clients, see [Excel JavaScript API requirement sets](https://learn.microsoft.com/javascript/api/requirement-sets/excel/excel-api-requirement-sets). Current manifest support requires Excel for Windows version 2501 or later, or Excel for Mac version 16.103 or later. To check manifest compatibility across clients, see [unified manifest platform support](https://learn.microsoft.com/office/dev/add-ins/develop/unified-manifest-overview#client-and-platform-support).

## Prerequisites

- A current LTS version of [Node.js](https://nodejs.org/) and [npm](https://www.npmjs.com/get-npm).
- A supported Excel installation connected to a Microsoft 365 account. Get a [free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) that provides a renewable 90-day Microsoft 365 E5 developer subscription.

## Run the sample

1. Download or clone the project and open its root folder in a terminal.
1. Install the dependencies.

   ```console
   npm install
   ```

1. Start the development server and sideload the add-in.

   ```console
   npm start
   ```

   Trust the development certificate if prompted.

1. In Excel, choose **Track sales** from the ribbon to open the add-in's task pane.
1. In the task pane, select a product and a date range. Both endpoints are included, and single-day ranges are supported.
1. Choose **Get sales data**. The add-in creates or refreshes the **Sales Tracker** worksheet's table and chart, then activates the worksheet.

   > **Note**: Each import resets the table style and column visibility. The sample updates its named table and chart rather than creating duplicates.
1. To format the table, choose the data columns and color of the table, then choose **Apply table settings**.
1. To format the chart, choose a data source, chart type, and chart theme. Then, choose **Apply chart settings**.
1. When finished testing the add-in, stop the web server and uninstall the add-in.

   ```console
   npm stop
   ```

## Key parts of the sample

The sample follows a sales-analysis workflow: import the selected records, format the table, and visualize a metric in a chart. The following sections explain how Office.js supports each step.

### Import data into an Excel table

After the user selects a product and date range, the sample prepares the matching sales records for Excel. It first converts the source dates to numeric Excel dates with [`Workbook.functions.date`](https://learn.microsoft.com/javascript/api/excel/excel.functions), respecting the workbook's date system. It then writes the records with [`Range.values`](https://learn.microsoft.com/javascript/api/excel/excel.range) and organizes them into a table with [`TableCollection.add`](https://learn.microsoft.com/javascript/api/excel/excel.tablecollection).

Subsequent imports update the same table. [`Table.resize`](https://learn.microsoft.com/javascript/api/excel/excel.table#excel-excel-table-resize-member(1)) adjusts its size to the new selection, and [`Table.getDataBodyRange`](https://learn.microsoft.com/javascript/api/excel/excel.table#excel-excel-table-getdatabodyrange-member(1)) provides the range for writing the replacement records.

To learn more about table creation, resizing, and formatting, see [Create, read, and manage Excel tables with JavaScript](https://learn.microsoft.com/office/dev/add-ins/excel/excel-add-ins-tables).

### Format tables and control column visibility

Once the records are in the table, formatting makes them easier to read. The sample combines a built-in style through [`Table.style`](https://learn.microsoft.com/javascript/api/excel/excel.table) with [`Range.numberFormat`](https://learn.microsoft.com/javascript/api/excel/excel.range) for dates, unit counts, and currency amounts. [`RangeFormat.autofitColumns`](https://learn.microsoft.com/javascript/api/excel/excel.rangeformat) then sizes the columns to fit their contents.

Users can also hide columns to focus on the information they need. The sample obtains each column's range with [`TableColumn.getRange`](https://learn.microsoft.com/javascript/api/excel/excel.tablecolumn) and sets [`Range.columnHidden`](https://learn.microsoft.com/javascript/api/excel/excel.range). This hides the corresponding worksheet column, not just its cells within the table.

### Create and customize a chart

The chart turns the imported records into a view of the selected metric over time. The sample creates it with [`ChartCollection.add`](https://learn.microsoft.com/javascript/api/excel/excel.chartcollection), then uses [`ChartSeries.setXAxisValues` and `ChartSeries.setValues`](https://learn.microsoft.com/javascript/api/excel/excel.chartseries) to connect the series to the table's dates and metric values.

Users can change the presentation without replacing the chart: [`Chart.chartType`](https://learn.microsoft.com/javascript/api/excel/excel.chart) switches between line, clustered column, and area visualizations, while [chart formatting](https://learn.microsoft.com/javascript/api/excel/excel.chartformat) and [series formatting and marker properties](https://learn.microsoft.com/javascript/api/excel/excel.chartseries) apply the color themes. Setting [`Chart.plotVisibleOnly = false`](https://learn.microsoft.com/javascript/api/excel/excel.chart) also keeps the metric plotted when its worksheet column is hidden.

To explore chart series and formatting patterns, see [Create and customize Excel charts with JavaScript](https://learn.microsoft.com/office/dev/add-ins/excel/excel-add-ins-charts).

### Work with the workbook asynchronously

All of these steps use the same asynchronous workbook model. Within [`Excel.run`](https://learn.microsoft.com/javascript/api/excel), the sample queues reads with object [`load`](https://learn.microsoft.com/javascript/api/office/officeextension.clientobject) calls and submits queued reads and updates through [`context.sync`](https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext).

To support repeated imports, it checks the [worksheet](https://learn.microsoft.com/javascript/api/excel/excel.worksheetcollection), [table](https://learn.microsoft.com/javascript/api/excel/excel.tablecollection), and [chart](https://learn.microsoft.com/javascript/api/excel/excel.chartcollection) collections with `getItemOrNullObject` before creating sample objects. After populating the table, [`Worksheet.activate`](https://learn.microsoft.com/javascript/api/excel/excel.worksheet) and [`Range.select`](https://learn.microsoft.com/javascript/api/excel/excel.range) bring the results into view.

For the workbook operations, see [sales-workbook.ts](src/sales-workbook.ts). The entry point in [taskpane.ts](src/taskpane.ts) waits for [`Office.onReady`](https://learn.microsoft.com/javascript/api/office) and checks the required API support with [`Office.context.requirements.isSetSupported`](https://learn.microsoft.com/javascript/api/office/office.requirementsets) before enabling the workflow.

To understand batching, loading properties, and synchronization, see [Using the application-specific API model](https://learn.microsoft.com/office/dev/add-ins/develop/application-specific-api-model).

## For VSTO developers

Office Add-ins extend Excel on Windows, Mac, and the web, whereas VSTO add-ins run only on Windows. Familiar operations, such as styling tables, hiding columns, and changing chart types, use the Excel JavaScript object model. Unlike direct calls in VSTO, workbook operations are queued and submitted asynchronously.

The [Key parts of the sample](#key-parts-of-the-sample) section explains these operations and links to their APIs. For broader guidance, see the [VSTO add-in developer's guide to Office Web Add-ins](https://learn.microsoft.com/office/dev/add-ins/overview/learning-path-transition).

For a corresponding C# implementation, explore the archived [Tables and Charts VSTO sample](https://github.com/microsoftarchive/msdn-code-gallery-community-s-z/tree/master/VSTO%20-%20Generate%20tables%20and%20charts%20by%20using%20an%20Excel%20add-in). Use it as an object-model comparison rather than current setup guidance.

## Solution

| Solution | Author(s) |
| --- | --- |
| Import and visualize sales data in Excel | Microsoft |

## Version history

| Version | Date | Comments |
| --- | --- | --- |
| 1.0 | November 12, 2026 | Initial release. |

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2026 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/excel-sales-tracker" />
