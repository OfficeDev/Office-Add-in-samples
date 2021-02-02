---
page_type: sample
products:
- office-excel
- office-powerpoint
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 02/02/2021 1:25:00 PM
description: "Learn how to create a contextual tab that displays on the ribbon in response to the context of the Office UI."
---

# (Preview) Create custom contextual tabs on the ribbon

This sample shows how to create a custom contextual tab on the ribbon in the Office UI. The sample creates a table, and when the user moves the focus inside the table, the custom tab is displayed. When the user moves outside the table, the custom tab is hidden.

> **Note:** The features used in this sample are currently in preview and subject to change. They are not currently supported for use in production environments. To try the preview features, you'll need to [join Office Insider](https://insider.office.com/join). A good way to try out preview features is to sign up for an Office 365 subscription. If you don't already have an Office 365 subscription, get one by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).

## Applies to

-  Excel and PowerPoint on Windows 10.

## Prerequisites

To use this sample, you'll need to [join Office Insider](https://insider.office.com/join).

Before running this sample, you need a recent version of [npm](https://www.npmjs.com/get-npm) and [Node.js](https://nodejs.org/en/) installed on your computer. To verify if you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.

## Solution

Solution | Author(s)
---------|----------
Create custom contextual tabs on the ribbon | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.0  | February 2, 2021 | Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

## Scenario: Create and use a contextual tab

This sample inserts a table of fictitious sales data for Contoso. The data is pulled from one of two mock data sources; a mock Excel file, or a mock SQL database. The user can select which data source to use either in the task pane, or in the contextual tab.

After the sales table is created, the sample creates a contextual tab named **Table Data**. When you select any cell or range inside the table, the contextual tab is displayed on the ribbon. When you select any cell or range outside the table, the contextual tab is hidden.

The contextual tab supports commands related to working with the sales data. When you make changes you can submit them and update the mock data source. Or you can refresh the table from the mock data source.

## Build and run the sample

1. Clone or download this repository.
2. In the command line, go to the **office-contextual-tabs** folder from your root directory.
3. Run the following command to download the dependencies required to run the sample.
    
    ```command&nbsp;line
    $ npm install
    
4. Run the following command to start the localhost web server, and sideload the sample add-in to Excel on Windows.
    
    ```command&nbsp;line
    $ npm run start
    
You can take the following actions to try out the add-in and the contextual tab.

- Use the task pane to import data from either the the mock Excel file, or mock SQL Database. Selecting **Import data** in the task pane creates the sales table.
- Select a cell, or range, inside the sales table to display the **Table Data** contextual tab on the ribbon.
- On the **Table Data** contextual tab you can:
    - **Submit** changes you made in the table to the mock data source.
    - **Refresh** the table from the mock data source which overwrites any changes you made.
    - Use **Show task pane** to show the add-in's task pane if it was closed.
    - **Import data** and update the table to use data from the mock Excel file or mock SQL database.
- Select a cell, or range, outside the sales table to hide the **Table Data** contextual tab.

## Key parts of this sample

### Shared JavaScript runtime

Contextual tabs requires the manifestl.xml file to specify loading the shared JavaScript runtime. For more information on configuring the shared runtime, see [Configure your Office Add-in to use a shared JavaScript runtime](https://docs.microsoft.com/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime)

### Describe and update the contextual tab using JSON

The `src/commands/ribbonJSON.js` file describes the full contextual tab's buttons, groups, and menu items. It returns the JSON from the `getContextualRibbonJSON()` function and the sample code stores this JSON in a global variable. 

### Update the contextual tab UI visibility

As buttons, or the tab itself are set to be visible or not, the `g.contextualTab` global variable is used to always maintain the correct contextual tab state. When the tab needs to be updated on the ribbon (to turn a UI element on or off) a call is made to `Office.ribbon.requestUpdate()`. This occurs in the `setContextualTabVisibility()` function in `src/utilities/utilities.js`.

### Handle context changes

When you build your own add-in, you'll need to decide what context determines which tabs or UI elements are shown. For this sample, we want to show the contextual tab when the focus is in the table. Also we want to enable the **Refresh** and **Submit** buttons when changes are made to the table.

When the user imports data to create the table, `createSampleWorkSheet()` adds a `onSelectionChanged` event handler, and `onChanged` event handler. Later, as the user moves the selection into or out of the table, the onSelectinChanged() function is called, which can display the contextual tab when the selection is inside the table. When the user makes changes to the table, `onSelectionChange()` is called, and the **Refresh** and **Submit** buttons are enabled.

You can see more details in the following code excerpt, or refer to these functions in the `src/utilities/utilities.js` file.

```javascript
export async function createSampleWorkSheet(mockDataSource) {
    //...//
    salesTable.onSelectionChanged.add(onSelectionChange);
    //...//
    salesTable.onChanged.add(onChanged);
    //...//
}

/**
 * Handles the onSelectionChange event. If selection is inside the table, the Contoso custom tab is shown.
 * Otherwise the Contoso custom tab is hidden.
 * @param  {} args The arguments for the selection changed event.
 */
function onSelectionChange(args) {
  let g = getGlobal();
  setContextualTabVisibility(args.isInsideTable);
  g.isTableSelected = args.isInsideTable;
}

/**
 * Handles the onChanged event. When data in the sales table is changed,
 * enable the refresh and submit buttons.
 */
function onChanged() {
  let g = getGlobal();

  //When the add-in creates the table, it will generate 4 events that we must ignore.
  //We only want to respond to the change events from the user.
  if (g.tableEventCount > 0) {
    g.tableEventCount--;
    return; //count down to throw away events caused by the table creation code
  }

  //check if dirty flag was set (flag avoids extra unnecessary ribbon operations)
  if (!g.isTableDirty) {
    g.isTableDirty = true;

    //Enable the Refresh and Submit buttons
    setSyncButtonEnabled(true);
  }
}
```




## Security notes

In the webpack.config.js file, a header is set to  `"Access-Control-Allow-Origin": "*"`. This is only for development purposes. In production code, you should list the allowed domains and not leave this header open to all domains.

You'll be prompted to install certificates for trusted access to https://localhost. The certificates are intended only for running and studying this code sample. Do not reuse them in your own code solutions or in production environments.

You can install or uninstall the certificates by running the following commands in the project folder.

```command&nbsp;line
npx office-addin-dev-certs install
npx office-addin-dev-certs uninstall
```

## Questions and comments

We'd love to get your feedback about this sample. You can send your feedback to us in the Issues section of this repository. Questions about developing Office Add-ins should be posted to Stack Overflow. Ensure your questions are tagged with [office-js].

## Additional resources

- [Office Add-ins documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)

## Copyright

Copyright (c) 2021 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/samples/Office.Contextual-tabs" />