---
page_type: sample
products:
- office-excel
- office-365
languages:
- typescript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 3/09/2020 1:25:00 PM
description: "This sample shows how to create contextual ribbon buttons that are enabled based on the state of your add-in. It also shows how to use the Office.js API to show or hide the task pane. This sample also demonstrates how to run code when the task pane is closed, such as on document open."
---

# (preview) Manage ribbon and task pane UI, and run code on doc open

## Summary

This sample shows how to create contextual ribbon buttons that are enabled based on the state of your add-in. It also shows how to use the Office.js API to show or hide the task pane. This sample also demonstrates how to run code when the task pane is closed, such as on document open.

![Screen shot of the add-in with ribbon buttons enabled and disabled](excel-shared-runtime.png)

> **Note:** The features used in this sample are currently in preview and subject to change. They are not currently supported for use in production environments. To try the preview features, you will need to [join Office Insider](https://insider.office.com/join). A good way to try out preview features is by using an Office 365 subscription. If you don't already have an Office 365 subscription, you can get one by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).

## Features

- Contextual ribbon UI that enables or disables the buttons.
- Set load behavior to load the add-in and run code when the document is opened.
- Open and close the task pane through the Office.js API.
- Handle Office.js events even when the task pane is closed.
- Share data globally, such as between custom functions and the task pane.

## Applies to

-  Excel on Windows, Mac, and in a browser.

## Prerequisites

To try the preview features used by this sample, you will need to [join Office Insider](https://insider.office.com/join).

Before running this sample, make sure you have installed a recent version of [npm](https://www.npmjs.com/get-npm) and [Node.js](https://nodejs.org/en/) on your computer. To check if you have already installed these tools, run the commands `node -v` and `npm -v` in your terminal.

## Solution

Solution | Author(s)
---------|----------
Office Add-in Shared Runtime Ribbon/Task pane APIs | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.0 | 3-9-2020 | Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

----------

## Scenario: A contextual add-in

This sample demonstrates a fictional scenario where the add-in connects to a backend data service to help the user import and work with Contoso data. To keep things simple, the data is mock data and the sample does not require an actual backend data service.

The add-in is aware of whether it is connected. When connected you will see the task pane update to allow you to import data, and also the ribbon buttons will be enabled to let you insert a table and work with the data. 

Additionally the add-in has a custom function that can display a filtered view of the data. The custom function is aware of the connection status, so that when connected, it will display the mock data. When disconnected, it will show `#N/A`.

## Build and run the solution

In the command prompt in the root of the project, run the command `npm install`.

When the installation completes, run the command `start npm start`. This will open a second command prompt, build the project and then start a server (with dev mode settings). It takes from 5 to 30 seconds. When it finishes, the last line should say `Compiled successfully`. Minimize this command prompt.

Back in the original command prompt, run the command `npm run sideload`. This will launch Excel and sideload the add-in. After a few seconds, a ribbon named `Contoso Data` will appear.

The add-in's ribbon buttons have the following behavior:

- **Connect service:** Connects to a mock Contoso data service. You can choose a CSV file, or database.
- **Disconnect service:** Disonnects from the mock Contoso data service.
- **Insert data:** Inserts a table from the mock Contoso data service.
- **Sum:** Enabled when you are in the table. Select a range of numerical cells and it will output the sum of those cells.
- **Load on doc open:** Choose this to enable the add-in to run the next time the document is opened. The `Sum` button will work immediately when in the table the next time the document is opened. Note that you need to save the document first to save this change.
- **No load behavior:** Choose this to disable the add-in from running on document open. The Sum button will not work until you activate the add-in in some way (ribbon, task pane or custom function action).
- **Open task pane:** Opens the task pane.
- **Close task pane:** Closes the task pane. The task pane is not shut down and will remember its state.

If the add-in is not connected to a service, the task pane will show a button to connect. Once connected, the task pane lets you choose a category from the data and insert a custom function. The custom function will filter data displayed to the selected category.


## Key parts of this sample

Global state is tracked in a window object retrieved using a `getGlobal()` function. This is accessible to custom functions, the task pane, and the ribbon (because all the code is running in the same JavaScript runtime.) The `ensureStateInitialized()` method initializes global state on startup. The global state tracks many items such as whether the add-in is connected to a service, and whether the task pane is open or closed.

### Enable and disable ribbon buttons

To enable and disable ribbon buttons the add-in uses the `Office.ribbon.requestUpdate` method and passes a JSON description of the buttons. Anytime there is a state change that requires the ribbon buttons to change, the `updateRibbon()` is called which sets if buttons are enabled based on global state variables.

### Open and close the task pane

To open or close the task pane, the add-in uses the `Office.addin.showAsTaskpane()` and `Office.addin.hide()` methods.

### Set runtime load behavior

If the `Load on doc open` button is chosen, the add-in configures the document so that the add-in begins running as soon as the document is opened. This allows the add-in to start running and monitoring Office events before the task pane is displayed. To configure the document to load the add-in on open, the add-in uses the `Office.addin.setStartupBehavior(Office.StartupBehavior.load)` method. To turn off the document load behavior, the add-in uses the `Office.addin.setStartupBehavior(Office.StartupBehavior.none)` method.

### Run code in the background

The add-in has a `Sum` button that is enabled when you move the range selection inside the expenses table. An event in the table is used to detect when the range selection is in or out of the table. If the add-in was configured to load on doc open, the event code will be operational as soon as the document is opened, and the `Sum` button will work even though the task pane is not yet opened.

To run code when the document opens, the add-in relies on the `Office.Initialize` event. This event is called when the document is opened and the load behavior is set for doc open. The add-in calls `ensureInitialize()` to set up the initial global state.

```typescript
Office.initialize = async () => {
    ensureStateInitialized(true);
    
    ...
```

The `ensureStateInitialized()` method will call the `monitorSheetCHanges()` method which will then search for the expenses table. If the expenses table was inserted, it adds an event handler for the `onTableSelectionChange` event. This is how the event handler code is set up on doc open.

```typescript
export async function monitorSheetChanges() {
  try {
    let g = getGlobal() as any;
    if (g.state.isInitialized) {
      await Excel.run(async context => {
        let table = context.workbook.tables.getItem('ExpensesTable');
        if (table !== undefined) {
          table.onSelectionChanged.add(onTableSelectionChange);
          await context.sync();
          updateRibbon();
        } else {
          g.state.isSumEnabled = false;
          updateRibbon();
        }
      ...
```

## Security notes

In the webpack.dev.js file, a header is set to `"Access-Control-Allow-Origin": "*"`. This is only for development purposes. In production code, you should list the allowed domains and not leave this header open to all domains.

You'll be prompted to install certificates for trusted access to https://localhost. The certificates are intended only for running and studying this code sample. Do not reuse them in your own code solutions or in production environments.

You can install or uninstall the certificates by running the following commands in the project folder.

```
npx office-addin-dev-certs install
npx office-addin-dev-certs uninstall
```

## Copyright

Copyright (c) 2020 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://telemetry.sharepointpnp.com/officedev/samples/excel-shared-runtime-scenario" />
