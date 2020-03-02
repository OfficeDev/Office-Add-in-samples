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
  createdDate: 3/02/2020 1:25:00 PM
description: "This sample shows how to create contextual ribbon buttons that are enabled based on the state of your add-in. It also shows how to use the Office.js API to show or hide the task pane. This sample also demonstrates how to run code when the task pane is closed, such as on document open."
---

# Manage ribbon UI and run code without using the task pane

## Summary

This sample shows how to create contextual ribbon buttons that are enabled based on the state of your add-in. It also shows how to use the Office.js API to show or hide the task pane. This sample also demonstrates how to run code when the task pane is closed, such as on document open.

## Features

- Contextual ribbon UI that enables or disables the buttons.
- Handle Office.js events even when the task pane is closed.
- Share data globally, such as between custom functions and the task pane.

## Applies to

-  Excel on Windows (one-time purchase and subscription)

## Prerequisites


Before running this sample, make sure you have installed a recent version of [npm](https://www.npmjs.com/get-npm) and [Node.js](https://nodejs.org/en/) on your computer. To check if you have already installed these tools, run the commands `node -v` and `npm -v` in your terminal.

## Solution

Solution | Author(s)
---------|----------
Office Add-in Shared Runtime Ribbon/Task pane APIs | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.0 | 3-2-2020 | Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

----------

## Scenario: A contextual add-in

This sample demonstrates a fictional scenario where the add-in connects to a backend data service to help the user import and work with Contoso data. There is no actual data service and the add-in mocks the data you see.

The add-in is aware of whether it is connected. When connected you will see the task pane update to allow you to import data, and also the ribbon buttons will be enabled to let you insert a table and work with the data. 

Additionally the add-in has a custom function that can display a filtered view of the data. The custom function is aware of the connection status, so that when connected, it will display the mock data. When disconnected, it will show `#N/A`.

## Build and run the solution

In the command prompt, run the command `start npm start`. This will open a second command prompt, build the project and then start a server (with dev mode settings). It takes from 5 to 30 seconds. When it finishes, the last line should say `Compiled successfully`. Minimize this command prompt.

Back in the original command prompt, run the command `npm run sideload`. This will launch Excel and sideload the add-in. After a few seconds, a ribbon named `Contoso Data` will appear.

The add-in's ribbon buttons have the following behavior:

- **Connect service:** Connects to a mock Contoso data service. You can choose a CSV file, or database.
- **Disconnect service:** Disonnects from the mock Contoso data service.
- **Insert data:** Inserts a table from the mock Contoso data service.
- **Sum:** Enabled when you are in the table. Select a range of numerical cells and it will output the sum of those cells.
- **Enable startup:** Choose this to enable the add-in to run the next time the document is opened. The `Sum` button will work immediately when in the table the next time the document is opened. Note that you need to save the document first to save this change.
- **Disable startup:** Choose this to disable the add-in from running on document open. The Sum button will not work until you activate the add-in in some way (ribbon, task pane or custom function action).
- **Open task pane:** Opens the task pane.
- **Close task pane:** Closes the task pane. The task pane is not shut down and will remember its state.

If the add-in is not connected to a service, the task pane will show a button to connect. Once connected, the task pane lets you choose a category from the data and insert a custom function. The custom function will filter data displayed to the selected category.

## Copyright

Copyright (c) 2020 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://telemetry.sharepointpnp.com/officedev/samples/shared-runtime-ribbon" />
