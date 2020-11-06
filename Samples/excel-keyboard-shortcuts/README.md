---
page_type: sample
products:
- office-excel
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 11/5/2020 10:00:00 AM
description: "This sample shows how to add keyboard shortcuts to your Office Add-in."
---

# (Preview) Use keyboard shortcuts for Office add-in actions

## Summary

This sample shows how to set up a basic Excel add-in project that utilizes keyboard shortcuts. Currently, the shortcuts are configured to show and hide the task pane as well as cycle through colors for a selected cell. Keyboard shortcuts can be used to achieve any action within the add-in runtime.

> **Note:** The features used in this sample are currently in preview and subject to change. They are not currently supported for use in production environments. To try the preview features, you'll need to [join Office Insider](https://insider.office.com/join). A good way to try out preview features is to sign up for an Office 365 subscription. If you don't already have an Office 365 subscription, get one by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).

## Features

- Add keyboard shortcuts to your Office Add-in. 
- Enable users to use those keyboard shortcuts to invoke any action within the Office Add-in runtime.

## Applies to

-  Excel on Windows, Mac, and in a browser.

## Prerequisites

To use this sample, you'll need to [join Office Insider](https://insider.office.com/join).

Before running this sample, you need a recent version of [npm](https://www.npmjs.com/get-npm) and [Node.js](https://nodejs.org/en/) installed on your computer. To verify if you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.

## Solution

Solution | Author(s)
---------|----------
Office Add-in share global data with a shared runtime | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.0 | 11-5-2020 | Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

----------

## Scenario: Open/Close taskpane and modify cell color

This sample adds three different shortcuts to the Office Add-in. This enables the user to:
- Use the "Ctrl+Shift+Up" keyboard shortcut to open the taskpane.
- Use the "Ctrl+Shift+Down" keyboard shortcut to close the taskpane.
- Use the "Ctrl+Shift+O" keyboard shortcut to cycle through colors for a selected cell.

## Build and run the solution

1. In the command prompt, run the command `npm install`.
2. Run the command `npm run start`. This starts the Node.js server, and opens Excel on Windows.

If you're running Excel on the web or Mac, run `npm run start:web` instead and see the following articles for instructions on how to sideload:

- [Sideload Office Add-ins in Office on the web for testing](https://docs.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing) 
- [Sideload Office Add-ins on Mac for testing](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-ipad-and-mac#sideload-an-add-in-in-office-on-mac)

Once the add-in is loaded use the following steps to try out the functionality.

1. Make sure the focus of the keyboard is on the worksheet by selecting a cell.
2. Press "Ctrl+Shift+Up" on the keyboard to trigger the Show Taskpane action. 
3. In the task pane, you will see the additional shortcuts available to try in the sample.


## Key parts of this sample

The manifest.xml is pre-configured to use the shared runtime. To see how to add shared runtime to your own add-in, use the following article:

- [Configure your Excel Add-in to use a shared JavaScript runtime](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/configure-your-add-in-to-use-a-shared-runtime)


Additionally, the following changes have been made to enable keyboard shortcuts:

1. Configured the add-in's manifest by adding the new element `ExtendedOverrides` to the end of the manifest.
2. Created the shortcuts JSON file `shortcuts.json`, in the `src/` folder to define actions and their keyboard shortcuts. Ensure the new file is properly bundled by configuring the `webpack.config.js` file.
3. Mapped actions to runtime calls with the associate method in `src/taskpane/taskpane.js`.


## Security notes

In the webpack.config.js file, a header is set to `"Access-Control-Allow-Origin": "*"`. This is only for development purposes. In production code, you should list the allowed domains and not leave this header open to all domains.

You'll be prompted to install certificates for trusted access to https://localhost. The certificates are intended only for running and studying this code sample. Do not reuse them in your own code solutions or in production environments.

You can install or uninstall the certificates by running the following commands in the project folder.

```
npx office-addin-dev-certs install
npx office-addin-dev-certs uninstall
```

## Copyright

Copyright (c) 2020 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/samples/excel-keyboard-shortcuts" />
