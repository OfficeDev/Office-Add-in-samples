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
  createdDate: 3/15/2020 1:25:00 PM
description: "This sample shows how to share data across the ribbon, task pane, and custom functions."
---

# (preview) Share global data with a shared runtime

## Summary

This sample shows how to set up a basic project that uses the shared runtime. The shared runtime runs all parts of your Excel add-in (ribbon buttons, task pane, custom functions) in a single browser runtime. This makes it easy to shared data through local storage, or through global variables.

![Screen shot of the add-in with ribbon buttons enabled and disabled](excel-shared-runtime-global.png)

> **Note:** The features used in this sample are currently in preview and subject to change. They are not currently supported for use in production environments. To try the preview features, you will need to [join Office Insider](https://insider.office.com/join). A good way to try out preview features is by using an Office 365 subscription. If you don't already have an Office 365 subscription, you can get one by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).

## Features

- Share data globally with ribbon buttons, the task pane, and custom functions.
- Use a provided manifest XML file to quick start a new project with a shared runtime.

## Applies to

-  Excel on Windows, Mac, and in a browser.

## Prerequisites

To try the preview features used by this sample, you will need to [join Office Insider](https://insider.office.com/join).

Before running this sample, make sure you have installed a recent version of [npm](https://www.npmjs.com/get-npm) and [Node.js](https://nodejs.org/en/) on your computer. To check if you have already installed these tools, run the commands `node -v` and `npm -v` in your terminal.

## Solution

Solution | Author(s)
---------|----------
Office Add-in share global data with a shared runtime | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.0 | 3-15-2020 | Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

----------

## Scenario: Sharing key/value pairs

This sample enables a user to store and retrieve key/value pairs by using the task pane or custom functions. The user can select which type of storage is used. They can choose to store key/value pairs in local storage, or choose to use a global variable.

## Build and run the solution

In the command prompt, run the command `npm run start`. This will start the node server, and automatically open Excel on the desktop.
If you are running Excel on the web or Mac, see the following articles for instructions on how to sideload:

- [Sideload Office Add-ins in Office on the web for testing](https://docs.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing) 
- [Sideload Office Add-ins on iPad and Mac for testing](https://docs.microsoft.com/office/dev/add-ins/testing/sideload-an-office-add-in-on-ipad-and-mac)

Once the add-in is loaded use the following steps to try out the functionality.

1. On the `Home` tab, choose `Show TaskPane`.
2. In the task pane, enter a key/value pair and choose `Store key/value pair`.
![Screen shot of both key and value input fields, and both store and get buttons.](task-pane-buttons.png)
3. In any spreadsheet cell, enter the formula `=CONTOSO.GETVALUEFORKEYCF("1")`. Pass the value of the key you created from the task pane.
4. In any spreadsheet cell, enter the formula `=CONTOSO.SETVALUEFORKEYCF("2","oranges")`. The formula should return the text `Stored key/value pair`.
5. In the task pane, enter the key from the previous formula `2` and choose `Get value for key`. The task pane should display the value `oranges`.

The task pane and custom function are sharing data via a global variable in the shared runtime. You can switch the method of storage by choosing either the `Global variable` or `Local storage` radio buttons on the task pane.

## Key parts of this sample

The manifest.xml is configured to use the shared runtime by using the `Runtimes` element as follows:

```xml
<Runtimes>
   <Runtime resid="Shared.Url" lifetime="long" />
</Runtimes>
```

If you read through other parts of the manifest you will see that the custom functions and task pane are also configured to use the `Shared.Url` because they will all run in the same runtime. `Shared.Url` points to `taskpane.html` which will load the shared runtime.

Global state is tracked in a window object retrieved using a `getGlobal()` function. This is accessible to custom functions, the task pane, and the ribbon (because all the code is running in the same JavaScript runtime.) 

There are no commands.html or functions.html files. These are not necessary because their purpose is to load individual runtimes. These do not apply when using the shared runtime.

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

<img src="https://telemetry.sharepointpnp.com/officedev/samples/excel-shared-runtime-global-state" />
