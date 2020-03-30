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
  createdDate: 3/30/2020 1:25:00 PM
description: "A starter template for creating add-ins that use preview shared runtime features."
---

# (Preview) Shared runtime starter template

## Summary

This template creates a new project similar to running `yo office` but instead it will use a shared runtime. You can use this template to start trying out features in the shared runtime. 

> **Note:** The shared runtime features are currently in preview and subject to change. They are not currently supported for use in production environments. To try the preview features, you'll need to [join Office Insider](https://insider.office.com/join). A good way to try out preview features is to sign up for an Office 365 subscription. If you don't already have an Office 365 subscription, get one by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).

## Features

- Lets you get started using the shared runtime.

## Applies to

-  Excel on Windows, Mac, and in a browser.

## Prerequisites

To use this sample, you'll need to [join Office Insider](https://insider.office.com/join).

Before running this template, you need a recent version of [npm](https://www.npmjs.com/get-npm) and [Node.js](https://nodejs.org/en/) installed on your computer. To verify if you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.

## Solution

Solution | Author(s)
---------|----------
Office Add-in template for shared runtime | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.0 | 3-26-2020 | Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

----------

## Start a new Excel add-in project with this template

1. Use a tool to generate a new GUID. If you are using Visual Studio Code, there are several extensions available that generate GUIDs.
2. Open the manifest.xml file, and insert the new GUID into the `<Id></Id>` near the top of the file. 
3. Run the command `npm install` in a terminal window or command prompt to install all the required dependencies.
4. Run the command `npm run start` in a terminal window or command prompt. This starts the node server, and opens Excel on the desktop.

If you're running Excel on the web or Mac, see the following articles for instructions on how to sideload:

- [Sideload Office Add-ins in Office on the web for testing](https://docs.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing) 
- [Sideload Office Add-ins on iPad and Mac for testing](https://docs.microsoft.com/office/dev/add-ins/testing/sideload-an-office-add-in-on-ipad-and-mac)

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
