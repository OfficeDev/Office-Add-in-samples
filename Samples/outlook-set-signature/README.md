---
page_type: sample
products:
- office-outlook
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 02/25/2020 10:00:00 AM
description: "Use event-based activation to manage Outlook signatures."
---

# (Preview) Event-based activation to set signature

## Summary

This sample uses event-based activation to run an Outlook add-in when the user sends a new message. The add-in enables a user to manage their Outlook signatures without requiring them to manually open the task pane.

> **Note:** The features used in this sample are currently in preview and subject to change. They are not currently supported for use in production environments. To try the preview features, you'll need to [join Office Insider](https://insider.office.com/join). A good way to try out preview features is to sign up for a Microsoft 365 subscription. If you don't already have a Microsoft 365 subscription, get one by joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/office/dev-program).

## Features

- Use event-based activation to run code when the task pane is not open.
- Let the user choose and set a signature for emails.

## Applies to

-  Outlook on Windows, and in a browser.

## Prerequisites

To use this sample, you'll need to [join Office Insider](https://insider.office.com/join).

Before running this sample, you need a recent version of [npm](https://www.npmjs.com/get-npm) and [Node.js](https://nodejs.org/en/) installed on your computer. To verify if you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.

## Solution

Solution | Author(s)
---------|----------
Use event-based activation to manage Outlook signatures | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.0 | 2-25-2021 | Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

----------

## Scenario: Event-based activation

In this scenario, you want your add-in to help the user manage signatures, even when the task pane is not open. When the user sends a new email, the add-in displays  

## Build and run the solution

1. In the command prompt, run the command `npm install`.
2. Run the command `npm run dev-server`. This starts the Node.js server.
3. Sideload the add-in to Outlook on the desktop, or in a browser by following the manual instructions in the article [Sideload Outlook add-ins for testing](https://docs.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing).

Once the add-in is loaded use the following steps to try out the functionality.
1. In Outlook, send a new message.
    
    You should see a notification at the top of the message that reads: "Please set your signature with the PnP signature manager add-in."
    
2. Choose **Set signatures**. This will open the task pane for the add-in.
3. In the task pane you can use the UI to set up and try out several signature styles.

## Key parts of this sample

TBD

## Security notes

In the webpack.config.js file, a header is set to `"Access-Control-Allow-Origin": "*"`. This is only for development purposes. In production code, you should list the allowed domains and not leave this header open to all domains.

You'll be prompted to install certificates for trusted access to https://localhost. The certificates are intended only for running and studying this code sample. Do not reuse them in your own code solutions or in production environments.

You can install or uninstall the certificates by running the following commands in the project folder.

```
npx office-addin-dev-certs install
npx office-addin-dev-certs uninstall
```

## Copyright

Copyright (c) 2021 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/samples/outlook-autorun-set-signature" />