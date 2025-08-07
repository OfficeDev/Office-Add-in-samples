---
page_type: sample
urlFragment: excel-add-in-commands
products:
  - office-add-ins
  - office-excel
  - office
  - m365
  - office-teams
languages:
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: '12/09/2021 10:00:00 AM'
description: 'Create an Excel add-in with command buttons.'
---

# Create an Excel add-in with command buttons

## Summary

Learn how to build an Office Add-in that has a command button to show the task pane, and a menu dropdown button that can show the task pane, or get data.

![Excel showing ribbon with Home tab selected and two buttons for the sample that show the task pane or show a dropdown menu.](../images/excel-ribbon-buttons.png)

## Features

- A button in the ribbon that shows the task pane.
- A dropdown button in the ribbon with two menu commands.

## Applies to

- Excel on Windows, Mac, and in a browser.

## Prerequisites

- Microsoft 365 - Get a [free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) that provides a renewable 90-day Microsoft 365 E5 developer subscription.

## Version history

| Version  | Date | Comments |
|----------|------|----------|
| 1.0 | 12-09-2021 | Initial release |
| 1.1 | 08-07-2025 | Add support for the unified manifest for Microsoft 365 |

## Decide on a version of the manifest

- Add-in only manifest
  - To run the add-in only manifest, which is the **manifest.xml** file in the sample's root directory **Samples/add-in-commands/excel**, go to the [Add-in only manifest](#add-in-only-manifest) section.
- [Unified manifest for Microsoft 365](https://learn.microsoft.com/office/dev/add-ins/develop/json-manifest-overview)
  - To run the unified manifest for Microsoft 365 (**manifest.json**), go to the [Unified manifest](#unified-manifest) section.

## Add-in only manifest

### Run the sample

Use one of the following add-in file hosting options to run the sample.

#### Use GitHub as the web host

You can run this sample in Excel on Windows, on Mac, or in a browser. The add-in web files are served from this repo on GitHub.

1. Download the **manifest.xml** file from this sample to a folder on your computer.
1. Sideload the add-in manifest in Excel by following the appropriate instructions in the article [Sideload an Office Add-in for testing](https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing).
1. Follow the steps in [Try it out](#try-it-out) to test the sample.

#### Configure a localhost web server and run the sample from localhost

If you prefer to configure a web server and host the add-in's web files from your computer, use the following steps.

1. Install a recent version of [npm](https://www.npmjs.com/get-npm) and [Node.js](https://nodejs.org/) on your computer. To verify if you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.

1. You need http-server to run the local web server. If you haven't installed this yet, you can do this with the following command.

    ```console
    npm install --global http-server
    ```

1. You need Office-Addin-dev-certs to generate self-signed certificates to run the local web server. If you haven't installed this yet, you can do this with the following command.

    ```console
    npm install --global office-addin-dev-certs
    ```

1. Clone or download this sample to a folder on your computer. Then go to that folder in a console or terminal window.
1. Run the following command to generate a self-signed certificate that you can use for the web server.

    ```console
    npx office-addin-dev-certs install
    ```

    The previous command will display the folder location where it generated the certificate files.

1. Go to the folder location where the certificate files were generated. Copy the localhost.crt and localhost.key files to the hello world sample folder.

1. Run the following command.

    ```console
    http-server -S -C localhost.crt -K localhost.key --cors . -p 3000
    ```

    The http-server will run and host the current folder's files on localhost:3000.

1. Sideload **manifest-localhost.xml** in Excel by following the appropriate instructions in the article [Sideload an Office Add-in for testing](https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing).

1. Follow the steps in [Try it out](#try-it-out) to test the sample.

### Key parts of this sample

#### Commands UI

The **manifest.xml** file defines all of the commands UI in the `<ExtensionPoint>` element.
The ribbon buttons and dropdown menu are specified in the `<OfficeTab>` section. Because `<OfficeTab id="TabHome">` specifies `TabHome`, the buttons are located on the **Home** ribbon tab.

For more information about ExtensionPoint elements and options, see [Add ExtensionPoint elements](https://learn.microsoft.com/office/dev/add-ins/develop/create-addin-commands#step-4-add-extensionpoint-elements).

#### Commands JavaScript

The **manifest.xml** file contains a `<FunctionFile resid="Commands.Url"/>` element that specifies where to find the JavaScript commands to run when buttons are used. The `Commands.Url` resource ID points to `/src/commands/commands.html`. When a button command is chosen, **commands.html** is loaded, which then loads `/src/commands/commands.js`. This is where the `ExecuteFunction` actions are mapped from the **manifest.xml** file.

For example, the following manifest XML maps to the `writeValue` function in **commands.js**.

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>writeValue</FunctionName>
</Action>
```

```javascript
async function writeValue(event) {
...
```

For more information about adding commands, see [Add the FunctionFile element](https://learn.microsoft.com/office/dev/add-ins/develop/create-addin-commands#step-3-add-the-functionfile-element).

## Unified manifest

### Prerequisites

- If you want to run the web server on localhost, install a recent version of [npm](https://www.npmjs.com/get-npm) and [Node.js](https://nodejs.org) on your computer. To check if you've already installed these tools, from a command prompt, run the following commands.

    ```console
    node -v
    npm -v
    ```

- If you want to run the sample using GitHub as the web host, install the [Microsoft 365 Agents Toolkit command line interface (CLI)](https://learn.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-cli). From a command prompt, run the following command.

    ```console
    npm install -g @microsoft/teamsapp-cli
    ```

### Run the sample

You can run this sample in Excel on Windows, on Mac, or in a browser. Use one of the following add-in file hosting options.

#### Use GitHub as the web host

The quickest way to run the sample is to use GitHub as the web host. However, you can't debug or change the source code. The add-in web files are served from this GitHub repository.

1. Download the **manifest-configurations/unified/excel-add-in-commands.zip** file from this sample to a folder on your computer.
1. Sideload the add-in manifest in Excel by following the appropriate instructions in the article [Sideload Office Add-ins that use the unified manifest for Microsoft 365](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-add-in-with-unified-manifest).
1. Follow the steps in [Try it out](#try-it-out) to test the sample.

#### Use localhost

If you prefer to host the web server on localhost, follow these steps:

1. Clone or download this repository.
1. From a command prompt, go to the root of the project folder **/Samples/add-in-commands/excel**.
1. Copy the files from the **manifest-configurations/unified** subfolder to the root folder.
1. Run the following commands.

    ```console
    npm install
    ```

    ```console
    npm start
    ```

    This starts the web server on localhost and sideloads the **manifest.json** file to Excel.

1. Follow the steps in [Try it out](#try-it-out) to test the sample.

1. To stop the web server and uninstall the add-in from Excel, run the following command.

    ```console
    npm stop
    ```

## Key parts of this sample

### Commands UI

The **manifest.json** file defines all of the commands UI in the "extensions" array.
The ribbon buttons and dropdown menu are specified in the "ribbons.tabs" array. Because "builtInTabId" specifies `TabHome`, the buttons are located on the **Home** ribbon tab.

For more information about adding ribbon buttons, see [Menu and menu items](https://learn.microsoft.com/office/dev/add-ins/develop/create-addin-commands-unified-manifest#menu-and-menu-items).

### Commands JavaScript

The **manifest.json** file contains a runtime in the "extensions.runtimes" array with "id" set to "CommandsRuntime" that specifies where to find the JavaScript commands to run when buttons are used. The "code.page" property points to `/src/commands/commands.html`. When a button command is chosen, **commands.html** is loaded, which then loads `/src/commands/commands.js`. This is where the `executeFunction` actions are mapped from the **manifest.json** file.

For example, the following manifest JSON maps to the `writeValue` function in **commands.js**.

```json
"actions": [
  {
    "id": "writeValue",
    "type": "executeFunction",
    "displayName": "Write Value"
  }
]
```

```javascript
async function writeValue(event) {
...
```

For more information about adding commands, see [Add a function command](https://learn.microsoft.com/office/dev/add-ins/develop/create-addin-commands-unified-manifest#add-a-function-command).

## Try it out

1. Verify that the add-in loaded successfully. You'll see a **Show task pane** button and **Dropdown menu** button on the **Home** tab on the ribbon.

1. On the **Home** tab, choose the **Show task pane** button to display the task pane of the add-in. Choose the **Dropdown menu** button to see a dropdown menu. From the menu, you can show the task pane, or choose **Write value** to call a command that writes the button's ID to the current cell.

## Copyright

Copyright (c) 2021 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/excel-add-in-commands" />
