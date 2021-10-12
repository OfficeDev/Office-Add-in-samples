---
page_type: sample
urlFragment: office-add-in-hello-world
products:
- office-add-ins
- office-excel
- office-word
- office-outlook
- office-powerpoint
- office
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: "10/11/2021 10:00:00 AM"
description: "Create a simple Office Add-in that displays hello world."
---

# Create an Office Add-in that displays hello world

## Summary

Learn how to build the simplest Office Add-in with only a manifest, HTML web page, and a logo. This sample will help you understand the fundamental parts of an Office Add-in.

## Features

- Display hello world in Outlook, Word, Excel, or PowerPoint.
- Learn fundamentals of the manifest.
- Learn how to initialize the Office JavaScript API library.
- Interact with document content through Office JavaScript APIs.

## Applies to

- Office on Windows, Mac, and in a browser.

## Prerequisites

- Microsoft 365

## Understand an Office Add-in

An Office Add-in is a web application that can extend Office with additional functionality for the user. For example, an add-in can add ribbon buttons, a task pane, or a content pane with the functionality you want. Because an Office Add-in is a web application you must provide a web server to host the files.

This sample has the following four folders. Each folder is a sample that is designed to run in the indicated Office application.
- outlook-hello-world: Hello world sample for Outlook
- excel-hello-world: Hello world sample for Excel
- word-hello-world
- powerpoint-hello-world

To work with the samples, clone or download this repo. Then go to the folder containing the sample for the Office application you want to work with. All of the following guidance will apply to the sample you choose to work with.

### Key components

The hello world sample implements the **Manifest** and **Web app** components identified in [Components of an Office Add-in](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins#components-of-an-office-add-in).

- Manifest: You only need one manifest file for your add-in. The hello world sample contains two manifest files to support two different hosting scenarios.
    - **manifest.xml**: This manifest file will load the add-in from the GitHub repo (through GitHub page hosting). You can run the sample and don't need to configure your own web server.
    - **manifest.localhost.xml**: This manifest file will load the add-in from a local web server that you configure. See <TBD> later in this readme for instructions on configuring your own web server.
- Web app: The hello world sample implements a task pane named **taskpane.html** that contains HTML and JavaScript. The **taskpane.html** file contains all the code necessary to display a task pane, interact with the user, and write "Hello world!" to the document.

### Initialize the Office JavaScript API library

The sample initializes the Office JavaScript API library with a call to `office.onReady()` in the **taskpane.html** file. This is required before you can make any calls to the Office JavaScript APIs. for more information about initialization, see [Initialize your Office Add-in](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/initialize-add-in).

```javascript
Office.onReady((info) => {
  });
```

### Write to the document

When the user chooses the **button**, the `sayHello()` function is called. This function Then calls `Excel.run` to run code and call the Office JavaScript APIs. It uses a `context` object provided by the Office JS API library to get the active worksheet's `A1` range value and set the value to "Hello world!". Calling `context.sync()` runs the command.

For more information see []()


```javascript
 function sayHello() {
        Excel.run(context => {
            context.workbook.worksheets.getActiveWorksheet().getRange("A1").values = [['Hello world!']];
            return context.sync();
        });
    }
```

## Run the sample

An Office Add-in requires you to configure a web server to provide all the resources, such as HTML, image, and JavaScript files. The hello world sample is configured so that the files are hosted directly from this GitHub repo. Use the following steps to sideload the manifest.xml file to see the sample run.

1. Download the **manifest.xml** file from the sample folder for Outlook, Word, Excel, or PowerPoint.
1. Open [Office on the web](https://office.live.com/).
1. Choose **Excel**, and then open a new document.
1. Open the **Insert** tab on the ribbon and choose **Office Add-ins**.
1. On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.
   ![The Office Add-ins dialog with a drop-down in the upper right reading "Manage my add-ins" and a drop-down below it with the option "Upload My Add-in"](../../Samples/images/office-add-ins-my-account.png)
1. Browse to the add-in manifest file, and then select **Upload**.
   ![The upload add-in dialog with buttons for browse, upload, and cancel.
](../../Samples/images/upload-add-in.png)
1. Verify that the add-in loaded successfully. You will see a **Hello world** button on the **Home** tab on the ribbon.

Choose the **Hello world** button to see the add-in display "Hello world!" in cell A1.

## Run the sample from Localhost

If you prefer to run the web server and host the add-in's web files from your computer, use the following steps:

1. Clone or download this sample to a folder on your computer. Then go to that folder in a console or terminal window.
1. You need http-server to run the local web server. If you haven't installed this yet you can do this with the following command:
    
    ```console
    npm install --global http-server
    ```
    
2. Run the following command to generate a self-signed certificate that you can use for the web server.

    ```console
    npx office-addin-dev-certs install
    ```

    The previous command will display the folder location where it generated the certificate files.

3. Go to the folder location where the certificate files were generated. Copy the localhost.crt and localhost.key files to the hello world sample folder.
4. Run the following command:
    
    ```console
    http-server -S -C localhost.crt -K localhost.key --cors . -p 3000
    ```
    
    The http-server will run and host the current folder's files on localhost:3000.
    
5. Follow the steps in [Run the sample](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/excel-keyboard-shortcuts#run-the-sample), but upload the `manifest-localhost.xml` file for step 6.

## Run the sample

You can run the 

## Solution

Solution | Author(s)
---------|----------
Use keyboard shortcuts for Office add-in actions | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.0 | 11-5-2020 | Initial release
1.1 | May 11, 2021 | Removed yo office and modified to be GitHub hosted

----------

## Scenario: Open/Close taskpane and modify cell color

This sample adds three different shortcuts to the Office Add-in. This enables the user to:

- Use the "Ctrl+Alt+1" keyboard shortcut to open the task pane.
- Use the "Ctrl+Alt+2" keyboard shortcut to close the task pane.
- Use the "Ctrl+Alt+3" keyboard shortcut to cycle through colors for a selected cell.
- Use the "Ctrl+R" keyboard shortcut to test the shortcut conflict modal.

## Run the sample

You can run this sample in Excel in a browser. The add-in web files are served from this repo on GitHub.

1. Download the **manifest.xml** file from this sample to a folder on your computer.
1. Open [Office on the web](https://office.live.com/).
1. Choose **Excel**, and then open a new document.
1. Open the **Insert** tab on the ribbon and choose **Office Add-ins**.
1. On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.
   ![The Office Add-ins dialog with a drop-down in the upper right reading "Manage my add-ins" and a drop-down below it with the option "Upload My Add-in"](../../Samples/images/office-add-ins-my-account.png)
1. Browse to the add-in manifest file, and then select **Upload**.
   ![The upload add-in dialog with buttons for browse, upload, and cancel.
](../../Samples/images/upload-add-in.png)
1. Verify that the add-in loaded successfully. You will see a **PnP keyboard shortcuts** button on the **Home** tab on the ribbon.

Once the add-in is loaded use the following steps to try out the functionality.

1. Press "Ctrl+Alt+1" on the keyboard to trigger the Show Taskpane action.
2. In the task pane, you will see the additional shortcuts available to try in the sample.

## Key parts of this sample

The manifest.xml is pre-configured to use the shared runtime. To see how to add shared runtime to your own add-in, use the following article:

- [Configure your Excel Add-in to use a shared JavaScript runtime](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/configure-your-add-in-to-use-a-shared-runtime)


Additionally, the following changes have been made to enable keyboard shortcuts:

1. Configured the add-in's manifest by adding the new element `ExtendedOverrides` to the end of the manifest.
2. Created the shortcuts JSON file `shortcuts.json`, in the `src/` folder to define actions and their keyboard shortcuts. Ensure the new file is properly bundled by configuring the `webpack.config.js` file.
3. Mapped actions to runtime calls with the associate method in `src/taskpane.js`.


## Run the sample from Localhost

If you prefer to host the web server for the sample on your computer, follow these steps:

1. Open the **/src/commands/ribbonJSON.js** file.
1. Edit line 9 to refer to the localhost:3000 endpoint as shown in the following code.
    
    ```javascript
    const sourceUrl = "https://localhost:3000";
    ```
    
1. Save the file.
1. You need http-server to run the local web server. If you haven't installed this yet you can do this with the following command:
    
    ```console
    npm install --global http-server
    ```
    
2. Use a tool such as openssl to generate a self-signed certificate that you can use for the web server. Move the cert.pem and key.pem files to the webworker-customfunction folder for this sample.
3. From a command prompt, go to the web-worker folder and run the following command:
    
    ```console
    http-server -S --cors . -p 3000
    ```
    
4. To reroute to localhost run office-addin-https-reverse-proxy. If you haven't installed this you can do this with the following command:
    
    ```console
    npm install --global office-addin-https-reverse-proxy
    ```
    
    To reroute run the following in another command prompt:
    
    ```console
    office-addin-https-reverse-proxy --url http://localhost:3000
    ```
    
5. Follow the steps in [Run the sample](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/excel-keyboard-shortcuts#run-the-sample), but upload the `manifest-localhost.xml` file for step 6.

## Copyright

Copyright (c) 2020 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/samples/excel-keyboard-shortcuts" />