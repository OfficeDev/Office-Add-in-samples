---
page_type: sample
urlFragment: excel-add-in-hello-world
products:
  - office-excel
  - office
  - m365
languages:
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: '10/11/2021 10:00:00 AM'
description: 'Create a simple Excel add-in that displays hello world.'
---

# Create an Excel add-in that displays hello world

## Summary

Learn how to build the simplest Office Add-in with only a manifest, HTML web page, and a logo. This sample will help you understand the fundamental parts of an Office Add-in.

## Features

- Display hello world in Excel.
- Learn fundamentals of the manifest.
- Learn how to initialize the Office JavaScript API library.
- Interact with document content through Office JavaScript APIs.

## Applies to

- Excel on Windows, Mac, and in a browser.

## Prerequisites

- Microsoft 365 - You can get a free developer sandbox by joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program#Subscription).

## Understand an Office Add-in

An Office Add-in is a web application that can extend Office with additional functionality for the user. For example, an add-in can add ribbon buttons, a task pane, or a content pane with the functionality you want. Because an Office Add-in is a web application you must provide a web server to host the files.

The sample contained in this folder is designed to run in Excel.

## Key components

The hello world sample implements the **Manifest** and **Web app** components identified in [Components of an Office Add-in](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins#components-of-an-office-add-in).

### Manifest

The manifest file describes your add-in to Office. It contains information such as a unique identifier, name, what buttons to show on the ribbon, and more. Importantly, the manifest provides URL locations for where Office can find and download the add-in's resource files. The manifest and two icon files are combined into a zip package file that is sideloaded to Office.

### Web app

The hello world sample implements a task pane named **taskpane.html** that contains HTML and JavaScript. The **taskpane.html** file contains all the code necessary to display a task pane, interact with the user, and write "Hello world!" into cell 'A1' on the active Worksheet.

### Initialize the Office JavaScript API library

The sample initializes the Office JavaScript API library with a call to `office.onReady()` in the **taskpane.html** file. This is required before you can make any calls to the Office JavaScript APIs. For more information about initialization, see [Initialize your Office Add-in](https://learn.microsoft.com/office/dev/add-ins/develop/initialize-add-in).

```javascript
Office.onReady((info) => {});
```

### Write to the worksheet

When the user chooses the **Say hello** button, the `sayHello()` function is called. This function then calls `Excel.run` to run code and call the Office JavaScript APIs. It uses a `context` object provided by the Office JS API library to get the active worksheet's `A1` range value and set the value to "Hello world!". Calling `context.sync()` runs the command.

For more information see [Tutorial: Create an Excel task pane add-in](https://learn.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial)

```javascript
function sayHello() {
  Excel.run((context) => {
    context.workbook.worksheets.getActiveWorksheet().getRange('A1').values = [
      ['Hello world!'],
    ];
    return context.sync();
  });
}
```

## Run the sample with GitHub as host

An Office Add-in requires you to configure a web server to provide all the resources, such as HTML, image, and JavaScript files. The Hello World sample is configured so that the files are hosted directly from this GitHub repo, so all you need to do is build the manifest and package, and then sideload the package. 

1. Clone or download this sample to a folder on your computer. Then in a command prompt, bash shell, or **TERMINAL** in Visual Studio Code, navigate to the root of the sample folder.
1. Run the command `npm install`.
1. Run the command `npm run build`.
1. Run the command `npm run start:prod`.

After a few seconds, desktop Word opens, and after a few seconds more, a **Hello World** button appears on the right end of the **Home** ribbon. 

1. Choose the **Hello world** button to display the task pane of the add-in.
1. Choose the **Say hello** button to insert "Hello world!" in cell A1.

When you're finished working with the add-in, close Excel, and then in the window where you ran the three npm commands, run `npm run stop:prod`.

## Run the sample on Excel on Windows or Mac

Office Add-ins are cross-platform so you can also run them on Windows, Mac, and iPad. The following links will take you to documentation for how to sideload on Windows, Mac, or iPad. Be sure you have a local copy of the manifest.xml file for the Hello world sample. Then follow the sideloading instructions for your platform.

- [Sideload Office Add-ins for testing from a network share](https://learn.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)
- [Sideload Office Add-ins on Mac for testing](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)
- [Sideload Office Add-ins on iPad for testing](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-an-office-add-in-on-ipad)

## Configure a localhost web server and run the sample from localhost

If you prefer to configure a web server and host the add-in's web files from your computer, use the following steps.

1. Clone or download this sample to a folder on your computer. Then in a command prompt, bash shell, or **TERMINAL** in Visual Studio Code, navigate to the root of the sample folder.
1. Run the command `npm install`.
1. Run the command `npm start`.

- If you've never developed an Office add-in on this computer before or it has been more than 30 days since you last did, you'll be prompted to delete an old security cert and/or install a new one. Agree to both prompts. 
- After a few seconds a **webpack** dev-server window will open and your files will be hosted there on localhost:3000.
- When the server is successfully running, desktop Excel opens, and after a few seconds more, a **Hello World** button appears on the right end of the **Home** ribbon.

1.  Choose the **Hello world** button to display the task pane of the add-in.
1.  Choose the **Say Hello** button to insert "Hello World" into the document.

When you're finished working with the add-in, close Excel, and then in the window where you ran the two npm commands, run `npm stop`.

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2021 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

**Note**: The taskpane.html file contains an image URL that tracks diagnostic data for this sample add-in. Please remove the image tag if you reuse this sample in your own code project.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/excel-add-in-hello-world" />
