---
page_type: sample
urlFragment: powerpoint-add-in-hello-world
products:
  - office-powerpoint
  - office
languages:
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: '10/11/2021 10:00:00 AM'
description: 'Create a simple PowerPoint add-in that displays hello world.'
---

# Create a PowerPoint add-in that displays "Hello World"

## Summary

Learn how to build the simplest Office Add-in with only a manifest, HTML web page, and a logo. This sample will help you understand the fundamental parts of an Office Add-in.

## Features

- Display "Hello World" in PowerPoint.
- Learn fundamentals of the manifest.
- Learn how to initialize the Office JavaScript API library.
- Interact with document content through Office JavaScript APIs.

## Applies to

- PowerPoint on Windows, Mac, and in a browser.

## Prerequisites

- Microsoft 365 - You can get a free developer sandbox by joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program#Subscription).

## Understand an Office Add-in

An Office Add-in is a web application that can extend Office with additional functionality for the user. For example, an add-in can add ribbon buttons, a task pane, or a content pane with the functionality you want. Because an Office Add-in is a web application you must provide a web server to host the files.

The sample contained in this folder is designed to run in PowerPoint.

## Key components

The Hello World sample implements the **Manifest** and **Web app** components identified in [Components of an Office Add-in](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins#components-of-an-office-add-in).

### Manifest

The manifest file describes your add-in to Office. It contains information such as a unique identifier, name, what buttons to show on the ribbon, and more. Importantly, the manifest provides URL locations for where Office can find and download the add-in's resource files. The manifest and two icon files are combined into a zip package file that is sideloaded to Office.

### Web app

The Hello World sample implements a task pane in a file named **taskpane.html** that contains HTML and JavaScript. The **taskpane.html** file contains all the code necessary to display a task pane, interact with the user, and write "Hello world!" into a text box on a presentation slide.

### Initialize the Office JavaScript API library

The sample initializes the Office JavaScript API library with a call to `office.onReady()` in the **taskpane.html** file. This is required before you can make any calls to the Office JavaScript APIs. For more information about initialization, see [Initialize your Office Add-in](https://learn.microsoft.com/office/dev/add-ins/develop/initialize-add-in).

```javascript
Office.onReady((info) => {});
```

### Write to a PowerPoint object

When the user chooses the **Say hello** button, the `sayHello()` function is called. This function calls `Office.context.document.setSelectedDataAsync()` with an empty string (space). This will clear the text of the currently selected object. Then it calls `setSelectedDataAsync()` again and passes "Hello world!". The currently selected object will now display "Hello world!".

For more information see [Build your first PowerPoint task pane add-in](https://learn.microsoft.com/office/dev/add-ins/quickstarts/powerpoint-quickstart)

```javascript
async function sayHello() {
  // Set coercion type to text since
  const options = { coercionType: Office.CoercionType.Text };

  // clear current selection
  await Office.context.document.setSelectedDataAsync(' ', options);

  // Set text in selection to 'Hello world!'
  await Office.context.document.setSelectedDataAsync('Hello world!', options);
}
```

## Decide on a manifest type

There are two types of manifests for Office Add-ins. For more information about the differences between them, see [Office Add-ins manifest](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests).

- **Add-in only manifest**: By default, the sample supports the add-in only manifest. In the root of the sample, there are two versions of the add-in only manifest to support the two ways of hosting the web app part of the add-in: **manifest.xml** and **manifest-localhost.xml**. For convenience, a copy of the files needed for using the add-in only manifest can be found in the **manifest-configurations/add-in-only** subfolder.

   To work with the add-in only manifest continue with the [Use the add-in only manifest](#use-the-add-in-only-manifest) section.

- **Unified manifest for Microsoft 365**: To work the unified manifest (**manifest.json**), you need to copy all the files from the **manifest-configurations/unified** subfolder to the sample's root directory, replacing any existing files that have the same names. (We recommend that you also delete the **manifest.xml** and **manifest-localhost.xml** files from root directory, so only files needed for the unified manifest are present in the root.)

   To work with the unified manifest continue with the [Use the unified manifest](#use-the-unified-manifest) section.

   > **Note:** If you ever want to switch back to the add-in only manifest, copy the files in the **manifest-configurations/add-in-only** subfolder to the sample's root directory. We recommend that you delete the following files the root of the sample, so only files needed for the add-only manifest are present in the root.
   >
   > - **manifest.json**
   > - **package.json**
   > - **package-lock.json**
   > - **webpack.config.js**

### Use the add-in only manifest

#### Run the sample on PowerPoint on web

An Office Add-in requires you to configure a web server to provide all the resources, such as HTML, image, and JavaScript files. The hello world sample is configured so that the files are hosted directly from this GitHub repo. Use the following steps to sideload the manifest.xml file to see the sample run.

1. Download the **manifest.xml** file from the sample folder for PowerPoint.
1. Open [Office on the web](https://office.live.com/).
1. Choose **PowerPoint**, and then open a new blank presentation.
1. On the **Insert** tab on the ribbon in the **Add-ins** section, choose **Add-ins**.
1. On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.

    ![The Office Add-ins dialog with a drop-down in the upper right reading "Manage my add-ins" and a drop-down below it with the option "Upload My Add-in"](../images/office-add-ins-powerpoint-web.png)

1. Browse to the add-in manifest file, and then select **Upload**.

    ![The upload add-in dialog with buttons for browse, upload, and cancel.](../images/office-upload-add-ins-powerpoint-web.png)

1. Verify that the add-in loaded successfully. You will see a **Hello world** button on the **Home** tab on the ribbon.
1. Choose the **Hello world** button to display the task pane of the add-in.
1. Position your cursor in the Slide where you want to insert the text.
1. Choose the **Say Hello** button to insert "Hello world!" into the current PowerPoint slide.

#### Run the sample on PowerPoint on Windows or Mac

Office Add-ins are cross-platform so you can also run them on Windows, Mac, and iPad. The following links will take you to documentation for how to sideload on Windows, Mac, or iPad. Be sure you have a local copy of the **manifest.xml** file for the Hello world sample. Then follow the sideloading instructions for your platform.

- [Sideload Office Add-ins for testing from a network share](https://learn.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)
- [Sideload Office Add-ins on Mac for testing](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)
- [Sideload Office Add-ins on iPad for testing](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-an-office-add-in-on-ipad)

#### Configure a localhost web server and run the sample from localhost

If you prefer to configure a web server and host the add-in's web files from your computer, use the following steps:

1. Install a recent version of [npm](https://www.npmjs.com/get-npm) and [Node.js](https://nodejs.org/) on your computer. To verify if you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.
1. You need http-server to run the local web server. If you haven't installed this yet you can do this with the following command:

    ```console
    npm install --global http-server
    ```

1. You need Office-Addin-dev-certs to generate self-signed certificates to run the local web server. If you haven't installed this yet you can do this with the following command:

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
1. Run the following command:

    ```console
    http-server -S -C localhost.crt -K localhost.key --cors . -p 3000
    ```

    The http-server will run and host the current folder's files on localhost:3000.

Now that your localhost web server is running, you can sideload the **manifest-localhost.xml** file provided in the powerpoint-hello-world folder. Using the **manifest-localhost.xml** file, follow the steps in [Run the sample on PowerPoint on web](#run-the-sample-on-powerpoint-on-web) to sideload and run the add-in.

### Use the unified manifest

#### Run the sample with GitHub as the host

An Office Add-in requires you to configure a web server to provide all the resources, such as HTML, image, and JavaScript files. The Hello World sample is configured so that the files are hosted directly from this GitHub repo, so all you need to do is build the manifest and package, and then sideload the package. 

1. Clone or download this sample to a folder on your computer. Then in a command prompt, bash shell, or **TERMINAL** in Visual Studio Code, navigate to the root of the sample folder.
1. Run the command `npm install`.
1. Run the command `npm run build`.
1. Run the command `npm run start:prod`.

   After a few seconds, desktop PowerPoint opens, and after a few seconds more, a **Hello World** button appears on the right end of the **Home** ribbon. 

1. Choose the **Hello world** button to display the task pane of the add-in.
1. Position your cursor in the slide where you want to insert the text.
1. Choose the **Say Hello** button to insert "Hello world!" into the current PowerPoint slide.

When you're finished working with the add-in, close PowerPoint, and then in the window where you ran the three npm commands, run `npm run stop:prod`.

#### Configure a localhost web server and run the sample from localhost

If you prefer to configure a web server and host the add-in's web files from your computer, use the following steps.

1. Clone or download this sample to a folder on your computer. Then in a command prompt, bash shell, or **TERMINAL** in Visual Studio Code, navigate to the root of the sample folder.
1. Run the command `npm install`.
1. Run the command `npm start`.

   - If you've never developed an Office add-in on this computer before or it has been more than 30 days since you last did, you'll be prompted to delete an old security cert and/or install a new one. Agree to both prompts. 
   - After a few seconds a **webpack** dev-server window will open and your files will be hosted there on localhost:3000.
   - When the server is successfully running, desktop PowerPoint opens, and after a few seconds more, a **Hello World** button appears on the right end of the **Home** ribbon. 

1. Choose the **Hello world** button to display the task pane of the add-in.
1. Position your cursor in the slide where you want to insert the text.
1. Choose the **Say Hello** button to insert "Hello world!" into the current PowerPoint slide.

When you're finished working with the add-in, close PowerPoint, and then in the window where you ran the two npm commands, run `npm stop`.

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2021 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

**Note**: The taskpane.html file contains an image URL that tracks diagnostic data for this sample add-in. Please remove the image tag if you reuse this sample in your own code project.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/powerpoint-add-in-hello-world" />
