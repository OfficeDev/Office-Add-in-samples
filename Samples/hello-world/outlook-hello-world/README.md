---
page_type: sample
urlFragment: outlook-add-in-hello-world
products:
  - office-outlook
  - office
languages:
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: '10/11/2021 10:00:00 AM'
description: 'Create a simple Outlook add-in that displays hello world.'
---

# Create an Outlook add-in that displays hello world

## Summary

Learn how to build the simplest Office Add-in with only a manifest, HTML web page, and a logo. This sample will help you understand the fundamental parts of an Office Add-in.

## Features

- Display hello world in an Outlook email message.
- Learn fundamentals of the manifest.
- Learn how to initialize the Office JavaScript API library.
- Interact with message content through Office JavaScript APIs.

## Applies to

- Outlook on Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic), Mac, and in a browser.

## Prerequisites

- Microsoft 365 - You can get a [free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) that provides a renewable 90-day Microsoft 365 E5 developer subscription.

## Understand an Office Add-in

An Office Add-in is a web application that can extend Office with additional functionality for the user. For example, an add-in can add ribbon buttons, and a task pane with the functionality you want. Because an Office Add-in is a web application you must provide a web server to host the files.

The sample contained in this folder is a sample that is designed to run in Outlook.

## Key components

The hello world sample implements the **Manifest** and **Web app** components identified in [Components of an Office Add-in](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins#components-of-an-office-add-in).

### Manifest

The manifest file is an XML file that describes your add-in to Office. It contains information such as a unique identifier, name, what buttons to show on the ribbon, and more. Importantly the manifest provides URL locations for where Office can find and download the add-in's resource files.

The hello world sample contains two manifest files to support two different web hosting scenarios.

- **manifest.xml**: This manifest file gets the add-in's HTML page from the original GitHub repo location. This is the quickest way to try out the sample. To get started running the add-in with this manifest, see [Run the sample using GitHub as a web host](#run-the-sample-using-github-as-a-web-host).
- **manifest.localhost.xml**: This manifest file gets the add-in's HTML page from a local web server that you configure. Use this manifest if you want to change the code and experiment. For more information, see [Configure a localhost web server and run the sample from localhost](#configure-a-localhost-web-server-and-run-the-sample-from-localhost).

### Web app

The hello world sample implements a task pane named **taskpane.html** that contains HTML and JavaScript. The **taskpane.html** file contains all the code necessary to display a task pane, interact with the user, and write "Hello world!" into a new email message.

### Initialize the Office JavaScript API library

The sample initializes the Office JavaScript API library with a call to `office.onReady()` in the **taskpane.html** file. This is required before you can make any calls to the Office JavaScript APIs. For more information about initialization, see [Initialize your Office Add-in](https://learn.microsoft.com/office/dev/add-ins/develop/initialize-add-in).

```javascript
Office.onReady((info) => {});
```

### Write to the email message

When the user chooses the **Say hello** button, the `sayHello()` function is called as shown in the following code sample. This function then calls `Office.context.mailbox.item.body.setAsync()` which is an Office JavaScript API. The `setAsync()` method overwrites the body of the message with "Hello world!". Then it calls the anonymous callback method `function (asyncResult)`. Most Outlook functions in the Office JavaScript API use this callback pattern. In this sample, the callback method checks that the call was successful. If not it writes an error message to the console.

```javascript
/**
 * Writes 'Hello world!' to a new message body.
 */
function sayHello() {
  Office.context.mailbox.item.body.setAsync(
    'Hello world!',
    {
      coercionType: 'html', // Write text as HTML
    },

    // Callback method to check that setAsync succeeded
    function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write(asyncResult.error.message);
      }
    }
  );
}
```

For more information see [Build your first Outlook add-in](https://learn.microsoft.com/office/dev/add-ins/quickstarts/outlook-quickstart)

## Run the sample

An Office Add-in requires you to configure a web server to provide all the resources, such as HTML, image, and JavaScript files. Select one of the following options to run the hello world sample.

### Run the sample using GitHub as a web host

The hello world sample is configured so that the files are hosted directly from this GitHub repo.

1. Download the **manifest.xml** file from this sample to a folder on your computer.
1. Sideload the manifest in Outlook on Windows (new or classic), on Mac, or on the web by following the instructions in [Sideload Outlook add-in on Windows or Mac](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing).

### Configure a localhost web server and run the sample from localhost

If you prefer to configure a web server and host the add-in's web files from your computer, use the following steps.

1. Install a recent version of [npm](https://www.npmjs.com/get-npm) and [Node.js](https://nodejs.org/) on your computer. To verify if you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.

1. You need http-server to run the local web server. If you haven't installed this yet you can do this with the following command:

   ```console
   npm install --global http-server
   ```

1. You need Office-Addin-dev-certs to generate self-signed certificates to run the local web server. If you haven't installed this yet you can do this with the following command:

   ```console
   npm install --global office-addin-dev-certs
   ```

1. Clone or download this sample to a folder on your computer. Then, go to that folder in a console or terminal window.
1. Run the following command to generate a self-signed certificate to use for the web server.

   ```console
   npx office-addin-dev-certs install
   ```

   This command will display the folder location where it generated the certificate files.

1. Go to the folder location where the certificate files were generated. Copy the **localhost.crt** and **localhost.key** files to the cloned or downloaded sample folder.

1. Run the following command.

   ```console
   http-server -S -C localhost.crt -K localhost.key --cors . -p 3000
   ```

   The http-server will run and host the current folder's files on `localhost:3000`.

1. Now that your localhost web server is running, you can sideload the **manifest-localhost.xml** file provided in the **outlook-hello-world** folder. Using the **manifest-localhost.xml** file, follow the steps in [Sideload Outlook add-in on Windows or Mac](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing) to sideload and run the add-in.

## Test the sample on Outlook

1. Verify that the add-in loaded successfully. You'll see a **Hello World** button on the Message tab of the ribbon.
1. Choose the **Hello World** button on the ribbon to see the add-in task pane with the text, "This add-in will insert the text 'Hello world!' in a new message."
1. Choose the **Say hello** button to insert "Hello world!" in the message body.

![A new email message in Outlook showing the hello world button and task pane.](../images/outlook-for-windows-new-message.png)

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2021 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

**Note**: The **taskpane.html** file contains an image URL that tracks diagnostic data for this sample add-in. Please remove the image tag if you reuse this sample in your own code project.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/outlook-add-in-hello-world" />
