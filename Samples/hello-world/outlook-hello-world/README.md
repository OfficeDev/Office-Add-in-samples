---
page_type: sample
urlFragment: outlook-add-in-hello-world
products:
  - office-outlook
  - office
  - m365
languages:
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: '10/11/2021 10:00:00 AM'
description: 'Create a simple Outlook add-in that displays Hello World.'
---

# Create an Outlook add-in that displays "Hello World"

## Summary

Learn how to build the simplest Office Add-in with only a manifest, HTML web page, and a logo. This sample will help you understand the fundamental parts of an Office Add-in.

## Features

- Display "Hello World" in an Outlook email message.
- Learn fundamentals of the manifest.
- Learn how to initialize the Office JavaScript API library.
- Interact with message content through Office JavaScript APIs.

## Applies to

- Outlook on the web
- Outlook on Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic)
- Outlook on Mac

## Prerequisites

- Microsoft 365 - You can get a free developer sandbox by joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program#Subscription).

## Understand an Office Add-in

An Office Add-in is a web application that can extend Office with additional functionality for the user. For example, an add-in can add ribbon buttons, and a task pane with the functionality you want. Because an Office Add-in is a web application you must provide a web server to host the files.

The sample contained in this folder is designed to run in Outlook.

## Key components

The Hello World sample implements the **Manifest** and **Web app** components identified in [Components of an Office Add-in](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins#components-of-an-office-add-in).

### Manifest

The manifest file describes your add-in to Office. It contains information such as a unique identifier, name, what buttons to show on the ribbon, and more. Importantly, the manifest provides URL locations for where Office can find and download the add-in's resource files.

### Web app

The Hello World sample implements a task pane in a file named **taskpane.html** that contains HTML and JavaScript. The **taskpane.html** file contains all the code necessary to display a task pane, interact with the user, and write "Hello world!" into a new email message.

### Initialize the Office JavaScript API library

The sample initializes the Office JavaScript API library with a call to `office.onReady()` in the **taskpane.html** file. This is required before you can make any calls to the Office JavaScript APIs. For more information about initialization, see [Initialize your Office Add-in](https://learn.microsoft.com/office/dev/add-ins/develop/initialize-add-in).

```javascript
Office.onReady((info) => {});
```

### Write to the email message

When the user chooses the **Say hello** button from the task pane, the `sayHello()` function is called, as shown in the following code sample. This function then calls `Office.context.mailbox.item.body.setAsync()`, which is an Office JavaScript API. The `setAsync()` method overwrites the body of the message with "Hello world!". Then, it calls the anonymous callback method `function (asyncResult)`. Most Outlook functions in the Office JavaScript API use this callback pattern. In this sample, the callback method checks that the call was successful. If not, it writes an error message to the console.

```javascript
/**
 * Writes "Hello world!" to a new message body.
 */
function sayHello() {
  Office.context.mailbox.item.body.setAsync(
    "Hello world!",
    {
      coercionType: Office.CoercionType.Html, // Write text as HTML.
    },

    // Callback method to check that setAsync succeeded.
    function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write(asyncResult.error.message);
      }
    }
  );
}
```

For more information see [Build your first Outlook add-in](https://learn.microsoft.com/office/dev/add-ins/quickstarts/outlook-quickstart).

## Run the sample

An Office Add-in requires you to configure a web server to provide all the resources, such as HTML, image, and JavaScript files. To run the Hello World sample, use any of the add-in file hosting options applicable to your Outlook client.

- [Run in Outlook on the web or on Windows (new or classic)](#run-in-outlook-on-the-web-or-on-windows-new-or-classic)
- [Run in Outlook on Mac](#run-in-outlook-on-mac)

> [!NOTE]
> The unified manifest for Microsoft 365 is directly supported in Outlook on the web and on Windows (new and classic). For more information on manifests and their supported platforms, see [Office Add-in manifest](https://learn.microsoft.com/office/dev/add-ins/develop/add-in-manifests).

### Run in Outlook on the web or on Windows (new or classic)

#### Use GitHub as the web host

The Hello World sample is configured so that the add-in files are hosted directly from this GitHub repository.

1. Clone or download this sample to a folder on your computer. Then, in a command prompt, bash shell, or **TERMINAL** in Visual Studio Code, navigate to the root of the sample folder.
1. Run the command `npm install`.
1. Run the command `npm run build`.
1. Run the command `npm run start:prod`.

    After a few seconds, classic Outlook on Windows opens, and after a few seconds more, a **Hello World** button appears on the Message tab of the ribbon in Message Compose mode. The add-in is also sideloaded to other supported Outlook clients, such as Outlook on the web and new Outlook on Windows.

1. [Test the sample on Outlook](#test-the-sample-on-outlook).

When you're finished working with the add-in, close Outlook. Then, in the window where you ran the npm commands, run `npm run stop:prod`.

#### Use localhost

If you prefer to configure a web server and host the add-in's web files from your computer, use the following steps.

1. Clone or download this sample to a folder on your computer. Then in a command prompt, bash shell, or **TERMINAL** in Visual Studio Code, navigate to the root of the sample folder.
1. Run the command `npm install`.
1. Run the command `npm start`.

    - If you've never developed an Office add-in on this computer before or it has been more than 30 days since you last did, you'll be prompted to delete an old security certificate and/or install a new one. Agree to both prompts.
    - After a few seconds, a webpack dev-server window will open and your files will be hosted there on localhost:3000.
    - When the server is successfully running, classic Outlook on Windows opens, and after a few seconds, a **Hello World** button appears on the Message tab of the ribbon in Message Compose mode. The add-in is also sideloaded to other supported Outlook clients, such as Outlook on the web and new Outlook on Windows.

1. [Test the sample on Outlook](#test-the-sample-on-outlook).

When you're finished working with the add-in, close Outlook. Then, in the window where you ran the npm commands, run `npm stop`.

### Run in Outlook on Mac

#### Use GitHub as the web host

The Hello World sample is configured so that the add-in files are hosted directly from this GitHub repository.

1. Download the **manifest.xml** file from this sample to a folder on your computer.
1. Sideload the manifest by following the manual instructions in [Sideload Outlook add-ins for testing](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing?tabs=xmlmanifest#sideload-manually).

    The **Hello World** button appears on the Message tab of the ribbon in Message Compose mode. The add-in is also sideloaded to other supported Outlook clients, such as Outlook on the web and new Outlook on Windows.

1. Follow the steps in [Try it out](#try-it-out) to test the sample.

When you're finished working with the add-in, [remove the sideloaded add-in](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing?tabs=xmlmanifest#remove-a-sideloaded-add-in).

#### Use localhost

If you prefer to configure a web server and host the add-in's web files from your computer, use the following steps.

1. Clone or download this repository.
1. From a command prompt, run the following commands.

    ```console
    npm install
    npm run start:xml
    ```

    - If you've never developed an Office add-in on this computer before or it has been more than 30 days since you last did, you'll be prompted to delete an old security certificate and/or install a new one. Agree to both prompts.
    - After a few seconds, a webpack dev-server window will open and your files will be hosted there on localhost:3000.

1. After starting the server, sideload the manifest by following the manual instructions in [Sideload Outlook add-ins for testing](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing?tabs=xmlmanifest#sideload-manually).

    The **Hello World** button appears on the Message tab of the ribbon in Message Compose mode. The add-in is also sideloaded to other supported Outlook clients, such as Outlook on the web.

1. Follow the steps in [Try it out](#try-it-out) to test the sample.

When you're finished working with the add-in, close Outlook. Then, in the window where you ran the npm commands, run `npm run stop:xml`. Then, [remove the sideloaded add-in](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing?tabs=xmlmanifest#remove-a-sideloaded-add-in).

## Test the sample on Outlook

1. Create a new message.
1. Choose the **Hello world** button on the ribbon to see the add-in task pane with the text, "This add-in will insert the text 'Hello world!' in a new message."
1. Choose the **Say hello** button to insert "Hello world!" in the message body.

![A new email message in Outlook showing the Hello world button and task pane.](../images/outlook-for-windows-new-message.png)

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2021 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

**Note**: The **taskpane.html** file contains an image URL that tracks diagnostic data for this sample add-in. Please remove the image tag if you reuse this sample in your own code project.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/outlook-add-in-hello-world" />
