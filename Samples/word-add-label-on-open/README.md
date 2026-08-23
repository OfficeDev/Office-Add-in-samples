---
page_type: sample
urlFragment: word-add-label-on-open
products:
  - office-word
  - office
  - m365
languages:
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: 06/30/2025 4:00:00 PM
description: "Shows how to configure a Word add-in to activate when a document opens."
---

# Automatically add labels with an add-in when any Word document opens

## Summary

This sample shows how to configure an add-in to automatically run when any Word document opens. It adds a header to indicate the content's sensitivity. 

> **Note**: Once this add-in is installed, it works on every document created or opened in Word. This is different from the per-document system of running code when a document opens that is described in [Run code in your Office Add-in when the document opens](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/run-code-on-document-open).

## Description

The add-in acts when the `OnDocumentOpened` event occurs. The `changeHeader` function is a JavaScript event handler for this event. It adds either a "Public" header to new documents or a "Highly Confidential" header to existing documents that already have content. Some of the functionality is duplicated in the task pane to allow for manual changes.

This sample is designed for Word, but the event-based activation parts will also work for Excel and PowerPoint.

The sample is configured to use the [unified manifest for Microsoft 365](https://learn.microsoft.com/office/dev/add-ins/develop/unified-manifest-overview) (**manifest.json**), and it also ships the equivalent add-in only manifest (**manifest.xml**). See [Choose a manifest type](#choose-a-manifest-type).

### Event-based activation deployment limitations

Event-based add-ins work only when deployed by a Microsoft 365 administrator. If users install them directly from AppSource or the Office Store, they will not automatically launch. Moreover, they cannot be sideloaded.

## Applies to

- Word on the web
- Word on Windows

## Prerequisites

- Office connected to a Microsoft 365 subscription (including Office on the web).
- [Node.js](https://nodejs.org/) (latest recommended version).
- [npm](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm) version 8 or greater.

## Solution

| Solution | Authors |
|----------|-----------|
| How to configure a Word add-in to activate when any document opens. | Microsoft |

## Version history

| Version  | Date | Comments |
|----------|------|----------|
| 1.0 | 06-30-2025 | Initial release |
| 1.1 | 08-21-2026 | Converted the project to the unified manifest for Microsoft 365 |

## Choose a manifest type

By default, the sample uses the unified manifest for Microsoft 365 (**manifest.json**). However, you can switch the project between the unified manifest and the add-in only manifest (**manifest.xml**). For more information about the differences between them, see [Office Add-ins manifest](https://learn.microsoft.com/office/dev/add-ins/develop/add-in-manifests). To continue with the unified manifest, skip ahead to the [Install the add-in](#install-the-add-in) section.

### To switch to the add-in only manifest

Copy all the files from the **manifest-configurations/add-in-only** subfolder to the sample's root folder, replacing any existing files that have the same names. We recommend that you delete the **manifest.json** file from the root folder, so only files needed for the add-in only manifest are present. Then, continue with [Install the add-in](#install-the-add-in).

### To switch back to the unified manifest for Microsoft 365

To switch back to the unified manifest, copy the files from the **manifest-configurations/unified** subfolder to the sample's root folder. We recommend that you delete the **manifest.xml** file from the root folder.

## Install the add-in

You must install the add-in in the Microsoft 365 Admin Portal. It can't be sideloaded.

1. Clone or download this repo.
1. Navigate to the **Office-Add-in-samples\Samples\word-add-label-on-open** folder in a command prompt, terminal, or bash shell.
1. Continue with the section below for your type of manifest.

### Install the unified manifest version

1. Using any zip utility, create an app package zip file that contains the manifest.json and the two icon files specified in the "icons" property of the manifest. The icon files must have the same relative path in the zip file as specified in the manifest. Since the path of the two image files is assets/icon-192.png and assets/icon-32.png, then you must include an assets folder with the two files in the zip package. 
1. Sign in as an admin to your Microsoft 365 tenancy.
1. In the admin portal, expand the **Settings** section in the navigation pane then select **Integrated apps**.
1. On the **Integrated apps** page, choose the **Upload custom apps** action.
1. In the **App type** drop down box, select **Teams app**. *Not* **Office Add-in**!
1. Use the file chooser control to navigate to and select the app package zip file.
1. Follow the instructions on the page to complete the installation.

For more information about how to deploy an add-in, please refer to [Deploy and publish Office Add-ins in the Microsoft 365 admin center](https://learn.microsoft.com/microsoft-365/admin/manage/office-addins).

### Install the add-in only manifest version

1. Sign in as an admin to your Microsoft 365 tenancy.
1. In the admin portal, expand the **Settings** section in the navigation pane then select **Integrated apps**.
1. On the **Integrated apps** page, choose the **Upload custom apps** action. 
1. In the **App type** drop down box, select **Office Add-in**.
1. Use the file chooser control to navigate to and select the manifest.
1. Follow the instructions on the page to complete the installation.

For more information about how to deploy an add-in, please refer to [Deploy and publish Office Add-ins in the Microsoft 365 admin center](https://learn.microsoft.com/microsoft-365/admin/manage/office-addins).

## Start the sample

After you have installed the add-in, you can test it in Word on the web right away. You many need to wait as much as 24 hours to test it on Word on Windows. It is not supported on Mac or mobile.

### Test on the web

1. Navigate to the **Office-Add-in-samples\Samples\word-add-label-on-open** folder in a command prompt, terminal, or bash shell.
1. Run `npm install`.
1. Run `npm run build:dev`. This will create a **\dist** folder and put the **commands.js** file there. The **webpack.config.js** file gives this folder the alias **public**, which is where the manifest has told Word to expect it.
1. Run `npm run dev-server`. This launches the server. 
1. In a browser, open Word in your Microsoft 365 tenant, and then open a blank document. A header should immediately be inserted specifying that the document has "Public" sensitivity. If this doesn't happen with the very first document, select the **File** tab in Word on the web and open another blank document.
1. Select the **File** tab in Word on the web and open a document that already has content. A header should immediately be inserted specifying that the document has "Highly Confidential" sensitivity.
1. Select the **My add-in** button on the **Event-based add-in activation** in the **Home** tab to open the task pane.
1. Select any of the sensitivity levels that are listed in the task pane to replace the header.

> **Note**: When you are finished testing, run `npm stop` to stop the server. 

### Test on Windows

It may take up to 24 hours for your Word on Windows client to synchronize with app catalog in your Microsoft 365 tenant. 

1. Cycle through the following deployment check steps periodically. 

    1. Close Word if it's running, and then immediately reopen it.
    1. Navigate to the **Office-Add-in-samples\Samples\word-add-label-on-open** folder in a command prompt, terminal, or bash shell.
    1. If you haven't already, run `npm install`.
    1. Run `npm run signin`. You will be prompted to sign in to your Microsoft 365 tenant, even if you are already signed in. Follow the prompts to sign in.
    1. Run `npm run build:dev`. This will create a **\dist** folder and put the **commands.js** file there. The **webpack.config.js** file gives this folder the alias **public**, which is where the manifest has told Word to expect it.
    1. Run `npm run dev-server`. This launches the server.
    1. Open Word, and then open a blank document.
    1. The last deployment check step depends on whether the **Event-based add-in activation** group is on the **Home** tab of the ribbon or the add-in is listed in the flyout that opens when you click the **Add-ins** button.
    
      - If the group isn't on the ribbon and the add-in isn't listed on the flyout, run `npm stop`. Repeat these deployment check steps in a little while.
       - If the group isn't on the **Home** tab, but the add-in is listed on the flyout, select it on the flyout to add it to the ribbon. Then go on to the next major step.
       - If the group is on the ribbon, go on to the next major step. 

    > **Note**: If the add-in isn't installed in Word within 24 hours, force a synchronization with the Integrated Apps catalog: Open Word on your computer. Open the Admin portal and on the **Integrated Apps** page, open the add-ins listing. Edit the add-ins properties to change what users it is available to. For example, if it is initially available only to one user, change that to make it available to a group. Wait a minute and then reverse the change. Repeat the deployment check steps.

1. The handler for the **OnDocumentOpened** event doesn't run on the first document you open in Word on Windows. Select the **File** tab in Word and open another blank document. A header should immediately be inserted specifying that the document has "Public" sensitivity. 
1. Select the **File** tab in Word on the web and open a document that already has content. A header should immediately be inserted specifying that the document has "Highly Confidential" sensitivity.
1. Select the **My add-in** button on the **Event-based add-in activation** in the **Home** tab to open the task pane.
1. Select any of the sensitivity levels that are listed in the task pane to replace the header.

> **Note**: When you are finished testing, run `npm stop` to stop the server. 

## Make it yours

The following are a few suggestions for how you could tailor this to your scenario.

- Add more complex logic to categorize the headers based on the content of the file.
- Apply the `OnDocumentOpened` event logic to an Excel or PowerPoint add-in.

## Related content

- [Activate add-ins with events](https://learn.microsoft.com/office/dev/add-ins/develop/event-based-activation)
- [Debug event-based or spam-reporting add-ins](https://learn.microsoft.com/office/dev/add-ins/testing/debug-autolaunch)
- [Office Add-ins with the unified app manifest for Microsoft 365](https://learn.microsoft.com/office/dev/add-ins/develop/unified-manifest-overview)
- [Convert an add-in to use the unified manifest for Microsoft 365](https://learn.microsoft.com/office/dev/add-ins/develop/convert-xml-to-json-manifest)
- [Microsoft 365 app manifest reference](https://learn.microsoft.com/microsoft-365/extensibility/schema/root)
- [Word add-ins documentation](https://learn.microsoft.com/office/dev/add-ins/word/)

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2025 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/word-add-in-label-on-open" />
