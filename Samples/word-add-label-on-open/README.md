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

# Automatically add labels with an add-in when a Word document opens

## Summary

This sample shows how to configure an add-in to automatically run when a Word document opens. It adds a header to indicate the content's sensitivity.

## Description

The add-in acts when the `OnDocumentOpened` event occurs. The `changeHeader` function is a JavaScript event handler for this event. It adds either a "Public" header to new documents or a "Highly Confidential" header to existing documents that already have content. Some of the functionality is duplicated in the task pane to allow for manual changes.

This sample is designed for Word, but the event-based activation parts will also work for Excel and PowerPoint.

The sample is configured to use the [unified manifest for Microsoft 365](https://learn.microsoft.com/office/dev/add-ins/develop/unified-manifest-overview) (**manifest.json**), and it also ships the equivalent add-in only manifest (**manifest.xml**). See [Choose a manifest type](#choose-a-manifest-type).

### Event-based activation deployment limitations

Event-based add-ins work only when deployed by an administrator. If users install them directly from AppSource or the Office Store, they will not automatically launch. Moreover, for add-ins that handle the `OnDocumentOpened` event, *the auto-open feature won't work on desktop, only in Office on the web.* Other features of the add-in do work on desktop.

For the purposes of this sample, sideloading the manifest with the script `npm start` sideloads the add-in in Word desktop and in Word on the web But the auto-open feature won't work on desktop, only in Office on the web. 

If you prefer to perform an admin deployment rather than sideload, take the following steps in the Microsoft 365 admin center.

1. In the admin portal, expand the **Settings** section in the navigation pane then select **Integrated apps**.
1. On the **Integrated apps** page, choose the **Upload custom apps** action.
1. The next steps depend on what manifest is being used. 

    - **Unified manifest for Microsoft 365**: 
        1. In the **App type** drop down box, select **Teams app**. *Not* **Office Add-in**!
        1. Use the file chooser control to navigate and select the app package zip file.
        1. Follow the instructions on the page to complete the installation.

    - **Add-in only manifest**: 
        1. In the **App type** drop down box, select **Office Add-in**.
        1. Use the file chooser control to navigate and select the manifest.
        1. Follow the instructions on the page to complete the installation.

For more information about how to deploy an add-in, please refer to [Deploy and publish Office Add-ins in the Microsoft 365 admin center](https://learn.microsoft.com/microsoft-365/admin/manage/office-addins).

## Applies to

- Word on the web

## Prerequisites

- Office connected to a Microsoft 365 subscription (including Office on the web).
- To sideload the unified manifest, Office on Windows Version 2501 (18407.20002) or later. For details, see [Sideload Office Add-ins that use the unified manifest for Microsoft 365](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-add-in-with-unified-manifest).
- [Node.js](https://nodejs.org/) (latest recommended version).
- [npm](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm) version 8 or greater.

## Solution

| Solution | Authors |
|----------|-----------|
| How to configure a Word add-in to activate when a document opens. | Microsoft |

## Version history

| Version  | Date | Comments |
|----------|------|----------|
| 1.0 | 06-30-2025 | Initial release |
| 1.1 | 08-21-2026 | Converted the project to the unified manifest for Microsoft 365 |

## Choose a manifest type

By default, the sample uses the unified manifest for Microsoft 365 (**manifest.json**). However, you can switch the project between the unified manifest and the add-in only manifest (**manifest.xml**). For more information about the differences between them, see [Office Add-ins manifest](https://learn.microsoft.com/office/dev/add-ins/develop/add-in-manifests). To continue with the unified manifest, skip ahead to the [Run the sample](#run-the-sample) section.

### To switch to the add-in only manifest

Copy all the files from the **manifest-configurations/add-in-only** subfolder to the sample's root folder, replacing any existing files that have the same names. We recommend that you delete the **manifest.json** file from the root folder, so only files needed for the add-in only manifest are present. Then, [run the sample](#run-the-sample).

### To switch back to the unified manifest for Microsoft 365

To switch back to the unified manifest, copy the files from the **manifest-configurations/unified** subfolder to the sample's root folder. We recommend that you delete the **manifest.xml** file from the root folder.

## Start the sample

1. Clone or download this repo.

1. Go to the **Samples\word-add-label-on-open** folder via the command line.

1. Run `npm install`.

1. Close Word if it's running, then run `npm start` to build the project, sideload the add-in, and launch the web server. Sideloading the unified manifest on Windows also makes the add-in available in Office on the web.

    The command creates the app package (a zip file that contains **manifest.json** and the two icon files referenced by the manifest's `"icons"` property) and installs it for you.

1. It may take as much as three minutes to sideload. Word desktop will open. *Close it. The add-in's auto-open feature doesn't work on desktop clients.*

## Try it out

1. *In Word on the web*, try opening both new and existing Word documents. Headers should automatically be added when they open. A new or empty document gets the "Public" header, and a document that already has content gets the "Highly Confidential" header.

    If no header is added, see [Event-based activation deployment limitations](#event-based-activation-deployment-limitations).

1. On the **Home** tab, in the **Event-based add-in activation** group, select **My add-in** to open the task pane.

1. Select any sensitivity level in the task pane, and verify that the corresponding label replaces the header that the event handler added.

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
