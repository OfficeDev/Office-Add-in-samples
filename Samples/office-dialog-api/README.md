---
title: Open dialogs in Office Add-ins
page_type: sample
urlFragment: office-dialog-api
products:
  - office
  - office-excel
  - office-outlook
  - office-powerpoint
  - office-word
  - m365
languages:
  - typescript
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: "11/25/2026 12:00:00 AM"
description: "Open a number picker from a task pane or function command in Excel, Outlook, PowerPoint, and Word."
---

# Open dialogs in Office Add-ins

Use the Office Dialog API to collect a number in a separate window and return it to your add-in.

- Open pop-up and iframe dialogs from a task pane.
- Open a dialog directly from a ribbon function command.
- Return the selected number to the task pane, document selection, or Outlook message notification.
- Handle dialog creation errors, page errors, and closure.

## Applies to

The sample supports Excel, Outlook, PowerPoint, and Word on Windows, Mac, and the web. The following table describes client support for the manifests included in this sample. Excel, PowerPoint, and Word also support the add-in only manifest, but this sample provides an add-in only manifest only for Outlook.

| Office client | Add-in only manifest | Unified manifest for Microsoft 365 |
| --- | --- | --- |
| Excel, PowerPoint, and Word on Windows | Not included in this sample | Version 2501 (Build 18407.20002) or later |
| Excel, PowerPoint, and Word on Mac | Not included in this sample | Version 16.103 (25101427) or later |
| Excel, PowerPoint, and Word on the web | Not included in this sample | Supported |
| Classic Outlook on Windows | Supported with Mailbox 1.4 | Version 2307 (Build 16626.20132) or later |
| New Outlook on Windows and Outlook on the web | Supported | Supported |
| Outlook on Mac | Supported with Mailbox 1.4 | Not supported; use the add-in only manifest |

Excel, PowerPoint, and Word require **DialogApi 1.1**, **TextCoercion 1.1**, and **AddinCommands 1.1** for the complete workflow. Outlook requires **Mailbox 1.4** and runs in message read mode. Mobile clients aren't targeted by these manifests.

The unified manifest uses the most recent generally available schema version. Its client requirements apply in addition to the API requirements. For more information, see the following:

- [Unified manifest support](https://learn.microsoft.com/office/dev/add-ins/develop/unified-manifest-overview#client-and-platform-support)
- [Dialog API requirement sets](https://learn.microsoft.com/javascript/api/requirement-sets/common/dialog-api-requirement-sets)
- [TextCoercion requirement set](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/common/office-add-in-requirement-sets#textcoercion)
- [Add-in commands requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/common/add-in-commands-requirement-sets)
- [Outlook requirement sets](https://learn.microsoft.com/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)

## Prerequisites

- An active LTS release of [Node.js](https://nodejs.org/) and npm.
- A supported Office client and an account that permits sideloading add-ins.
- For Outlook, a Microsoft 365 account with an Exchange Online mailbox and a message to open.

## Choose a manifest type

By default, the sample uses the unified manifest for Microsoft 365, which supports Excel, Outlook, PowerPoint, and Word. You can also run the sample in Outlook with the add-in only manifest. For more information about the differences between them, see [Office Add-ins manifest](https://learn.microsoft.com/office/dev/add-ins/develop/add-in-manifests). To continue with the unified manifest, skip ahead to the [Run the sample](#run-the-sample) section.

For Outlook, install only one manifest configuration at a time to avoid duplicate ribbon buttons. The configurations have distinct add-in IDs but share the same source code and local HTTPS server. The unified manifest requests `Document.ReadWrite.User` to insert text at the selection and `MailboxItem.Read.User` for Outlook item notifications. The Outlook add-in only manifest requests `ReadItem` to show item notifications without modifying the message.

> [!NOTE]
> The unified manifest for Microsoft 365 isn't directly supported in Outlook on Mac. Run the sample with the add-in only manifest instead. For more information about clients and platforms supported by the unified manifest, see [Office Add-ins with the unified app manifest for Microsoft 365](https://learn.microsoft.com/office/dev/add-ins/develop/unified-manifest-overview#client-and-platform-support).

### To switch to the add-in only manifest

Remove any sideloaded unified-manifest installation from Outlook, then follow [Start Outlook with the add-in only manifest](#start-outlook-with-the-add-in-only-manifest), using **manifest-configurations/add-in-only/manifest-outlook.xml**.

### To switch back to the unified manifest for Microsoft 365

Remove any sideloaded add-in only installation from Outlook, then follow [Start with the unified manifest](#start-with-the-unified-manifest).

## Run the sample

### Install dependencies

1. Clone or download this repository.
1. Open a terminal in **Samples/office-dialog-api** and run:

   ```console
   npm install
   ```

### Start with the unified manifest

1. Start the unified-manifest configuration using **manifest.json** in the sample's root folder:

   ```console
   npm start
   ```

   This command starts HTTPS hosting, sideloads the add-in to all supported Office hosts, and opens Word by default. If prompted, trust the development certificate.

   To open a particular desktop Office host instead, specify it with `--app`:

   ```console
   npm start -- --app excel
   ```

   Replace `excel` with `outlook`, `powerpoint`, or `word` as needed. The `--` separator passes `--app` and its value to the Office debugging tool. This selects the application to open; it doesn't restrict which supported hosts can use the sideloaded add-in. Outlook on Mac requires the add-in only manifest described in [Start Outlook with the add-in only manifest](#start-outlook-with-the-add-in-only-manifest).

   To test in Office on the web, run the applicable command to sideload the add-in in your preferred desktop application. Once sideloaded, the add-in also appears on the web client.

### Start Outlook with the add-in only manifest

The add-in only manifest supports Outlook on Windows, Mac, and the web.

1. Start the local HTTPS server.

   ```console
   npm run start:outlook
   ```

1. Follow [Sideload Outlook add-ins for testing](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing?tabs=xmlmanifest) for your Outlook client, using **manifest-configurations/add-in-only/manifest-outlook.xml**.

### Open dialogs from a task pane

After sideloading the add-in with either manifest type, open the task pane for your Office host.

- In Excel, PowerPoint, or Word, select **Home** > **Dialog API sample** > **Dialog from Task Pane**.
- In Outlook, open an existing email message, then open **Office dialog API** and choose **Dialog from Task Pane** from the message's apps menu or ribbon, depending on your Outlook client. The sample's commands aren't available while composing a message.

With the task pane open, follow these steps in any supported Office host.

1. Select **Pick a number (pop-up)**.
1. Select **1** or **2**. The dialog closes, and **Picked 1** or **Picked 2** appears in the task pane.
1. Repeat with **Pick a number (iframe)**. The `displayInIframe` option applies to Office on the web; desktop clients display a separate dialog.
1. Open a dialog, then select either button from the task pane again. The task pane reports that a dialog is already open. Choose a number in the original dialog to confirm that it still works.
1. Open another dialog and close it. The task pane displays **Dialog closed.**

### Open a dialog from a function command

1. Prepare the destination for the result.
   - In Excel, select one editable cell.
   - In Outlook, open an existing message. The result appears as a notification above the message; its body is unchanged.
   - In PowerPoint, create a slide and place the cursor inside a text box.
   - In Word, place the cursor in an editable document.
1. Close the task pane, then choose **Dialog from Function**.
1. Choose **1** or **2**. The dialog closes, and the result is inserted at the document selection or shown in the Outlook message notification.
1. Run the command again and close the dialog. **Dialog closed.** appears in the same destination.

The command completes after the result is written. If the selection isn't editable or a notification can't be displayed, the command logs the Office error in its runtime console and still completes. Dialog creation and navigation errors are reported in the task pane or command destination.

### Stop and remove the sample

Run the command that matches the configuration you started.

- **Unified manifest**: `npm stop`
- **Add-in only manifest**: `npm run stop:outlook`

## Key parts of the sample

### Open a dialog from either runtime

When you select a task-pane button or function command, [dialog-service.ts](src/dialog/dialog-service.ts) opens the same number picker with [`displayDialogAsync`](https://learn.microsoft.com/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)). The initial page is resolved relative to the parent page, so it remains on the same origin even when deployed under a subpath.

The task pane offers pop-up and iframe options. The function command uses the iframe option where supported. Both use a separate dialog runtime rather than reading or changing the Office document inside the dialog. For the platform rules, see [Use the Office dialog API](https://learn.microsoft.com/office/dev/add-ins/develop/dialog-api-in-office-add-ins).

### Return a choice and close the dialog

After Office initializes in the picker, [dialog.ts](src/dialog/dialog.ts) uses [`messageParent`](https://learn.microsoft.com/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) to send **Picked 1** or **Picked 2** to its parent. The parent registers `DialogMessageReceived` with [`Dialog.addEventHandler`](https://learn.microsoft.com/javascript/api/office/office.dialog#office-office-dialog-addeventhandler-member(1)), validates the selection and available origin information, and calls [`Dialog.close`](https://learn.microsoft.com/javascript/api/office/office.dialog#office-office-dialog-close-member(1)).

Each opened dialog owns its handlers. A failed second open doesn't replace the first dialog's reference, and duplicate terminal events don't repeat the result or complete a command twice.

### Display the result in the current application

The task pane displays the returned message in an accessible status region. For function commands, [commands.ts](src/commands/commands.ts) dispatches on the Office host: Excel, PowerPoint, and Word share [`setSelectedDataAsync`](https://learn.microsoft.com/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) with text coercion. Outlook instead uses [`notificationMessages.replaceAsync`](https://learn.microsoft.com/javascript/api/outlook/office.notificationmessages#outlook-office-notificationmessages-replaceasync-member(1)).

The function command keeps its event alive while the dialog is open, then calls `event.completed()` after the asynchronous document write or notification finishes. Both manifest formats associate the same `showDialog` handler with their function command.

### Distinguish errors from closure

A failed `displayDialogAsync` call returns an error before a dialog exists. Once it opens, `DialogEventReceived` reports page loading errors, HTTPS errors, or closure. Both paths provide a readable message. See [Handle dialog errors and events](https://learn.microsoft.com/office/dev/add-ins/develop/dialog-handle-errors-events) and [Dialog API best practices](https://learn.microsoft.com/office/dev/add-ins/develop/dialog-best-practices).

## Solution

| Solution | Author(s) |
| --- | --- |
| Open dialogs in Office Add-ins | Microsoft |

## Version history

| Version | Date | Comments |
| ------- | ---- | -------- |
|1.0 | November 25, 2026 | Initial release in this repository with Excel, Outlook, PowerPoint, and Word support and both manifest configurations. |

## Additional resources

- [Office Dialog API requirement sets](https://learn.microsoft.com/javascript/api/requirement-sets/common/dialog-api-requirement-sets)
- [Add-in commands](https://learn.microsoft.com/office/dev/add-ins/design/add-in-commands)
- [Manage unified and add-in only manifest versions](https://learn.microsoft.com/office/dev/add-ins/concepts/duplicate-legacy-metaos-add-ins)

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2026 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/office-dialog-api" />
