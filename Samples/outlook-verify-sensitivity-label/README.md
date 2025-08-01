---
title: "Verify the sensitivity label of a message"
page_type: sample
urlFragment: outlook-verify-sensitivity-label
products:
  - office-outlook
  - office
  - m365
  - office-teams
languages:
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: 04/18/2023 10:00:00 AM
description: "Learn how to verify and update the sensitivity label of a message using an event-based add-in."
---

# Verify the sensitivity label of a message

**Applies to**: Outlook on Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic) | Outlook on Mac | modern Outlook on the web

## Summary

This sample uses the sensitivity label API in an event-based add-in to verify and apply the **Highly Confidential** sensitivity label to outgoing messages that contain at least one attachment or a recipient who's a member of the fictitious Fabrikam legal team. When the **Highly Confidential** label is applied, a fictitious legal hold account is added to the **Bcc** field of the message.

![The Smart Alert dialog notifying the sender that the sensitivity label was updated to Highly Confidential because of the presence of an attachment or the email address of a Fabrikam legal team member in the message.](assets/outlook-verify-sensitivity-label.png)

## Features

- The sensitivity label API is used to:
  - Verify that the catalog of sensitivity labels is enabled on the mailbox where the add-in is installed.
  - Get the available sensitivity labels from the catalog.
  - Get the sensitivity label of a message.
  - Set the sensitivity label of a message.
- Event-based activation is used to handle the following events.
  - When the `OnMessageRecipientsChanged` event occurs, the add-in performs the following:
    - It checks for the legal hold account (`legalhold@fabrikam.com`). If the account is added to the **To** or **Cc** field, it's automatically removed. If the account is added to the **Bcc** field, the add-in verifies the message has the **Highly Confidential** sensitivity label and adds it if it isn't set.
    - It checks for an email address with the `-legal@fabrikam.com` format. If the address is present in the message, the add-in verifies the message has the **Highly Confidential** sensitivity label and adds it if it isn't set.
  - When the `OnSensitivityLabelChanged` event occurs, if the message has an attachment or a recipient who's a member of the Fabrikam legal team (`-legal@fabrikam.com`), the add-in verifies that the **Highly Confidential** label is set. It then adds the legal hold account, if applicable.
  - When the `OnMessageAttachmentsChanged` event occurs, if it contains at least one attachment, the add-in verifies that the **Highly Confidential** sensitivity label is set on the message.
  - When the `OnMessageSend` event occurs, if the **Highly Confidential** label is set and the legal hold account is in the **Bcc** field of the message, a Smart Alerts dialog is displayed to notify the user.

For documentation related to this sample, see the following:

- [Manage the sensitivity label of your message or appointment in compose mode](https://learn.microsoft.com/office/dev/add-ins/outlook/sensitivity-label)
- [Activate add-ins with events](https://learn.microsoft.com/office/dev/add-ins/develop/event-based-activation)
- [Handle OnMessageSend and OnAppointmentSend events in your Outlook add-in with Smart Alerts](https://learn.microsoft.com/office/dev/add-ins/outlook/onmessagesend-onappointmentsend-events)

## Applies to

- Outlook on the web (modern)
- classic Outlook on Windows starting in Version 2304 (Build 16327.20248)
- new Outlook on Windows
- Outlook on Mac starting in Version 16.77.816.0

## Prerequisites

- A Microsoft 365 E5 subscription. You can get a [free developer sandbox](https://aka.ms/m365/devprogram#Subscription) that provides a renewable 90-day Microsoft 365 E5 subscription for development purposes.
- An enabled catalog of sensitivity labels in Outlook that includes the **Highly Confidential** label. To learn how to configure the sensitivity labels in your tenant, see the following:
  - [Get started with sensitivity labels](https://learn.microsoft.com/microsoft-365/compliance/get-started-with-sensitivity-labels)
  - [Create and configure sensitivity labels and their policies](https://learn.microsoft.com/microsoft-365/compliance/create-sensitivity-labels)
  - [Default labels and policies to protect your data](https://learn.microsoft.com/microsoft-365/compliance/mip-easy-trials)
- (Optional) If you want to run the web server on localhost, install a recent version of [npm](https://www.npmjs.com/get-npm) and [Node.js](https://nodejs.org) on your computer. To check if you've already installed these tools, from a command prompt, run the following commands.

    ```console
    node -v
    npm -v
    ```

## Choose a manifest type

By default, the sample uses an add-in only manifest. However, you can switch the project between the add-in only manifest and the unified manifest for Microsoft 365. For more information about the differences between them, see [Office Add-ins manifest](https://learn.microsoft.com/office/dev/add-ins/develop/add-in-manifests). To continue with the add-in only manifest, skip ahead to the [Run the sample](#run-the-sample) section.

> [!NOTE]
> To run the sample in Outlook on the web, on Mac, or in the new Outlook on Windows, use the [add-in only manifest](#run-with-the-add-in-only-manifest). For more information on manifests and their supported platforms, see [Office Add-in manifest](https://learn.microsoft.com/office/dev/add-ins/develop/add-in-manifests). For information about events and their supported platforms, see [Activate add-ins with events](https://learn.microsoft.com/office/dev/add-ins/develop/event-based-activation#supported-events).

### To switch to the unified manifest for Microsoft 365

Copy all the files from the **manifest-configurations/unified** subfolder to the sample's root folder, replacing any existing files that have the same names. We recommend that you delete the **manifest.xml** and **manifest-localhost.xml** files from the root folder, so only files needed for the unified manifest are present. Then, [run the sample](#run-the-sample).

### To switch back to the add-in only manifest

To switch back to the add-in only manifest, copy the files from the **manifest-configurations/add-in-only** subfolder to the sample's root folder. We recommend that you delete the **manifest.json** file from the root folder.

## Run the sample

To run the sample, choose whether to host the add-in's web files on localhost or on GitHub.

### Use localhost

If you prefer to host the web server on localhost, follow these steps.

1. Clone or download this repository.
1. From a command prompt, go to the root of the project folder **/Samples/outlook-verify-sensitivity-label**.
1. Run the following commands.

    ```console
    npm install
    npm start
    ```

    This starts the web server on localhost and sideloads the manifest file.

1. Follow the steps in [Try it out](#try-it-out) to test the sample.

1. To stop the web server and uninstall the add-in from Outlook, run the following command.

    ```console
    npm stop
    ```

### Use GitHub

> [!NOTE]
> The option to use GitHub as the web host only applies to the add-in only manifest.

The quickest way to run the sample is to use GitHub as the web host. However, you can't debug or change the source code. The add-in web files are served from this GitHub repository.

1. Download the **manifest.xml** file from this sample to a folder on your computer.
1. Sideload the add-in only manifest by following the manual instructions in [Sideload Outlook add-ins for testing](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing?tabs=xmlmanifest#sideload-manually).
1. Follow the steps in [Try it out](#try-it-out) to test the sample.
1. To uninstall the add-in from Outlook, follow the instructions in [Remove a sideloaded add-in](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing?tabs=xmlmanifest#remove-a-sideloaded-add-in).

## Try it out

Once the add-in is loaded, use the following steps to try out its functionality.

1. Create a new message.
1. Do one of the following:

    - Add an attachment to the message.
    - Add the email address of a fictitious Fabrikam legal team member to the **To**, **Cc**, or **Bcc** field using the format, `-legal@fabrikam.com`. For example, `eli-legal@fabrikam.com`.

    The sensitivity label of the message is set to **Highly Confidential** and the `legalhold@fabrikam.com` account is added to the **Bcc** field.

1. (Optional) Add a subject or content to the body of the message.
1. Select **Send**.

    A Smart Alerts dialog appears that reads, "Due to the contents of your message, the sensitivity label has been set to Highly Confidential and the legal hold account has been added to the **Bcc** field. To learn more, see Fabrikam's information protection policy. Do you need to make changes to your message?"
1. If you're ready to send your message, select **Send anyway**. Otherwise, select **Don't send**.

   > **Note**: Sending a message to the fabrikam.com domain will result in an undeliverable message.

### Test changing the sensitivity label of a message

If you manually change the sensitivity label of a message to **Highly Confidential**, the `legalhold@fabrikam.com` account is automatically added to the **Bcc** field. Use the following steps to try out this functionality.

1. Create a new message.
1. Change the sensitivity label to **Highly Confidential**. For guidance on how to change the sensitivity label of a message, see [Apply sensitivity labels to your files and email](https://support.microsoft.com/office/2f96e7cd-d5a4-403b-8bd7-4cc636bae0f9).
1. If you're prompted with a **Justification Required** dialog, select the applicable option, then select **Change**.

    The sensitivity label of the message is set to **Highly Confidential** and the `legalhold@fabrikam.com` account is added to the **Bcc** field.

### Test removing the legal hold account from a Highly Confidential message

If you attempt to remove the `legalhold@fabrikam.com` account from a message that's labeled **Highly Confidential**, the account will be automatically re-added to the **Bcc** field. Use the following steps to try out this functionality.

1. Navigate to the message you previously created in [Test changing the sensitivity label of the message](#test-changing-the-sensitivity-label-of-a-message).
1. Navigate to the **Bcc** field and delete `legalhold@fabrikam.com`.

    The `legalhold@fabrikam.com` account is re-added to the **Bcc** field.

### Test manually adding the legal hold account as a recipient

In this sample, the `legalhold@fabrikam.com` account can only be added to the **Bcc** field when the sensitivity label of a message is set to **Highly Confidential**. Use the following steps to try out this functionality.

1. Create a new message.
1. Ensure that the sensitivity label is set to something other than **Highly Confidential**.
1. Add `legalhold@fabrikam.com` to the **To**, **Cc**, or **Bcc** field of the message.

    The `legalhold@fabrikam.com` account is automatically removed from the **To**, **Cc**, or **Bcc** field of the message.

## Key parts of the sample

### Configure event-based activation in the manifest

To use the sensitivity label API, the permission level of your add-in's manifest must be set to **read/write item**.

- **Unified manifest for Microsoft 365**: The ["name"](https://learn.microsoft.com/microsoft-365/extensibility/schema/root-authorization-permissions-resource-specific#name) property of the object in the ["authorization.permissions.resourceSpecific"](https://learn.microsoft.com/microsoft-365/extensibility/schema/root-authorization-permissions-resource-specific) array must be set to "MailboxItem.ReadWrite.User".

    ```json
    "authorization": {
      "permissions": {
        "resourceSpecific": [
          {
            "name": "MailboxItem.ReadWrite.User",
            "type": "Delegated"
          }
        ]
      }
    },
    ```

- **Add-in only manifest**: The [<Permissions\>](https://learn.microsoft.com/javascript/api/manifest/permissions) element must be set to **ReadWriteItem**.

    ```xml
    <Permissions>ReadWriteItem</Permissions>
    ```

The manifest configures the runtime to handle event-based activation. Because the Outlook platform uses the client to determine whether to use HTML or JavaScript to load the runtime, both of these files must be referenced in the manifest. Classic Outlook on Windows uses the referenced JavaScript file to load the runtime, while Outlook on the web, on Mac, and new Outlook on Windows use the HTML file. The runtime configuration varies depending on the manifest your add-in uses.

- **Unified manifest for Microsoft 365**: The ["extensions.runtimes.code"](https://learn.microsoft.com/microsoft-365/extensibility/schema/extension-runtime-code) property has a child ["page"](https://learn.microsoft.com/microsoft-365/extensibility/schema/extension-runtime-code#page) property that is set to the HTML file and a child ["script"](https://learn.microsoft.com/microsoft-365/extensibility/schema/extension-runtime-code#script) property that is set to the JavaScript file.

    ```json
    "runtimes": [
      {
        ...
        "id": "event_runtime",
        "type": "general",
        "code": {
            "page": "https://officedev.github.io/Office-Add-in-samples/Samples/outlook-verify-sensitivity-label/src/launchevent/launchevent.html",
            "script": "https://officedev.github.io/Office-Add-in-samples/Samples/outlook-verify-sensitivity-label/src/launchevent/launchevent.js"
        },
        ...
      },
      ...
    ]
    ```

- **Add-in only manifest**: The HTML page resource ID is specified in the [\<Runtime\>](https://learn.microsoft.com/javascript/api/manifest/runtime) element and a JavaScript file resource ID is specified in the [\<Override\>](https://learn.microsoft.com/javascript/api/manifest/override#override-element-for-runtime) element.

    ```xml
    <!-- HTML file that references the JavaScript event handlers. This is used by Outlook on the web and on Mac, and in new Outlook on Windows. -->
    <Runtime resid="WebViewRuntime.Url">
        <!-- JavaScript file that contains the event handlers. This is used by classic Outlook on Windows. -->
        <Override type="javascript" resid="JSRuntime.Url"/>
    </Runtime>
    ...
    <bt:Url id="JSRuntime.Url" DefaultValue="https://officedev.github.io/Office-Add-in-samples/Samples/outlook-verify-sensitivity-label/src/launchevent/launchevent.js"/>
    <bt:Url id="WebViewRuntime.Url" DefaultValue="https://officedev.github.io/Office-Add-in-samples/Samples/outlook-verify-sensitivity-label/src/launchevent/launchevent.html"/>
    ```

The manifest also maps the events that activate the add-in to the functions that handle each event.

- **Unified manifest for Microsoft 365**: The events and their handlers are specified in the ["extensions.autoRunEvents"](https://learn.microsoft.com/microsoft-365/extensibility/schema/extension-auto-run-events-array) array. The function name provided in the ["actionId"](https://learn.microsoft.com/microsoft-365/extensibility/schema/extension-auto-run-events-array-events#actionid) property must match the name used in the ["id"](https://learn.microsoft.com/microsoft-365/extensibility/schema/extension-runtimes-actions-item#id) property of the applicable object in the ["extensions.runtimes.actions"](https://learn.microsoft.com/microsoft-365/extensibility/schema/extension-runtimes-actions-item) array.

    ```json
    "autoRunEvents": [
      {
        ...
        "events": [
          {
            "type": "messageRecipientsChanged",
            "actionId": "onMessageRecipientsChangedHandler"
          },
          {
            "type": "messageSending",
            "actionId": "onMessageSendHandler",
            "options": {
              "sendMode": "promptUser"
            }
          },
          {
            "type": "sensitivityLabelChanged",
            "actionId": "onSensitivityLabelChangedHandler"
          },
          {
            "type": "messageAttachmentsChanged",
            "actionId": "onMessageAttachmentsChangedHandler"
          }
        ]
      }
    ],
    ```

- **Add-in only manifest**: The events and their handlers are specified in the [\<LaunchEvents\>](https://learn.microsoft.com/javascript/api/manifest/launchevents) element.

    ```xml
    <!-- Indicates on which events the add-in activates. -->
    <LaunchEvents>
        <LaunchEvent Type="OnMessageRecipientsChanged" FunctionName="onMessageRecipientsChangedHandler"/>
        <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="PromptUser"/>
        <LaunchEvent Type="OnSensitivityLabelChanged" FunctionName="onSensitivityLabelChangedHandler"/>
        <LaunchEvent Type="OnMessageAttachmentsChanged" FunctionName="onMessageAttachmentsChangedHandler"/>
    </LaunchEvents>
    ```

In this sample, the **prompt user** send mode option is implemented for the `OnMessageSend` event to notify the sender that the sensitivity label of a message has been updated to meet the company's data loss prevention policies. To learn more about send mode options, see [Available send mode options](https://learn.microsoft.com/office/dev/add-ins/outlook/onmessagesend-onappointmentsend-events#available-send-mode-options).

### Configure the event handlers

The event object is passed to its respective handler in the **launchevent.js** file for processing. To ensure that the event-based add-in runs in Outlook, the JavaScript file that contains your handlers (in this case, **launchevent.js**) must call `Office.actions.associate`. This method maps the function ID specified in the manifest to its respective event handler in the JavaScript file.

```javascript
/** 
 * Maps the event handler name specified in the manifest to its JavaScript counterpart.
 */
Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("onSensitivityLabelChangedHandler", onSensitivityLabelChangedHandler);
Office.actions.associate("onMessageAttachmentsChangedHandler", onMessageAttachmentsChangedHandler);
```

The handler calls the [event.completed](https://learn.microsoft.com/javascript/api/outlook/office.mailboxevent#outlook-office-mailboxevent-completed-member(1)) method to signify when it completes processing an event. In the `onMessageSendHandler` function, the `event.completed` call specifies the [allowEvent](https://learn.microsoft.com/javascript/api/outlook/office.smartalertseventcompletedoptions#outlook-office-smartalertseventcompletedoptions-allowevent-member) property to indicate whether the event can continue to execute or must terminate. It also specifies the [errorMessage](https://learn.microsoft.com/javascript/api/outlook/office.smartalertseventcompletedoptions#outlook-office-smartalertseventcompletedoptions-errormessage-member) property to display the Smart Alerts dialog to indicate that the sensitivity label was updated.

```javascript
event.completed({ allowEvent: false, errorMessage: "Due to the contents of your message, the sensitivity label has been set to Highly Confidential and the legal hold account has been added to the Bcc field.\nTo learn more, see Fabrikam's information protection policy.\n\nDo you need to make changes to your message?" });
```

### Call the sensitivity label API

The sensitivity label API methods can only be called in compose mode. Before the add-in can get or set the sensitivity label on a message, it calls [Office.context.sensitivityLabelsCatalog.getIsEnabledAsync](https://learn.microsoft.com/javascript/api/outlook/office.sensitivitylabelscatalog#outlook-office-sensitivitylabelscatalog-getisenabledasync-member(1)) to verify that the catalog of sensitivity labels is enabled on the mailbox. The catalog of sensitivity labels is configured by an organization's administrator. For more information, see [Learn about sensitivity labels](https://learn.microsoft.com/microsoft-365/compliance/sensitivity-labels).

```javascript
// Verifies that the catalog of sensitivity labels is enabled on the mailbox where the add-in is installed.
Office.context.sensitivityLabelsCatalog.getIsEnabledAsync({ asyncContext: event }, (result) => {
    const event = result.asyncContext;
    if (result.status === Office.AsyncResultStatus.Failed) {
        console.log("Unable to retrieve the status of the sensitivity label catalog.");
        console.log(`Error: ${result.error.message}`);
        event.completed();
        return;
    }

    ...
});
```

The [Office.context.mailbox.item.sensitivityLabel.getAsync](https://learn.microsoft.com/javascript/api/outlook/office.sensitivitylabel#outlook-office-sensitivitylabel-getasync-member(1)) method only returns the unique identifier (GUID) of the sensitivity label applied to the current message. To help determine the name of the label, the add-in first calls [Office.context.sensitivityLabelsCatalog.getAsync](https://learn.microsoft.com/javascript/api/outlook/office.sensitivitylabelscatalog#outlook-office-sensitivitylabelscatalog-getasync-member(1)). This method retrieves the sensitivity labels available to the mailbox in the form of [SensitivityLabelDetails](https://learn.microsoft.com/javascript/api/outlook/office.sensitivitylabeldetails) objects. These objects provide details about the labels, including their names.

```javascript
// Gets the sensitivity labels available to the mailbox.
Office.context.sensitivityLabelsCatalog.getAsync({ asyncContext: event }, (result) => {
    const event = result.asyncContext;
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.log("Unable to retrieve the catalog of sensitivity labels.");
      console.log(`Error: ${result.error.message}`);
      if (callback) {
        callback();
      } else {
        event.completed();
      }
      return;
    }

    // Gets the sensitivity label of the current message.
    const sensitivityLabelCatalog = result.value;
    Office.context.mailbox.item.sensitivityLabel.getAsync({ asyncContext: { event: event, sensitivityLabelCatalog: sensitivityLabelCatalog } }, (result) => {
      const event = result.asyncContext;
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log("Unable to get the sensitivity label of the message.");
        console.log(`Error: ${result.error.message}`);
        if (callback) {
          callback();
        } else {
          event.completed();
        }
        return;
      }

      ...
    });
});
```

To set the sensitivity label of a message to **Highly Confidential**, the add-in passes the label's unique identifier (GUID) as a parameter to [Office.context.mailbox.item.sensitivityLabel.setAsync](https://learn.microsoft.com/javascript/api/outlook/office.sensitivitylabel#outlook-office-sensitivitylabel-setasync-member(1)).

> **Tip**: When you test this sample and adopt it for your scenario, you can also pass the `SensitivityLabelDetails` object returned by `Office.context.sensitivityLabelsCatalog.getAsync` to the `setAsync` method.

```javascript
// Sets the sensitivity label of the message to Highly Confidential using the label's GUID.
Office.context.mailbox.item.sensitivityLabel.setAsync(labelId, { asyncContext: event }, (result) => {
    const event = result.asyncContext;
    if (result.status === Office.AsyncResultStatus.Failed) {
        console.log("Unable to set the Highly Confidential sensitivity label to the message.");
        console.log(`Error: ${result.error.message}`);
        if (callback) {
          callback();
        } else {
          event.completed();
        }
        return;
    }
    ...
});
```

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Solution

|Solution|Authors|
|--------|-------|
|Verify the sensitivity label of a message using an event-based add-in.|Microsoft|

## Version history

|Version|Date|Comments|
|-------|----|--------|
|1.0|April 18, 2023|Initial release|
|1.1|May 19, 2023|Update for General Availability (GA) of the sensitivity label API|
|1.2|October 12, 2023|Update supported version of Outlook on Mac|
|1.3|January 11, 2024|Remove Microsoft 365 Insider Program requirement|
|1.4|July 28, 2025|Add support for the unified manifest for Microsoft 365 and move add-in logic to other event handlers|

## Copyright

Copyright (c) 2023 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/outlook-verify-sensitivity-label" />
