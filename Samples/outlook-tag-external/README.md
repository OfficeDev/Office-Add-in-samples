---
title: "Identify and tag external recipients using Outlook event-based activation"
page_type: sample
urlFragment: outlook-add-in-tag-external-recipients
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
  createdDate: 07/06/2021 2:00:00 PM
description: "Use Outlook event-based activation to tag external recipients."
---

# Identify and tag external recipients using Outlook event-based activation

**Applies to:** Outlook on Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic) | Outlook on the web

## Summary

This sample uses event-based activation to run an Outlook add-in when the user changes recipients while composing a message. The add-in also uses the [appendOnSendAsync API](https://learn.microsoft.com/javascript/api/outlook/office.body?view=outlook-js-1.11#appendOnSendAsync_data__options__callback_). If external recipients are added, the add-in prepends "[External]" to the message subject and appends a disclaimer to the message body on send.

![Screen shot of PnP sample displaying an information bar prompting the user to set up signatures, and sample signature inserted into the email.](./assets/outlook-tag-external-overview.png)

For documentation related to this sample, see [Configure your Outlook add-in for event-based activation](https://learn.microsoft.com/office/dev/add-ins/outlook/autolaunch).

## Features

- Use event-based activation to respond to changes in message recipients during compose mode.
- Update the message subject to indicate there are external recipients.
- Add a disclaimer to messages sent to external recipients.

## Applies to

- Outlook
  - Windows (new and classic)
  - web browser

## Prerequisites

- Microsoft 365

    > **Note**: If you don't have a Microsoft 365 subscription, you might qualify for a Microsoft 365 E5 developer subscription for development purposes through the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram); for details, see the [FAQ](https://learn.microsoft.com/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-). Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g).

## Solution

| Solution | Authors |
|---------|----------|
| Use Outlook event-based activation to tag a message with external recipients | Microsoft |

## Version history

| Version | Date | Comments |
|---------|------|---------|
| 1.0 | 7-6-2021 | Initial release |
| 1.1 | 11-1-2021 | Update for GA of SessionData API and OnMessageRecipientsChanged event |

----------

## Scenario: Event-based activation

In this scenario, if the message has external recipients, the add-in prepends "[External]" to the message subject. When the user sends an email message that includes external recipients, the add-in appends a disclaimer to the message.

## Run the sample

You can run this sample in Outlook on Windows (new or classic) or in a browser. The add-in web files are served from this repo on GitHub.

1. Download the **manifest.xml** file from this sample to a folder on your computer.
1. Sideload the add-in manifest in Outlook on the web or on Windows (new or classic) by following the manual instructions in the article [Sideload Outlook add-ins for testing](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing).

### Try it out

Once the add-in is loaded, use the following steps to try out the functionality.

1. Open Outlook on Windows (new or classic) or in a browser.
1. Create a new message.
1. Add a recipient email address that's external to your organization.

    > Notice that "[External]" is inserted at the beginning of the subject.

1. Send the email.

    > Navigate to your **Sent Items** folder, open the email you sent, and notice the included disclaimer.

## Run the sample from localhost

If you prefer to host the web server for the sample on your computer, follow these steps:

1. You need http-server to run the local web server. If you haven't installed this yet, run the following command.

    ```console
    npm install --global http-server
    ```

1. Use a tool such as openssl to generate a self-signed certificate that you can use for the web server. Move the cert.pem and key.pem files to the root folder for this sample.
1. From a command prompt, go to the root folder and run the following command.

    ```console
    http-server -S --cors . -p 3000
    ```

1. To reroute to localhost, run office-addin-https-reverse-proxy. If you haven't installed this, run the following command.

    ```console
    npm install --global office-addin-https-reverse-proxy 
    ```

    To reroute, run the following in another command prompt.

    ```console
    office-addin-https-reverse-proxy --url http://localhost:3000 
    ```

1. Sideload `manifest-localhost.xml` in Outlook on the web or on Windows (new or classic) by following the manual instructions in the article [Sideload Outlook add-ins for testing](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing).
1. [Try out the sample!](#try-it-out)

## Configure event-based activation and AppendOnSend in the manifest

The manifest configures a runtime that is loaded specifically to handle event-based activation. The following `<Runtime>` element specifies an HTML page resource ID that loads the runtime in Outlook on the web and new Outlook on Windows. The `<Override>` element specifies the JavaScript file to load the runtime for classic Outlook on Windows because the client doesn't use the HTML page to load the runtime.

```xml
<Runtime resid="WebViewRuntime.Url">
  <Override type="javascript" resid="JSRuntime.Url"/>
</Runtime>
...
<bt:Url id="WebViewRuntime.Url" DefaultValue="https://officedev.github.io/Office-Add-in-samples/Samples/outlook-tag-external/src/commands.html" />
<bt:Url id="JSRuntime.Url" DefaultValue="https://officedev.github.io/Office-Add-in-samples/Samples/outlook-tag-external/src/commands/commands.js" />
```

The add-in handles the `OnMessageRecipientsChanged` event that is mapped to the `tagExternal_onMessageRecipientsChangedHandler` function in the `commands.js` file.

```xml
<LaunchEvents>
  <LaunchEvent Type="OnMessageRecipientsChanged" FunctionName="tagExternal_onMessageRecipientsChangedHandler" />
</LaunchEvents>
```

```js
Office.actions.associate("tagExternal_onMessageRecipientsChangedHandler", tagExternal_onMessageRecipientsChangedHandler);
```

Since the add-in calls `Office.context.mailbox.item.body.appendOnSendAsync`, the `AppendOnSend` extended permission is declared in the manifest.

```xml
<ExtendedPermission>AppendOnSend</ExtendedPermission>
```

## Handle the OnMessageRecipientsChanged event, manage session data, and call the appendOnSendAsync API

When the user composes a message (including replies and forwards) and changes any recipients, Outlook will load the files specified in the manifest to handle the `OnMessageRecipientsChanged` event. Outlook on the web and new Outlook on Windows load the **commands.html** page, which then also loads **commands.js**. In classic Outlook on Windows, **commands.js** is loaded directly but **commands.html** is not loaded.

The **commands.js** file contains the `tagExternal_onMessageRecipientsChangedHandler` function that handles the `OnMessageRecipientsChanged` event from Outlook.

Also, the **commands.js** file contains the following helper functions.

- `checkForExternalTo`: Determines if there are any external users in the **To** field then sets a [SessionData](https://learn.microsoft.com/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11#sessionData) key named **tagExternalTo**.
- `checkForExternalCc`: Determines if there are any external users in the **Cc** field then sets a [SessionData](https://learn.microsoft.com/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11#sessionData) key named **tagExternalCc**.
- `checkForExternalBcc`: Determines if there are any external users in the **Bcc** field then sets a [SessionData](https://learn.microsoft.com/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11#sessionData) key named **tagExternalBcc**.
- `_checkForExternal`: Checks if any property is set to `true` in the [SessionData](https://learn.microsoft.com/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11#sessionData) property bag.
- `_tagExternal`:
  - Updates the [subject](https://learn.microsoft.com/javascript/api/outlook/office.messagecompose?view=outlook-js-1.11#subject) to prepend or remove the "[External]" tag.
  - Calls the [appendOnSendAsync](https://learn.microsoft.com/javascript/api/outlook/office.body?view=outlook-js-1.11#appendOnSendAsync_data__options__callback_) to set or clear the disclaimer.

> **Note**
>
> You can use a different pattern to handle events if needed. For example, if you need code that applies only to Outlook on the web and new Outlook on Windows, you can define separate JavaScript files. For a sample using this pattern, see [Use Outlook event-based activation to set the signature](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-set-signature).

## Known issues

- In classic Outlook on Windows, the `OnMessageRecipientsChanged` event occurs again when sending a reply or reply-all message, even if there are no further changes to the recipients.

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2021 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/outlook-autorun-tag-external" />
