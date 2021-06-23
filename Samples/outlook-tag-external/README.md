---
page_type: sample
products:
- office-outlook
- office
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 06/30/2021 10:00:00 AM
description: "Use Outlook event-based activation to tag external recipients."
---

# Use Outlook event-based activation to tag external recipients (preview)

**Applies to:** Outlook on Windows | Outlook on the web

## Summary

This sample uses event-based activation to run an Outlook add-in when the user changes recipients while composing a message. The add-in also uses the [appendOnSendAsync API](https://docs.microsoft.com/javascript/api/outlook/office.body?view=outlook-js-preview#appendOnSendAsync_data__options__callback_). If external recipients are added, the add-in prepends "[External]" to the message subject and appends a disclaimer to the message body on send.

For documentation related to this sample, see [Configure your Outlook add-in for event-based activation](https://docs.microsoft.com/office/dev/add-ins/outlook/autolaunch).

> **Note:** Features used in this sample are currently in preview and subject to change. They are not currently supported for use in production environments. To try the preview features, you'll need to [join Office Insider](https://insider.office.com/join). A good way to try out preview features is to sign up for a Microsoft 365 subscription. If you don't already have a Microsoft 365 subscription, get one by joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/office/dev-program). For more information on how to use events currently in preview, see the [How to preview](https://docs.microsoft.com/office/dev/add-ins/outlook/autolaunch#how-to-preview) section of the event-based activation article.

## Features

- Use event-based activation to respond to changes in message recipients during compose mode.
- Update the message subject to indicate there are external recipients.
- Add a disclaimer if the message is being sent to external recipients.

## Applies to

- Outlook
  - Windows
  - web browser

## Prerequisites

- To use this sample, you'll need to follow the instructions in the [How to preview](https://docs.microsoft.com/office/dev/add-ins/outlook/autolaunch#how-to-preview) section of the event-based activation article.
- Before running this sample, you need a recent version of [npm](https://www.npmjs.com/get-npm) and [Node.js](https://nodejs.org/) installed on your computer. To verify if you've already installed these tools, run the commands `node -v` and `npm -v` in your command prompt.

## Solution

| Solution | Authors |
|---------|----------|
| Use Outlook event-based activation to tag a message with external recipients | Microsoft |

## Version history

Version  | Date | Comments
|---------|------|---------|
| 1.0 | 6-30-2021 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

----------

## Scenario: Event-based activation

In this scenario, the add-in helps the user indicate if their message has recipients external to their organization by prepending "[External]" to the message subject. When the user sends an email message that includes external recipients, the add-in appends a disclaimer to the message.

## Build and run the solution

1. Clone or download this repository.
1. In the command line, go to the **outlook-tag-external** folder from your root directory.
1. Run the following command to download the dependencies required to run the sample.

    ```command&nbsp;line
    npm install
    ```

1. Run the following command to start the localhost web server.

    ```command&nbsp;line
    npm run dev-server
    ```

1. Sideload the add-in to Outlook on Windows, or Outlook on the web by following the manual instructions in the article [Sideload Outlook add-ins for testing](https://docs.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing).

Once the add-in is loaded, use the following steps to try out the functionality.

1. Open Outlook on Windows or in a browser.
1. Create a new message.
1. Add a recipient email address that's external to your organization.

    > Notice that "[External]" is inserted at the beginning of the subject.

1. Send the email.

    > Navigate to your **Sent Items** folder, open the email you sent, and notice the included disclaimer.

## Key parts of this sample

### Configure event-based activation and AppendOnSend in the manifest

The manifest configures a runtime that is loaded specifically to handle event-based activation. The following `<Runtime>` element specifies an HTML page resource ID that loads the runtime in Outlook on the web. The `<Override>` element specifies the JavaScript file to load the runtime for Outlook on Windows because Outlook on Windows doesn't use the HTML page to load the runtime.

```xml
<Runtime resid="WebViewRuntime.Url">
  <Override type="javascript" resid="JSRuntime.Url"/>
</Runtime>
...
<bt:Url id="WebViewRuntime.Url" DefaultValue="https://localhost:3000/commands.html" />
<bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/src/commands/commands.js" />
```

The add-in handles the `OnMessageRecipientsChanged` event that is mapped to the `tagExternal_onMessageRecipientsChangedHandler` function in the JavaScript file.

```xml
<LaunchEvents>
  <LaunchEvent Type="OnMessageRecipientsChanged" FunctionName="tagExternal_onMessageRecipientsChangedHandler" />
</LaunchEvents>
```

Since the add-in calls `Office.context.mailbox.item.body.appendOnSendAsync`, the `AppendOnSend` permission is declared in the manifest.

```xml
<ExtendedPermission>AppendOnSend</ExtendedPermission>
```

### Handle the event and call the appendOnSendAsync API

When the user creates a new message (including replies and forwards) and changes any recipients, Outlook will load the files specified in the manifest to handle the `OnMessageRecipientsChanged` event. Outlook on the web will load the `commands.html` page, which then also loads `commands.js`.

The `commands.js` file contains the `tagExternal` function that handles the `OnMessageRecipientsChanged` event from Outlook. In Outlook on Windows, this file is loaded directly and `commands.html` is not loaded.

Also, the `commands.js` file contains the `_tagExternal` helper function that calls the [appendOnSendAsync() API](https://docs.microsoft.com/javascript/api/outlook/office.body?view=outlook-js-preview#appendOnSendAsync_data__options__callback_) to set or clear the disclaimer as appropriate.

Note that you can use a different pattern to handle events if needed. For example, if you need code that applies only to Outlook on the web but other code that applies to both Outlook on the web and on Windows, you can define separate JavaScript files. For a sample using this pattern, see [Use Outlook event-based activation to set the signature](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-set-signature).

## Security notes

In the webpack.config.js file, a header is set to `"Access-Control-Allow-Origin": "*"`. This is only for development purposes. In production code, you should list the allowed domains and not leave this header open to all domains.

You'll be prompted to install certificates for trusted access to https://localhost. The certificates are intended only for running and studying this code sample. Do not reuse them in your own code solutions or in production environments.

Install or uninstall the certificates by running the following commands in the project folder.

```command&nbsp;line
npx office-addin-dev-certs install
npx office-addin-dev-certs uninstall
```

## Copyright

Copyright (c) 2021 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/samples/outlook-autorun-tag-external" />
