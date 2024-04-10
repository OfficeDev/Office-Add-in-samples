---
title: "Report spam or phishing emails in Outlook (preview)"
page_type: sample
urlFragment: outlook-spam-reporting
products:
- office-add-ins
- office-outlook
- office
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 03/26/2024 10:00:00 AM
description: "Learn how to create an integrated spam-reporting add-in in Outlook."
---

# Report spam or phishing emails in Outlook (preview)

**Applies to**: Outlook on Windows

![A sample spam-reporting dialog.](./assets/readme/outlook-spam-processing-dialog.png)

## Summary

This sample showcases how to build an integrated spam-reporting solution that:

- Is easily discoverable in the Outlook client ribbon.
- Provides the user with a processing dialog to report an email.
- Facilitates saving a copy of the reported email to a file to submit it to your backend system for further processing.

To learn about key components of this sample, see [Implement an integrated spam-reporting add-in (preview)](https://learn.microsoft.com/office/dev/add-ins/outlook/spam-reporting).

> [!IMPORTANT]
> The integrated spam-reporting feature is currently in preview in Outlook on Windows. Features in preview shouldn't be used in production add-ins. We invite you to try out this feature in test or development environments and welcome feedback on your experience through [GitHub](https://github.com/OfficeDev/office-js/issues/new/choose).

## Applies to

Outlook on Windows starting in Version 2307 (Build 16626.10000).

> [!NOTE]
> If you don't have a Microsoft 365 subscription, you might qualify for a free developer subscription that's renewable for 90 days and comes configured with sample data. For details, see the [Microsoft 365 Developer Program FAQ](https://learn.microsoft.com/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-).

## Prerequisites

- Microsoft 365 subscription
- You must join the [Microsoft 365 Insider program](https://insider.microsoft365.com/join/windows) and select the **Beta Channel** option to access Office beta builds.

> [!TIP]
> If you're unable to choose a channel in your Outlook client, see [Let users choose which Microsoft 365 Insider channel to install on Windows devices](https://learn.microsoft.com/deployoffice/insider/deploy/user-choice).

> [!IMPORTANT]
> To test the `getAsFileAsync` method while it's still in preview in Outlook on Windows, you must configure your computer's registry.
>
> Outlook on Windows includes a local copy of the production and beta versions of Office.js instead of loading from the content delivery network (CDN). By default, the local production copy of the API is referenced. To reference the local beta copy of the API, you must configure your computer's registry as follows:
>
> 1. In the registry, navigate to `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer`. If the key doesn't exist, create it.
> 1. Create an entry named `EnableBetaAPIsInJavaScript` and set its value to `1`.
>
>    ![The EnableBetaAPIsInJavaScript registry value is set to 1.](https://learn.microsoft.com/office/dev/add-ins/images/outlook-beta-registry-key.png)

## Run the sample

---

Run this sample in Outlook on Windows using one of the following add-in file hosting options.

### Run the sample from GitHub

1. Download the **manifest.xml** file from this sample to a folder on your computer.
1. Sideload the add-in manifest in Outlook on Windows by following the manual instructions in [Sideload Outlook add-ins for testing](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing#sideload-manually).
1. Follow the steps in [Try it out](#try-it-out) to test the sample.

### Run the sample from localhost

If you prefer to host the web server for the sample on your computer, follow these steps.

1. Install a recent version of [npm](https://www.npmjs.com/get-npm) and [Node.js](https://nodejs.org/) on your computer. To verify if you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.
1. You need http-server to run the local web server. If you haven't installed this yet, run the following command.

   ```console
   npm install --global http-server
   ```

1. You need Office-Addin-dev-certs to generate self-signed certificates to run the local web server. If you haven't installed this yet, you can do this with the following command.

    ```console
    npm install --global office-addin-dev-certs
    ```

1. Clone or download this sample to a folder on your computer, then go to that folder in a console or terminal window.
1. Run the following command to generate a self-signed certificate to use for the web server.

   ```console
    npx office-addin-dev-certs install
    ```

    This command will display the folder location where it generated the certificate files.

1. Go to the folder location where the certificate files were generated, then copy the **localhost.crt** and **localhost.key** files to the cloned or downloaded sample folder.
1. Run the following command.

    ```console
    http-server -S -C localhost.crt -K localhost.key --cors . -p 3000
    ```

    The http-server will run and host the current folder's files on localhost:3000.

1. Now that your localhost web server is running, you can sideload the **manifest-localhost.xml** file provided in the sample folder. To sideload the manifest, follow the manual instructions in [Sideload Outlook add-ins for testing](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing?tabs=windows-web#sideload-manually).
1. Follow the steps in [Try it out](#try-it-out) to test the sample.

## Try it out

Once the add-in is loaded in Outlook, use the following steps to try out its functionality.

1. Choose a message from your inbox, then select the add-in's button from the ribbon.

    ![The spam-reporting add-in button is selected from the ribbon.](./assets//readme/outlook-spam-ribbon-button.png)
1. In the preprocessing dialog, choose a reason for reporting the message and add information about the message, if configured. Then, select **Report**.
1. (Optional) In the post-processing dialog, select **OK**.

    ![The post-processing dialog of the sample spam-reporting add-in.](./assets//readme/outlook-spam-post-processing-dialog.png)

## References

- [Implement an integrated spam-reporting add-in](https://learn.microsoft.com/office/dev/add-ins/outlook/spam-reporting)
- [ReportPhishingCommandSurface Extension Point](https://learn.microsoft.com/javascript/api/manifest/extensionpoint?view=outlook-js-preview&preserve-view=true#reportphishingcommandsurface-preview)
- [Office.MessageRead.getAsFileAsync() method](https://learn.microsoft.com/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#outlook-office-messageread-getasfileasync-member(1))
- [Troubleshoot event-based and spam-reporting add-ins](https://learn.microsoft.com/office/dev/add-ins/outlook/troubleshoot-event-based-and-spam-reporting-add-ins)
- [Debug your event-based or spam-reporting Outlook add-in](https://learn.microsoft.com/office/dev/add-ins/outlook/debug-autolaunch)
- [Microsoft Office Add-in Debugger Extension for Visual Studio Code](https://learn.microsoft.com/office/dev/add-ins/testing/debug-with-vs-extension)
- [Develop Office Add-ins with Visual Studio Code](https://learn.microsoft.com/office/dev/add-ins/develop/develop-add-ins-vscode)
- [Office Add-ins with Visual Studio Code](https://code.visualstudio.com/docs/other/office)
- [Debugging with Visual Studio Code](https://code.visualstudio.com/docs/editor/debugging)
- [Node.js debugging in VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging)
- [Office-Addin-Debugging](https://www.npmjs.com/package/office-addin-debugging)

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2024 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Solution

| Solution | Author(s) |
| ----- | ----- |
| Report spam or phishing emails in Outlook (preview) | [Eric Legault](https://www.linkedin.com/in/ericlegault/) |

## Version history

| Version | Date | Comments |
| ----- | ----- | ----- |
| 1.0 | March 26, 2024 | Initial release |

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/outlook-spam-reporting" />
