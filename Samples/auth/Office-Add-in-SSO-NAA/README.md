---
page_type: sample
urlFragment: office-add-in-sso-naa
products:
  - m365
  - office
  - office-excel
  - office-word
  - office-powerpoint
languages:
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: "03/19/2024 10:00:00 AM"
description: "This sample shows how to implement SSO in an Office Add-in by using nested app authentication."
---

# Office Add-in with SSO using nested app authentication (preview)

## Summary

This sample shows how to use MSAL.js nested app authentication (NAA) in an Office Add-in to access Microsoft Graph APIs for the signed in user. The sample displays the signed in user's name and email. It also inserts the names of files from the user's Microsoft OneDrive account into the document.

> [!IMPORTANT]
> Nested app authentication is currently in preview. To try this feature, you need to join the [Microsoft 365 Insider Program](https://insider.microsoft365.com/join) and choose **Current Channel (Preview)**. Don't use NAA in production add-ins. We invite you to try out NAA in test or development environments and welcome feedback on your experience through GitHub (see https://github.com/OfficeDev/office-js/issues).

## Features

- Use MSAL.js NAA to get an access token to call Microsoft Graph APIs.
- Use MSAL.js NAA to get information about the user signed in to Office.

## Applies to

For a list of supported hosts, see [NAA supported accounts and hosts](https://learn.microsoft.com/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in#naa-supported-accounts-and-hosts).

## Prerequisites

- Office connected to a Microsoft 365 subscription (including Office on the web).
- You need to join the [Microsoft 365 Insider Program](https://insider.microsoft365.com/join) to use the NAA preview features. Choose the **Current Channel (Preview)** insider level.
- [Node.js](https://nodejs.org/) version 16 or greater.
- [npm](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm) version 8 or greater.

## Build and run the solution

### Create an application registration

1. Go to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page to register your app.
1. Sign in with the ***admin*** credentials to your Microsoft 365 tenancy. For example, **MyName@contoso.onmicrosoft.com**.
1. Select **New registration**. On the **Register an application** page, set the values as follows.

    - Set **Name** to `Office-Add-in-SSO-NAA`.
    - Set **Supported account types** to **Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)**.
    - In the **Redirect URI** section, ensure that **Single-page application (SPA)** is selected in the drop down and then set the URI to `brk-multihub://localhost:3000`.
    - Select **Register**.

1. On the **Office-Add-in-SSO-NAA** page, copy and save the value for the **Application (client) ID**. You'll use it in the next section.

For more information on how to register your application, see [Register an application with the Microsoft Identity Platform](https://learn.microsoft.com/graph/auth-register-app-v2).

### Configure the sample

1. Clone or download this repository.
1. From the command line, or a terminal window, go to the root folder of this sample at `/samples/auth/Office-Add-in-SSO-NAA`.
1. Open the `src/taskpane/authConfig.ts` file.
1. Replace the placeholder "Enter_the_Application_Id_Here" with the Application ID that you copied.
1. Save the file.

## Choose a manifest type

By default, the sample uses an add-in only manifest. However, you can switch the project between the add-in only manifest and the unified manifest. For more information about the differences between them, see [Office Add-ins manifest](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests).
If you want to continue with the add-in only manifest, skip ahead to the [Run the sample](#run-the-sample) section.

### To switch to the Unified manifest for Microsoft 365

Copy all files from the **manifest-configurations/unified** subfolder to the sample's root folder, replacing any existing files that have the same names. We recommend that you delete the **manifest.xml** file from the root folder, so only files needed for the unified manifest are present in the root. Then continue with the [Run the sample](#run-the-sample) section.

### To switch back to the Add-in only manifest

If you want to switch back to the add-in only manifest, copy the files in the **manifest-configurations/add-in-only** subfolder to the sample's root folder. We recommend that you delete the **manifest.json** file from the root folder.

## Run the sample

1. Run the following commands.

`npm install`
`npm run start`

This will start the web server and sideload the add-in to Excel.

1. In Excel, look for the **Show task pane** button and select it.
1. When the task pane opens, there are two buttons: **Get user data** and **Get user files**.
1. To see the signed in user's name and email, select **Get user data**.
1. To insert the first 10 filenames from the signed in user's Microsoft OneDrive, select **Get user files**.

You'll be prompted to consent to the scopes the sample needs when you select the buttons.  

## Selecting hosts and debugging steps

If you want to choose Word or PowerPoint, modify the `start` command in the `package.json` file to match one of the following entries.

- For Word: `"start": "office-addin-debugging start manifest.xml desktop --app word",`
- For PowerPoint: `"start": "office-addin-debugging start manifest.xml desktop --app powerpoint",`

You can also debug the sample by opening the project in VS Code.

1. Select the **Run and Debug** icon in the **Activity Bar** on the side of VS Code. You can also use the keyboard shortcut <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>D</kbd>.
1. Select the launch configuration you want from the **Configuration dropdown** in the **Run and Debug** view. For example, **Word Desktop (Edge Chromium)**.
1. Start your debug session with **F5**, or **Run** > **Start Debugging**.

![VS Code debug view](./assets/vs-code-debug-view.png)

For more information on debugging with VS Code, see [Debugging](https://code.visualstudio.com/Docs/editor/debugging). For more information on debugging Office Add-ins in VS Code, see [Debug Office Add-ins on Windows using Visual Studio Code and Microsoft Edge WebView2 (Chromium-based)](https://learn.microsoft.com/office/dev/add-ins/testing/debug-desktop-using-edge-chromium)

## Key parts of this sample

The `src/taskpane/authConfig.ts` file contains the MSAL code for configuring and using NAA. It contains a class named AccountManager which manages getting user account and token information.

- The `initialize` function is called from Office.onReady to configure and initialize MSAL to use NAA.
- The `ssoGetToken` function gets an access token for the signed in user to call Microsoft Graph APIs.
- The `ssoGetUserIdentity` function gets the account information of the signed in user. This can be used to get user details such as name and email.

The `src/taskpane/document.ts` file contains code to write a list of file names, retrieved from Microsoft Graph, into the document. This works for Word, Excel, and PowerPoint documents.

The `src/taskpane/taskpane.ts` file contains code that runs when the user chooses buttons in the task pane. They use the AccountManager class to get tokens or user information depending on which button is chosen.

The `src/taskpane/msgraph-helper.ts` file contains code to construct and make a REST call to the Microsoft Graph API.

## Security reporting

If you find a security issue with our libraries or services, report the issue to [secure@microsoft.com](mailto:secure@microsoft.com) with as much detail as you can provide. Your submission may be eligible for a bounty through the [Microsoft Bounty](https://aka.ms/bugbounty) program. Don't post security issues to [GitHub Issues](https://github.com/AzureAD/microsoft-authentication-library-for-android/issues) or any other public site. We'll contact you shortly after receiving your issue report. We encourage you to get new security incident notifications by visiting [Microsoft technical security notifications](https://technet.microsoft.com/security/dd252948) to subscribe to Security Advisory Alerts.

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2024 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/office-add-in-sso-naa" />
