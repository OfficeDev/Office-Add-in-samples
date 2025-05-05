---
page_type: sample
urlFragment: office-add-in-auth-graph-react
products:
  - m365
  - office
  - office-excel
  - office-powerpoint
  - office-word
  - ms-graph
languages:
  - javascript
description: Learn how to build a Microsoft Office Add-in, as a single-page application with no backend, that connects to Microsoft Graph, finds the first three workbooks stored in OneDrive for Business, fetches their filenames, and inserts the names into an Office document using Office.js.
extensions:
  contentType: samples
  technologies:
    - Add-ins
    - Microsoft Graph
  services:
    - Excel
    - Microsoft 365
  createdDate: 5/1/2017 2:09:09 PM
---
# Get OneDrive data using Microsoft Graph and msal.js in an Office Add-in

## Summary

Learn how to build a Microsoft Office Add-in, as a single-page application (SPA) with no backend, that connects to Microsoft Graph, finds the first three workbooks stored in OneDrive for Business, fetches their filenames, and inserts the names into an Office document using Office.js.

## Features

Integrating data from online service providers increases the value and adoption of your add-ins. This code sample shows you how to connect your SPA add-in to Microsoft Graph. Use this code sample to:

* Connect to Microsoft Graph from an Office Add-in.
* Use the MSAL.js Library to implement the OAuth 2.0 authorization framework in an add-in, using the Auth Code Flow w/ PKCE for SPAs.
* Use the OneDrive REST APIs from Microsoft Graph.
* Show a dialog using the Office UI namespace.
* Build an Add-in using React, MSAL.js, and Office.js.
* Use add-in commands in an add-in.

## Applies to

* Excel on Windows (one-time purchase and subscription)

## Prerequisites

To run this code sample, the following are required.

* [Node and npm](https://nodejs.org/en/), version 18.20.2 or later (npm version 10.5.0 or later).

* TypeScript version 5.4.3 or later.

* A Microsoft 365 account. You can get one if you qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram); for details, see the [FAQ](https://learn.microsoft.com/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-). Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g).

* At least three Excel workbooks stored on OneDrive for Business in your Office 365 subscription.

* Office on Windows, version 16.0.6769.2001 or higher.

* A Microsoft Azure Tenant. This add-in requires Azure Active Directory (AD). Azure AD provides identity services that applications use for authentication and authorization. A trial subscription can be acquired here: [Microsoft Azure](https://account.windowsazure.com/SignUp).

* A code editor.

## Solution

Solution | Author(s)
---------|----------
Office Add-in Microsoft Graph React | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.1  | December 10th, 2020 | Upgrade MSAL.js to v2
1.0  | August 29th, 2019| Initial release
1.1  | January 14th, 2021| Changed system for creating and installing the SSL certificates for HTTPS.
1.2  | April 4th, 2024 | Updated to MSAL 3.7.1. Refactored code.

----------

## Build and run the solution

### Create an application registration

1. Navigate to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page to register your app.
1. Sign in with the ***admin*** credentials to your Microsoft 365 tenancy. For example, MyName@contoso.onmicrosoft.com.
1. Select **New registration**. On the **Register an application** page, set the values as follows.

    * Set **Name** to `ExcelGraphDemo`.
    * Set **Supported account types** to **Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)**.
    * In the **Redirect URI** section, ensure that **Single-page application (SPA)** is selected in the drop down and then set the URI to `https://localhost:3000/login/login.html`.
    * Select **Register**.

    For more information on how to register your application, see [Register an application with the Microsoft Identity Platform](https://learn.microsoft.com/graph/auth-register-app-v2).

    > Note: The sample uses the OAuth 2.0 Auth Code Flow w/ PKCE for SPAs, which requires no secrets.

1. On the **ExcelGraphDemo** page, copy and save the value for the **Application (client) ID**. You'll use it in the next section.

### Configure the sample

1. In a code editor, open the `/login/login.ts` file in the project. Near the top is a configuration property called `clientId`. Replace the `YOUR APP ID HERE` placeholder value with the application ID you copied in the previous step. Save and close the file.
1. Open the `/logout/logout.ts` file in the project. Near the top is a configuration property called `clientId`. Replace the `YOUR APP ID HERE` placeholder value with the application ID you copied in the previous step. Save and close the file.
1. Open a **Command Prompt** *as an administrator*.
1. Navigate to the root of the sample, which would normally be `[PATH-TO-YOUR-PROJECTS]\Office-Add-in-samples\Samples\auth\Office-Add-in-Microsoft-Graph-React`.
1. Run the command `npm install`.
1. Run the command: ```npx office-addin-dev-certs install --machine```.

    If you get the following prompt, click **Yes**.

    <img src="ReadmeImages/CertificateWarningPrompt.png" alt="Screenshot of a dialog that warns about the SSL certificate and asks user to accept or deny installation of it" />

    > Note: If you have worked with another Office Add-in within the last 30 days that was originally created with the Yo Office tool, you may have unexpired certs for localhost already, in which case you will get a message saying that localhost is already trusted. If so, continue with the next section.

### Run the solution

1. In the command prompt, run the command `start npm start`. This will open a second command prompt, build the project and then start a server (with dev mode settings). It takes from 5 to 30 seconds. When it finishes, the last line should say `Compiled successfully`. Minimize this command prompt.
1. Back in the original command prompt, run the command `npm run sideload`. This will launch Excel and install the add-in in it. After a few seconds, a **OneDrive Files** group appears on the right end of the **Home** ribbon with a button named **Open Add-in**.
1. Click the **Open Add-in** to open the task pane add-in.
1. The pages and buttons in the add-in are self-explanatory. 

    > Note: The first time that you press the **Connect to Office 365** button and sign in, you will be prompted to consent to the add-in.

## Known issues

* When a dialog is opened (with either the **Connect to Office 365** or the **Sign out from Office 365** buttons) on a Windows computer, a process named **Desktop App Web Viewer** starts on the computer. (You can see it in **Task Manager**.) These processes don't always close when the dialog closes. If you are working with the sample a lot, opening and closing dialogs, these processes use more and more memory. Eventually, the login dialog will start to flicker and seem to reload itself over and over. If this happens, use  **Task Manager** to kill the processes.

## Questions and feedback

* Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
* We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
* For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Additional resources

* [Microsoft Graph documentation](https://learn.microsoft.com/graph/)
* [Office Add-ins documentation](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins)

## Copyright

Copyright (c) 2019 and 2021 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

**Note**: The Index.html file contains an image URL that tracks diagnostic data for this sample add-in. Please remove the image tag if you reuse this sample in your own code project.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/auth/Office-Add-in-Microsoft-Graph-React" />
