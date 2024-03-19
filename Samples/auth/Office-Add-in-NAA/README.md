---
page_type: sample
urlFragment: office-add-in-naa
products:
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
description: "This sample shows how to use Office SSO through nested app authentication."
---

# Office Add-in with nested app authentication

## Summary

This sample shows how to use nested app authentication (NAA) in an Office Add-in to access Microsoft Graph APIs for the signed in user. Nested app authentication works with Office SSO so the sample does not require a separate sign-in process. The sample displays the user's information such as name and email. It also displays the names of files in the user's Microsoft OneDrive account.

## Features

- Use NAA to get an access token to call Microsoft Graph APIs.
- Use NAA to get information about the user signed in to Office.

## Applies to

- Word, Excel, and Powerpoint on Windows, Mac, and in a browser.

## Prerequisites

- Office connected to a Microsoft 365 subscription (including Office on the web).
- [Node.js](https://nodejs.org/) version 16 or greater.
- [npm](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm) version 8 or greater.

## Build and run the solution

### Create an application registration

1. Navigate to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page to register your app.
1. Sign in with the ***admin*** credentials to your Microsoft 365 tenancy. For example, MyName@contoso.onmicrosoft.com.
1. Select **New registration**. On the **Register an application** page, set the values as follows.

    - Set **Name** to `OfficeAddinNAA`.
    - Set **Supported account types** to **Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)**.
    - In the **Redirect URI** section, ensure that **Single-page application (SPA)** is selected in the drop down and then set the URI to `brk-multihub://localhost:3000`.
    - Select **Register**.

1. On the **OfficeAddinNAA** page, copy and save the value for the **Application (client) ID**. You'll use it in the next section.

For more information on how to register your application, see [Register an application with the Microsoft Identity Platform](https://learn.microsoft.com/graph/auth-register-app-v2).

### Configure the sample

1. Open the **src/taskpane/authConfig.ts** file.
1. Replace the placeholder "Enter_the_Application_Id_Here" with the Application ID that you copied.
1. Replace the placeholder "Enter_the_Login_Hint_Here" with your Microsoft 365 user name. For example: *admin@contoso.onmicrosoft.com*.
1. Save the file.

## Key parts of this sample

The **src/taskpane/authConfig.ts** file contains the MSAL code for configuring and using NAA. It contains a class named AccountManager which manages getting user account and token information.

- The **initialize** function is called from Office.onReady to configure and intitialize MSAL to use NAA.
- The **ssoGetToken** function gets an access token for the signed in user to call Microsoft Graph APIs.
- The **ssoGetUserIdentity** function gets the account information of the signed in user. This can be used to get user details such as name and email.

The **src/taskpane/document.ts** file contains code to write a list of file names, retrieved from Microsoft Graph, into the document. This works for Word, Excel, and PowerPoint documents.

The **src/taskpane/taskpane.ts** file contains code that runs when the user chooses buttons in the task pane. They use the AccountManager class to get tokens or user information depending on which button is chosen.

The **src/taskpane/msgraph-helper.ts** file contains code to construct and make a REST call to the Microsoft Graph API.

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2024 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/office-wxp-naa" />