---
page_type: sample
urlFragment: office-add-in-sso-nodejs
products:
  - m365
  - office
  - office-excel
  - office-powerpoint
  - office-word
languages:
  - javascript
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
# Office Add-in that supports legacy Office Single Sign-on, the Add-in, and Microsoft Graph

**Note:** This sample uses legacy Office single sign-on (SSO). For a modern authentication experience with support across a wider range of platforms, use the Microsoft Authentication Library (MSAL) with nested app authentication (NAA). For more information, see [Enable single sign-on in an Office Add-in with nested app authentication](https://learn.microsoft.com/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in).

The `getAccessToken` API in Office.js enables users who are signed into Office to get access to an AAD-protected add-in and to Microsoft Graph without needing to sign-in again.

This sample is built on Node.JS, Express, and [Microsoft Authentication Library for Node (msal-node)](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-node).

## Features

Integrating data from online service providers increases the value and adoption of your add-ins. This code sample shows you how to connect your add-in to Microsoft Graph. Use this code sample to:

- Build an add-in using Node.js, Express, msal-node, and Office.js
- Connect to Microsoft Graph from an Office Add-in
- Use the OneDrive REST APIs from Microsoft Graph
- Use the Express routes and middleware to implement the OAuth 2.0 authorization framework in an add-in
- See how to use the Single Sign-on (SSO) API
- See how an add-in can fall back to an interactive sign-in in scenarios where SSO is not available
- Use the msal.js library to implement a fallback authentication/authorization system that is invoked when Office SSO is not available
- Show a dialog using the Office UI namespace when Office SSO is not available
- Use add-in commands in an add-in

## Applies to

- Excel on Windows (subscription)
- PowerPoint on Windows (subscription)
- Word on Windows (subscription)

## Prerequisites

To run this code sample, the following are required:

- [Node.js](https://nodejs.org/) (the latest [LTS](https://nodejs.org/en/about/previous-releases) version)
- [Git Bash](https://git-scm.com/downloads) (or another git client)
- A code editor - we recommend Visual Studio Code
- At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription
- A build of Microsoft 365 that supports the [IdentityAPI 1.3 requirement set](/javascript/api/requirement-sets/common/identity-api-requirement-sets). You might qualify for a Microsoft 365 E5 developer subscription, which includes a developer sandbox, through the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram); for details, see the [FAQ](/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-). The [developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) includes a Microsoft Azure subscription that you can use for app registrations in later steps in this article. If you prefer, you can use a separate Microsoft Azure subscription for app registrations. Get a trial subscription at [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Configure the app registration

Follow the instructions at [Register an Office Add-in that uses single sign-on (SSO) with the Microsoft identity platform](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/register-sso-add-in-aad-v2). Use the following values for placeholders in the app registration steps from the previous article.

| Placeholder                     | Value                                 |
|---------------------------------|---------------------------------------|
| `<add-in-name>`                 | **Office-Add-in-NodeJS-SSO**          |
| `<fully-qualified-domain-name>` | `localhost:3000`                      |
| Microsoft Graph permissions     | profile, openid, Files.Read           |

## Configure the add-in to use the app registration

1. Open a command prompt in the root folder of this project.

1. Enter `npm install` in the console to install all of the dependencies itemized in the package.json file.

1. Run the command `npm run install-dev-certs`. Select **Yes** to the prompt to install the certificate.

1. Open the project in a code editor.

1. Open the `.ENV` file and use the values that you copied earlier from the **Office-Add-in-NodeJS-SSO** app registration. Set the values as follows:

    | Name              | Value                                                            |
    | ----------------- | ---------------------------------------------------------------- |
    | **CLIENT_ID**     | **Application (client) ID** from app registration overview page. |
    | **CLIENT_SECRET** | **Client secret** saved from **Certificates & Secrets** page.    |

    The values should **not** be in quotation marks.

1. Open the add-in manifest file "manifest\manifest_local.xml" and then scroll to the bottom of the file. Just above the `</VersionOverrides>` end tag, you'll find the following markup.

   ```xml
   <WebApplicationInfo>
     <Id>$app-id-guid$</Id>
     <Resource>api://localhost:3000/$app-id-guid$</Resource>
     <Scopes>
         <Scope>Files.Read</Scope>
         <Scope>profile</Scope>
         <Scope>openid</Scope>
     </Scopes>
   </WebApplicationInfo>
   ```

1. Replace the placeholder "$app-id-guid$" _in both places_ in the markup with the **Application ID** that you copied when you created the **Office-Add-in-NodeJS-SSO** app registration. The "$" symbols are not part of the ID, so don't include them. This is the same ID you used for the CLIENT_ID in the .ENV file.

   **Note:** The `<Resource>` value is the **Application ID URI** you set when you registered the add-in. The `<Scopes>` section is used only to generate a consent dialog box if the add-in is sold through Microsoft Marketplace.

1. Open the `\public\javascripts\fallback-msal\authConfig.js` file. Replace the placeholder "$app-id-guid$" with the application ID that you saved from the **Office-Add-in-NodeJS-SSO** app registration you created previously.

1. Save the changes to the file.

## Run the project

1. Ensure that you have some files in your OneDrive so that you can verify the results.

1. Open a command prompt in the root of this project.

1. Run the command `npm install` to install all package dependencies.

1. Run the command `npm start` to start the middle-tier server.

1. You need to sideload the add-in into an Office application (Excel, Word, or PowerPoint) to test it. The instructions depend on your platform. There are links to instructions at [Sideload an Office Add-in for Testing](https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing).

1. In the Office application, on the **Home** ribbon, select the **Show Add-in** button in the **SSO Node.js** group to open the task pane add-in.

1. Click the **Get OneDrive File Names** button. If you're logged into Office with either a Microsoft 365 Education or work account, or a Microsoft account, and SSO is working as expected the first 10 file and folder names in your OneDrive for Business are inserted into the document. (It may take as much as 15 seconds the first time.) If you're not logged in, or you're in a scenario that doesn't support SSO, or SSO isn't working for any reason, you'll be prompted to sign in. After you sign in, the file and folder names appear.

**Note:** If you were previously signed into Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so. If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned. To prevent this, be sure to _close all other Office applications_ before you press **Get OneDrive File Names**.

## Stop running the project

When you're ready to stop the middle-tier server and uninstall the add-in, follow these instructions:

1. Run the following command to stop the middle-tier server.

    ```console
    npm stop
    ```

1. To uninstall or remove the add-in, see the specific sideload article you used for details.

## Security notes

- The `/getuserfilenames` route in `getFilesroute.js` uses a literal string to compose the call for Microsoft Graph. If you change the call so that any part of the string comes from user input, sanitize the input so that it cannot be used in a Response header injection attack.

- In `app.js` the following content security policy is in place for scripts. You may want to specify additional restrictions depending on your add-in security needs.

    `"Content-Security-Policy": "script-src https://appsforoffice.microsoft.com https://ajax.aspnetcdn.com https://alcdn.msauth.net " +  process.env.SERVER_SOURCE,`

Always follow security best practices in the [Microsoft identity platform documentation](https://learn.microsoft.com/entra/identity-platform/).

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Join the Microsoft 365 Developer Program

Join the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram) to get resources and information to help you build solutions for the Microsoft 365 platform, including recommendations tailored to your areas of interest.

You might also qualify for a free developer subscription that's renewable for 90 days and comes configured with sample data; for details, see the [FAQ](https://learn.microsoft.com/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-).

## Additional resources

- [Microsoft Graph documentation](https://learn.microsoft.com/graph/)
- [Office Add-ins documentation](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins)

## Copyright

Copyright (c) 2021 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Solution

| Solution | Author(s) |
| --------- | ---------- |
| Office Add-in Microsoft Graph Node.js | Microsoft |

## Version history

| Version | Date | Comments |
| --------- | ----- | -------- |
| 1.0 | May 10, 2017 | Initial release |
| 1.0 | September 15, 2017 | Added support for 2FA. |
| 1.0 | December 8, 2017 | Added extensive error handling. |
| 1.0 | January 7, 2019 | Added information about web application security practices. |
| 2.0 | October 26, 2019 | Changed to use new API and added Display Dialog API fallback. |
| 2.1 | August 11, 2020 | Removed preview note because the API has released. |
| 2.2 | July 7, 2022 | Fixed middle-tier token handling and MSAL fallback approach to be consistent with Microsoft identity platform guidance. |
| 2.3 | February 16, 2023 | Refactored code to simplify. |
| 2.4 | January 21, 2026 | Security reviewed. Updated as legacy content as MSAL SSO via nested app authentication is preferred. |

**Note**: The index.pug file contains an image URL that tracks diagnostic data for this sample add-in. Please remove the image tag if you reuse this sample in your own code project.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/auth/Office-Add-in-NodeJS-SSO" />
