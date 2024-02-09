---
page_type: sample
urlFragment: office-add-in-sso-nodejs
products:
- office-excel
- office-powerpoint
- office-word
- m365
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
# Office Add-in that supports Single Sign-on to Office, the Add-in, and Microsoft Graph

The `getAccessToken` API in Office.js enables users who are signed into Office to get access to an AAD-protected add-in and to Microsoft Graph without needing to sign-in again. 

There are two versions of the sample in this repo:

- In the **Begin** folder is the starting point for the SSO walkthrough at [Create a Node.js Office Add-in that uses single sign-on](https://learn.microsoft.com/office/dev/add-ins/develop/create-sso-office-add-ins-nodejs). Please follow the instructions in the article.
- In the **Complete** folder is the completed sample. To use this version, follow the instructions in [Create a Node.js Office Add-in that uses single sign-on](https://learn.microsoft.com/office/dev/add-ins/develop/create-sso-office-add-ins-nodejs), but with the following changes:
  - Substitute "Complete" for "Begin"
  - Skip the sections **Code the client-side** and **Code the server-side**

These samples are built on Node.JS, Express, and Microsoft Authentication Library for JavaScript (msal.js).

## Features

Integrating data from online service providers increases the value and adoption of your add-ins. This code sample shows you how to connect your add-in to Microsoft Graph. Use this code sample to:

- Build an add-in using Node.js, Express, msal.js, and Office.js
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

- A code editor. We recommend Visual Studio Code which was used to create the sample.
- A Microsoft 365 account. You can get one if you qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram); for details, see the [FAQ](/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-). Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g).
- At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription.
- A Microsoft Azure Tenant. This add-in requires Azure Active Directory (AD). Azure AD provides identity services that applications use for authentication and authorization. A trial subscription can be acquired here: [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Solution

Solution | Author(s)
---------|----------
Office Add-in Microsoft Graph Node.js | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.0 | May 10, 2017| Initial release
1.0 | September 15, 2017 | Added support for 2FA.
1.0 | December 8, 2017 | Added extensive error handling.
1.0 | January 7, 2019 | Added information about web application security practices.
2.0 | October 26, 2019 | Changed to use new API and added Display Dialog API fallback.
2.1 | August 11, 2020 | Removed preview note because the API has released.
2.2 | July 7, 2022 | Fixed middle-tier token handling and MSAL fallback approach to be consistent with Microsoft identity platform guidance.
2.3 | February 16, 2023 | Refactored code to simplify.

## Security note

These samples send a hardcoded query parameter on the URL for the Microsoft Graph REST API. If you modify this code in a production add-in and any part of query parameter comes from user input, be sure that it is sanitized so that it cannot be used in a Response header injection attack.

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Join the Microsoft 365 Developer Program

Get a free sandbox, tools, and other resources you need to build solutions for the Microsoft 365 platform.

- [Free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) Get a free, renewable 90-day Microsoft 365 E5 developer subscription.
- [Sample data packs](https://developer.microsoft.com/microsoft-365/dev-program#Sample) Automatically configure your sandbox by installing user data and content to help you build your solutions.
- [Access to experts](https://developer.microsoft.com/microsoft-365/dev-program#Experts) Access community events to learn from Microsoft 365 experts.
- [Personalized recommendations](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations) Find developer resources quickly from your personalized dashboard.

## Additional resources

- [Microsoft Graph documentation](https://learn.microsoft.com/graph/)
- [Office Add-ins documentation](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins)

## Copyright

Copyright (c) 2021 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

**Note**: The index.pug file contains an image URL that tracks diagnostic data for this sample add-in. Please remove the image tag if you reuse this sample in your own code project.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/auth/Office-Add-in-NodeJS-SSO" />
