---
page_type: sample
urlFragment: office-add-in-sso-aspnet
products:
  - m365
  - office
  - office-excel
  - office-powerpoint
  - office-word
languages:
  - javascript
  - aspx
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
# Office Add-in that supports Single Sign-on to Office, the add-in, and Microsoft Graph

This sample implements an Office Add-in that uses the `getAccessToken` API in Office.js to give the add-in access to Microsoft Graph data. This sample is built on ASP.NET and Microsoft Authentication Library (MSAL) .NET.

There are two versions of the sample in this repo, one of which has its own README file:

- **Begin** folder. The starting point for the SSO walkthrough at [Create an ASP.NET Office Add-in that uses single sign-on](https://learn.microsoft.com/office/dev/add-ins/develop/create-sso-office-add-ins-aspnet). Please follow the instructions in the article.
- **Complete** folder. The completed sample. To use this version, follow the instructions in the article [Create an ASP.NET Office Add-in that uses single sign-on](https://learn.microsoft.com/office/dev/add-ins/develop/create-sso-office-add-ins-aspnet), but substitute "Complete" for "Begin" in those instructions and skip the sections **Code the client-side** and **Code the server-side**.

## Features

Integrating data from online service providers increases the value and adoption of your add-ins. This code sample shows you how to connect your add-in to Microsoft Graph. Use this code sample to:

- See how to use the Single Sign-on (SSO) API.
- Access Microsoft Graph on behalf of an Office Add-in.
- Build an add-in using ASP.NET Core, MSAL 4.x.x for .NET, and Office.js.
- Use the MSAL.NET Library to implement the OAuth 2.0 authorization framework in an add-in.
- Use the OneDrive REST APIs from Microsoft Graph.

## Applies to

-  Any platform and Office host combination that supports the [IdentityAPI 1.3 requirement set](https://learn.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).

## Prerequisites

To run this code sample, the following are required:

- Visual Studio 2019 or later.
- [Office Developer Tools](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx).
- Microsoft 365 - You can get a [free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) that provides a renewable 90-day Microsoft 365 E5 developer subscription.
- At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription.
- A Microsoft Azure Tenant. This add-in requires Azure Active Directory (AD). Azure AD provides identity services that applications use for authentication and authorization. A trial subscription can be acquired here: [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Solution

Solution | Author(s)
---------|----------
Office Add-in Microsoft Graph ASP.NET | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.0 | May 10, 2017| Initial release
1.0 | September 15, 2017 | Added support for 2FA.
1.0 | December 8, 2017 | Added extensive error handling.
1.0 | January 7, 2019 | Added information about web application security practices.
2.0 | November 5, 2019 | Added Display Dialog API fall back and use new version of SSO API.
2.1 | August 11, 2020 | Removed preview note because the APIs have released.
2.2 | June 15, 2021 | Updated NuGet packages and adjust code for breaking changes.
3.0 | November 7, 2022 | Updated to use ASP.NET Core. Removed fallback dialog approach.

## Security note

The sample sends a hardcoded query parameter on the URL for the Microsoft Graph REST API. If you modify this code in a production add-in and any part of query parameter comes from user input, be sure that it is sanitized so that it cannot be used in a Response header injection attack.

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

Copyright (c) 2019 - 2021 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

**Note**: The Index.cshtml file contains an image URL that tracks diagnostic data for this sample add-in. Please remove the image tag if you reuse this sample in your own code project.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/auth/Office-Add-in-ASPNET-SSO" />
