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

**Note:** This sample uses legacy Office single sign-on (SSO). For a modern authentication experience with support across a wider range of platforms, use the Microsoft Authentication Library (MSAL) with nested app authentication (NAA). For more information, see [Enable single sign-on in an Office Add-in with nested app authentication](https://learn.microsoft.com/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in).

This sample implements an Office Add-in that uses the legacy `getAccessToken` API in Office.js to give the add-in access to Microsoft Graph data. This sample is built on ASP.NET and Microsoft Authentication Library (MSAL) .NET.

## Features

Integrating data from online service providers increases the value and adoption of your add-ins. This code sample shows you how to connect your add-in to Microsoft Graph. Use this code sample to:

- See how to use the legacy Office Single Sign-on (SSO) API.
- Access Microsoft Graph on behalf of an Office Add-in.
- Build an add-in using ASP.NET Core, MSAL 4.x for .NET, and Office.js.
- Use the MSAL.NET Library to implement the OAuth 2.0 authorization framework in an add-in.
- Use the OneDrive REST APIs from Microsoft Graph.

## Applies to

- Any platform and Office host combination that supports the [IdentityAPI 1.3 requirement set](https://learn.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).

## Prerequisites

To run this code sample, the following are required:

- Visual Studio 2019 or later.
- [Office Developer Tools](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx).
- A build of Microsoft 365 that supports the [IdentityAPI 1.3 requirement set](/javascript/api/requirement-sets/common/identity-api-requirement-sets). You might qualify for a Microsoft 365 E5 developer subscription, which includes a developer sandbox, through the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram); for details, see the [FAQ](/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-). The [developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) includes a Microsoft Azure subscription that you can use for app registrations in later steps in this article. If you prefer, you can use a separate Microsoft Azure subscription for app registrations. Get a trial subscription at [Microsoft Azure](https://account.windowsazure.com/SignUp).
- At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription.

## Configure the app registration

Follow the instructions at [Register an Office Add-in that uses single sign-on (SSO) with the Microsoft identity platform](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/register-sso-add-in-aad-v2). Use the following values for placeholders in the app registration steps from the previous article.

| Placeholder                     | Value                                 |
|---------------------------------|---------------------------------------|
| `<add-in-name>`                 | **Office-Add-in-ASPNET-SSO**          |
| `<fully-qualified-domain-name>` | `localhost:44355`                     |
| Microsoft Graph permissions     | profile, openid, Files.Read           |

## Configure the solution

1. In the root of this sample folder, open the solution (.sln) file in **Visual Studio**. Right-click (or select and hold) the top node in **Solution Explorer** (the Solution node, not either of the project nodes), and then select **Set startup projects**.

1. Under **Common Properties**, select **Startup Project**, and then **Multiple startup projects**. Ensure that the **Action** for both projects is set to **Start**, and that the **Office-Add-in-ASPNETCoreWebAPI** project is listed first. Close the dialog.

1. In **Solution Explorer**, choose the **Office-Add-in-ASPNET-SSO-manifest** project and open the add-in manifest file “Office-Add-in-ASPNET-SSO.xml” and then scroll to the bottom of the file. Just above the end `</VersionOverrides>` tag, you'll find the following markup.

    ```xml
    <WebApplicationInfo>
        <Id>Enter_client_ID_here</Id>
        <Resource>api://localhost:44355/Enter_client_ID_here</Resource>
        <Scopes>
            <Scope>Files.Read</Scope>
            <Scope>profile</Scope>
            <Scope>openid</Scope>
        </Scopes>
    </WebApplicationInfo>
    ```

1. Replace the placeholder "Enter_client_ID_here" _in both places_ in the markup with the **Application ID** that you copied when you created the **Office-Add-in-ASPNET-SSO** app registration. This is the same ID you used for the application ID in the appsettings.json file.

   > [!NOTE]
   > The `<Resource>` value is the **Application ID URI** you set when you registered the add-in. The `<Scopes>` section is used only to generate a consent dialog box if the add-in is sold through Microsoft Marketplace.

1. Save and close the manifest file.

1. In **Solution Explorer**, choose the **Office-Add-in-ASPNET-SSO-web** project and open the **appsettings.json** file.

1. Replace the placeholder **Enter_client_id_here** with the **Application (client) ID** value you saved previously.

1. Replace the placeholder **Enter_client_secret_here** with the client secret value you saved previously.

    > [!NOTE]
    > You must also change the **TenantId** to support single-tenant if you configured your app registration for single-tenant. Replace the **Common** value with the **Application (client) ID** for single-tenant support.

1. Save and close the appsettings.json file.

## Run the solution

1. In Visual Studio, on the **Build** menu, select **Clean Solution**. When it finishes, open the **Build** menu again and select **Build Solution**.
1. In **Solution Explorer**, select the **Office-Add-in-ASPNET-SSO-manifest** project node.
1. In the **Properties** pane, open the **Start Document** drop down and choose one of the three options (Excel, Word, or PowerPoint).

    :::image type="content" source="../images/select-host.png" alt-text="Choose the desired Office client application: Excel, PowerPoint, or Word.":::

1. Press <kbd>F5</kbd>. Or select **Debug** > **Start Debugging**.
1. In the Office application, select the **Show Add-in** in the **SSO ASP.NET** group to open the task pane add-in.
1. Select **Get OneDrive File Names**. If you're logged into Office with either a Microsoft 365 Education or work account, or a Microsoft account, and SSO is working as expected, the first 10 file and folder names in your OneDrive for Business are displayed on the task pane. If you are not logged in, or you are in a scenario that does not support SSO, or SSO is not working for any reason, you will be prompted to sign in. After you sign in, the file and folder names appear.


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
3.1 | January 26, 2026 | Updated to reflect this is legacy code. No longer a walkthrough.

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
