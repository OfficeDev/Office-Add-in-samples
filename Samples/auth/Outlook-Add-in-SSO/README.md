---
page_type: sample
urlFragment: outlook-add-in-sso-aspnet
products:
- office
- office-outlook
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 05/06/2021 10:00:00 AM
description: "An Outlook add-in sample that accesses Microsoft Graph using single sign-on and adds buttons to the Outlook ribbon."
---

# Single Sign-on (SSO) sample Outlook add-in

**Applies to:** Outlook on Windows | Outlook on Mac | Outlook on the web

## Summary

The sample implements an Outlook add-in that uses Office's SSO feature to give the add-in access to Microsoft Graph data. Specifically, it enables the user to save all attachments to their OneDrive. It also shows how to add custom buttons to the Outlook ribbon. The sample illustrates the following concepts:

- [Use the SSO access token](https://learn.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) to call the Microsoft Graph API without prompting the user
- If the SSO token is not available, authenticate to the user's OneDrive using the [Microsoft Authentication Library (MSAL)](https://learn.microsoft.com/azure/active-directory/develop/msal-overview).
- Use the [Microsoft Graph API](https://developer.microsoft.com/graph/docs/api-reference/v1.0/resources/onedrive) to create files in OneDrive
- Implement a WebAPI that uses the [Microsoft identity platform and OAuth 2.0 On-Behalf-Of flow](https://learn.microsoft.com/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow) to exchange the user's access token for a new access token with permissions to Microsoft Graph and OneDrive.

## Register the add-in with Microsoft identity platform

Use the following value for the placeholder for the subsequent app registration steps.

| Placeholder or section | Value                               |
|---------------------|-------------------------------------|
| <add-in-name>       | AttachmentDemo                      |
| Microsoft Graph permissions | openid, profile, offline_access, Files.ReadWrite, Mail.Read|

Follow the steps in [Register an Office Add-in that uses single sign-on (SSO) with the Microsoft identity platform](https://learn.microsoft.com/office/dev/add-ins/develop/register-sso-add-in-aad-v2)

## Add an SPA redirect to the app registration

Fallback authentication uses the MSAL library to sign in the user. This requires the add-in to acquire an access token from the client side in the task pane. To support this you will need an SPA redirect in the app registration.

1. Open the app registration you created in the Azure portal.
1. Under **Manage**, select **Authentication**.
1. Select **Add a platform**.
1. In the **Configure platforms** pane, select **Single-page application**.
1. In the **Configure single-page application** pane, enter `https://localhost:44355/dialog.html` for the **Redirect URIs** and then select **Configure**.

Now your app registration supports fallback authentication in the task pane.

## Configure the Sample

Before you run the sample, you'll need to do a few things to make it work properly.

1. In Visual Studio, open the **AttachmentDemo.sln** solution file for this sample.

1. In the **Solution Explorer**, open **AttachmentDemo > AttachmentDemoManifest > AttachmentDemo.xml**.
1. Find the `<WebApplicationInfo>` section near the bottom of the manifest. Then replace the `Enter_client_ID_here` value, in both places where it appears, with the application ID you generated as part of the app registration process.

    **Note:** Make sure that the port number in the `Resource` element matches the port used by your project. It should also match the port you used when registering the application.

1. In the **Solution Explorer**, open **AttachmentDemo-ASPNETCore > appsetings.json**.
1. Replace the `Enter_client_ID_here` placeholder value with the application ID you generated as part of the app registration process.
1. Replace the `Enter_client_secret_here` value with the client secret you generated as part of the app registration process.

## Provide user consent to the app

If you want to try the add-in using a different tenant than the one where you registered the app, you need to do this step.

You have two choices for providing consent:

- All users. Use an administrator account and consent once for all users in your Office 365 tenant
- Single user. Use any account to consent for just that user

### Provide admin consent for all users

If you have access to a tenant administrator account, this method allows you to provide consent for all users in your organization, which can be convenient if you have multiple developers that need to develop and test your add-in.

1. Browse to `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`, where `{application_ID}` is the application ID shown in your app registration.
1. Sign in with your administrator account.
1. Review the permissions and click **Accept**.

The browser will attempt to redirect back to your app, which may not be running. You might see a "this site cannot be reached" error after clicking **Accept**. This is OK, the consent was still recorded.

### Provide consent for a single user

If you don't have access to a tenant administrator account, or you just want to limit consent to a few users, this method allows you to provide consent for a single user.

1. Browse to `https://login.microsoftonline.com/common/oauth2/authorize?client_id={application_ID}&state=12345&response_type=code`, where `{application_ID}` is the application ID shown in your app registration.
1. Sign in with your account.
1. Review the permissions and click **Accept**.

The browser will attempt to redirect back to your app, which may not be running. You might see a "this site cannot be reached" error after clicking **Accept**. This is OK, the consent was still recorded.

## Run the Sample

1. Select the **AttachmentDemo** project in **Solution Explorer**, then choose the **Start Action** value you want (under **Add-in** in the properties window). Choose any installed browser to launch Outlook on the web, or you can choose **Office Desktop Client** to launch Outlook on Windows. If you choose **Office Desktop Client**, be sure to configure Outlook to connect to the Office or Outlook.com user you want to install the add-in for.

1. Press **F5** to build and debug the project. You may be prompted to trust the developer certificate.

1. You should be prompted for a user account and password. Provide a user in your Office tenant, or an Outlook.com account. The add-in will be installed for that user, and either Outlook on the web or Outlook on Windows will open.

1. Select any message, **that has one or more attachments**.

1. Open the task pane:
    
    * If you're in Outlook on the web: Select the **...** (**More actions**) drop down menu, and then choose **AttachmentDemoWeb**.<br>
![Screen shot of the elipses button in Outlook on the web](buttons-outlook-web.png)
    * If you're in Outlook on Windows or Mac: On the **Home** tab, select **Choose attachments**. Note that if the Outlook app window is too small, that **Choose attachments** will instead be located on the **Home** tab's **...** (**More commands**) button.<br>
![Screen shot of the Choose attachments button on Home tab in Outlook on the web](buttons-outlook-desktop.png)
    
1. In the **AttachmentDemoWeb** task pane that opens, select the attachments you want to save.

1. Choose **Save to OneDrive**.

1. You should see a success message in the task pane.<br>
![Screen shot of the task pane displaying attachments successfully saved](successful-save.png)

1. Open OneDrive and you should see the attachments saved in a new folder named **Outlook Attachments**.<br>
![Screen shot of the Outlook Attachments folder in OneDrive](onedrive-attachments-folder.png)

## Testing the fallback dialog

It's recommended to test all paths when working with SSO. In some scenarios, you'll have to use the fallback dialog by modifying the code as follows:

1. In Visual Studio, open the **MessageRead.js** file.
1. Change the authSSO value to false.
    
    ```javascript
    let authSSO = false;
    ```

There is also a `MockError` method available in the **Saveattachments.cs** controller to throw various MSAL or Microsoft Graph errors for testing.    

## Troubleshoot manifest issues

Visual Studio may show a warning or error about the `WebApplicationInfo` element being invalid. The error may not show up until you try to build the solution. As of this writing, Visual Studio has not updated the schema files to include the `WebApplicationInfo` element. To work around this problem, use the updated schema file in this repository: [MailAppVersionOverridesV1_1.xsd](manifest-schema-fix/MailAppVersionOverridesV1_1.xsd).

1. On your development machine, locate the existing MailAppVersionOverridesV1_1.xsd. This should be located in your Visual Studio installation directory under `./Xml/Schemas/{lcid}`. For example, on a typical installation of VS 2017 32-bit on an English (US) system, the full path would be `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.
1. Rename the existing file to `MailAppVersionOverridesV1_1.old`.
1. Move the version of the file from this repository into the folder.

## Copyright

Copyright (c) 2021 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/auth/Outlook-Add-in-SSO" />
