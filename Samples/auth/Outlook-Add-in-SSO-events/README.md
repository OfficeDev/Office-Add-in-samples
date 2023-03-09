---
title: "Use SSO with event-based activation in an Outlook add-in"
page_type: sample
urlFragment: outlook-add-in-sso-event-based-activation
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
  createdDate: 02/21/2023 10:00:00 AM
description: "Use SSO with event-based activation in an Outlook add-in."
---

# Use SSO with event-based activation in an Outlook add-in

**Applies to**: Outlook on Windows | Outlook on the web | [new Outlook on Mac](https://support.microsoft.com/office/6283be54-e74d-434e-babb-b70cefc77439)

## Summary

The sample shows how to use SSO to access a user's Microsoft Graph data from an event fired in an Outlook add-in. The sample illustrates the following concepts:

- [Get a user access token using SSO](https://learn.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) when the [OnNewMessageCompose event](https://learn.microsoft.com/office/dev/add-ins/outlook/autolaunch) fires.
- Implement a server REST API that uses the [Microsoft identity platform and OAuth 2.0 On-Behalf-Of flow](https://learn.microsoft.com/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow) to exchange the user's access token for a new access token with permissions to the users profile on Microsoft Graph.
- Use the [Microsoft Graph API](https://developer.microsoft.com/graph/docs/api-reference/v1.0/resources/onedrive) to get user profile data, such as display name and job title.
- Construct a signature in the mail item containing the user profile data.

## Register the add-in with the Microsoft identity platform

Use the following values for the subsequent app registration steps.

| Placeholder or section | Value          |
|------------------------|----------------|
| `<add-in-name>`          | `outlook-event-sso-sample` |
| `<fully-qualified-domain-name>` | `localhost:3000` |
| Microsoft Graph permissions | profile, openid, User.Read |

Follow the steps in [Register an Office Add-in that uses single sign-on (SSO) with the Microsoft identity platform](https://learn.microsoft.com/office/dev/add-ins/develop/register-sso-add-in-aad-v2).

> Note: The instructions tell you to create a redirect URI for a single-page application. This step isn't necessary for this sample because it doesn't use a fallback authentication approach if SSO fails.

## Configure the sample

Before you run the sample, you'll need to do a few things to make it work properly.

1. Clone or download this repo.
1. In Visual Studio Code (or editor of your choice), open the root folder for this sample.

### Update manifest.xml

1. Open the **manifest.xml** file.
1. Find the `<WebApplicationInfo>` section near the bottom of the manifest. Then, replace the `Enter_client_ID_here` value, in both places where it appears, with the application ID you generated as part of the app registration process.

> Note: Make sure that the port number in the `Resource` element matches the port used by your project. It should also match the port you used when registering the application.

1. Save your changes.

### Update .ENV

1. Open the **.ENV** file.
1. Replace the `Enter_client_ID_here` placeholder value with the application ID you generated as part of the app registration process.
1. Replace the `Enter_client_secret_here` placeholder value with the client secret you generated as part of the app registration process.
1. Save your changes.

## Provide user consent to the app

If you want to try the add-in using a different tenant than the one where you registered the app, you need to complete this step.

You have two choices for providing consent:

- All users. Use an administrator account and consent once for all users in your Office 365 tenant.
- Single user. Use any account to consent for just that user.

### Provide admin consent for all users

If you have access to a tenant administrator account, this method allows you to provide consent for all users in your organization. This is convenient if you have multiple developers that need to develop and test your add-in.

1. Browse to `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`, where `{application_ID}` is the application ID shown in your app registration.
1. Sign in with your administrator account.
1. Review the permissions and click **Accept**.

The browser will attempt to redirect back to your app, which may not be running. You might see a "this site cannot be reached" error after clicking **Accept**. This is OK, the consent was still recorded.

### Provide consent for a single user

If you don't have access to a tenant administrator account, or you want to limit consent to a few users, this method allows you to provide consent for a single user.

1. Browse to `https://login.microsoftonline.com/common/oauth2/authorize?client_id={application_ID}&state=12345&response_type=code`, where `{application_ID}` is the application ID shown in your app registration.
1. Sign in with your account.
1. Review the permissions and click **Accept**.

The browser will attempt to redirect back to your app, which may not be running. You might see a "this site cannot be reached" error after clicking **Accept**. This is OK, the consent was still recorded.

## Run the sample

1. Open a terminal window and run the command `npm install` to install all package dependencies.
1. Run the command `npm start` to start the web server.
1. To test the add-in in Outlook, you need to sideload it. Follow the instructions at [Sideload an Office Add-in for testing](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing) to sideload the sample.
1. Compose a new email. The email will display a notification that it will append a signature.
1. Send the email to yourself. Check when it arrives that the signature is appended.

## SSO and fallback

It's recommended to always have a fallback authentication approach if SSO fails for any reason. However, fallback authentication requires a popup dialog for the user to sign in. It's not possible to open a dialog from an event in Outlook, so this sample doesn't use fallback authentication. If an error occurs, the sample displays the error as a notification on the message, and the signature is not appended.

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2023 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/samples/outlook-add-in-sso-event-based-activation" />
