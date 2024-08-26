---
page_type: sample
urlFragment: outlook-event-sso-naa
products:
  - office
  - office-outlook
languages:
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: "08/20/2024 10:00:00 AM"
description: "This sample shows how to implement SSO in an event in an Outlook add-in by using nested app authentication."
---

# Use SSO in events in an Outlook add-in using nested app authentication (preview)

## Summary

This sample shows how to implement single sign-on (SSO) in an event in an Outlook add-in. It uses the Microsoft Authentication Library for JavaScript (MSAL.js) and nested app authentication (NAA) to access Microsoft Graph APIs for the signed-in user. The sample displays the signed-in user's name as a signature in the body of a new email or calendar item.

> [!IMPORTANT]
> Nested app authentication is currently in preview. To try this feature, you need to join the [Microsoft 365 Insider Program](https://insider.microsoft365.com/join) and choose **Beta Channel**. Don't use NAA in production add-ins. We invite you to try out NAA in test or development environments and welcome feedback on your experience through GitHub (see https://github.com/OfficeDev/office-js/issues).

## Features

- Use MSAL.js NAA to get an access token for the signed in user to call Microsoft Graph APIs.
- Get an access token through NAA in the `OnNewMessageCompose` and `OnNewAppointmentOrganizer` events.
- Add a signature to an email or calendar invite with the signed-in user's name.

## Applies to

- Outlook (Beta Channel for classic Outlook only, new Outlook coming soon)

For more information on supported platforms, see [NAA supported accounts and hosts](https://learn.microsoft.com/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in#naa-supported-accounts-and-hosts).

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

    - Set **Name** to `Outlook-Event-SSO-NAA`.
    - Set **Supported account types** to **Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)**.
    - In the **Redirect URI** section, ensure that **Single-page application (SPA)** is selected in the drop down and then set the URI to `brk-multihub://localhost:3000`.
    - Select **Register**.

1. On the **Outlook-Add-in-SSO-NAA** page, copy and save the value for the **Application (client) ID**. You'll use it in the next section.
1. Select the link to modify redirect URIs which should appear as **0 web, 1 spa, 0 public client**.

      ![The redirect URIs link.](./assets/ui-add-redirect-link.png)

1. In the **Single-page application Redirect URIs** section, select **Add URI**.

      ![The Add URI link.](./assets/ui-add-redirects-link.png)

1. Enter the new URI value `https://localhost:3000/taskpane.html` and select **Save**.

      ![The completed redirects in the application registration.](./assets/ui-completed-redirects.png)

For more information on how to register your application, see [Register an application with the Microsoft Identity Platform](https://learn.microsoft.com/graph/auth-register-app-v2).

### Configure the sample

1. Clone or download this repository.
1. From the command line, or a terminal window, go to the root folder of this sample at `/samples/auth/Outlook-Event-SSO-NAA`.
1. Open the `src/launchevent/launchevent.js` file.
1. Replace the placeholder "Enter_the_Application_Id_Here" with the Application ID that you copied.
1. Open the `src/taskpane/taskpane.js` file.
1. Replace the placeholder "Enter_the_Application_Id_Here" with the Application ID that you copied.
1. Save the file.

## Run the sample

First ensure you have signed in and consented to the add-in's scopes. Once you approve consent, you no longer need to do those steps.

1. Run the following commands.

    `npm install`
    `npm run build:dev`
    `npm run start`

    This will start the web server and sideload the add-in to Outlook.

> [!IMPORTANT]
> The dev build uses the `https://localhost:3000/public/launchevent.js` path. If you change to the production build, you also need to change the URL to `https://localhost:3000/launchevent.js` in the `manifest.xml` file.

1. Start Outlook (classic) and sign in.
1. Open an existing email item.
1. Choose "Show Task Pane" from the ribbon. This will open the task pane of the add-in.
1. Select the **Sign in** button to sign in. You may be prompted to consent to the scopes of the add-in. The task pane will indicate it signed in by displaying your user name.

Now you can use the event-based code.

1. Create a new email. The add-in will automatically add a signature with your signed in name.

> [!NOTE]
> You can also consent using the following URL. This avoids the steps of signing in by using the task pane.

https://login.microsoftonline.com/{tenant}/v2.0/adminconsent
        ?client_id={appRegistrationID}
        &scope=https://graph.microsoft.com/User.Read https://graph.microsoft.com/openid https://graph.microsoft.com/profile
        &redirect_uri=brk-multihub://localhost:3000

- {tenant} is the ID of the tenant that is granting admin consent.
- {appRegistrationID} is the ID of the app registration you created for the add-in.

For more information, see [Admin consent on the Microsoft identity platform](https://learn.microsoft.com/entra/identity-platform/v2-admin-consent)

## Debugging steps

To debug this sample, follow the instructions in [Debug your event-based or spam-reporting Outlook add-in](https://learn.microsoft.com/office/dev/add-ins/outlook/debug-autolaunch). All `console.log` statements from the event code appear in the [runtime log](https://learn.microsoft.com/office/dev/add-ins/testing/runtime-logging).

## Key parts of this sample

### Events using MSAL.js with NAA

The `src/launchevent/launchevent.js` file contains the code for the Outlook events. It initializes the public client application (PCA) for MSAL and calls `acquireTokenSilent` to get the access token. It does not call `acquireTokenPopup` because event code cannot interact with UI. If `acquiretokenSilent` fails, it will log the error. `console.write` statements will write messages to the [runtime log](https://learn.microsoft.com/office/dev/add-ins/testing/runtime-logging).

### Webpack configuration and hot reload

The `webpack.config.js` file is updated from what yo office generates. Hot reload code from webpack is not supported in the JS runtime for Outlook events. The webpack config modifiations ensure you can import the MSAL.js library without the hot reload code.

### Well-known URI

The `src/well-known/microsoft-officeaddins-allowed.json` file lists `https://localhost:3000/public/launchevent.js` as an allowed file to access SSO. For more information, see [Use single sign-on (SSO) or cross-origin resource sharing (CORS) in your event-based or spam-reporting Outlook add-in](https://learn.microsoft.com/office/dev/add-ins/outlook/use-sso-in-event-based-activation).

## Security reporting

If you find a security issue with our libraries or services, report the issue to [secure@microsoft.com](mailto:secure@microsoft.com) with as much detail as you can provide. Your submission may be eligible for a bounty through the [Microsoft Bounty](https://aka.ms/bugbounty) program. Don't post security issues to [GitHub Issues](https://github.com/AzureAD/microsoft-authentication-library-for-android/issues) or any other public site. We'll contact you shortly after receiving your issue report. We encourage you to get new security incident notifications by visiting [Microsoft technical security notifications](https://technet.microsoft.com/security/dd252948) to subscribe to Security Advisory Alerts.

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2024 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/outlook-event-sso-naa" />