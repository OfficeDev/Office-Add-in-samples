# Complete sample for Office Add-in using SSO with Node.JS

This sample shows how to use single sign-on (SSO) in your Office Add-in. The `getAccessToken` API in Office.js enables your add-in to access user data and Microsoft Graph services on behalf of the user who is currently signed in to Office. There is no need for the user to sign in a second time to your add-in.

This sample is also used as a template for yo office. If you installed this sample by using yo office, there are additional steps to get the code running. Follow the instructions in [Create a Node.js Office Add-in that uses single sign-on](https://docs.microsoft.com/office/dev/add-ins/develop/create-sso-office-add-ins-nodejs), but with the following changes:
    - Substitute "Complete" for "Begin"
    - Skip the sections **Code the client-side** and **Code the server-side**
