---
page_type: sample
urlFragment: office-excel-add-in-open-in-teams
products:
- office-excel
- office-teams
- m365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 02/11/2021 1:25:00 PM
description: "Learn how to open data from your web site in a spreadsheet in Teams for team collaboration."
---

# Open data from your web site in a spreadsheet in Microsoft Teams

This sample accomplishes the following tasks.

- Creates a new Excel spreadsheet in Microsoft Teams containing data you define.
- Embeds your add-in into the Excel spreadsheet.
- Creates a message in Microsoft Teams with a link to the new spreadsheet.

## Applies to

- Microsoft Teams
- Microsoft Excel

## Prerequisites

- Microsoft 365

## Register the add-in with Azure AD v2.0 endpoint

1. Navigate to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page to register your app.

1. Sign in with the ***admin*** credentials to your Microsoft 365 tenancy. For example, MyName@contoso.onmicrosoft.com.

1. Select **New registration**. On the **Register an application** page, set the values as follows:

    * Set **Name** to `OpenInTeamsSample`.
    * Set **Supported account types** to **Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)**.
    * In the **Redirect URI** section, ensure that **Web** is selected in the dropdown and then set the URI to `https://localhost:44326/`.
    
    **Note:** The port number used for the redirect URI (`44326`) must match the port your web server is running on. When you open the Visual Studio solution in later steps, you can find the web server's port number by selecting the **ContosoWebApp** project in **Solution Explorer**, then looking at the **SSL URL** setting in the properties window.

1. Choose **Register**.

1. On the **OpenInTeamsSample** page, copy and save the **Application (client) ID**. You'll use it in later procedures.

//1. Under **Manage**, select **Authentication**. Under **Implicit grant**, check the **Access tokens** checkbox, then select **Save**.

1. Under **Manage**, select **Certificates & secrets**. Select the **New client secret** button. Enter a value for **Description**, then select an appropriate option for **Expires**, and choose **Add**.

1. Copy and save the client secret value. You'll use it in later procedures.

1. Under **Manage**, select **API permissions** and then select **Add a permission**. On the panel that opens, choose **Microsoft Graph** and then choose **Delegated permissions**.

1. Use the **Select permissions** search box to search the following permissions:

    * Channel.ReadBasic.All
    * ChannelMessage.Send
    * Chat.ReadWrite
    * Files.ReadWrite.All
    * openid
    * profile
    * Team.ReadBasic.All
    * User.Read

    **Note:** The `User.Read` permission may already be listed by default. It's a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission if your add-in does not actually need it.

1. Select the check box for each permission as it appears. After selecting the permissions, select the **Add permissions** button at the bottom of the panel.

1. On the same page, choose the **Grant admin consent for [tenant name]** button, and then select **Yes** for the confirmation that appears.

    **Note:** After choosing **Grant admin consent for [tenant name]**, you may see a banner message asking you to try again in a few minutes so that the consent prompt can be constructed. If so, you can start work on the next section, ***but don't forget to come back to the portal and press this button***!

## Configure the Sample

Before you run the sample, you'll need to do a few things to make it work properly.

1. In Visual Studio, open the **excel-open-in-teams.sln** solution file for this sample.
1. In the **Solution Explorer**, open **ContosoWebWeb > Web.config**.
1. Replace the `[Application (client) ID]` value in both places where it appears with the application ID you generated as part of the app registration process.
1. Replace the `[Application secret]` value with the client secret you generated as part of the app registration process.

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

## Run the sample

1. Press **F5** to build and debug the project. You may be prompted to trust the developer certificate.
    The Contoso Web application will open in a browser.
1. Choose the **Sign in with Microsoft** button on the ribbon.
1. You should be prompted for a user account and password. Sign in with a user name and password from your Microsoft 365 account.
    You'll see a new option on the ribbon named **Product data**.
1. Now that you are signed in you can choose **Product data**.
    The Product sales page is displayed.
1. Choose **Open in Microsoft Teams** to start the process of opening the data in Microsoft Teams
1. A page will appear titled **Select the Team**. Choose a team from the dropdown list and then choose **Submit**. This will select the team where you want to open the data.
1. A page will appear titled **Select the channel**. Choose a channel from the dropdown list and then choose **Submit**. This will select the channel where you want to open the data.
1. The page will now redirect to Microsoft Teams. Choose if you want to open Microsoft teams in the browser, or in the app.
1. When Microsoft Teams opens, you will see a chat message in the channel you chose containing a spreadsheet named productdata.xlsx. Choose the spreadsheet and open it. You will see the product data, and also the script lab add-in is embedded in the spreadsheet.

## Key parts of this sample

### Authentication

This sample reuses code from a Microsoft Azure OpenID Connect sample to handle authentication and authorization. For more information see [Use OpenID Connect to sign in users to Microsoft identity platform and execute Microsoft Graph operations using incremental consent](https://github.com/Azure-Samples/ms-identity-aspnet-webapp-openidconnect).

### Constructing the spreadsheet

This sample uses the [Open XML SDK](https://docs.microsoft.com/office/open-xml/open-xml-sdk) to construct the spreadsheet in memory before uploading it to OneDrive. The code that constructs the spreadsheet is in **Helpers\SpreadsheetBuilder.cs**.

* The `InsertHeader` method inserts the header for the product data table.
* The `InsertData` method inserts the data values for the product data table.
* The `EmbedAddin` method embeds the script lab add-in.
* Modify the `GenerateWebExtensionPart1Content` method to embed your add-in instead of the script lab add-in. Note that there is a *CUSTOM MODIFICATION BEGIN/END* section where you can specify and custom properties that your add-in needs to load when it starts.

### Interacting with Microsoft Teams through the Microsoft Graph API

This sample uses the Microsoft Graph API to upload the spreadsheet to the OneDrive for Microsoft Teams, and also to create the message that links to the spreadsheet. The **ProductsController.cs** file contains the code that constructs the URL calls for Microsoft Graph. The **Helpers\GraphAPIHelper.cs** file contains code that gets the access token, makes the Microsoft Graph call, and returns the result.

### Sequence of events

1. The user chooses to open in Teams. The `TeamsList` action in **ProductsController.cs** is called.
    - `TeamsList` constructs a URL to query Microsoft Graph for all teams the user belongs to.
    - `TeamsList` calls `GraphAPIHelper.CallGraphAPIGet` to run the URL call and get back the JSON list of teams.
    - `TeamsList` returns the `TeamsList` view containing a dropdown list of teams.
2. The user chooses which Team they want to open in. The `ChannelsListForTeam` action in  **ProductsController.cs** is called.
    - `ChannelsListForTeam` constructs a URL to query Microsoft Graph for all channels for the selected team.
    - `ChannelsListForTeam` calls `GraphAPIHelper.CallGraphAPIGet` to run the URL call and get back the JSON list of channels.
    - `ChannelsListForTeam` returns the `ChannelsListForTeam` view containing a dropdown list of channels.
3. The user chooses which channel they want to open in. The `UploadSpreadsheet` action in  **ProductsController.cs** is called.
    - `UploadSpreadsheet` calls `CreateSpreadsheet` to construct an in memory spreadsheet.
    - `UploadSpreadsheet` calls `UploadSpreadsheetToOneDrive` to upload the spreadsheet to the select team and channel.
    - `UploadSpreadsheetToOneDrive` constructs a URL to query Microsoft Graph for the name of the folder on OneDrive for the selected team and channel.
    - `UploadSpreadsheetToOneDrive` calls `GraphAPIHelper.CallGraphAPIGet` to run the URL call and get back the OneDrive folder name.
    - `UploadSpreadsheetToOneDrive` constructs a URL to create a new message.
    - `UploadSpreadsheetToOneDrive` calls `GraphAPIHelper.CallGraphAPIWithBody` to upload (HTTP Put) the spreadsheet to the OneDrive folder.
4. `UploadSpreadsheet` creates the message by calling `CreateChannelMessage`.
    - `CreateChannelMessage` constructs a URL to create a new message through Microsoft Graph.
    - `CreateChannelMessage` extracts the redirect URI from the `file.eTag` we received from uploading the spreadsheet file to OneDrive.
    - `CreateChannelMessage` constructs the text for the message and attaches a link to the spreadsheet file.
    - `CreateChannelMessage` calls `GraphAPIHelper.CallGraphAPIWithBody` to run the URL and create the message.
    - `UploadSpreadsheet` returns the `UploadToTeams` view containing the redirect URI to the chat message.
5. The `UploadToTeams` view loads and redirects to the chat message that was created.

## Questions and comments

We'd love to get your feedback about this sample. Please send your feedback to us in the Issues section of this repository. Questions about developing Office Add-ins should be posted to [Microsoft Q&A](https://docs.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.


## Solution

Solution | Authors
---------|----------
Open data from your web site in a spreadsheet in Microsoft Teams | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.0  | August 13, 2021 | Initial release

## Copyright

Copyright (c) 2021 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/samples/excel-open-in-teams" />