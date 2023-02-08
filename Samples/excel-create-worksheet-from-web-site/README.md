---
page_type: sample
urlFragment: excel-add-in-create-spreadsheet-from-web-page
products:
- office-excel
- m365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 01/31/2023 1:25:00 PM
description: "Learn how to create a spreadsheet from your web site, populate it with data, and embed your Excel add-in."
---

# Create a spreadsheet from your web site, populate it with data, and embed your Excel add-in

This sample accomplishes the following tasks.

- Creates a new Excel spreadsheet from a web site.
- Populates the spreadsheet with data from the web site.
- Embeds the Script Lab add-in into the Excel spreadsheet.
- Opens the spreadsheet on a new browser tab.

![Sequence diagram showing an "Open in Microsoft Excel" button on your web site that creates a spreadsheet with your data which contains your add-in](./images/open-in-excel-overview.png)

This sample implements the pattern described in [Create an Excel spreadsheet from your web page, populate it with data, and embed your Office Add-in](https://learn.microsoft.com/office/dev/add-ins/excel/pnp-open-in-excel)

## Applies to

- Microsoft Excel

## Prerequisites

- [Visual Studio 2022 or later](https://aka.ms/VSDownload). Add the Office/SharePoint development workload when configuring Visual Studio.
- [Visual Studio Code](https://code.visualstudio.com/Download).
- A Microsoft 365 account. To get one, join the [Microsoft 365 Developer Program](https://aka.ms/devprogramsignup). 
- At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription.

## Set up the sample

### Step 1: Clone or download this repository

From your shell or command line:

```console
git clone https://github.com/OfficeDev/Office-Add-in-samples.git
```

or download and extract the repository *.zip* file.

> :warning: To avoid path length limitations on Windows, we recommend cloning into a directory near the root of your drive.

### Step 2: Install project dependencies

```console
    cd <WebApplication-folder>
    npm install
```

### Step 3: Register the sample application(s) in your tenant

The **WebApp** must be registered in Azure AD. To register it, you can:

- follow the steps below for manually register your apps
- or use PowerShell scripts that:
  - **automatically** creates the Azure AD applications and related objects (passwords, permissions, dependencies) for you.
  - modify the projects' configuration files.

<details>
   <summary>Expand this section if you want to use this automation:</summary>

    > :warning: If you have never used **Microsoft Graph PowerShell** before, we recommend you go through the [App Creation Scripts Guide](./WebApplication/AppCreationScripts/AppCreationScripts.md) once to ensure that your environment is prepared correctly for this step.
  
    1. On Windows, run PowerShell as **Administrator** and navigate to the root of the cloned directory
    1. In PowerShell run:

       ```PowerShell
       Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process -Force
       ```

    1. Run the script to create your Azure AD application and configure the code of the sample application accordingly.
    1. For interactive process -in PowerShell, run:

       ```PowerShell
       cd .\AppCreationScripts\
       .\Configure.ps1 -TenantId "[Optional] - your tenant id" -AzureEnvironmentName "[Optional] - Azure environment, defaults to 'Global'"
       ```

    > Other ways of running the scripts are described in [App Creation Scripts guide](./AppCreationScripts/AppCreationScripts.md). The scripts also provide a guide to automated application registration, configuration and removal which can help in your CI/CD scenarios.
</details>

#### Choose the Azure AD tenant where you want to create your applications

To manually register the apps, as a first step you'll need to:

1. Sign in to the [Azure portal](https://portal.azure.com).
1. If your account is present in more than one Azure AD tenant, select your profile at the top right corner in the menu on top of the page, and then **switch directory** to change your portal session to the desired Azure AD tenant.

#### Register the client app (contoso-addin-data-to-excel)

1. Navigate to the [Azure portal](https://portal.azure.com) and select the **Azure Active Directory** service.
1. Select the **App Registrations** blade on the left, then select **New registration**.
1. In the **Register an application page** that appears, enter your application's registration information:
    1. In the **Name** section, enter a meaningful application name that will be displayed to users of the app, for example `contoso-addin-data-to-excel`.
    1. Under **Supported account types**, select **Accounts in this organizational directory only**
    1. Select **Register** to create the application.
1. In the **Overview** blade, find and note the **Application (client) ID**. You use this value in your app's configuration file(s) later in your code.
1. In the app's registration screen, select the **Authentication** blade to the left.
1. If you don't have a platform added, select **Add a platform** and select the **Single-page application** option.
    1. In the **Redirect URI** section enter the following redirect URIs:
        1. `http://localhost:3000`
        1. `http://localhost:3000/redirect`
    1. Click **Save** to save your changes.
1. Since this app signs-in users, we will now proceed to select **delegated permissions**, which is is required by apps signing-in users.
    1. In the app's registration screen, select the **API permissions** blade in the left to open the page where we add access to the APIs that your application needs:
    1. Select the **Add a permission** button and then:
    1. Ensure that the **Microsoft APIs** tab is selected.
    1. In the *Commonly used Microsoft APIs* section, select **Microsoft Graph**
    1. In the **Delegated permissions** section, select **User.Read**, **Contacts.Read**, and **Files.ReadWrite** in the list. Use the search box if necessary.
    1. Select the **Add permissions** button at the bottom.

##### Configure Optional Claims

1. Still on the same app registration, select the **Token configuration** blade to the left.
1. Select **Add optional claim**:
    1. Select **optional claim type**, then choose **ID**.
    1. Select the optional claim **acct**.
    > Provides user's account status in tenant. If the user is a **member** of the tenant, the value is *0*. If they're a **guest**, the value is *1*.
    1. Select the optional claim **login_hint**.
    > An opaque, reliable login hint claim. This claim is the best value to use for the login_hint OAuth parameter in all flows to get SSO.See $[optional claims](https://docs.microsoft.com/azure/active-directory/develop/active-directory-optional-claims) for more details on this optional claim.
    1. Select **Add** to save your changes.

##### Configure the client app (contoso-addin-data-to-excel) to use your app registration

Open the project in your IDE (like Visual Studio or Visual Studio Code) to configure the code.

> In the steps below, "ClientID" is the same as "Application ID" or "AppId".
1. Open the `WebApplication/App/authConfig.js` file.
1. Find the key `Enter_the_Application_Id_Here` and replace the existing value with the application ID (clientId) of `contoso-addin-data-to-excel` app copied from the Azure portal.
1. Find the key `Enter_the_Tenant_Id_Here` and replace the existing value with your Azure AD tenant/directory ID.

## Run the sample

### Start the Azure Functions project

1. Open FunctionCreateSpreadsheet.sln in Visual Studio.
1. Press **F5** (or choose **Debug** > **Start Debugging**) to build and start the Azure function project. The function will run locally using the Azure Functions Core Tools. You should see the following output in a new console window.

![Console output after starting the Azure Functions project.](./images/azure-function-running.png)

### Start the web application

1. From your shell or command line go to the `WebApplication/` folder, then run the following command:

    ```console
    npm start
    ```

1. In a browser, go to the URL `http://localhost:3000/index.html`.
    ![Contoso web app with sign-in button.](./images/web-page-sign-in-button.png)
1. Choose the **Sign In** button.
1. You will be prompted to sign in. Sign in with a user name and password from your Microsoft 365 account.

    > Note: You may also be prompted to consent to the app permissions. You'll need to consent before the app can continue successfully.

    Once you sign in, the page will display a table of sales data.
    ![Screenshot of Contoso web app listing rows of data with product name, quarter 1, quarter 2, quarter 3, and quarter 4 sales numbers](./images/web-page-product-data.png)
1. Choose the Excel icon to open a new tab with a new spreadsheet.

When the spreadsheet opens, you will see the sales data. The embedded Script Lab add-in will be available on the ribbon.

## Key parts of this sample

### Authentication

This sample was built using the code from [Vanilla JavaScript single-page application using MSAL.js to authenticate users to call Microsoft Graph](https://github.com/Azure-Samples/ms-identity-javascript-tutorial/blob/main/2-Authorization-I/1-call-graph/README.md). Please refer to the [readme](https://github.com/Azure-Samples/ms-identity-javascript-tutorial/blob/main/2-Authorization-I/1-call-graph/README.md) for more information on how the authentication works.

### Implement the Excel button

The `WebApplicatoin/App/index.html` page has an `<img>` tag that displays the Excel icon. The click handler calls `openInExcel()` which is in the `WebApplication/App/authPopup.js` file. The `openInExcel` function sends the sales data from `WebApplication/App/tableData.js` in a POST request to the `FunctionCreateSpreadsheet` Azure Functions app.

### Construct the spreadsheet

The **FunctionCreateSpreadsheet** app uses Azure Functions to provide a function that constructs the spreadsheet. The function is triggered by an HTTP POST request. The body of the request contains JSON describing rows and columns of data to populate the spreadsheet. The function expects data in the format shown in `./WebApplication/App/tableData.js`. The function returns the raw data of the new spreadsheet as a Base64 string.

The function uses the [Open XML SDK](https://learn.microsoft.com/office/open-xml/open-xml-sdk) to construct the spreadsheet in memory. The code that constructs the spreadsheet is in `FunctionCreateSpreadsheet/SpreadsheetBuilder.cs`.

- The `InsertData` method inserts the data values for the sales data table.
- The `EmbedAddin` method embeds the script lab add-in.
- Modify the `GenerateWebExtensionPart1Content` method to embed your add-in instead of the script lab add-in. Note that there is a *CUSTOM MODIFICATION BEGIN/END* section where you can specify custom properties that your add-in needs to load when it starts.

### Upload the spreadsheet to OneDrive

Once the Base64 encoded string of the new spreadsheet is returned to the `openInExcel` function, it calls `uploadFile`. The `uploadFile` function uses the Microsoft Graph API to upload the spreadsheet to the OneDrive. It creates the URI `'https://graph.microsoft.com/v1.0/me/drive/root:/` for the Microsoft Graph API and adds the folder location and filename. It adds the Base64 string as the body, and calls the `callGraph` function to make the actual REST API call.

## Modify the sample for your own web site

To repurpose the code in this sample for your own web site, you'll want to make the following changes.

### Use your own data

The sample uses mock data described in `WebApplication/App/tableData.js`. You'll need to replace this code to use the actual data from your web site. If your data uses a different data model, you'll need to update the `FunctionCreateSpreadsheet/Product.cs` file.
The `FunctionCreateSpreadsheet/SpreadsheetBuilder.cs` file contains an `InsertData` method that is bound to the product model data of this sample. You'll need to update it handle any changes you make to the data model.

### Embed your add-in

The sample embeds the script lab add-in. You'll need to change the code to embed your own add-in.
In the **SpreadsheetBuilder.cs** file, the `GenerateWebExtensionPart1Content` method sets the reference to Script Lab.

```csharp
We.WebExtension webExtension1 = new We.WebExtension() { Id = "{635BF0CD-42CC-4174-B8D2-6D375C9A759E}" };
webExtension1.AddNamespaceDeclaration("we", "http://schemas.microsoft.com/office/webextensions/webextension/2010/11");
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

In the previous code:

- The **StoreType** value is "OMEX", an alias for the Office Store.
- The **Store** value is "en-US" the culture section of the store where Script Lab is.
- The **Id** value is the Office Store's asset ID for Script Lab.

The `GenerateWebExtensionPart1Content` method contains commented code that shows how to set values for a centrally deployed add-in.

> **Note**: For more information about alternative values for these attributes, see [Automatically open a task pane with a document](https://learn.microsoft.com/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document).

The `GeneratePartContent` method specifies the visibility of the task pane when the file opens.

```csharp
Wetp.WebExtensionTaskpane webExtensionTaskpane1 = new Wetp.WebExtensionTaskpane() { DockState = "right", Visibility = true, Width = 350D, Row = (UInt32Value)4U };
```

In the previous code, the `Visibility` property of the `WebExtensionTaskpane` object is set to `true`. This ensures that the first time that the file is opened after the code is run, the task pane opens with Script Lab in it (after the user accepts the prompt to trust Script Lab). This is what we want for this sample. However, in most scenarios you will probably want this set to `false`. The effect of setting it to false is that the *first* time the file is opened, the user has to install the add-in, from the **Add-in** button on the ribbon. On every *subsequent* opening of the file, the task pane with the add-in opens automatically.

The advantage of setting this property to `false` is that you can use the Office.js to give users the ability to turn on and off the auto-opening of the add-in. Specifically, your script sets the **Office.AutoShowTaskpaneWithDocument** document setting to `true` or `false`. However, if `WebExtensionTaskpane.Visibility` is set to `true`, there is no way for Office.js or, hence, your users to turn off the auto-opening of the add-in. Only editing the OOXML of the document can change `WebExtensionTaskpane.Visibility` to false.

> **Note**: For more information about task pane visibility at the level of the Open XML that these .NET APIs represent, see [Automatically open a task pane with a document](https://learn.microsoft.com/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document).

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Solution

Solution | Authors
---------|----------
Open data from your web site in a spreadsheet | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.0  | January 31, 2023 | Initial release

## Copyright

Copyright (c) 2023 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/excel-add-in-create-spreadsheet-from-web-page" />