---
page_type: sample
urlFragment: excel-add-in-create-spreadsheet-from-web-page
products:
  - m365
  - office
  - office-excel
languages:
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: 01/31/2023 1:25:00 PM
description: "Learn how to create a spreadsheet from your web page, populate it with data, and embed your Excel add-in."
---

# Create a spreadsheet from your web page, populate it with data, and embed your Excel add-in

This sample accomplishes the following tasks.

- Creates a new Excel spreadsheet from a web page.
- Populates the spreadsheet with data from the web page.
- Embeds the Script Lab add-in into the Excel spreadsheet.
- Opens the spreadsheet in a new browser tab.

![Sequence diagram showing an "Open in Microsoft Excel" button on your web page that creates a spreadsheet with your data which contains your add-in](./images/open-in-excel-overview.png)

This sample implements the pattern described in [Create an Excel spreadsheet from your web page, populate it with data, and embed your Office Add-in](https://learn.microsoft.com/office/dev/add-ins/excel/pnp-open-in-excel).

## Applies to

- Microsoft Excel

## Prerequisites

- [Node.js](https://nodejs.org/) version 16 or later.
- [Visual Studio Code](https://code.visualstudio.com/Download).
- A Microsoft 365 account. You can get one if you qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram); for details, see the [FAQ](https://learn.microsoft.com/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-). Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g).
- At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription.

## Set up the sample

### Step 1: Clone or download this repository

From your shell or command line, use the following command:

```console
git clone https://github.com/OfficeDev/Office-Add-in-samples.git
```

Or, you can download and extract the repository *.zip* file.

> :warning: To avoid path length limitations on Windows, clone the repository into a directory near the root of your drive.

### Step 2: Install project dependencies

Go to the sample folder and install the dependencies:

```console
cd Samples/excel-create-worksheet-from-web-site
npm install
```

### Step 3: Register the sample applications in your tenant

#### Register the client app (contoso-addin-data-to-excel)

1. To register your app, go to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page.
1. Sign in with the **_admin_** credentials to your Microsoft 365 tenancy. For example, **MyName@contoso.onmicrosoft.com**.
1. Select **New registration**. On the **Register an application** page, set the values as follows.

   - Set **Name** to `contoso-addin-data-to-excel`.
   - Set **Supported account types** to **Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)**.
   - Select **Register**.

1. In the **Overview** blade, find and note the **Application (client) ID**. Save this value to use in the project files later in these steps.

#### Add redirect URIs

1. In the app's registration screen, select the **Manage > Authentication** blade to the left.
1. Select **Add Redirect URI**.
1. Select the **Single-page application** option.
1. In the **Redirect URI** section enter `http://localhost:3000` as the redirect URI:
1. Select **Configure**.
1. Select the **Single-page application** option again.
1. In the **Redirect URI** section enter `http://localhost:3000/redirect` as a new redirect URI.
1. Select **Configure**.

#### Add delegated permissions

Since this app signs-in users, add delegated permissions, which are required by apps that sign in users.

1. In the app's registration screen, select the **Manage > API permissions** blade on the left pane.
1. Select **Add a permission**.
1. Select **Microsoft Graph**. Then select **Delegated permissions**.
1. Use the search box and list of permissions to find and select the following permissions.
    - **User.Read**
    - **Contacts.Read**
    - **Files.ReadWrite**
1. Select the **Add permissions** button at the bottom.

##### Configure optional claims

1. Still on the same app registration, select the **Manage > Token configuration** blade on the left pane.
1. Select **Add optional claim**:
    1. Select **optional claim type**, and then choose **ID**.
    1. Select the optional claim **acct**.
    > Provides user's account status in tenant. If the user is a **member** of the tenant, the value is *0*. If they're a **guest**, the value is *1*.
    1. Select the optional claim **login_hint**.
    > An opaque, reliable login hint claim. This claim is the best value to use for the `login_hint` OAuth parameter in all flows to get SSO. See [optional claims](https://learn.microsoft.com/entra/identity-platform/optional-claims) for more details on this optional claim.
    1. Select **Add** to save your changes.

##### Configure the client app (contoso-addin-data-to-excel) to use your app registration

Open the project in your IDE (like Visual Studio or Visual Studio Code) to configure the code.

> In the steps below, "ClientID" is the same as "Application ID" or "AppId".

1. Open the `WebApplication/App/authConfig.js` file.
1. Find the key `Enter_the_Application_Id_Here` and replace the existing value with the application ID (clientId) of `contoso-addin-data-to-excel` app copied from the Azure portal.
1. Find the key `Enter_the_Tenant_Id_Here` and replace the existing value with your Microsoft Entra ID tenant/directory ID.

## Run the sample

### Start the Node.js server

1. From your shell or command line in the sample root folder, run the following command:

    ```console
    npm start
    ```

   This command starts the Node.js server. The server serves the web application and provides the API endpoint for creating spreadsheets.

1. In a browser, go to the URL `http://localhost:3000/index.html`.
    ![Contoso web app with sign-in button.](./images/web-page-sign-in-button.png)
1. Select the **Sign In** button.
1. You're prompted to sign in. Sign in with a user name and password from your Microsoft 365 account.

    > Note: You might also be prompted to consent to the app permissions. You need to consent before the app can continue successfully.

    After you sign in, the page displays a table of sales data.
    ![Screenshot of Contoso web app listing rows of data with product name, quarter 1, quarter 2, quarter 3, and quarter 4 sales numbers](./images/web-page-product-data.png)
1. Select the Excel icon to open a new tab with a new spreadsheet.

When the spreadsheet opens, you see the sales data. The embedded Script Lab add-in is available on the ribbon. After the first time you open the file and accept the prompt to trust Script Lab, the add-in task pane automatically opens whenever you open the spreadsheet.

## Key parts of this sample

### Authentication

This sample uses code from [Vanilla JavaScript single-page application using MSAL.js to authenticate users to call Microsoft Graph](https://github.com/Azure-Samples/ms-identity-javascript-tutorial/blob/main/2-Authorization-I/1-call-graph/README.md). For more information about how authentication works, see the [readme](https://github.com/Azure-Samples/ms-identity-javascript-tutorial/blob/main/2-Authorization-I/1-call-graph/README.md).

### Implement the Excel button

The [WebApplication/App/index.html](WebApplication/App/index.html) page has an `<img>` tag that displays the Excel icon. The click handler calls `openInExcel()`, which is in the [WebApplication/App/authPopup.js](WebApplication/App/authPopup.js) file. The `openInExcel` function sends the sales data from [WebApplication/App/tableData.js](WebApplication/App/tableData.js) in a POST request to the Node.js server endpoint at `/api/create-spreadsheet`.

### Construct the spreadsheet

The Node.js server in [server.js](server.js) provides an API endpoint at `/api/create-spreadsheet` that constructs the spreadsheet. The endpoint is triggered by an HTTP POST request. The body of the request contains JSON describing rows and columns of data to populate the spreadsheet. The endpoint expects data in the format shown in [WebApplication/App/tableData.js](WebApplication/App/tableData.js). The endpoint returns the raw data of the new spreadsheet as a binary blob.

The server uses the [ExcelJS](https://github.com/exceljs/exceljs) library to construct the spreadsheet in memory, then uses [JSZip](https://www.npmjs.com/package/jszip) and [xml2js](https://www.npmjs.com/package/xml2js) to manipulate the Office Open XML (OOXML) structure to embed the Script Lab add-in. The code that constructs the spreadsheet is in the `/api/create-spreadsheet` endpoint handler in [server.js](server.js).

- The endpoint inserts the data values from the request into the worksheet.
- Formatting is applied to the header row (bold font, gray background).
- The `embedAddin` function manipulates the OOXML structure by:
  - Adding `webextension1.xml` with Script Lab add-in reference.
  - Adding `taskpanes.xml` to configure the task pane behavior.
  - Updating `[Content_Types].xml` to register the web extension parts.
  - Updating `workbook.xml.rels` to link the taskpane configuration.
  - Setting `visibility="0"` so users install the add-in once, then it auto-opens on subsequent opens.

### Upload the spreadsheet to OneDrive

After the `openInExcel` function receives the binary blob of the new spreadsheet, it calls `uploadFile`. The `uploadFile` function uses the Microsoft Graph API to upload the spreadsheet to OneDrive. It creates the URI `'https://graph.microsoft.com/v1.0/me/drive/root:/` for the Microsoft Graph API and adds the folder location and filename. It passes the binary blob as the body, and calls the `callGraph` function to make the actual REST API call.

## Modify the sample for your own web site

To repurpose the code in this sample for your own web site, make the following changes.

### Use your own data

The sample uses mock data described in [WebApplication/App/tableData.js](WebApplication/App/tableData.js). Replace this code to use the actual data from your web site. The data structure is simple JSON with rows and columns.

```javascript
{
  rows: [
    { columns: [{ value: 'Header1' }, { value: 'Header2' }] },
    { columns: [{ value: 'Data1' }, { value: 'Data2' }] }
  ]
}
```

The spreadsheet creation endpoint in [server.js](server.js) iterates through this structure to populate the Excel worksheet. If you need to change the data model, update both the client-side data structure and the server endpoint handler.

### Embed your add-in

The sample embeds the Script Lab add-in from Microsoft AppSource. To embed your own add-in instead, modify the `createWebExtensionXml` function in [server.js](server.js).

Use the following key configuration in the webextension XML:

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{YOUR-ADDIN-GUID}">
    <we:reference id="ASSET-OR-GUID" version="1.0.0.0" store="STORE-LOCALE" storeType="STORE-TYPE"/>
    <we:properties>
        <we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
    </we:properties>
</we:webextension>
```

**Key attributes to configure:**

- **id**: A GUID for the web extension instance (can be any GUID).
- **reference id**: Depends on `storeType`:
  - For AppSource (`storeType="OMEX"`): Use the AppSource asset ID (for example, "wa104380862" for Script Lab).
  - For centralized deployment (`storeType="EXCatalog"`): Use your add-in's GUID from the manifest.
  - For network share (`storeType="FileSystem"`): Use your add-in's GUID from the manifest.
- **version**: Your add-in version.
- **store**: For AppSource, use locale ("en-US"). For centralized deployment, use "EXCatalog".
- **storeType**: "OMEX" (AppSource), "EXCatalog" (centralized), "FileSystem" (network share), or "WOPICatalog" (WOPI hosts).

**Task pane visibility:**

In the `createTaskpaneXml` function, the `visibility` attribute controls behavior:

- `visibility="0"`: User must install add-in from ribbon first. After installation, task pane auto-opens on subsequent file opens. Users can toggle auto-open via Office.js.
- `visibility="1"`: Task pane opens immediately when file opens (prompts to trust add-in). Users can't toggle off via Office.js.

For more information, see [Automatically open a task pane with a document](https://learn.microsoft.com/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document).

## Security notes

- This sample might use packages that have security problems. Run `npm audit` to find any security vulnerabilities and update packages regularly.
- This sample runs on `localhost` for development purposes only. In production, you should:
  - Use HTTPS.
  - Implement proper CORS configuration.
  - Add authentication and authorization for the API endpoint.
  - Validate and sanitize all input data.
  - Consider rate limiting to prevent abuse.

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Solution

Solution | Authors
---------|----------
Open data from your web page in a spreadsheet | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.0  | January 31, 2023 | Initial release
1.1  | April 24, 2024 | Update package versions
1.2 | January 7, 2026 | Changed to no longer require Visual Studio

## Copyright

Copyright (c) 2023 Microsoft Corporation. All rights reserved.

This project adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/excel-add-in-create-spreadsheet-from-web-page" />