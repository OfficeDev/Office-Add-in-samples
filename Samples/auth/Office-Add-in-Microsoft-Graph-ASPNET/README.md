---
page_type: sample
urlFragment: office-add-in-auth-aspnet-graph
products:
- office-excel
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 5/1/2019 1:25:00 PM
description: "Learn how to build a Microsoft Office Add-in that connects to Microsoft Graph, finds the first three workbooks stored in OneDrive for Business, fetches their filenames, and inserts the names into an Office document using Office.js."
---
# Get OneDrive data using Microsoft Graph and MSAL.NET in an Office Add-in 

Learn how to build a Microsoft Office Add-in that connects to Microsoft Graph, finds the first three workbooks stored in OneDrive for Business, fetches their filenames, and inserts the names into an Office document using Office.js.

## Features
Integrating data from online service providers increases the value and adoption of your add-ins. This code sample shows you how to connect your add-in to Microsoft Graph. Use this code sample to:

* Connect to Microsoft Graph from an Office Add-in.
* Use the MSAL.NET Library to implement the OAuth 2.0 authorization framework in an add-in.
* Use the OneDrive REST APIs from Microsoft Graph.
* Show a dialog using the Office UI namespace.
* Build an Add-in using ASP.NET MVC, MSAL 3.x.x for .NET,  and Office.js. 
* Use add-in commands in an add-in.

## Applies to

-  Excel on Windows (one-time purchase and subscription)
-  PowerPoint on Windows (one-time purchase and subscription)
-  Word on Windows (one-time purchase and subscription)

## Prerequisites

To run this code sample, the following are required.

* Visual Studio 2019 or later.

* SQL Server Express (No longer automatically installed with recent versions of Visual Studio.)

* An Office 365 account which you can get by joining the [Office 365 Developer Program](https://aka.ms/devprogramsignup) that includes a free 1 year subscription to Office 365.

* At least three Excel workbooks stored on OneDrive for Business in your Office 365 subscription.

* Office on Windows, version 16.0.6769.2001 or higher.

* [Office Developer Tools](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* A Microsoft Azure Tenant. This add-in requires Azure Active Directiory (AD). Azure AD provides identity services that applications use for authentication and authorization. A trial subscription can be acquired here: [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Solution

Solution | Author(s)
---------|----------
Office Add-in Microsoft Graph ASP.NET | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.0  | July 8th, 2019| Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

----------

## Build and run the solution

### Configure the solution

1. In **Visual Studio**, choose the **Office-Add-in-Microsoft-Graph-ASPNETWeb** project. In **Properties**, ensure **SSL Enabled** is **True**. Verify that the **SSL URL** uses the same domain name and port number as those listed in the next step.
 
2. Register your application using the [Azure Management Portal](https://manage.windowsazure.com). **Log in with the identity of an administrator of your Office 365 tenancy to ensure that you are working in an Azure Active Directory that is associated with that tenancy.** To learn how to register your application, see [Register an application with the Microsoft Identity Platform](https://learn.microsoft.com/graph/auth-register-app-v2). Use the following settings:

 - REDIRCT URI: https://localhost:44301/AzureADAuth/Authorize
 - SUPPORTED ACCOUNT TYPES: "Accounts in this organizational directory only"
 - IMPLICIT GRANT: Do not enable any Implicit Grant options
 - API PERMISSIONS (Delegated permissions, not Application permissions): **Files.Read.All** and **User.Read**

	> Note: After you register your application, copy the **Application (client) ID** and the **Directory (tenant) ID** on the **Overview** blade of the App Registration in the Azure Management Portal. When you create the client secret on the **Certificates & secrets** blade, copy it too. 
	 
3.  In web.config, use the values that you copied in the previous step. Set **AAD:ClientID** to your client id, set **AAD:ClientSecret** to your client secret, and set **"AAD:O365TenantID"** to your tenant ID. 

### Run the solution

1. Open the Visual Studio solution file. 
2. Right-click **Office-Add-in-Microsoft-Graph-ASPNET** solution in **Solution Explorer** (not the project nodes), and then choose **Set startup projects**. Select the **Multiple startup projects** radio button. Make sure the project that ends with "Web" is listed first.
3. On the **Build** menu, select **Clean Solution**. When it finishes, open the **Build** menu again and select **Build Solution**.
4. In **Solution Explorer**, select the **Office-Add-in-Microsoft-Graph-ASPNET** project node (not the top solution node and not the project whose name ends in "Web").
5. In the **Properties** pane, open the **Start Document** drop down and choose one of the three options (Excel, Word, or PowerPoint).

    ![Choose the desired Office host application: Excel or PowerPoint or Word](images/SelectHost.JPG)

6. Press F5. 
7. In the Office application, choose **Insert** > **Open Add-in** in the **OneDrive Files** group to open the task pane add-in.
8. The pages and buttons in the add-in are self-explanatory. 

## Known issues

* The Fabric spinner control appears only briefly or not at all.

## Questions and comments

We'd love to get your feedback about this sample. You can send your feedback to us in the *Issues* section of this repository.
Questions about developing Office Add-ins should be posted to [Stack Overflow](http://stackoverflow.com). Ensure your questions are tagged with [office-js] and [MicrosoftGraph].

## Additional resources

* [Microsoft Graph documentation](https://learn.microsoft.com/graph/)
* [Office Add-ins documentation](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins)

## Copyright
Copyright (c) 2019 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

**Note:** The Index.cshtml file contains an image URL that tracks diagnostic data for this sample add-in. Please remove the image tag if you reuse this sample in your own code project.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/auth/Office-Add-in-Microsoft-Graph-ASPNET" />
