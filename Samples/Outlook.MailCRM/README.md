# [ARCHIVED] DXDemos.Office365 #

**Note:** This sample is archived and no longer actively maintained. Security vulnerabilities may exist in the project, or its dependencies. If you plan to reuse or run any code from this repo, be sure to perform appropriate security checks on the code or dependencies first. Do not use this project as the starting point of a production Office Add-in. Always start your production code by using the Office/SharePoint development workload in Visual Studio, or the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), and follow security best practices as you develop the add-in.

To restore the project dependencies, rename the following file.

- DXDemos.Office365/packages-archive.config -> DXDemos.Office365/packages.config

DXDemos.Office365 is a solution containing several store patterns and practices. The apps and add-ins in this solution are near store-ready and illustrate a number of interesting patterns.

# Getting Started #
After you clone the Store PnP repo, you should open the web.config and update the configuration with your own values. This includes appsettings values for **ida:ClientID**, **ida:Password**, and **baseUrl**. The DXDemos.Office365 solution also uses DocumentDB (Azure's NoSQL alternative to Mongo). It could be re-factored to to use any data store, but it is easiest to provision a DocumentDB account and update the appsettings for **ddb:endpoint**, **ddb:authKey**, and **ddb:database** (the values checked-in are invalid).

    <!-- The following are AAD settings for the app -->
    <add key="ida:ClientID" value="cb88b4df-db4b-4cbe-be95-b40f76dccb14" />
    <add key="ida:Password" value="c23vRAjoINSuKnj7tEDYCqwi7pN3cXy2pPdOecv54O4=" />
    <add key="ida:AuthorizationUri" value="https://login.microsoftonline.com" />

    <!--baseUrl is used for reply URL in OAuth flows...two listed to handle debug/release-->
    <!--<add key="baseUrl" value="https://dxsamples.azurewebsites.net/" />-->
    <add key="baseUrl" value="https://localhost:44321/" />
    
    <!-- The following are setting for Azure DocumentDB, which is the data store for the app -->
    <add key="ddb:endpoint" value="https://dxdemo.documents.azure.com:443/" />
    <add key="ddb:authKey" value="WkPRneEPSrhCdaEVd30e+ag00pbe8B0Ilzn4idJqakWMtFgz7oFBXlrjZvNTqPKzHG25ZHAwZxJrtydo1gBiAw==" />
    <add key="ddb:database" value="dxdemo" />

When you setup this app in Azure AD, it needs the following permissions: **Access directory as signed-in user** and **Read and Write the signed-in users files**. It also needs **Read and write user mail** from Office 365 Exchange Online** (this is for user pictures, which is currently broken in the Unified API).

# Apps/Add-ins in DXDemos.Office365 #
This section will outline all the apps and add-ins contained in the DXDemos.Office365 solution.
## DXDemos.Office365.MailCRM ##
DXDemos.Office365.MailCRM is a Read Outlook Mail Add-in that functions as a CRM solution with mail. This follows a popular scenario of providing a mail add-in that looks up contact details in another system.

The solution is hosted in a ASP.NET MVC project, but is built as a single-page application (SPA) with AngularJS. Web APIs provide the back-end services and integration with DocumentDB. A single OAuth controller provides the sign-in logic and helps cache the users refresh tokens.


<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/Outlook.MailCRM" />