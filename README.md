# Office Add-ins code samples

Office Add-ins code samples are provided in this repo to help you learn, study, and built great Office Add-ins!

## Auth, Identity and SSO

| Name           | Description  |
| -------------- | ------------ |
| [Single Sign-on (SSO) Sample Outlook Add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Outlook-Add-in-SSO) | The sample implements an Outlook add-in that uses Office's SSO feature to give the add-in access to Microsoft Graph data.|
| [Get OneDrive data using Microsoft Graph and msal.js in an Office Add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React) | Learn how to build a Microsoft Office Add-in, as a single-page application (SPA) with no backend, that connects to Microsoft Graph, and access workbooks stored in OneDrive for Business to update a spreadsheet.  |
| [Office Add-in auth to Microsoft Graph](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET) | Learn how to build a Microsoft Office Add-in that connects to Microsoft Graph, and access workbooks stored in OneDrive for Business to update a spreadsheet. |
| [Outlook Add-in auth to Microsoft Graph](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET). | Learn how to build a Microsoft Outlook Add-in that connects to Microsoft Graph, and access workbooks stored in OneDrive for Business to compose a new email message. |
| [Single Sign-on (SSO) Office Add-in with ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO) | Implements an Office Add-in that uses the getAccessToken API in Office.js to give the add-in access to Microsoft Graph data. This sample is built on ASP.NET. |
| [Single Sign-on (SSO) Office Add-in with Node.js](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO) | Implements an Office Add-in that uses the getAccessToken API in Office.js to give the add-in access to Microsoft Graph data. This sample is built on Node.js.|

## Outlook

| Name           | Description  |
| -------------- | ------------ |
| [Use Outlook event-based activation to indicate external recipients (preview)](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-tag-external) | Uses event-based activation to run an Outlook add-in when the user changes recipients while composing a message. The add-in also uses the appendOnSendAsync API to add a disclaimer.|
| [Use Outlook event-based activation to set the signature](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-set-signature) | Uses event-based activation to run an Outlook add-in when the user creates a new message or appointment.|

## Excel

| Name           | Description  |
| -------------- | ------------ |
| [Open in Teams](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/excel-open-in-teams) | Creates a new Excel spreadsheet in Microsoft Teams containing data you define.|
| [Create custom contextual tabs on the ribbon](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/office-contextual-tabs) | This sample shows how to create a custom contextual tab on the ribbon in the Office UI. The sample creates a table, and when the user moves the focus inside the table, the custom tab is displayed. When the user moves outside the table, the custom tab is hidden. |
| [Use keyboard shortcuts for Office add-in actions](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/excel-keyboard-shortcuts) | Shows how to set up a basic Excel add-in project that utilizes keyboard shortcuts. |
| [Custom function sample using web worker](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Excel-custom-functions/web-worker) | Shows how to use web workers in custom functions to prevent blocking the UI of your Office Add-in. |
| [Use storage techniques to access data from an Office Add-in when offline](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/Excel.OfflineStorageAddin) | Demonstrates how you can implement localStorage to enable limited functionality for your Office Add-in when a user experiences lost connection. |
| [Custom function batching pattern](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Excel-custom-functions/Batching)| Batch multiple calls into a single call to reduce the number of network calls to a remote service.|

## Shared JavaScript runtime

Check out these samples if you want to take advantage of the [shared runtime](https://docs.microsoft.com/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime) for your Office Add-in.

| Date               | Name           | Description  |
| ------------------ | -------------- | ------------ |
[Share global data with a shared runtime](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/excel-shared-runtime-global-state) | Shows how to set up a basic project that uses the shared runtime to run code for ribbon buttons, task pane, and custom functions in a single browser runtime. |
| [Manage ribbon and task pane UI, and run code on doc open](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/excel-shared-runtime-scenario) | Shows how to create contextual ribbon buttons that are enabled based on the state of your add-in. |

## Additional samples

| Name           | Description  |
| -------------- | ------------ |
| [Use a shared library to migrate your Visual Studio Tools for Office add-in to an Office web add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/VSTO-shared-code-migration) | Provides a strategy for code reuse when migrating from VSTO Add-ins to Office Add-ins. |
| [Integrate an Azure function with your Excel custom function](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Excel-custom-functions/AzureFunction) | Learn how to integrate Azure functions with custom functions to move to the cloud or integrate additional services. |
| [Dynamic DPI code samples](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/dynamic-dpi) | A collection of samples for handling DPI changes in COM, VSTO, and Office Add-ins. |

## Learn more

To learn more about Office Add-ins, see the [Office Add-ins documentation](https://aka.ms/office-add-ins-docs).

## Join the Microsoft 365 Developer Program

Get a free sandbox, tools, and other resources you need to build solutions for the Microsoft 365 platform.

- [Free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) Get a free, renewable 90-day Microsoft 365 E5 developer subscription.
- [Sample data packs](https://developer.microsoft.com/microsoft-365/dev-program#Sample) Automatically configure your sandbox by installing user data and content to help you build your solutions.
- [Access to experts](https://developer.microsoft.com/microsoft-365/dev-program#Experts) Access community events to learn from Microsoft 365 experts.
- [Personalized recommendations](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations) Find developer resources quickly from your personalized dashboard.

## Patterns and Practices community

Office Add-ins PnP is a community driven effort that helps developers extend, build, and provision customizations for the Office platform. The source is maintained on this GitHub repo where anyone can participate. You can provide contributions to the samples, reusable components, and documentation. Office Add-ins PnP is owned and coordinated by Office engineering teams, but the work is done by the community for the community.

## Code of conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
