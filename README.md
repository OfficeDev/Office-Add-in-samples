# Office Add-ins Patterns and Practices (PnP)

Office Add-ins PnP is a community driven effort that helps developers extend, build, and provision customizations for the Office platform. The source is maintained on this GitHub repo where anyone can participate. You can provide contributions to the samples, reusable components, and documentation. Office Add-ins PnP is owned and coordinated by Office engineering teams, but the work is done by the community for the community.

## List of samples

| Date               | Name           | Description  |
| ------------------ | -------------- | ------------ |
| December 28, 2020 | [Custom function sample using web worker](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/web-worker) | This sample shows how to use web workers in custom functions to prevent blocking the UI of your Office Add-in. |
| November 5, 2020   | [(Preview) Use keyboard shortcuts for Office add-in actions](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) | Shows how to set up a basic Excel add-in project that utilizes keyboard shortcuts. Currently, the shortcuts are configured to show and hide the task pane as well as cycle through colors for a selected cell. |
| March 9, 2020         | [Manage ribbon and task pane UI, and run code on doc open](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-shared-runtime-scenario) | This sample shows how to create contextual ribbon buttons that are enabled based on the state of your add-in. It also shows how to use the Office.js API to show or hide the task pane. This sample also demonstrates how to run code when the task pane is closed, such as on document open. |
| January 29, 2020   | [Use a shared library to migrate your Visual Studio Tools for Office add-in to an Office web add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/VSTO-shared-code-migration) | Provides a strategy for code reuse when migrating from VSTO Add-ins to Office Add-ins. |
| November 12, 2019  | [Get OneDrive data using Microsoft Graph and msal.js in an Office Add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-React) | Learn how to build a Microsoft Office Add-in, as a single-page application (SPA) with no backend, that connects to Microsoft Graph, and access workbooks stored in OneDrive for Business to update a spreadsheet.  |
| October 2, 2019    | [Integrate an Azure function with your Excel custom function](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AzureFunction) | Learn how to integrate Azure functions with custom functions to move to the cloud or integrate additional services. |
| September 30, 2019 | [Dynamic DPI code samples](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/dynamic-dpi) | A collection of samples for handling DPI changes in COM, VSTO, and Office Add-ins. |
| July 19, 2019      | [Use storage techniques to access data from an Office Add-in when offline](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/Excel.OfflineStorageAddin) | Demonstrates how you can implement localStorage to enable limited functionality for your Office Add-in when a user experiences lost connection. |
| May 1, 2019        | [Office Add-in auth to Microsoft Graph](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET) | Learn how to build a Microsoft Office Add-in that connects to Microsoft Graph, and access workbooks stored in OneDrive for Business to update a spreadsheet. |
| May 1, 2019 | [Outlook Add-in auth to Microsoft Graph](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET).  | Learn how to build a Microsoft Outlook Add-in that connects to Microsoft Graph, and access workbooks stored in OneDrive for Business to compose a new email message. |
| May 1, 2019 | [Custom function batching pattern](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/Batching)| Batch multiple calls into a single call to reduce the number of network calls to a remote service.|

## Learn more

To learn more about Office Add-ins, see the [Office Add-ins documentation](https://aka.ms/office-add-ins-docs).

## Join the Microsoft 365 Developer Program
Get a free sandbox, tools, and other resources you need to build solutions for the Microsoft 365 platform.
- [Free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) Get a free, renewable 90-day Microsoft 365 E5 developer subscription.
- [Sample data packs](https://developer.microsoft.com/microsoft-365/dev-program#Sample) Automatically configure your sandbox by installing user data and content to help you build your solutions.
- [Access to experts](https://developer.microsoft.com/microsoft-365/dev-program#Experts) Access community events to learn from Microsoft 365 experts.
- [Personalized recommendations](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations) Find developer resources quickly from your personalized dashboard.

## Code of conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
