# Office Add-ins Patterns and Practices (PnP)

Office Add-ins PnP is a community driven effort that helps developers extend, build, and provision customizations for the Office platform. The source is maintained on this GitHub repo where anyone can participate. You can provide contributions to the samples, reusable components, and documentation. Office Add-ins PnP is owned and coordinated by Office engineering teams, but the work is done by the community for the community.

## List of recent samples

- [Use a shared library to migrate your Visual Studio Tools for Office add-in to an Office web add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/VSTO-shared-code-migration). Provides a strategy for code reuse when migrating from VSTO Add-ins to Office Add-ins.
- [Dynamic DPI code samples](https://github.com/OfficeDev/PnP-OfficeAddins/tree/vstoshared/Samples/dynamic-dpi)A collection of samples for handling DPI changes in COM, VSTO, and Office Add-ins.
- [Use storage techniques to access data from an Office Add-in when offline](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/Excel.OfflineStorageAddin). Demonstrates how you can implement localStorage to enable limited functionality for your Office Add-in when a user experiences lost connection.
- **Authentication and authorization:** We're adding a new auth section and have two new samples.
  - [Office Add-in auth to Microsoft Graph](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET). Learn how to build a Microsoft Office Add-in that connects to Microsoft Graph, finds the first three workbooks stored in OneDrive for Business, fetches their filenames, and inserts the names into an Office document using Office.js.
  - [Outlook Add-in auth to Microsoft Graph](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET). Learn how to build a Microsoft Outlook Add-in that connects to Microsoft Graph, finds the first three workbooks stored in OneDrive for Business, fetches their filenames, and inserts the names into a new message compose form in Outlook.
- [Custom function batching pattern](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/Batching). If your custom functions call a remote service you may want to use a batching pattern to reduce the number of network calls to the remote service. This is useful when a spreadsheet recalculates and it contains many of your custom functions. Recalculate will result in many calls to your custom functions, but you can batch them into one or a few calls to the remote service.
- [Custom function storage](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/Storage) for custom functions. Custom functions and task panes cannot directly communicate with each other. See how to use the Storage object to send data between custom functions and task panes. This is especially useful for sharing an access token.

## Learn more

To learn more about Office Add-ins, see the [Office Add-ins documentation](https://aka.ms/office-add-ins-docs).

## Code of conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
