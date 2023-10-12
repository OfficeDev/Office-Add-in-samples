# Office Add-ins code samples

Office Add-ins code samples are provided in this repo to help you learn, study, and build great Office Add-ins!

## Getting started

The following samples show how to build the simplest Office Add-in with only a manifest, HTML web page, and a logo. They will help you understand the fundamental parts of an Office Add-in. For additional getting started information, see our [quick starts](https://learn.microsoft.com/office/dev/add-ins/quickstarts/excel-quickstart-jquery) and [tutorials](https://learn.microsoft.com/search/?terms=tutorial&scope=Office%20Add-ins).

* [Excel "Hello world" add-in](Samples/hello-world/excel-hello-world)
* [Outlook "Hello world" add-in](Samples/hello-world/outlook-hello-world)
* [PowerPoint "Hello world" add-in](Samples/hello-world/powerpoint-hello-world)
* [Word "Hello world" add-in](Samples/hello-world/word-hello-world)

### Completed tutorials

The following samples are the completed versions of various tutorials for Office Add-ins.

| Name           | Description  |
| -------------- | ------------ |
| [Excel Tutorial - Completed](Samples/tutorials/excel-tutorial) | This sample is the completed version of the [Tutorial: Create an Excel task pane add-in](https://learn.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial) that shows how to create an Excel add-in with a task pane and command ribbon buttons. The add-in shows how to create and sort a table, make a chart, freeze a row, protect a worksheet, and display a dialog box. |
| [Outlook Tutorial - Completed](Samples/tutorials/outlook-tutorial) | This sample is the completed version of the [Tutorial: Build a message compose Outlook add-in](https://learn.microsoft.com/office/dev/add-ins/tutorials/outlook-tutorial) that shows how to build an Outlook add-in that can be used in message compose mode to insert content into the body of a message. The add-in shows how to collect information from the user, fetch data from an external service, implement a function command, implement a task pane, and display a dialog box. |
| [PowerPoint Tutorial (Visual Studio) - Completed](Samples/tutorials/powerpoint-tutorial) | This sample is the completed version of the [Tutorial: Create a PowerPoint task pane add-in](https://learn.microsoft.com/office/dev/add-ins/tutorials/powerpoint-tutorialtabs=visualstudio) that shows how to create a PowerPoint add-in with a task pane. The add-in shows how to add the [Bing](https://www.bing.com) photo of the day to a slide, add text to a slide, get slide metadata, and navigate between slides. |
| [PowerPoint Tutorial (yo office) - Completed](Samples/tutorials/powerpoint-tutorial-yo) | This sample is the completed version of the [Tutorial: Create a PowerPoint task pane add-in](https://learn.microsoft.com/office/dev/add-ins/tutorials/powerpoint-tutorial?tabs=yeomangenerator) that shows how to create a PowerPoint add-in with a task pane. The add-in shows how to add an image to a slide, add text to a slide, get slide metadata, and navigate between slides. |
| [Word Tutorial - Completed](Samples/tutorials/word-tutorial) | This sample is the completed version of the [Tutorial: Create a Word task pane add-in](https://learn.microsoft.com/office/dev/add-ins/tutorials/word-tutorial) that shows how to create a Word add-in with a task pane. The add-in shows how to insert and replace text ranges, paragraphs, images, HTML, tables, and content controls. The add-in also shows how to format text and how to manage content in content controls. |

## Blazor WebAssembly

| Name           | Description  |
| -------------- | ------------ |
| [Create a Blazor WebAssembly Excel add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/blazor-add-in/excel-blazor-add-in) | Uses .NET Blazor technologies to build an Excel add-in. |
| [Create a Blazor WebAssembly Word add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/blazor-add-in/word-blazor-add-in) | Uses .NET Blazor technologies to build a Word add-in. |
| [Create a Blazor WebAssembly Outlook add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/blazor-add-in/outlook-blazor-add-in) | Uses .NET Blazor technologies to build an Outlook add-in. |

## Auth, Identity and Single Sign-on (SSO)

All of the following samples show how to access and work with a user's Microsoft Graph data using the Microsoft identity platform.

| Host | Auth approach | Platform | Name |
| ---- | ------------- | -------- | ---- |
| Outlook add-in | SSO | ASP.NET server |[Use SSO in an Outlook Add-in with ASP.NET](Samples/auth/Outlook-Add-in-SSO)|
| Outlook add-in | SSO | Node.js server |[Use SSO with event-based activation in an Outlook add-in](Samples/auth/Outlook-Add-in-SSO-events)|
| Office Add-in | SSO | ASP.NET server |[Use SSO in an Office Add-in with ASP.NET](Samples/auth/Office-Add-in-ASPNET-SSO)|
| Excel add-in | SSO | Node.js server |[Use SSO in an Office Add-in with Node.js](Samples/auth/Office-Add-in-NodeJS-SSO)|
| Excel add-in | MSAL | React SPA |[Use MSAL.js for auth and Microsoft Graph in an Excel add-in](Samples/auth/Office-Add-in-Microsoft-Graph-React)|
| Excel add-in | MSAL | ASP.NET server |[Use MSAL.NET for auth and Microsoft Graph in an Excel add-in](Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)|
| Outlook add-in | MSAL | ASP.NET server |[Use MSAL.NET for auth and Microsoft Graph in an Outlook add-in](Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)|

## Office

| Name           | Description  |
| -------------- | ------------ |
| [Save custom settings in your Office Add-in](Samples/office-add-in-save-custom-settings) | Shows how to save custom settings inside an Office Add-in. The add-in stores data as key/value pairs, using the JavaScript API for Office property bag, browser cookies, web storage (localStorage and sessionStorage), or by storing the data in a hidden div in the document. |

## Outlook

| Name           | Description  |
| -------------- | ------------ |
| [Encrypt attachments, process meeting request attendees, and react to appointment date/time changes using Outlook event-based activation](Samples/outlook-encrypt-attachments) | Shows how to use event-based activation to encrypt attachments when added by the user. Also shows event handling for recipients changed in a meeting request, and changes to the start or end date or time in a meeting request. |
| [Identify and tag external recipients using Outlook event-based activation](Samples/outlook-tag-external) | Uses event-based activation to run an Outlook add-in when the user changes recipients while composing a message. The add-in also uses the appendOnSendAsync API to add a disclaimer.|
| [Set your signature using Outlook event-based activation](Samples/outlook-set-signature) | Uses event-based activation to run an Outlook add-in when the user creates a new message or appointment.|
| [Verify the color categories of a message or appointment before it's sent using Smart Alerts](Samples/outlook-check-item-categories/) | Uses Outlook Smart Alerts to verify that required color categories are applied to a new message or appointment before it's sent.|
| [Verify the sensitivity label of a message](Samples/outlook-verify-sensitivity-label/) | Uses the sensitivity label API in an event-based add-in to verify and apply the **Highly Confidential** sensitivity label to applicable outgoing messages. |

## Excel

| Name           | Description  |
| -------------- | ------------ |
| [Data types explorer](Samples/excel-data-types-explorer) | Builds an Excel add-in that allows you to create and explore data types in your workbooks. Data types enable add-in developers to organize complex data structures as objects, such as formatted number values, web images, and entity values. |
| [Open in Teams](Samples/excel-open-in-teams) | Creates a new Excel spreadsheet in Microsoft Teams containing data you define.|
| [Insert an external Excel file and populate it with JSON data](Samples/excel-insert-file)  | Insert an existing template from an external Excel file into the currently open Excel file. Then retrieve data from a JSON web service and populate the template for the customer. |
| [Create custom contextual tabs on the ribbon](Samples/office-contextual-tabs) | This sample shows how to create a custom contextual tab on the ribbon in the Office UI. The sample creates a table, and when the user moves the focus inside the table, the custom tab is displayed. When the user moves outside the table, the custom tab is hidden. |
| [Use keyboard shortcuts for Office add-in actions](Samples/excel-keyboard-shortcuts) | Shows how to set up a basic Excel add-in project that utilizes keyboard shortcuts. |
| [Custom function sample using web worker](Excel-custom-functions/web-worker) | Shows how to use web workers in custom functions to prevent blocking the UI of your Office Add-in. |
| [Use storage techniques to access data from an Office Add-in when offline](Samples/Excel.OfflineStorageAddin) | Demonstrates how you can implement localStorage to enable limited functionality for your Office Add-in when a user experiences lost connection. |
| [Custom function batching pattern](Excel-custom-functions/Batching)| Batch multiple calls into a single call to reduce the number of network calls to a remote service.|

## Word

| Name           | Description  |
| -------------- | ------------ |
| [Get, edit, and set OOXML content in a Word document with a Word add-in](Samples/word-add-in-get-set-edit-openxml)| Shows how to get, edit, and set OOXML content in a Word document.|
| [Load and write Open XML in your Word add-in](Samples/word-add-in-load-and-write-open-xml) | Shows how to add a variety of rich content types to a Word document using the **setSelectedDataAsync** method with **ooxml** coercion type.|

## Shared JavaScript runtime

Check out these samples if you want to take advantage of the [shared runtime](https://learn.microsoft.com/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime) for your Office Add-in.

| Date               | Name           | Description  |
| ------------------ | -------------- | ------------ |
| [Share global data with a shared runtime](Samples/excel-shared-runtime-global-state) | Shows how to set up a basic project that uses the shared runtime to run code for ribbon buttons, task pane, and custom functions in a single browser runtime. |
| [Manage ribbon and task pane UI, and run code on doc open](Samples/excel-shared-runtime-scenario) | Shows how to create contextual ribbon buttons that are enabled based on the state of your add-in. |

## Additional samples

| Name           | Description  |
| -------------- | ------------ |
| [Use a shared library to migrate your Visual Studio Tools for Office add-in to an Office web add-in](Samples/VSTO-shared-code-migration) | Provides a strategy for code reuse when migrating from VSTO Add-ins to Office Add-ins. |
| [Integrate an Azure function with your Excel custom function](Excel-custom-functions/AzureFunction) | Learn how to integrate Azure functions with custom functions to move to the cloud or integrate additional services. |
| [Dynamic DPI code samples](Samples/dynamic-dpi) | A collection of samples for handling DPI changes in COM, VSTO, and Office Add-ins. |

## Learn more

To learn more about Office Add-ins, see the [Office Add-ins documentation](https://aka.ms/office-add-ins-docs).

## Questions and feedback

* Did you experience any problems with a sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
* We'd love to get your feedback about the samples. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
* For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Join the Microsoft 365 Developer Program

Get a free sandbox, tools, and other resources you need to build solutions for the Microsoft 365 platform.

* [Free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) Get a free, renewable 90-day Microsoft 365 E5 developer subscription.
* [Sample data packs](https://developer.microsoft.com/microsoft-365/dev-program#Sample) Automatically configure your sandbox by installing user data and content to help you build your solutions.
* [Access to experts](https://developer.microsoft.com/microsoft-365/dev-program#Experts) Access community events to learn from Microsoft 365 experts.
* [Personalized recommendations](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations) Find developer resources quickly from your personalized dashboard.

## Community

The Office Add-ins code samples are a community-driven effort that helps developers extend, build, and provision customizations for the Office platform. The source is maintained on this GitHub repo where anyone can participate. You can provide contributions to the samples, reusable components, and documentation. Office Add-ins code samples is owned and coordinated by Office engineering teams, but the work is done by the community for the community.

Please read the [Contribute](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/CONTRIBUTING.md) page to learn how to be an active part of this community.

## Code of conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
