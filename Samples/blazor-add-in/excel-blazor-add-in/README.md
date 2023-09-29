---
page_type: sample
urlFragment: excel-blazor-add-in 
products:
  - office-add-ins
  - office
languages:
  - javascript
  - C#
extensions:
  contentType: samples
  technologies: 
  - Add-ins
  createdDate: '05/26/2022 10:00:00 PM'
description: 'Create a Blazor Webassembly Excel add-in showcasing some samples.'
---

# Create a Blazor Webassembly Excel add-in

This sample shows how to build an Excel add-in using .NET Blazor technologies. Blazor Webassembly allows you to build Office Add-ins using .NET, C#, and JavaScript to interact with the Office JS API. The add-in uses JavaScript to work with the document and Office JS APIs, but you build the user interface and all other non-Office interactions in C# and .NET Core Blazor technologies.

- Work with Blazor Webassembly in the context of Office.
- Build cross-platform Office Add-ins using Blazor, C# and JavaScript Interop.
- Initialize the Office JavaScript API library in Blazor context.
- Interact with Excel to manipulate worksheets.
- Interact with workbook content through Office JavaScript APIs.

## Applies to

- Excel on the web, Windows, and Mac.

## Prerequisites

- Microsoft 365 - Get a [free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) that provides a renewable 90-day Microsoft 365 E5 developer subscription.

## Run the sample

1. Download or clone the [Office Add-ins samples repository](https://github.com/OfficeDev/Office-Add-in-samples).
1. Open Visual Studio 2022 and open the **Office-Add-in-samples\Samples\blazor-add-in\excel-blazor-add-in\excel-blazor-add-in.sln** solution. (Do **not** open Visual Studio "as administrator". There is a bug that will prevent the add-in from sideloading when Visual Studio is elevated in this way.)
1. Choose **Debug** > **Start Debugging**. Or press F5 to start the solution.
1. When Excel opens, choose **Home** > **Show Taskpane**.

Next, try out the controls.

## Understand an Office Add-in in Blazor Context

An Office Add-in is a web application that extends Office with additional functionality for the user. For example, an add-in can add ribbon buttons, a task pane, or a content pane with the functionality you want. Because an Office Add-in is a web application, you must provide a web server to host the files.
Building the Office Add-in as a Blazor Webassembly allows you to build a .NET Core compliant website that interacts with the Office JS APIs. If your background is with VBA, VSTO, or COM add-in development, you may find that building Office Add-ins using Blazor Webassembly is a familiar development technique.

## Key parts of this sample

This sample uses a Blazor Webassembly file that runs cross-platform in various browsers supporting WASM (Webassembly). The Blazor WASM App demonstrates some basic Excel functions.

The purpose of this sample is to show you how to build and interact with the Blazor, C# and JavaScript Interop options. If you're looking for more examples of interacting with Excel and Office JS APIs, see [Script Lab](https://aka.ms/getscriptlab).

### Blazor pages

The **Pages** folder contains the Blazor pages, such as **Index.razor**. Each **.razor** page also contains two code-behind pages, for example, **Index.razor.cs** and **Index.razor.js**. The C# file first establishes an interop connection with the JavaScript file.

```csharp
protected override async Task OnAfterRenderAsync(bool firstRender)
{
  if (firstRender)
  {
    JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/Index.razor.js");
  }
}
```

For any events that need to interact with the Office document, the C# file calls through interop to the JavaScript file.

```csharp
private async Task HelloButton() =>
  await JSModule.InvokeVoidAsync("helloButton");
```

The JavaScript runs the code to interact with the worksheet and returns.

```javascript
export function helloButton() {
    return Excel.run(context => {
        // Insert text 'Hello world!' into cell A1.
        context.workbook.worksheets.getActiveWorksheet().getRange("A1").values = [['Hello world!']];

        // sync the context to run the previous API call, and return.
        return context.sync();
    });
}
```

The fundamental pattern includes the following steps.

1. Call **JSRuntime.InvokeAsync** to set up the interop between C# and JavaScript.
1. Use **JSModule.InvokeVoidAsync** to call JavaScript functions from your C# code.
1. Call Office JS APIs to interact with the worksheet from JavaScript code.

## Debugging

This sample is configured to support debugging both JavaScript and C# files. If you create a new Blazor project, you will need to make the following file updates to support C# debugging.

1. In the **launchSettings.json** file of the web project, make sure all instances of `launchBrowser` are set to `false`.
1. In the **<projectName>.csproj.user** file of the add-in project, add the `<BlazorAppUrl>` and `<InspectUri>` elements as shown in the following example XML.

**Note:** The port number in the following XML is 7126. You must change the number to the port number specified in the **launchSettings.json** file for your web project.

```xml
<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="Current" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <PropertyGroup>
    <BlazorAppUrl>https://localhost:7126/</BlazorAppUrl>
    <InspectUri>{wsProtocol}://{url.hostname}:{url.port}/_framework/debug/ws-proxy?browser={browserInspectUri}</InspectUri>
  </PropertyGroup>
</Project>
```

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Solution

Solution | Authors
---------|----------
Create a Blazor Webassembly Excel add-in | [Maarten van Stam](https://mvp.microsoft.com/en-us/PublicProfile/33535)

## Version history

Version  | Date | Comments
---------| -----| --------
1.0  | May 27, 2022 | Initial release

## Copyright

Copyright(c) Maarten van Stam. All rights reserved.Licensed under the MIT License.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

**Note**: The index.html file contains an image URL that tracks diagnostic data for this sample add-in. Please remove the image tag if you reuse this sample in your own code project.

<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/samples/blazor-add-in/excel-blazor-add-in" />
