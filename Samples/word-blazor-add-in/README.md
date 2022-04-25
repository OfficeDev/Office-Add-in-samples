---
page_type: sample
urlFragment: word-blazor-add-in
products:
  - office-add-ins
  - office
languages:
  - javascript
  - C#
extensions:
  contentType: samples
  technologies: Add-ins
  createdDate: '04/14/2022 10:00:00 PM'
description: 'Create a Blazor Webassembly Word add-in showcasing some samples.'
---

# Create a Blazor Webassembly Word add-in

This sample shows how to build a Word add-in using .NET Blazor technologies. Blazor Webassembly allows you to build your addins using .NET, C# and JavaScript to interact with the OfficeJS API. The add-in uses JavaScript to work with the document and Office JS APIs, but you build the user interface and all other non-Office interactions in C# and .NET Core Blazor technologies.

- Work with Blazor Webassembly in the context of Office.
- Build cross-platform Office Add-ins using Blazor, C# and JavaScript Interop.
- Initialize the Office JavaScript API library in Blazor context.
- Interact with Word to manipulate paragraphs and content controls.
- Interact with document content through Office JavaScript APIs.

## Applies to

- Word on Windows, Mac, and in a browser.

## Prerequisites

- Microsoft 365 - You can get a [free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) that provides a renewable 90-day Microsoft 365 E5 developer subscription.

## Run the sample

1. Download or clone this repository.
1. Open Visual Studio 2022, and open the **Office-Add-in-samples\Samples\word-blazor-add-inword-blazor-add-in.sln** solution.
1. Choose **Debug** > **Start Debugging**. Or press F5 to start the solution.
1. When Word opens, choose **Home** > **Show Taskpane**.

Now you can try out the controls.

## Understand an Office Add-in in Blazor Context

An Office Add-in is a web application that can extend Office with additional functionality for the user. For example, an add-in can add ribbon buttons, a task pane, or a content pane with the functionality you want. Because an Office Add-in is a web application you must provide a web server to host the files.
When building the Office Add-in as a Blazor Webassembly this will allow you to build a .NET Core compliant website that interacts with the Office JS APIs. If your background is with VBA, VSTO, or COM Add-in development, you may find that building Office Add-ins using Blazor Webassembly is a more familiar development technique.

## Key parts of this sample

This sample uses a Blazor Webassembly file that can run cross-platform in various browsers supporting WASM (Webassembly). The Blazor WASM App demonstrates some basic Word functions using paragraphs and content controls including event handlers.

The purpose of this sample is to show how to build and interact with the Blazor, C# and JavaScript Interop options. If you are looking for more examples of interacting with Word and Office JS APIs, see [Script Lab](https://https://aka.ms/getscriptlab).

### Blazor pages

The **Pages** folder contains the Blazor pages, such as **HelloWorld.razor**. These also contain two code-behind pages, such as **HelloWorld.razor.cs** and **HelloWorld.razor.js**. The C# file first establishes an interop connection with the JavaScript file.

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
private async Task InsertParagraph() =>
  await JSModule.InvokeVoidAsync("insertParagraph");
```

The JavaScript runs the code to interact with the document and returns.

```javascript
export function insertParagraph() {
  return Word.run((context) => {
    // insert a paragraph at the start of the document.
    const paragraph = context.document.body.insertParagraph("Hello World from Page2.razor.js", Word.InsertLocation.start);

    // sync the context to run the previous API call, and return.
    return context.sync();
  });
}
```

This is the fundamental pattern.

1. Call **JSRuntime.InvokeAsync** to set up the interop between C# and JavaScript.
1. Use **JSModule.InvokeVoidAsync** to call JavaScript functions from your C# code.
1. Always call Office JS APIs to interact with the document from JavaScript code.

## Questions and comments

We'd love to get your feedback about this sample. Please send your feedback to us in the Issues section of this repository. Questions about developing Office Add-ins should be posted to [Microsoft Q&A](https://docs.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Solution

Solution | Authors
---------|----------
Create a Blazor Webassembly Word add-in | [Maarten van Stam](https://mvp.microsoft.com/en-us/PublicProfile/33535)

## Version history

Version  | Date | Comments
---------| -----| --------
1.0  | April 25, 2022 | Initial release

## Copyright

Copyright(c) Maarten van Stam.All rights reserved.Licensed under the MIT License.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/samples/word-blazor-add-in" />