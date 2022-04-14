---
page_type: sample
urlFragment: office-add-in-samples
products:
  - office-add-ins
  - office
languages:
  - javascript
  - C#
  - Blazor
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: '04/14/2022 10:00:00 PM'
description: 'Create a Blazor Webassembly Office Add-in showcasing some samples.'
---

# Create a Blazor Webassembly Office Add-in showcasing some samples

## Summary

This sample shows how to build an Office Add-in using .NET Blazor technologies.
Blazor Webassembly allows you to build your addins using .NET, C# and JavaScript to interact with the OfficeJS API.

The Add-in works exactly the same as JavaScript based Office Add-ins but this will allow you to build the User Interface and all other not Office interactions in C# and .NET Core Blazor technologies.
![Diagram showing a hello project consists of a manifest, HTML page, and image assets.](./images/hello-world-introduction.png)

## Features

- Learn how to work with Blazor Webassembly in the context of Office
- Learn how to build cross platform Office Add-ins using Blazor, C# and JavaScript Interop
- Learn how to initialize the Office JavaScript API library in Blazor context.
- Learn fundamentals of the manifest.
- Interact with Word to manipulate Paragraphs and Content Controls.
- Interact with document content through Office JavaScript APIs.

## Applies to

- Office on Windows, Mac, and in a browser.

## Prerequisites

- Microsoft 365 - You can get a [free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) that provides a renewable 90-day Microsoft 365 E5 developer subscription.

## Understand an Office Add-in in Blazor Context

An Office Add-in is a web application that can extend Office with additional functionality for the user. For example, an add-in can add ribbon buttons, a task pane, or a content pane with the functionality you want. Because an Office Add-in is a web application you must provide a web server to host the files.
When building the Office Add-in as a Blazor Webassembly this will allow you to build a .NET Core compliant website that interacts with the OfficeJS APIs. This can lower the bar for existing Office developers used to develop their applications using VBA, VSTO or COM Add-ins.

To work with the sample, clone or download this repo. Then go to the folder containing the sample for the Office application you want to work with.

## Key components

The hello world sample implements the **Manifest** and **Web app** components identified in [Components of an Office Add-in](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins#components-of-an-office-add-in).
What is different in this sample is that the Add-in is provided as a Blazor Webassembly file that can run cross-platform in various browsers supporting WASM (Webassembly).

### Manifest

You only need one manifest file for your add-in. The Blazor Add-in sample contains two manifest files to support two different hosting scenarios.
Using Visual Studio the manifest is swapped out to either run in localhost or on the deployed host (the manifest_live.xml needs to be adjusted to your specific host settings)

- **manifest_live.xml**: This manifest file will load the add-in from a hosted scenario such as Azure Web App.
- **manifest_local.xml**: This manifest file will load the add-in from a local web server that you configure. See the hello world README files for the specific hosts on how to run the local web server.

### Web app

The Blazor WASM App will demo some basic Word functions using Paragraphs and Content Controls including Event Handlers.
The purpose is not so much to show how to interact with Word. This can be adopted from the many Gists available in Script Lab. The purpose of this sample is to show how to build and interact with the Blazor, C# and JavaScript Interop options.

## Details and running the Add-in

For details on running the Add-in in Word see the README in the hello world samples for details and to learn how to run the local web server.

## Copyright

Copyright (c) 2022 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/samples/word-blazor-add-in" />

