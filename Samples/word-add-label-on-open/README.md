---
page_type: sample
urlFragment: word-add-label-on-open
products:
  - office-word
  - office
  - m365
languages:
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: 06/30/2025 4:00:00 PM
description: "Shows how to configure a Word add-in to activate when a document opens."
---

# Automatically add labels with an add-in when a Word document opens

## Summary

This sample shows how to configure an add-in to automatically run when a Word document opens. It adds a header to indicate the content's sensitivity.

## Description

The add-in acts when the `OnDocumentOpened` event occurs. The `changeHeader` function is a JavaScript event handler for this event. It adds either a "Public" header to new documents or a "Highly Confidential" header to existing documents that already have content. Some of the functionality is duplicated in the task pane to allow for manual changes.

This sample is designed for Word, but the event-based activation parts will also work for Excel and PowerPoint.

### Event-based activation deployment limitations

Event-based add-ins work only when deployed by an administrator. If users install them directly from AppSource or the Office Store, they will not automatically launch. For the purposes of this sample, sideloading the manifest in Word on the web should be sufficient to explore the functionality.  To perform an admin deployment, upload the manifest to the Microsoft 365 admin center by taking the following actions.

1. In the admin portal, expand the **Settings** section in the navigation pane then select **Integrated apps**.
1. On the **Integrated apps** page, choose the **Upload custom apps** action.

For more information about how to deploy an add-in, please refer to [Deploy and publish Office Add-ins in the Microsoft 365 admin center](https://learn.microsoft.com/microsoft-365/admin/manage/office-addins).

## Applies to

- Word on Windows
- Word on Mac
- Word on the web

## Prerequisites

- Office connected to a Microsoft 365 subscription (including Office on the web).
- [Node.js](https://nodejs.org/) version 16 or greater.
- [npm](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm) version 8 or greater.

## Solution

| Solution | Authors |
|----------|-----------|
| How to configure a Word add-in to activate when a document opens. | Microsoft |

## Version history

| Version  | Date | Comments |
|----------|------|----------|
| 1.0 | 06-30-2025 | Initial release |

## Run the sample

1. Clone or download this repo.

1. Go to the **Samples\word-add-label-on-open** folder via the command line.

1. Run `npm install`.

1. Run `npm run build`.

1. Run `npm start` to launch the web server. **Ignore the Word document that is opened**.

1. Manually sideload your add-in in Word on the web by following the guidance at [Sideload Office Add-ins to Office on the web](../testing/sideload-office-add-ins-for-testing.md#manually-sideload-an-add-in-to-office-on-the-web)

## Try it out

1. Try opening both new and existing Word documents. Headers should automatically be added when they open.

## Make it yours

The following are a few suggestions for how you could tailor this to your scenario.

- Add more complex logic to categorize the headers based on the content of the file.
- Apply the `OnDocumentOpened` event logic to an Excel or PowerPoint add-in.

## Related content

- [Implement event-based activation in Excel, PowerPoint, and Word add-ins](https://learn.microsoft.com/office/dev/add-ins/develop/wxp-event-based-activation.md)
- [Word add-ins documentation](https://learn.microsoft.com/office/dev/add-ins/word/)

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2025 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/word-add-in-label-on-open" />