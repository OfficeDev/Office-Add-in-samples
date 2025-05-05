---
page_type: sample
urlFragment: office-word-add-in-tutorial
products:
  - m365
  - office
  - office-word
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 09/20/2023 4:00:00 PM
description: "A completed version of the step-by-step Word tutorial hosted on learn.microsoft.com."
---

# Word Tutorial - Completed

## Summary

This sample is the result of completing the [Tutorial: Create a Word task pane add-in](https://learn.microsoft.com/office/dev/add-ins/tutorials/word-tutorial). It was constructed with the [Yeoman generator for Office Add-ins](https://learn.microsoft.com/office/dev/add-ins/develop/yeoman-generator-overview).

The tutorial gives step-by-step instructions on how to add functionality alongside explanations as to why code is being added. Use this sample if you want to explore and try the completed code, or if you need to debug any issues you encountered while following the tutorial.

## Features

This sample demonstrates the basics of working with a document in Word. The add-in inserts and replaces text ranges, paragraphs, images, HTML, tables, and content controls. The add-in also shows how to format text and how to manage content with content controls.

## Applies to

- Word on Windows
- Word on Mac
- Word on the web

## Prerequisites

- Office connected to a Microsoft 365 subscription (including Office on the web).
- [Node.js](https://nodejs.org/) version 16 or greater.
- [npm](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm) version 8 or greater.

## Solution

| Solution | Author(s) |
|----------|-----------|
| Learn the basics of Word add-ins | Microsoft |

## Version history

| Version  | Date | Comments |
|----------|------|----------|
| 1.0 | 9-20-2023 | Initial release |

## Run the sample

1. Fork and download this repo.

1. Go to the **Samples/tutorials/word-tutorial/My Office Add-in** folder via the command line.

1. Run `npm install`.

1. Run `npm run build`.

1. Start the local web server and sideload your add-in.

    - To test your add-in in Word, run the following command in the root directory of your project. This starts the local web server (if it's not already running) and opens Word with your add-in loaded.

      - Windows: `npm start`
      - Mac: `npm run dev-server`

    - To test your add-in in Word on the web, run the following command in the root directory of the `My Office Add-in` project. When you run this command, the local web server starts. Replace "{url}" with the URL of a Word document on your OneDrive or a SharePoint library to which you have permissions.

      ```command line
      npm run start:web -- --document {url}
      ```

      The following are examples.

      - `npm run start:web -- --document https://contoso.sharepoint.com/:t:/g/EZGxP7ksiE5DuxvY638G798BpuhwluxCMfF1WZQj3VYhYQ?e=F4QM1R`
      - `npm run start:web -- --document https://1drv.ms/x/s!jkcH7spkM4EGgcZUgqthk4IK3NOypVw?e=Z6G1qp`
      - `npm run start:web -- --document https://contoso-my.sharepoint-df.com/:t:/p/user/EQda453DNTpFnl1bFPhOVR0BwlrzetbXvnaRYii2lDr_oQ?e=RSccmNP`

      > NOTE: If you're developing on a Mac, enclose the {url} in single quotation marks. Don't do this on Windows.

      If your add-in doesn't sideload in the document, manually sideload it by following the instructions in [Manually sideload add-ins to Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).

1. If the add-in task pane isn't already open in Word, go to the Home tab and choose the **Show Taskpane** button in the ribbon to open it.

1. Use the buttons in the task pane to interact with the document through your add-in. For details on the expected behavior of each button, see the "Test the add-in" sections of the [Tutorial: Create a Word task pane add-in](https://learn.microsoft.com/office/dev/add-ins/tutorials/word-tutorial).

## See also

The version of this sample that you create step-by-step is found in the article [Tutorial: Create a Word task pane add-in](https://learn.microsoft.com/office/dev/add-ins/tutorials/word-tutorial).

## Copyright

Copyright (c) 2023 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/office-word-add-in-tutorial" />