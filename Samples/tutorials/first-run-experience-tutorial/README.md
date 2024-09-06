---
page_type: sample
urlFragment: office-add-in-first-run-experience-tutorial
products:
  - office
languages:
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: 09/05/2024 4:00:00 PM
description: "A completed version of the step-by-step first-run experience tutorial hosted on learn.microsoft.com."
---

# Office Add-in First-run Experience Tutorial - Completed

## Summary

This sample is the result of completing the [Build an Office Add-in with a basic first-run experience](https://learn.microsoft.com/office/dev/add-ins/tutorials/first-run-experience-tutorial) tutorial. It's built from the [Yeoman generator for Office Add-ins](https://learn.microsoft.com/office/dev/add-ins/develop/yeoman-generator-overview).

The tutorial gives step-by-step instructions on how to add functionality alongside explanations as to why code is being added. Use this sample if you want to explore and try the completed code, or if you need to debug any issues you encountered while following the tutorial.

## Features

This sample demonstrates the basics of implementing a first-run experience (FRE) in Excel. The pattern can be applied to other Office applications where Office Web Add-ins are supported.

## Applies to

- Excel on Windows
- Excel on Mac
- Excel on the web

## Prerequisites

- Office connected to a Microsoft 365 subscription (including Office on the web).
- [Node.js](https://nodejs.org/) version 16 or greater.
- [npm](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm) version 8 or greater.

## Solution

Solution | Authors
---------|----------
Learn the basics of implementing a first-run experience | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.0 | 9-5-2024 | Initial release

## Run the sample

1. Fork and download this repo.

1. Go to the **Samples/tutorials/first-run-experience-tutorial/My Office Add-in** folder via the command line.

1. Run `npm install`.

1. Run `npm run build`.

1. Start the local web server and sideload your add-in.

    - To test your add-in in Excel, run the following command in the root directory of your project. This starts the local web server (if it's not already running) and opens Excel with your add-in loaded.

      - Windows: `npm start`
      - Mac: `npm run dev-server`

    - To test your add-in in Excel on the web, run the following command in the root directory of the `My Office Add-in` project. When you run this command, the local web server starts. Replace "{url}" with the URL of an Excel document on your OneDrive or a SharePoint library to which you have permissions.

      ```command line
      npm run start:web -- --document {url}
      ```

      The following are examples.

      - `npm run start:web -- --document https://contoso.sharepoint.com/:t:/g/EZGxP7ksiE5DuxvY638G798BpuhwluxCMfF1WZQj3VYhYQ?e=F4QM1R`
      - `npm run start:web -- --document https://1drv.ms/x/s!jkcH7spkM4EGgcZUgqthk4IK3NOypVw?e=Z6G1qp`
      - `npm run start:web -- --document https://contoso-my.sharepoint-df.com/:t:/p/user/EQda453DNTpFnl1bFPhOVR0BwlrzetbXvnaRYii2lDr_oQ?e=RSccmNP`

      > NOTE: If you are developing on a Mac, enclose the {url} in single quotation marks. Do not do this on Windows.

      If your add-in doesn't sideload in the document, manually sideload it by following the instructions in [Manually sideload add-ins to Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).

1. If the add-in task pane isn't already open in Excel, go to the Home tab and choose the **Show Taskpane** button in the ribbon to open it.

1. Verify that the list of Office Add-ins features is displayed. This indicates that the first-run experience was shown.

1. Close the task pane then reopen it. Verify that the task pane no longer displays the list of of features.

1. When you want to stop the local web server and uninstall the add-in, follow the applicable instructions:

    - To stop the server, run `npm stop`. This also uninstalls the add-in if it was automatically sideloaded.
    - If you manually sideloaded the add-in, see [Remove a sideloaded add-in](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing#remove-a-sideloaded-add-in).

## See also

The version of this sample that you create step-by-step is found in the article [Build an Office Add-in with a basic first-run experience](https://learn.microsoft.com/office/dev/add-ins/tutorials/first-run-experience-tutorial).

## Copyright

Copyright (c) 2024 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/office-add-in-first-run-experience-tutorial" />