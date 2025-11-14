---
page_type: sample
urlFragment: office-outlook-add-in-tutorial
products:
  - m365
  - office
  - office-outlook
languages:
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: 09/12/2023 4:00:00 PM
description: "A completed version of the step-by-step Outlook tutorial hosted on learn.microsoft.com."
---

# Outlook Tutorial - Completed

## Summary

This sample is the result of completing the [Tutorial: Build a message compose Outlook add-in](https://learn.microsoft.com/office/dev/add-ins/tutorials/outlook-tutorial). It was constructed with the [Yeoman generator for Office Add-ins](https://learn.microsoft.com/office/dev/add-ins/develop/yeoman-generator-overview).

The tutorial gives step-by-step instructions on how to add functionality alongside explanations as to why code is being added. Use this sample if you want to explore and try the completed code, or if you need to debug any issues you encountered while following the tutorial.

## Features

This sample demonstrates the basics of working with a compose message in Outlook. The functions collect information from the user, fetch data from an external service, implement a function command, and implement a task pane that inserts content into the body of a message. The sample also shows how to use a dialog box.

## Applies to

- Outlook on Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic)
- Outlook on Mac
- Outlook on the web

## Prerequisites

- Office connected to a Microsoft 365 subscription (including Office on the web).
- [Node.js](https://nodejs.org/) (latest recommended version).
- [npm](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm) version 8 or greater.
- [Showdown](https://github.com/showdownjs/showdown).
- [URI.js](https://github.com/medialize/URI.js).
- [jQuery](https://jquery.com/).
- [Set up GitHub gists](https://learn.microsoft.com/office/dev/add-ins/tutorials/outlook-tutorial#setup) on your account.

## Choose a manifest type

By default, the sample uses an add-in only manifest. However, you can switch the project between the add-in only manifest and the unified manifest for Microsoft 365. For more information about the differences between them, see [Office Add-ins manifest](https://learn.microsoft.com/office/dev/add-ins/develop/add-in-manifests). To continue with the add-in only manifest, skip ahead to the [Run the sample](#run-the-sample) section.

> [!NOTE]
> To run the sample in Outlook on Mac, use the [add-in only manifest](#run-with-the-add-in-only-manifest). For more information on manifests and their supported platforms, see [Office Add-in manifest](https://learn.microsoft.com/office/dev/add-ins/develop/add-in-manifests).

### To switch to the unified manifest for Microsoft 365

Copy all the files from the **manifest-configurations/unified** subfolder to the sample's root folder, replacing any existing files that have the same names. We recommend that you delete the **manifest.xml** and **manifest-localhost.xml** files from the root folder, so only files needed for the unified manifest are present. Then, [run the sample](#run-the-sample).

### To switch back to the add-in only manifest

To switch back to the add-in only manifest, copy the files from the **manifest-configurations/add-in-only** subfolder to the sample's root folder. We recommend that you delete the **manifest.json** file from the root folder.

## Run the sample

1. Fork and download this repository.

1. Go to the **Samples/tutorials/outlook-tutorial/Git the gist** folder via the command line.

1. Run `npm install`.

1. Run the following command to start the local web server and sideload your add-in in Outlook.

    ```console
    npm start
    ```

    > If your add-in doesn't sideload in Outlook, manually sideload it by following the instructions in [Sideload Outlook add-ins for testing](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing#sideload-manually).

1. In Outlook, compose a new message.

1. In the message window, choose the **Insert default gist** button in the ribbon. This opens a dialog where you add your GitHub username and select the default gist.

    > In classic Outlook on Windows, you may need to close and reopen the new message window to pick up the latest settings from the dialog.

1. In the message window, choose the **Insert gist** button in ribbon. This opens a task pane where you select the GitHub gist you want to insert into the message body.

## See also

The version of this sample that you create step-by-step is found in the article [Tutorial: Build a message compose Outlook add-in](https://learn.microsoft.com/office/dev/add-ins/tutorials/outlook-tutorial).

## Solution

| Solution | Author(s) |
|----------|-----------|
| Learn the basics of Outlook add-ins | Microsoft |

## Version history

| Version  | Date | Comments |
|----------|------|----------|
| 1.0 | 9-12-2023 | Initial release |
| 1.1 | 11-14-2025 | Added support for the unified manifest for Microsoft 365 |

## Copyright

Copyright (c) 2023 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/office-outlook-add-in-tutorial" />
