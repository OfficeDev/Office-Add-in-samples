---
page_type: sample
urlFragment: office-add-in-keyboard-shortcuts
products:
  - office-excel
  - office-word
  - office
  - m365
languages:
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: "11/5/2020 10:00:00 AM"
description: "This sample shows how to add keyboard shortcuts to your Office Add-in."
---

# Use keyboard shortcuts for Office Add-in actions

## Summary

This sample shows how to create custom keyboard shortcuts for an Office Add-in. Keyboard shortcuts let power users quickly use your add-in's features and give accessibility options to avoid using a mouse. In this sample, the following shortcuts are configured for each supported platform.

| Action | Windows | Mac | Web |
| ----- | ----- | ----- | ----- |
| Open the add-in's task pane | <kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>Up arrow key</kbd> | <kbd>Command</kbd>+<kbd>Shift</kbd>+<kbd>Up arrow key</kbd> | <kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>1</kbd> |
| Hide the add-in's task pane | <kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>Down arrow key</kbd> | <kbd>Command</kbd>+<kbd>Shift</kbd>+<kbd>Down arrow key</kbd> | <kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>2</kbd> |
| Perform an action that's specific to the current Office host<br>- **Excel**: Cycle through colors in the currently selected cell<br>- **Word**: Add text to the document | <kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>Q</kbd> | <kbd>Command</kbd>+<kbd>Shift</kbd>+<kbd>Q</kbd> | <kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>3</kbd> |

Keyboard shortcuts can be used to achieve any action within the add-in runtime.

![The sample's task pane displaying a list of the available keyboard shortcuts.](./assets/office-keyboard-shortcuts-overview.png)

## Features

- Add keyboard shortcuts to your Office Add-in.
- Provide users with keyboard shortcuts to invoke any action within the Office Add-in runtime.

## Applies to

- Office on the web
  - Excel
  - Word
- Office on Windows
  - Excel: Version 2111 (Build 14701.10000)
  - Word: Version 2408 (Build 17928.20114)
- Office on Mac
  - Excel: Version 16.55 (21111400)
  - Word: Version 16.88 (24081116)

## Prerequisites

- Microsoft 365

## Solution

| Solution | Authors |
| -------- | --------- |
| Use keyboard shortcuts for Office Add-in actions | Microsoft |

## Version history

| Version | Date | Comments |
| ------- | ---- | -------- |
| 1.0 | November 5, 2020 | Initial release |
| 1.1 | May 11, 2021 | Removed yo office and modified to be GitHub hosted |
| 2.0 | September 27, 2024 | Added support for Word |
| 2.1 | December 5, 2024 | Updated keyboard shortcuts |
| 2.2 | July 29, 2025 | Added support for the unified manifest for Microsoft 365 and updated support in Word on the web |

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

## Choose a manifest type

By default, the sample uses an add-in only manifest. However, you can switch the project between the add-in only manifest and the unified manifest for Microsoft 365. For more information about the differences between them, see [Office Add-ins manifest](https://learn.microsoft.com/office/dev/add-ins/develop/add-in-manifests). To continue with the add-in only manifest, skip ahead to the [Run the sample](#run-the-sample) section.

### To switch to the unified manifest for Microsoft 365

Copy all the files from the **manifest-configurations/unified** subfolder to the sample's root folder, replacing any existing files that have the same names. We recommend that you delete the **manifest.xml** and **manifest-localhost.xml** files from the root folder, so only files needed for the unified manifest are present. Then, [run the sample](#run-the-sample).

### To switch back to the add-in only manifest

To switch back to the add-in only manifest, copy the files from the **manifest-configurations/add-in-only** subfolder to the sample's root folder. We recommend that you delete the **manifest.json** file from the root folder.

## Run the sample

Run this sample with the [add-in only manifest](#run-with-the-add-in-only-manifest) or [unified manifest for Microsoft 365](#run-with-the-unified-manifest-for-microsoft-365). Use one of the following add-in file hosting options.

### Run with the unified manifest for Microsoft 365

1. Clone or download this repository.
1. From a command prompt, go to the root of the project folder **/samples/office-keyboard-shortcuts**.
1. Run `npm install`.
1. Run the applicable command for your Office application.

  - **Excel**: `npm run start:desktop:excel`
  - **Word**: `npm run start:desktop:word`

    This starts the web server on localhost and sideloads the **manifest.json** file.

1. Verify that the add-in loaded successfully. You'll see a **Keyboard shortcuts** button on the **Home** tab of the ribbon.
1. Follow the steps in [Try it out](#try-it-out) to test the sample.
1. To stop the web server and uninstall the add-in, run the following command.

    ```console
    npm stop
    ```

### Run with the add-in only manifest

To run the sample using the add-in only manifest, you can choose to host the web server on [GitHub](#use-github-as-the-web-host) or (localhost)[#use-localhost].

#### Use GitHub as the web host

The quickest way to run the sample is to use GitHub as the web host. However, you can't debug or change the source code. The add-in web files are served from this GitHub repository.

1. Download the **manifest.xml** file from this sample to a folder on your computer.
1. Sideload the manifest file in Excel or Word. The sideloading process varies depending on your platform.

    - **Office on the web**: [Manually sideload an add-in to Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing#manually-sideload-an-add-in-to-office-on-the-web)
    - **Office on Windows**: [Sideload Office Add-ins for testing from a network share](https://learn.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)
    - **Office on Mac**: [Sideload Office Add-ins on Mac for testing](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)
1. Verify that the add-in loaded successfully. You'll see a **Keyboard shortcuts** button on the **Home** tab of the ribbon.
1. Follow the steps in [Try it out](#try-it-out) to test the sample.
1. To uninstall the add-in, follow the instructions for the applicable platform.

    - [Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing#remove-a-sideloaded-add-in)
    - [Office on Windows](https://learn.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins#remove-a-sideloaded-add-in)
    - [Office on Mac](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac#remove-a-sideloaded-add-in)

#### Use localhost

To host the web server on localhost:

1. Clone or download this repository.
1. From a command prompt, go to the root of the project folder **/samples/office-keyboard-shortcuts**.
1. Run `npm install`.
1. Run the applicable command for your Office application.

  - **Excel**: `npm run start:desktop:excel`
  - **Word**: `npm run start:desktop:word`

    This starts the web server on localhost and sideloads the **manifest-localhost.xml** file.

1. Verify that the add-in loaded successfully. You'll see a **Keyboard shortcuts** button on the **Home** tab of the ribbon.
1. Follow the steps in [Try it out](#try-it-out) to test the sample.
1. To stop the web server and uninstall the add-in, run the following command.

    ```console
    npm stop
    ```

## Try it out

Once the add-in is loaded, try out its functionality.

1. Press the [applicable shortcut](#summary) on your keyboard to open the add-in's task pane.

  > [!NOTE]
  > If the keyboard shortcut is already in use in Excel or Word, a dialog will be shown so that you can select which action you'd like to map to the shortcut. Once you select an action, you can change your preference by invoking the **Reset Office Add-in Shortcut Preferences** command from the search field.
  >
  > ![The Reset Office Add-in Shortcut Preferences option in Excel.](./assets/office-keyboard-shortcuts-reset.png)

1. Try the other available shortcuts shown in the task pane.

## Key parts of this sample

The custom keyboard shortcuts implemented in this sample rely on the following components.

- The add-in manifest is configured to use a shared runtime. For guidance on how to implement a shared runtime in your add-in, see [Configure your Office Add-in to use a shared runtime](https://learn.microsoft.com/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime).
- The method to define an add-in's actions and their keyboard shortcuts differs depending on the type of manifest your add-in uses.
  - **Unified manifest for Microsoft 365**: The actions and keyboard shortcuts are defined in the ["extensions.keyboardShortcuts.shortcuts"](https://learn.microsoft.com/microsoft-365/extensibility/schema/extension-shortcut?view=m365-app-prev&preserve-view=true) array of the **manifest.json** file.
  - **Add-in only manifest**: The actions and keyboard shortcuts are defined in a shortcuts JSON file (**shortcuts.json**). For guidance on how to construct the JSON file, see the [JSON file schema](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).
- The custom actions defined are then mapped to their specific JavaScript functions (**taskpane.js**) using the [Office.actions.associate](https://learn.microsoft.com/javascript/api/office/office.actions#office-office-actions-associate-member(1)) method.

To learn more about each component, see [Add custom keyboard shortcuts to your Office Add-ins](https://learn.microsoft.com/office/dev/add-ins/design/keyboard-shortcuts).

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2020 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/office-keyboard-shortcuts" />
