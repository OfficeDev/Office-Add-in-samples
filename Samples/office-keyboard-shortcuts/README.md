---
page_type: sample
urlFragment: office-add-in-keyboard-shortcuts
products:
  - office-excel
  - office-word
  - office
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

This sample shows how to create custom keyboard shortcuts for an Office Add-in. Keyboard shortcuts let power users quickly use your add-in's features and give accessibility options to avoid using a mouse. In this sample, the following shortcuts are configured.

- <kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>Up arrow key</kbd>: Opens the add-in's task pane.
- <kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>Down arrow key</kbd>: Hides the add-in's task pane.
- <kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>Q</kbd>: Performs an action that's specific to the current Office host.
  - **Excel**: Cycles through colors in the currently selected cell.
  - **Word**: Adds text to the document.

Keyboard shortcuts can be used to achieve any action within the add-in runtime.

![The sample's task pane displaying a list of the available keyboard shortcuts.](./assets/office-keyboard-shortcuts-overview.png)

## Features

- Add keyboard shortcuts to your Office Add-in.
- Provide users with keyboard shortcuts to invoke any action within the Office Add-in runtime.

## Applies to

- Office on the web
  - Excel
  - Word

    > **Note**: The keyboard shortcut feature is currently being rolled out to Word on the web. If you test the feature in Word on the web at this time, the shortcuts may not work if they're activated from within the add-in's task pane. We recommend to periodically check [Keyboard Shortcuts requirement sets](https://learn.microsoft.com/javascript/api/requirement-sets/common/keyboard-shortcuts-requirement-sets) to find out when the feature is fully supported.

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
| 1.0 | 11-5-2020 | Initial release |
| 1.1 | May 11, 2021 | Removed yo office and modified to be GitHub hosted |
| 2.0 | September 27, 2024 | Added support for Word |
| 2.1 | December 5, 2024 | Updated keyboard shortcuts |

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

## Run the sample

Run this sample with a [unified manifest for Microsoft 365](#run-with-the-unified-manifest-for-microsoft-365) or [add-in only manifest](#run-with-the-add-in-only-manifest). Use one of the following add-in file hosting options.

> [!NOTE]
>
> - Implementing keyboard shortcuts with the unified manifest for Microsoft 365 is currently in public developer preview. This shouldn't be used in production add-ins. We invite you to try it out in test or development environments. For more information, see the [Microsoft 365 app manifest schema reference](https://learn.microsoft.com/microsoft-365/extensibility/schema/?view=m365-app-prev&preserve-view=true).

### Run with the unified manifest for Microsoft 365

#### Use GitHub as the web host

The quickest way to run the sample is to use GitHub as the web host. However, you can't debug or change the source code. The add-in web files are served from this GitHub repository.

1. Download the **office-keyboard-shortcuts.zip** file from this sample to a folder on your computer.
1. Sideload the sample to Excel or Word by following the instructions in [Sideload with the Teams Toolkit CLI (command-line interface)](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-add-in-with-unified-manifest#sideload-with-the-teams-toolkit-cli-command-line-interface).
1. Verify that the add-in loaded successfully. You'll see a **Keyboard shortcuts** button on the **Home** tab of the ribbon.
1. Follow the steps in [Try it out](#try-it-out) to test the sample.
1. To uninstall the add-in, run the following command. Replace *{title ID}* with the add-in's title ID that was generated when you sideloaded the add-in.

    ```console
    teamsapp uninstall --mode title-id --title-id {title ID} --interactive false
    ```

#### Use localhost

If you prefer to host the web server on localhost, follow these steps.

1. Clone or download this repository.
1. From a command prompt, go to the root of the project folder **/samples/office-keyboard-shortcuts**.
1. Run the following commands.

    ```console
    npm install
    npm start
    ```

    This starts the web server on localhost and sideloads the **manifest.json** file.

1. Verify that the add-in loaded successfully. You'll see a **Keyboard shortcuts** button on the **Home** tab of the ribbon.
1. Follow the steps in [Try it out](#try-it-out) to test the sample.
1. To stop the web server and uninstall the add-in, run the following command.

    ```console
    npm stop
    ```

#### Use Microsoft Azure

You can deploy this sample with the unified manifest to Microsoft Azure using the Teams Toolkit extension in Visual Studio Code.

1. In Visual Studio Code, go to the activity bar, then open the Teams Toolkit extension.
1. In the Accounts section of the Teams Toolkit pane, choose **Sign in to Azure**.
1. After you sign in, select a subscription under your account.
1. In the Development section of the Teams Toolkit pane, choose **Provision in the cloud**. Alternatively, open the command palette and choose **Teams: Provision in the cloud**.
1. Choose **Deploy to the cloud**. Alternatively, open the command palette and choose **Teams: Deploy to the cloud**.

Once the sample is successfully deployed, follow these steps.

1. Copy the endpoint of your new Azure deployment. Use one of the following methods.
    - In Visual Studio Code, select **View** > **Output** to open the Output window. Then, copy the endpoint for your new Azure deployment.
    - In the Azure portal, go to the new storage account. Then, choose **Data management** > **Static website** and copy the **Primary endpoint** value.
1. Open the **./webpack.config.js** file.
1. Change the `urlProd` constant to use the endpoint of your Azure deployment.
1. Save your change. Then, run the following command.

    ```console
    npm run build
    ```

    This generates a new **manifest.json** file in the **dist** folder of your project that will load the add-in resources from your storage account.
1. Run the following command.

    ```console
    npm run start:prod
    ```

    //TODO starts and the **manifest.json** file is sideloaded from the **dist** folder.
1. Verify that the add-in loaded successfully. You'll see a **Keyboard shortcuts** button on the **Home** tab of the ribbon.
1. Follow the steps in [Try it out](#try-it-out) to test the sample.
1. To stop the web server and uninstall the add-in, run the following command.

    ```console
    npm run stop:prod
    ```

### Run with the add-in only manifest

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

If you prefer to host the web server on localhost, follow these steps.

1. Clone or download this repository.
1. From a command prompt, run the following commands.

    ```console
    npm install
    npm run start:xml
    ```

    This starts the web server on localhost and sideloads the **manifest-localhost.xml** file.
1. Verify that the add-in loaded successfully. You'll see a **Keyboard shortcuts** button on the **Home** tab of the ribbon.
1. Follow the steps in [Try it out](#try-it-out) to test the sample.
1. To stop the web server and uninstall the add-in, run the following command.

    ```console
    npm run stop:xml
    ```

## Try it out

Once the add-in is loaded, try out its functionality.

1. Press <kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>Up arrow key</kbd> on your keyboard to open the add-in's task pane.

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
