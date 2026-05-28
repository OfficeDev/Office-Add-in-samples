---
title: Define KeyTips for the ribbon controls of an Office Add-in
page_type: sample
urlFragment: office-keytips
products:
  - office-excel
  - office-word
  - office-powerpoint
  - office
  - m365
languages:
  - typescript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: "05/20/2026 10:00:00 AM"
description: "This sample shows how to define KeyTips for ribbon tabs, buttons, and menus of an Office Add-in that uses the unified manifest for Microsoft 365."
---

# Define KeyTips for the ribbon controls of an Office Add-in

This sample shows how to define [KeyTips](https://support.microsoft.com/office/954cd3f7-2f77-4983-978d-c09b20e31f0e) for custom ribbon controls of an Office Add-in that uses the unified manifest for Microsoft 365. KeyTips, also known as sequential key shortcuts or access keys, appear around the ribbon controls when <kbd>Alt</kbd> is pressed. They enable keyboard-based navigation of the ribbon, making the user experience more accessible and efficient.

In this sample, a custom ribbon tab named **Contoso KeyTips** contains buttons and menus to select a color option for the following host-specific actions.

- **Excel**: Changes the color of the selected cell.
-  **PowerPoint**: Inserts a text box with a fill.
-  **Word**: Inserts colored text.

Each tab, button, and menu is assigned a custom KeyTip in the manifest.

To learn more about custom KeyTips for Office Add-ins, see [Add custom KeyTips to your Office Add-ins](https://learn.microsoft.com/office/dev/add-ins/design/add-custom-key-tips).

## Applies to

Custom KeyTips are supported in Excel, PowerPoint, and Word on the following platforms.
- Web
- Windows (Version 2603 (Build 19822.20000) and later)
- Mac (Version 16.107 (Build 26030819) and later)

## Prerequisites

- [Node.js](https://nodejs.org/) (the latest LTS version).
- Office connected to a Microsoft 365 subscription (including Office on the web). Get a [free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) that provides a renewable 90-day Microsoft 365 E5 developer subscription.

## Version history

| Version  | Date        | Comments        |
|----------|-------------|-----------------|
| 1.0      | 05-28-2026  | Initial release |

## Solution

| Solution | Author(s) |
|----------|-----------|
| Define KeyTips for the ribbon controls of an Office Add-in | Microsoft |

## Run the sample

Run this sample in Excel, Word, or PowerPoint. The add-in web files are served from `https://localhost:3000` on your computer.

1. Clone or download this repository to your computer.
1. In a console or terminal, go to the root of the project folder **Samples/office-keytips**.
1. Run the following command to install dependencies.

    ```console
    npm install
    ```

1. Run the following command to start the local web server and sideload the add-in in Excel.

    ```console
    npm start
    ```

    If you want to run the add-in on Word or PowerPoint, run one of the following commands instead.

    ```console
    npm run start:desktop:word
    npm run start:desktop:powerpoint
    ```

    Your preferred application opens and the add-in is sideloaded. When the add-in is sideloaded, the **Contoso KeyTips** tab appears on the ribbon.
    
    > **Tip**: To test the add-in in Office on the web, run the applicable command to sideload the add-in in your preferred desktop application. Once sideloaded, the add-in also appears on the web client.

1. Follow the steps in [Try it out](#try-the-sample) to test the sample.

1. To stop the web server, run the following command.

```console
npm stop
    

## Try the sample

1. In Excel, PowerPoint, or Word, press <kbd>Alt</kbd>.

    KeyTips appear for each ribbon tab.
1. Press <kbd>C</kbd> then <kbd>K</kbd> to open the **Contoso KeyTips** tab.
1. With the **Contoso KeyTips** tab active, if the KeyTips have disappeared from the ribbon, press <kbd>Alt</kbd>+<kbd>CK</kbd> again. The KeyTips appear for the controls in the tab. Select the key for the color you want to apply to the host action.
    - **Red**: <kbd>R</kbd>
    - **Orange**: <kbd>O</kbd>
    - **Yellow**: <kbd>Y</kbd>
    - **More colors**: <kbd>M</kbd>
    - **Green**: <kbd>G</kbd>
    - **Blue**: <kbd>B</kbd>
    - **Purple**: <kbd>P</kbd>   

## Key parts of the sample

Custom KeyTips are defined in the ["keytip"](https://learn.microsoft.com/microsoft-365/extensibility/schema/extension-ribbons-array-tabs-item#keytip) property of each ribbon control in the manifest. The `"keytip"` property is supported for the following controls.

- [Built-in Office tabs and custom tabs](https://learn.microsoft.com/microsoft-365/extensibility/schema/extension-ribbons-array-tabs-item)
- [Buttons and menus](https://learn.microsoft.com/microsoft-365/extensibility/schema/extension-common-custom-group-controls-item)

The following sample defines a KeyTip for the add-in's tab and button.

```json
{
    "id": "ContosoKeyTipsTab",
    "label": "Contoso KeyTips",
    "keytip": "CK",
    "groups": [
        {
            "id": "KeyTipGroup1",
            "label": "Colors Group 1",
            ...
            "controls": [
                {
                    "id": "Btn1",
                    "type": "button",
                    "label": "Red",
                    ...
                    "actionId": "btn1Action",
                    "keytip": "R"
                }
            ]
        }
    ]
}
```

## Additional resources

- [Office Add-ins manifest](https://learn.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
- [Unified manifest for Microsoft 365](https://learn.microsoft.com/office/dev/add-ins/develop/unified-manifest-overview)
- [Add-in commands](https://learn.microsoft.com/office/dev/add-ins/design/add-in-commands)
- [Add custom KeyTips to your Office Add-ins](https://learn.microsoft.com/office/dev/add-ins/design/add-custom-key-tips)

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2026 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/office-keytips" />
