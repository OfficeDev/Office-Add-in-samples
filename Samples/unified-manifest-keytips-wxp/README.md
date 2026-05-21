---
title: Define keytips for ribbon controls with the unified manifest
page_type: sample
urlFragment: unified-manifest-keytips-wxp
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
description: "This sample shows how to define keytips (keyboard shortcuts shown on Alt) for ribbon tabs, groups, buttons, menus, and menu items in an Office Add-in that uses the unified manifest for Microsoft 365."
---

# Define keytips for ribbon controls with the unified manifest

This sample shows how to define **keytips** for custom ribbon controls in an Office Add-in that uses the unified manifest for Microsoft 365. Keytips are the letters that appear next to ribbon controls when the user presses **Alt**, enabling keyboard-only navigation of the ribbon.

The sample defines a custom ribbon tab named **Contoso Keytips** in Excel, Word, and PowerPoint, with two groups of buttons and a menu. Each tab, group, button, menu, and menu item has a `keytip` value assigned in the manifest.

## Applies to

- Excel, Word, and PowerPoint on Windows, Mac and the web.

## Prerequisites

- [Node.js](https://nodejs.org/) (the latest LTS version).
- The latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). To install these tools globally, run the following command via the command prompt:

    ```console
    npm install -g yo generator-office
    ```

- Office connected to a Microsoft 365 subscription (including Office on the web). Get a [free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) that provides a renewable 90-day Microsoft 365 E5 developer subscription.

## Version history

| Version  | Date        | Comments        |
|----------|-------------|-----------------|
| 1.0      | 05-20-2026  | Initial release |

## Solution

| Solution | Author(s) |
|----------|-----------|
| Define keytips for ribbon controls with the unified manifest | Microsoft |

## Run the sample

This sample is configured to run with the unified manifest for Microsoft 365. Two manifest files are provided:

- **manifest.json** — references the add-in's web files hosted on GitHub Pages at `https://officedev.github.io/Office-Add-in-samples/Samples/unified-manifest-keytips-wxp/`.
- **manifest-localhost.json** — references the add-in's web files served from `https://localhost:3000` for local development.

### Run the sample on localhost

1. Clone or download this repo to your computer.
1. In a console or terminal, navigate to the **Samples/unified-manifest-keytips-wxp** folder.
1. Install dependencies.

    ```console
    npm install
    ```

1. Start the local web server and sideload the add-in. This uses **manifest.json** by default; to use the localhost-only manifest, replace `manifest.json` with `manifest-localhost.json` in the `start` script or run the command directly.

    ```console
    npm start
    ```

    Or to target a specific application:

    ```console
    npm run start:desktop:excel
    npm run start:desktop:word
    npm run start:desktop:powerpoint
    ```

1. When the add-in loads, you'll see a new **Contoso Keytips** tab on the ribbon.

### Stop the local web server

```console
npm stop
```

## Try the sample

1. Open Excel, Word, or PowerPoint with the add-in sideloaded.
1. Press the **Alt** key. Keytip letters appear over each ribbon tab.
1. Press **C** then **K** to navigate to the **Contoso Keytips** tab. (`CK` is the keytip assigned to the tab in this sample.)
1. With the tab active, press **Alt** again (if keytips have disappeared) to view the keytips for the controls on the tab:
    - **Buttons 1** group: `R` (Button 1), `O` (Button 2), `Y` (Button 3), `M` (Menu).
    - Open the menu with `M`, then choose: `P`, `C`, `M`, `L`, `B`, or `G` to invoke a menu item.
    - **Buttons 2** group: `G` (Button 4), `B` (Button 5), `P` (Button 6).
1. Each button and menu item runs a function defined in **src/commands/commands.ts** that displays a notification message.

## Key parts of the sample

The keytip values are defined in **manifest.json** (and **manifest-localhost.json**) on each ribbon element:

```json
{
    "id": "ContosoKeytipsTab",
    "label": "Contoso Keytips",
    "keytip": "CK",
    "groups": [
        {
            "id": "KeytipGroup1",
            "label": "Buttons 1",
            "controls": [
                {
                    "id": "Btn1",
                    "type": "button",
                    "label": "Button 1",
                    "actionId": "btn1Action",
                    "keytip": "R"
                }
            ]
        }
    ]
}
```

The `keytip` property is supported on the following elements in the unified manifest:

- Custom ribbon `tabs`
- `groups` within a tab
- `controls` (buttons and menus)
- `items` (menu items inside a menu)

## Use the sample in your own project

To add keytips to ribbon controls in your own add-in:

1. Open your **manifest.json** file.
1. Add a `"keytip"` property with a one-to-three character string on any custom ribbon `tab`, `group`, `control`, or menu `item`.
1. Use unique keytip letters within each parent scope so each control can be selected unambiguously.

## Additional resources

- [Office Add-ins manifest](https://learn.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
- [Unified manifest for Microsoft 365](https://learn.microsoft.com/office/dev/add-ins/develop/unified-manifest-overview)
- [Add-in commands](https://learn.microsoft.com/office/dev/add-ins/design/add-in-commands)
- [Keyboard shortcuts in Office Add-ins](https://learn.microsoft.com/office/dev/add-ins/design/keyboard-shortcuts)

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2026 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/unified-manifest-keytips-wxp" />
