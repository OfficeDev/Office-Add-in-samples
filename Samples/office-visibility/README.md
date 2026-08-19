---
title: Show or hide the ribbon controls of an Office Add-in
page_type: sample
urlFragment: office-visibility
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
  createdDate: "08/10/2026 10:00:00 AM"
description: "This sample shows how to set and change the visibility of groups and controls on a custom Office Add-in ribbon tab."
---

# Show or hide controls on a custom ribbon tab

This sample shows how to set the initial visibility of groups and controls on a custom ribbon tab and how to change their visibility at runtime. It uses the unified manifest for Microsoft 365, a shared runtime, and `Office.ribbon.requestUpdate()`.

The **Contoso Visibility** tab contains the following UI.

- A **Task pane** group that always remains visible so the sample controls stay accessible.
- A **Primary commands** group that's initially visible.
- **Button 1**, which is initially visible.
- **Button 2**, which is initially hidden.
- A **Sample menu** whose visibility can be changed at runtime.
- A **Secondary commands** group that's initially hidden.

The task pane lets you show or hide any sample button or menu, show or hide entire groups, update several elements in one request, and restore the manifest defaults.

To learn more about managing the visibility of controls on a custom tab, see [Show or hide add-in commands on a custom tab](https://learn.microsoft.com/office/dev/add-ins/design/show-hide-controls-custom-tab).

## Applies to

This sample is supported in Excel, PowerPoint, and Word on the following platforms.

- Office on the web
- Office on Windows: Version 2606 (Build 16.0.20112.15190) and later
- Office on Mac: Version 16.109.1 (Build 260512.115) and later

For more information, see [Ribbon API 1.3 requirement set](https://learn.microsoft.com/javascript/api/requirement-sets/common/ribbon-api-requirement-sets).

## Prerequisites

- [Node.js](https://nodejs.org/) (the latest LTS version).
- Office connected to a Microsoft 365 subscription. Get a [free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) that provides a renewable 90-day Microsoft 365 E5 developer subscription.
- A [supported Office version](#applies-to).

## Version history

| Version | Date | Comments |
| ------- | ---- | -------- |
| 1.0 | 08-10-2026 | Initial release |

## Solution

| Solution | Author(s) |
| -------- | --------- |
| Show or hide controls on a custom ribbon tab | Microsoft |

## Run the sample

The add-in web files are served from `https://localhost:3000` on your computer.

1. Clone or download this repository.
1. In a console or terminal, go to the root of the project folder **Samples/office-visibility**.
1. Run the following command to install the dependencies.

    ```console
    npm install
    ```

1. Run the following command to start the local web server and sideload the add-in in Excel.

    ```console
    npm start
    ```

    To use another desktop application, run the applicable command instead.

    ```console
    npm run start:desktop:word
    npm run start:desktop:powerpoint
    ```

    To test in Office on the web, run the applicable command to sideload the add-in in your preferred desktop application. Once sideloaded, the add-in also appears on the web client.
1. When the Office application opens, select the **Contoso Visibility** tab and then select **Show task pane**.
1. Follow the steps in [Try the sample](#try-the-sample).
1. To stop the web server, run the following command.

    ```console
    npm stop
    ```

## Try the sample

1. On the **Contoso Visibility** tab, verify the initial ribbon state.
    - The **Taskpane** and **Primary commands** groups are visible.
    - **Button 1** is visible.
    - **Button 2** isn't visible.
    - The **Secondary commands** group isn't visible.
1. In the task pane, under **Control visibility**, select **Button 2**, and then select **Show control**.

    The button appears in the **Primary commands** group.
1. Select **Sample menu**, and then select **Hide control**.

    The menu is removed from the ribbon.
1. Under **Group visibility**, select **Secondary commands**, and then select **Show group**.

    The group and its two buttons appear on the ribbon.
1. Under **Control visibility**, select **Secondary button one**, and then select **Hide control**.

    **Secondary button one** disappears while **Secondary button two** remains visible.
1. Under **Batch operations**, try **Hide all sample content**, **Show all sample content**, and **Restore manifest defaults**. Each operation updates the ribbon with one call to `Office.ribbon.requestUpdate()`.

## Key parts of the sample

### Set initial visibility in the manifest

The `visible` property in the **manifest.json** file sets the initial visibility of a group, button, or menu in a custom tab. The default value is `true`.

The following example initially hides a button.

```json
{
  "id": "Button2",
  "type": "button",
  "label": "Button 2",
  "actionId": "button2Action",
  "visible": false
}
```

The `visible` property is supported starting in Version 1.28 and later of the unified manifest for Microsoft 365.

### Change visibility at runtime

The shared runtime calls `Office.ribbon.requestUpdate()` using the control IDs from the manifest. The following example shows how to set the visibility of a control using the `requestUpdate` method.

```typescript
await Office.ribbon.requestUpdate({
  tabs: [
    {
      id: "ContosoVisibilityTab",
      groups: [
        {
          id: "PrimaryGroup",
          controls: [
            {
              id: "Button2",
              visible: true
            }
          ]
        }
      ]
    }
  ]
});
```

To show or hide a group, set the `visible` property on the applicable group object.

```typescript
await Office.ribbon.requestUpdate({
  tabs: [
    {
      id: "ContosoVisibilityTab",
      groups: [
        {
          id: "SecondaryGroup",
          visible: true
        }
      ]
    }
  ]
});
```

Only the properties included in the request are changed. The sample batches related group and control changes into one request when possible.

## Additional resources

- [Show or hide add-in commands on a custom tab](https://learn.microsoft.com/office/dev/add-ins/design/show-hide-controls-custom-tab)
-  [Ribbon API 1.3 requirement set](https://learn.microsoft.com/javascript/api/requirement-sets/common/ribbon-api-requirement-sets)
- [Add-in commands](https://learn.microsoft.com/office/dev/add-ins/design/add-in-commands)
- [Office.Ribbon interface](https://learn.microsoft.com/javascript/api/office/office.ribbon)
- [Office.RibbonUpdaterData interface](https://learn.microsoft.com/javascript/api/office/office.ribbonupdaterdata?view=common-js-preview)
- [Unified manifest for Microsoft 365](https://learn.microsoft.com/office/dev/add-ins/develop/unified-manifest-overview)
- [Ribbon group schema](https://learn.microsoft.com/microsoft-365/extensibility/schema/extension-ribbons-custom-tab-groups-item)
- [Ribbon control schema](https://learn.microsoft.com/microsoft-365/extensibility/schema/extension-common-custom-group-controls-item)

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2026 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/office-visibility" />
