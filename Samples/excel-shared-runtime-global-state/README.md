---
page_type: sample
urlFragment: office-add-in-shared-runtime-global-data
products:
  - m365
  - office
  - office-excel
languages:
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: 3/15/2020 1:25:00 PM
description: "This sample shows how to share data across the ribbon, task pane, and custom functions."
---

# Share global data across add-in components

## Summary

This sample shows how to set up a basic project that uses the shared runtime. The shared runtime runs all parts of the Excel add-in (ribbon buttons, task pane, custom functions) in a single browser runtime. This makes it easy to shared data through local storage, or through global variables.

![Screen shot of the add-in with ribbon buttons enabled and disabled](excel-shared-runtime-global.png)

## Features

- Share data globally with ribbon buttons, the task pane, and custom functions.
- Demonstrates shared runtime with custom functions in a unified manifest.
- To get started, use either the unified manifest (manifest.json) or the XML manifest (manifest.xml).

## Applies to

- Excel on Windows, Mac, and the web.

## Prerequisites

- Microsoft 365.
- Office 2304 (Build 16320.20000) or later for unified manifest support.

## Solution

Solution | Author(s)
---------|----------
Office Add-in share global data with a shared runtime | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.0 | 3-15-2020 | Initial release.
1.1 | May 26, 2021 | Updated to use GitHub pages for hosting.
1.2 | April 2026 | Added unified manifest support.

----------

## Scenario: Sharing key/value pairs

This sample enables a user to store and retrieve key/value pairs by using the task pane or custom functions. The user can select which type of storage is used. They can choose to store key/value pairs in local storage, or choose to use a global variable.

## Run the sample with the unified manifest

This sample includes a **manifest.json** file that uses the unified manifest format for Microsoft 365. The unified manifest provides:

- Modern JSON format instead of XML
- Support for custom functions with shared runtime
- Streamlined configuration for Office Add-ins

**Important:** Custom functions are only available in preview with the unified manifest. Do not use custom functions with the unified manifest in a production add-in.

**Note:** The unified manifest requires Office 2304 (Build 16320.20000) or later.

### Using the unified manifest

The unified manifest configures the shared runtime and custom functions in a single `runtime` object:

```json
{
  "runtimes": [
    {
      "id": "SharedRuntime",
      "type": "general",
      "lifetime": "long",
      "code": {
        "page": "https://localhost:3000/src/taskpane/taskpane.html",
        "script": "https://localhost:3000/src/functions/functions.js"
      },
      "actions": [...],
      "customFunctions": {
        "functions": [...],
        "namespace": {
          "id": "CONTOSO",
          "name": "CONTOSO"
        },
        "metadataUrl": "https://localhost:3000/src/functions/functions.json"
      }
    }
  ]
}
```

### Building and Running with Unified Manifest

1. Install dependencies:
   ```bash
   npm install
   ```

1. Start the development server:
   ```bash
   npm start
   ```

1. The add-in loads in Excel with the unified manifest.

For production builds with GitHub Pages URLs:
```bash
npm run build
npm run start:prod
```

**Note:** The unified manifest requires Office 2304 (Build 16320.20000) or later. Custom functions support in unified manifest is available in schema version 1.25+.

## Run the sample with the XML manifest in Excel on the web

You can run this sample in Excel on the web. The add-in web files are served from this repo on GitHub.

1. Download the **manifest.xml** file from this sample to a folder on your computer.
1. Open [Office on the web](https://office.live.com/).
1. Choose **Excel**, and then open a new document.
1. Open the **Insert** tab on the ribbon and choose **Office Add-ins**.
1. On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.
   ![The Office Add-ins dialog with a drop-down in the upper right reading "Manage my add-ins" and a drop-down below it with the option "Upload My Add-in"](../../Samples/images/office-add-ins-my-account.png)
1. Browse to the add-in manifest file, and then select **Upload**.
   ![The upload add-in dialog with buttons for browse, upload, and cancel.
](../../Samples/images/upload-add-in.png)
1. Verify that the add-in loaded successfully. You will see a **Show Taskpane** button on the **Home** tab on the ribbon.

Once the add-in is loaded use the following steps to try out the functionality.

1. On the `Home` tab, choose `Show TaskPane`.
1. In the task pane, enter a key/value pair, and choose `Store key/value pair`.
![Screen shot of both key and value input fields, and both store and get buttons.](task-pane-buttons.png)
1. In any spreadsheet cell, enter the formula `=CONTOSO.GETVALUEFORKEYCF("1")`. Pass the value of the key you created from the task pane.
1. In any spreadsheet cell, enter the formula `=CONTOSO.SETVALUEFORKEYCF("2","oranges")`. The formula should return the text `Stored key/value pair`.
1. In the task pane, enter the key from the previous formula `2` and choose `Get value for key`. The task pane should display the value `oranges`.

The task pane and custom function share data via a global variable in the shared runtime. You can switch the method of storage by choosing either the `Global variable` or `Local storage` radio buttons on the task pane.

## Run the sample with the XML manifest from localhost

If you prefer to host the web server for the sample on your computer, follow these steps:

1. You need http-server to run the local web server. If you haven't installed this yet you can do this with the following command:

    ```console
    npm install --global http-server
    ```

1. Use a tool such as openssl to generate a self-signed certificate that you can use for the web server. Move the cert.pem and key.pem files to the root folder for this sample.
1. From a command prompt, go to the root folder and run the following command:

    ```console
    http-server -S --cors . -p 3000
    ```

1. To reroute to localhost run office-addin-https-reverse-proxy. If you haven't installed this you can do this with the following command:

    ```console
    npm install --global office-addin-https-reverse-proxy
    ```

    To reroute run the following in another command prompt:

    ```console
    office-addin-https-reverse-proxy --url http://localhost:3000
    ```

1. Follow the steps in [Run the sample with the XML manifest in Excel on the web](#run-the-sample-with-the-xml-manifest-in-excel-on-the-web), but upload the `manifest-localhost.xml` file for step 6.

## Key parts of this sample

The manifest.xml is configured to use the shared runtime by using the `Runtimes` element as follows:

```xml
<Runtimes>
   <Runtime resid="Shared.Url" lifetime="long" />
</Runtimes>
```

In other parts of the manifest, you'll see that the custom functions and task pane are also configured to use the `Shared.Url` because they all run in the same runtime. `Shared.Url` points to `taskpane.html` which loads the shared runtime.

Global state is tracked in a window object retrieved using a `getGlobal()` function. This is accessible to custom functions, the task pane, and the ribbon (because all the code is running in the same JavaScript runtime.) 

There are no commands.html or functions.html files. These are not necessary because their purpose is to load individual runtimes. These do not apply when using the shared runtime.

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2020 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/officedev/samples/excel-shared-runtime-global-state" />
