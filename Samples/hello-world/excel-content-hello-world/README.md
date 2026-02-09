---
page_type: sample
urlFragment: excel-content-add-in-hello-world
products:
  - office-excel
  - office
  - m365
languages:
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: '02/04/2026 10:00:00 AM'
description: 'Create a Excel content add-in that gets the selected text then displays it.'
---

# Create a Excel content add-in that gets and displays selected text

## Summary

Learn how to build an Office content add-in that gets and displays the text selected from a Excel worksheet.

![The content add-in open in Excel.](../images/excel-content-add-in-hello-world.png)

## Applies to

- Excel on Windows, Mac, and in a browser.

## Version history

| Version  | Date | Comments |
|----------|------|----------|
| 1.0 | 02-04-2026 | Initial release |

## Prerequisites

- Microsoft 365 - You can get a free developer sandbox by joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program#Subscription).

### Get and display selected text

When the user chooses the **Get data from selection** button, the `getDataFromSelection()` function is called. This function calls `Office.context.document.getSelectedDataAsync()` to get the selected text from the worksheet. "Hello, world!" and the selected text are then displayed in the content add-in.

For more information, see [Content add-ins](https://learn.microsoft.com/office/dev/add-ins/design/content-add-ins?tabs=jsonmanifest).

```javascript
// Reads data from current document selection and displays it.
function getDataFromSelection() {
    if (Office.context.document.getSelectedDataAsync) {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    document.getElementById("selected-data").textContent = 'Hello, world! The selected text is: ' + result.value;
                } else {
                    document.getElementById("selected-data").textContent = 'Error getting selected text.';
                    console.error('Error:', result.error.message);
                }
            });
    } else {
        document.getElementById("selected-data").textContent = 'Error: Reading selection data isn\'t supported by this host application.';
        console.error('Error:', 'Reading selection data isn\'t supported by this host application.');
    }
}
```

## Run the sample

### Run the sample with GitHub as the host

An Office Add-in requires you to configure a web server to provide all the resources, such as HTML, image, and JavaScript files. The Hello World sample is configured so that the files are hosted directly from this GitHub repo, so all you need to do is build the manifest and package, and then sideload the package.

1. Clone or download this sample to a folder on your computer. Then in a command prompt, bash shell, or **TERMINAL** in Visual Studio Code, navigate to the root of the sample folder.
1. Run the command `npm install`.
1. Run the command `npm run build`.
1. Run the command `npm run start:prod`.

   After a few seconds, desktop Excel opens, and after a few seconds more, the content add-in appears over the current worksheet with a **Get data from selection** button.
     - If the content add-in doesn't appear, open the **Add-ins** button in the **Home** tab of the ribbon, then select the name of the content add-in, "Excel Content Add-in".

1. Choose the **Get data from selection** button to display "Hello, world!" and the selected text.

When you're finished working with the add-in, close Excel, and then in the window where you ran the three npm commands, run `npm run stop:prod`.

### Configure a localhost web server and run the sample from localhost

If you prefer to configure a web server and host the add-in's web files from your computer, use the following steps.

1. Clone or download this sample to a folder on your computer. Then in a command prompt, bash shell, or **TERMINAL** in Visual Studio Code, navigate to the root of the sample folder.
1. Run the command `npm install`.
1. Run the command `npm start`.

   - If you've never developed an Office Add-in on this computer before or it has been more than 30 days since you last did, you'll be prompted to delete an old security cert or install a new one. Agree to both prompts.
   - After a few seconds, a **webpack** dev-server window will open and your files will be hosted there on localhost:3000.
   - When the server is successfully running, desktop Excel opens, and after a few seconds more, the content add-in appears over the current worksheet with a **Get data from selection** button.
     - If the content add-in doesn't appear, open the **Add-ins** button in the **Home** tab of the ribbon, then select the name of the content add-in, "Excel Content Add-in".

1. Ensure that there's text in the worksheet, then select any range of cells containing text in the worksheet.
1. Choose the **Get data from selection** button to display "Hello, world!" and the selected text.

When you're finished working with the add-in, close Excel, and then in the window where you ran the two npm commands, run `npm stop`.

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

**Note**: The content.html file contains an image URL that tracks diagnostic data for this sample add-in. Please remove the image tag if you reuse this sample in your own code project.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/excel-content-add-in-hello-world" />
