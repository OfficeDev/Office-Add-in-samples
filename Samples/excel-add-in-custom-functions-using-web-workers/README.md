# Build Asynchronous Custom Functions in Excel with Web Workers

<img src="./assets/thumbnail.png" width="800" alt="A workbook with custom function using web worker.">

This sample shows how to use web workers in custom functions to prevent blocking UI of your add-in. You'll learn how to:

- Create custom functions in Excel
- Use web workers

## How to run this sample

### Prerequisites

- Download and install [Visual Studio Code](https://visualstudio.microsoft.com/downloads/).
- Install the latest version of the [Office Add-ins Development Kit](https://marketplace.visualstudio.com/items?itemName=msoffice.microsoft-office-add-in-debugger) into Visual Studio Code.
- Node.js (the latest LTS version). Visit the [Node.js site](https://nodejs.org/) to download and install the right version for your operating system. To verify if you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.
- Microsoft Office connected to a Microsoft 365 subscription. You might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program), see [FAQ](https://learn.microsoft.com/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-) for details. Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try?rtc=1) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/buy/compare-all-microsoft-365-products).
  
### Run the add-in from the Office Add-ins Development Kit

1. Create a new project with the sample code.

   Open the Office Add-ins Development Kit extension and view samples in the **Sample gallery**. Select the **Create** button in the top-right corner of the sample page. The new project will open in a second Visual Studio Code window. Close the original VSC window.
   
1. Open the Office Add-ins Development Kit.
    
    Select the <img src="./assets/Icon_Office_Add-ins_Development_Kit.png" width="30" alt="The Office Add-ins Development Kit icon in the activity bar of VSCode"/> icon in the **Activity Bar** to open the extension.

1. Preview Your Office Add-in (F5).

    Select **Preview Your Office Add-in(F5)** to launch the add-in and debug the code. In the drop down menu, select the option **Desktop (Edge Chromium)**.

    <img src="./assets/devkit_preview.png" width="500" alt="The 'Preview your Office Add-in' option in the Office Add-ins Development Kit's task pane."/>

    The extension checks that the prerequisites are met before debugging starts. The terminal will alert you to any issues with your environment. After this process, the Excel desktop application launches and opens a new workbook with the sample add-in sideloaded. The add-in automatically opens as well.

    If this is the first time that you have sideloaded an Office Add-in on your computer (or the first time in over a month), you may be prompted to delete an old certificate and/or to install a new one. Agree to both prompts. The first run requires installing dependency of this project, which might take 2~3 minutes or longer. During this time, there might be a dialog pop up at the lower right of the VSC screen. You should not interact with this dialog before the Office application launched.

1. Stop Previewing Your Office Add-in.

    Once you are finished testing and debugging the add-in, select **Stop Previewing Your Office Add-in**. This closes the web server and removes the add-in from the registry and cache.

## Use the sample add-in

1. Select the **Run function without web workers** button. A custom function without a web worker will be inserted in the cell **A1** and run. Note that this stops the bouncing ball inside the task pane.
2. Select the **Run function with web worker** button. A custom function with a web worker will be inserted in cell **A1** and run. Note that the bouncing ball inside the task pane doesn't stop.

The add-in adds the following custom functions to the workbook.

- `=WebWorkerSample.TEST(2)`: Post the TEST function to web worker to do the calculation that returns the computing result.
- `=WebWorkerSample.TEST_PROMISE(2)`: Post the TEST_PROMISE function to web worker that returns a promise that resolves with the calculation result.
- `=WebWorkerSample.TEST_ERROR(2)`: Post the TEST_ERROR function to web worker that returns an error.
- `=WebWorkerSample.TEST_ERROR_PROMISE(2)`: Post the TEST_ERROR_PROMISE function to web worker that returns a promise that rejects with an error.
- `=WebWorkerSample.TEST_UI_THREAD(2)`: Do the calculation using the UI thread.

Open the task pane to see an animated ball bouncing. This shows the effect of blocking the UI thread. Enter `=WebWorkerSample.TEST_UI_THREAD(50000)` into a cell to cause the thread to be blocked for five seconds. The bouncing ball will stop while the function result is calculated.

## Explore sample files

These are the important files in the sample project.

```
| .eslintrc.json
| .gitignore
| .vscode/
|   | extensions.json
|   | launch.json               Launch and debug configurations
|   | settings.json             
|   | tasks.json                
| assets/                       Static assets, such as images
| babel.config.json
| manifest*.xml                 Manifest file
| package.json                  
| README.md                     
| RUN_WITH_EXTENSION.md         
| SECURITY.md
| src/                          Add-ins source code
|   | taskpane/
|   |   | taskpane.html         Task pane entry HTML
|   |   | taskpane.js           Add API calls and logic here
|   | functions/
|   |   | functions.js          Custom function JavaScript
|   |   | functions-worker.js   Web worker JavaScript
| webpack.config.js             Webpack config
```

## Details

### Dispatch to web worker

To have a custom function use a web worker, turn the calculation into a job and dispatch it to the web worker. In this sample, the `dispatchCalculationJob` function takes the function name and parameters from a custom function. It then creates a job object that is posted to a web worker. For more details see the `dispatchCalculationJob` function in [functions.js](src/functions/functions.js).

### Run the job and return the result

The web worker runs the job specified in the job object for the actual calculation. This sample's web worker code is in a separate file, [functions-worker.js](src/functions/functions-worker.js).

The functions-worker.js will:

1. Receive a message that contains the job to run.
1. Invoke a function to perform the calculation.
1. Call **postMessage** to post the result back to the main thread.

Most of the code handles the error case and Promise case.

### Process results from the web worker

In [functions.js](src/functions/functions.js), when a new web worker is created, it's provided a callback function to process the result. The callback function parses the data to determine the outcome of the job. It resolves or rejects the promise, as determined by the job result data.

## Troubleshooting

If you have problems running the sample, take the following steps.

- Close any open instances of Excel.
- Close the previous web server started for the sample with the **Stop Previewing Your Office Add-in** Office Add-ins Development Kit extension option.
- Try to run the sample again.

If you still have problems, see [troubleshoot development errors](https://learn.microsoft.com//office/dev/add-ins/testing/troubleshoot-development-errors) or [create a GitHub issue](https://aka.ms/officedevkitnewissue) and we'll help you.  

For information on running the sample on Excel on the web, see [Sideload Office Add-ins to Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).

For information on debugging on older versions of Office, see [Debug add-ins using developer tools in Microsoft Edge Legacy](https://learn.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-devtools-edge-legacy).

## Make code changes

Once you understand the sample, make it your own! All the information about Office Add-ins is found in our [official documentation](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins). You can also explore more samples in the Office Add-ins Development Kit. Select **View Samples** to see more samples of real-world scenarios.

If you edit the manifest as part of your changes, use the **Validate Manifest File** option in the Office Add-ins Development Kit. This shows you any errors in the manifest syntax.

## Engage with the team

Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.

Want to learn more about new features and best practices for the Office platform? [Join the Microsoft Office Add-ins community call](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins-community-call).

## Copyright

Copyright (c) 2024 Microsoft Corporation. All rights reserved.
This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
