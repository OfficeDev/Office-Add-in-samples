# Mail merge in an Excel add-in

<img src="./assets/thumbnail.png" width="800" alt="A workbook with mail merge add-in open.">

This sample Office Add-in sends emails from inside Excel using the Microsoft Graph JavaScript SDK. You'll learn how to:

- Verify and validate data, such as email addresses.
- Send email with Microsoft Graph.
- Sign-in to Microsoft Graph to get proper permissions.

## How to run this sample

### Prerequisites

- Download and install [Visual Studio Code](https://visualstudio.microsoft.com/downloads/).
- Install the latest version of the [Office Add-ins Development Kit](https://marketplace.visualstudio.com/items?itemName=msoffice.microsoft-office-add-in-debugger) into Visual Studio Code.
- Node.js (the latest LTS version). Visit the [Node.js site](https://nodejs.org/) to download and install the right version for your operating system. To verify if you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.
- Microsoft Office connected to a Microsoft 365 subscription. You might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program), see [FAQ](https://learn.microsoft.com/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-) for details. Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try?rtc=1) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/buy/compare-all-microsoft-365-products).
  
### Run the add-in from the Office Add-ins Development Kit

1. **Download the sample code**

   Open the Office Add-ins Development Kit extension and view samples in the **Sample gallery**. Select the **Create** button in the top-right corner of the sample page.
   
1. **Open the Office Add-ins Development Kit**
    
    Select the <img src="./assets/Icon_Office_Add-ins_Development_Kit.png" width="30" alt="The Office Add-ins Development Kit icon in the activity bar of Visual Studio Code."/> icon in the **Activity Bar** to open the extension.

1. **Preview Your Office Add-in (F5)**

    Select **Preview Your Office Add-in(F5)** to launch the add-in and debug the code. In the drop down menu, select the option **Desktop (Edge Chromium)**.

    <img src="./assets/devkit_preview.png" width="500" alt="The 'Preview your Office Add-in' option in the Office Add-ins Development Kit's task pane."/>

    The extension checks that the prerequisites are met before debugging starts. The terminal will alert you to any issues with your environment. After this process, the Excel desktop application launches and opens a new workbook with the sample add-in sideloaded. The add-in automatically opens as well.

1. **Stop Previewing Your Office Add-in**

    Once you are finished testing and debugging the add-in, select **Stop Previewing Your Office Add-in**. This closes the web server and removes the add-in from the registry and cache.


## Use the sample add-in

An Excel desktop application will be auto-launched and the Mail Merge add-in will be auto-run on the right task pane area. The sideload steps has been integrated into the process, eliminating the need for manual intervention.

<img src="./assets/thumbnail.png" width="800" altText="A workbook with mail merge add-in open.">

Please follow the steps below:

1. Create sample data, including valid email address (required) and other information.
2. Verify template and data. the To Line must contain the column name of the email address.
3. Send email, which will pop up a dialog to get the consent of Microsoft Graph. After sign-in, the email will be sent. <br><img src="./assets/mail.png" width="600" altText="A mail to be sent">

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
| manifest.xml                  Manifest file
| package.json                  
| README.md                     
| RUN_WITH_EXTENSION.md         
| src/                          Add-in source code
|   | taskpane/
|   |   | consent.html          Consent HTML
|   |   | consent.js            Consent JavaScript
|   |   | taskpane.css          Task pane style
|   |   | taskpane.html         Task pane entry HTML
|   |   | taskpane.js           Add API calls and logic here
| webpack.config.js             Webpack config
```

## Feature details

`./src/taskpane` contains all the main page rendering, mail merge code logic, and Graph consent process.

1. The `taskpane.html` file is the main page of this project. In our sample project, we use several `text-area` boxes and `buttons` to interact with the backend for data and commands.
2. The `taskpane.js` file contains the main code logic of this add-in:
- The `createSampleData()` function uses the Excel JavaScript API to interact with the workbook. It inserts a sample table named "InvoiceTable" and fills it with the necessary data (email addresses) and other information.
- Your add-in can `get authorization to Microsoft Graph data` by obtaining an access token from the Microsoft identity platform. Use either the Authorization Code flow or the Implicit flow just as you would in other web applications, but with one exception: The Microsoft identity platform doesn't allow its sign-in page to open in an iframe. When an Office Add-in is running in Office on the web, the task pane is an iframe. This means you'll need to open the sign-in page in a dialog box using the Office dialog API. This affects how you use authentication and authorization helper libraries. For more information, see [Authentication with the Office dialog API](https://learn.microsoft.com/office/dev/add-ins/develop/auth-with-office-dialog-api).

- In our project, we use the `DialogAPIAuthProvider` class to open the sign-in page and get consent for Graph, which contains two functions: `getAccessToken()` and `login()`.
    - `getAccessToken()` checks whether the token already exists.
    - `login()` constructs a URL for the Graph login dialog and uses the Office JavaScript API to display this dialog. It sets up event handlers for dialog events. If the dialog sends a message with a status of 'success', it stores the received access token and resolves the Promise with it.

- The `sendEmail()` function replaces the column names in the to/subject/content text areas with the corresponding values in the table and sends emails row by row.
    
    The code for sending an email via `Microsoft Graph` is as follows:
    ```
    const sendMail = 
    {
        message: {
            subject: finalSubject,
            body: {
                contentType: 'Text',
                content: finalContent
            },
            toRecipients: [{
                emailAddress: {
                    address: addressValue[i][0]
                }
            }]
        }
    };

    await graphClient.api('me/SendMail')
        .post(sendMail);
    ```

## Reference

-   [Microsoft Graph website](https://graph.microsoft.io)
- For more information about Graph, please visit to the official documation: [Overview of Microsoft Graph](https://learn.microsoft.com/en-us/graph/overview)
-   The Microsoft Graph TypeScript definitions enable editors to provide intellisense on Microsoft Graph objects including users, messages, and groups.
    -   [@microsoft/microsoft-graph-types](https://www.npmjs.com/package/@microsoft/microsoft-graph-types) or [@types/microsoft-graph](https://www.npmjs.com/package/@types/microsoft-graph)
    -   [@microsoft/microsoft-graph-types-beta](https://www.npmjs.com/package/@microsoft/microsoft-graph-types-beta)
-   [Microsoft Graph Toolkit: UI Components and Authentication Providers for Microsoft Graph](https://docs.microsoft.com/graph/toolkit/overview)
-   [Office Dev Center](http://dev.office.com/)

## Additional resources

- Step-by-step training exercises that guide you through creating a basic application that accesses data via the Microsoft Graph:

    -   [Build Angular single-page apps with Microsoft Graph](https://docs.microsoft.com/graph/tutorials/angular)
    -   [Build Node.js Express apps with Microsoft Graph](https://docs.microsoft.com/graph/tutorials/node)
    -   [Build React Native apps with Microsoft Graph](https://docs.microsoft.com/graph/tutorials/react-native)
    -   [Build React single-page apps with Microsoft Graph](https://docs.microsoft.com/graph/tutorials/react)
    -   [Build JavaScript single-page apps with Microsoft Graph](https://docs.microsoft.com/graph/tutorials/javascript)
    -   [Explore Microsoft Graph scenarios for JavaScript development](https://docs.microsoft.com/learn/paths/m365-msgraph-scenarios/)

## Tips and Tricks

- [Microsoft Graph SDK `n.call is not a function` by Lee Ford](https://www.lee-ford.co.uk/posts/graph-sdk-is-not-a-function/)
- [Example of using the Graph JS library with ESM and `importmaps` ](https://github.com/waldekmastykarz/js-graph-101/blob/main/index_esm.html)

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
