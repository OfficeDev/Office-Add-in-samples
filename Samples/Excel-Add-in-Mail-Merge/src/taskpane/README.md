# Project Code Explaination

This sample demonstrates how to use the Microsoft Graph JavaScript SDK to send emails in Excel from Office Add-ins.

## Manifest

The `manifest.xml`  file of an Office Add-in describes how your add-in should be activated when an end user installs and uses it with Office documents and applications. A manifest file enables an Office Add-in to do the following:

- Describe itself by providing an ID, version, description, display name, and default locale.

- Specify the images used for branding the add-in and iconography used for add-in commands in the Office app ribbon.

- Specify how the add-in integrates with Office, including any custom UI, such as ribbon buttons the add-in creates.

- Specify the requested default dimensions for content add-ins, and requested height for Outlook add-ins.

- Declare permissions that the Office Add-in requires, such as reading or writing to the document.

As you modify your `manifest.xml` file, use the included [Office Toolbox](https://github.com/OfficeDev/office-toolbox) to ensure that your XML file is correct and complete. It will also give you information on against what platforms to test your add-ins before submitting to the store.

To run Office Add-in Validator, use the following command in your project directory:

```bash
npm run validate
```

For more information on manifest validation, refer to our [add-in manifests documentation](https://learn.microsoft.com/office/dev/add-ins/develop/add-in-manifests).

## Taskpane folder

Taskpane folder contains all the Main Page Drawing, Mail Merge code-logic and Graph Consent Process.

1. The `taskpane.html` file is the main page of the whole project. In our sample project, we use several `text-area boxes` and `buttons` to interact with the backend for data and commands.

- **Note**: The taskpane.html file contains an image URL that tracks diagnostic data for this sample add-in. Please remove the image tag if you reuse this sample in your own code project.

    < img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/excel-add-in-mail-merge" >

2. The `taskpane.js` file contains the main code logic of the whole add-in.
- `createSampleData()` function using Excel JavaScript API to interact with the workbook. Insert a sample table named "InvoiceTable" and fill it with the necessary data(email address) and other information.
- Your add-in can `get authorization to Microsoft Graph data` by obtaining an access token to Microsoft Graph from the Microsoft identity platform. Use either the Authorization Code flow or the Implicit flow just as you would in other web applications but with one exception: The Microsoft identity platform doesn't allow its sign-in page to open in an iframe. When an Office Add-in is running in Office on the web, the task pane is an iframe. This means you'll need to open the sign-in page in a dialog box by using the Office dialog API. This affects how you use authentication and authorization helper libraries. For more information, see [Authentication with the Office dialog API](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/auth-with-office-dialog-api).

    In our project, we use `DialogAPIAuthProvider` class to open the sign-in page and get the consent of Graph, which contains two functions: `getAccessToken()` and `login()`.
    - `getAccessToken()` checks whether the token already exists.
    - `login()` constructs a URL for the Graph login dialog and uses the Office JavaScript API to display this dialog. It sets up event handlers for dialog events. If the dialog sends a message with a status of 'success', it stores the received access token and resolves the Promise with it.

- `sendEmail()` function replace the column name in the to/subject/content textarea with the corresponding value in the table, and send email row by row.
    
    The code of sending email via `Microsoft Graph` is as below:
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

