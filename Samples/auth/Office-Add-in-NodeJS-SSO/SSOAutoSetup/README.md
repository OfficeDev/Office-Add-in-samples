---
page_type: sample
products:
- office-excel
- office-powerpoint
- office-word
- microsoft-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  - Microsoft Graph
  services:
  - Excel
  - Microsoft 365
  createdDate: 5/1/2017 2:09:09 PM
description: "This sample shows how to add support for SSO in your add-in."
---

# Office Add-in that supports Single Sign-on to Office, the add-in, and Microsoft Graph

THIS README ASSUMES THAT YOU HAVE READ THE TOP LEVEL README IN THE ROOT OF THIS REPO. IT CONTAINS IMPORTANT INFORMATION ABOUT THE SAMPLES, **INCLUDING PREREQUISITES**.

## To use the project

### Register the add-in

1. npm install
1. Register your application in Azure by running the following NPM script at the root of your project folder where package.json is located: **npm run configure-sso**

- Your browser will open and prompt for authentication. Enter the user name and password of a user with tenant admin permissions. If you created an account using [Microsoft 365 developer program](https://aka.ms/devprogramsignup), this should suffice.
- Once you have successfully logged in, you will see the script reporting each step it takes in the command shell.

### Run the solution

1. Open a command prompt in the root of the project.
2. Run the command `npm start`.
3. You may be prompted to register the dev-certificates for the dev-server.  Choose 'Yes" for this dialog.  **NOTE:** The dev-certs dialog may not be readily visible if you have many windows open, so you may need to minimize other windows to see it.
4. Excel will automatically start by default.  You can change the default desktop application to Word or PowerPoint by updating the **app-to-debug** in the config section of package.json
5. In the Office application, on the **Home** ribbon, select the **Show Add-in** button in the **SSO Node.js** group to open the task pane add-in.
6. Click the **Get OneDrive File Names** button. If you are logged into Office with either a Microsoft 365 Education or work account, or a Microsoft account, and SSO is working as expected, the first 10 file and folder names in your OneDrive for Business are inserted into the document. (It may take as much as 15 seconds the first time.) If you are not logged in, or you are in a scenario that does not support SSO, or SSO is not working for any reason, you will be prompted to log in. After you log in, the file and folder names appear.

## Security note

The sample sends a hardcoded query parameter on the URL for the Microsoft Graph REST API. If you modify this code in a production add-in and any part of query parameter comes from user input, be sure that it is sanitized so that it cannot be used in a Response header injection attack.

## Questions and comments

We'd love to get your feedback about this sample. You can send your feedback to us in the *Issues* section of this repository.
Questions about developing Office Add-ins should be posted to [Microsoft Q&A](https://aka.ms/office-js-dev-questions).

## Additional resources

- [Microsoft Graph documentation](https://docs.microsoft.com/graph/)
- [Office Add-ins documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)

## Copyright

Copyright (c) 2019 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
