# Excel Tutorial - Completed

This sample is the result of completing the [Tutorial: Create an Excel task pane add-in](https://learn.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial). It was constructed with the [Yeoman generator for Office Add-ins](https://learn.microsoft.com/office/dev/add-ins/develop/yeoman-generator-overview).

The tutorial gives step-by-step instructions on how to add functionality alongside explanations as to why code is being added. This sample is best used if you've completed the tutorial and want to start from a more complete project. It's also helpful when debugging problems you encounter during the tutorial.

## Run the add-in

1. Navigate to the root directory of the `My Office Add-in` project.

1. Run `npm install`

1. Start the local web server and sideload your add-in:

- To test your add-in in Excel, run the following command in the root directory of your project. This starts the local web server (if it's not already running) and opens Excel with your add-in loaded.

  - Windows: `npm start`
  - Mac: `npm run dev-server`

- To test your add-in in Excel on the web, run the following command in the root directory of the `My Office Add-in` project. When you run this command, the local web server starts. Replace "{url}" with the URL of an Excel document on your OneDrive or a SharePoint library to which you have permissions.

  ```command line
  npm run start:web -- --document {url}
  ```

  The following are examples.

  - `npm run start:web -- --document https://contoso.sharepoint.com/:t:/g/EZGxP7ksiE5DuxvY638G798BpuhwluxCMfF1WZQj3VYhYQ?e=F4QM1R`
  - `npm run start:web -- --document https://1drv.ms/x/s!jkcH7spkM4EGgcZUgqthk4IK3NOypVw?e=Z6G1qp`
  - `npm run start:web -- --document https://contoso-my.sharepoint-df.com/:t:/p/user/EQda453DNTpFnl1bFPhOVR0BwlrzetbXvnaRYii2lDr_oQ?e=RSccmNP`

  > NOTE: If you are developing on a Mac, enclose the {url} in single quotation marks. Do not do this on Windows.

    If your add-in doesn't sideload in the document, manually sideload it by following the instructions in Manually sideload add-ins to Office on the web.

1. If the add-in task pane isn't already open in Excel, go to the Home tab and choose the **Show Taskpane** button in the ribbon to open it.

1. Use the buttons in the taskpane and the **Toggle Worksheet Protection** button to interact with the workbook through your add-in.
