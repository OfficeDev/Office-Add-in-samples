# Custom functions in Excel (Preview)

Learn how to use custom functions in Excel (similar to user-defined functions, or UDFs). Custom functions are JavaScript functions that you can add to Excel, and then use them like any native Excel function (for example =Sum). This sample accompanies the [Custom Functions Overview](https://docs.microsoft.com/office/dev/add-ins/excel/custom-functions-overview) topic.

## Table of Contents
* [Change History](#change-history)
* [Prerequisites](#prerequisites)
* [To use the project](#to-use-the-project)
* [Making changes](#making-changes)
* [Debugging](#debugging)
* [IntelliSense for the JSON file in Visual Studio Code](#intellisense-for-the-json-file-in-visual-studio-code)
* [Questions and comments](#questions-and-comments)
* [Additional resources](#additional-resources)

## Change History

* Oct 27, 2017: Initial version.
* April 23, 2018: Revised and expanded.
* June 1, 2018: Bug fixes.

## Prerequisites

* Install Office 2016 for Windows and join the [Office Insider](https://products.office.com/en-us/office-insider) program. You must have Office build number 10827 or later.

## To use the project

On a machine with a valid instance of an Excel Insider build installed, follow these instructions to use this custom function sample add-in:

1. On the machine where your custom functions project is installed, follow the instructions to install the self-signed certificates (https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) . 
2. From a command prompt from within your custom functions project directory, run `npm run start` to start a localhost server instance. 
3. Run `npm run sideload` to launch Excel and load the custom functions add-in. Additonal information on sideloading can be found at <https://aka.ms/sideload-addins>.
4. After Excel launches, you will need to register the custom-functions add-in to work around a bug:
    a. On the upper-left-hand side of Excel, there is a small hexagon icon with a dropdown arrow. The icon is to right of the Save icon.
    b. Click on this dropdown arrow and then click on the Custom Functions Sample add-in to register it.
5. Test a custom function by entering `=CONTOSO.ADD(num1, num2)` in a cell.
6. Try the other functions in the sample: `=CONTOSO.ADDASYNC(num1, num2)`, `CONTOSO.INCREMENTVALUE(increment)`.
7. If you make changes to the sample add-in, copy the updated files to your website, and then close and reopen Excel. If your functions are not available in Excel, re-insert the add-in using **Insert** > **My Add-ins**.

## Making changes
If you make changes to the sample functions code (in the JS file), close and reopen Excel to test them.

If you change the functions metadata (in the JSON file), close Excel and delete your cache folder `Users/<user>/AppData/Local/Microsoft/Office/16.0/Wef/CustomFunctions`. Then re-insert the add-in using **Insert** > **My Add-ins**.

## Debugging
Currently, the best method for debugging Excel custom functions is to first [sideload](https://docs.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing) your add-in within **Excel Online**. Then you can debug your custom functions by using the [F12 debugging tool native to your browser](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online). Use `console.log` statements within your custom functions code to send output to the console in real time.

If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.

If you are testing your add-in in Office 2016 desktop you can enable [runtime logging](https://docs.microsoft.com/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in) to debug issues with your add-in's XML manifest file as well as several installation and runtime conditions.

## IntelliSense for the JSON file in Visual Studio Code	
For intelliSense to help you edit the JSON file, follow these steps:

1. Open the JSON file (it has a .json extension) in Visual Studio Code.	
2. If you are starting a new file from scratch, add the following to the top of the file:	
	
     ```js	
    {	
        "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",	
    ```	
3. Press **Ctrl+Space** and intelliSense will prompt you with a list of all items that are valid at the cursor point. For example, if you pressed **Ctrl+Space** immediately after the `"$schema"` line, you are prompted to enter `functions`, which is the only key that is valid at that point. Select it and the `"functions": []` array is entered. If the cursor is between the `[]`, then you are prompted to enter an empty object as a member of the array. If the cursor is in the object, then you are prompted with a list of the keys that are valid in the object.

## Questions and comments

We'd love to get your feedback about this sample. You can send your feedback to us in the *Issues* section of this repository.

Questions about Microsoft Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). If your question is about the Office JavaScript APIs, make sure that your questions are tagged with [office-js] and [API].

## Additional resources

* [Custom functions overview](https://docs.microsoft.com/office/dev/add-ins/excel/custom-functions-overview)
* [Custom functions best practices](https://docs.microsoft.com/office/dev/add-ins/excel/custom-functions-best-practices)
* [Custom functions runtime](https://docs.microsoft.com/office/dev/add-ins/excel/custom-functions-runtime) 
* [Office add-in documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
* More Office Add-in samples at [OfficeDev on Github](https://github.com/officedev)

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Copyright
Copyright (c) 2017 Microsoft Corporation. All rights reserved.