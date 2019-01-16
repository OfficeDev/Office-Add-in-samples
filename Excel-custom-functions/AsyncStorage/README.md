# Using AsyncStorage in custom functions

Learn how to use pass data values from custom functions to task panes, or from task panes to custom functions.

## Change History

* Jan 15, 2019: Initial version.

## Prerequisites

* Install Office 2016 for Windows and join the [Office Insider](https://products.office.com/en-us/office-insider) program. You must have Office build number 10827 or later.

## To use the project

On a machine with a valid instance of an Excel Insider build installed, follow these instructions to use this custom function sample add-in:

1. On the machine where your custom functions project is installed, follow the instructions to install the self-signed certificates (https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) . 
2. From a command prompt from within your custom functions project directory, run `npm run start` to start a localhost server instance. 
4. After Excel launches, you will need to register the custom-functions add-in to work around a bug:
    a. On the upper-left-hand side of Excel, there is a small hexagon icon with a dropdown arrow. The icon is to right of the Save icon.
    b. Click on this dropdown arrow and then click on the Custom Functions Sample add-in to register it.
5. The task pane has further instructions on how to store and retrieve values in AsyncStorage.

## Making changes
If you make changes to the sample functions code (in the JS file), close and reopen Excel to test them.

If you change the functions metadata (in the JSON file), close Excel and delete your cache folder `Users/<user>/AppData/Local/Microsoft/Office/16.0/Wef/CustomFunctions`. Then re-insert the add-in using **Insert** > **My Add-ins**.


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
Copyright (c) 2019 Microsoft Corporation. All rights reserved.
