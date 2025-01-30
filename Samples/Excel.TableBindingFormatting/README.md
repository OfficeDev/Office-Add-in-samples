# [ARCHIVED] Excel Task Pane Table Binding and Formatting #

**Note:** This sample is archived and no longer actively maintained. Security vulnerabilities may exist in the project, or its dependencies. If you plan to reuse or run any code from this repo, be sure to perform appropriate security checks on the code or dependencies first. Do not use this project as the starting point of a production Office Add-in. Always start your production code by using the Office/SharePoint development workload in Visual Studio, or the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), and follow security best practices as you develop the add-in.

For current samples, showing how to work with tables and formatting, see [Script Lab](https://learn.microsoft.com/office/dev/add-ins/overview/explore-with-script-lab).

To restore the project dependencies, rename the following file.

- Excel.TableBindingFormattingWeb/packages-archive.config -> Excel.TableBindingFormattingWeb/packages.config

### Summary ###
This code sample demonstrates techniques for creating a table binding, adding rows to an existing binding, applying table styles and applying cell formatting.

![](http://i.imgur.com/dex6lyr.png)

### Applies to ###
-  Excel 2013
-  Excel Online 2013

### Solution ###
Solution | Author(s)
---------|----------
Excel.TableBindingFormatting.sln | Doug Perkes (**Microsoft**)

### Version history ###
Version  | Date | Comments
---------| -----| --------
1.0  | April 23rd 2015 | Initial release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**


## Building the Sample ##
1. Open `ExcelFormattingSample.sln` in Visual Studio 2013. 
2. Right-click (or select and hold) on the `ExcelFormattingSample` project in Solution Explorer and choose **Set as startup project**. 
3. Press the <kbd>F5</kbd> key to build the solution and run it in Excel 2013.

## Description ##

This example demonstrates several techniques for working with table bindings in Excel 2013 using the JavaScript API for Office.

To get started, run the solution as described above. Once running, specify a number of rows and click the Create Sample Table button. This will create a sample table with random data containing five columns: Number of Widgets, Order Needed By, Month, Color and Customer. The color column is a hex value and the font color of the cell is set to the value.

- Once the table has been inserted, the additional buttons on the add-in become available for use. Clicking Add Rows adds the specified number of rows of random data to the table.

![](http://i.imgur.com/2n4kNew.png)

- Select a table style from the Table Options drop down list and click the Table Options button to apply the style.

![](http://i.imgur.com/4tkMndG.png)

- Use the Range Formatting section to apply a border around the cells specified by the start and end row/col input boxes.
![](http://i.imgur.com/dgzD5kp.png)

> Note: The rows and columns are zero based and relative to the start of the table. They are not representative of the row numbers and column letters in Excel.

![](http://i.imgur.com/RT9YLob.png)

- To clear all formatting from the table, click the Clear Format button.

- Clicking Clear Data will remove all the rows from the table, but leave the table headers.

## Cell Formatting ##

The code samples demonstrate a valuable technique for applying cell formatting to a large range of cells. When applying cell formatting in Excel Online, the number of format groups passed to the `cellFormat` parameter cannot exceed 100. To get around this limitation the code demonstrates the use of a queue and a recursive function for applying the formatting in sets of 100.

## Source Code Files ##

The key source code files in this project are the following

- `Excel.TableBindingFormattingWeb\App\Home\Home.html` - contains the html controls and formatting for the UI of the add-in. 
- `Excel.TableBindingFormattingWeb\App\Home\Home.js` - contains the application logic for creating and manipulating the table 

## More Information ##

For more information see the [JavaScript API for Office](https://msdn.microsoft.com/en-us/library/office/fp142185.aspx "JavaScript API for Office").

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/Excel.TableBindingFormatting" />