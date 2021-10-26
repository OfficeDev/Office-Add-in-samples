# [ARCHIVED] Word-Add-in-Load-and-write-Open-XML

**Note:** This repo is archived and no longer actively maintained. Security vulnerabilities may exist in the project, or its dependencies. If you plan to reuse or run any code from this repo, be sure to perform appropriate security checks on the code or dependencies first. Do not use this project as the starting point of a production Office Add-in. Always start your production code by using the Office/SharePoint development workload in Visual Studio, or the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), and follow security best practices as you develop the add-in. 

This sample add-in shows you how to add a variety of rich content types to a Word document using the setSelectedDataAsync method with ooxml coercion type. The app also gives you the ability to show the Office Open XML markup for each sample content type right on the page.

**Description of the sample**

The add-in initializes in a blank Word document. You choose an option to insert the content or its markup at the selection point in the active Word document and then click the object type you want from the following options:

* formatted text
* styled text
* a simple image
* a formatted image
* a text box
* an Office drawing shape
* a content control
* a formatted table
* a styled table
* a SmartArt diagram
* a chart

Figure 1 shows how the task pane for the sample add-in appears when the solution starts.

![Figure 1. The Loading and Writing OOXML task pane](/description/9a7aa2da-4f99-4519-8cd1-f341060ff9beimage.png)

**Note**

When you choose the option to see the markup for a selected type of content, what you're seeing is the Office Open XML edited to remove unnecessary markup, along with a few tips for additional guidance. You can also review any piece of markup used in the add-in (with formatting to make it easier to navigate) directly in the Visual Studio solution. For further help interpreting, editing, and simplifying your work with Office Open XML for Word add-ins, see  Creating Better Add-ins for Word with Office Open XML.

Figures 2a - 2b show how the document surface and task pane appear after extracting Office Open XML from the selection.

![Figure 2a. Document surface appearance after using the 'Get…' button to extract Office Open XML for selected content](/description/70dee213-4853-47b2-abcf-55a982abb2c4image.png)

![Figure 2b. Task pane appearance after using the 'Get…' button to extract Office Open XML for selected content](/description/image.png)



**Prerequisites**

This sample requires:

* Visual Studio 2012
* Office 2013 tools for Visual Studio 2012
* Word 2013

**Key components of the sample**

The sample app contains:

* The LoadingAndWritingOOXML project, which contains:
* The LoadingAndWritingOOXML.xml manifest file
* The LoadingAndWritingOOXML Web project, which contains multiple template files
* However, the files that have been developed as part of this sample solution include:
* LoadingAndWritingOOXML.html (in the App folder, LoadingAndWritingOOXML subfolder). This contains the HTML user interface that is displayed in the task pane. It consists of two HTML radio buttons for choosing the option to insert a selected content type or display its markup in Word, several buttons for selecting a content type, and instructional text
* LoadingAndWritingOOXML.js (in the same folder as above). This script file contains code that runs when the app is loaded. This startup wires up the Click event handlers for the eleven buttons in LoadingAndWritingOOXML.html that represent different content types. The handler in the JavaScript connects each button to the correct function based on the actively-selected radio button, to either write the content or its markup into the document.
* Several XML files containing the markup for each of the content types you can insert via the app. These are located in the folder named OOXMLSamples. (Note that some content types have a separate XML file for the markup when inserting the object vs. displaying the markup on the page because chunks of binary data where applicable (i.e., for pictures and charts) are removed from the markup displayed on the page for ease of review. To learn more about the binary data contained in some types of Office Open XML markup, see the previously-referenced article  Creating Better Add-ins for Word with Office Open XML

All other files are automatically provided by the Visual Studio project template for Add-ins for Office, and they have not been modified in the development of this sample app.

**Configure the sample**

To configure the sample, open the LoadingAndWritingOOXML.sln file with Visual Studio. No other configuration is necessary.

**Build the sample**

To build the sample, choose the Ctrl+Shift+B keys.

**Run and test the sample**

To run the sample, choose the F5 key.

**Troubleshooting**

If the app fails to respond as described, try reloading it. (In the task pane, choose the down arrow, and then choose Reload.)

**Change log**

* First release: Aug 2013.
* GitHub release: Aug 2015.

**Related content**

* [Build Add-ins for Office](http://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Standard ECMA-376: Office Open XML File Formats](http://www.ecma-international.org/publications/standards/Ecma-376.htm)
* [Creating Better Apps for Word with Office Open XML](http://msdn.microsoft.com/EN-US/library/office/apps/dn423225.aspx)

 


This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
