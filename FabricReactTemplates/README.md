# Office UI Fabric React templates for Visual Studio 2019

## Summary
These templates can be used to create Office Add-ins that use [Office UI Fabric React](https://github.com/OfficeDev/office-ui-fabric-react). [Office UI Fabric JS](https://github.com/OfficeDev/office-ui-fabric-js) is no longer supported. [Office UI Fabric](https://developer.microsoft.com/fabric#) is recommended instead. 

**Templates**
- excel-ts: Creates an Excel add-in using TypeScript and Office UI Fabric React.
- excel-js: Creates an Excel add-in using JavaScript and Office UI Fabric React.

## Applies to
- Office 365
- Visual Studio 2019

## Prerequisites
- Visual Studio 2019 with the following workloads
  - ASP.NET & Web Development
  - Node.JS
  - Office / SharePoint Development
  And the following optional component
  - Net Core 2.2 development tools

## Instructions

### Create a new Excel add-in project

1. Copy the templates from this repo to %UserProfile%\Documents\Visual Studio 2019\Templates\ProjectTemplates
2. Start Visual Studio 2019.
3. Choose **Create a new project**.
4. Enter "Excel" into the search box
5. Choose the **Excel Web Add-in** project template, then choose **Next**.
6. The **Configure your new project** page has options to name your project (and solution), choose a disk location, and select a Framework version. Set the values you want for your new project and then choose **Create**.
7. On the **choose the add-in type** page, select **Add new functionalities to Excel**. Then choose **Finish**.

### Add the Office UI Fabric React template

1. Right-click the solution in **Solution Explorer** and choose **Add > New Project**.
2. In the **Add a new project** page, enter "excel" in the search window.
    - If you want to use TypeScript, choose the **Excel_Web Add_in_ASP.NET_Core_React** project template.
    - If you want to use JavaScript, choose the **Excel_WebAdd-in_AsP.NET_Core_ React_TypeScript** project template.
3. Select the add-in project in **Solution Explorer**. It will have the name you gave the project when you created it and a Manifest entry under it. 
4. Press F4 to view the **Properties** window for the project (if it is not already visible.)
5. Change the **Web Project** property to the new Office UI Fabric React project you just added.
6. Press F5 to see the completed Excel add-in project run in Excel.
> **Note:** npm install should run and install the packages but you may need to watch the output window for errors.


## Version history

Version  | Date | Comments
---------| -----| --------
1.0 | August 15, 2019 | Initial Release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**


