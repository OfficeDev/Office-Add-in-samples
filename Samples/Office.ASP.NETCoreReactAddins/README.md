# ASP .NET Core, React.js and Office UI Fabric React sample TaskPane Web projects for Visual Studio 2019

## Summary

These are sample web projects for Office TaskPane Add-ins. These samples use [ASP.NET Core](https://github.com/aspnet/AspNetCore), [React.js](https://reactjs.org/) and [Office UI Fabric React](https://github.com/OfficeDev/office-ui-fabric-react).

## Sample Web Projects

- excel-ts: contains sample TypeScript web projects that can be used with Office Web-Addin in Visual Studio.
- excel-js: contains sample JavaScript web projects that can be used with Office Web-Addin in Visual Studio.

## Applies to

- Office 365
- Visual Studio 2019

## Prerequisites

- Visual Studio 2019 (16.3 Preview 3 or newer) and the following Workloads and Optional Components installed:
  - ASP.NET & Web Development
  - Node.JS
  - Office / SharePoint Development
  - Net Core 2.2 development tools (optional component)

- Node.js, install from https://nodejs.org

>**Note:** Update 3 is currently in Preview.  It can be downloaded from https://visualstudio.microsoft.com/vs/preview/

## Instructions

### Create a new Office Web add-in project

1. Start Visual Studio 2019.
2. Choose **Create a new project**.
3. Type  "Excel" into the search box at the top of the dialog.
4. Choose the Excel Web Add-in template, then choose **Next**.
5. Give the new project a name on the **Configure your new project** page and use the default values for the remaining fields and then choose **Create**.
6. On the choose the add-in type page, select **Add new functionalities to Excel**. Then choose Finish.

### Add the ASP.NET Core React.js web project to the solution

1. Download the files for the sample project from this repo that you wish to use.
2. In Visual Studio, right-click the solution in **Solution Explorer** and choose **Add > Existing Project**.
3. In the **Add Existing Project** dialog, browse to the project you just downloaded and select the .csproj file.
4. Build the project you just added, right-click the project in **Solution Explorer** and choose **Build**.
5. Select the add-in project in Solution Explorer. It will have the name you gave the project when you created it and a Manifest entry under it (i.e. ExcelWebAddin1).
6. Press F4 to view the **Properties** window for the project (if it is not already visible).
7. Change the **Web Project** property to the name of ASP .NET Core web project you just added to the solution.
8. Copy the contents of the manifest.xml file in the ASP .NET Core project and replace all of the contents in the maninfest file in the Web Add-in project (i.e. ExcelWebAddin1Manifest.xml) with it.
9. Press F5 to debug the Office Web Add-in project.

> **Note:** npm install should run and install the packages prior to building the ASP.NET Core web project but you may need to watch the output window for errors.  If errors occur, please try running npm install from the ./ClientApp folder.

## Version history

Version  | Date | Comments
---------| -----| --------
1.0 | September 10, 2019 | Initial Release

### Disclaimer ###

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
