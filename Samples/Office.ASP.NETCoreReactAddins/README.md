# ASP .NET Core, React.js and Office UI Fabric React sample task pane web projects for Visual Studio 2019

## Summary

These are sample web projects for Office task pane add-ins. These samples use [ASP.NET Core](https://github.com/aspnet/AspNetCore), [React.js](https://reactjs.org/) and [Office UI Fabric React](https://github.com/OfficeDev/office-ui-fabric-react).

## Sample Web Projects

- excel-ts: contains sample TypeScript web project that can be used with Office web add-in in Visual Studio.
- excel-js: contains sample JavaScript web project that can be used with Office web add-in in Visual Studio.

## Applies to

- Office 365
- Visual Studio 2019

## Prerequisites

- Visual Studio 2019 (16.3 Preview 3 or newer) and the following Workloads and Optional Components installed:
  - ASP.NET and web development
  - Node.js development
  - Office/SharePoint development
  - Net Core 2.2 Runtime (individual component)

- Node.js, install from https://nodejs.org

>**Note:** Update 3 is currently in Preview.  It can be downloaded from https://visualstudio.microsoft.com/vs/preview/

## Solution

Solution | Author(s)
---------|----------
Add ASP.NET Core with Office UI Fabric React to your Office Add-in | Microsoft

## Instructions

### Create a new Office Web add-in project

1. Start Visual Studio 2019.
2. Choose **Create a new project**.
3. Type  "Excel" into the search box at the top of the dialog.
4. Choose the Excel Web Add-in template, then choose **Next**.
5. Give the new project a name on the **Configure your new project** page and use the default values for the remaining fields and then choose **Create**.
6. On the choose the add-in type page, select **Add new functionalities to Excel**. Then choose Finish.

### Add the ASP.NET Core React.js web project to the solution

1. Download or clone this repo. This will create a **PnP-OfficeAddins** folder.
2. In Visual Studio, right-click the solution in **Solution Explorer** and choose **Add > Existing Project**.
3. In the **Add Existing Project** dialog, go to the **PnP-OfficeAddins/Samples/Office.ASP.NETCoreReactAddins** folder. There are two project folders there: **excel-js** (JavaScript) and **excel-ts** (TypeScript). Choose the folder for the language you want to use, and then choose the .csproj file in the folder.
4. Build the project you just added, right-click the project in **Solution Explorer** and choose **Build**.
5. Select the add-in project in Solution Explorer. It will have the name you gave the project when you created it and a Manifest entry under it (i.e. ExcelWebAddin1).
6. Press F4 to view the **Properties** window for the project (if it is not already visible).
7. Change the **Web Project** property to the name of ASP .NET Core web project you just added to the solution.
8. In the **Solution Explorer** open the manifest.xml file under the Office Add-in project (the project only contains a manifest file). Copy and save the <Id> that has the GUID for your project.
8. Copy the contents of the manifest.xml file in the ASP .NET Core project and replace all of the contents in the maninfest file in the Web Add-in project (i.e. ExcelWebAddin1Manifest.xml) with it.
9. Press F5 to debug the Office Web Add-in project.

> **Note:** npm install should run and install the packages prior to building the ASP.NET Core web project but you may need to watch the output window for errors.  If errors occur, please try running npm install from the ./ClientApp folder.

## Version history

Version  | Date | Comments
---------| -----| --------
1.0 | September 10, 2019 | Initial Release

### Disclaimer ###

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
