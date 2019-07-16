# Dynamic DPI code samples

## Summary

Many computer and display configurations now support high DPI (dots-per-inch) resolutions, and can connect multiple monitors with different sizes and pixel densities. This requires applications to adjust when the user moves the app to a monitor with a different DPI, or changes the zoom level. Applications that don’t support DPI scaling might look fine on low DPI monitors, but will look stretched and blurry when shown on a high DPI monitor.

The code samples included here will help you with handling DPI changes in your code for VSTO and COM Add-in projects. More information about the code samples and handling DPI can be found in the accompnying article: [Handle high DPI and DPI scaling in your Office solution](https://docs.microsoft.com/office/client-developer/ddpi/handle-high-dpi-and-dpi-scaling-in-your-office-solution)

## Applies to

- VSTO Add-ins
- Custom task panes
- COM Add-ins
- ActiveX controls

## Prerequisites

- Visual Studio 2017 or later
    - latest version of Windows SDK (not you may have to retarget steps to do so)
    - Visual Basic (for the VB one)<TBD>
- An Office 365 account which you can get by joining the Office 365 Developer Program that includes a free 1 year subscription to Office 365.
- Many of the samples use the **Developer** tab in Microsoft Excel. If you haven't enabled the **Developer** tab, follow these instructions in the article [Show the Developer tab](https://support.office.com/article/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45)

## Solution

Solution | Author(s)
---------|----------
Dynamic DPI samples | Shawn McDowell (Microsoft)

## Version history

Version  | Date | Comments
---------| -----| --------
1.0  | July 15 2019 | Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

----------

## COMAddinCPP

This sample shows how to handle DPI changes in a COM Add-in using C++.

### To run the sample

1. Clone or download this repo.
2. Run Visual Studio 2017 (or later) as administrator.
3. Open the **COMAddinCPP.sln** file.
4. On the menu bar choose **Build** > **Build Solution**.
    > **Note:** Set your build target to **Win32** or **x64** to match the version of Microsoft Excel you will run (32-bit or 64-bit).
5. Run Excel.
6. On the **Developer** tab, choose **COM Add-ins**.
7. Ensure **COM Add-in C++ with Custom TaskPane** is enabled (checked). Then choose **OK**.

A custom task pane will appear titled **COM Add-in C++**. It will show rectangle coordinates based on the current DPI. You can drag Excel to a monitor with a different DPI to see the coordinates update.

### DPI code

You can find more info about the DPI code in the ATLControl.cpp file.
<TBD: why does DPIHelper.cpp exist and is empty?>

## VSTO SharedAddin

This shows how to handle DPI changes in a VSTO Add-in. It contains folders for each Office host listed below:
- VSTO Excel Addin
- VSTO OutlookAddIn
- VSTO PowerPointAddIn
- VSTO VisioAddIn
- VSTO WordAddIn

### To run the sample

1. Clone or download this repo.
2. Run Visual Studio 2017 (or later) as administrator.
3. Open the **VSTOSharedAddin.sln** file.
4. On the menu bar choose **Build** > **Build Solution**.
    > **Note:** Set your build target to **Win32** or **x64** to match the version of Microsoft Excel you will run (32-bit or 64-bit).
5. Set one of the projects as the startup project. For example, right-click the **ExcelAddin1** project and choose **Set as StartUp Project**.
6. Choose **Start** (or press F5). The debugger will launch Excel and load the add-in.

The task pane for the VSTO Add-in will appear. You can drag Excel to a monitor with a different DPI to see displayed information change.

### DPI code

You can find more info about the DPI code in the DPIContextBlock.cs and DPIHelper.cs files.
<TBD: Can you change it from system aware?>
<TBD: Why won't it launch Word or other hosts in debug mode? Only Excel seems to launch correctly.>
<TBD: You have to create a test certificate for each project manually. is there an easier way to do this?>
<TBD: Following draw scenario draws incorrectly: 
1. set to per monitor aware
2.  click open top-level form
3. drag to window set to 150%>
<TBD: should the button1 label have a different name?>

## MFC ActiveX Not DPI Aware

This is an ActiveX control created from the MFC template that uses the DPI of the window to determine how to scale its contents.

<TBD: Can build and register, but cannot insert it. Get "Cannot insert object" error. >

## MFCApplicationDPIAware

This is an ActiveX control created from the MFC template that is dynamic DPI aware.

<TBD: Can we rename folder to "MFC ActiveX DPI Aware" since this appears to be counterpart of "ActiveX Not DPI Aware" folder?>



### To run the sample

1. Clone or download this repo.
2. Run Visual Studio 2017 (or later) as administrator.
3. Open the **MFCApplication1.sln** file.
4. On the menu bar choose **Build** > **Build Solution**.
    > **Note:** Set your build target to **Win32** or **x64** to match the version of Microsoft Excel you will run (32-bit or 64-bit).
5. Run Excel.
6. On the **Developer** > **Controls** tab, choose **Insert**. Then choose the **More Controls** icon which is in the **ActiveX Controls** section.
7. Choose **MFCActiveX Control**. Then choose **OK**.
5. Insert the control on the workbook by drawing a rectangle representing the size it should be.
6. You can right-click on the control an dchoose **MFCActiveX Control Object** > **Properties**.
7. On the **MFCActiveX Control Properties** box, enable the **Utilize Dynamic DPI Code** checkbox.

<TBD: is there a way to show how this works with an example text or something?>

### DPI code

You can find more info about the DPI code in the MFCApplication1.cpp file.

## Window Based ActiveX

This shows how to handle DPI changes in a window-based ActiveX control.

### To run the sample

1. Clone or download this repo.
2. Run Visual Studio 2017 (or later) as administrator.
3. Open the **MFCActiveX.sln** file.
4. On the menu bar choose **Build** > **Build Solution**.
    > **Note:** Set your build target to **Win32** or **x64** to match the version of Microsoft Excel you will run (32-bit or 64-bit).
5. Run Excel.
6. On the **Developer** > **Controls** tab, choose **Insert**. Then choose the **More Controls** icon which is in the **ActiveX Controls** section.
7. Choose **MFCActiveX Control**. Then choose **OK**.
5. Insert the control on the workbook by drawing a rectangle representing the size it should be.
6. Choose the Design Mode button to turn off design mode.

The control will display a pie chart and some other information. You can drag Excel to a monitor with different DPI settings to see how the control redraws.

### DPI code

This is an MFC window-based ActiveX control that supports dynamic DPI on the WM_SIZE event.

You can find more info about the DPI code in the MFCActiveXCtrl.cpp file.

## Windowles ActiveX

This shows how to handle DPI changes in a windowless ActiveX control.

### To run the sample

1. Clone or download this repo.
2. Run Visual Studio 2017 (or later) as administrator.
3. Open the **ODActiveX.sln** file.
4. On the menu bar choose **Build** > **Build Solution**.
    > **Note:** Set your build target to **Win32** or **x64** to match the version of Microsoft Excel you will run (32-bit or 64-bit).
5. Run Excel.
6. On the **Developer** > **Controls** tab, choose **Insert**. Then choose the **More Controls** icon which is in the **ActiveX Controls** section.
7. Choose **ODActiveX Control**. Then choose **OK**.
5. Insert the control on the workbook by drawing a rectangle representing the size it should be.
6. Choose the Design Mode button to turn off design mode.

The control will display some DPI information. You can drag Excel to a monitor with different DPI settings to see how the control redraws.

### DPI code

This is an MFC Windowless ActiveX control that supports dynamic DPI on WM_PAINT. It gets the HWND of the host window from HDC. <TBD: does this need more info?>


You can find more info about the DPI code in the ODActiveXCtrl.cpp file.

<img src="https://telemetry.sharepointpnp.com/officedev/samples/dynamic-dpi" />
