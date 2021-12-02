---
page_type: sample
products:
- office-excel
- office-word
- office-powerpoint
- office-365
languages:
- typescript
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 11/6/2015 1:25:00 PM
description: "The examples in this repo show you how to use add-in commands in Excel, Word and PowerPoint add-ins."
---

# Office Add-in Commands Samples 

## Overview
Add-in commands enable developers to extend the Office user interface such as the Office Ribbon to create awesome, efficient to use add-ins. Watch this [channel9 video](https://channel9.msdn.com/events/Build/2016/P551) for a complete overview. The examples in this repo show you how to use add-in commands in Excel, Word and PowerPoint add-ins. If you are looking for information about commands for **Outlook** head to [http://dev.outlook.com](http://dev.outlook.com)
 
Here is how the samples look when running: 

### Custom Tab (Simple Example)
![](https://i.imgur.com/HRCbRFO.png)

### Excel
![](http://i.imgur.com/OsRIk5E.png)

### Word
Existing Tab
![](http://i.imgur.com/wrA6R3T.png)

### PowerPoint
![](http://i.imgur.com/jwkkNsQ.png)


## Quick Start
### Step 1. Setup your environment


- **Office Desktop**: Ensure that you have the latest version of Office installed. Add-in commands require build 16.0.6769.0000 or higher (**16.0.6868.0000** recommended). Learn how to [Install the latest version of Office applications](http://aka.ms/latestoffice). 
 
- **Office Online**: There is no additional setup. Please note that support for commands in Office Online for work/school accounts is in preview.

- **Office for Mac**: Ensure that you have build 15.33+

### Step 2. Create and validate your manifest
We strongly recommend you to use one of our sample manifests as a starting point, the [Simple example](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Simple) is a good one to get going. Once you make it work then you can start making small modifications and test your changes often. If you make modifications, use the [Manifest reference](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests?tabs=tabid-1) as a guide. You can also validate your xml using the [Office Add-in Validator](https://github.com/OfficeDev/office-addin-validator) **. For Office Windows clients you can also use [Runtime Logging](Tools/RuntimeLogging.md) to debug your manifest.

You can also use the latest Visual Studio Tools to create and debug your add-in. See next step. 

### Step 3. Deploy add-in manifest and test the add-in
To test your add-in you must register it with Office. Two methods are currently supported:
#### Sideload directly to the client
- **Office Desktop**. Sideload your add-in via a [network share](https://msdn.microsoft.com/EN-US/library/office/fp123503.aspx). 
	- Once sideloaded, go to `Insert>My Add-ins>Shared Folder` and click the `Refresh` button to ensure the Add-in shows. Do this any time you need to refresh your Ribbon.
- **Office Online**. Open the Add-ins dialog via `Insert>Office Add-ins` then select `[Manage My Add-ins]>Upload My Add-in` and upload the manifest file you want to test. To remove a sideloaded add-in you have to [clear your HTML LocalStorage](http://superuser.com/questions/519628/clear-html5-local-storage-on-a-specific-page) 

- **Office for Mac**. [Sideload your add-in on the Mac](https://dev.office.com/docs/add-ins/testing/sideload-an-office-add-in-on-ipad-and-mac)
	- Once sideloaded, goto `Insert>Office Add-ins` and click on the add-in to install it.

#### Visual Studio F5
- Make sure you have at least version 16.0.6868.0000 of Office for Windows installed. 
- Make sure you have the latest [Visual Studio tools](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx). 

Once you have the latest, the new VS templates include support for add-in commands. You can also deploy your add-ins to Windows Desktop clients using F5. 

## Documentation
- [FAQ](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/FAQ.md)
- [Manifest reference](http://dev.office.com/docs/add-ins/overview/add-in-manifests)

## Join the Microsoft 365 Developer Program
Get a free sandbox, tools, and other resources you need to build solutions for the Microsoft 365 platform.
- [Free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) Get a free, renewable 90-day Microsoft 365 E5 developer subscription.
- [Sample data packs](https://developer.microsoft.com/microsoft-365/dev-program#Sample) Automatically configure your sandbox by installing user data and content to help you build your solutions.
- [Access to experts](https://developer.microsoft.com/microsoft-365/dev-program#Experts) Access community events to learn from Microsoft 365 experts.
- [Personalized recommendations](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations) Find developer resources quickly from your personalized dashboard.


        
    


This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
