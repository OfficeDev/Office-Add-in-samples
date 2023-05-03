# Auto-opening a Taskpane with a Document

## Overview
Add-ins with commands enable developers to extend the Office UI, such as create Office Ribbon buttons. When users click your command an action, such as showing a pane is executed. Some scenarios, however, require a pane to be automatically opened with certain documents without explicit user interaction. The auto-open taskpane feature, part of AddInCommands 1.1 allows you automatically open a pane in those scenarios. 


### How is this different from “inserting” a taskpane? 
Add-ins without commands, such as when users run add-ins in Office 2013, are by default inserted into the document, making it stick to that document without any explicit user or developer intent to do so. As a result,  any other user that opens that document is prompted to install the add-in and the pane opens.  The challenge with this model is that in many cases users don’t want the add-in to stick with the document, all they want is to use a particular add-in. For example, a student using a dictionary add-in doesn’t want his class-mates or teachers to open that same document and be prompted to install a dictionary add-in.  

In contrast, the auto-open pane feature is driven by the add-in developer as opposed to being a default behavior.  This means that developers explicitly opt-in, or provide affordances for users to opt-in,  to use the feature on specific add-ins and specific documents that require it. 

## Support and availability
The auto-open feature is currently in **developer preview** and it is only supported in the following environments:

Products

- Word
- Excel
- PowerPoint

Versions

- Office for Windows Desktop. Build 16.0.811.1000+ (Insiders Fast)
- Office for Mac. Build 15.34(170414)+   (Insiders Fast)
- Office Online 

Please note that for Windows and Mac you need to be on **[Insiders Fast](https://products.office.com/en-us/office-insider?tab=tab-1)** and have updates turned on to have access to this feature during the preview. In other words, even if you have a more recent build the feature won't work if you are not part of Insiders Fast. 

## Best Practices
### Do


- Use add-in commands along with auto-open to make users more efficient using your add-in. Sample scenarios:
	- The document needs the add-in to function properly. For example, a spreadsheet with stock values that are periodically refreshed by an add-in, the add-in needs to auto-open with the document otherwise the values of the stocks would remain stale. 
	- The user is likely to always use the add-in for that document. For example, an add-in that helps users fill-in or change data inside a document by pulling information from a backed system. 


- Provide users control to turn on/off if a pane in your add-in should auto-open. For example a UI affordance in case users no longer want your add-in to auto-open a pane. 
- Use requirement set detection to determine if this feature is available and provide a fallback behavior if it isn’t.

### Don’t


- Abuse this feature as means to artificially increase usage of your add-in. If it doesn’t make sense for your add-in to auto-open with certain documents you should not use this feature; it will annoy users and your add-in might get rejected from the Office store if Microsoft detects an abuse. 


- Use this feature as a generic “pinning” capability for panes. This feature doesn’t allow you to pin a specific pane in place, instead, it  lets you to designate ONE pane per add-in to auto-open along with a document. If your add-in has multiple panes, you can only designate one to auto-open. 

## Implementation
There are 2 main steps required to use this feature: Specifying the pane to be opened and tagging documents to trigger auto-open.


#### 1.-Specifying the pane to open
You indicate which pane will be auto-opened on your manifest by setting its TaskpaneId value to the well-known value of ***Office.AutoShowTaskpaneWithDocument***. You can only set ONE pane to be opened; if you set this value multiple times, the second pane will be ignored. 
          
    <Action xsi:type="ShowTaskpane">
         <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
         <SourceLocation resid="Contoso.Taskpane.Url" />
    </Action>
     

#### 2.-Tagging a document to trigger auto-open
To trigger auto-open a document must be appropriately tagged. Documents that are not tagged will not trigger auto-open. You can tag a document in 2 main ways, choose the one that makes the most sense for your scenario:


##### Client side
Set the ***Office.AutoShowTaskpaneWithDocument*** setting to ***true*** using Office.js. Use this method if you need to tag the document as part of your add-in interaction (E.g. as soon as the user creates a binding, or clicks on a UI affordance on your add-in to indicate they want the pane to auto-open) 

    Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
    Office.context.document.settings.saveAsync();

##### Via OpenXML
You can use OpenXML to create/modify a document and add the appropriate tag. A sample to show how to do this using OpenXML is in the works, in the meantime, here is the snippet that shows how the setting looks like inside a webextension part. 

    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="[ADDIN GUID PER MANIFEST]">
      <we:reference id="[GUID]" version="[your add-in version]" store="[Pointer to store]" storeType="[StoreType]"/>
      <we:alternateReferences/>
      <we:properties>
    	<we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
      </we:properties>
      <we:bindings/>
      <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
    </we:webextension>

An easy way to figure out the XML you need to write is to first run your add-in and using the client side technique to write the value, then save the document and inspect the XML that is generated. 

### Add-in installation requirement##
It is important to highlight that the **pane that you designate will only automatically open IF** , by the time the user opens the document, your **add-in is already installed on the users device**.  If users open a document and they do not have your add-in already installed then nothing will happen, the setting will be ignored. 

If you require to also distribute the add-in with the document, so that users are prompted to install it, you also need to set the pane visibility property to 1, you can only do this via OpenXML.

## Samples
The folder in this repo contains a simple example that shows you how to specify what pane to open on your add-in manifest as well as how to tag a document via Office.js. Additional samples are in the works. 

![](http://i.imgur.com/JtHwr47.png)
