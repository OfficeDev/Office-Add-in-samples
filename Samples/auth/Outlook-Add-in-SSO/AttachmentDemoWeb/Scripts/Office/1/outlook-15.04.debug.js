/* Outlook specific API library */
/* Version: 15.0.4927.1000 */
/* Update: 2 */
/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

/*
	Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
*/

Type.registerNamespace("Microsoft.Office.WebExtension.MailboxEnums");
Microsoft.Office.WebExtension.MailboxEnums.EntityType={
	MeetingSuggestion: "meetingSuggestion",
	TaskSuggestion: "taskSuggestion",
	Address: "address",
	EmailAddress: "emailAddress",
	Url: "url",
	PhoneNumber: "phoneNumber",
	Contact: "contact",
	FlightReservations: "flightReservations",
	ParcelDeliveries: "parcelDeliveries"
};
Microsoft.Office.WebExtension.MailboxEnums.ItemType={
	Message: "message",
	Appointment: "appointment"
};
Microsoft.Office.WebExtension.MailboxEnums.ResponseType={
	None: "none",
	Organizer: "organizer",
	Tentative: "tentative",
	Accepted: "accepted",
	Declined: "declined"
};
Microsoft.Office.WebExtension.MailboxEnums.RecipientType={
	Other: "other",
	DistributionList: "distributionList",
	User: "user",
	ExternalUser: "externalUser"
};
Microsoft.Office.WebExtension.MailboxEnums.AttachmentType={
	File: "file",
	Item: "item",
	Cloud: "cloud"
};
Microsoft.Office.WebExtension.MailboxEnums.BodyType={
	Text: "text",
	Html: "html"
};
Microsoft.Office.WebExtension.MailboxEnums.ItemNotificationMessageType={
	ProgressIndicator: "progressIndicator",
	InformationalMessage: "informationalMessage",
	ErrorMessage: "errorMessage"
};
Microsoft.Office.WebExtension.CoercionType={
	Text: "text",
	Html: "html"
};
Microsoft.Office.WebExtension.MailboxEnums.UserProfileType={
	Office365: "office365",
	OutlookCom: "outlookCom",
	Enterprise: "enterprise"
};
Microsoft.Office.WebExtension.MailboxEnums.RestVersion={
	v1_0: "v1.0",
	v2_0: "v2.0",
	Beta: "beta"
};
Type.registerNamespace("OSF.DDA");
var OSF=window.OSF || {};
OSF.DDA=OSF.DDA || {};
window["OSF"]["DDA"]["OutlookAppOm"]=OSF.DDA.OutlookAppOm=function(officeAppContext, targetWindow, appReadyCallback)
{
	this.$$d__callAppReadyCallback$p$0=Function.createDelegate(this,this._callAppReadyCallback$p$0);
	this.$$d_displayContactCardAsync=Function.createDelegate(this,this.displayContactCardAsync);
	this.$$d_displayNewMessageFormApi=Function.createDelegate(this,this.displayNewMessageFormApi);
	this.$$d__displayNewAppointmentFormApi$p$0=Function.createDelegate(this,this._displayNewAppointmentFormApi$p$0);
	this.$$d_windowOpenOverrideHandler=Function.createDelegate(this,this.windowOpenOverrideHandler);
	this.$$d__getEwsUrl$p$0=Function.createDelegate(this,this._getEwsUrl$p$0);
	this.$$d__getDiagnostics$p$0=Function.createDelegate(this,this._getDiagnostics$p$0);
	this.$$d__getUserProfile$p$0=Function.createDelegate(this,this._getUserProfile$p$0);
	this.$$d__getItem$p$0=Function.createDelegate(this,this._getItem$p$0);
	this.$$d__getInitialDataResponseHandler$p$0=Function.createDelegate(this,this._getInitialDataResponseHandler$p$0);
	window["OSF"]["DDA"]["OutlookAppOm"]._instance$p=this;
	this._officeAppContext$p$0=officeAppContext;
	this._appReadyCallback$p$0=appReadyCallback;
	var $$t_4=this;
	var stringLoadedCallback=function()
		{
			if(appReadyCallback)
				if(!officeAppContext["get_isDialog"]())
					$$t_4.invokeHostMethod(1,null,$$t_4.$$d__getInitialDataResponseHandler$p$0);
				else
					window.setTimeout($$t_4.$$d__callAppReadyCallback$p$0,0)	
		};
	if(this._areStringsLoaded$p$0())
		stringLoadedCallback();
	else
		this._loadLocalizedScript$p$0(stringLoadedCallback)
};
window["OSF"]["DDA"]["OutlookAppOm"]._throwOnPropertyAccessForRestrictedPermission$i=function(currentPermissionLevel)
{
	if(!currentPermissionLevel)
		throw Error.create(window["_u"]["ExtensibilityStrings"]["l_ElevatedPermissionNeeded_Text"]);
};
window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i=function(value, minValue, maxValue, argumentName)
{
	if(value < minValue || value > maxValue)
		throw Error.argumentOutOfRange(argumentName);
};
window["OSF"]["DDA"]["OutlookAppOm"]._getHtmlBody$p=function(data)
{
	var htmlBody="";
	if("htmlBody" in data)
	{
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidHtmlBody$p(data["htmlBody"]);
		htmlBody=data["htmlBody"]
	}
	return htmlBody
};
window["OSF"]["DDA"]["OutlookAppOm"]._getAttachments$p=function(data)
{
	var attachments=[];
	if("attachments" in data)
	{
		attachments=data["attachments"];
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentsArray$p(attachments)
	}
	return attachments
};
window["OSF"]["DDA"]["OutlookAppOm"]._getOptionsAndCallback$p=function(data)
{
	var args=[];
	if("options" in data)
		args[0]=data["options"];
	if("callback" in data)
		args[args.length]=data["callback"];
	return args
};
window["OSF"]["DDA"]["OutlookAppOm"]._createAttachmentsDataForHost$p=function(attachments)
{
	var attachmentsData=new Array(0);
	if(Array.isInstanceOfType(attachments))
		for(var i=0; i < attachments["length"]; i++)
			if(Object.isInstanceOfType(attachments[i]))
			{
				var attachment=attachments[i];
				window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachment$p(attachment);
				attachmentsData[i]=window["OSF"]["DDA"]["OutlookAppOm"]._createAttachmentData$p(attachment)
			}
			else
				throw Error.argument("attachments");
	return attachmentsData
};
window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidHtmlBody$p=function(htmlBody)
{
	if(!String.isInstanceOfType(htmlBody))
		throw Error.argument("htmlBody");
	if($h.ScriptHelpers.isNullOrUndefined(htmlBody))
		throw Error.argument("htmlBody");
	window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(htmlBody["length"],0,window["OSF"]["DDA"]["OutlookAppOm"].maxBodyLength,"htmlBody")
};
window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentsArray$p=function(attachments)
{
	if(!Array.isInstanceOfType(attachments))
		throw Error.argument("attachments");
};
window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachment$p=function(attachment)
{
	if(!Object.isInstanceOfType(attachment))
		throw Error.argument("attachments");
	if(!("type" in attachment) || !("name" in attachment))
		throw Error.argument("attachments");
	if(!("url" in attachment || "itemId" in attachment))
		throw Error.argument("attachments");
};
window["OSF"]["DDA"]["OutlookAppOm"]._createAttachmentData$p=function(attachment)
{
	var attachmentData=null;
	if(attachment["type"]==="file")
	{
		var url=attachment["url"];
		var name=attachment["name"];
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentUrlOrName$p(url,name);
		attachmentData=window["OSF"]["DDA"]["OutlookAppOm"]._createFileAttachmentData$p(url,name)
	}
	else if(attachment["type"]==="item")
	{
		var itemId=window["OSF"]["DDA"]["OutlookAppOm"].getItemIdBasedOnHost(attachment["itemId"]);
		var name=attachment["name"];
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentItemIdOrName$p(itemId,name);
		attachmentData=window["OSF"]["DDA"]["OutlookAppOm"]._createItemAttachmentData$p(itemId,name)
	}
	else
		throw Error.argument("attachments");
	return attachmentData
};
window["OSF"]["DDA"]["OutlookAppOm"]._createFileAttachmentData$p=function(url, name)
{
	return["file",name,url]
};
window["OSF"]["DDA"]["OutlookAppOm"]._createItemAttachmentData$p=function(itemId, name)
{
	return["item",name,itemId]
};
window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentUrlOrName$p=function(url, name)
{
	if(!String.isInstanceOfType(url) || !String.isInstanceOfType(name))
		throw Error.argument("attachments");
	if(url["length"] > 2048)
		throw Error.argumentOutOfRange("attachments",url["length"],window["_u"]["ExtensibilityStrings"]["l_AttachmentUrlTooLong_Text"]);
	window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentName$p(name)
};
window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentItemIdOrName$p=function(itemId, name)
{
	if(!String.isInstanceOfType(itemId) || !String.isInstanceOfType(name))
		throw Error.argument("attachments");
	if(itemId["length"] > 200)
		throw Error.argumentOutOfRange("attachments",itemId["length"],window["_u"]["ExtensibilityStrings"]["l_AttachmentItemIdTooLong_Text"]);
	window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentName$p(name)
};
window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentName$p=function(name)
{
	if(name["length"] > 255)
		throw Error.argumentOutOfRange("attachments",name["length"],window["_u"]["ExtensibilityStrings"]["l_AttachmentNameTooLong_Text"]);
};
window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidRestVersion$p=function(restVersion)
{
	if(!restVersion)
		throw Error.argumentNull("restVersion");
	if(restVersion !==window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RestVersion"]["v1_0"] && restVersion !==window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RestVersion"]["v2_0"] && restVersion !==window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RestVersion"]["Beta"])
		throw Error.argument("restVersion");
};
window["OSF"]["DDA"]["OutlookAppOm"].getItemIdBasedOnHost=function(itemId)
{
	if(window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._initialData$p$0 && window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._initialData$p$0.get__isRestIdSupported$i$0())
		return window["OSF"]["DDA"]["OutlookAppOm"]._instance$p["convertToRestId"](itemId,window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RestVersion"]["v1_0"]);
	return window["OSF"]["DDA"]["OutlookAppOm"]._instance$p["convertToEwsId"](itemId,window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RestVersion"]["v1_0"])
};
window["OSF"]["DDA"]["OutlookAppOm"]._throwOnArgumentType$p=function(value, expectedType, argumentName)
{
	if(Object["getType"](value) !==expectedType)
		throw Error.argumentType(argumentName);
};
window["OSF"]["DDA"]["OutlookAppOm"]._validateOptionalStringParameter$p=function(value, minLength, maxLength, name)
{
	if($h.ScriptHelpers.isNullOrUndefined(value))
		return;
	window["OSF"]["DDA"]["OutlookAppOm"]._throwOnArgumentType$p(value,String,name);
	var stringValue=value;
	window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(stringValue["length"],minLength,maxLength,name)
};
window["OSF"]["DDA"]["OutlookAppOm"]._convertToOutlookParameters$p=function(dispid, data)
{
	var executeParameters=null;
	switch(dispid)
	{
		case 1:
		case 2:
		case 12:
		case 3:
		case 14:
		case 18:
		case 26:
		case 32:
		case 41:
		case 34:
			break;
		case 4:
			var jsonProperty=window["JSON"]["stringify"](data["customProperties"]);
			executeParameters=[jsonProperty];
			break;
		case 5:
			executeParameters=[data["body"]];
			break;
		case 8:
		case 9:
			executeParameters=[data["itemId"]];
			break;
		case 7:
			executeParameters=[window["OSF"]["DDA"]["OutlookAppOm"]._convertRecipientArrayParameterForOutlookForDisplayApi$p(data["requiredAttendees"]),window["OSF"]["DDA"]["OutlookAppOm"]._convertRecipientArrayParameterForOutlookForDisplayApi$p(data["optionalAttendees"]),data["start"],data["end"],data["location"],window["OSF"]["DDA"]["OutlookAppOm"]._convertRecipientArrayParameterForOutlookForDisplayApi$p(data["resources"]),data["subject"],data["body"]];
			break;
		case 44:
			executeParameters=[window["OSF"]["DDA"]["OutlookAppOm"]._convertRecipientArrayParameterForOutlookForDisplayApi$p(data["toRecipients"]),window["OSF"]["DDA"]["OutlookAppOm"]._convertRecipientArrayParameterForOutlookForDisplayApi$p(data["ccRecipients"]),window["OSF"]["DDA"]["OutlookAppOm"]._convertRecipientArrayParameterForOutlookForDisplayApi$p(data["bccRecipients"]),data["subject"],data["htmlBody"],data["attachments"]];
			break;
		case 43:
			executeParameters=[data["ewsIdOrEmail"]];
			break;
		case 40:
			executeParameters=[data["extensionId"],data["consentState"]];
			break;
		case 11:
		case 10:
			executeParameters=[data["htmlBody"]];
			break;
		case 31:
		case 30:
			executeParameters=[data["htmlBody"],data["attachments"]];
			break;
		case 23:
		case 13:
		case 38:
		case 29:
			executeParameters=[data["data"],data["coercionType"]];
			break;
		case 37:
		case 28:
			executeParameters=[data["coercionType"]];
			break;
		case 17:
			executeParameters=[data["subject"]];
			break;
		case 15:
			executeParameters=[data["recipientField"]];
			break;
		case 22:
		case 21:
			executeParameters=[data["recipientField"],window["OSF"]["DDA"]["OutlookAppOm"]._convertComposeEmailDictionaryParameterForSetApi$p(data["recipientArray"])];
			break;
		case 19:
			executeParameters=[data["itemId"],data["name"]];
			break;
		case 16:
			executeParameters=[data["uri"],data["name"]];
			break;
		case 20:
			executeParameters=[data["attachmentIndex"]];
			break;
		case 25:
			executeParameters=[data["TimeProperty"],data["time"]];
			break;
		case 24:
			executeParameters=[data["TimeProperty"]];
			break;
		case 27:
			executeParameters=[data["location"]];
			break;
		case 33:
		case 35:
			executeParameters=[data["key"],data["type"],data["persistent"],data["message"],data["icon"]];
			break;
		case 36:
			executeParameters=[data["key"]];
			break;
		default:
			Sys.Debug.fail("Unexpected method dispid");
			break
	}
	return executeParameters
};
window["OSF"]["DDA"]["OutlookAppOm"]._convertRecipientArrayParameterForOutlookForDisplayApi$p=function(array)
{
	return array ? array["join"](";") : null
};
window["OSF"]["DDA"]["OutlookAppOm"]._convertComposeEmailDictionaryParameterForSetApi$p=function(recipients)
{
	if(!recipients)
		return null;
	var results=new Array(recipients.length);
	for(var i=0; i < recipients.length; i++)
		results[i]=[recipients[i]["address"],recipients[i]["name"]];
	return results
};
window["OSF"]["DDA"]["OutlookAppOm"]._validateAndNormalizeRecipientEmails$p=function(emailset, name)
{
	if($h.ScriptHelpers.isNullOrUndefined(emailset))
		return null;
	window["OSF"]["DDA"]["OutlookAppOm"]._throwOnArgumentType$p(emailset,Array,name);
	var originalAttendees=emailset;
	var updatedAttendees=null;
	var normalizationNeeded=false;
	window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(originalAttendees["length"],0,window["OSF"]["DDA"]["OutlookAppOm"]._maxRecipients$p,String.format("{0}.length",name));
	for(var i=0; i < originalAttendees["length"]; i++)
		if($h.EmailAddressDetails.isInstanceOfType(originalAttendees[i]))
		{
			normalizationNeeded=true;
			break
		}
	if(normalizationNeeded)
		updatedAttendees=[];
	for(var i=0; i < originalAttendees["length"]; i++)
		if(normalizationNeeded)
		{
			updatedAttendees[i]=$h.EmailAddressDetails.isInstanceOfType(originalAttendees[i]) ? originalAttendees[i]["emailAddress"] : originalAttendees[i];
			window["OSF"]["DDA"]["OutlookAppOm"]._throwOnArgumentType$p(updatedAttendees[i],String,String.format("{0}[{1}]",name,i))
		}
		else
			window["OSF"]["DDA"]["OutlookAppOm"]._throwOnArgumentType$p(originalAttendees[i],String,String.format("{0}[{1}]",name,i));
	return updatedAttendees
};
OSF.DDA.OutlookAppOm.prototype={
	_initialData$p$0: null,
	_item$p$0: null,
	_userProfile$p$0: null,
	_diagnostics$p$0: null,
	_officeAppContext$p$0: null,
	_appReadyCallback$p$0: null,
	_clientEndPoint$p$0: null,
	get_clientEndPoint: function()
	{
		if(!this._clientEndPoint$p$0)
			this._clientEndPoint$p$0=window["OSF"]["_OfficeAppFactory"]["getClientEndPoint"]();
		return this._clientEndPoint$p$0
	},
	set_clientEndPoint: function(value)
	{
		this._clientEndPoint$p$0=value;
		return value
	},
	get_initialData: function()
	{
		return this._initialData$p$0
	},
	get__appName$i$0: function()
	{
		return this._officeAppContext$p$0["get_appName"]()
	},
	initialize: function(initialData)
	{
		var ItemTypeKey="itemType";
		this._initialData$p$0=new $h.InitialData(initialData);
		if(1===initialData[ItemTypeKey])
			this._item$p$0=new $h.Message(this._initialData$p$0);
		else if(3===initialData[ItemTypeKey])
			this._item$p$0=new $h.MeetingRequest(this._initialData$p$0);
		else if(2===initialData[ItemTypeKey])
			this._item$p$0=new $h.Appointment(this._initialData$p$0);
		else if(4===initialData[ItemTypeKey])
			this._item$p$0=new $h.MessageCompose(this._initialData$p$0);
		else if(5===initialData[ItemTypeKey])
			this._item$p$0=new $h.AppointmentCompose(this._initialData$p$0);
		else
			Sys.Debug.trace("Unexpected item type was received from the host.");
		this._userProfile$p$0=new $h.UserProfile(this._initialData$p$0);
		this._diagnostics$p$0=new $h.Diagnostics(this._initialData$p$0,this._officeAppContext$p$0["get_appName"]());
		this._initializeMethods$p$0();
		$h.InitialData._defineReadOnlyProperty$i(this,"item",this.$$d__getItem$p$0);
		$h.InitialData._defineReadOnlyProperty$i(this,"userProfile",this.$$d__getUserProfile$p$0);
		$h.InitialData._defineReadOnlyProperty$i(this,"diagnostics",this.$$d__getDiagnostics$p$0);
		$h.InitialData._defineReadOnlyProperty$i(this,"ewsUrl",this.$$d__getEwsUrl$p$0);
		if(window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.get__appName$i$0()===64)
			if(this._initialData$p$0.get__overrideWindowOpen$i$0())
				window.open=this.$$d_windowOpenOverrideHandler
	},
	windowOpenOverrideHandler: function(url, targetName, features, replace)
	{
		this.invokeHostMethod(403,{launchUrl: url},null)
	},
	makeEwsRequestAsync: function(data)
	{
		var args=[];
		for(var $$pai_5=1; $$pai_5 < arguments.length;++$$pai_5)
			args[$$pai_5 - 1]=arguments[$$pai_5];
		if($h.ScriptHelpers.isNullOrUndefined(data))
			throw Error.argumentNull("data");
		if(data["length"] > window["OSF"]["DDA"]["OutlookAppOm"]._maxEwsRequestSize$p)
			throw Error.argument("data",window["_u"]["ExtensibilityStrings"]["l_EwsRequestOversized_Text"]);
		this._throwOnMethodCallForInsufficientPermission$i$0(3,"makeEwsRequestAsync");
		var parameters=$h.CommonParameters.parse(args,true,true);
		var ewsRequest=new $h.EwsRequest(parameters._asyncContext$p$0);
		var $$t_4=this;
		ewsRequest.onreadystatechange=function()
		{
			if(4===ewsRequest.get__requestState$i$1())
				parameters._callback$p$0(ewsRequest._asyncResult$p$0)
		};
		ewsRequest.send(data)
	},
	recordDataPoint: function(data)
	{
		if($h.ScriptHelpers.isNullOrUndefined(data))
			throw Error.argumentNull("data");
		this.invokeHostMethod(402,data,null)
	},
	recordTrace: function(data)
	{
		if($h.ScriptHelpers.isNullOrUndefined(data))
			throw Error.argumentNull("data");
		this.invokeHostMethod(401,data,null)
	},
	trackCtq: function(data)
	{
		if($h.ScriptHelpers.isNullOrUndefined(data))
			throw Error.argumentNull("data");
		this.invokeHostMethod(400,data,null)
	},
	convertToLocalClientTime: function(timeValue)
	{
		var date=new Date(timeValue["getTime"]());
		var offset=date["getTimezoneOffset"]() * -1;
		if(this._initialData$p$0 && this._initialData$p$0.get__timeZoneOffsets$i$0())
		{
			date["setUTCMinutes"](date["getUTCMinutes"]() - offset);
			offset=this._findOffset$p$0(date);
			date["setUTCMinutes"](date["getUTCMinutes"]()+offset)
		}
		var retValue=this._dateToDictionary$i$0(date);
		retValue["timezoneOffset"]=offset;
		return retValue
	},
	convertToUtcClientTime: function(input)
	{
		var retValue=this._dictionaryToDate$i$0(input);
		if(this._initialData$p$0 && this._initialData$p$0.get__timeZoneOffsets$i$0())
		{
			var offset=this._findOffset$p$0(retValue);
			retValue["setUTCMinutes"](retValue["getUTCMinutes"]() - offset);
			offset=!input["timezoneOffset"] ? retValue["getTimezoneOffset"]() * -1 : input["timezoneOffset"];
			retValue["setUTCMinutes"](retValue["getUTCMinutes"]()+offset)
		}
		return retValue
	},
	convertToRestId: function(itemId, restVersion)
	{
		if(!itemId)
			throw Error.argumentNull("itemId");
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidRestVersion$p(restVersion);
		return itemId["replace"](new RegExp("[/]","g"),"-")["replace"](new RegExp("[+]","g"),"_")
	},
	convertToEwsId: function(itemId, restVersion)
	{
		if(!itemId)
			throw Error.argumentNull("itemId");
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidRestVersion$p(restVersion);
		return itemId["replace"](new RegExp("[-]","g"),"/")["replace"](new RegExp("[_]","g"),"+")
	},
	getUserIdentityTokenAsync: function()
	{
		var args=[];
		for(var $$pai_2=0; $$pai_2 < arguments.length;++$$pai_2)
			args[$$pai_2]=arguments[$$pai_2];
		this._throwOnMethodCallForInsufficientPermission$i$0(1,"getUserIdentityTokenAsync");
		var parameters=$h.CommonParameters.parse(args,true,true);
		this._invokeGetTokenMethodAsync$p$0(2,"GetUserIdentityToken",parameters._callback$p$0,parameters._asyncContext$p$0)
	},
	getCallbackTokenAsync: function()
	{
		var args=[];
		for(var $$pai_2=0; $$pai_2 < arguments.length;++$$pai_2)
			args[$$pai_2]=arguments[$$pai_2];
		this._throwOnMethodCallForInsufficientPermission$i$0(1,"getCallbackTokenAsync");
		var parameters=$h.CommonParameters.parse(args,true,true);
		this._invokeGetTokenMethodAsync$p$0(12,"GetCallbackToken",parameters._callback$p$0,parameters._asyncContext$p$0)
	},
	displayMessageForm: function(itemId)
	{
		if($h.ScriptHelpers.isNullOrUndefined(itemId))
			throw Error.argumentNull("itemId");
		this.invokeHostMethod(8,{itemId: window["OSF"]["DDA"]["OutlookAppOm"].getItemIdBasedOnHost(itemId)},null)
	},
	displayAppointmentForm: function(itemId)
	{
		if($h.ScriptHelpers.isNullOrUndefined(itemId))
			throw Error.argumentNull("itemId");
		this.invokeHostMethod(9,{itemId: window["OSF"]["DDA"]["OutlookAppOm"].getItemIdBasedOnHost(itemId)},null)
	},
	RegisterConsentAsync: function(consentState)
	{
		if(consentState !==2 && consentState !==1 && consentState)
			throw Error.argumentOutOfRange("consentState");
		var parameters={};
		parameters["consentState"]=consentState["toString"]();
		parameters["extensionId"]=this["GetExtensionId"]();
		this.invokeHostMethod(40,parameters,null)
	},
	CloseApp: function()
	{
		this.invokeHostMethod(42,null,null)
	},
	GetIsRead: function()
	{
		return this._initialData$p$0.get__isRead$i$0()
	},
	GetEndNodeUrl: function()
	{
		return this._initialData$p$0.get__endNodeUrl$i$0()
	},
	GetConsentMetadata: function()
	{
		return this._initialData$p$0.get__consentMetadata$i$0()
	},
	GetEntryPointUrl: function()
	{
		return this._initialData$p$0.get__entryPointUrl$i$0()
	},
	GetMarketplaceContentMarket: function()
	{
		return this._initialData$p$0.get__marketplaceContentMarket$i$0()
	},
	GetMarketplaceAssetId: function()
	{
		return this._initialData$p$0.get__marketplaceAssetId$i$0()
	},
	GetExtensionId: function()
	{
		return this._initialData$p$0.get__extensionId$i$0()
	},
	createAsyncResult: function(value, errorCode, detailedErrorCode, userContext, errorMessage)
	{
		var initArgs={};
		var errorArgs=null;
		initArgs[window["OSF"]["DDA"]["AsyncResultEnum"]["Properties"]["Value"]]=value;
		initArgs[window["OSF"]["DDA"]["AsyncResultEnum"]["Properties"]["Context"]]=userContext;
		if(0 !==errorCode)
		{
			errorArgs={};
			var errorProperties=$h.OutlookErrorManager.getErrorArgs(detailedErrorCode);
			errorArgs[window["OSF"]["DDA"]["AsyncResultEnum"]["ErrorProperties"]["Name"]]=errorProperties["name"];
			errorArgs[window["OSF"]["DDA"]["AsyncResultEnum"]["ErrorProperties"]["Message"]]=!errorMessage ? errorProperties["message"] : errorMessage;
			errorArgs[window["OSF"]["DDA"]["AsyncResultEnum"]["ErrorProperties"]["Code"]]=detailedErrorCode
		}
		return new window["OSF"]["DDA"]["AsyncResult"](initArgs,errorArgs)
	},
	_throwOnMethodCallForInsufficientPermission$i$0: function(requiredPermissionLevel, methodName)
	{
		if(this._initialData$p$0._permissionLevel$p$0 < requiredPermissionLevel)
			throw Error.create(String.format(window["_u"]["ExtensibilityStrings"]["l_ElevatedPermissionNeededForMethod_Text"],methodName));
	},
	_displayReplyForm$i$0: function(obj)
	{
		this._displayReplyFormHelper$p$0(obj,false)
	},
	_displayReplyAllForm$i$0: function(obj)
	{
		this._displayReplyFormHelper$p$0(obj,true)
	},
	_displayReplyFormHelper$p$0: function(obj, isReplyAll)
	{
		if(String.isInstanceOfType(obj))
			this._doDisplayReplyForm$p$0(obj,isReplyAll);
		else if(Object.isInstanceOfType(obj) && Object.getTypeName(obj)==="Object")
			this._doDisplayReplyFormWithAttachments$p$0(obj,isReplyAll);
		else
			throw Error.argumentType();
	},
	_doDisplayReplyForm$p$0: function(htmlBody, isReplyAll)
	{
		if(!$h.ScriptHelpers.isNullOrUndefined(htmlBody))
			window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(htmlBody["length"],0,window["OSF"]["DDA"]["OutlookAppOm"].maxBodyLength,"htmlBody");
		this.invokeHostMethod(isReplyAll ? 11 : 10,{htmlBody: htmlBody},null)
	},
	_doDisplayReplyFormWithAttachments$p$0: function(data, isReplyAll)
	{
		var htmlBody=window["OSF"]["DDA"]["OutlookAppOm"]._getHtmlBody$p(data);
		var attachments=window["OSF"]["DDA"]["OutlookAppOm"]._getAttachments$p(data);
		var parameters=$h.CommonParameters.parse(window["OSF"]["DDA"]["OutlookAppOm"]._getOptionsAndCallback$p(data),false);
		var $$t_6=this;
		this._standardInvokeHostMethod$i$0(isReplyAll ? 31 : 30,{
			htmlBody: htmlBody,
			attachments: window["OSF"]["DDA"]["OutlookAppOm"]._createAttachmentsDataForHost$p(attachments)
		},function(rawInput)
		{
			return rawInput
		},parameters._asyncContext$p$0,parameters._callback$p$0)
	},
	_standardInvokeHostMethod$i$0: function(dispid, data, format, userContext, callback)
	{
		var $$t_B=this;
		this.invokeHostMethod(dispid,data,function(resultCode, response)
		{
			if(callback)
			{
				var asyncResult=null;
				if(Object.isInstanceOfType(response))
				{
					var responseDictionary=response;
					if("error" in responseDictionary || "data" in responseDictionary || "errorCode" in responseDictionary)
						if(!responseDictionary["error"])
						{
							var formattedData=format ? format(responseDictionary["data"]) : responseDictionary["data"];
							asyncResult=$$t_B.createAsyncResult(formattedData,0,0,userContext,null)
						}
						else
						{
							var errorCode=responseDictionary["errorCode"];
							asyncResult=$$t_B.createAsyncResult(null,1,errorCode,userContext,null)
						}
				}
				if(!asyncResult && resultCode)
					asyncResult=$$t_B.createAsyncResult(null,1,9002,userContext,null);
				callback(asyncResult)
			}
		})
	},
	invokeHostMethod: function(dispid, data, responseCallback)
	{
		if(64===this._officeAppContext$p$0["get_appName"]())
		{
			var args={ApiParams: data};
			args["MethodData"]={
				ControlId: window["OSF"]["_OfficeAppFactory"]["getId"](),
				DispatchId: dispid
			};
			if(dispid===1)
				this.get_clientEndPoint()["invoke"]("GetInitialData",responseCallback,args);
			else
				this.get_clientEndPoint()["invoke"]("ExecuteMethod",responseCallback,args)
		}
		else if(!this._isOwaOnlyMethod$p$0(dispid))
		{
			var executeParameters=window["OSF"]["DDA"]["OutlookAppOm"]._convertToOutlookParameters$p(dispid,data);
			var $$t_B=this;
			window["OSF"]["ClientHostController"]["execute"](dispid,executeParameters,function(nativeData, resultCode)
			{
				if(responseCallback)
				{
					var responseData=nativeData.toArray();
					var rawData=window["JSON"]["parse"](responseData[0]);
					if(Object.isInstanceOfType(rawData))
					{
						var deserializedData=rawData;
						if(responseData["length"] > 1 && responseData[1])
						{
							deserializedData["error"]=true;
							deserializedData["errorCode"]=responseData[1]
						}
						else
							deserializedData["error"]=false;
						responseCallback(resultCode,deserializedData)
					}
					else if(Number.isInstanceOfType(rawData))
					{
						var returnDict={};
						returnDict["error"]=true;
						returnDict["errorCode"]=rawData;
						responseCallback(resultCode,returnDict)
					}
					else
						throw Error.notImplemented("Return data type from host must be Dictionary or int");
				}
			})
		}
		else if(responseCallback)
			responseCallback(-2,null)
	},
	_dictionaryToDate$i$0: function(input)
	{
		var retValue=new Date(input["year"],input["month"],input["date"],input["hours"],input["minutes"],input["seconds"],!input["milliseconds"] ? 0 : input["milliseconds"]);
		if(window["isNaN"](retValue["getTime"]()))
			throw Error.format(window["_u"]["ExtensibilityStrings"]["l_InvalidDate_Text"]);
		return retValue
	},
	_dateToDictionary$i$0: function(input)
	{
		var retValue={};
		retValue["month"]=input["getMonth"]();
		retValue["date"]=input["getDate"]();
		retValue["year"]=input["getFullYear"]();
		retValue["hours"]=input["getHours"]();
		retValue["minutes"]=input["getMinutes"]();
		retValue["seconds"]=input["getSeconds"]();
		retValue["milliseconds"]=input["getMilliseconds"]();
		return retValue
	},
	_isOwaOnlyMethod$p$0: function(dispId)
	{
		switch(dispId)
		{
			case 402:
			case 401:
			case 400:
			case 403:
				return true;
			default:
				return false
		}
	},
	_displayNewAppointmentFormApi$p$0: function(parameters)
	{
		var normalizedRequiredAttendees=window["OSF"]["DDA"]["OutlookAppOm"]._validateAndNormalizeRecipientEmails$p(parameters["requiredAttendees"],"requiredAttendees");
		var normalizedOptionalAttendees=window["OSF"]["DDA"]["OutlookAppOm"]._validateAndNormalizeRecipientEmails$p(parameters["optionalAttendees"],"optionalAttendees");
		window["OSF"]["DDA"]["OutlookAppOm"]._validateOptionalStringParameter$p(parameters["location"],0,window["OSF"]["DDA"]["OutlookAppOm"]._maxLocationLength$p,"location");
		window["OSF"]["DDA"]["OutlookAppOm"]._validateOptionalStringParameter$p(parameters["body"],0,window["OSF"]["DDA"]["OutlookAppOm"].maxBodyLength,"body");
		window["OSF"]["DDA"]["OutlookAppOm"]._validateOptionalStringParameter$p(parameters["subject"],0,window["OSF"]["DDA"]["OutlookAppOm"]._maxSubjectLength$p,"subject");
		if(!$h.ScriptHelpers.isNullOrUndefined(parameters["start"]))
		{
			window["OSF"]["DDA"]["OutlookAppOm"]._throwOnArgumentType$p(parameters["start"],Date,"start");
			var startDateTime=parameters["start"];
			parameters["start"]=startDateTime["getTime"]();
			if(!$h.ScriptHelpers.isNullOrUndefined(parameters["end"]))
			{
				window["OSF"]["DDA"]["OutlookAppOm"]._throwOnArgumentType$p(parameters["end"],Date,"end");
				var endDateTime=parameters["end"];
				if(endDateTime < startDateTime)
					throw Error.argumentOutOfRange("end",endDateTime,window["_u"]["ExtensibilityStrings"]["l_InvalidEventDates_Text"]);
				parameters["end"]=endDateTime["getTime"]()
			}
		}
		var updatedParameters=null;
		if(normalizedRequiredAttendees || normalizedOptionalAttendees)
		{
			updatedParameters={};
			var $$dict_7=parameters;
			for(var $$key_8 in $$dict_7)
			{
				var entry={
						key: $$key_8,
						value: $$dict_7[$$key_8]
					};
				updatedParameters[entry.key]=entry.value
			}
			if(normalizedRequiredAttendees)
				updatedParameters["requiredAttendees"]=normalizedRequiredAttendees;
			if(normalizedOptionalAttendees)
				updatedParameters["optionalAttendees"]=normalizedOptionalAttendees
		}
		this.invokeHostMethod(7,updatedParameters || parameters,null)
	},
	displayNewMessageFormApi: function(parameters)
	{
		var updatedParameters={};
		if(parameters)
		{
			var normalizedToRecipients=window["OSF"]["DDA"]["OutlookAppOm"]._validateAndNormalizeRecipientEmails$p(parameters["toRecipients"],"toRecipients");
			var normalizedCcRecipients=window["OSF"]["DDA"]["OutlookAppOm"]._validateAndNormalizeRecipientEmails$p(parameters["ccRecipients"],"ccRecipients");
			var normalizedBccRecipients=window["OSF"]["DDA"]["OutlookAppOm"]._validateAndNormalizeRecipientEmails$p(parameters["bccRecipients"],"bccRecipients");
			window["OSF"]["DDA"]["OutlookAppOm"]._validateOptionalStringParameter$p(parameters["htmlBody"],0,window["OSF"]["DDA"]["OutlookAppOm"].maxBodyLength,"htmlBody");
			window["OSF"]["DDA"]["OutlookAppOm"]._validateOptionalStringParameter$p(parameters["subject"],0,window["OSF"]["DDA"]["OutlookAppOm"]._maxSubjectLength$p,"subject");
			var attachments=window["OSF"]["DDA"]["OutlookAppOm"]._getAttachments$p(parameters);
			var $$dict_7=parameters;
			for(var $$key_8 in $$dict_7)
			{
				var entry={
						key: $$key_8,
						value: $$dict_7[$$key_8]
					};
				updatedParameters[entry.key]=entry.value
			}
			if(normalizedToRecipients)
				updatedParameters["toRecipients"]=normalizedToRecipients;
			if(normalizedCcRecipients)
				updatedParameters["ccRecipients"]=normalizedCcRecipients;
			if(normalizedBccRecipients)
				updatedParameters["bccRecipients"]=normalizedBccRecipients;
			if(attachments)
				updatedParameters["attachments"]=window["OSF"]["DDA"]["OutlookAppOm"]._createAttachmentsDataForHost$p(attachments)
		}
		this.invokeHostMethod(44,updatedParameters || parameters,null)
	},
	displayContactCardAsync: function(ewsIdOrEmail, callback)
	{
		if($h.ScriptHelpers.isNullOrUndefined(ewsIdOrEmail))
			throw Error.argumentNull("ewsIdOrEmail");
		var asyncContext=null;
		var asyncFunctionCallback=null;
		if(callback)
		{
			var commonParameters=$h.CommonParameters.parse([callback],false);
			asyncContext=commonParameters._asyncContext$p$0;
			asyncFunctionCallback=commonParameters._callback$p$0
		}
		this._standardInvokeHostMethod$i$0(43,{ewsIdOrEmail: ewsIdOrEmail.trim()},null,asyncContext,asyncFunctionCallback)
	},
	_initializeMethods$p$0: function()
	{
		var currentInstance=this;
		if($h.Item.isInstanceOfType(this._item$p$0))
		{
			currentInstance["displayNewAppointmentForm"]=this.$$d__displayNewAppointmentFormApi$p$0;
			currentInstance["displayNewMessageForm"]=this.$$d_displayNewMessageFormApi;
			currentInstance["displayContactCardAsync"]=this.$$d_displayContactCardAsync
		}
	},
	_getInitialDataResponseHandler$p$0: function(resultCode, data)
	{
		if(resultCode)
			return;
		this.initialize(data);
		this["displayName"]="mailbox";
		window.setTimeout(this.$$d__callAppReadyCallback$p$0,0)
	},
	_callAppReadyCallback$p$0: function()
	{
		this._appReadyCallback$p$0()
	},
	_invokeGetTokenMethodAsync$p$0: function(outlookDispid, methodName, callback, userContext)
	{
		if($h.ScriptHelpers.isNullOrUndefined(callback))
			throw Error.argumentNull("callback");
		var $$t_8=this;
		this.invokeHostMethod(outlookDispid,null,function(resultCode, response)
		{
			var asyncResult;
			if(resultCode)
				asyncResult=$$t_8.createAsyncResult(null,1,9017,userContext,String.format(window["_u"]["ExtensibilityStrings"]["l_InternalProtocolError_Text"],resultCode));
			else
			{
				var responseDictionary=response;
				if(responseDictionary["wasSuccessful"])
					asyncResult=$$t_8.createAsyncResult(responseDictionary["token"],0,0,userContext,null);
				else
					asyncResult=$$t_8.createAsyncResult(null,1,responseDictionary["errorCode"],userContext,responseDictionary["errorMessage"])
			}
			callback(asyncResult)
		})
	},
	_getItem$p$0: function()
	{
		return this._item$p$0
	},
	_getUserProfile$p$0: function()
	{
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnPropertyAccessForRestrictedPermission$i(this._initialData$p$0._permissionLevel$p$0);
		return this._userProfile$p$0
	},
	_getDiagnostics$p$0: function()
	{
		return this._diagnostics$p$0
	},
	_getEwsUrl$p$0: function()
	{
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnPropertyAccessForRestrictedPermission$i(this._initialData$p$0._permissionLevel$p$0);
		return this._initialData$p$0.get__ewsUrl$i$0()
	},
	_findOffset$p$0: function(value)
	{
		var ranges=this._initialData$p$0.get__timeZoneOffsets$i$0();
		for(var r=0; r < ranges["length"]; r++)
		{
			var range=ranges[r];
			var start=window["parseInt"](range["start"]);
			var end=window["parseInt"](range["end"]);
			if(value["getTime"]() - start >=0 && value["getTime"]() - end < 0)
				return window["parseInt"](range["offset"])
		}
		throw Error.format(window["_u"]["ExtensibilityStrings"]["l_InvalidDate_Text"]);
	},
	_areStringsLoaded$p$0: function()
	{
		var stringsLoaded=false;
		try
		{
			stringsLoaded=!$h.ScriptHelpers.isNullOrUndefined(window["_u"]["ExtensibilityStrings"]["l_EwsRequestOversized_Text"])
		}
		catch($$e_1){}
		return stringsLoaded
	},
	_loadLocalizedScript$p$0: function(stringLoadedCallback)
	{
		var url=null;
		var baseUrl="";
		var scripts=document.getElementsByTagName("script");
		for(var i=scripts.length - 1; i >=0; i--)
		{
			var filename=null;
			var attributes=scripts[i].attributes;
			if(attributes)
			{
				var attribute=attributes.getNamedItem("src");
				if(attribute)
					filename=attribute.value;
				if(filename)
				{
					var debug=false;
					filename=filename["toLowerCase"]();
					var officeIndex=filename["indexOf"]("office_strings.js");
					if(officeIndex < 0)
					{
						officeIndex=filename["indexOf"]("office_strings.debug.js");
						debug=true
					}
					if(officeIndex > 0 && officeIndex < filename["length"])
					{
						url=filename["replace"](debug ? "office_strings.debug.js" : "office_strings.js","outlook_strings.js");
						var languageUrl=filename["substring"](0,officeIndex);
						var lastIndexOfSlash=languageUrl["lastIndexOf"]("/",languageUrl["length"] - 2);
						if(lastIndexOfSlash===-1)
							lastIndexOfSlash=languageUrl["lastIndexOf"]("\\",languageUrl["length"] - 2);
						if(lastIndexOfSlash !==-1 && languageUrl["length"] > lastIndexOfSlash+1)
							baseUrl=languageUrl["substring"](0,lastIndexOfSlash+1);
						break
					}
				}
			}
		}
		if(url)
		{
			var head=document.getElementsByTagName("head")[0];
			var scriptElement=null;
			var $$t_H=this;
			var scriptElementCallback=function()
				{
					if(stringLoadedCallback && (!scriptElement.readyState || scriptElement.readyState && (scriptElement.readyState==="loaded" || scriptElement.readyState==="complete")))
					{
						scriptElement.onload=null;
						scriptElement.onreadystatechange=null;
						stringLoadedCallback()
					}
				};
			var $$t_I=this;
			var failureCallback=function()
				{
					if(!$$t_I._areStringsLoaded$p$0())
					{
						var fallbackUrl=baseUrl+"en-us/"+"outlook_strings.js";
						scriptElement.onload=null;
						scriptElement.onreadystatechange=null;
						scriptElement=$$t_I._createScriptElement$p$0(fallbackUrl);
						scriptElement.onload=scriptElementCallback;
						scriptElement.onreadystatechange=scriptElementCallback;
						head.appendChild(scriptElement)
					}
				};
			scriptElement=this._createScriptElement$p$0(url);
			scriptElement.onload=scriptElementCallback;
			scriptElement.onreadystatechange=scriptElementCallback;
			window.setTimeout(failureCallback,2e3);
			head.appendChild(scriptElement)
		}
	},
	_createScriptElement$p$0: function(url)
	{
		var scriptElement=document.createElement("script");
		scriptElement.type="text/javascript";
		scriptElement.src=url;
		return scriptElement
	}
};
window["OSF"]["DDA"]["Settings"]=OSF.DDA.Settings=function(data)
{
	this._rawData$p$0=data
};
window["OSF"]["DDA"]["Settings"]._convertFromRawSettings$p=function(rawSettings)
{
	if(!rawSettings)
		return{};
	if(window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.get__appName$i$0()===8)
	{
		var outlookSettings=rawSettings["SettingsKey"];
		if(outlookSettings)
			return window["OSF"]["DDA"]["SettingsManager"]["deserializeSettings"](outlookSettings)
	}
	return rawSettings
};
OSF.DDA.Settings.prototype={
	_rawData$p$0: null,
	_settingsData$p$0: null,
	get__data$p$0: function()
	{
		if(!this._settingsData$p$0)
		{
			this._settingsData$p$0=window["OSF"]["DDA"]["Settings"]._convertFromRawSettings$p(this._rawData$p$0);
			this._rawData$p$0=null
		}
		return this._settingsData$p$0
	},
	get: function(name)
	{
		return this.get__data$p$0()[name]
	},
	set: function(name, value)
	{
		this.get__data$p$0()[name]=value
	},
	remove: function(name)
	{
		delete this.get__data$p$0()[name]
	},
	saveAsync: function()
	{
		var args=[];
		for(var $$pai_4=0; $$pai_4 < arguments.length;++$$pai_4)
			args[$$pai_4]=arguments[$$pai_4];
		var commonParameters=$h.CommonParameters.parse(args,false);
		if(window["JSON"]["stringify"](window["OSF"]["DDA"]["SettingsManager"]["serializeSettings"](this.get__data$p$0()))["length"] > 32768)
		{
			var asyncResult=window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.createAsyncResult(null,1,9019,commonParameters._asyncContext$p$0,"");
			var $$t_3=this;
			window.setTimeout(function()
			{
				commonParameters._callback$p$0(asyncResult)
			},0);
			return
		}
		if(window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.get__appName$i$0()===64)
			this._saveSettingsForOwa$p$0(commonParameters._callback$p$0,commonParameters._asyncContext$p$0);
		else
			this._saveSettingsForOutlook$p$0(commonParameters._callback$p$0,commonParameters._asyncContext$p$0)
	},
	_saveSettingsForOutlook$p$0: function(callback, userContext)
	{
		var storedException=null;
		try
		{
			var serializedSettings=window["OSF"]["DDA"]["SettingsManager"]["serializeSettings"](this.get__data$p$0());
			var jsonSettings=window["JSON"]["stringify"](serializedSettings);
			var settingsObjectToSave={SettingsKey: jsonSettings};
			window["OSF"]["DDA"]["RichClientSettingsManager"]["write"](settingsObjectToSave)
		}
		catch(ex)
		{
			storedException=ex
		}
		if(callback)
		{
			var asyncResult;
			if(storedException)
				asyncResult=window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.createAsyncResult(null,1,9019,userContext,storedException["message"]);
			else
				asyncResult=window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.createAsyncResult(null,0,0,userContext,null);
			callback(asyncResult)
		}
	},
	_saveSettingsForOwa$p$0: function(callback, userContext)
	{
		var serializedSettings=window["OSF"]["DDA"]["SettingsManager"]["serializeSettings"](this.get__data$p$0());
		var $$t_7=this;
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.invokeHostMethod(404,[serializedSettings],function(resultCode, response)
		{
			if(callback)
			{
				var asyncResult;
				if(resultCode)
					asyncResult=window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.createAsyncResult(null,1,9017,userContext,String.format(window["_u"]["ExtensibilityStrings"]["l_InternalProtocolError_Text"],resultCode));
				else
				{
					var responseDictionary=response;
					if(!responseDictionary["error"])
						asyncResult=window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.createAsyncResult(null,0,0,userContext,null);
					else
						asyncResult=window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.createAsyncResult(null,1,9019,userContext,responseDictionary["errorMessage"])
				}
				callback(asyncResult)
			}
		})
	}
};
Type.registerNamespace("$h");
var $h=window.$h || {};
Type.registerNamespace("Office.cast");
var Office=window.Office || {};
Office.cast=Office.cast || {};
$h.Appointment=function(dataDictionary)
{
	this.$$d__getOrganizer$p$2=Function.createDelegate(this,this._getOrganizer$p$2);
	this.$$d__getNormalizedSubject$p$2=Function.createDelegate(this,this._getNormalizedSubject$p$2);
	this.$$d__getSubject$p$2=Function.createDelegate(this,this._getSubject$p$2);
	this.$$d__getResources$p$2=Function.createDelegate(this,this._getResources$p$2);
	this.$$d__getRequiredAttendees$p$2=Function.createDelegate(this,this._getRequiredAttendees$p$2);
	this.$$d__getOptionalAttendees$p$2=Function.createDelegate(this,this._getOptionalAttendees$p$2);
	this.$$d__getLocation$p$2=Function.createDelegate(this,this._getLocation$p$2);
	this.$$d__getEnd$p$2=Function.createDelegate(this,this._getEnd$p$2);
	this.$$d__getStart$p$2=Function.createDelegate(this,this._getStart$p$2);
	$h.Appointment.initializeBase(this,[dataDictionary]);
	$h.InitialData._defineReadOnlyProperty$i(this,"start",this.$$d__getStart$p$2);
	$h.InitialData._defineReadOnlyProperty$i(this,"end",this.$$d__getEnd$p$2);
	$h.InitialData._defineReadOnlyProperty$i(this,"location",this.$$d__getLocation$p$2);
	$h.InitialData._defineReadOnlyProperty$i(this,"optionalAttendees",this.$$d__getOptionalAttendees$p$2);
	$h.InitialData._defineReadOnlyProperty$i(this,"requiredAttendees",this.$$d__getRequiredAttendees$p$2);
	$h.InitialData._defineReadOnlyProperty$i(this,"resources",this.$$d__getResources$p$2);
	$h.InitialData._defineReadOnlyProperty$i(this,"subject",this.$$d__getSubject$p$2);
	$h.InitialData._defineReadOnlyProperty$i(this,"normalizedSubject",this.$$d__getNormalizedSubject$p$2);
	$h.InitialData._defineReadOnlyProperty$i(this,"organizer",this.$$d__getOrganizer$p$2)
};
$h.Appointment.prototype={
	getEntities: function()
	{
		return this._data$p$0._getEntities$i$0()
	},
	getEntitiesByType: function(entityType)
	{
		return this._data$p$0._getEntitiesByType$i$0(entityType)
	},
	getRegExMatches: function()
	{
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"getRegExMatches");
		return this._data$p$0._getRegExMatches$i$0()
	},
	getFilteredEntitiesByName: function(name)
	{
		return this._data$p$0._getFilteredEntitiesByName$i$0(name)
	},
	getRegExMatchesByName: function(name)
	{
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"getRegExMatchesByName");
		return this._data$p$0._getRegExMatchesByName$i$0(name)
	},
	displayReplyForm: function(obj)
	{
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._displayReplyForm$i$0(obj)
	},
	displayReplyAllForm: function(obj)
	{
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._displayReplyAllForm$i$0(obj)
	},
	getItemType: function()
	{
		return window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ItemType"]["Appointment"]
	},
	_getStart$p$2: function()
	{
		return this._data$p$0.get__start$i$0()
	},
	_getEnd$p$2: function()
	{
		return this._data$p$0.get__end$i$0()
	},
	_getLocation$p$2: function()
	{
		return this._data$p$0.get__location$i$0()
	},
	_getOptionalAttendees$p$2: function()
	{
		return this._data$p$0.get__cc$i$0()
	},
	_getRequiredAttendees$p$2: function()
	{
		return this._data$p$0.get__to$i$0()
	},
	_getResources$p$2: function()
	{
		return this._data$p$0.get__resources$i$0()
	},
	_getSubject$p$2: function()
	{
		return this._data$p$0.get__subject$i$0()
	},
	_getNormalizedSubject$p$2: function()
	{
		return this._data$p$0.get__normalizedSubject$i$0()
	},
	_getOrganizer$p$2: function()
	{
		return this._data$p$0.get__organizer$i$0()
	}
};
$h.AppointmentCompose=function(data)
{
	this.$$d__getLocation$p$2=Function.createDelegate(this,this._getLocation$p$2);
	this.$$d__getEnd$p$2=Function.createDelegate(this,this._getEnd$p$2);
	this.$$d__getStart$p$2=Function.createDelegate(this,this._getStart$p$2);
	this.$$d__getOptionalAttendees$p$2=Function.createDelegate(this,this._getOptionalAttendees$p$2);
	this.$$d__getRequiredAttendees$p$2=Function.createDelegate(this,this._getRequiredAttendees$p$2);
	$h.AppointmentCompose.initializeBase(this,[data]);
	$h.InitialData._defineReadOnlyProperty$i(this,"requiredAttendees",this.$$d__getRequiredAttendees$p$2);
	$h.InitialData._defineReadOnlyProperty$i(this,"optionalAttendees",this.$$d__getOptionalAttendees$p$2);
	$h.InitialData._defineReadOnlyProperty$i(this,"start",this.$$d__getStart$p$2);
	$h.InitialData._defineReadOnlyProperty$i(this,"end",this.$$d__getEnd$p$2);
	$h.InitialData._defineReadOnlyProperty$i(this,"location",this.$$d__getLocation$p$2)
};
$h.AppointmentCompose.prototype={
	_requiredAttendees$p$2: null,
	_optionalAttendees$p$2: null,
	_start$p$2: null,
	_end$p$2: null,
	_location$p$2: null,
	getItemType: function()
	{
		return window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ItemType"]["Appointment"]
	},
	_getRequiredAttendees$p$2: function()
	{
		this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
		if(!this._requiredAttendees$p$2)
			this._requiredAttendees$p$2=new $h.ComposeRecipient(0,"requiredAttendees");
		return this._requiredAttendees$p$2
	},
	_getOptionalAttendees$p$2: function()
	{
		this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
		if(!this._optionalAttendees$p$2)
			this._optionalAttendees$p$2=new $h.ComposeRecipient(1,"optionalAttendees");
		return this._optionalAttendees$p$2
	},
	_getStart$p$2: function()
	{
		this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
		if(!this._start$p$2)
			this._start$p$2=new $h.ComposeTime(1);
		return this._start$p$2
	},
	_getEnd$p$2: function()
	{
		this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
		if(!this._end$p$2)
			this._end$p$2=new $h.ComposeTime(2);
		return this._end$p$2
	},
	_getLocation$p$2: function()
	{
		this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
		if(!this._location$p$2)
			this._location$p$2=new $h.ComposeLocation;
		return this._location$p$2
	}
};
$h.AttachmentConstants=function(){};
$h.AttachmentDetails=function(data)
{
	this.$$d__getIsInline$p$0=Function.createDelegate(this,this._getIsInline$p$0);
	this.$$d__getAttachmentType$p$0=Function.createDelegate(this,this._getAttachmentType$p$0);
	this.$$d__getSize$p$0=Function.createDelegate(this,this._getSize$p$0);
	this.$$d__getContentType$p$0=Function.createDelegate(this,this._getContentType$p$0);
	this.$$d__getName$p$0=Function.createDelegate(this,this._getName$p$0);
	this.$$d__getId$p$0=Function.createDelegate(this,this._getId$p$0);
	this._data$p$0=data;
	$h.InitialData._defineReadOnlyProperty$i(this,"id",this.$$d__getId$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"name",this.$$d__getName$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"contentType",this.$$d__getContentType$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"size",this.$$d__getSize$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"attachmentType",this.$$d__getAttachmentType$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"isInline",this.$$d__getIsInline$p$0)
};
$h.AttachmentDetails.prototype={
	_data$p$0: null,
	_getId$p$0: function()
	{
		return this._data$p$0["id"]
	},
	_getName$p$0: function()
	{
		return this._data$p$0["name"]
	},
	_getContentType$p$0: function()
	{
		return this._data$p$0["contentType"]
	},
	_getSize$p$0: function()
	{
		return this._data$p$0["size"]
	},
	_getAttachmentType$p$0: function()
	{
		var response=this._data$p$0["attachmentType"];
		return response < $h.AttachmentDetails._attachmentTypeMap$p.length ? $h.AttachmentDetails._attachmentTypeMap$p[response] : window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["AttachmentType"]["File"]
	},
	_getIsInline$p$0: function()
	{
		return this._data$p$0["isInline"]
	}
};
$h.Body=function(){};
$h.Body._tryMapToHostCoercionType$i=function(coercionType, hostCoercionType)
{
	hostCoercionType.val=undefined;
	if(coercionType===window["Microsoft"]["Office"]["WebExtension"]["CoercionType"]["Html"])
		hostCoercionType.val=3;
	else if(coercionType===window["Microsoft"]["Office"]["WebExtension"]["CoercionType"]["Text"])
		hostCoercionType.val=0;
	else
		return false;
	return true
};
$h.Body.prototype={getAsync: function(coercionType)
	{
		var args=[];
		for(var $$pai_7=1; $$pai_7 < arguments.length;++$$pai_7)
			args[$$pai_7 - 1]=arguments[$$pai_7];
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"body.getAsync");
		var commonParameters=$h.CommonParameters.parse(args,true);
		var hostCoercionType;
		var $$t_5,
			$$t_6;
		if(!($$t_6=$h.Body._tryMapToHostCoercionType$i(coercionType,$$t_5={val: hostCoercionType}),hostCoercionType=$$t_5["val"],$$t_6))
			throw Error.argument("coercionType");
		var dataToHost={coercionType: hostCoercionType};
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(37,dataToHost,null,commonParameters._asyncContext$p$0,commonParameters._callback$p$0)
	}};
$h.ComposeBody=function()
{
	$h.ComposeBody.initializeBase(this)
};
$h.ComposeBody._createParameterDictionaryToHost$i=function(data, parameters)
{
	var dataToHost={data: data};
	if(parameters._options$p$0 && !$h.ScriptHelpers.isNull(parameters._options$p$0["coercionType"]))
	{
		var hostCoercionType;
		var $$t_4,
			$$t_5;
		if(!($$t_5=$h.Body._tryMapToHostCoercionType$i(parameters._options$p$0["coercionType"],$$t_4={val: hostCoercionType}),hostCoercionType=$$t_4["val"],$$t_5))
		{
			if(parameters._callback$p$0)
				parameters._callback$p$0(window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.createAsyncResult(null,1,1e3,parameters._asyncContext$p$0,null));
			return null
		}
		dataToHost["coercionType"]=hostCoercionType
	}
	else
		dataToHost["coercionType"]=0;
	return dataToHost
};
$h.ComposeBody.prototype={
	getTypeAsync: function()
	{
		var args=[];
		for(var $$pai_2=0; $$pai_2 < arguments.length;++$$pai_2)
			args[$$pai_2]=arguments[$$pai_2];
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"body.getTypeAsync");
		var parameters=$h.CommonParameters.parse(args,true);
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(14,null,null,parameters._asyncContext$p$0,parameters._callback$p$0)
	},
	setSelectedDataAsync: function(data)
	{
		var args=[];
		for(var $$pai_4=1; $$pai_4 < arguments.length;++$$pai_4)
			args[$$pai_4 - 1]=arguments[$$pai_4];
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"body.setSelectedDataAsync");
		var parameters=$h.CommonParameters.parse(args,false);
		if(!String.isInstanceOfType(data))
			throw Error.argumentType("data",Object["getType"](data),String);
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(data["length"],0,1e6,"data");
		var dataToHost=$h.ComposeBody._createParameterDictionaryToHost$i(data,parameters);
		if(!dataToHost)
			return;
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(13,dataToHost,null,parameters._asyncContext$p$0,parameters._callback$p$0)
	},
	prependAsync: function(data)
	{
		var args=[];
		for(var $$pai_4=1; $$pai_4 < arguments.length;++$$pai_4)
			args[$$pai_4 - 1]=arguments[$$pai_4];
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"body.prependAsync");
		var parameters=$h.CommonParameters.parse(args,false);
		if(!String.isInstanceOfType(data))
			throw Error.argumentType("data",Object["getType"](data),String);
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(data["length"],0,1e6,"data");
		var dataToHost=$h.ComposeBody._createParameterDictionaryToHost$i(data,parameters);
		if(!dataToHost)
			return;
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(23,dataToHost,null,parameters._asyncContext$p$0,parameters._callback$p$0)
	},
	setAsync: function(data)
	{
		var args=[];
		for(var $$pai_4=1; $$pai_4 < arguments.length;++$$pai_4)
			args[$$pai_4 - 1]=arguments[$$pai_4];
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"body.setAsync");
		var parameters=$h.CommonParameters.parse(args,false);
		if(!String.isInstanceOfType(data))
			throw Error.argumentType("data",Object["getType"](data),String);
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(data["length"],0,1e6,"data");
		var dataToHost=$h.ComposeBody._createParameterDictionaryToHost$i(data,parameters);
		if(!dataToHost)
			return;
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(38,dataToHost,null,parameters._asyncContext$p$0,parameters._callback$p$0)
	}
};
$h.ComposeItem=function(data)
{
	this.$$d__getBody$p$1=Function.createDelegate(this,this._getBody$p$1);
	this.$$d__getSubject$p$1=Function.createDelegate(this,this._getSubject$p$1);
	$h.ComposeItem.initializeBase(this,[data]);
	$h.InitialData._defineReadOnlyProperty$i(this,"subject",this.$$d__getSubject$p$1);
	$h.InitialData._defineReadOnlyProperty$i(this,"body",this.$$d__getBody$p$1)
};
$h.ComposeItem.prototype={
	_subject$p$1: null,
	_body$p$1: null,
	addFileAttachmentAsync: function(uri, attachmentName)
	{
		var args=[];
		for(var $$pai_5=2; $$pai_5 < arguments.length;++$$pai_5)
			args[$$pai_5 - 2]=arguments[$$pai_5];
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"addFileAttachmentAsync");
		if(!$h.ScriptHelpers.isNonEmptyString(uri))
			throw Error.argument("uri");
		if(!$h.ScriptHelpers.isNonEmptyString(attachmentName))
			throw Error.argument("attachmentName");
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(uri["length"],0,2048,"uri");
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(attachmentName["length"],0,255,"attachmentName");
		var commonParameters=$h.CommonParameters.parse(args,false);
		var parameters={
				uri: uri,
				name: attachmentName,
				__timeout__: 6e5
			};
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(16,parameters,null,commonParameters._asyncContext$p$0,commonParameters._callback$p$0)
	},
	addItemAttachmentAsync: function(itemId, attachmentName)
	{
		var args=[];
		for(var $$pai_5=2; $$pai_5 < arguments.length;++$$pai_5)
			args[$$pai_5 - 2]=arguments[$$pai_5];
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"addItemAttachmentAsync");
		if(!$h.ScriptHelpers.isNonEmptyString(itemId))
			throw Error.argument("itemId");
		if(!$h.ScriptHelpers.isNonEmptyString(attachmentName))
			throw Error.argument("attachmentName");
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(itemId["length"],0,200,"itemId");
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(attachmentName["length"],0,255,"attachmentName");
		var commonParameters=$h.CommonParameters.parse(args,false);
		var parameters={
				itemId: window["OSF"]["DDA"]["OutlookAppOm"].getItemIdBasedOnHost(itemId),
				name: attachmentName,
				__timeout__: 6e5
			};
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(19,parameters,null,commonParameters._asyncContext$p$0,commonParameters._callback$p$0)
	},
	removeAttachmentAsync: function(attachmentId)
	{
		var args=[];
		for(var $$pai_3=1; $$pai_3 < arguments.length;++$$pai_3)
			args[$$pai_3 - 1]=arguments[$$pai_3];
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"removeAttachmentAsync");
		if(!$h.ScriptHelpers.isNonEmptyString(attachmentId))
			throw Error.argument("attachmentId");
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(attachmentId["length"],0,200,"attachmentId");
		var commonParameters=$h.CommonParameters.parse(args,false);
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(20,{attachmentIndex: attachmentId},null,commonParameters._asyncContext$p$0,commonParameters._callback$p$0)
	},
	getSelectedDataAsync: function(coercionType)
	{
		var args=[];
		for(var $$pai_7=1; $$pai_7 < arguments.length;++$$pai_7)
			args[$$pai_7 - 1]=arguments[$$pai_7];
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"getSelectedDataAsync");
		var commonParameters=$h.CommonParameters.parse(args,true);
		var hostCoercionType;
		var $$t_5,
			$$t_6;
		if(coercionType !==window["Microsoft"]["Office"]["WebExtension"]["CoercionType"]["Html"] && coercionType !==window["Microsoft"]["Office"]["WebExtension"]["CoercionType"]["Text"] || !($$t_6=$h.Body._tryMapToHostCoercionType$i(coercionType,$$t_5={val: hostCoercionType}),hostCoercionType=$$t_5["val"],$$t_6))
			throw Error.argument("coercionType");
		var dataToHost={coercionType: hostCoercionType};
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(28,dataToHost,null,commonParameters._asyncContext$p$0,commonParameters._callback$p$0)
	},
	setSelectedDataAsync: function(data)
	{
		var args=[];
		for(var $$pai_4=1; $$pai_4 < arguments.length;++$$pai_4)
			args[$$pai_4 - 1]=arguments[$$pai_4];
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"setSelectedDataAsync");
		var parameters=$h.CommonParameters.parse(args,false);
		if(!String.isInstanceOfType(data))
			throw Error.argumentType("data",Object["getType"](data),String);
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(data["length"],0,1e6,"data");
		var dataToHost=$h.ComposeBody._createParameterDictionaryToHost$i(data,parameters);
		if(!dataToHost)
			return;
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(29,dataToHost,null,parameters._asyncContext$p$0,parameters._callback$p$0)
	},
	close: function()
	{
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(41,null,null,null,null)
	},
	_getBody$p$1: function()
	{
		this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
		if(!this._body$p$1)
			this._body$p$1=new $h.ComposeBody;
		return this._body$p$1
	},
	_getSubject$p$1: function()
	{
		this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
		if(!this._subject$p$1)
			this._subject$p$1=new $h.ComposeSubject;
		return this._subject$p$1
	},
	saveAsync: function()
	{
		var args=[];
		for(var $$pai_2=0; $$pai_2 < arguments.length;++$$pai_2)
			args[$$pai_2]=arguments[$$pai_2];
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"saveAsync");
		var parameters=$h.CommonParameters.parse(args,false);
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(32,null,null,parameters._asyncContext$p$0,parameters._callback$p$0)
	}
};
$h.ComposeRecipient=function(type, propertyName)
{
	this._type$p$0=type;
	this._propertyName$p$0=propertyName
};
$h.ComposeRecipient._throwOnInvalidDisplayNameOrEmail$p=function(displayName, emailAddress)
{
	if(!displayName && !emailAddress)
		throw Error.argument("recipients");
	if(displayName && displayName["length"] > 255)
		throw Error.argumentOutOfRange("recipients",displayName["length"],window["_u"]["ExtensibilityStrings"]["l_DisplayNameTooLong_Text"]);
	if(emailAddress && emailAddress["length"] > 571)
		throw Error.argumentOutOfRange("recipients",emailAddress["length"],window["_u"]["ExtensibilityStrings"]["l_EmailAddressTooLong_Text"]);
};
$h.ComposeRecipient._getAsyncFormatter$p=function(rawInput)
{
	var input=rawInput;
	var output=[];
	for(var i=0; i < input["length"]; i++)
	{
		var email=new $h.EmailAddressDetails(input[i]);
		output[i]=email
	}
	return output
};
$h.ComposeRecipient._createEmailDictionaryForHost$p=function(address, name)
{
	return{
			address: address,
			name: name
		}
};
$h.ComposeRecipient.prototype={
	_propertyName$p$0: null,
	_type$p$0: 0,
	getAsync: function()
	{
		var args=[];
		for(var $$pai_2=0; $$pai_2 < arguments.length;++$$pai_2)
			args[$$pai_2]=arguments[$$pai_2];
		var parameters=$h.CommonParameters.parse(args,true);
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,this._propertyName$p$0+".getAsync");
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(15,{recipientField: this._type$p$0},$h.ComposeRecipient._getAsyncFormatter$p,parameters._asyncContext$p$0,parameters._callback$p$0)
	},
	setAsync: function(recipients)
	{
		var args=[];
		for(var $$pai_2=1; $$pai_2 < arguments.length;++$$pai_2)
			args[$$pai_2 - 1]=arguments[$$pai_2];
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,this._propertyName$p$0+".setAsync");
		this.setAddHelper(recipients,args,true)
	},
	addAsync: function(recipients)
	{
		var args=[];
		for(var $$pai_2=1; $$pai_2 < arguments.length;++$$pai_2)
			args[$$pai_2 - 1]=arguments[$$pai_2];
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,this._propertyName$p$0+".addAsync");
		this.setAddHelper(recipients,args,false)
	},
	setAddHelper: function(recipients, args, isSet)
	{
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(recipients["length"],0,100,"recipients");
		var parameters=$h.CommonParameters.parse(args,false);
		var recipientData=[];
		if(Array.isInstanceOfType(recipients))
			for(var i=0; i < recipients["length"]; i++)
				if(String.isInstanceOfType(recipients[i]))
				{
					$h.ComposeRecipient._throwOnInvalidDisplayNameOrEmail$p(recipients[i],recipients[i]);
					recipientData[i]=$h.ComposeRecipient._createEmailDictionaryForHost$p(recipients[i],recipients[i])
				}
				else if($h.EmailAddressDetails.isInstanceOfType(recipients[i]))
				{
					var address=recipients[i];
					$h.ComposeRecipient._throwOnInvalidDisplayNameOrEmail$p(address["displayName"],address["emailAddress"]);
					recipientData[i]=$h.ComposeRecipient._createEmailDictionaryForHost$p(address["emailAddress"],address["displayName"])
				}
				else if(Object.isInstanceOfType(recipients[i]))
				{
					var input=recipients[i];
					var emailAddress=input["emailAddress"];
					var displayName=input["displayName"];
					$h.ComposeRecipient._throwOnInvalidDisplayNameOrEmail$p(displayName,emailAddress);
					recipientData[i]=$h.ComposeRecipient._createEmailDictionaryForHost$p(emailAddress,displayName)
				}
				else
					throw Error.argument("recipients");
		else
			throw Error.argument("recipients");
		var $$t_B=this;
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(isSet ? 21 : 22,{
			recipientField: this._type$p$0,
			recipientArray: recipientData
		},function(rawInput)
		{
			return rawInput
		},parameters._asyncContext$p$0,parameters._callback$p$0)
	}
};
$h.ComposeRecipient.RecipientField=function(){};
$h.ComposeRecipient.RecipientField.prototype={
	to: 0,
	cc: 1,
	bcc: 2,
	requiredAttendees: 0,
	optionalAttendees: 1
};
$h.ComposeRecipient.RecipientField.registerEnum("$h.ComposeRecipient.RecipientField",false);
$h.ComposeLocation=function(){};
$h.ComposeLocation.prototype={
	getAsync: function()
	{
		var args=[];
		for(var $$pai_2=0; $$pai_2 < arguments.length;++$$pai_2)
			args[$$pai_2]=arguments[$$pai_2];
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"location.getAsync");
		var parameters=$h.CommonParameters.parse(args,true);
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(26,null,null,parameters._asyncContext$p$0,parameters._callback$p$0)
	},
	setAsync: function(location)
	{
		var args=[];
		for(var $$pai_3=1; $$pai_3 < arguments.length;++$$pai_3)
			args[$$pai_3 - 1]=arguments[$$pai_3];
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"location.setAsync");
		var parameters=$h.CommonParameters.parse(args,false);
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(location["length"],0,255,"location");
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(27,{location: location},null,parameters._asyncContext$p$0,parameters._callback$p$0)
	}
};
$h.ComposeSubject=function(){};
$h.ComposeSubject.prototype={
	getAsync: function()
	{
		var args=[];
		for(var $$pai_2=0; $$pai_2 < arguments.length;++$$pai_2)
			args[$$pai_2]=arguments[$$pai_2];
		var parameters=$h.CommonParameters.parse(args,true);
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"subject.getAsync");
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(18,null,null,parameters._asyncContext$p$0,parameters._callback$p$0)
	},
	setAsync: function(data)
	{
		var args=[];
		for(var $$pai_3=1; $$pai_3 < arguments.length;++$$pai_3)
			args[$$pai_3 - 1]=arguments[$$pai_3];
		var parameters=$h.CommonParameters.parse(args,false);
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"subject.setAsync");
		if(!String.isInstanceOfType(data))
			throw Error.argument("data");
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(data["length"],0,255,"data");
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(17,{subject: data},null,parameters._asyncContext$p$0,parameters._callback$p$0)
	}
};
$h.ComposeTime=function(type)
{
	this.$$d__ticksToDateFormatter$p$0=Function.createDelegate(this,this._ticksToDateFormatter$p$0);
	this._timeType$p$0=type
};
$h.ComposeTime.prototype={
	_timeType$p$0: 0,
	getAsync: function()
	{
		var args=[];
		for(var $$pai_2=0; $$pai_2 < arguments.length;++$$pai_2)
			args[$$pai_2]=arguments[$$pai_2];
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,this._getPropertyName$p$0()+".getAsync");
		var parameters=$h.CommonParameters.parse(args,true);
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(24,{TimeProperty: this._timeType$p$0},this.$$d__ticksToDateFormatter$p$0,parameters._asyncContext$p$0,parameters._callback$p$0)
	},
	setAsync: function(dateTime)
	{
		var args=[];
		for(var $$pai_3=1; $$pai_3 < arguments.length;++$$pai_3)
			args[$$pai_3 - 1]=arguments[$$pai_3];
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,this._getPropertyName$p$0()+".setAsync");
		if(!Date.isInstanceOfType(dateTime))
			throw Error.argumentType("dateTime",Object["getType"](dateTime),Date);
		if(window["isNaN"](dateTime["getTime"]()))
			throw Error.argument("dateTime");
		if(dateTime["getTime"]() < -864e13 || dateTime["getTime"]() > 864e13)
			throw Error.argumentOutOfRange("dateTime");
		var parameters=$h.CommonParameters.parse(args,false);
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(25,{
			TimeProperty: this._timeType$p$0,
			time: dateTime["getTime"]()
		},null,parameters._asyncContext$p$0,parameters._callback$p$0)
	},
	_ticksToDateFormatter$p$0: function(rawInput)
	{
		var ticks=rawInput;
		return new Date(ticks)
	},
	_getPropertyName$p$0: function()
	{
		return this._timeType$p$0===1 ? "start" : "end"
	}
};
$h.ComposeTime.TimeType=function(){};
$h.ComposeTime.TimeType.prototype={
	start: 1,
	end: 2
};
$h.ComposeTime.TimeType.registerEnum("$h.ComposeTime.TimeType",false);
$h.Contact=function(data)
{
	this.$$d__getContactString$p$0=Function.createDelegate(this,this._getContactString$p$0);
	this.$$d__getAddresses$p$0=Function.createDelegate(this,this._getAddresses$p$0);
	this.$$d__getUrls$p$0=Function.createDelegate(this,this._getUrls$p$0);
	this.$$d__getEmailAddresses$p$0=Function.createDelegate(this,this._getEmailAddresses$p$0);
	this.$$d__getPhoneNumbers$p$0=Function.createDelegate(this,this._getPhoneNumbers$p$0);
	this.$$d__getBusinessName$p$0=Function.createDelegate(this,this._getBusinessName$p$0);
	this.$$d__getPersonName$p$0=Function.createDelegate(this,this._getPersonName$p$0);
	this._data$p$0=data;
	$h.InitialData._defineReadOnlyProperty$i(this,"personName",this.$$d__getPersonName$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"businessName",this.$$d__getBusinessName$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"phoneNumbers",this.$$d__getPhoneNumbers$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"emailAddresses",this.$$d__getEmailAddresses$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"urls",this.$$d__getUrls$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"addresses",this.$$d__getAddresses$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"contactString",this.$$d__getContactString$p$0)
};
$h.Contact.prototype={
	_data$p$0: null,
	_phoneNumbers$p$0: null,
	_getPersonName$p$0: function()
	{
		return this._data$p$0["PersonName"]
	},
	_getBusinessName$p$0: function()
	{
		return this._data$p$0["BusinessName"]
	},
	_getAddresses$p$0: function()
	{
		return $h.Entities._getExtractedStringProperty$i(this._data$p$0,"Addresses")
	},
	_getEmailAddresses$p$0: function()
	{
		return $h.Entities._getExtractedStringProperty$i(this._data$p$0,"EmailAddresses")
	},
	_getUrls$p$0: function()
	{
		return $h.Entities._getExtractedStringProperty$i(this._data$p$0,"Urls")
	},
	_getPhoneNumbers$p$0: function()
	{
		if(!this._phoneNumbers$p$0)
		{
			var $$t_1=this;
			this._phoneNumbers$p$0=$h.Entities._getExtractedObjects$i($h.PhoneNumber,this._data$p$0,"PhoneNumbers",function(data)
			{
				return new $h.PhoneNumber(data)
			})
		}
		return this._phoneNumbers$p$0
	},
	_getContactString$p$0: function()
	{
		return this._data$p$0["ContactString"]
	}
};
$h.CustomProperties=function(data)
{
	if($h.ScriptHelpers.isNullOrUndefined(data))
		throw Error.argumentNull("data");
	if(Array.isInstanceOfType(data))
	{
		var customPropertiesArray=data;
		if(customPropertiesArray.length > 0)
			this._data$p$0=customPropertiesArray[0];
		else
			throw Error.argument("data");
	}
	else
		this._data$p$0=data
};
$h.CustomProperties.prototype={
	_data$p$0: null,
	get: function(name)
	{
		var value=this._data$p$0[name];
		if(typeof value==="string")
		{
			var valueString=value;
			if(valueString["length"] > 6 && valueString.startsWith("Date(") && valueString.endsWith(")"))
			{
				var ticksString=valueString["substring"](5,valueString["length"] - 1);
				var ticks=window["parseInt"](ticksString);
				if(!window["isNaN"](ticks))
				{
					var dateTimeValue=new Date(ticks);
					if(dateTimeValue)
						value=dateTimeValue
				}
			}
		}
		return value
	},
	set: function(name, value)
	{
		if(window["OSF"]["OUtil"]["isDate"](value))
			value="Date("+value["getTime"]()+")";
		this._data$p$0[name]=value
	},
	remove: function(name)
	{
		delete this._data$p$0[name]
	},
	saveAsync: function()
	{
		var args=[];
		for(var $$pai_4=0; $$pai_4 < arguments.length;++$$pai_4)
			args[$$pai_4]=arguments[$$pai_4];
		var MaxCustomPropertiesLength=2500;
		if(window["JSON"]["stringify"](this._data$p$0)["length"] > MaxCustomPropertiesLength)
			throw Error.argument();
		var parameters=$h.CommonParameters.parse(args,false,true);
		var saveCustomProperties=new $h.SaveDictionaryRequest(parameters._callback$p$0,parameters._asyncContext$p$0);
		saveCustomProperties._sendRequest$i$0(4,"SaveCustomProperties",{customProperties: this._data$p$0})
	}
};
$h.Diagnostics=function(data, appName)
{
	this.$$d__getOwaView$p$0=Function.createDelegate(this,this._getOwaView$p$0);
	this.$$d__getHostVersion$p$0=Function.createDelegate(this,this._getHostVersion$p$0);
	this.$$d__getHostName$p$0=Function.createDelegate(this,this._getHostName$p$0);
	this._data$p$0=data;
	this._appName$p$0=appName;
	$h.InitialData._defineReadOnlyProperty$i(this,"hostName",this.$$d__getHostName$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"hostVersion",this.$$d__getHostVersion$p$0);
	if(64===this._appName$p$0)
		$h.InitialData._defineReadOnlyProperty$i(this,"OWAView",this.$$d__getOwaView$p$0)
};
$h.Diagnostics.prototype={
	_data$p$0: null,
	_appName$p$0: 0,
	_getHostName$p$0: function()
	{
		if(8===this._appName$p$0)
			return"Outlook";
		else if(64===this._appName$p$0)
			return"OutlookWebApp";
		return null
	},
	_getHostVersion$p$0: function()
	{
		return this._data$p$0.get__hostVersion$i$0()
	},
	_getOwaView$p$0: function()
	{
		return this._data$p$0.get__owaView$i$0()
	}
};
$h.EmailAddressDetails=function(data)
{
	this.$$d__getRecipientType$p$0=Function.createDelegate(this,this._getRecipientType$p$0);
	this.$$d__getAppointmentResponse$p$0=Function.createDelegate(this,this._getAppointmentResponse$p$0);
	this.$$d__getDisplayName$p$0=Function.createDelegate(this,this._getDisplayName$p$0);
	this.$$d__getEmailAddress$p$0=Function.createDelegate(this,this._getEmailAddress$p$0);
	this._data$p$0=data;
	$h.InitialData._defineReadOnlyProperty$i(this,"emailAddress",this.$$d__getEmailAddress$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"displayName",this.$$d__getDisplayName$p$0);
	if($h.ScriptHelpers.dictionaryContainsKey(data,"appointmentResponse"))
		$h.InitialData._defineReadOnlyProperty$i(this,"appointmentResponse",this.$$d__getAppointmentResponse$p$0);
	if($h.ScriptHelpers.dictionaryContainsKey(data,"recipientType"))
		$h.InitialData._defineReadOnlyProperty$i(this,"recipientType",this.$$d__getRecipientType$p$0)
};
$h.EmailAddressDetails._createFromEmailUserDictionary$i=function(data)
{
	var emailAddressDetailsDictionary={};
	var displayName=data["Name"];
	var emailAddress=data["UserId"];
	emailAddressDetailsDictionary["name"]=displayName || $h.EmailAddressDetails._emptyString$p;
	emailAddressDetailsDictionary["address"]=emailAddress || $h.EmailAddressDetails._emptyString$p;
	return new $h.EmailAddressDetails(emailAddressDetailsDictionary)
};
$h.EmailAddressDetails.prototype={
	_data$p$0: null,
	toJSON: function()
	{
		var result={};
		result["emailAddress"]=this._getEmailAddress$p$0();
		result["displayName"]=this._getDisplayName$p$0();
		if($h.ScriptHelpers.dictionaryContainsKey(this._data$p$0,"appointmentResponse"))
			result["appointmentResponse"]=this._getAppointmentResponse$p$0();
		if($h.ScriptHelpers.dictionaryContainsKey(this._data$p$0,"recipientType"))
			result["recipientType"]=this._getRecipientType$p$0();
		return result
	},
	_getEmailAddress$p$0: function()
	{
		return this._data$p$0["address"]
	},
	_getDisplayName$p$0: function()
	{
		return this._data$p$0["name"]
	},
	_getAppointmentResponse$p$0: function()
	{
		var response=this._data$p$0["appointmentResponse"];
		return response < $h.EmailAddressDetails._responseTypeMap$p.length ? $h.EmailAddressDetails._responseTypeMap$p[response] : window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ResponseType"]["None"]
	},
	_getRecipientType$p$0: function()
	{
		var response=this._data$p$0["recipientType"];
		return response < $h.EmailAddressDetails._recipientTypeMap$p.length ? $h.EmailAddressDetails._recipientTypeMap$p[response] : window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RecipientType"]["Other"]
	}
};
$h.Entities=function(data, filteredEntitiesData, timeSent, permissionLevel)
{
	this.$$d__createMeetingSuggestion$p$0=Function.createDelegate(this,this._createMeetingSuggestion$p$0);
	this.$$d__getParcelDeliveries$p$0=Function.createDelegate(this,this._getParcelDeliveries$p$0);
	this.$$d__getFlightReservations$p$0=Function.createDelegate(this,this._getFlightReservations$p$0);
	this.$$d__getContacts$p$0=Function.createDelegate(this,this._getContacts$p$0);
	this.$$d__getPhoneNumbers$p$0=Function.createDelegate(this,this._getPhoneNumbers$p$0);
	this.$$d__getUrls$p$0=Function.createDelegate(this,this._getUrls$p$0);
	this.$$d__getEmailAddresses$p$0=Function.createDelegate(this,this._getEmailAddresses$p$0);
	this.$$d__getMeetingSuggestions$p$0=Function.createDelegate(this,this._getMeetingSuggestions$p$0);
	this.$$d__getTaskSuggestions$p$0=Function.createDelegate(this,this._getTaskSuggestions$p$0);
	this.$$d__getAddresses$p$0=Function.createDelegate(this,this._getAddresses$p$0);
	this._data$p$0=data || {};
	this._filteredData$p$0=filteredEntitiesData || {};
	this._dateTimeSent$p$0=timeSent;
	$h.InitialData._defineReadOnlyProperty$i(this,"addresses",this.$$d__getAddresses$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"taskSuggestions",this.$$d__getTaskSuggestions$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"meetingSuggestions",this.$$d__getMeetingSuggestions$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"emailAddresses",this.$$d__getEmailAddresses$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"urls",this.$$d__getUrls$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"phoneNumbers",this.$$d__getPhoneNumbers$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"contacts",this.$$d__getContacts$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"flightReservations",this.$$d__getFlightReservations$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"parcelDeliveries",this.$$d__getParcelDeliveries$p$0);
	this._permissionLevel$p$0=permissionLevel
};
$h.Entities._getExtractedObjects$i=function(T, data, name, creator, removeDuplicates, stringPropertyName)
{
	var results=null;
	var extractedObjects=data[name];
	if(!extractedObjects)
		return new Array(0);
	if(removeDuplicates)
		extractedObjects=$h.Entities._removeDuplicate$p(Object,extractedObjects,$h.Entities._entityDictionaryEquals$p,stringPropertyName);
	results=new Array(extractedObjects.length);
	var count=0;
	for(var $$arr_9=extractedObjects, $$len_A=$$arr_9["length"], $$idx_B=0; $$idx_B < $$len_A;++$$idx_B)
	{
		var extractedObject=$$arr_9[$$idx_B];
		if(creator)
			results[count++]=creator(extractedObject);
		else
			results[count++]=extractedObject
	}
	return results
};
$h.Entities._getExtractedStringProperty$i=function(data, name, removeDuplicate)
{
	var extractedProperties=data[name];
	if(!extractedProperties)
		return new Array(0);
	if(removeDuplicate)
		extractedProperties=$h.Entities._removeDuplicate$p(String,extractedProperties,$h.Entities._stringEquals$p,null);
	return extractedProperties
};
$h.Entities._createContact$p=function(data)
{
	return new $h.Contact(data)
};
$h.Entities._createTaskSuggestion$p=function(data)
{
	return new $h.TaskSuggestion(data)
};
$h.Entities._createPhoneNumber$p=function(data)
{
	return new $h.PhoneNumber(data)
};
$h.Entities._entityDictionaryEquals$p=function(dictionary1, dictionary2, entityPropertyIdentifier)
{
	if(dictionary1===dictionary2)
		return true;
	if(!dictionary1 || !dictionary2)
		return false;
	if(dictionary1[entityPropertyIdentifier]===dictionary2[entityPropertyIdentifier])
		return true;
	return false
};
$h.Entities._stringEquals$p=function(string1, string2, entityProperty)
{
	return string1===string2
};
$h.Entities._removeDuplicate$p=function(T, array, entityEquals, entityPropertyIdentifier)
{
	for(var matchIndex1=array.length - 1; matchIndex1 >=0; matchIndex1--)
	{
		var removeMatch=false;
		for(var matchIndex2=matchIndex1 - 1; matchIndex2 >=0; matchIndex2--)
			if(entityEquals(array[matchIndex1],array[matchIndex2],entityPropertyIdentifier))
			{
				removeMatch=true;
				break
			}
		if(removeMatch)
			Array.removeAt(array,matchIndex1)
	}
	return array
};
$h.Entities.prototype={
	_dateTimeSent$p$0: null,
	_data$p$0: null,
	_filteredData$p$0: null,
	_filteredEntitiesCache$p$0: null,
	_permissionLevel$p$0: 0,
	_taskSuggestions$p$0: null,
	_meetingSuggestions$p$0: null,
	_phoneNumbers$p$0: null,
	_contacts$p$0: null,
	_addresses$p$0: null,
	_emailAddresses$p$0: null,
	_urls$p$0: null,
	_flightReservations$p$0: null,
	_parcelDeliveries$p$0: null,
	_getByType$i$0: function(entityType)
	{
		if(entityType===window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["EntityType"]["MeetingSuggestion"])
			return this._getMeetingSuggestions$p$0();
		else if(entityType===window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["EntityType"]["TaskSuggestion"])
			return this._getTaskSuggestions$p$0();
		else if(entityType===window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["EntityType"]["Address"])
			return this._getAddresses$p$0();
		else if(entityType===window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["EntityType"]["PhoneNumber"])
			return this._getPhoneNumbers$p$0();
		else if(entityType===window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["EntityType"]["EmailAddress"])
			return this._getEmailAddresses$p$0();
		else if(entityType===window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["EntityType"]["Url"])
			return this._getUrls$p$0();
		else if(entityType===window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["EntityType"]["Contact"])
			return this._getContacts$p$0();
		else if(entityType===window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["EntityType"]["FlightReservations"])
			return this._getFlightReservations$p$0();
		else if(entityType===window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["EntityType"]["ParcelDeliveries"])
			return this._getParcelDeliveries$p$0();
		return null
	},
	_getFilteredEntitiesByName$i$0: function(name)
	{
		if(!this._filteredEntitiesCache$p$0)
			this._filteredEntitiesCache$p$0={};
		if(!$h.ScriptHelpers.dictionaryContainsKey(this._filteredEntitiesCache$p$0,name))
		{
			var found=false;
			for(var i=0; i < $h.Entities._allEntityKeys$p.length; i++)
			{
				var entityTypeKey=$h.Entities._allEntityKeys$p[i];
				var perEntityTypeDictionary=this._filteredData$p$0[entityTypeKey];
				if(!perEntityTypeDictionary)
					continue;
				if($h.ScriptHelpers.dictionaryContainsKey(perEntityTypeDictionary,name))
				{
					switch(entityTypeKey)
					{
						case"EmailAddresses":
						case"Urls":
							this._filteredEntitiesCache$p$0[name]=$h.Entities._getExtractedStringProperty$i(perEntityTypeDictionary,name);
							break;
						case"Addresses":
							this._filteredEntitiesCache$p$0[name]=$h.Entities._getExtractedStringProperty$i(perEntityTypeDictionary,name,true);
							break;
						case"PhoneNumbers":
							this._filteredEntitiesCache$p$0[name]=$h.Entities._getExtractedObjects$i($h.PhoneNumber,perEntityTypeDictionary,name,$h.Entities._createPhoneNumber$p,false,null);
							break;
						case"TaskSuggestions":
							this._filteredEntitiesCache$p$0[name]=$h.Entities._getExtractedObjects$i($h.TaskSuggestion,perEntityTypeDictionary,name,$h.Entities._createTaskSuggestion$p,true,"TaskString");
							break;
						case"MeetingSuggestions":
							this._filteredEntitiesCache$p$0[name]=$h.Entities._getExtractedObjects$i($h.MeetingSuggestion,perEntityTypeDictionary,name,this.$$d__createMeetingSuggestion$p$0,true,"MeetingString");
							break;
						case"Contacts":
							this._filteredEntitiesCache$p$0[name]=$h.Entities._getExtractedObjects$i($h.Contact,perEntityTypeDictionary,name,$h.Entities._createContact$p,true,"ContactString");
							break
					}
					found=true;
					break
				}
			}
			if(!found)
				this._filteredEntitiesCache$p$0[name]=null
		}
		return this._filteredEntitiesCache$p$0[name]
	},
	_createMeetingSuggestion$p$0: function(data)
	{
		return new $h.MeetingSuggestion(data,this._dateTimeSent$p$0)
	},
	_getAddresses$p$0: function()
	{
		if(!this._addresses$p$0)
			this._addresses$p$0=$h.Entities._getExtractedStringProperty$i(this._data$p$0,"Addresses",true);
		return this._addresses$p$0
	},
	_getEmailAddresses$p$0: function()
	{
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnPropertyAccessForRestrictedPermission$i(this._permissionLevel$p$0);
		if(!this._emailAddresses$p$0)
			this._emailAddresses$p$0=$h.Entities._getExtractedStringProperty$i(this._data$p$0,"EmailAddresses",false);
		return this._emailAddresses$p$0
	},
	_getUrls$p$0: function()
	{
		if(!this._urls$p$0)
			this._urls$p$0=$h.Entities._getExtractedStringProperty$i(this._data$p$0,"Urls",false);
		return this._urls$p$0
	},
	_getPhoneNumbers$p$0: function()
	{
		if(!this._phoneNumbers$p$0)
			this._phoneNumbers$p$0=$h.Entities._getExtractedObjects$i($h.PhoneNumber,this._data$p$0,"PhoneNumbers",$h.Entities._createPhoneNumber$p);
		return this._phoneNumbers$p$0
	},
	_getTaskSuggestions$p$0: function()
	{
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnPropertyAccessForRestrictedPermission$i(this._permissionLevel$p$0);
		if(!this._taskSuggestions$p$0)
			this._taskSuggestions$p$0=$h.Entities._getExtractedObjects$i($h.TaskSuggestion,this._data$p$0,"TaskSuggestions",$h.Entities._createTaskSuggestion$p,true,"TaskString");
		return this._taskSuggestions$p$0
	},
	_getMeetingSuggestions$p$0: function()
	{
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnPropertyAccessForRestrictedPermission$i(this._permissionLevel$p$0);
		if(!this._meetingSuggestions$p$0)
			this._meetingSuggestions$p$0=$h.Entities._getExtractedObjects$i($h.MeetingSuggestion,this._data$p$0,"MeetingSuggestions",this.$$d__createMeetingSuggestion$p$0,true,"MeetingString");
		return this._meetingSuggestions$p$0
	},
	_getContacts$p$0: function()
	{
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnPropertyAccessForRestrictedPermission$i(this._permissionLevel$p$0);
		if(!this._contacts$p$0)
			this._contacts$p$0=$h.Entities._getExtractedObjects$i($h.Contact,this._data$p$0,"Contacts",$h.Entities._createContact$p,true,"ContactString");
		return this._contacts$p$0
	},
	_getParcelDeliveries$p$0: function()
	{
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnPropertyAccessForRestrictedPermission$i(this._permissionLevel$p$0);
		if(!this._parcelDeliveries$p$0)
			this._parcelDeliveries$p$0=$h.Entities._getExtractedObjects$i(Object,this._data$p$0,"ParcelDeliveries",null);
		return this._parcelDeliveries$p$0
	},
	_getFlightReservations$p$0: function()
	{
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnPropertyAccessForRestrictedPermission$i(this._permissionLevel$p$0);
		if(!this._flightReservations$p$0)
			this._flightReservations$p$0=$h.Entities._getExtractedObjects$i(Object,this._data$p$0,"FlightReservations",null);
		return this._flightReservations$p$0
	}
};
$h.ReplyConstants=function(){};
$h.AsyncConstants=function(){};
window["Office"]["cast"]["item"]=Office.cast.item=function(){};
window["Office"]["cast"]["item"]["toItemRead"]=function(item)
{
	if($h.Item.isInstanceOfType(item))
		return item;
	throw Error.argumentType();
};
window["Office"]["cast"]["item"]["toItemCompose"]=function(item)
{
	if($h.ComposeItem.isInstanceOfType(item))
		return item;
	throw Error.argumentType();
};
window["Office"]["cast"]["item"]["toMessage"]=function(item)
{
	return window["Office"]["cast"]["item"]["toMessageRead"](item)
};
window["Office"]["cast"]["item"]["toMessageRead"]=function(item)
{
	if($h.Message.isInstanceOfType(item))
		return item;
	throw Error.argumentType();
};
window["Office"]["cast"]["item"]["toMessageCompose"]=function(item)
{
	if($h.MessageCompose.isInstanceOfType(item))
		return item;
	throw Error.argumentType();
};
window["Office"]["cast"]["item"]["toMeetingRequest"]=function(item)
{
	if($h.MeetingRequest.isInstanceOfType(item))
		return item;
	throw Error.argumentType();
};
window["Office"]["cast"]["item"]["toAppointment"]=function(item)
{
	return window["Office"]["cast"]["item"]["toAppointmentRead"](item)
};
window["Office"]["cast"]["item"]["toAppointmentRead"]=function(item)
{
	if($h.Appointment.isInstanceOfType(item))
		return item;
	throw Error.argumentType();
};
window["Office"]["cast"]["item"]["toAppointmentCompose"]=function(item)
{
	if($h.AppointmentCompose.isInstanceOfType(item))
		return item;
	throw Error.argumentType();
};
$h.Item=function(data)
{
	this.$$d__getBody$p$1=Function.createDelegate(this,this._getBody$p$1);
	this.$$d__getAttachments$p$1=Function.createDelegate(this,this._getAttachments$p$1);
	this.$$d__getItemClass$p$1=Function.createDelegate(this,this._getItemClass$p$1);
	this.$$d__getItemId$p$1=Function.createDelegate(this,this._getItemId$p$1);
	this.$$d__getDateTimeModified$p$1=Function.createDelegate(this,this._getDateTimeModified$p$1);
	this.$$d__getDateTimeCreated$p$1=Function.createDelegate(this,this._getDateTimeCreated$p$1);
	$h.Item.initializeBase(this,[data]);
	$h.InitialData._defineReadOnlyProperty$i(this,"dateTimeCreated",this.$$d__getDateTimeCreated$p$1);
	$h.InitialData._defineReadOnlyProperty$i(this,"dateTimeModified",this.$$d__getDateTimeModified$p$1);
	$h.InitialData._defineReadOnlyProperty$i(this,"itemId",this.$$d__getItemId$p$1);
	$h.InitialData._defineReadOnlyProperty$i(this,"itemClass",this.$$d__getItemClass$p$1);
	$h.InitialData._defineReadOnlyProperty$i(this,"attachments",this.$$d__getAttachments$p$1);
	$h.InitialData._defineReadOnlyProperty$i(this,"body",this.$$d__getBody$p$1)
};
$h.Item.prototype={
	_body$p$1: null,
	_getItemId$p$1: function()
	{
		return this._data$p$0.get__itemId$i$0()
	},
	_getItemClass$p$1: function()
	{
		return this._data$p$0.get__itemClass$i$0()
	},
	_getDateTimeCreated$p$1: function()
	{
		return this._data$p$0.get__dateTimeCreated$i$0()
	},
	_getDateTimeModified$p$1: function()
	{
		return this._data$p$0.get__dateTimeModified$i$0()
	},
	_getAttachments$p$1: function()
	{
		return this._data$p$0.get__attachments$i$0()
	},
	_getBody$p$1: function()
	{
		if(!this._body$p$1)
			this._body$p$1=new $h.Body;
		return this._body$p$1
	}
};
$h.ItemBase=function(data)
{
	this.$$d__createCustomProperties$i$0=Function.createDelegate(this,this._createCustomProperties$i$0);
	this.$$d__getNotificationMessages$p$0=Function.createDelegate(this,this._getNotificationMessages$p$0);
	this.$$d_getItemType=Function.createDelegate(this,this.getItemType);
	this._data$p$0=data;
	$h.InitialData._defineReadOnlyProperty$i(this,"itemType",this.$$d_getItemType);
	$h.InitialData._defineReadOnlyProperty$i(this,"notificationMessages",this.$$d__getNotificationMessages$p$0)
};
$h.ItemBase.prototype={
	_data$p$0: null,
	_notificationMessages$p$0: null,
	get_data: function()
	{
		return this._data$p$0
	},
	loadCustomPropertiesAsync: function()
	{
		var args=[];
		for(var $$pai_3=0; $$pai_3 < arguments.length;++$$pai_3)
			args[$$pai_3]=arguments[$$pai_3];
		var parameters=$h.CommonParameters.parse(args,true,true);
		var loadCustomProperties=new $h._loadDictionaryRequest(this.$$d__createCustomProperties$i$0,"customProperties",parameters._callback$p$0,parameters._asyncContext$p$0);
		loadCustomProperties._sendRequest$i$0(3,"LoadCustomProperties",{})
	},
	_createCustomProperties$i$0: function(data)
	{
		return new $h.CustomProperties(data)
	},
	_getNotificationMessages$p$0: function()
	{
		if(!this._notificationMessages$p$0)
			this._notificationMessages$p$0=new $h.NotificationMessages;
		return this._notificationMessages$p$0
	}
};
$h.MeetingRequest=function(data)
{
	this.$$d__getRequiredAttendees$p$3=Function.createDelegate(this,this._getRequiredAttendees$p$3);
	this.$$d__getOptionalAttendees$p$3=Function.createDelegate(this,this._getOptionalAttendees$p$3);
	this.$$d__getLocation$p$3=Function.createDelegate(this,this._getLocation$p$3);
	this.$$d__getEnd$p$3=Function.createDelegate(this,this._getEnd$p$3);
	this.$$d__getStart$p$3=Function.createDelegate(this,this._getStart$p$3);
	$h.MeetingRequest.initializeBase(this,[data]);
	$h.InitialData._defineReadOnlyProperty$i(this,"start",this.$$d__getStart$p$3);
	$h.InitialData._defineReadOnlyProperty$i(this,"end",this.$$d__getEnd$p$3);
	$h.InitialData._defineReadOnlyProperty$i(this,"location",this.$$d__getLocation$p$3);
	$h.InitialData._defineReadOnlyProperty$i(this,"optionalAttendees",this.$$d__getOptionalAttendees$p$3);
	$h.InitialData._defineReadOnlyProperty$i(this,"requiredAttendees",this.$$d__getRequiredAttendees$p$3)
};
$h.MeetingRequest.prototype={
	_getStart$p$3: function()
	{
		return this._data$p$0.get__start$i$0()
	},
	_getEnd$p$3: function()
	{
		return this._data$p$0.get__end$i$0()
	},
	_getLocation$p$3: function()
	{
		return this._data$p$0.get__location$i$0()
	},
	_getOptionalAttendees$p$3: function()
	{
		return this._data$p$0.get__cc$i$0()
	},
	_getRequiredAttendees$p$3: function()
	{
		return this._data$p$0.get__to$i$0()
	}
};
$h.MeetingSuggestion=function(data, dateTimeSent)
{
	this.$$d__getEndTime$p$0=Function.createDelegate(this,this._getEndTime$p$0);
	this.$$d__getStartTime$p$0=Function.createDelegate(this,this._getStartTime$p$0);
	this.$$d__getSubject$p$0=Function.createDelegate(this,this._getSubject$p$0);
	this.$$d__getLocation$p$0=Function.createDelegate(this,this._getLocation$p$0);
	this.$$d__getAttendees$p$0=Function.createDelegate(this,this._getAttendees$p$0);
	this.$$d__getMeetingString$p$0=Function.createDelegate(this,this._getMeetingString$p$0);
	this._data$p$0=data;
	this._dateTimeSent$p$0=dateTimeSent;
	$h.InitialData._defineReadOnlyProperty$i(this,"meetingString",this.$$d__getMeetingString$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"attendees",this.$$d__getAttendees$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"location",this.$$d__getLocation$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"subject",this.$$d__getSubject$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"start",this.$$d__getStartTime$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"end",this.$$d__getEndTime$p$0)
};
$h.MeetingSuggestion.prototype={
	_dateTimeSent$p$0: null,
	_data$p$0: null,
	_attendees$p$0: null,
	_getMeetingString$p$0: function()
	{
		return this._data$p$0["MeetingString"]
	},
	_getLocation$p$0: function()
	{
		return this._data$p$0["Location"]
	},
	_getSubject$p$0: function()
	{
		return this._data$p$0["Subject"]
	},
	_getStartTime$p$0: function()
	{
		var time=this._createDateTimeFromParameter$p$0("StartTime");
		var resolvedTime=$h.MeetingSuggestionTimeDecoder.resolve(time,this._dateTimeSent$p$0);
		if(resolvedTime["getTime"]() !==time["getTime"]())
			return window["OSF"]["DDA"]["OutlookAppOm"]._instance$p["convertToUtcClientTime"](window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._dateToDictionary$i$0(resolvedTime));
		return time
	},
	_getEndTime$p$0: function()
	{
		var time=this._createDateTimeFromParameter$p$0("EndTime");
		var resolvedTime=$h.MeetingSuggestionTimeDecoder.resolve(time,this._dateTimeSent$p$0);
		if(resolvedTime["getTime"]() !==time["getTime"]())
			return window["OSF"]["DDA"]["OutlookAppOm"]._instance$p["convertToUtcClientTime"](window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._dateToDictionary$i$0(resolvedTime));
		return time
	},
	_createDateTimeFromParameter$p$0: function(keyName)
	{
		var dateTimeString=this._data$p$0[keyName];
		if(!dateTimeString)
			return null;
		return new Date(dateTimeString)
	},
	_getAttendees$p$0: function()
	{
		if(!this._attendees$p$0)
		{
			var $$t_1=this;
			this._attendees$p$0=$h.Entities._getExtractedObjects$i($h.EmailAddressDetails,this._data$p$0,"Attendees",function(data)
			{
				return $h.EmailAddressDetails._createFromEmailUserDictionary$i(data)
			})
		}
		return this._attendees$p$0
	}
};
$h.MeetingSuggestionTimeDecoder=function(){};
$h.MeetingSuggestionTimeDecoder.resolve=function(inTime, sentTime)
{
	if(!sentTime)
		return inTime;
	try
	{
		var tod;
		var outDate;
		var extractedDate;
		var sentDate=new Date(sentTime["getFullYear"](),sentTime["getMonth"](),sentTime["getDate"](),0,0,0,0);
		var $$t_7,
			$$t_8,
			$$t_9;
		if(!($$t_9=$h.MeetingSuggestionTimeDecoder._decode$p(inTime,$$t_7={val: extractedDate},$$t_8={val: tod}),extractedDate=$$t_7["val"],tod=$$t_8["val"],$$t_9))
			return inTime;
		else
		{
			if($h._preciseDate.isInstanceOfType(extractedDate))
				outDate=$h.MeetingSuggestionTimeDecoder._resolvePreciseDate$p(sentDate,extractedDate);
			else if($h._relativeDate.isInstanceOfType(extractedDate))
				outDate=$h.MeetingSuggestionTimeDecoder._resolveRelativeDate$p(sentDate,extractedDate);
			else
				outDate=sentDate;
			if(window["isNaN"](outDate["getTime"]()))
				return sentTime;
			outDate["setMilliseconds"](outDate["getMilliseconds"]()+tod);
			return outDate
		}
	}
	catch($$e_6)
	{
		return sentTime
	}
};
$h.MeetingSuggestionTimeDecoder._isNullOrUndefined$i=function(value)
{
	return null===value || value===undefined
};
$h.MeetingSuggestionTimeDecoder._resolvePreciseDate$p=function(sentDate, precise)
{
	var year=precise._year$i$1;
	var month=!precise._month$i$1 ? sentDate["getMonth"]() : precise._month$i$1 - 1;
	var day=precise._day$i$1;
	if(!day)
		return sentDate;
	var candidate;
	if($h.MeetingSuggestionTimeDecoder._isNullOrUndefined$i(year))
	{
		candidate=new Date(sentDate["getFullYear"](),month,day);
		if(candidate["getTime"]() < sentDate["getTime"]())
			candidate=new Date(sentDate["getFullYear"]()+1,month,day)
	}
	else
		candidate=new Date(year < 50 ? 2e3+year : 1900+year,month,day);
	if(candidate["getMonth"]() !==month)
		return sentDate;
	return candidate
};
$h.MeetingSuggestionTimeDecoder._resolveRelativeDate$p=function(sentDate, relative)
{
	var date;
	switch(relative._unit$i$1)
	{
		case 0:
			date=new Date(sentDate["getFullYear"](),sentDate["getMonth"](),sentDate["getDate"]());
			date["setDate"](date["getDate"]()+relative._offset$i$1);
			return date;
		case 5:
			return $h.MeetingSuggestionTimeDecoder._findBestDateForWeekDate$p(sentDate,relative._offset$i$1,relative._tag$i$1);
		case 2:
			var days=1;
			switch(relative._modifier$i$1)
			{
				case 1:
					break;
				case 2:
					days=16;
					break;
				default:
					if(!relative._offset$i$1)
						days=sentDate["getDate"]();
					break
			}
			date=new Date(sentDate["getFullYear"](),sentDate["getMonth"](),days);
			date["setMonth"](date["getMonth"]()+relative._offset$i$1);
			if(date["getTime"]() < sentDate["getTime"]())
				date["setDate"](date["getDate"]()+sentDate["getDate"]() - 1);
			return date;
		case 1:
			date=new Date(sentDate["getFullYear"](),sentDate["getMonth"](),sentDate["getDate"]());
			date["setDate"](sentDate["getDate"]()+7 * relative._offset$i$1);
			if(relative._modifier$i$1===1 || !relative._modifier$i$1)
			{
				date["setDate"](date["getDate"]()+1 - date["getDay"]());
				if(date["getTime"]() < sentDate["getTime"]())
					return sentDate;
				return date
			}
			else if(relative._modifier$i$1===2)
			{
				date["setDate"](date["getDate"]()+5 - date["getDay"]());
				return date
			}
			break;
		case 4:
			return $h.MeetingSuggestionTimeDecoder._findBestDateForWeekOfMonthDate$p(sentDate,relative);
		case 3:
			if(relative._offset$i$1 > 0)
				return new Date(sentDate["getFullYear"]()+relative._offset$i$1,0,1);
			break;
		default:
			break
	}
	return sentDate
};
$h.MeetingSuggestionTimeDecoder._findBestDateForWeekDate$p=function(sentDate, offset, tag)
{
	if(offset > -5 && offset < 5)
	{
		var dayOfWeek;
		var days;
		dayOfWeek=(tag+6) % 7+1;
		days=7 * offset+(dayOfWeek - sentDate["getDay"]());
		sentDate["setDate"](sentDate["getDate"]()+days);
		return sentDate
	}
	else
	{
		var days=(tag - sentDate["getDay"]()) % 7;
		if(days < 0)
			days+=7;
		sentDate["setDate"](sentDate["getDate"]()+days);
		return sentDate
	}
};
$h.MeetingSuggestionTimeDecoder._findBestDateForWeekOfMonthDate$p=function(sentDate, relative)
{
	var date;
	var firstDay;
	var newDate;
	date=sentDate;
	if(relative._tag$i$1 <=0 || relative._tag$i$1 > 12 || relative._offset$i$1 <=0 || relative._offset$i$1 > 5)
		return sentDate;
	var monthOffset=(12+relative._tag$i$1 - date["getMonth"]() - 1) % 12;
	firstDay=new Date(date["getFullYear"](),date["getMonth"]()+monthOffset,1);
	if(relative._modifier$i$1===1)
		if(relative._offset$i$1===1 && firstDay["getDay"]() !==6 && firstDay["getDay"]())
			return firstDay;
		else
		{
			newDate=new Date(firstDay["getFullYear"](),firstDay["getMonth"](),firstDay["getDate"]());
			newDate["setDate"](newDate["getDate"]()+(7+(1 - firstDay["getDay"]())) % 7);
			if(firstDay["getDay"]() !==6 && firstDay["getDay"]() && firstDay["getDay"]() !==1)
				newDate["setDate"](newDate["getDate"]() - 7);
			newDate["setDate"](newDate["getDate"]()+7 * (relative._offset$i$1 - 1));
			if(newDate["getMonth"]()+1 !==relative._tag$i$1)
				return sentDate;
			return newDate
		}
	else
	{
		newDate=new Date(firstDay["getFullYear"](),firstDay["getMonth"](),$h.MeetingSuggestionTimeDecoder._daysInMonth$p(firstDay["getMonth"](),firstDay["getFullYear"]()));
		var offset=1 - newDate["getDay"]();
		if(offset > 0)
			offset=offset - 7;
		newDate["setDate"](newDate["getDate"]()+offset);
		newDate["setDate"](newDate["getDate"]()+7 * (1 - relative._offset$i$1));
		if(newDate["getMonth"]()+1 !==relative._tag$i$1)
			if(firstDay["getDay"]() !==6 && firstDay["getDay"]())
				return firstDay;
			else
				return sentDate;
		else
			return newDate
	}
};
$h.MeetingSuggestionTimeDecoder._decode$p=function(inDate, date, time)
{
	var DateValueMask=32767;
	date.val=null;
	time.val=0;
	if(!inDate)
		return false;
	time.val=$h.MeetingSuggestionTimeDecoder._getTimeOfDayInMillisecondsUTC$p(inDate);
	var inDateAtMidnight=inDate["getTime"]() - time.val;
	var value=(inDateAtMidnight - $h.MeetingSuggestionTimeDecoder._baseDate$p["getTime"]()) / 864e5;
	if(value < 0)
		return false;
	else if(value >=262144)
		return false;
	else
	{
		var type=value >> 15;
		value=value & DateValueMask;
		switch(type)
		{
			case 0:
				return $h.MeetingSuggestionTimeDecoder._decodePreciseDate$p(value,date);
			case 1:
				return $h.MeetingSuggestionTimeDecoder._decodeRelativeDate$p(value,date);
			default:
				return false
		}
	}
};
$h.MeetingSuggestionTimeDecoder._decodePreciseDate$p=function(value, date)
{
	var c_SubTypeMask=7;
	var c_MonthMask=15;
	var c_DayMask=31;
	var c_YearMask=127;
	var year=null;
	var month=0;
	var day=0;
	date.val=null;
	var subType=value >> 12 & c_SubTypeMask;
	if((subType & 4)===4)
	{
		year=value >> 5 & c_YearMask;
		if((subType & 2)===2)
		{
			if((subType & 1)===1)
				return false;
			month=value >> 1 & c_MonthMask
		}
	}
	else
	{
		if((subType & 2)===2)
			month=value >> 8 & c_MonthMask;
		if((subType & 1)===1)
			day=value >> 3 & c_DayMask
	}
	date.val=new $h._preciseDate(day,month,year);
	return true
};
$h.MeetingSuggestionTimeDecoder._decodeRelativeDate$p=function(value, date)
{
	var TagMask=15;
	var OffsetMask=63;
	var UnitMask=7;
	var ModifierMask=3;
	var tag=value & TagMask;
	value >>=4;
	var offset=$h.MeetingSuggestionTimeDecoder._fromComplement$p(value & OffsetMask,6);
	value >>=6;
	var unit=value & UnitMask;
	value >>=3;
	var modifier=value & ModifierMask;
	try
	{
		date.val=new $h._relativeDate(modifier,offset,unit,tag);
		return true
	}
	catch($$e_A)
	{
		date.val=null;
		return false
	}
};
$h.MeetingSuggestionTimeDecoder._fromComplement$p=function(value, n)
{
	var signed=1 << n - 1;
	var mask=(1 << n) - 1;
	if((value & signed)===signed)
		return-((value ^ mask)+1);
	else
		return value
};
$h.MeetingSuggestionTimeDecoder._daysInMonth$p=function(month, year)
{
	return 32 - new Date(year,month,32)["getDate"]()
};
$h.MeetingSuggestionTimeDecoder._getTimeOfDayInMillisecondsUTC$p=function(inputTime)
{
	var timeOfDay=0;
	timeOfDay+=inputTime["getUTCHours"]() * 3600;
	timeOfDay+=inputTime["getUTCMinutes"]() * 60;
	timeOfDay+=inputTime["getUTCSeconds"]();
	timeOfDay *=1e3;
	timeOfDay+=inputTime["getUTCMilliseconds"]();
	return timeOfDay
};
$h._extractedDate=function(){};
$h._preciseDate=function(day, month, year)
{
	$h._preciseDate.initializeBase(this);
	if(day < 0 || day > 31)
		throw Error.argumentOutOfRange("day");
	if(month < 0 || month > 12)
		throw Error.argumentOutOfRange("month");
	this._day$i$1=day;
	this._month$i$1=month;
	if(!$h.MeetingSuggestionTimeDecoder._isNullOrUndefined$i(year))
	{
		if(!month && day)
			throw Error.argument("Invalid arguments");
		if(year < 0 || year > 2099)
			throw Error.argumentOutOfRange("year");
		this._year$i$1=year % 100
	}
	else if(!this._month$i$1 && !this._day$i$1)
		throw Error.argument("Invalid datetime");
};
$h._preciseDate.prototype={
	_day$i$1: 0,
	_month$i$1: 0,
	_year$i$1: null
};
$h._relativeDate=function(modifier, offset, unit, tag)
{
	$h._relativeDate.initializeBase(this);
	if(offset < -32 || offset > 31)
		throw Error.argumentOutOfRange("offset");
	if(tag < 0 || tag > 15)
		throw Error.argumentOutOfRange("tag");
	if(!unit && offset < 0)
		throw Error.argument("unit & offset do not form a valid date");
	this._modifier$i$1=modifier;
	this._offset$i$1=offset;
	this._unit$i$1=unit;
	this._tag$i$1=tag
};
$h._relativeDate.prototype={
	_modifier$i$1: 0,
	_offset$i$1: 0,
	_unit$i$1: 0,
	_tag$i$1: 0
};
$h.Message=function(dataDictionary)
{
	this.$$d__getConversationId$p$2=Function.createDelegate(this,this._getConversationId$p$2);
	this.$$d__getInternetMessageId$p$2=Function.createDelegate(this,this._getInternetMessageId$p$2);
	this.$$d__getCc$p$2=Function.createDelegate(this,this._getCc$p$2);
	this.$$d__getTo$p$2=Function.createDelegate(this,this._getTo$p$2);
	this.$$d__getFrom$p$2=Function.createDelegate(this,this._getFrom$p$2);
	this.$$d__getSender$p$2=Function.createDelegate(this,this._getSender$p$2);
	this.$$d__getNormalizedSubject$p$2=Function.createDelegate(this,this._getNormalizedSubject$p$2);
	this.$$d__getSubject$p$2=Function.createDelegate(this,this._getSubject$p$2);
	$h.Message.initializeBase(this,[dataDictionary]);
	$h.InitialData._defineReadOnlyProperty$i(this,"subject",this.$$d__getSubject$p$2);
	$h.InitialData._defineReadOnlyProperty$i(this,"normalizedSubject",this.$$d__getNormalizedSubject$p$2);
	$h.InitialData._defineReadOnlyProperty$i(this,"sender",this.$$d__getSender$p$2);
	$h.InitialData._defineReadOnlyProperty$i(this,"from",this.$$d__getFrom$p$2);
	$h.InitialData._defineReadOnlyProperty$i(this,"to",this.$$d__getTo$p$2);
	$h.InitialData._defineReadOnlyProperty$i(this,"cc",this.$$d__getCc$p$2);
	$h.InitialData._defineReadOnlyProperty$i(this,"internetMessageId",this.$$d__getInternetMessageId$p$2);
	$h.InitialData._defineReadOnlyProperty$i(this,"conversationId",this.$$d__getConversationId$p$2)
};
$h.Message.prototype={
	getEntities: function()
	{
		return this._data$p$0._getEntities$i$0()
	},
	getEntitiesByType: function(entityType)
	{
		return this._data$p$0._getEntitiesByType$i$0(entityType)
	},
	getFilteredEntitiesByName: function(name)
	{
		return this._data$p$0._getFilteredEntitiesByName$i$0(name)
	},
	getRegExMatches: function()
	{
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"getRegExMatches");
		return this._data$p$0._getRegExMatches$i$0()
	},
	getRegExMatchesByName: function(name)
	{
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"getRegExMatchesByName");
		return this._data$p$0._getRegExMatchesByName$i$0(name)
	},
	displayReplyForm: function(obj)
	{
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._displayReplyForm$i$0(obj)
	},
	displayReplyAllForm: function(obj)
	{
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._displayReplyAllForm$i$0(obj)
	},
	getItemType: function()
	{
		return window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ItemType"]["Message"]
	},
	_getSubject$p$2: function()
	{
		return this._data$p$0.get__subject$i$0()
	},
	_getNormalizedSubject$p$2: function()
	{
		return this._data$p$0.get__normalizedSubject$i$0()
	},
	_getSender$p$2: function()
	{
		return this._data$p$0.get__sender$i$0()
	},
	_getFrom$p$2: function()
	{
		return this._data$p$0.get__from$i$0()
	},
	_getTo$p$2: function()
	{
		return this._data$p$0.get__to$i$0()
	},
	_getCc$p$2: function()
	{
		return this._data$p$0.get__cc$i$0()
	},
	_getInternetMessageId$p$2: function()
	{
		return this._data$p$0.get__internetMessageId$i$0()
	},
	_getConversationId$p$2: function()
	{
		return this._data$p$0.get__conversationId$i$0()
	}
};
$h.MessageCompose=function(data)
{
	this.$$d__getConversationId$p$2=Function.createDelegate(this,this._getConversationId$p$2);
	this.$$d__getBcc$p$2=Function.createDelegate(this,this._getBcc$p$2);
	this.$$d__getCc$p$2=Function.createDelegate(this,this._getCc$p$2);
	this.$$d__getTo$p$2=Function.createDelegate(this,this._getTo$p$2);
	$h.MessageCompose.initializeBase(this,[data]);
	$h.InitialData._defineReadOnlyProperty$i(this,"to",this.$$d__getTo$p$2);
	$h.InitialData._defineReadOnlyProperty$i(this,"cc",this.$$d__getCc$p$2);
	$h.InitialData._defineReadOnlyProperty$i(this,"bcc",this.$$d__getBcc$p$2);
	$h.InitialData._defineReadOnlyProperty$i(this,"conversationId",this.$$d__getConversationId$p$2)
};
$h.MessageCompose.prototype={
	_to$p$2: null,
	_cc$p$2: null,
	_bcc$p$2: null,
	getItemType: function()
	{
		return window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ItemType"]["Message"]
	},
	_getTo$p$2: function()
	{
		this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
		if(!this._to$p$2)
			this._to$p$2=new $h.ComposeRecipient(0,"to");
		return this._to$p$2
	},
	_getCc$p$2: function()
	{
		this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
		if(!this._cc$p$2)
			this._cc$p$2=new $h.ComposeRecipient(1,"cc");
		return this._cc$p$2
	},
	_getBcc$p$2: function()
	{
		this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
		if(!this._bcc$p$2)
			this._bcc$p$2=new $h.ComposeRecipient(2,"bcc");
		return this._bcc$p$2
	},
	_getConversationId$p$2: function()
	{
		return this._data$p$0.get__conversationId$i$0()
	}
};
$h.NotificationMessages=function(){};
$h.NotificationMessages._mapToHostItemNotificationMessageType$p=function(dataToHost)
{
	var notificationType;
	var hostItemNotificationMessageType;
	notificationType=dataToHost["type"];
	if(notificationType===window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ItemNotificationMessageType"]["ProgressIndicator"])
		hostItemNotificationMessageType=1;
	else if(notificationType===window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ItemNotificationMessageType"]["InformationalMessage"])
		hostItemNotificationMessageType=0;
	else if(notificationType===window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ItemNotificationMessageType"]["ErrorMessage"])
		hostItemNotificationMessageType=2;
	else
		throw Error.argument("type");
	dataToHost["type"]=hostItemNotificationMessageType
};
$h.NotificationMessages._validateKey$p=function(key)
{
	if(!$h.ScriptHelpers.isNonEmptyString(key))
		throw Error.argument("key");
	window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(key["length"],0,32,"key")
};
$h.NotificationMessages._validateDictionary$p=function(dictionary)
{
	if(!$h.ScriptHelpers.isNonEmptyString(dictionary["type"]))
		throw Error.argument("type");
	if(dictionary["type"]===window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ItemNotificationMessageType"]["InformationalMessage"])
	{
		if(!$h.ScriptHelpers.isNonEmptyString(dictionary["icon"]))
			throw Error.argument("icon");
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(dictionary["icon"]["length"],0,32,"icon");
		if($h.ScriptHelpers.isUndefined(dictionary["persistent"]))
			throw Error.argument("persistent");
		if(!Boolean.isInstanceOfType(dictionary["persistent"]))
			throw Error.argumentType("persistent",Object["getType"](dictionary["persistent"]),Boolean);
	}
	else
	{
		if(!$h.ScriptHelpers.isUndefined(dictionary["icon"]))
			throw Error.argument("icon");
		if(!$h.ScriptHelpers.isUndefined(dictionary["persistent"]))
			throw Error.argument("persistent");
	}
	if(!$h.ScriptHelpers.isNonEmptyString(dictionary["message"]))
		throw Error.argument("message");
	window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(dictionary["message"]["length"],0,150,"message")
};
$h.NotificationMessages.prototype={
	addAsync: function(key, dictionary)
	{
		var args=[];
		for(var $$pai_5=2; $$pai_5 < arguments.length;++$$pai_5)
			args[$$pai_5 - 2]=arguments[$$pai_5];
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(0,"NotificationMessages.addAsync");
		var commonParameters=$h.CommonParameters.parse(args,false);
		$h.NotificationMessages._validateKey$p(key);
		$h.NotificationMessages._validateDictionary$p(dictionary);
		var dataToHost={};
		dataToHost=$h.ScriptHelpers.deepClone(dictionary);
		dataToHost["key"]=key;
		$h.NotificationMessages._mapToHostItemNotificationMessageType$p(dataToHost);
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(33,dataToHost,null,commonParameters._asyncContext$p$0,commonParameters._callback$p$0)
	},
	getAllAsync: function()
	{
		var args=[];
		for(var $$pai_2=0; $$pai_2 < arguments.length;++$$pai_2)
			args[$$pai_2]=arguments[$$pai_2];
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(0,"NotificationMessages.getAllAsync");
		var commonParameters=$h.CommonParameters.parse(args,true);
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(34,null,null,commonParameters._asyncContext$p$0,commonParameters._callback$p$0)
	},
	replaceAsync: function(key, dictionary)
	{
		var args=[];
		for(var $$pai_5=2; $$pai_5 < arguments.length;++$$pai_5)
			args[$$pai_5 - 2]=arguments[$$pai_5];
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(0,"NotificationMessages.replaceAsync");
		var commonParameters=$h.CommonParameters.parse(args,false);
		$h.NotificationMessages._validateKey$p(key);
		$h.NotificationMessages._validateDictionary$p(dictionary);
		var dataToHost={};
		dataToHost=$h.ScriptHelpers.deepClone(dictionary);
		dataToHost["key"]=key;
		$h.NotificationMessages._mapToHostItemNotificationMessageType$p(dataToHost);
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(35,dataToHost,null,commonParameters._asyncContext$p$0,commonParameters._callback$p$0)
	},
	removeAsync: function(key)
	{
		var args=[];
		for(var $$pai_4=1; $$pai_4 < arguments.length;++$$pai_4)
			args[$$pai_4 - 1]=arguments[$$pai_4];
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(0,"NotificationMessages.removeAsync");
		var commonParameters=$h.CommonParameters.parse(args,false);
		$h.NotificationMessages._validateKey$p(key);
		var dataToHost={key: key};
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(36,dataToHost,null,commonParameters._asyncContext$p$0,commonParameters._callback$p$0)
	}
};
$h.OutlookErrorManager=function(){};
$h.OutlookErrorManager.getErrorArgs=function(errorCode)
{
	if(!$h.OutlookErrorManager._isInitialized$p)
		$h.OutlookErrorManager._initialize$p();
	return window["OSF"]["DDA"]["ErrorCodeManager"]["getErrorArgs"](errorCode)
};
$h.OutlookErrorManager._initialize$p=function()
{
	$h.OutlookErrorManager._addErrorMessage$p(9e3,"AttachmentSizeExceeded",window["_u"]["ExtensibilityStrings"]["l_AttachmentExceededSize_Text"]);
	$h.OutlookErrorManager._addErrorMessage$p(9001,"NumberOfAttachmentsExceeded",window["_u"]["ExtensibilityStrings"]["l_ExceededMaxNumberOfAttachments_Text"]);
	$h.OutlookErrorManager._addErrorMessage$p(9002,"InternalFormatError",window["_u"]["ExtensibilityStrings"]["l_InternalFormatError_Text"]);
	$h.OutlookErrorManager._addErrorMessage$p(9003,"InvalidAttachmentId",window["_u"]["ExtensibilityStrings"]["l_InvalidAttachmentId_Text"]);
	$h.OutlookErrorManager._addErrorMessage$p(9004,"InvalidAttachmentPath",window["_u"]["ExtensibilityStrings"]["l_InvalidAttachmentPath_Text"]);
	$h.OutlookErrorManager._addErrorMessage$p(9005,"CannotAddAttachmentBeforeUpgrade",window["_u"]["ExtensibilityStrings"]["l_CannotAddAttachmentBeforeUpgrade_Text"]);
	$h.OutlookErrorManager._addErrorMessage$p(9006,"AttachmentDeletedBeforeUploadCompletes",window["_u"]["ExtensibilityStrings"]["l_AttachmentDeletedBeforeUploadCompletes_Text"]);
	$h.OutlookErrorManager._addErrorMessage$p(9007,"AttachmentUploadGeneralFailure",window["_u"]["ExtensibilityStrings"]["l_AttachmentUploadGeneralFailure_Text"]);
	$h.OutlookErrorManager._addErrorMessage$p(9008,"AttachmentToDeleteDoesNotExist",window["_u"]["ExtensibilityStrings"]["l_DeleteAttachmentDoesNotExist_Text"]);
	$h.OutlookErrorManager._addErrorMessage$p(9009,"AttachmentDeleteGeneralFailure",window["_u"]["ExtensibilityStrings"]["l_AttachmentDeleteGeneralFailure_Text"]);
	$h.OutlookErrorManager._addErrorMessage$p(9010,"InvalidEndTime",window["_u"]["ExtensibilityStrings"]["l_InvalidEndTime_Text"]);
	$h.OutlookErrorManager._addErrorMessage$p(9011,"HtmlSanitizationFailure",window["_u"]["ExtensibilityStrings"]["l_HtmlSanitizationFailure_Text"]);
	$h.OutlookErrorManager._addErrorMessage$p(9012,"NumberOfRecipientsExceeded",String.format(window["_u"]["ExtensibilityStrings"]["l_NumberOfRecipientsExceeded_Text"],500));
	$h.OutlookErrorManager._addErrorMessage$p(9013,"NoValidRecipientsProvided",window["_u"]["ExtensibilityStrings"]["l_NoValidRecipientsProvided_Text"]);
	$h.OutlookErrorManager._addErrorMessage$p(9014,"CursorPositionChanged",window["_u"]["ExtensibilityStrings"]["l_CursorPositionChanged_Text"]);
	$h.OutlookErrorManager._addErrorMessage$p(9016,"InvalidSelection",window["_u"]["ExtensibilityStrings"]["l_InvalidSelection_Text"]);
	$h.OutlookErrorManager._addErrorMessage$p(9017,"AccessRestricted","");
	$h.OutlookErrorManager._addErrorMessage$p(9018,"GenericTokenError","");
	$h.OutlookErrorManager._addErrorMessage$p(9019,"GenericSettingsError","");
	$h.OutlookErrorManager._addErrorMessage$p(9020,"GenericResponseError","");
	$h.OutlookErrorManager._addErrorMessage$p(9021,"SaveError",window["_u"]["ExtensibilityStrings"]["l_SaveError_Text"]);
	$h.OutlookErrorManager._addErrorMessage$p(9022,"MessageInDifferentStoreError",window["_u"]["ExtensibilityStrings"]["l_MessageInDifferentStoreError_Text"]);
	$h.OutlookErrorManager._addErrorMessage$p(9023,"DuplicateNotificationKey",window["_u"]["ExtensibilityStrings"]["l_DuplicateNotificationKey_Text"]);
	$h.OutlookErrorManager._addErrorMessage$p(9024,"NotificationKeyNotFound",window["_u"]["ExtensibilityStrings"]["l_NotificationKeyNotFound_Text"]);
	$h.OutlookErrorManager._addErrorMessage$p(9025,"NumberOfNotificationsExceeded",window["_u"]["ExtensibilityStrings"]["l_NumberOfNotificationsExceeded_Text"]);
	$h.OutlookErrorManager._addErrorMessage$p(9026,"PersistedNotificationArrayReadError",window["_u"]["ExtensibilityStrings"]["l_PersistedNotificationArrayReadError_Text"]);
	$h.OutlookErrorManager._addErrorMessage$p(9027,"PersistedNotificationArraySaveError",window["_u"]["ExtensibilityStrings"]["l_PersistedNotificationArraySaveError_Text"]);
	$h.OutlookErrorManager._addErrorMessage$p(9028,"CannotPersistPropertyInUnsavedDraftError",window["_u"]["ExtensibilityStrings"]["l_CannotPersistPropertyInUnsavedDraftError_Text"]);
	$h.OutlookErrorManager._isInitialized$p=true
};
$h.OutlookErrorManager._addErrorMessage$p=function(errorCode, errorName, errorMessage)
{
	window["OSF"]["DDA"]["ErrorCodeManager"]["addErrorMessage"](errorCode,{
		name: errorName,
		message: errorMessage
	})
};
$h.OutlookErrorManager.OutlookErrorCodes=function(){};
$h.OutlookErrorManager.OsfDdaErrorCodes=function(){};
$h.PhoneNumber=function(data)
{
	this.$$d__getPhoneType$p$0=Function.createDelegate(this,this._getPhoneType$p$0);
	this.$$d__getOriginalPhoneString$p$0=Function.createDelegate(this,this._getOriginalPhoneString$p$0);
	this.$$d__getPhoneString$p$0=Function.createDelegate(this,this._getPhoneString$p$0);
	this._data$p$0=data;
	$h.InitialData._defineReadOnlyProperty$i(this,"phoneString",this.$$d__getPhoneString$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"originalPhoneString",this.$$d__getOriginalPhoneString$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"type",this.$$d__getPhoneType$p$0)
};
$h.PhoneNumber.prototype={
	_data$p$0: null,
	_getPhoneString$p$0: function()
	{
		return this._data$p$0["PhoneString"]
	},
	_getOriginalPhoneString$p$0: function()
	{
		return this._data$p$0["OriginalPhoneString"]
	},
	_getPhoneType$p$0: function()
	{
		return this._data$p$0["Type"]
	}
};
$h.TaskSuggestion=function(data)
{
	this.$$d__getAssignees$p$0=Function.createDelegate(this,this._getAssignees$p$0);
	this.$$d__getTaskString$p$0=Function.createDelegate(this,this._getTaskString$p$0);
	this._data$p$0=data;
	$h.InitialData._defineReadOnlyProperty$i(this,"taskString",this.$$d__getTaskString$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"assignees",this.$$d__getAssignees$p$0)
};
$h.TaskSuggestion.prototype={
	_data$p$0: null,
	_assignees$p$0: null,
	_getTaskString$p$0: function()
	{
		return this._data$p$0["TaskString"]
	},
	_getAssignees$p$0: function()
	{
		if(!this._assignees$p$0)
		{
			var $$t_1=this;
			this._assignees$p$0=$h.Entities._getExtractedObjects$i($h.EmailAddressDetails,this._data$p$0,"Assignees",function(data)
			{
				return $h.EmailAddressDetails._createFromEmailUserDictionary$i(data)
			})
		}
		return this._assignees$p$0
	}
};
$h.UserProfile=function(data)
{
	this.$$d__getUserProfileType$p$0=Function.createDelegate(this,this._getUserProfileType$p$0);
	this.$$d__getTimeZone$p$0=Function.createDelegate(this,this._getTimeZone$p$0);
	this.$$d__getEmailAddress$p$0=Function.createDelegate(this,this._getEmailAddress$p$0);
	this.$$d__getDisplayName$p$0=Function.createDelegate(this,this._getDisplayName$p$0);
	this._data$p$0=data;
	$h.InitialData._defineReadOnlyProperty$i(this,"displayName",this.$$d__getDisplayName$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"emailAddress",this.$$d__getEmailAddress$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"timeZone",this.$$d__getTimeZone$p$0);
	$h.InitialData._defineReadOnlyProperty$i(this,"type",this.$$d__getUserProfileType$p$0)
};
$h.UserProfile.prototype={
	_data$p$0: null,
	_getUserProfileType$p$0: function()
	{
		return this._data$p$0.get__userProfileType$i$0()
	},
	_getDisplayName$p$0: function()
	{
		return this._data$p$0.get__userDisplayName$i$0()
	},
	_getEmailAddress$p$0: function()
	{
		return this._data$p$0.get__userEmailAddress$i$0()
	},
	_getTimeZone$p$0: function()
	{
		return this._data$p$0.get__userTimeZone$i$0()
	}
};
$h.OutlookDispid=function(){};
$h.OutlookDispid.prototype={
	owaOnlyMethod: 0,
	getInitialData: 1,
	getUserIdentityToken: 2,
	loadCustomProperties: 3,
	saveCustomProperties: 4,
	ewsRequest: 5,
	displayNewAppointmentForm: 7,
	displayMessageForm: 8,
	displayAppointmentForm: 9,
	displayReplyForm: 10,
	displayReplyAllForm: 11,
	getCallbackToken: 12,
	bodySetSelectedDataAsync: 13,
	getBodyTypeAsync: 14,
	getRecipientsAsync: 15,
	addFileAttachmentAsync: 16,
	setSubjectAsync: 17,
	getSubjectAsync: 18,
	addItemAttachmentAsync: 19,
	removeAttachmentAsync: 20,
	setRecipientsAsync: 21,
	addRecipientsAsync: 22,
	bodyPrependAsync: 23,
	getTimeAsync: 24,
	setTimeAsync: 25,
	getLocationAsync: 26,
	setLocationAsync: 27,
	getSelectedDataAsync: 28,
	setSelectedDataAsync: 29,
	displayReplyFormWithAttachments: 30,
	displayReplyAllFormWithAttachments: 31,
	saveAsync: 32,
	addNotficationMessageAsync: 33,
	getAllNotficationMessagesAsync: 34,
	replaceNotficationMessageAsync: 35,
	removeNotficationMessageAsync: 36,
	getBodyAsync: 37,
	setBodyAsync: 38,
	appCommands1: 39,
	registerConsentAsync: 40,
	close: 41,
	closeApp: 42,
	displayContactCardAsync: 43,
	displayNewMessageForm: 44,
	appCommands2: 94,
	trackCtq: 400,
	recordTrace: 401,
	recordDataPoint: 402,
	windowOpenOverrideHandler: 403,
	saveSettingsRequest: 404
};
$h.OutlookDispid.registerEnum("$h.OutlookDispid",false);
$h.RequestState=function(){};
$h.RequestState.prototype={
	unsent: 0,
	opened: 1,
	headersReceived: 2,
	loading: 3,
	done: 4
};
$h.RequestState.registerEnum("$h.RequestState",false);
$h.CommonParameters=function(options, callback, asyncContext)
{
	this._options$p$0=options;
	this._callback$p$0=callback;
	this._asyncContext$p$0=asyncContext
};
$h.CommonParameters.parse=function(args, isCallbackRequired, tryLegacy)
{
	var legacyParameters;
	var $$t_8,
		$$t_9;
	if(tryLegacy && ($$t_9=$h.CommonParameters._tryParseLegacy$p(args,$$t_8={val: legacyParameters}),legacyParameters=$$t_8["val"],$$t_9))
		return legacyParameters;
	var argsLength=args.length;
	var options=null;
	var callback=null;
	var asyncContext=null;
	if(argsLength===1)
		if(Function.isInstanceOfType(args[0]))
			callback=args[0];
		else if(Object.isInstanceOfType(args[0]))
			options=args[0];
		else
			throw Error.argumentType();
	else if(argsLength===2)
	{
		if(!Object.isInstanceOfType(args[0]))
			throw Error.argument("options");
		if(!Function.isInstanceOfType(args[1]))
			throw Error.argument("callback");
		options=args[0];
		callback=args[1]
	}
	else if(argsLength)
		throw Error.parameterCount(window["_u"]["ExtensibilityStrings"]["l_ParametersNotAsExpected_Text"]);
	if(isCallbackRequired && !callback)
		throw Error.argumentNull("callback");
	if(options && !$h.ScriptHelpers.isNullOrUndefined(options["asyncContext"]))
		asyncContext=options["asyncContext"];
	return new $h.CommonParameters(options,callback,asyncContext)
};
$h.CommonParameters._tryParseLegacy$p=function(args, commonParameters)
{
	commonParameters.val=null;
	var argsLength=args.length;
	var callback=null;
	var userContext=null;
	if(!argsLength || argsLength > 2)
		return false;
	if(!Function.isInstanceOfType(args[0]))
		return false;
	callback=args[0];
	if(argsLength > 1)
		userContext=args[1];
	commonParameters.val=new $h.CommonParameters(null,callback,userContext);
	return true
};
$h.CommonParameters.prototype={
	_options$p$0: null,
	_callback$p$0: null,
	_asyncContext$p$0: null,
	get_options: function()
	{
		return this._options$p$0
	},
	get_callback: function()
	{
		return this._callback$p$0
	},
	get_asyncContext: function()
	{
		return this._asyncContext$p$0
	}
};
$h.EwsRequest=function(userContext)
{
	$h.EwsRequest.initializeBase(this,[userContext])
};
$h.EwsRequest.prototype={
	readyState: 1,
	status: 0,
	statusText: null,
	onreadystatechange: null,
	responseText: null,
	get__statusCode$i$1: function()
	{
		return this.status
	},
	set__statusCode$i$1: function(value)
	{
		this.status=value;
		return value
	},
	get__statusDescription$i$1: function()
	{
		return this.statusText
	},
	set__statusDescription$i$1: function(value)
	{
		this.statusText=value;
		return value
	},
	get__requestState$i$1: function()
	{
		return this.readyState
	},
	set__requestState$i$1: function(value)
	{
		this.readyState=value;
		return value
	},
	get_hasOnReadyStateChangeCallback: function()
	{
		return!$h.ScriptHelpers.isNullOrUndefined(this.onreadystatechange)
	},
	get__response$i$1: function()
	{
		return this.responseText
	},
	set__response$i$1: function(value)
	{
		this.responseText=value;
		return value
	},
	send: function(data)
	{
		this._checkSendConditions$i$1();
		if($h.ScriptHelpers.isNullOrUndefined(data))
			this._throwInvalidStateException$i$1();
		this._sendRequest$i$0(5,"EwsRequest",{body: data})
	},
	_callOnReadyStateChangeCallback$i$1: function()
	{
		if(!$h.ScriptHelpers.isNullOrUndefined(this.onreadystatechange))
			this.onreadystatechange()
	},
	_parseExtraResponseData$i$1: function(response){},
	executeExtraFailedResponseSteps: function(){}
};
$h.InitialData=function(data)
{
	this._data$p$0=data;
	this._permissionLevel$p$0=this._calculatePermissionLevel$p$0()
};
$h.InitialData._defineReadOnlyProperty$i=function(o, methodName, getter)
{
	var propertyDescriptor={
			get: getter,
			configurable: false
		};
	Object.defineProperty(o,methodName,propertyDescriptor)
};
$h.InitialData.prototype={
	_toRecipients$p$0: null,
	_ccRecipients$p$0: null,
	_attachments$p$0: null,
	_resources$p$0: null,
	_entities$p$0: null,
	_data$p$0: null,
	_permissionLevel$p$0: 0,
	get__isRestIdSupported$i$0: function()
	{
		return this._data$p$0["isRestIdSupported"]
	},
	get__itemId$i$0: function()
	{
		return this._data$p$0["id"]
	},
	get__itemClass$i$0: function()
	{
		return this._data$p$0["itemClass"]
	},
	get__dateTimeCreated$i$0: function()
	{
		return new Date(this._data$p$0["dateTimeCreated"])
	},
	get__dateTimeModified$i$0: function()
	{
		return new Date(this._data$p$0["dateTimeModified"])
	},
	get__dateTimeSent$i$0: function()
	{
		return new Date(this._data$p$0["dateTimeSent"])
	},
	get__subject$i$0: function()
	{
		this._throwOnRestrictedPermissionLevel$i$0();
		return this._data$p$0["subject"]
	},
	get__normalizedSubject$i$0: function()
	{
		this._throwOnRestrictedPermissionLevel$i$0();
		return this._data$p$0["normalizedSubject"]
	},
	get__internetMessageId$i$0: function()
	{
		return this._data$p$0["internetMessageId"]
	},
	get__conversationId$i$0: function()
	{
		return this._data$p$0["conversationId"]
	},
	get__sender$i$0: function()
	{
		this._throwOnRestrictedPermissionLevel$i$0();
		var sender=this._data$p$0["sender"];
		return $h.ScriptHelpers.isNullOrUndefined(sender) ? null : new $h.EmailAddressDetails(sender)
	},
	get__from$i$0: function()
	{
		this._throwOnRestrictedPermissionLevel$i$0();
		var from=this._data$p$0["from"];
		return $h.ScriptHelpers.isNullOrUndefined(from) ? null : new $h.EmailAddressDetails(from)
	},
	get__to$i$0: function()
	{
		this._throwOnRestrictedPermissionLevel$i$0();
		if(null===this._toRecipients$p$0)
			this._toRecipients$p$0=this._createEmailAddressDetails$p$0("to");
		return this._toRecipients$p$0
	},
	get__cc$i$0: function()
	{
		this._throwOnRestrictedPermissionLevel$i$0();
		if(null===this._ccRecipients$p$0)
			this._ccRecipients$p$0=this._createEmailAddressDetails$p$0("cc");
		return this._ccRecipients$p$0
	},
	get__attachments$i$0: function()
	{
		this._throwOnRestrictedPermissionLevel$i$0();
		if(null===this._attachments$p$0)
			this._attachments$p$0=this._createAttachmentDetails$p$0();
		return this._attachments$p$0
	},
	get__ewsUrl$i$0: function()
	{
		return this._data$p$0["ewsUrl"]
	},
	get__marketplaceAssetId$i$0: function()
	{
		return this._data$p$0["marketplaceAssetId"]
	},
	get__extensionId$i$0: function()
	{
		return this._data$p$0["extensionId"]
	},
	get__marketplaceContentMarket$i$0: function()
	{
		return this._data$p$0["marketplaceContentMarket"]
	},
	get__consentMetadata$i$0: function()
	{
		return this._data$p$0["consentMetadata"]
	},
	get__isRead$i$0: function()
	{
		return this._data$p$0["isRead"]
	},
	get__endNodeUrl$i$0: function()
	{
		return this._data$p$0["endNodeUrl"]
	},
	get__entryPointUrl$i$0: function()
	{
		return this._data$p$0["entryPointUrl"]
	},
	get__start$i$0: function()
	{
		return new Date(this._data$p$0["start"])
	},
	get__end$i$0: function()
	{
		return new Date(this._data$p$0["end"])
	},
	get__location$i$0: function()
	{
		return this._data$p$0["location"]
	},
	get__userProfileType$i$0: function()
	{
		return this._data$p$0["userProfileType"]
	},
	get__resources$i$0: function()
	{
		this._throwOnRestrictedPermissionLevel$i$0();
		if(null===this._resources$p$0)
			this._resources$p$0=this._createEmailAddressDetails$p$0("resources");
		return this._resources$p$0
	},
	get__organizer$i$0: function()
	{
		this._throwOnRestrictedPermissionLevel$i$0();
		var organizer=this._data$p$0["organizer"];
		return $h.ScriptHelpers.isNullOrUndefined(organizer) ? null : new $h.EmailAddressDetails(organizer)
	},
	get__userDisplayName$i$0: function()
	{
		return this._data$p$0["userDisplayName"]
	},
	get__userEmailAddress$i$0: function()
	{
		return this._data$p$0["userEmailAddress"]
	},
	get__userTimeZone$i$0: function()
	{
		return this._data$p$0["userTimeZone"]
	},
	get__timeZoneOffsets$i$0: function()
	{
		return this._data$p$0["timeZoneOffsets"]
	},
	get__hostVersion$i$0: function()
	{
		return this._data$p$0["hostVersion"]
	},
	get__owaView$i$0: function()
	{
		return this._data$p$0["owaView"]
	},
	get__overrideWindowOpen$i$0: function()
	{
		return this._data$p$0["overrideWindowOpen"]
	},
	_getEntities$i$0: function()
	{
		if(!this._entities$p$0)
			this._entities$p$0=new $h.Entities(this._data$p$0["entities"],this._data$p$0["filteredEntities"],this.get__dateTimeSent$i$0(),this._permissionLevel$p$0);
		return this._entities$p$0
	},
	_getEntitiesByType$i$0: function(entityType)
	{
		var entites=this._getEntities$i$0();
		return entites._getByType$i$0(entityType)
	},
	_getFilteredEntitiesByName$i$0: function(name)
	{
		var entities=this._getEntities$i$0();
		return entities._getFilteredEntitiesByName$i$0(name)
	},
	_getRegExMatches$i$0: function()
	{
		if(!this._data$p$0["regExMatches"])
			return null;
		return this._data$p$0["regExMatches"]
	},
	_getRegExMatchesByName$i$0: function(regexName)
	{
		var regexMatches=this._getRegExMatches$i$0();
		if(!regexMatches || !regexMatches[regexName])
			return null;
		return regexMatches[regexName]
	},
	_throwOnRestrictedPermissionLevel$i$0: function()
	{
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnPropertyAccessForRestrictedPermission$i(this._permissionLevel$p$0)
	},
	_createEmailAddressDetails$p$0: function(key)
	{
		var to=this._data$p$0[key];
		if($h.ScriptHelpers.isNullOrUndefined(to))
			return[];
		var recipients=[];
		for(var i=0; i < to["length"]; i++)
			if(!$h.ScriptHelpers.isNullOrUndefined(to[i]))
				recipients[i]=new $h.EmailAddressDetails(to[i]);
		return recipients
	},
	_createAttachmentDetails$p$0: function()
	{
		var attachments=this._data$p$0["attachments"];
		if($h.ScriptHelpers.isNullOrUndefined(attachments))
			return[];
		var attachmentDetails=[];
		for(var i=0; i < attachments["length"]; i++)
			if(!$h.ScriptHelpers.isNullOrUndefined(attachments[i]))
				attachmentDetails[i]=new $h.AttachmentDetails(attachments[i]);
		return attachmentDetails
	},
	_calculatePermissionLevel$p$0: function()
	{
		var HostReadItem=1;
		var HostReadWriteMailbox=2;
		var HostReadWriteItem=3;
		var permissionLevelFromHost=this._data$p$0["permissionLevel"];
		if($h.ScriptHelpers.isUndefined(this._permissionLevel$p$0))
			return 0;
		switch(permissionLevelFromHost)
		{
			case HostReadItem:
				return 1;
			case HostReadWriteItem:
				return 2;
			case HostReadWriteMailbox:
				return 3;
			default:
				return 0
		}
	}
};
$h._loadDictionaryRequest=function(createResultObject, dictionaryName, callback, userContext)
{
	$h._loadDictionaryRequest.initializeBase(this,[userContext]);
	this._createResultObject$p$1=createResultObject;
	this._dictionaryName$p$1=dictionaryName;
	this._callback$p$1=callback
};
$h._loadDictionaryRequest.prototype={
	_dictionaryName$p$1: null,
	_createResultObject$p$1: null,
	_callback$p$1: null,
	handleResponse: function(response)
	{
		if(response["wasSuccessful"])
		{
			var value=response[this._dictionaryName$p$1];
			var responseData=window["JSON"]["parse"](value);
			this.createAsyncResult(this._createResultObject$p$1(responseData),0,0,null)
		}
		else
			this.createAsyncResult(null,1,9020,response["errorMessage"]);
		this._callback$p$1(this._asyncResult$p$0)
	}
};
$h.ProxyRequestBase=function(userContext)
{
	$h.ProxyRequestBase.initializeBase(this,[userContext])
};
$h.ProxyRequestBase.prototype={
	handleResponse: function(response)
	{
		if(!response["wasProxySuccessful"])
		{
			this.set__statusCode$i$1(500);
			this.set__statusDescription$i$1("Error");
			var errorMessage=response["errorMessage"];
			this.set__response$i$1(errorMessage);
			this.createAsyncResult(null,1,9020,errorMessage)
		}
		else
		{
			this.set__statusCode$i$1(response["statusCode"]);
			this.set__statusDescription$i$1(response["statusDescription"]);
			this.set__response$i$1(response["body"]);
			this.createAsyncResult(this.get__response$i$1(),0,0,null)
		}
		this._parseExtraResponseData$i$1(response);
		this._cycleReadyStateFromHeadersReceivedToLoadingToDone$i$1()
	},
	_throwInvalidStateException$i$1: function()
	{
		throw Error.create("DOMException",{
			code: 11,
			message: "INVALID_STATE_ERR"
		});
	},
	_cycleReadyStateFromHeadersReceivedToLoadingToDone$i$1: function()
	{
		var $$t_0=this;
		this._changeReadyState$i$1(2,function()
		{
			$$t_0._changeReadyState$i$1(3,function()
			{
				$$t_0._changeReadyState$i$1(4,null)
			})
		})
	},
	_changeReadyState$i$1: function(state, nextStep)
	{
		this.set__requestState$i$1(state);
		var $$t_2=this;
		window.setTimeout(function()
		{
			try
			{
				$$t_2._callOnReadyStateChangeCallback$i$1()
			}
			finally
			{
				if(!$h.ScriptHelpers.isNullOrUndefined(nextStep))
					nextStep()
			}
		},0)
	},
	_checkSendConditions$i$1: function()
	{
		if(this.get__requestState$i$1() !==1)
			this._throwInvalidStateException$i$1();
		if(this._isSent$p$0)
			this._throwInvalidStateException$i$1()
	}
};
$h.RequestBase=function(userContext)
{
	this._userContext$p$0=userContext
};
$h.RequestBase.prototype={
	_isSent$p$0: false,
	_asyncResult$p$0: null,
	_userContext$p$0: null,
	get_asyncResult: function()
	{
		return this._asyncResult$p$0
	},
	_sendRequest$i$0: function(dispid, methodName, dataToSend)
	{
		this._isSent$p$0=true;
		var $$t_5=this;
		window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.invokeHostMethod(dispid,dataToSend,function(resultCode, response)
		{
			if(resultCode)
				$$t_5.createAsyncResult(null,1,9017,String.format(window["_u"]["ExtensibilityStrings"]["l_InternalProtocolError_Text"],resultCode));
			else
				$$t_5.handleResponse(response)
		})
	},
	createAsyncResult: function(value, errorCode, detailedErrorCode, errorDescription)
	{
		this._asyncResult$p$0=window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.createAsyncResult(value,errorCode,detailedErrorCode,this._userContext$p$0,errorDescription)
	}
};
$h.SaveDictionaryRequest=function(callback, userContext)
{
	$h.SaveDictionaryRequest.initializeBase(this,[userContext]);
	if(!$h.ScriptHelpers.isNullOrUndefined(callback))
		this._callback$p$1=callback
};
$h.SaveDictionaryRequest.prototype={
	_callback$p$1: null,
	handleResponse: function(response)
	{
		if(response["wasSuccessful"])
			this.createAsyncResult(null,0,0,null);
		else
			this.createAsyncResult(null,1,9020,response["errorMessage"]);
		if(!$h.ScriptHelpers.isNullOrUndefined(this._callback$p$1))
			this._callback$p$1(this._asyncResult$p$0)
	}
};
$h.ScriptHelpers=function(){};
$h.ScriptHelpers.isNull=function(value)
{
	return null===value
};
$h.ScriptHelpers.isNullOrUndefined=function(value)
{
	return $h.ScriptHelpers.isNull(value) || $h.ScriptHelpers.isUndefined(value)
};
$h.ScriptHelpers.isUndefined=function(value)
{
	return value===undefined
};
$h.ScriptHelpers.dictionaryContainsKey=function(obj, keyName)
{
	return Object.isInstanceOfType(obj) ? keyName in obj : false
};
$h.ScriptHelpers.isNonEmptyString=function(value)
{
	if(!value)
		return false;
	return String.isInstanceOfType(value)
};
$h.ScriptHelpers.deepClone=function(obj)
{
	return window["JSON"]["parse"](window["JSON"]["stringify"](obj))
};
window["OSF"]["DDA"]["OutlookAppOm"].registerClass("OSF.DDA.OutlookAppOm");
window["OSF"]["DDA"]["Settings"].registerClass("OSF.DDA.Settings");
$h.ItemBase.registerClass("$h.ItemBase");
$h.Item.registerClass("$h.Item",$h.ItemBase);
$h.Appointment.registerClass("$h.Appointment",$h.Item);
$h.ComposeItem.registerClass("$h.ComposeItem",$h.ItemBase);
$h.AppointmentCompose.registerClass("$h.AppointmentCompose",$h.ComposeItem);
$h.AttachmentConstants.registerClass("$h.AttachmentConstants");
$h.AttachmentDetails.registerClass("$h.AttachmentDetails");
$h.Body.registerClass("$h.Body");
$h.ComposeBody.registerClass("$h.ComposeBody",$h.Body);
$h.ComposeRecipient.registerClass("$h.ComposeRecipient");
$h.ComposeLocation.registerClass("$h.ComposeLocation");
$h.ComposeSubject.registerClass("$h.ComposeSubject");
$h.ComposeTime.registerClass("$h.ComposeTime");
$h.Contact.registerClass("$h.Contact");
$h.CustomProperties.registerClass("$h.CustomProperties");
$h.Diagnostics.registerClass("$h.Diagnostics");
$h.EmailAddressDetails.registerClass("$h.EmailAddressDetails");
$h.Entities.registerClass("$h.Entities");
$h.ReplyConstants.registerClass("$h.ReplyConstants");
$h.AsyncConstants.registerClass("$h.AsyncConstants");
window["Office"]["cast"]["item"].registerClass("Office.cast.item");
$h.Message.registerClass("$h.Message",$h.Item);
$h.MeetingRequest.registerClass("$h.MeetingRequest",$h.Message);
$h.MeetingSuggestion.registerClass("$h.MeetingSuggestion");
$h.MeetingSuggestionTimeDecoder.registerClass("$h.MeetingSuggestionTimeDecoder");
$h._extractedDate.registerClass("$h._extractedDate");
$h._preciseDate.registerClass("$h._preciseDate",$h._extractedDate);
$h._relativeDate.registerClass("$h._relativeDate",$h._extractedDate);
$h.MessageCompose.registerClass("$h.MessageCompose",$h.ComposeItem);
$h.NotificationMessages.registerClass("$h.NotificationMessages");
$h.OutlookErrorManager.registerClass("$h.OutlookErrorManager");
$h.OutlookErrorManager.OutlookErrorCodes.registerClass("$h.OutlookErrorManager.OutlookErrorCodes");
$h.OutlookErrorManager.OsfDdaErrorCodes.registerClass("$h.OutlookErrorManager.OsfDdaErrorCodes");
$h.PhoneNumber.registerClass("$h.PhoneNumber");
$h.TaskSuggestion.registerClass("$h.TaskSuggestion");
$h.UserProfile.registerClass("$h.UserProfile");
$h.CommonParameters.registerClass("$h.CommonParameters");
$h.RequestBase.registerClass("$h.RequestBase");
$h.ProxyRequestBase.registerClass("$h.ProxyRequestBase",$h.RequestBase);
$h.EwsRequest.registerClass("$h.EwsRequest",$h.ProxyRequestBase);
$h.InitialData.registerClass("$h.InitialData");
$h._loadDictionaryRequest.registerClass("$h._loadDictionaryRequest",$h.RequestBase);
$h.SaveDictionaryRequest.registerClass("$h.SaveDictionaryRequest",$h.RequestBase);
$h.ScriptHelpers.registerClass("$h.ScriptHelpers");
window["OSF"]["DDA"]["OutlookAppOm"].asyncMethodTimeoutKeyName="__timeout__";
window["OSF"]["DDA"]["OutlookAppOm"].ewsIdOrEmailParamName="ewsIdOrEmail";
window["OSF"]["DDA"]["OutlookAppOm"]._maxRecipients$p=100;
window["OSF"]["DDA"]["OutlookAppOm"]._maxSubjectLength$p=255;
window["OSF"]["DDA"]["OutlookAppOm"].maxBodyLength=32768;
window["OSF"]["DDA"]["OutlookAppOm"]._maxLocationLength$p=255;
window["OSF"]["DDA"]["OutlookAppOm"]._maxEwsRequestSize$p=1e6;
window["OSF"]["DDA"]["OutlookAppOm"].executeMethodName="ExecuteMethod";
window["OSF"]["DDA"]["OutlookAppOm"].getInitialDataMethodName="GetInitialData";
window["OSF"]["DDA"]["OutlookAppOm"].itemIdParameterName="itemId";
window["OSF"]["DDA"]["OutlookAppOm"].restVersionParameterName="restVersion";
window["OSF"]["DDA"]["OutlookAppOm"]._instance$p=null;
$h.AttachmentConstants.maxAttachmentNameLength=255;
$h.AttachmentConstants.maxUrlLength=2048;
$h.AttachmentConstants.maxItemIdLength=200;
$h.AttachmentConstants.maxRemoveIdLength=200;
$h.AttachmentConstants.attachmentParameterName="attachments";
$h.AttachmentConstants.attachmentTypeParameterName="type";
$h.AttachmentConstants.attachmentUrlParameterName="url";
$h.AttachmentConstants.attachmentItemIdParameterName="itemId";
$h.AttachmentConstants.attachmentNameParameterName="name";
$h.AttachmentConstants.attachmentTypeFileName="file";
$h.AttachmentConstants.attachmentTypeItemName="item";
$h.AttachmentDetails._attachmentTypeMap$p=[window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["AttachmentType"]["File"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["AttachmentType"]["Item"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["AttachmentType"]["Cloud"]];
$h.Body.coercionTypeParameterName="coercionType";
$h.ComposeRecipient.displayNameLengthLimit=255;
$h.ComposeRecipient.maxSmtpLength=571;
$h.ComposeRecipient.recipientsLimit=100;
$h.ComposeRecipient.totalRecipientsLimit=500;
$h.ComposeRecipient.addressParameterName="address";
$h.ComposeRecipient.nameParameterName="name";
$h.ComposeLocation.locationKey="location";
$h.ComposeLocation.maximumLocationLength=255;
$h.ComposeSubject.maximumSubjectLength=255;
$h.ComposeTime.timeTypeName="TimeProperty";
$h.ComposeTime.timeDataName="time";
$h.Diagnostics.outlookAppName="Outlook";
$h.Diagnostics.outlookWebAppName="OutlookWebApp";
$h.EmailAddressDetails._emptyString$p="";
$h.EmailAddressDetails._responseTypeMap$p=[window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ResponseType"]["None"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ResponseType"]["Organizer"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ResponseType"]["Tentative"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ResponseType"]["Accepted"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ResponseType"]["Declined"]];
$h.EmailAddressDetails._recipientTypeMap$p=[window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RecipientType"]["Other"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RecipientType"]["DistributionList"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RecipientType"]["User"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RecipientType"]["ExternalUser"]];
$h.Entities._allEntityKeys$p=["Addresses","EmailAddresses","Urls","PhoneNumbers","TaskSuggestions","MeetingSuggestions","Contacts","FlightReservations","ParcelDeliveries"];
$h.ReplyConstants.htmlBodyKeyName="htmlBody";
$h.AsyncConstants.optionsKeyName="options";
$h.AsyncConstants.callbackKeyName="callback";
$h.AsyncConstants.asyncResultKeyName="asyncResult";
$h.MeetingSuggestionTimeDecoder._baseDate$p=new Date("0001-01-01T00:00:00Z");
$h.NotificationMessages.maximumKeyLength=32;
$h.NotificationMessages.maximumIconLength=32;
$h.NotificationMessages.maximumMessageLength=150;
$h.NotificationMessages.notificationsKeyParameterName="key";
$h.NotificationMessages.notificationsTypeParameterName="type";
$h.NotificationMessages.notificationsIconParameterName="icon";
$h.NotificationMessages.notificationsMessageParameterName="message";
$h.NotificationMessages.notificationsPersistentParameterName="persistent";
$h.OutlookErrorManager.errorNameKey="name";
$h.OutlookErrorManager.errorMessageKey="message";
$h.OutlookErrorManager._isInitialized$p=false;
$h.OutlookErrorManager.OutlookErrorCodes.attachmentSizeExceeded=9e3;
$h.OutlookErrorManager.OutlookErrorCodes.numberOfAttachmentsExceeded=9001;
$h.OutlookErrorManager.OutlookErrorCodes.internalFormatError=9002;
$h.OutlookErrorManager.OutlookErrorCodes.invalidAttachmentId=9003;
$h.OutlookErrorManager.OutlookErrorCodes.invalidAttachmentPath=9004;
$h.OutlookErrorManager.OutlookErrorCodes.cannotAddAttachmentBeforeUpgrade=9005;
$h.OutlookErrorManager.OutlookErrorCodes.attachmentDeletedBeforeUploadCompletes=9006;
$h.OutlookErrorManager.OutlookErrorCodes.attachmentUploadGeneralFailure=9007;
$h.OutlookErrorManager.OutlookErrorCodes.attachmentToDeleteDoesNotExist=9008;
$h.OutlookErrorManager.OutlookErrorCodes.attachmentDeleteGeneralFailure=9009;
$h.OutlookErrorManager.OutlookErrorCodes.invalidEndTime=9010;
$h.OutlookErrorManager.OutlookErrorCodes.htmlSanitizationFailure=9011;
$h.OutlookErrorManager.OutlookErrorCodes.numberOfRecipientsExceeded=9012;
$h.OutlookErrorManager.OutlookErrorCodes.noValidRecipientsProvided=9013;
$h.OutlookErrorManager.OutlookErrorCodes.cursorPositionChanged=9014;
$h.OutlookErrorManager.OutlookErrorCodes.invalidSelection=9016;
$h.OutlookErrorManager.OutlookErrorCodes.accessRestricted=9017;
$h.OutlookErrorManager.OutlookErrorCodes.genericTokenError=9018;
$h.OutlookErrorManager.OutlookErrorCodes.genericSettingsError=9019;
$h.OutlookErrorManager.OutlookErrorCodes.genericResponseError=9020;
$h.OutlookErrorManager.OutlookErrorCodes.saveError=9021;
$h.OutlookErrorManager.OutlookErrorCodes.messageInDifferentStoreError=9022;
$h.OutlookErrorManager.OutlookErrorCodes.duplicateNotificationKey=9023;
$h.OutlookErrorManager.OutlookErrorCodes.notificationKeyNotFound=9024;
$h.OutlookErrorManager.OutlookErrorCodes.numberOfNotificationsExceeded=9025;
$h.OutlookErrorManager.OutlookErrorCodes.persistedNotificationArrayReadError=9026;
$h.OutlookErrorManager.OutlookErrorCodes.persistedNotificationArraySaveError=9027;
$h.OutlookErrorManager.OutlookErrorCodes.cannotPersistPropertyInUnsavedDraftError=9028;
$h.OutlookErrorManager.OutlookErrorCodes.ooeInvalidDataFormat=2006;
$h.OutlookErrorManager.OsfDdaErrorCodes.ooeCoercionTypeNotSupported=1e3;
$h.CommonParameters.asyncContextKeyName="asyncContext";
$h.InitialData.userProfileTypeKey="userProfileType";
$h.ScriptHelpers.emptyString="";
OSF.ClientMode={
	ReadWrite: 0,
	ReadOnly: 1
}
OSF.DDA.RichInitializationReason={
	1: Microsoft.Office.WebExtension.InitializationReason.Inserted,
	2: Microsoft.Office.WebExtension.InitializationReason.DocumentOpened
};
Microsoft.Office.WebExtension.FileType={
	Text: "text",
	Compressed: "compressed"
};
OSF.DDA.RichClientSettingsManager={
	read: function OSF_DDA_RichClientSettingsManager$Read(onCalling, onReceiving) {
		var keys=[];
		var values=[];
		if (onCalling) {
			onCalling();
		}
		if (typeof OsfOMToken !='undefined' && OsfOMToken) {
			window.external.GetContext().GetSettings(OsfOMToken).Read(keys, values);
		}
		else
		{
			window.external.GetContext().GetSettings().Read(keys, values);
		}
		if (onReceiving) {
			onReceiving();
		}
		var serializedSettings={};
		for (var index=0; index < keys.length; index++) {
			serializedSettings[keys[index]]=values[index];
		}
		return serializedSettings;
	},
	write: function OSF_DDA_RichClientSettingsManager$Write(serializedSettings, overwriteIfStale, onCalling, onReceiving) {
		var keys=[];
		var values=[];
		for (var key in serializedSettings) {
			keys.push(key);
			values.push(serializedSettings[key]);
		}
		if (onCalling) {
			onCalling();
		}
		if (typeof OsfOMToken !='undefined' && OsfOMToken) {
			window.external.GetContext().GetSettings(OsfOMToken).Write(keys, values);
		}
		else {
			window.external.GetContext().GetSettings().Write(keys, values);
		}
		if (onReceiving) {
			onReceiving();
		}
	}
};
OSF.DDA.DispIdHost.getRichClientDelegateMethods=function (actionId) {
	var delegateMethods={};
	delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=OSF.DDA.SafeArray.Delegate.executeAsync;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync]=OSF.DDA.SafeArray.Delegate.registerEventAsync;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync]=OSF.DDA.SafeArray.Delegate.unregisterEventAsync;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.MessageParent]=OSF.DDA.SafeArray.Delegate.MessageParent;
	function getSettingsExecuteMethod(hostDelegateMethod) {
		return function (args) {
			var status, response;
			try {
				response=hostDelegateMethod(args.hostCallArgs, args.onCalling, args.onReceiving);
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess;
			} catch (ex) {
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
				response={ name : Strings.OfficeOM.L_InternalError, message : ex };
			}
			if (args.onComplete) {
				args.onComplete(status, response);
			}
		};
	}
	function readSerializedSettings(hostCallArgs, onCalling, onReceiving) {
		return OSF.DDA.RichClientSettingsManager.read(onCalling, onReceiving);
	}
	function writeSerializedSettings(hostCallArgs, onCalling, onReceiving) {
		return OSF.DDA.RichClientSettingsManager.write(
			hostCallArgs[OSF.DDA.SettingsManager.SerializedSettings],
			hostCallArgs[Microsoft.Office.WebExtension.Parameters.OverwriteIfStale],
			onCalling,
			onReceiving
		);
	}
	switch (actionId) {
		case OSF.DDA.AsyncMethodNames.RefreshAsync.id:
			delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=getSettingsExecuteMethod(readSerializedSettings);
			break;
		case OSF.DDA.AsyncMethodNames.SaveAsync.id:
			delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=getSettingsExecuteMethod(writeSerializedSettings);
			break;
		default:
			break;
	}
	return delegateMethods;
}
OSF.DDA.DispIdHost.getClientDelegateMethods=function (actionId) {
	var delegateMethods={};
	delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=OSF.DDA.SafeArray.Delegate.executeAsync;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync]=OSF.DDA.SafeArray.Delegate.registerEventAsync;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync]=OSF.DDA.SafeArray.Delegate.unregisterEventAsync;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.MessageParent]=OSF.DDA.SafeArray.Delegate.MessageParent;
	if (OSF.DDA.AsyncMethodNames.RefreshAsync && actionId==OSF.DDA.AsyncMethodNames.RefreshAsync.id) {
		var readSerializedSettings=function (hostCallArgs, onCalling, onReceiving) {
			return OSF.DDA.ClientSettingsManager.read(onCalling, onReceiving);
		};
		delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=OSF.DDA.ClientSettingsManager.getSettingsExecuteMethod(readSerializedSettings);
	}
	if (OSF.DDA.AsyncMethodNames.SaveAsync && actionId==OSF.DDA.AsyncMethodNames.SaveAsync.id) {
		var writeSerializedSettings=function (hostCallArgs, onCalling, onReceiving) {
			return OSF.DDA.ClientSettingsManager.write(hostCallArgs[OSF.DDA.SettingsManager.SerializedSettings], hostCallArgs[Microsoft.Office.WebExtension.Parameters.OverwriteIfStale], onCalling, onReceiving);
		};
		delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=OSF.DDA.ClientSettingsManager.getSettingsExecuteMethod(writeSerializedSettings);
	}
	return delegateMethods;
}
OSF.DDA.File=function OSF_DDA_File(handle, fileSize, sliceSize) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"size": {
			value: fileSize
		},
		"sliceCount": {
			value: Math.ceil(fileSize / sliceSize)
		}
	});
	var privateState={};
	privateState[OSF.DDA.FileProperties.Handle]=handle;
	privateState[OSF.DDA.FileProperties.SliceSize]=sliceSize;
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(
		this, [
			am.GetDocumentCopyChunkAsync,
			am.ReleaseDocumentCopyAsync
		],
		privateState
	);
}
OSF.DDA.FileSliceOffset="fileSliceoffset";
OSF.DDA.CustomXmlParts=function OSF_DDA_CustomXmlParts() {
	this._eventDispatches=[];
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.AddDataPartAsync,
		am.GetDataPartByIdAsync,
		am.GetDataPartsByNameSpaceAsync
	]);
};
OSF.DDA.CustomXmlPart=function OSF_DDA_CustomXmlPart(customXmlParts, id, builtIn) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"builtIn": {
			value: builtIn
		},
		"id": {
			value: id
		},
		"namespaceManager": {
			value: new OSF.DDA.CustomXmlPrefixMappings(id)
		}
	});
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.DeleteDataPartAsync,
		am.GetPartNodesAsync,
		am.GetPartXmlAsync
	]);
	var customXmlPartEventDispatches=customXmlParts._eventDispatches;
	var dispatch=customXmlPartEventDispatches[id];
	if (!dispatch) {
		var et=Microsoft.Office.WebExtension.EventType;
		dispatch=new OSF.EventDispatch([
			et.DataNodeDeleted,
			et.DataNodeInserted,
			et.DataNodeReplaced
		]);
		customXmlPartEventDispatches[id]=dispatch;
	}
	OSF.DDA.DispIdHost.addEventSupport(this, dispatch);
};
OSF.DDA.CustomXmlPrefixMappings=function OSF_DDA_CustomXmlPrefixMappings(partId) {
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(
		this,
		[
			am.AddDataPartNamespaceAsync,
			am.GetDataPartNamespaceAsync,
			am.GetDataPartPrefixAsync
		],
		partId
	);
};
OSF.DDA.CustomXmlNode=function OSF_DDA_CustomXmlNode(handle, nodeType, ns, baseName) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"baseName": {
			value: baseName
		},
		"namespaceUri": {
			value: ns
		},
		"nodeType": {
			value: nodeType
		}
	});
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(
		this,
		[
			am.GetRelativeNodesAsync,
			am.GetNodeValueAsync,
			am.GetNodeXmlAsync,
			am.SetNodeValueAsync,
			am.SetNodeXmlAsync,
			am.GetNodeTextAsync,
			am.SetNodeTextAsync
		],
		handle
	);
};
OSF.DDA.NodeInsertedEventArgs=function OSF_DDA_NodeInsertedEventArgs(newNode, inUndoRedo) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.DataNodeInserted
		},
		"newNode": {
			value: newNode
		},
		"inUndoRedo": {
			value: inUndoRedo
		}
	});
};
OSF.DDA.NodeReplacedEventArgs=function OSF_DDA_NodeReplacedEventArgs(oldNode, newNode, inUndoRedo) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.DataNodeReplaced
		},
		"oldNode": {
			value: oldNode
		},
		"newNode": {
			value: newNode
		},
		"inUndoRedo": {
			value: inUndoRedo
		}
	});
};
OSF.DDA.NodeDeletedEventArgs=function OSF_DDA_NodeDeletedEventArgs(oldNode, oldNextSibling, inUndoRedo) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.DataNodeDeleted
		},
		"oldNode": {
			value: oldNode
		},
		"oldNextSibling": {
			value: oldNextSibling
		},
		"inUndoRedo": {
			value: inUndoRedo
		}
	});
};
var OSFRichclient;
(function(OSFRichclient)
{
	var RichClientHostController=function()
		{
			function RichClientHostController(){}
			RichClientHostController.prototype.execute=function(id, params, callback)
			{
				if (typeof OsfOMToken !='undefined' && OsfOMToken) {
					window.external.Execute(id,params,callback,OsfOMToken);
				}
				else{
					window.external.Execute(id,params,callback);
				}
			};
			RichClientHostController.prototype.registerEvent=function(id, targetId, handler, callback)
			{
				if (typeof OsfOMToken !='undefined' && OsfOMToken) {
					window.external.RegisterEvent(id,targetId,handler,callback,OsfOMToken);
				}
				else {
					window.external.RegisterEvent(id,targetId,handler,callback);
				}
			};
			RichClientHostController.prototype.unregisterEvent=function(id, targetId, callback)
			{
				if (typeof OsfOMToken !='undefined' && OsfOMToken) {
					window.external.UnregisterEvent(id,targetId,callback,OsfOMToken);
				}
				else {
					window.external.UnregisterEvent(id,targetId,callback);
				}
			};
			return RichClientHostController
		}();
	OSFRichclient.RichClientHostController=RichClientHostController
})(OSFRichclient || (OSFRichclient={}));
OSF.ClientHostController=new OSFRichclient.RichClientHostController;
var OfficeExt;
(function(OfficeExt)
{
	(function(AppCommand)
	{
		var AppCommandManager=function()
			{
				function AppCommandManager()
				{
					var _this=this;
					this._pseudoDocument=null;
					this._eventDispatch=null;
					this._processAppCommandInvocation=function(args)
					{
						var verifyResult=_this._verifyManifestCallback(args.callbackName);
						if(verifyResult.errorCode !=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
						{
							_this._invokeAppCommandCompletedMethod(args.appCommandId,verifyResult.errorCode,"");
							return
						}
						var eventObj=_this._constructEventObjectForCallback(args);
						if(eventObj)
							window.setTimeout(function()
							{
								verifyResult.callback(eventObj)
							},0);
						else
							_this._invokeAppCommandCompletedMethod(args.appCommandId,OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError,"")
					}
				}
				AppCommandManager.initializeOsfDda=function()
				{
					OSF.DDA.AsyncMethodNames.addNames({AppCommandInvocationCompletedAsync: "appCommandInvocationCompletedAsync"});
					OSF.DDA.AsyncMethodCalls.define({
						method: OSF.DDA.AsyncMethodNames.AppCommandInvocationCompletedAsync,
						requiredArguments: [{
								name: Microsoft.Office.WebExtension.Parameters.Id,
								types: ["string"]
							},{
								name: Microsoft.Office.WebExtension.Parameters.Status,
								types: ["number"]
							},{
								name: Microsoft.Office.WebExtension.Parameters.Data,
								types: ["string"]
							}]
					});
					OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{AppCommandInvoked: "appCommandInvoked"});
					OSF.OUtil.setNamespace("AppCommand",OSF.DDA);
					OSF.DDA.AppCommand.AppCommandInvokedEventArgs=OfficeExt.AppCommand.AppCommandInvokedEventArgs
				};
				AppCommandManager.prototype.initializeAndChangeOnce=function(callback)
				{
					AppCommand.registerDdaFacade();
					this._pseudoDocument={};
					OSF.DDA.DispIdHost.addAsyncMethods(this._pseudoDocument,[OSF.DDA.AsyncMethodNames.AppCommandInvocationCompletedAsync]);
					this._eventDispatch=new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.AppCommandInvoked]);
					var onRegisterCompleted=function(result)
						{
							if(callback)
								if(result.status=="succeeded")
									callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
								else
									callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError)
						};
					OSF.DDA.DispIdHost.addEventSupport(this._pseudoDocument,this._eventDispatch);
					this._pseudoDocument.addHandlerAsync(Microsoft.Office.WebExtension.EventType.AppCommandInvoked,this._processAppCommandInvocation,onRegisterCompleted)
				};
				AppCommandManager.prototype._verifyManifestCallback=function(callbackName)
				{
					var defaultResult={
							callback: null,
							errorCode: OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidCallback
						};
					callbackName=callbackName.trim();
					try
					{
						var callList=callbackName.split(".");
						var parentObject=window;
						for(var i=0; i < callList.length - 1; i++)
							if(parentObject[callList[i]] && typeof parentObject[callList[i]]=="object")
								parentObject=parentObject[callList[i]];
							else
								return defaultResult;
						var callbackFunc=parentObject[callList[callList.length - 1]];
						if(typeof callbackFunc !="function")
							return defaultResult
					}
					catch(e)
					{
						return defaultResult
					}
					return{
							callback: callbackFunc,
							errorCode: OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess
						}
				};
				AppCommandManager.prototype._invokeAppCommandCompletedMethod=function(appCommandId, resultCode, data)
				{
					this._pseudoDocument.appCommandInvocationCompletedAsync(appCommandId,resultCode,data)
				};
				AppCommandManager.prototype._constructEventObjectForCallback=function(args)
				{
					var _this=this;
					var eventObj=new AppCommandCallbackEventArgs;
					try
					{
						var jsonData=JSON.parse(args.eventObjStr);
						this._translateEventObjectInternal(jsonData,eventObj);
						Object.defineProperty(eventObj,"completed",{
							value: function()
							{
								var jsonString=JSON.stringify(eventObj);
								_this._invokeAppCommandCompletedMethod(args.appCommandId,OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess,jsonString)
							},
							enumerable: true
						})
					}
					catch(e)
					{
						eventObj=null
					}
					return eventObj
				};
				AppCommandManager.prototype._translateEventObjectInternal=function(input, output)
				{
					for(var key in input)
					{
						if(!input.hasOwnProperty(key))
							continue;
						var inputChild=input[key];
						if(typeof inputChild=="object" && inputChild !=null)
						{
							OSF.OUtil.defineEnumerableProperty(output,key,{value: {}});
							this._translateEventObjectInternal(inputChild,output[key])
						}
						else
							Object.defineProperty(output,key,{
								value: inputChild,
								enumerable: true,
								writable: true
							})
					}
				};
				AppCommandManager.prototype._constructObjectByTemplate=function(template, input)
				{
					var output={};
					if(!template || !input)
						return output;
					for(var key in template)
						if(template.hasOwnProperty(key))
						{
							output[key]=null;
							if(input[key] !=null)
							{
								var templateChild=template[key];
								var inputChild=input[key];
								var inputChildType=typeof inputChild;
								if(typeof templateChild=="object" && templateChild !=null)
									output[key]=this._constructObjectByTemplate(templateChild,inputChild);
								else if(inputChildType=="number" || inputChildType=="string" || inputChildType=="boolean")
									output[key]=inputChild
							}
						}
					return output
				};
				AppCommandManager.instance=function()
				{
					if(AppCommandManager._instance==null)
						AppCommandManager._instance=new AppCommandManager;
					return AppCommandManager._instance
				};
				AppCommandManager._instance=null;
				return AppCommandManager
			}();
		AppCommand.AppCommandManager=AppCommandManager;
		var AppCommandInvokedEventArgs=function()
			{
				function AppCommandInvokedEventArgs(appCommandId, callbackName, eventObjStr)
				{
					this.type=Microsoft.Office.WebExtension.EventType.AppCommandInvoked;
					this.appCommandId=appCommandId;
					this.callbackName=callbackName;
					this.eventObjStr=eventObjStr
				}
				AppCommandInvokedEventArgs.create=function(eventProperties)
				{
					return new AppCommandInvokedEventArgs(eventProperties[AppCommand.AppCommandInvokedEventEnums.AppCommandId],eventProperties[AppCommand.AppCommandInvokedEventEnums.CallbackName],eventProperties[AppCommand.AppCommandInvokedEventEnums.EventObjStr])
				};
				return AppCommandInvokedEventArgs
			}();
		AppCommand.AppCommandInvokedEventArgs=AppCommandInvokedEventArgs;
		var AppCommandCallbackEventArgs=function()
			{
				function AppCommandCallbackEventArgs(){}
				return AppCommandCallbackEventArgs
			}();
		AppCommand.AppCommandCallbackEventArgs=AppCommandCallbackEventArgs;
		AppCommand.AppCommandInvokedEventEnums={
			AppCommandId: "appCommandId",
			CallbackName: "callbackName",
			EventObjStr: "eventObjStr"
		}
	})(OfficeExt.AppCommand || (OfficeExt.AppCommand={}));
	var AppCommand=OfficeExt.AppCommand
})(OfficeExt || (OfficeExt={}));
OfficeExt.AppCommand.AppCommandManager.initializeOsfDda();
var OfficeExt;
(function(OfficeExt)
{
	(function(AppCommand)
	{
		function registerDdaFacade()
		{
			if(OSF.DDA.SafeArray)
			{
				var parameterMap=OSF.DDA.SafeArray.Delegate.ParameterMap;
				parameterMap.define({
					type: OSF.DDA.MethodDispId.dispidAppCommandInvocationCompletedMethod,
					toHost: [{
							name: Microsoft.Office.WebExtension.Parameters.Id,
							value: 0
						},{
							name: Microsoft.Office.WebExtension.Parameters.Status,
							value: 1
						},{
							name: Microsoft.Office.WebExtension.Parameters.Data,
							value: 2
						}]
				});
				parameterMap.define({
					type: OSF.DDA.EventDispId.dispidAppCommandInvokedEvent,
					fromHost: [{
							name: OSF.DDA.EventDescriptors.AppCommandInvokedEvent,
							value: parameterMap.self
						}],
					isComplexType: true
				});
				parameterMap.define({
					type: OSF.DDA.EventDescriptors.AppCommandInvokedEvent,
					fromHost: [{
							name: OfficeExt.AppCommand.AppCommandInvokedEventEnums.AppCommandId,
							value: 0
						},{
							name: OfficeExt.AppCommand.AppCommandInvokedEventEnums.CallbackName,
							value: 1
						},{
							name: OfficeExt.AppCommand.AppCommandInvokedEventEnums.EventObjStr,
							value: 2
						}],
					isComplexType: true
				})
			}
		}
		AppCommand.registerDdaFacade=registerDdaFacade
	})(OfficeExt.AppCommand || (OfficeExt.AppCommand={}));
	var AppCommand=OfficeExt.AppCommand
})(OfficeExt || (OfficeExt={}));
var OfficeJsClient_OutlookWin32;
(function(OfficeJsClient_OutlookWin32)
{
	function prepareRightAfterWebExtensionInitialize()
	{
		var appCommandHandler=OfficeExt.AppCommand.AppCommandManager.instance();
		appCommandHandler.initializeAndChangeOnce();
	}
	function prepareRightBeforeWebExtensionInitialize()
	{
		OSF._OfficeAppFactory.setHostFacade(new OSF.DDA.DispIdHost.Facade(OSF.DDA.DispIdHost.getClientDelegateMethods,OSF.DDA.SafeArray.Delegate.ParameterMap));
	}
	OfficeJsClient_OutlookWin32.prepareRightBeforeWebExtensionInitialize=prepareRightBeforeWebExtensionInitialize;
	OfficeJsClient_OutlookWin32.prepareRightAfterWebExtensionInitialize=prepareRightAfterWebExtensionInitialize;
})(OfficeJsClient_OutlookWin32 || (OfficeJsClient_OutlookWin32={}));
OSF.OUtil.setNamespace("SafeArray", OSF.DDA);
OSF.DDA.SafeArray.Response={
	Status: 0,
	Payload: 1
};
OSF.DDA.SafeArray.UniqueArguments={
	Offset: "offset",
	Run: "run",
	BindingSpecificData: "bindingSpecificData",
	MergedCellGuid: "{66e7831f-81b2-42e2-823c-89e872d541b3}"
};
OSF.OUtil.setNamespace("Delegate", OSF.DDA.SafeArray);
OSF.DDA.SafeArray.Delegate.SpecialProcessor=function OSF_DDA_SafeArray_Delegate_SpecialProcessor() {
	function _2DVBArrayToJaggedArray(vbArr) {
		var ret;
		try {
			var rows=vbArr.ubound(1);
			var cols=vbArr.ubound(2);
			vbArr=vbArr.toArray();
			if (rows==1 && cols==1) {
				ret=[vbArr];
			} else {
				ret=[];
				for (var row=0; row < rows; row++) {
					var rowArr=[];
					for (var col=0; col < cols; col++) {
						var datum=vbArr[row * cols+col];
						if (datum !=OSF.DDA.SafeArray.UniqueArguments.MergedCellGuid) {
							rowArr.push(datum);
						}
					}
					if (rowArr.length > 0) {
						ret.push(rowArr);
					}
				}
			}
		} catch (ex) {
		}
		return ret;
	}
	var complexTypes=[
		OSF.DDA.PropertyDescriptors.FileProperties,
		OSF.DDA.PropertyDescriptors.FileSliceProperties,
		OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor,
		OSF.DDA.PropertyDescriptors.BindingProperties,
		OSF.DDA.SafeArray.UniqueArguments.BindingSpecificData,
		OSF.DDA.SafeArray.UniqueArguments.Offset,
		OSF.DDA.SafeArray.UniqueArguments.Run,
		OSF.DDA.PropertyDescriptors.Subset,
		OSF.DDA.PropertyDescriptors.DataPartProperties,
		OSF.DDA.PropertyDescriptors.DataNodeProperties,
		OSF.DDA.EventDescriptors.BindingSelectionChangedEvent,
		OSF.DDA.EventDescriptors.DataNodeInsertedEvent,
		OSF.DDA.EventDescriptors.DataNodeReplacedEvent,
		OSF.DDA.EventDescriptors.DataNodeDeletedEvent,
		OSF.DDA.EventDescriptors.DocumentThemeChangedEvent,
		OSF.DDA.EventDescriptors.OfficeThemeChangedEvent,
		OSF.DDA.EventDescriptors.ActiveViewChangedEvent,
		OSF.DDA.EventDescriptors.AppCommandInvokedEvent,
		OSF.DDA.DataNodeEventProperties.OldNode,
		OSF.DDA.DataNodeEventProperties.NewNode,
		OSF.DDA.DataNodeEventProperties.NextSiblingNode,
		Microsoft.Office.Internal.Parameters.OfficeTheme,
		Microsoft.Office.Internal.Parameters.DocumentTheme
	];
	var dynamicTypes={};
	dynamicTypes[Microsoft.Office.WebExtension.Parameters.Data]=(function () {
		var tableRows=0;
		var tableHeaders=1;
		return {
			toHost: function OSF_DDA_SafeArray_Delegate_SpecialProcessor_Data$toHost(data) {
				if (typeof data !="string" && data[OSF.DDA.TableDataProperties.TableRows] !==undefined) {
					var tableData=[];
					tableData[tableRows]=data[OSF.DDA.TableDataProperties.TableRows];
					tableData[tableHeaders]=data[OSF.DDA.TableDataProperties.TableHeaders];
					data=tableData;
				}
				return data;
			},
			fromHost: function OSF_DDA_SafeArray_Delegate_SpecialProcessor_Data$fromHost(hostArgs) {
				var ret;
				if (hostArgs.toArray) {
					var dimensions=hostArgs.dimensions();
					if(dimensions===2) {
						ret=_2DVBArrayToJaggedArray(hostArgs);
					} else {
						var array=hostArgs.toArray();
						if(array.length===2 &&  ((array[0] !=null && array[0].toArray) || (array[1] !=null && array[1].toArray))) {
							ret={};
							ret[OSF.DDA.TableDataProperties.TableRows]=_2DVBArrayToJaggedArray(array[tableRows]);
							ret[OSF.DDA.TableDataProperties.TableHeaders]=_2DVBArrayToJaggedArray(array[tableHeaders]);
						} else {
							ret=array;
						}
					}
				} else {
					ret=hostArgs;
				}
				return ret;
			}
		}
	})();
	OSF.DDA.SafeArray.Delegate.SpecialProcessor.uber.constructor.call(this, complexTypes, dynamicTypes);
	this.pack=function OSF_DDA_SafeArray_Delegate_SpecialProcessor$pack(param, arg) {
		var value;
		if (this.isDynamicType(param)) {
			value=dynamicTypes[param].toHost(arg);
		} else {
			value=arg;
		}
		return value;
	};
	this.unpack=function OSF_DDA_SafeArray_Delegate_SpecialProcessor$unpack(param, arg) {
		var value;
		if (this.isComplexType(param) || OSF.DDA.ListType.isListType(param)) {
			try {
				value=arg.toArray();
			} catch (ex) {
				value=arg || {};
			}
		} else if (this.isDynamicType(param)) {
			value=dynamicTypes[param].fromHost(arg);
		} else {
			value=arg;
		}
		return value;
	};
	this.dynamicTypes=dynamicTypes;
}
OSF.OUtil.extend(OSF.DDA.SafeArray.Delegate.SpecialProcessor, OSF.DDA.SpecialProcessor);
OSF.DDA.SafeArray.Delegate.ParameterMap=(function () {
	var parameterMap=new OSF.DDA.HostParameterMap(new OSF.DDA.SafeArray.Delegate.SpecialProcessor());
	var ns;
	var self=parameterMap.self;
	function createObject(properties) {
		var obj=null;
		if (properties) {
			obj={};
			var len=properties.length;
			for (var i=0; i < len; i++) {
				obj[properties[i].name]=properties[i].value;
			}
		}
		return obj;
	}
	function define(definition) {
		var args={};
		var toHost=createObject(definition.toHost);
		if (definition.invertible) {
			args.map=toHost;
		}
		else if (definition.canonical) {
			args.toHost=args.fromHost=toHost;
		}
		else {
			args.toHost=toHost;
			args.fromHost=createObject(definition.fromHost);
		}
		parameterMap.setMapping(definition.type, args);
	}
	ns=OSF.DDA.FileProperties;
	define({
		type: OSF.DDA.PropertyDescriptors.FileProperties,
		fromHost: [
			{ name: ns.Handle, value: 0 },
			{ name: ns.FileSize, value: 1 }
		]
	});
	define({
		type: OSF.DDA.PropertyDescriptors.FileSliceProperties,
		fromHost: [
			{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 0 },
			{ name: ns.SliceSize, value: 1}
		]
	});
	ns=OSF.DDA.FilePropertiesDescriptor;
	define({
		type: OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor,
		fromHost: [
			{ name: ns.Url, value: 0 }
		]
	});
	ns=OSF.DDA.BindingProperties;
	define({
		type: OSF.DDA.PropertyDescriptors.BindingProperties,
		fromHost: [
			{ name: ns.Id, value: 0 },
			{ name: ns.Type, value: 1 },
			{ name: OSF.DDA.SafeArray.UniqueArguments.BindingSpecificData, value: 2 }
		]
	});
	define({
		type: OSF.DDA.SafeArray.UniqueArguments.BindingSpecificData,
		fromHost: [
			{ name: ns.RowCount, value: 0 },
			{ name: ns.ColumnCount, value: 1 },
			{ name: ns.HasHeaders, value: 2 }
		]
	});
	ns=OSF.DDA.SafeArray.UniqueArguments;
	define({
		type: OSF.DDA.PropertyDescriptors.Subset,
		toHost: [
			{ name: ns.Offset, value: 0 },
			{ name: ns.Run, value: 1 }
		],
		canonical: true
	});
	ns=Microsoft.Office.WebExtension.Parameters;
	define({
		type: OSF.DDA.SafeArray.UniqueArguments.Offset,
		toHost: [
			{ name: ns.StartRow, value: 0 },
			{ name: ns.StartColumn, value: 1 }
		],
		canonical: true
	});
	define({
		type: OSF.DDA.SafeArray.UniqueArguments.Run,
		toHost: [
			{ name: ns.RowCount, value: 0 },
			{ name: ns.ColumnCount, value: 1 }
		],
		canonical: true
	});
	ns=OSF.DDA.DataPartProperties;
	define({
		type: OSF.DDA.PropertyDescriptors.DataPartProperties,
		fromHost: [
			{ name: ns.Id, value: 0 },
			{ name: ns.BuiltIn, value: 1 }
		]
	});
	ns=OSF.DDA.DataNodeProperties;
	define({
		type: OSF.DDA.PropertyDescriptors.DataNodeProperties,
		fromHost: [
			{ name: ns.Handle, value: 0 },
			{ name: ns.BaseName, value: 1 },
			{ name: ns.NamespaceUri, value: 2 },
			{ name: ns.NodeType, value: 3 }
		]
	});
	define({
		type: OSF.DDA.EventDescriptors.BindingSelectionChangedEvent,
		fromHost: [
			{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: 0 },
			{ name: OSF.DDA.PropertyDescriptors.Subset, value: 1 }
		]
	});
	define({
		type: OSF.DDA.EventDescriptors.DocumentThemeChangedEvent,
		fromHost: [
			{ name: Microsoft.Office.Internal.Parameters.DocumentTheme, value: self}
		]
	})
	define({
		type: OSF.DDA.EventDescriptors.OfficeThemeChangedEvent,
		fromHost: [
			{ name: Microsoft.Office.Internal.Parameters.OfficeTheme, value: self}
		]
	})
	define({
		type: OSF.DDA.EventDescriptors.ActiveViewChangedEvent,
		fromHost: [
			{ name: Microsoft.Office.WebExtension.Parameters.ActiveView, value: 0}
		]
	})
	ns=OSF.DDA.DataNodeEventProperties;
	define({
		type: OSF.DDA.EventDescriptors.DataNodeInsertedEvent,
		fromHost: [
			{ name: ns.InUndoRedo, value: 0 },
			{ name: ns.NewNode, value: 1 }
		]
	});
	define({
		type: OSF.DDA.EventDescriptors.DataNodeReplacedEvent,
		fromHost: [
			{ name: ns.InUndoRedo, value: 0 },
			{ name: ns.OldNode, value: 1 },
			{ name: ns.NewNode, value: 2 }
		]
	});
	define({
		type: OSF.DDA.EventDescriptors.DataNodeDeletedEvent,
		fromHost: [
			{ name: ns.InUndoRedo, value: 0 },
			{ name: ns.OldNode, value: 1 },
			{ name: ns.NextSiblingNode, value: 2 }
		]
	});
	define({
		type: ns.OldNode,
		fromHost: [
			{ name: OSF.DDA.PropertyDescriptors.DataNodeProperties, value: self }
		]
	});
	define({
		type: ns.NewNode,
		fromHost: [
			{ name: OSF.DDA.PropertyDescriptors.DataNodeProperties, value: self }
		]
	});
	define({
		type: ns.NextSiblingNode,
		fromHost: [
			{ name: OSF.DDA.PropertyDescriptors.DataNodeProperties, value: self }
		]
	});
	ns=Microsoft.Office.WebExtension.AsyncResultStatus;
	define({
		type: OSF.DDA.PropertyDescriptors.AsyncResultStatus,
		fromHost: [
			{ name: ns.Succeeded, value: 0 },
			{ name: ns.Failed, value: 1 }
		]
	});
	ns=Microsoft.Office.WebExtension.CoercionType;
	define({
		type: Microsoft.Office.WebExtension.Parameters.CoercionType,
		toHost: [
			{ name: ns.Text, value: 0 },
			{ name: ns.Matrix, value: 1 },
			{ name: ns.Table, value: 2 },
			{ name: ns.Html, value: 3 },
			{ name: ns.Ooxml, value: 4 },
			{ name: ns.SlideRange, value:7 },
			{ name: ns.Image, value:8 }
		]
	});
	ns=Microsoft.Office.WebExtension.GoToType;
	define({
		type: Microsoft.Office.WebExtension.Parameters.GoToType,
		toHost: [
			{ name: ns.Binding, value: 0 },
			{ name: ns.NamedItem, value: 1 },
			{ name: ns.Slide, value: 2 },
			{ name: ns.Index, value: 3 }
		]
	});
	ns=Microsoft.Office.WebExtension.FileType;
	if (ns) {
		define({
			type: Microsoft.Office.WebExtension.Parameters.FileType,
			toHost: [
			{ name: ns.Text, value: 0 },
			{ name: ns.Compressed, value: 5 },
			{ name: ns.Pdf, value: 6 }
		]
		});
	}
	ns=Microsoft.Office.WebExtension.BindingType;
	if (ns) {
		define({
			type: Microsoft.Office.WebExtension.Parameters.BindingType,
			toHost: [
				{ name: ns.Text, value: 0 },
				{ name: ns.Matrix, value: 1 },
				{ name: ns.Table, value: 2 }
			],
			invertible: true
		});
	}
	ns=Microsoft.Office.WebExtension.ValueFormat;
	define({
		type: Microsoft.Office.WebExtension.Parameters.ValueFormat,
		toHost: [
			{ name: ns.Unformatted, value: 0 },
			{ name: ns.Formatted, value: 1 }
		]
	});
	ns=Microsoft.Office.WebExtension.FilterType;
	define({
		type: Microsoft.Office.WebExtension.Parameters.FilterType,
		toHost: [
			{ name: ns.All, value: 0 },
			{ name: ns.OnlyVisible, value: 1 }
		]
	});
	ns=Microsoft.Office.Internal.OfficeTheme;
	if (ns) {
		define({
			type:Microsoft.Office.Internal.Parameters.OfficeTheme,
			fromHost: [
						{name: ns.PrimaryFontColor, value: 0},
						{name: ns.PrimaryBackgroundColor, value: 1},
						{name: ns.SecondaryFontColor, value:2},
						{name: ns.SecondaryBackgroundColor, value:3}
			]
		})
	}
	ns=Microsoft.Office.WebExtension.ActiveView;
	if (ns) {
		define({
			type:Microsoft.Office.WebExtension.Parameters.ActiveView,
			fromHost: [
				{name: 0, value: ns.Read},
				{name: 1, value: ns.Edit}
			]
		})
	}
	ns=Microsoft.Office.Internal.DocumentTheme;
	if (ns) {
		define({
			type:Microsoft.Office.Internal.Parameters.DocumentTheme,
			fromHost: [
				{name: ns.PrimaryBackgroundColor, value: 0},
				{name: ns.PrimaryFontColor, value: 1},
				{name: ns.SecondaryBackgroundColor, value: 2},
				{name: ns.SecondaryFontColor, value: 3},
				{name: ns.Accent1, value: 4},
				{name: ns.Accent2, value: 5},
				{name: ns.Accent3, value: 6},
				{name: ns.Accent4, value: 7},
				{name: ns.Accent5, value: 8},
				{name: ns.Accent6, value: 9},
				{name: ns.Hyperlink, value: 10},
				{name: ns.FollowedHyperlink, value: 11},
				{name: ns.HeaderLatinFont, value: 12},
				{name: ns.HeaderEastAsianFont, value: 13},
				{name: ns.HeaderScriptFont, value: 14},
				{name: ns.HeaderLocalizedFont, value: 15},
				{name: ns.BodyLatinFont, value: 16},
				{name: ns.BodyEastAsianFont, value: 17},
				{name: ns.BodyScriptFont, value: 18},
				{name: ns.BodyLocalizedFont, value: 19}
			]
		})
	}
	ns=Microsoft.Office.WebExtension.SelectionMode;
	define({
		type: Microsoft.Office.WebExtension.Parameters.SelectionMode,
		toHost: [
			{ name: ns.Default, value: 0 },
			{ name: ns.Selected, value: 1 },
			{ name: ns.None, value: 2 },
		]
	});
	ns=Microsoft.Office.WebExtension.Parameters;
	var cns=OSF.DDA.MethodDispId;
	define({
		type: cns.dispidNavigateToMethod,
		toHost: [
			{ name: ns.Id, value: 0 },
			{ name: ns.GoToType, value: 1 },
			{ name: ns.SelectionMode, value: 2 }
		]
	});
	define({
		type: cns.dispidGetSelectedDataMethod,
		fromHost: [
			{ name: ns.Data, value: self }
		],
		toHost: [
			{ name: ns.CoercionType, value: 0 },
			{ name: ns.ValueFormat, value: 1 },
			{ name: ns.FilterType, value: 2 }
		]
	});
	define({
		type: cns.dispidSetSelectedDataMethod,
		toHost: [
			{ name: ns.CoercionType, value: 0 },
			{ name: ns.Data, value: 1 },
			{ name: ns.ImageLeft, value: 2 },
			{ name: ns.ImageTop, value: 3 },
			{ name: ns.ImageWidth, value: 4 },
			{ name: ns.ImageHeight, value: 5 }
		]
	});
	define({
		type: cns.dispidGetFilePropertiesMethod,
		fromHost: [
			{ name: OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor, value: self }
		]
	});
	define({
		type: cns.dispidGetDocumentCopyMethod,
		toHost: [{ name: ns.FileType, value: 0}],
		fromHost: [
			{ name: OSF.DDA.PropertyDescriptors.FileProperties, value: self }
		]
	});
	define({
		type: cns.dispidGetDocumentCopyChunkMethod,
		toHost: [
			{ name: OSF.DDA.FileProperties.Handle, value: 0 },
			{ name: OSF.DDA.FileSliceOffset, value: 1 },
			{ name: OSF.DDA.FileProperties.SliceSize, value: 2 }
		],
		fromHost: [
			{ name: OSF.DDA.PropertyDescriptors.FileSliceProperties, value: self }
		]
	});
	define({
		type: cns.dispidReleaseDocumentCopyMethod,
		toHost: [{ name: OSF.DDA.FileProperties.Handle, value: 0}]
	});
	define({
		type: cns.dispidAddBindingFromSelectionMethod,
		fromHost: [
			{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: self }
		],
		toHost: [
			{ name: ns.Id, value: 0 },
			{ name: ns.BindingType, value: 1 }
		]
	});
	define({
		type: cns.dispidAddBindingFromPromptMethod,
		fromHost: [
			{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: self }
		],
		toHost: [
			{ name: ns.Id, value: 0 },
			{ name: ns.BindingType, value: 1 },
			{ name: ns.PromptText, value: 2 }
		]
	});
	define({
		type: cns.dispidAddBindingFromNamedItemMethod,
		fromHost: [
			{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: self }
		],
		toHost: [
			{ name: ns.ItemName, value: 0 },
			{ name: ns.Id, value: 1 },
			{ name: ns.BindingType, value: 2 },
			{ name: ns.FailOnCollision, value: 3 }
		]
	});
	define({
		type: cns.dispidReleaseBindingMethod,
		toHost: [
			{ name: ns.Id, value: 0 }
		]
	});
	define({
		type: cns.dispidGetBindingMethod,
		fromHost: [
			{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: self }
		],
		toHost: [
			{ name: ns.Id, value: 0 }
		]
	});
	define({
		type: cns.dispidGetAllBindingsMethod,
		fromHost: [
			{ name: OSF.DDA.ListDescriptors.BindingList, value: self }
		]
	});
	define({
		type: cns.dispidGetBindingDataMethod,
		fromHost: [
			{ name: ns.Data, value: self }
		],
		toHost: [
			{ name: ns.Id, value: 0 },
			{ name: ns.CoercionType, value: 1 },
			{ name: ns.ValueFormat, value: 2 },
			{ name: ns.FilterType, value: 3 },
			{ name: OSF.DDA.PropertyDescriptors.Subset, value: 4 }
		]
	});
	define({
		type: cns.dispidSetBindingDataMethod,
		toHost: [
			{ name: ns.Id, value: 0 },
			{ name: ns.CoercionType, value: 1 },
			{ name: ns.Data, value: 2 },
			{ name: OSF.DDA.SafeArray.UniqueArguments.Offset, value: 3 }
		]
	});
	define({
		type: cns.dispidAddRowsMethod,
		toHost: [
			{ name: ns.Id, value: 0 },
			{ name: ns.Data, value: 1 }
		]
	});
	define({
		type: cns.dispidAddColumnsMethod,
		toHost: [
			{ name: ns.Id, value: 0 },
			{ name: ns.Data, value: 1 }
		]
	});
	define({
		type: cns.dispidClearAllRowsMethod,
		toHost: [
			{ name: ns.Id, value: 0 }
		]
	});
	define({
		type: cns.dispidClearFormatsMethod,
		toHost: [
			{ name: ns.Id, value: 0 }
		]
	});
		define({
		type: cns.dispidSetTableOptionsMethod,
		toHost: [
			{ name: ns.Id, value: 0 },
			{ name: ns.TableOptions, value: 1 },
		]
	});
	define({
		type: cns.dispidSetFormatsMethod,
		toHost: [
			{ name: ns.Id, value: 0 },
			{ name: ns.CellFormat, value: 1 },
		]
	});
	define({
		type: cns.dispidLoadSettingsMethod,
		fromHost: [
			{ name: OSF.DDA.SettingsManager.SerializedSettings, value: self }
		]
	});
	define({
		type: cns.dispidSaveSettingsMethod,
		toHost: [
			{ name: OSF.DDA.SettingsManager.SerializedSettings, value: OSF.DDA.SettingsManager.SerializedSettings },
			{ name: Microsoft.Office.WebExtension.Parameters.OverwriteIfStale, value: Microsoft.Office.WebExtension.Parameters.OverwriteIfStale }
		]
	});
	define({
		type: OSF.DDA.MethodDispId.dispidGetOfficeThemeMethod,
		fromHost: [
			{ name: Microsoft.Office.Internal.Parameters.OfficeTheme, value: self }
		]
	});
	define({
		type: OSF.DDA.MethodDispId.dispidGetDocumentThemeMethod,
		fromHost: [
			{ name: Microsoft.Office.Internal.Parameters.DocumentTheme, value: self }
		]
	});
	define({
		type: OSF.DDA.MethodDispId.dispidGetActiveViewMethod,
		fromHost: [
			{ name: Microsoft.Office.WebExtension.Parameters.ActiveView, value: self }
		]
	});
	define({
		type: cns.dispidAddDataPartMethod,
		fromHost: [
			{ name: OSF.DDA.PropertyDescriptors.DataPartProperties, value: self }
		],
		toHost: [
			{ name: ns.Xml, value: 0 }
		]
	});
	define({
		type: cns.dispidGetDataPartByIdMethod,
		fromHost: [
			{ name: OSF.DDA.PropertyDescriptors.DataPartProperties, value: self }
		],
		toHost: [
			{ name: ns.Id, value: 0 }
		]
	});
	define({
		type: cns.dispidGetDataPartsByNamespaceMethod,
		fromHost: [
			{ name: OSF.DDA.ListDescriptors.DataPartList, value: self }
		],
		toHost: [
			{ name: ns.Namespace, value: 0 }
		]
	});
	define({
		type: cns.dispidGetDataPartXmlMethod,
		fromHost: [
			{ name: ns.Data, value: self}
		],
		toHost: [
			{ name: ns.Id, value: 0 }
		]
	});
	define({
		type: cns.dispidGetDataPartNodesMethod,
		fromHost: [
			{ name: OSF.DDA.ListDescriptors.DataNodeList, value: self }
		],
		toHost: [
			{ name: ns.Id, value: 0 },
			{ name: ns.XPath, value: 1 }
		]
	});
	define({
		type: cns.dispidDeleteDataPartMethod,
		toHost: [
			{ name: ns.Id, value: 0 }
		]
	});
	define({
		type: cns.dispidGetDataNodeValueMethod,
		fromHost: [
			{ name: ns.Data, value: self}
		],
		toHost: [
			{ name: OSF.DDA.DataNodeProperties.Handle, value: 0 }
		]
	});
	define({
		type: cns.dispidGetDataNodeXmlMethod,
		fromHost: [
			{ name: ns.Data, value: self}
		],
		toHost: [
			{ name: OSF.DDA.DataNodeProperties.Handle, value: 0 }
		]
	});
	define({
		type: cns.dispidGetDataNodesMethod,
		fromHost: [
			{ name: OSF.DDA.ListDescriptors.DataNodeList, value: self }
		],
		toHost: [
			{ name: OSF.DDA.DataNodeProperties.Handle, value: 0 },
			{ name: ns.XPath, value: 1 }
		]
	});
	define({
		type: cns.dispidSetDataNodeValueMethod,
		toHost: [
			{ name: OSF.DDA.DataNodeProperties.Handle, value: 0 },
			{ name: ns.Data, value: 1 }
		]
	});
	define({
		type: cns.dispidSetDataNodeXmlMethod,
		toHost: [
			{ name: OSF.DDA.DataNodeProperties.Handle, value: 0 },
			{ name: ns.Xml, value: 1 }
		]
	});
	define({
		type: cns.dispidAddDataNamespaceMethod,
		toHost: [
			{ name: OSF.DDA.DataPartProperties.Id, value: 0 },
			{ name: ns.Prefix, value: 1 },
			{ name: ns.Namespace, value: 2 }
		]
	});
	define({
		type: cns.dispidGetDataUriByPrefixMethod,
		fromHost: [
			{ name: ns.Data, value: self}
		],
		toHost: [
			{ name: OSF.DDA.DataPartProperties.Id, value: 0 },
			{ name: ns.Prefix, value: 1 }
		]
	});
	define({
		type: cns.dispidGetDataPrefixByUriMethod,
		fromHost: [
			{ name: ns.Data, value: self}
		],
		toHost: [
			{ name: OSF.DDA.DataPartProperties.Id, value: 0 },
			{ name: ns.Namespace, value: 1 }
		]
	});
	define({
		type: cns.dispidGetDataNodeTextMethod,
		fromHost: [
			{ name: ns.Data, value: self}
		],
		toHost: [
			{ name: OSF.DDA.DataNodeProperties.Handle, value: 0 }
		]
	});
	define({
		type: cns.dispidSetDataNodeTextMethod,
		toHost: [
			{ name: OSF.DDA.DataNodeProperties.Handle, value: 0 },
			{ name: ns.Text, value: 1 }
		]
	});
	define({
		type: cns.dispidGetSelectedTaskMethod,
		fromHost: [
			{ name: ns.TaskId, value: self }
		]
	});
	define({
		type: cns.dispidGetTaskMethod,
		fromHost: [
			{ name: "taskName", value: 0 },
			{ name: "wssTaskId", value: 1 },
			{ name: "resourceNames", value: 2 }
		],
		toHost: [
			{ name: ns.TaskId, value: 0 }
		]
	});
	define({
		type: cns.dispidGetTaskFieldMethod,
		fromHost: [
			{ name: ns.FieldValue, value: self }
		],
		toHost: [
			{ name: ns.TaskId, value: 0 },
			{ name: ns.FieldId, value: 1 },
			{ name: ns.GetRawValue, value: 2 }
		]
	});
	define({
		type: cns.dispidGetWSSUrlMethod,
		fromHost: [
			{ name: ns.ServerUrl, value: 0 },
			{ name: ns.ListName, value: 1 }
		]
	});
	define({
		type: cns.dispidGetSelectedResourceMethod,
		fromHost: [
			{ name: ns.ResourceId, value: self }
		]
	});
	define({
		type: cns.dispidGetResourceFieldMethod,
		fromHost: [
			{ name: ns.FieldValue, value: self }
		],
		toHost: [
			{ name: ns.ResourceId, value: 0 },
			{ name: ns.FieldId, value: 1 },
			{ name: ns.GetRawValue, value: 2 }
		]
	});
	define({
		type: cns.dispidGetProjectFieldMethod,
		fromHost: [
			{ name: ns.FieldValue, value: self }
		],
		toHost: [
			{ name: ns.FieldId, value: 0 },
			{ name: ns.GetRawValue, value: 1 }
		]
	});
	define({
		type: cns.dispidGetSelectedViewMethod,
		fromHost: [
			{ name: ns.ViewType, value: 0 },
			{ name: ns.ViewName, value: 1 }
		]
	});
	cns=OSF.DDA.EventDispId
	define({ type: cns.dispidDocumentSelectionChangedEvent });
	define({
		type: cns.dispidBindingSelectionChangedEvent,
		fromHost: [
			{name: OSF.DDA.EventDescriptors.BindingSelectionChangedEvent, value: self}
		]
	});
	define({
		type: cns.dispidBindingDataChangedEvent,
		fromHost: [{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: self}]
	});
	define({ type: cns.dispidSettingsChangedEvent });
	define({
		type: cns.dispidDocumentThemeChangedEvent,
		fromHost: [
			{name: OSF.DDA.EventDescriptors.DocumentThemeChangedEvent, value: self}
		]
	});
	define({
		type: cns.dispidOfficeThemeChangedEvent,
		fromHost: [
			{name: OSF.DDA.EventDescriptors.OfficeThemeChangedEvent, value: self}
		]
	});
	define({
		type: cns.dispidActiveViewChangedEvent,
		fromHost: [{ name: OSF.DDA.EventDescriptors.ActiveViewChangedEvent, value: self}]
	});
	define({
		type: cns.dispidDataNodeAddedEvent,
		fromHost: [{ name: OSF.DDA.EventDescriptors.DataNodeInsertedEvent, value: self}]
	});
	define({
		type: cns.dispidDataNodeReplacedEvent,
		fromHost: [{ name: OSF.DDA.EventDescriptors.DataNodeReplacedEvent, value: self}]
	});
	define({
		type: cns.dispidDataNodeDeletedEvent,
		fromHost: [{ name: OSF.DDA.EventDescriptors.DataNodeDeletedEvent, value: self}]
	});
	define({ type: cns.dispidTaskSelectionChangedEvent });
	define({ type: cns.dispidResourceSelectionChangedEvent });
	define({ type: cns.dispidViewSelectionChangedEvent });
	parameterMap.define=define;
	return parameterMap;
})();
OSF.DDA.SafeArray.Delegate._onException=function OSF_DDA_SafeArray_Delegate$OnException(ex, args) {
	var status;
	var number=ex.number;
	if (number) {
		switch (number) {
			case -2146828218:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;
				break;
			case -2146827850:
			default:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
				break;
		}
	}
	if (args.onComplete) {
		args.onComplete(status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
	}
}
OSF.DDA.SafeArray.Delegate.executeAsync=function OSF_DDA_SafeArray_Delegate$ExecuteAsync(args) {
	try {
		if (args.onCalling) {
			args.onCalling();
		}
		function toArray(args) {
			var arrArgs=args;
			if (OSF.OUtil.isArray(args)) {
				var len=arrArgs.length;
				for (var i=0; i < len; i++) {
					arrArgs[i]=toArray(arrArgs[i]);
				}
			} else if (OSF.OUtil.isDate(args)) {
				arrArgs=args.getVarDate();
			} else if (typeof args==="object" && !OSF.OUtil.isArray(args)) {
				arrArgs=[];
				for (var index in args) {
					if (!OSF.OUtil.isFunction(args[index])) {
						arrArgs[index]=toArray(args[index]);
					}
				}
			}
			return arrArgs;
		}
		var startTime=(new Date()).getTime();
		if (typeof OsfOMToken !='undefined' && OsfOMToken) {
			window.external.Execute(
				args.dispId,
				toArray(args.hostCallArgs),
				function OSF_DDA_SafeArrayFacade$Execute_OnResponse(hostResponseArgs) {
					if (args.onReceiving) {
						args.onReceiving();
					}
					var result=hostResponseArgs.toArray();
					var status=result[OSF.DDA.SafeArray.Response.Status];
					if (args.onComplete) {
						var payload;
						if (status==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
							if (result.length > 2) {
								payload=[];
								for (var i=1; i < result.length; i++)
									payload[i - 1]=result[i];
							}
							else {
								payload=result[OSF.DDA.SafeArray.Response.Payload];
							}
						}
						else {
							payload=result[OSF.DDA.SafeArray.Response.Payload];
						}
						args.onComplete(status, payload);
					}
					if (OSF.AppTelemetry) {
						OSF.AppTelemetry.onMethodDone(args.dispId, args.hostCallArgs, Math.abs((new Date()).getTime() -  startTime), status);
					}
				},
				OsfOMToken
			);
		}
		else {
			window.external.Execute(
				args.dispId,
				toArray(args.hostCallArgs),
				function OSF_DDA_SafeArrayFacade$Execute_OnResponse1(hostResponseArgs) {
					if (args.onReceiving) {
						args.onReceiving();
					}
					var result=hostResponseArgs.toArray();
					var status=result[OSF.DDA.SafeArray.Response.Status];
					if (args.onComplete) {
						var payload;
						if (status==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
							if (result.length > 2) {
								payload=[];
								for (var i=1; i < result.length; i++)
									payload[i - 1]=result[i];
							}
							else {
								payload=result[OSF.DDA.SafeArray.Response.Payload];
							}
						}
						else {
							payload=result[OSF.DDA.SafeArray.Response.Payload];
						}
						args.onComplete(status, payload);
					}
					if (OSF.AppTelemetry) {
						OSF.AppTelemetry.onMethodDone(args.dispId, args.hostCallArgs, Math.abs((new Date()).getTime() -  startTime), status);
					}
				}
			);
		}
	}
	catch (ex) {
		OSF.DDA.SafeArray.Delegate._onException(ex, args);
	}
};
OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent=function OSF_DDA_SafeArrayDelegate$GetOnAfterRegisterEvent(register, args) {
	var startTime=(new Date()).getTime();
	return function OSF_DDA_SafeArrayDelegate$OnAfterRegisterEvent(hostResponseArgs) {
		if (args.onReceiving) {
			args.onReceiving();
		}
		var status=hostResponseArgs.toArray ? hostResponseArgs.toArray()[OSF.DDA.SafeArray.Response.Status] : hostResponseArgs;
		if (args.onComplete) {
			args.onComplete(status)
		}
		if (OSF.AppTelemetry) {
			OSF.AppTelemetry.onRegisterDone(register, args.dispId, Math.abs((new Date()).getTime() - startTime), status);
		}
	}
}
OSF.DDA.SafeArray.Delegate.registerEventAsync=function OSF_DDA_SafeArray_Delegate$RegisterEventAsync(args) {
	if (args.onCalling) {
		args.onCalling();
	}
	var callback=OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(true, args);
	try {
		if (typeof OsfOMToken !='undefined' && OsfOMToken) {
			window.external.RegisterEvent(
				args.dispId,
				args.targetId,
				function OSF_DDA_SafeArrayDelegate$RegisterEventAsync_OnEvent(eventDispId, payload) {
					if (args.onEvent) {
						args.onEvent(payload);
					}
					if (OSF.AppTelemetry) {
						OSF.AppTelemetry.onEventDone(args.dispId);
					}
				},
				callback,
				OsfOMToken
			);
		}
		else {
			window.external.RegisterEvent(
				args.dispId,
				args.targetId,
				function OSF_DDA_SafeArrayDelegate$RegisterEventAsync_OnEvent1(eventDispId, payload) {
					if (args.onEvent) {
						args.onEvent(payload);
					}
					if (OSF.AppTelemetry) {
						OSF.AppTelemetry.onEventDone(args.dispId);
					}
				},
				callback
			);
		}
	}
	catch (ex) {
		OSF.DDA.SafeArray.Delegate._onException(ex, args);
	}
};
OSF.DDA.SafeArray.Delegate.unregisterEventAsync=function OSF_DDA_SafeArray_Delegate$UnregisterEventAsync(args) {
	if (args.onCalling) {
		args.onCalling();
	}
	var callback=OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(false, args);
	try {
		  if (typeof OsfOMToken !='undefined' && OsfOMToken) {
				window.external.UnregisterEvent(
					args.dispId,
					args.targetId,
					callback,
					OsfOMToken
				);
		}
		else{
				window.external.UnregisterEvent(
					args.dispId,
					args.targetId,
					callback
				);
		}
	}
	catch (ex) {
		OSF.DDA.SafeArray.Delegate._onException(ex, args);
	}
};
OSF.DDA.SafeArray.Delegate.MessageParent=function OSF_DDA_SafeArray_Delegate$MessageParent(args){
	try {
		if (args.onCalling) {
			args.onCalling();
		}
		var startTime=(new Date()).getTime();
		var message=args.hostCallArgs[Microsoft.Office.WebExtension.Parameters.MessageToParent];
		if (typeof window.external.MessageParent2 !='undefined' && typeof OsfOMToken !='undefined' && OsfOMToken) {
			var targetOrigin=args.hostCallArgs[Microsoft.Office.WebExtension.Parameters.TargetOrigin];
			window.external.MessageParent2(message, targetOrigin, OsfOMToken);
		}
		else {
			window.external.MessageParent(message);
		}
		if (args.onReceiving) {
			args.onReceiving();
		}
		if (OSF.AppTelemetry) {
			OSF.AppTelemetry.onMethodDone(args.dispId, args.hostCallArgs, Math.abs((new Date()).getTime() - startTime), result);
		}
		return result;
	}
	catch (ex) {
		var status;
		var number=ex.number;
		if (number) {
		switch (number) {
			case -2146828218:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;
				break;
			case -2146827850:
			default:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
				break;
			}
		}
		return status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
	}
}

