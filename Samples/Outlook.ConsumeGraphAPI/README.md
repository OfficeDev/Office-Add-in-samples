# Outlook Add-in consuming the Microsoft Graph API

### Summary ###
This is a sample Outlook Add-in - built using Microsoft Visual Studio 2015 - that consumes the Microsoft Graph API via REST/AJAX, using ADAL.JS and the current user's context in Office 365.

The sample add-in is triggered by any email message received and with some specific keywords in the subject (*[Offer Request]*). The add-in searches for the email sender in the Contacts of the current tenant, and if there is any matching, it shows some related information. Moreover, it searches for documents in OneDrive for Business containing the sender's name in their content.

This sample is part of the code samples related to the book ["Programming Microsoft Office 365"](https://www.microsoftpressstore.com/store/programming-microsoft-office-365-includes-current-book-9781509300914) written by [Paolo Pialorsi](https://twitter.com/PaoloPia) and published by Microsoft Press.

### Applies to ###
-  Outlook 2013/2016, Outlook Web App, Outlook Mobile on Microsoft Office 365

### Solution ###
Solution | Author(s) | Twitter
---------|-----------|--------
Outlook.AddInSample.sln | Paolo Pialorsi (PiaSys.com) | [@PaoloPia](https://twitter.com/PaoloPia)

### Version history ###
Version  | Date | Comments
---------| -----| --------
1.0  | June 28th 2016 | Initial release

### Setup Instructions ###
In order to play with this add-in, you need to:

-  Sign up for a developer subscription for Office 365 [Office Dev Center](http://dev.office.com/), if you don't have one
-  Register the custom add-in as an Azure Active Directory application
-  Configure the Azure AD application with the following delegated permissions for Microsoft Graph: Read user files, Read user contacts
-  Enable OAuth 2.0 implicit flow capability for the Azure AD application
-  Configure the azureADTenant and azureADClientID variables in the [MessageRead.js](./Outlook.AddInSampleWeb/MessageRead.js) file of the solution


<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/Outlook.AddInSample" />