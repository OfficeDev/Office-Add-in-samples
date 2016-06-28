# Outlook Add-in consuming the Microsoft Graph API

### Summary ###
This is a sample Outlook Add-in - built using Microsoft Visual Studio 2015 - that consumes the Microsoft Graph API via REST/AJAX, using ADAL.JS and the current user's context in Office 365.

The sample add-in is triggered by any email message received and with some specific keywords in the subject (*[Offer Request]*). The add-in searches for the email sender in the Contacts of the current tenant, and if there is any matching, it shows some related information. Moreover, it searches for documents in OneDrive for Business containing the sender's name in their content.

This sample is part of the code samples related to the book ["Programming Office 365"](https://www.microsoftpressstore.com/store/programming-microsoft-office-365-includes-current-book-9781509300914) written by [Paolo Pialorsi](https://twitter.com/PaoloPia) and published by Microsoft Press.

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



