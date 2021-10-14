# Use Outlook event-based activation to encrypt attachments, process meeting request attendees and react to appointment date/time changes

**Applies to**: Outlook on Windows | Outlook on the web

## Summary
---
This sample showcases how to use event-based activation in an Outlook add-in when the user composes an email or appointment/meeting request.  It demonstrates how to run tasks based on events that fire when certain data changes when the user:
- adds an attachment to an email or appointment/meeting request
- adds recipients or distributions lists as required or optional attendeees in a meeting request
- changes the start or end date or time in an appointment/meeting request
- adds a notification message to the item when a user newly composes an email or appointment/meeting request, instructing the user to open the Task Pane for further information (for this sample, it shows instructions and details for this sample!)

## Features/Scenario

- Encrypts the first attachment that is added to a compose email or appointment and adds it as an another attachment with a "encrypted_" prefix to the file name, then decrypts that attachment and adds it as an another attachment with a "decrypted_" prefix to the file name
  - Also adds a notification message to the compose item to denote that encryption and decryption is in progress. When completed, that message is removed (it may only appear for a very brief time, depending on the complexity of the encryption process) and another notification message is added noting that the process has completed: ![Compose email](assets/readme/compose_email.png)  
- Adds notification messages to a meeting request when recipients are added or removed (these are removed when there are no longer any recipients):
  - Show a message with a running tally of the number of required and optional attendees
  - Show a message with a warning if one or more distribution lists are invited as an attendee
- Adds a notification message to an appointment when the user changes the date/time, showing the original date/time that was set when the appointment was opened (to serve as a reference for further date/time edits): ![Compose email](assets/readme/appointment_Outlook_desktop.png)  

## Applies to
---
- Outlook
  - Windows
  - web browser

## Prerequisites
---
- Microsoft 365

> Note: If you do not have a Microsoft 365 subscription, you can get one for development purposes by signing up for the Microsoft 365 developer program.

## Solution
---
| Solution                                                                                                                     | Author(s)    |
| ---------------------------------------------------------------------------------------------------------------------------- | ------------ |
| Use Outlook Event-based activation to process item attachments, meeting request recipients and appointment date/time changes | Eric Legault |

## Version history
---
| Version | Date       | Comments        |
| ------- | ---------- | --------------- |
| 1.0     | 10-14-2021 | Initial release |

## Run the sample
---
You can run this sample in Outlook on Windows or in a browser. The add-in web files are served from this repo on GitHub.

1. Download the **manifest.xml** file from this sample to a folder on your computer.
2. Sideload the add-in manifest in Outlook on the web or on Windows by following the manual instructions in the article [Sideload Outlook add-ins for testing](https://docs.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing).

### Try it out

Once the add-in is loaded use the following steps to try out the functionality:

1. Open Outlook on Windows or in a browser
2. Create a new message or appointment
    > You should see a notification at the top of the message that reads: **Open the Task Pane for details about running the Outlook Event-based Activation Sample Add-in.** | Show Task Pane | Dismiss
3. Add an attachment
    > You should see a notification at the top of the message that reads: **The '{file name} attachment has been encrypted and decrypted and added as reference attachments for your review** | Dismiss


1. Create a new meeting request
2. Add a user as a required or optional attendee
    > You should see a notification at the top of the message that reads: **Your appointment has 1 required and 0 optional attendees** | Dismiss
3. Add a distribution list as a required or optional attendee
    > You should see a notification at the top of the message that reads: **Warning! Your appointment has a distribution list! Make sure you have chosen the correct one!** | Dismiss
4. Change the start or end date or time
    > You should see a notification at the top of the message that reads: **Original date/time: Start = ##/##/#### #:##:## ##; End = ##/##/#### #:##:## ##** | Dismiss

## Run the sample from localhost

If you prefer to host the web server for the sample on your computer, follow these steps:

1. Install a recent version of [npm](https://www.npmjs.com/get-npm) and [Node.js](https://nodejs.org/) on your computer. To verify if you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.
1. You need http-server to run the local web server. If you haven't installed this yet, run the following command.

    ```console
    npm install --global http-server
    ```

1. Use a tool such as openssl to generate a self-signed certificate that you can use for the web server. Move the cert.pem and key.pem files to the root folder for this sample.
1. From a command prompt, go to the root folder and run the following command.

    ```console
    http-server -S --cors . -p 3000
    ```

1. To reroute to localhost, run office-addin-https-reverse-proxy. If you haven't installed this, run the following command.

    ```console
    npm install --global office-addin-https-reverse-proxy 
    ```

    To reroute, run the following in another command prompt.

    ```console
    office-addin-https-reverse-proxy --url http://localhost:3000 
    ```

1. Sideload `manifest.xml` in Outlook on the web or on Windows by following the manual instructions in the article [Sideload Outlook add-ins for testing](https://docs.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing).
1. [Try out the sample!](#try-it-out)

## Build and run the solution

- Install Types for Office JavaScript Preview API
  - `npm install --save-dev @types/office-js-preview`
- in the command line run:
  - `npm install`
  - `npm run build`
  - `npm start`

## To debug in Outlook Online:

- Upload the manifest.xml file in your "My add-ins" page in "Add-ins for Outlook" settings
- run `npm run start:web` (or from "RUN AND DEBUG" in VS Code, choose "Node.js..." -> "Run Script: start:web")

## To debug in Outlook for Windows:

> Prerequisites: See [Debug your event-based Outlook add-in (preview)](https://docs.microsoft.com/en-ca/office/dev/add-ins/outlook/debug-autolaunch)

- Run `npm start`; this should start Outlook for Windows
- If the "Register Outlook Developer Add-in Manifest" dialog appears, click OK:

![Register dialog](assets/readme/RegisterDialog.png "Register dialog")
- Verify that the `Computer\HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\Developer\93011807-161e-4cc1-846f-eb7721919e4e` registry key exists (this corresponds to the Id value in the manifest.xml file) and **UseDirectDebugger** is set to 1 (running `npm start` should do that automatically)
- Compose a new message or appointment
- Wait for the "Debug Event-based handler" dialog to appear; do NOT click OK (or Cancel)
- Open the `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]\bundle.js` file in Visual Studio Code and set breakpoints
- Run the "Direct Debugging" command in the RUN AND DEBUG dropdown
- Click OK in the "Debug Event-based handler" dialog
- Interact with the Outlook item to trigger breakpoints in your code
- **NOTE**: You may experience issues where the bundle.js file doesn't refresh after recent code changes. In these cases, try restarting the Outlook client. Breakpoints may also stop working. An alternate way of debugging allows `console` functions to write to a local text file. See [Runtime logging on Windows](https://docs.microsoft.com/en-ca/office/dev/add-ins/testing/runtime-logging#runtime-logging-on-windows) for more details

## Key parts of this sample

### Configure event-based activation in the manifest

The manifest configures a runtime that is loaded specifically to handle event-based activation. The following `<Runtime>` element specifies an HTML page resource id that loads the runtime on Outlook on the web. The `<Override>` element specifies the JavaScript file instead, to load the runtime for Outlook on Windows. Outlook on Windows doesn't use the HTML page to load the runtime.

```xml
<Runtime resid="WebViewRuntime.Url">
  <Override type="javascript" resid="JSRuntime.Url"/>
...
<bt:Url id="WebViewRuntime.Url" DefaultValue="https://localhost:3000/src/commands/commands.html" />
<bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/src/commands/commands.js" />
```

The add-in handles six events that are mapped to various functions:

```xml
<LaunchEvents>
  <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler" /> 
  <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler" />                 
  <LaunchEvent Type="OnAppointmentAttendeesChanged" FunctionName="onAppointmentAttendeesChangedHandler" />
  <LaunchEvent Type="OnAppointmentTimeChanged" FunctionName="onAppointmentTimeChangedHandler" /> 
  <!-- The same function handles both OnMessageAttachmentsChanged and OnAppointmentAttachmentsChanged-->
  <LaunchEvent Type="OnMessageAttachmentsChanged" FunctionName="onMessageAttachmentsChangedHandler" /> 
  <LaunchEvent Type="OnAppointmentAttachmentsChanged" FunctionName="onMessageAttachmentsChangedHandler" />
</LaunchEvents>
```

### Handling the events

When the user creates a new message or appointment, Outlook will load the files specified in the manifest to handle the following events:

| Event | Handler |
| --- | ---|
| `OnNewMessageCompose` | onMessageComposeHandler |
| `OnNewAppointmentOrganizer` | onAppointmentComposeHandler |
| `OnAppointmentAttendeesChanged` | onAppointmentAttendeesChangedHandler |
| `OnAppointmentTimeChanged` | onAppointmentTimeChangedHandler |
| `OnMessageAttachmentsChanged` | *onItemAttachmentsChangedHandler |
| `OnAppointmentAttachmentsChanged` | *onItemAttachmentsChangedHandler |

***NOTE**: The onItemAttachmentsChangedHandler function handles both OnMessageAttachmentsChanged and OnAppointmentAttachmentsChanged

Outlook on the web will load the `commands.html` page, which then also loads `commands.js`.

### Task pane code

The task pane code is located under the `taskpane` folder of this project. The task pane HTML and JavaScript files only provide a UI with details about this sample.

- `taskpane_appt_compose.html` is loaded when the user clicks the "Open Task Pane" link in the notification message or clicks the Show Task Pane button in the Ribbon.
- `taskpane_msg_compose.html` is loaded when the user clicks the "Open Task Pane" link in the notification message or clicks the Show Task Pane button in the Ribbon.

## Notes

- At present, imports are not supported in the JavaScript file where you implement the handling for event-based activation. This means that external libraries (like `cryto-js`) cannot be required directly in the `commands.js` file and must be loaded in `commands.html`. Since Outlook desktop only loads `commands.js`, encryption of attachments will not work on that platform
- `console.dir()` methods cannot be used in Outlook desktop
- Clicking the "Show Task Pane" link in the InfoBar doesn't work in Outlook Online. A fix is in progress and being tested: https://github.com/OfficeDev/office-js/issues/2125
- If you get eslint errors ("Parsing error: Cannot read file '.../tsconfig.json'"), ensure this line is in the .eslintrc.json file: `"project": "outlook-encrypt-attachments/tsconfig.json"`. Or add this to the .vscode\settings.json file: `"eslint.workingDirectories": [ "src" ]`

## References

- [Configure your Outlook add-in for event-based activation](https://docs.microsoft.com/en-ca/office/dev/add-ins/outlook/autolaunch)
- [Debug your event-based Outlook add-in (preview)](https://docs.microsoft.com/en-ca/office/dev/add-ins/outlook/debug-autolaunch)
- Other samples:
  - [https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-set-signature](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-set-signature)
  - [https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-tag-external](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-tag-external)
- [crypto-js](https://www.npmjs.com/package/crypto-js)
- [Office.SessionData interface](https://docs.microsoft.com/en-us/javascript/api/outlook/office.sessiondata?view=outlook-js-preview)
- [Microsoft Office Add-in Debugger Extension for Visual Studio Code](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/debug-with-vs-extension)
- [Develop Office Add-ins with Visual Studio Code](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/develop-add-ins-vscode)
- [Office Add-ins with Visual Studio Code](https://code.visualstudio.com/docs/other/office)
- [Debugging with Visual Studio Code](https://code.visualstudio.com/docs/editor/debugging)
- [Node.js debugging in VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging)
- [Office-Addin-Debugging](https://www.npmjs.com/package/office-addin-debugging)

## Code Attributions

- Getting recipients and attendees: `https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/outlook/30-recipients-and-attendees/get-set-required-attendees-appointment-organizer.yaml`

## Copyright

Copyright (c) 2021 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/samples/outlook-encrypt-attachments" />
