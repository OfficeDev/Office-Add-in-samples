# Outlook Event-based Activation Demo

## Summary

This sample showcases how to use event-based activation in an Outlook add-in.  It demonstrates how to run tasks based on events that fire when certain data changes after user interaction when composing an email message or editing an appointment or meeting request.

## Features

- Encrypts the first attachment that is added to a compose email or appointment and adds it as an another attachment with a "encrypted_" prefix to the file name, that decrypts that attachment and adds it as an another attachment with a "decrypted_" prefix to the file name
  - Also add a notification message to the email to denote that encryption and decryption is in progress. When completed, that message is removed and another notification message is added noting that the process has completed
- Adds notification messages to a meeting request when recipients are added or removed (these are removed when there are no more any recipients):
  - Show a message with a running tally of the number of required and optional attendees
  - Show a message with a warning if one or more distribution lists are invited as an attendee
- Adds a notification message to an appointment when the user changes the date/time, showing the original date/time that was set when the appointment was opened (to serve as a reference for further date/time edits)

[TODO: add pictures]

## Applies to

- Outlook
  - Windows
  - web browser

## Prerequisites

- Microsoft 365

> Note: If you do not have a Microsoft 365 subscription, you can get one for development purposes by signing up for the Microsoft 365 developer program.

## Solution

| Solution      | Author(s) |
| ------------- | --------- |
| Use Outlook Event-based activation to process item attachments, meeting request recipients and appointment date/time changes | Eric Legault    |

## Version history

| Version | Date                               | Comments        |
| ------- | ---------------------------------- | --------------- |
| 1.0     | 10-10-2021                  | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Build and run the solution

### Install Types for Office JavaScript Preview API

- `npm install --save-dev @types/office-js-preview`
- See https://docs.microsoft.com/en-ca/office/dev/add-ins/reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview
- Clone this repo, or download the sample.
- in the command line run:

  - `npm install`
  - `npm run build`
  - `npm start`

To debug in Outlook Online:

- Upload the manifest.xml file in your "My add-ins" page in "Add-ins for Outlook" settings
- run `npm run start:web` (or from "RUN AND DEBUG" in VS Code, choose "Node.js..." -> "Run Script: start:web")

## Notes

- At present, imports are not supported in the JavaScript file where you implement the handling for event-based activation.
- verify that crypto-js is referenced in the .html file (commands.html in this sample) that has a reference to the commands file (commands.js in this sample): `<script type="text/javascript" src="../../node_modules/crypto-js/crypto-js.js"></script>`
- Instead of using localStorage to manage state (as is done for caching the appointments original date/time in this sample), you can use RoamingSettings (see [&#39;Manage state and settings for an Outlook add-in&#39;](https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/manage-state-and-settings-outlook)) or use the [Office.SessionData interface](https://docs.microsoft.com/en-us/javascript/api/outlook/office.sessiondata?view=outlook-js-preview) in the Preview API.
- If you get eslint errors ("Parsing error: Cannot read file '.../tsconfig.json'"), ensure this line is in the .eslintrc.json file:
  `"project": "outlook-encrypt-attachments/tsconfig.json"`.
  Or add this to the .vscode\settings.json file:

```json
"eslint.workingDirectories": 
  [
    "src"
  ]
```

## References

- [Configure your Outlook add-in for event-based activation](https://docs.microsoft.com/en-ca/office/dev/add-ins/outlook/autolaunch)
- [Debug your event-based Outlook add-in (preview)](https://docs.microsoft.com/en-ca/office/dev/add-ins/outlook/debug-autolaunch)
- Other samples:
  - [https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-set-signature](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-set-signature)
  - [https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-tag-external](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-tag-external)
- [crypto-js](https://www.npmjs.com/package/crypto-js)
- [Microsoft Office Add-in Debugger Extension for Visual Studio Code](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/debug-with-vs-extension)
- [Develop Office Add-ins with Visual Studio Code](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/develop-add-ins-vscode)
- [Office Add-ins with Visual Studio Code](https://code.visualstudio.com/docs/other/office)
- [Debugging with Visual Studio Code](https://code.visualstudio.com/docs/editor/debugging)
- [Node.js debugging in VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging)
- [Office-Addin-Debugging](https://www.npmjs.com/package/office-addin-debugging)

## Code Attributions

- Getting recipients and attendees: `https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/outlook/30-recipients-and-attendees/get-set-required-attendees-appointment-organizer.yaml`

<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/samples/outlook-attachments-attendees-appointment-dates" />
