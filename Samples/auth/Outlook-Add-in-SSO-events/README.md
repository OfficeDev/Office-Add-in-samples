---
title: "Use SSO with event-based activation in an Outlook add-in"
page_type: sample
urlFragment: outlook-add-in-sso-event-based-activation
products:
- office-add-ins
- office-outlook
- office
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 05/05/2022 10:00:00 AM
description: "Use Outlook Smart Alerts to verify that required color categories are applied to a new message or appointment before it's sent."
---

# Use Outlook Smart Alerts

**Applies to**: Outlook on Windows | [new Outlook on Mac](https://support.microsoft.com/office/6283be54-e74d-434e-babb-b70cefc77439)

## Summary

This sample uses Outlook Smart Alerts to verify that required color categories are applied to a new message or appointment before it's sent.

For documentation related to this sample, see the following articles.

- [Use Smart Alerts and the OnMessageSend event in your Outlook add-in](https://learn.microsoft.com/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough)
- [Configure your Outlook add-in for event-based activation](https://learn.microsoft.com/office/dev/add-ins/outlook/autolaunch)

## Features

Use Outlook Smart Alerts to verify that required color categories are applied to a new message or appointment before it's sent.


## Applies to

- Outlook on Windows starting in Version 2206 (Build 15330.20196)
- [New Outlook on Mac](https://support.microsoft.com/office/6283be54-e74d-434e-babb-b70cefc77439) starting in Version 16.65.827.0

> **Note**: Although the Smart Alerts feature is supported in Outlook on the web, Windows, and new Mac UI (see the "Supported clients and platforms" section of [Use Smart Alerts and the onMessageSend and OnAppointmentSend events in your Outlook add-in](https://learn.microsoft.com/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough#supported-clients-and-platforms)), this sample only runs in Outlook on Windows and Mac.
>
> As the Office.Categories API can't be used in Compose mode in Outlook on the web, this sample isn't supported on that client. To learn how to develop a Smart Alerts add-in for Outlook on the web, see the [Smart Alerts walkthrough](https://learn.microsoft.com/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough).

## Prerequisites

- A Microsoft 365 subscription. If you don't have a Microsoft 365 subscription, you can get a [free developer sandbox](https://aka.ms/m365/devprogram#Subscription) that provides a renewable 90-day Microsoft 365 E5 subscription for development purposes.

## Solution

| Solution | Authors |
| -------- | --------- |
| Use Outlook Smart Alerts to verify that required color categories are applied to a new message or appointment before it's sent. | Microsoft |

## Version history

| Version | Date | Comments |
| ------- | ----- | -------- |
| 1.0 | 02-17-2023 | Initial release |

## Run the sample

Run this sample in Outlook on Windows or Mac. The add-in's web files are served from this repository on GitHub.

1. Download the **manifest.xml** file from this sample to a folder on your computer.

1. Sideload the add-in manifest in Outlook on Windows or Mac by following the manual instructions in [Sideload Outlook add-ins for testing](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing).

### Try it out




## Run the sample from localhost

If you prefer to configure a web server and host the add-in's web files from your computer, use the following steps.

1. Install a recent version of [npm](https://www.npmjs.com/get-npm) and [Node.js](https://nodejs.org/) on your computer. To verify if you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.

1. You need http-server to run the local web server. If you haven't installed this yet, you can do this with the following command.

    ```console
    npm install --global http-server
    ```

1. You need Office-Addin-dev-certs to generate self-signed certificates to run the local web server. If you haven't installed this yet, you can do this with the following command.

    ```console
    npm install --global office-addin-dev-certs
    ```

1. Clone or download this sample to a folder on your computer, then go to that folder in a console or terminal window.

1. Run the following command to generate a self-signed certificate to use for the web server.

   ```console
    npx office-addin-dev-certs install
    ```

    This command will display the folder location where it generated the certificate files.

1. Go to the folder location where the certificate files were generated, then copy the **localhost.crt** and **localhost.key** files to the cloned or downloaded sample folder.

1. Run the following command.

    ```console
    http-server -S -C localhost.crt -K localhost.key --cors . -p 3000
    ```

    The http-server will run and host the current folder's files on localhost:3000.

1. Now that your localhost web server is running, you can sideload the **manifest-localhost.xml** file provided in the sample folder. Using this file, follow the steps in [Run the sample](#run-the-sample) to sideload and run the add-in.

## Key parts of the sample


## Known issues

- In the Windows client, imports are not supported in the JavaScript file where you implement event-based activation handling.

## Copyright

Copyright (c) 2023 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

---

<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/samples/outlook-add-in-sso-event-based-activation" />
