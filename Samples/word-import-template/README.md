---
page_type: sample
urlFragment: word-import-template
products:
  - office-word
  - office
  - m365
  - office-teams
languages:
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: 03/08/2024 4:00:00 PM
description: "Shows how to import templates in a Word document."
---

# Import templates in a Word document

## Summary

This sample shows how to import a Word document template with an add-in.

## Description

The user updates their Word document with the content from another Word document, treating the external document like a template. The user selects a Word document through the add-in UI then it's applied to the current document.

![Import template add-in task pane.](./resources/word-import-template.png)

## Applies to

- Word on Windows
- Word on Mac
- Word on the web

## Solution

| Solution | Authors |
|----------|-----------|
| How to import a template in a Word document | Microsoft |

## Version history

| Version  | Date | Comments |
|----------|------|----------|
| 1.0 | 03-08-2024 | Initial release |
| 1.1 | 07-07-2025 | Add support for the unified manifest for Microsoft 365 |

## Decide on a version of the manifest

- Add-in only manifest
  - To run the add-in only manifest, which is the **manifest.xml** file in the sample's root directory **Samples/word-import-template**, go to the [Add-in only manifest](#add-in-only-manifest) section.
- [Unified manifest for Microsoft 365](https://learn.microsoft.com/office/dev/add-ins/develop/json-manifest-overview)
  - To run the unified manifest for Microsoft 365 (**manifest.json**), go to the [Unified manifest](#unified-manifest) section.

## Add-in only manifest

### Run the sample

Use one of the following add-in file hosting options to run the sample.

#### Use GitHub as the web host

You can run this sample in Word on Windows, on Mac, or in a browser. The add-in web files are served from this repo on GitHub.

1. Download the **manifest.xml** file from this sample to a folder on your computer.
1. Sideload the add-in manifest in Word by following the appropriate instructions in the article [Sideload an Office Add-in for testing](https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing).
1. Follow the steps in [Try it out](#try-it-out) to test the sample.

#### Use localhost

If you prefer to configure a web server and host the add-in's web files from your computer, use the following steps.

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

1. Sideload **manifest-localhost.xml** in Word by following the appropriate instructions in the article [Sideload an Office Add-in for testing](https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing).

1. Follow the steps in [Try it out](#try-it-out) to test the sample.

## Unified manifest

### Prerequisites

- If you want to run the web server on localhost, install a recent version of [npm](https://www.npmjs.com/get-npm) and [Node.js](https://nodejs.org) on your computer. To check if you've already installed these tools, from a command prompt, run the following commands.

    ```console
    node -v
    npm -v
    ```

- If you want to run the sample using GitHub as the web host, install the [Microsoft 365 Agents Toolkit command line interface (CLI)](https://learn.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-cli). From a command prompt, run the following command.

    ```console
    npm install -g @microsoft/teamsapp-cli
    ```

### Run the sample

You can run this sample in Word on Windows, on Mac, or in a browser. Use one of the following add-in file hosting options.

#### Use GitHub as the web host

The quickest way to run the sample is to use GitHub as the web host. However, you can't debug or change the source code. The add-in web files are served from this GitHub repository.

1. Download the **manifest-configurations/unified/word-import-template.zip** file from this sample to a folder on your computer.
1. Sideload the add-in manifest in Word by following the appropriate instructions in the article [Sideload Office Add-ins that use the unified manifest for Microsoft 365](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-add-in-with-unified-manifest).
1. Follow the steps in [Try it out](#try-it-out) to test the sample.

#### Use localhost

If you prefer to host the web server on localhost, follow these steps:

1. Clone or download this repository.
1. From a command prompt, go to the root of the project folder **/Samples/word-import-template**.
1. Copy the files from the **manifest-configurations/unified** subfolder to the root folder.
1. Run the following commands.

    ```console
    npm install
    ```

    ```console
    npm start
    ```

    This starts the web server on localhost and sideloads the **manifest.json** file to Word.

1. Follow the steps in [Try it out](#try-it-out) to test the sample.

1. To stop the web server and uninstall the add-in from Word, run the following command.

    ```console
    npm stop
    ```

## Try it out

Once the add-in is loaded, use the following steps to try out the functionality.

1. Open Word on Windows, on Mac, or in a browser.

1. To open the add-in task pane, go to the **Home** tab and choose **Show Task Pane**.

1. In the "Template" section of the add-in UI, select **Choose File**. Navigate to the location of your .docx file then open the file. The template is automatically applied to your document, replacing any preexisting content.

    ![The initial screen displaying the button to choose a file.](./resources/word-import-template-initial-screen.png)

    For convenience, the resources folder of this project includes a *template example.docx* file.

    ![Screen showing the imported template.](./resources/word-import-template-applied.png)

1. In the document, update the text and other content.

## Make it yours

The following are a few suggestions for how you could tailor this to your scenario.

- Include [single sign-on (SSO)](https://learn.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins) to support managing sessions and persisting settings for the user.
- Provide personalized or company-approved templates for users to access.
- Enable users to personalize templates and save to shared location.

## Related content

- [Import template](https://learn.microsoft.com/office/dev/add-ins/word/import-template)
- [Word add-ins documentation](https://learn.microsoft.com/office/dev/add-ins/word/)

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2024 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/word-import-template" />
