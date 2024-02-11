# Outlook Spam Reporting Add-in Sample

[Spam Reporting dialog](/assets/readme/outlook-spam-processing-dialog.png)

## Summary

This sample showcases how to build a spam reporting solution that:

- is easily discoverable in the Ribbon
- provides the user with a processing dialog for the email to be reported
- facilitate easily saving a copy of the email to a file that can be submitted to your back-end system

> [!IMPORTANT]
>
> The integrated spam-reporting feature is currently in preview in Outlook on Windows. Features in preview shouldn't be used in production add-ins. We invite you to try out this feature in test or development environments and welcome feedback on your experience through GitHub (see the **Questions and feedback** section at the end of this page).

## Prerequisites
- Microsoft 365

> Note: If you don't have a Microsoft 365 subscription, you can get a [free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) that provides a renewable 90-day Microsoft 365 E5 subscription for development purposes.

- To preview the integrated spam-reporting feature in Outlook on Windows, you must install Version 2307 (Build 16626.10000) or later. Then, join the [Microsoft 365 Insider program](https://insider.microsoft365.com/join/Windows) and select the **Beta Channel** option to access Office beta builds.

> [!TIP]
> If you're unable to choose a channel in your Outlook client, see [Let users choose which Microsoft 365 Insider channel to install on Windows devices](https://learn.microsoft.com/en-us/deployoffice/insider/deploy/user-choice?fbclid=IwAR3VzO4_HySIoNw735IZSRVBLQm_s83Cje4arT7kviE7HwaOoQYHPI0tF04).

## Configure the sample

> [!IMPORTANT]
> To test the `getAsFileAsync` method while it's still in preview in Outlook on Windows, you must configure your computer's registry.
>
> Outlook on Windows includes a local copy of the production and beta versions of Office.js instead of loading from the content delivery network (CDN). By default, the local production copy of the API is referenced. To reference the local beta copy of the API, you must configure your computer's registry as follows:
>
> 1. In the registry, navigate to `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer`. If the key doesn't exist, create it.
> 1. Create an entry named `EnableBetaAPIsInJavaScript` and set its value to `1`.
>
>    :::image type="content" source="../images/outlook-beta-registry-key.png" alt-text="The EnableBetaAPIsInJavaScript registry value is set to 1.":::

Manifest Sample

````xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.13">
        <bt:Set Name="Mailbox"/>
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <Runtimes>
          <Runtime resid="WebViewRuntime.Url">
            <!-- References the JavaScript file that contains the spam-reporting event handler. This is used by Outlook on Windows. -->
            <Override type="javascript" resid="JSRuntime.Url"/>
          </Runtime>
        </Runtimes>
        <DesktopFormFactor>
          <FunctionFile resid="WebViewRuntime.Url"/>
          <!-- Implements the integrated spam-reporting feature in the add-in. -->
          <ExtensionPoint xsi:type="ReportPhishingCommandSurface">
            <ReportPhishingCustomization>
              <!-- Configures the ribbon button. -->
              <Control xsi:type="Button" id="spamReportingButton">
                <Label resid="spamButton.Label"/>
                <Supertip>
                  <Title resid="spamButton.Label"/>
                  <Description resid="spamSuperTip.Text"/>
                </Supertip>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Action xsi:type="ExecuteFunction">
                  <FunctionName>onSpamReport</FunctionName>
                </Action>
              </Control>
              <!-- Configures the preprocessing dialog. -->
              <PreProcessingDialog>
                <Title resid="PreProcessingDialog.Label"/>
                <Description resid="PreProcessingDialog.Text"/>
                <ReportingOptions>
                  <Title resid="OptionsTitle.Label"/>
                  <Option resid="Option1.Label"/>
                  <Option resid="Option2.Label"/>
                  <Option resid="Option3.Label"/>
                </ReportingOptions>
                <FreeTextLabel resid="FreeText.Label"/>
                <MoreInfo>
                  <MoreInfoText resid="MoreInfo.Label"/>
                  <MoreInfoUrl resid="MoreInfo.Url"/>
                </MoreInfo>
              </PreProcessingDialog>
             <!-- Identifies the runtime to be used. This is also referenced by the Runtime element. -->
              <SourceLocation resid="WebViewRuntime.Url"/>
            </ReportPhishingCustomization> 
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="WebViewRuntime.Url" DefaultValue="https://localhost:3000/commands.html"/>
        <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/spamreporting.js"/>
        <bt:Url id="MoreInfo.Url" DefaultValue="https://www.contoso.com/spamreporting"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="spamButton.Label" DefaultValue="Report Spam Message"/>
        <bt:String id="PreProcessingDialog.Label" DefaultValue="Report Spam Message"/>
        <bt:String id="OptionsTitle.Label" DefaultValue="Why are you reporting this email?"/>
        <bt:String id="FreeText.Label" DefaultValue="Provide additional information, if any:"/>
        <bt:String id="MoreInfo.Label" DefaultValue="To learn more about reporting unsolicited messages, see "/>
        <bt:String id="Option1.Label" DefaultValue="Received spam email."/>
        <bt:String id="Option2.Label" DefaultValue="Received a phishing email."/>
        <bt:String id="Option3.Label" DefaultValue="I'm not sure this is a legitimate email."/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="spamSuperTip.Text" DefaultValue="Report an unsolicited message."/>
        <bt:String id="PreProcessingDialog.Text" DefaultValue="Thank you for reporting this message."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</VersionOverrides>
````

## Solution


| Solution                      | Author(s)    |
| ------------------------------- | -------------- |
| Outlook Spam Reporting Add-in | [Eric Legault](https://www.linkedin.com/in/ericlegault/) |

## Version history


| Version | Date              | Comments        |
| --------- | ------------------- | ----------------- |
| 1.0     | February 12 2024 | Initial release |

## Run the sample

---

Run this sample in Outlook on Windows or in a browser. The add-in web files are served from this repo on GitHub.

1. Download the **manifest.xml** file from this sample to a folder on your computer.
2. Sideload the add-in manifest in Outlook on the web or on Windows by following the manual instructions in the article [Sideload Outlook add-ins for testing](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing).
3. Choose a message from your inbox, then select the add-in's button from the ribbon:
[Spam Reporting button](/assets//readme/outlook-spam-ribbon-button.png)
4. In the preprocessing dialog, choose a reason for reporting the message and add information about the message, if configured. Then, select **Report**.
5. (Optional) In the post-processing dialog, select **OK**:
[Spam Reporting post-processing dialog](/assets//readme/outlook-spam-post-processing-dialog.png)

## Run the sample from localhost

If you prefer to host the web server for the sample on your computer, follow these steps.

1. Install a recent version of [npm](https://www.npmjs.com/get-npm) and [Node.js](https://nodejs.org/) on your computer. To verify if you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.
2. You need http-server to run the local web server. If you haven't installed this yet, run the following command.

   ```console
   npm install --global http-server
   ```
3. Use a tool such as openssl to generate a self-signed certificate that you can use for the web server. Move the cert.pem and key.pem files to the root folder for this sample.
4. From a command prompt, go to the root folder and run the following command.

   ```console
   http-server -S --cors . -p 3000
   ```
5. To reroute to localhost, run office-addin-https-reverse-proxy. If you haven't installed this, run the following command.

   ```console
   npm install --global office-addin-https-reverse-proxy
   ```

   To reroute, run the following in another command prompt.

   ```console
   office-addin-https-reverse-proxy --url http://localhost:3000
   ```
6. Sideload `manifest.xml` in Outlook on the web or on Windows by following the manual instructions in the article [Sideload Outlook add-ins for testing](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing).
7. Open Outlook, select an email, and click the Report Spam Message Ribbon button in the Report Group on the Home tab.

## References

- [Implement an integrated spam-reporting add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/spam-reporting)
- [ReportPhishingCommandSurface Extension Point](https://learn.microsoft.com/en-us/javascript/api/manifest/extensionpoint?view=outlook-js-preview&preserve-view=true#reportphishingcommandsurface-preview)
- [Office.MessageRead.getAsFileAsync() method](https://learn.microsoft.com/en-us/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#outlook-office-messageread-getasfileasync-member(1))
- [Configure your Outlook add-in for event-based activation](https://learn.microsoft.com/office/dev/add-ins/outlook/autolaunch)
- [Debug your event-based Outlook add-in](https://learn.microsoft.com/office/dev/add-ins/outlook/debug-autolaunch)
- Other samples:
  - [Encrypt attachments, process meeting request attendees, and react to appointment date/time changes using Outlook event-based activation](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-encrypt-attachments)
  - [Use Outlook event-based activation to set the signature](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-set-signature)
  - [Use Outlook event-based activation to tag external recipients](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-tag-external)
- [Microsoft Office Add-in Debugger Extension for Visual Studio Code](https://learn.microsoft.com/office/dev/add-ins/testing/debug-with-vs-extension)
- [Develop Office Add-ins with Visual Studio Code](https://learn.microsoft.com/office/dev/add-ins/develop/develop-add-ins-vscode)
- [Office Add-ins with Visual Studio Code](https://code.visualstudio.com/docs/other/office)
- [Debugging with Visual Studio Code](https://code.visualstudio.com/docs/editor/debugging)
- [Node.js debugging in VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging)
- [Office-Addin-Debugging](https://www.npmjs.com/package/office-addin-debugging)

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2024 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/outlook-spam-reporting" />