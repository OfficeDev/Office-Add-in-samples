# FAQ

### Setup/Getting Started

**I'm new to Office Add-ins, where do I start?**

If you have never developed an Office Web Add-in before we recommend you to visit our [Get Started](http://dev.office.com/getting-started/addins) experience to understand the basics. Afterwards you can then play with Add-in commands which is a new feature that is in preview currently by following the steps outlined in the [readme](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/README.md).

**How do I get the latest Office build needed to try the feature?**

Please see the [readme](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/README.md). The very first step explains how to get the build and activate the feature.

**Where can I find the manifest reference and samples?**

Please see the [readme](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/README.md).

**I deployed the Add-in manifest using a SharePoint App Catalog, which shows as "My organization" in the insertion dialog and I don't see buttons on the Ribbon, Why?**

Deploying Add-ins with commands via the SharePoint Add-in Catalog is not supported

### Debug: Add-in or buttons are not showing up

1.  Ensure you are using a supported client/build and catalog. As stated above, the SharePoint App Catalog is not a supported mechanism to deploy add-ins with commands.
2.  Start with samples
3.  Do small tweaks and validate you manifest using the [Office Add-in Validator](https://github.com/OfficeDev/office-addin-validator) 
4.  Double check the reference documentation.
5.  Verify that in your VersionOverrides you are targeting the correct host. Sometimes folks assume that the hosts declared on the top of the manifest
6.  Verify that you are using the correct Tab element. OfficeTab is to add commands to an existing Office Tab and requires that you pass an existing Id. CustomTab is to create a new tab. Consult the reference documentation for more details.
7.  Use **[RuntimeLogging](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Tools/RuntimeLogging.md)** to diagnose issues with your manifest.

### Debug: ExecuteFunction not working

**ExecuteFunction isn't working, what are the most common issues?**

1.  Check that the FunctionFile is loading properly, use Fiddler to see if a network call is being issued.
2.  Ensure you are using _HTTPS and that the certificate doesn't give any warnings_ as this would prevent the FunctionFile from loading. If you use a local server sometimes using the IP will warn but using localhost would work fine.
3.  Make sure you manifest has the correct resource ID and that the URL for your function file is correct
4.  Ensure that the name of your FunctionFile in the manifest is the same as your function in javascript.
5.  Verify that the function is defined in the GLOBAL scope for javascript. A function defined inside a different scope won't work.

### Debug: Icons not showing

**The buttons display but icons aren't showing, what are the most common issues?**

1.  Check that the URLs of the icons are valid.
2.  Check you are using a supported file format for your icon. We recommnend PNG. 
3.  Ensure you are using _HTTPS and that the certificate doesn't give any warnings_ as this would prevent icons from loading. If you use a local server sometimes using the IP will warn but using localhost would work fine.
4.  Make sure you _DO NOT_ send any **no-cache/no-store** headers back as this might prevent icons from being stored and used
5.  Make sure you manifest has the correct resource ID and that the URL for your icon file is correct

### Debug: Misc

**Will users still have to go to the insertion dialog to make the add-ins show their buttons?**

Once your add-in is installed it will have its buttons permanently displayed on the Ribbon

**I found an issue, I have a question or I have a feature request, where do I log that?**

*   Issues with the samples please use [Issues](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/issues) of this repo to log.
*   Question/Additional help use **StackOverflow** and tag with **office-js**..

New feature requests please log them at [user voice](https://officespdev.uservoice.com/)
