# Office Add-ins Patterns and Practices (PnP)

Office Add-ins PnP is a community driven effort that helps developers extend, build, and provision customizations for the Office platform. The source is maintained on this GitHub repo where anyone can participate. You can provide contributions to the samples, reusable components, and documentation. Office Add-ins PnP is owned and coordinated by Office engineering teams, but the work is done by the community for the community.

## List of recent samples

- [Custom function batching pattern](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/Batching). If your custom functions call a remote service you may want to use a batching pattern to reduce the number of network calls to the remote service. This is useful when a spreadsheet recalculates and it contains many of your custom functions. Recalculate will result in many calls to your custom functions, but you can batch them into one or a few calls to the remote service.
- [Custom function storage](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/Storage) for custom functions. Custom functions and task panes cannot directly communicate with each other. See how to use the Storage object to send data between custom functions and task panes. This is especially useful for sharing an access token.

## More information

Please use [http://aka.ms/OfficeDevPnP](http://aka.ms/OfficeDevPnP) for getting latest information around the whole *Office 365 Developer Patterns and Practices program*.

## Code of conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
