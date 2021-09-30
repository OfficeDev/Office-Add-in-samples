---
page_type: sample
urlFragment: excel-custom-function-batching-pattern
products:
- office-excel
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 5/1/2019 1:25:00 PM
description: "If your functions call a remote service you may want to use a batching pattern to reduce the number of network calls to the service."
---

# Custom function batching pattern

If your custom functions call a remote service you may want to use a batching pattern to reduce the number of network calls to the remote service. This is useful when a spreadsheet recalculates and it contains many of your custom functions. Recalculate will result in many calls to your custom functions, but you can batch them into one or a few calls to the remote service.

## Applies to

- Custom functions on Excel on Windows, Mac, and on the web

## Prerequisites

To get set up and working with custom functions, see [Custom functions requirements](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-requirements)

## Solution

Solution | Author(s)
---------|----------
Custom function batching | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.0  | April 5, 2019 | Initial release
1.1 | June 1, 2021 | Update to use GitHub hosting

## Scenario: Custom function batching

In this scenario your custom functions call a remote service. To reduce network round trips you will batch all the calls and send them in a single call to the web service. This is ideal when the spreadsheet is recalculated. For example, if someone used your custom function in 100 cells in a spreadsheet, and then recalculates the spreadsheet, your custom function would run 100 times and make 100 network calls. By using this batching pattern, the calls can be combined to make all 100 calculations in a single network call.

## Run the sample

You can run this sample in Excel in a browser. The add-in web files are served from this repo on GitHub.
1. Download the **manifest.xml** file from this sample to a folder on your computer.
1. Open [Office on the web](https://office.live.com/).
1. Choose **Excel**, and then open a new document.
1. Open the **Insert** tab on the ribbon and choose **Office Add-ins**.
1. On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.
   ![The Office Add-ins dialog with a drop-down in the upper right reading "Manage my add-ins" and a drop-down below it with the option "Upload My Add-in"](../../Samples/images/office-add-ins-my-account.png)
1. Browse to the add-in manifest file, and then select **Upload**.
   ![The upload add-in dialog with buttons for browse, upload, and cancel.
](../../Samples/images/upload-add-in.png)
1. Verify that the add-in loaded successfully. You will see a **Show Taskpane** button on the **Home** tab on the ribbon.
Once the add-in is loaded use the following steps to try out the functionality.

## Key parts of the sample

The code pattern contains two custom functions named `DIV2` and `MUL2`. Instead of performing the calculation, each of them calls a `_pushOperation` function to push the operation into a batch queue to be passed to a web service.

```javascript
function mul2(first, second) {
  return _pushOperation(
    "mul2",
    [first, second]
  );
}
```

### Batching the operation

The `_pushOperation` function pushes each operation into a _batch variable. It schedules the batch call to be made within 100 milliseconds. You can adjust this when using the code in your own solution.

```javascript
  // If a remote request hasn't been scheduled yet,
  // schedule it after a certain timeout, e.g. 100 ms.
  if (!_isBatchedRequestScheduled) {
    _isBatchedRequestScheduled = true;
    setTimeout(_makeRemoteRequest, 100);
  }
```

### Making the remote request

The `_makeRemoteRequest` function prepares the batch request and passes it to the `_fetchFromRemoteService` function. If you are adapting this code to your own solution you need to modify `_makeRemoteRequest` to actually call your remote service.

### The remote service

The `_fetchFromRemoteService` function processes the batch of operations, performs the operations, and then returns the results. In this sample, `_fetchFromRemoteService` is just another function to demonstrate the pattern. When adapting this code to your solution, use this method on the server-side to respond to the client call over the network.

## How to apply batching in your own solution

You can copy and paste this code into your own solution. When using this pattern, you'll need to evaluate and update the following areas of code.

### _pushOperation

Adjust the timeout value as needed. A longer time will be more noticeable to the user. A shorter time may result in more calls to the remote service.

### _makeRemoteRequest

Modify this function to actually make a network call to your remote service and pass the batch operations in a single call. For example, you may want to serialize the batch entries into a JSON body to be passed in the network call to the remote service.

### _fetchFromRemoteService

Place this function in your remote service to handle the network call from the client. You'll want to modify this to perform the actual operations of your custom functions (or call the correct methods to do so.)

**Note**: You should remove the call to `pause(1000)` which simulates network latency in the sample.

## Run the sample from Localhost

If you prefer to host the web server for the sample on your computer, follow these steps:

1. You need http-server to run the local web server. If you haven't installed this yet you can do this with the following command:
    
    ```console
    npm install --global http-server
    ```
    
2. Use a tool such as openssl to generate a self-signed certificate that you can use for the web server. Move the cert.pem and key.pem files to the root folder for this sample.
3. From a command prompt, go to the root folder and run the following command:
    
    ```console
    http-server -S --cors . -p 3000
    ```
    
4. To reroute to localhost run office-addin-https-reverse-proxy. If you haven't installed this you can do this with the following command:
    
    ```console
    npm install --global office-addin-https-reverse-proxy
    ```
    
    To reroute run the following in another command prompt:
    
    ```console
    office-addin-https-reverse-proxy --url http://localhost:3000
    ```
    
5. Follow the steps in [Run the sample](#run-the-sample), but upload the `manifest-localhost.xml` file for step 6.

## Security notes

When implementing the **_fetchFromRemoteService** function on a server, apply an appropriate authentication mechanism. Ensure that only the correct callers can access the function.

## Copyright

Copyright (c) 2019 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/excel-custom-functions/batching" />
