---
page_type: sample
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
---

# Using storage to share data between UI-less custom functions and the task pane

If you need to share data values between your UI-less custom functions and the task pane, you can use the OfficeRuntime.storage object. UI-less custom functions and task do not share the same runtime and cannot access the same data. OfficeRuntime.storage saves simple key/value pairs that you can access from both UI-less custom functions and the task pane.

This sample accompanies the article [Save and share state in UI-less custom functions](https://docs.microsoft.com/office/dev/add-ins/excel/custom-functions-save-state)

## Applies to

- UI-less custom functions on Excel desktop and online

**Note:** Shared runtime is now recommended for most custom functions scenarios. This sample applies to UI-less custom functions only. 

## Prerequisites

To get set up and working with custom functions, see [Custom functions requirements](https://docs.microsoft.com/office/dev/add-ins/excel/custom-functions-requirements)

### Solution ###

Solution | Author(s)
---------|----------
Using storage to share data between UI-less custom functions and the task pane | Microsoft

### Version history ###

Version  | Date | Comments
---------| -----| --------
1.0  | May 1, 2019 | Initial release

### Disclaimer ###

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

## Sample: Sharing data between custom functions and the task pane

This sample code shows how to share data between UI-less custom functions and the task pane. The task pane allows the user to enter a key/value pair and save it to storage. Then in a UI-less custom function, the value can be retrieved using the `GETVALUE(key)` custom function. Or the user can use the `STOREVALUE(key,value)` custom function to store a value, and then retrieve it in the task pane.

## Run the sample

To run this sample, download the code and go to the **Storage** folder in a command prompt window.

1. Run `npm install`.
2. Run `npm run build`.
3. Run `npm run start`. The sample will now sideload into Excel on desktop.

### How the custom functions work with storage

The /src/functions/functions.js file contains two custom functions named `StoreValue` and `GetValue`.

`StoreValue` takes a key and value from the user and stores them by calling the `OfficeRuntime.storage.setItem` method as shown in the following sample code.

```js
function StoreValue(key, value) {
  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

`GetValue` retrieves a value for a given key by calling the `OfficeRuntime.storage.getItem` method as shown in the following sample code.

```js
function GetValue(key) {
  return OfficeRuntime.storage.getItem(key);
}
```

### How the task pane works with storage

The /src/taskpane/taskpane.html has two JavaScript functions that are called from buttons on the UI. The `SendTokenToCustomFunction` function retrieves the key and token from text boxes on the task pane. Then it calls the `OfficeRuntime.storage.setItem` method to store the key/value pair as shown in the following sample code.

```js
function SendTokenToCustomFunction() {
  var token = document.getElementById('tokenTextBox').value;
  var tokenSendStatus = document.getElementById('tokenSendStatus');
  var key = "token";
  OfficeRuntime.storage.setItem(key, token).then(function () {
    tokenSendStatus.value = "Success: Item with key '" + key + "' saved to Storage.";
  }, function (error) {
    tokenSendStatus.value = "Error: Unable to save item with key '" + key + "' to Storage. " + error;
  });
}
```

The `ReceiveTokenFromCustomFunction` function retrieves the key from a text box on the task pane. Then it calls the `OfficeRuntime.storage.getItem` method to get value for the key and display it on the page.

```js
function ReceiveTokenFromCustomFunction() {
  var key = "token";
  var tokenSendStatus = document.getElementById('tokenSendStatus');
  OfficeRuntime.storage.getItem(key).then(function (result) {
    tokenSendStatus.value = "Success: Item with key '" + key + "' read from Storage.";
    document.getElementById('tokenTextBox2').value = result;
  }, function (error) {
    tokenSendStatus.value = "Error: Unable to read item with key '" + key + "' from Storage. " + error;
  });
}
```

## Security notes

### webpack.config.js

In the webpack.config.js file, a header is set to  `"Access-Control-Allow-Origin": "*"`. This is only for development purposes. You should lock this header down to only allowed domains in production code. 

### Self-signed certificates

You will be prompted to install self-signed certificates when you run this sample on your development computer. The certificates are intended only for running and studying this code sample. Do not reuse them in your own code solutions or in production environments.

You can install or uninstall the self-signed certificates by running the following commands in the project folder.

```cli
npx office-addin-dev-certs install
npx office-addin-dev-certs uninstall
```
<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/excel-custom-functions/storage" />
