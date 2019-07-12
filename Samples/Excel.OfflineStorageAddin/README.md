# Using storage techniques to access Office add-in data offline

## Summary

This sample demonstrates how you can implement `localStorage` to enable limited functionality for your Office add-in when a user experiences lost connection.

## Applies to

-  Office on Windows (Word, Excel, PowerPoint)
-  Office on Mac (Word, Excel, PowerPoint)
-  Office on the web (Word, Excel, PowerPoint)

## Prerequisites

Before running this sample, make sure you have installed a recent version of [npm](https://www.npmjs.com/get-npm) and [Node.js](https://nodejs.org/en/) on your computer. To check if you have already installed these tools, run the commands `node -v` and `npm -v` in your terminal.

## Solution

Solution | Author(s)
---------|----------
Excel.OfflineStorageAddin | Nancy Wang, Albert Dotson (**Microsoft**)

## Version history

Version  | Date | Comments
---------| -----| --------
1.0  | (update later) | Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

----------
# Scenario: Caching data using offline storage techniques
This sample Office add-in inserts a table of basketball players' stats in your file, retrieved from a local file named `sampleData.js`. In this sample code, data from the add-in is stored in `localStorage` to allow users who previously opened the add-in with online connection to insert the table of stats offline.

While this add-in gets its data from a local server, implementation of `localStorage` as shown in this sample can be extended to add-ins that get their data from online sources. Furthermore, although this sample runs only in Excel, `localStorage can be used to offline data across Word, Excel, and PowerPoint.

**Note**: `localStorage` can store up to 5MB of data. To store larger amounts of data offline and for improved performance, consider using [IndexedDB](https://developer.mozilla.org/en-US/docs/Web/API/IndexedDB_API). Note that as of now, IndexedDB isn't supported by all browsers used by Office add-ins and may cause your add-in to fail on some computers. 
Furthermore, `localStorage` is not secure and should only be used for data that can be made publicly available. To protect add-in data while offline, consider using offline cookies.
As another option, you can store your add-in's data in [`Office.Settings`](https://docs.microsoft.com/en-us/javascript/api/office/office.settings?view=office-js). Using Office.Settings will enable your add-in's offline capabilities to persist within the file (e.g. if you'd like to share your file with others).

## Build and run the sample

1. Clone or download this repository. 
2. In the command line, navigate to the **Excel.OfflineStorageAddin** folder from your root directory.

The following code sample shows you how to do these two steps: 
```command&nbsp;line

$ git clone https://github.com/OfficeDev/PnP-OfficeAddins.git
$ cd PnP-OfficeAddins
$ cd Samples
$ cd Excel.OfflineStorageAddin
```
You can try out this sample by running the following commands:
```command&nbsp;line
# this will download the node modules needed to run this add-in
$ npm install

# this will build the add-in 
$ npm run build

# this will start the server to run your add-in on Excel on your desktop
$ npm run start

# this will start the server to run your add-in on Excel on the web
$ npm run start:web
```
## Key parts of this sample

Navigate to *Excel.OfflineStorageAddin/src/taskpane/taskpane.js* to find the implementation of `localStorage` described below. 

### Implementing `localStorage` to offline data
The *Excel.OfflineStorageAddin/src/taskpane/taskpane.js* file contains the `loadTable()` function, that uses `localStorage` to display a table of basketball player stats when a user loses connection.

In the sample code, the `loadTable()` function first checks if the basketball player data was previously stored in `localStorage`, as shown in the code below. If it exists, the data is parsed from a JSON object back into readable text format before being passed to `createTable()`, a function which creates a table from the given data. 

```js
if (localStorage.DraftPlayerData) {
    var dataObject = JSON.parse(localStorage.DraftPlayerData);
    createTable(dataObject);
}
```

If the data wasn't previously stored, `loadTable()` attempts to access the offline data file, *sampleData.js*, through an AJAX call. If this attempt is successful, the function passes the data returned from the file to `createTable()` to produce a table. Before being stored into `localStorage`, the data is also converted into a JSON object so that it can be easily parsed when `localStorage` is accessed. However, if the function is unable to access the *sampleData.js* file, the function returns an error to the console. This process is shown in the following code:
```
else {
    $.ajax({
        dataType: "json",
        url: "sampleData.js",
        success: function (result, status, xhr) {
            localStorage.DraftPlayerData = JSON.stringify(result);
            createTable(result);
        },
        error: function (xhr, status, error) {
            console.log("Player data failed to load with error: " + error);
        }
    });
}
```

## Security notes

In the webpack.config.js file, a header is set to  `"Access-Control-Allow-Origin": "*"`. This is only for development purposes. You should lock this header down to only allowed domains in production code.

You will be prompted to install self-signed certificates when you run this sample on your development computer. The certificates are intended only for running and studying this code sample. Do not reuse them in your own code solutions or in production environments.

You can install or uninstall the self-signed certificates by running the following commands in the project folder.

```command&nbsp;line
npx office-addin-dev-certs install
npx office-addin-dev-certs uninstall
```
<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/excel-custom-functions/storage" />


<img src="https://telemetry.sharepointpnp.com/officedev/samples/readme-template" />
