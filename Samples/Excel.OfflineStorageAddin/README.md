# Using offline storage techniques to cache add-in data

## Summary

This sample demonstrates how you can implement local storage to enable limited functionality for your add-in when a user experiences lost connection.

## Applies to

-  Office on Windows (Word, Excel, PowerPoint)
-  Office on Mac (Word, Excel, PowerPoint)
-  Office on the web (Word, Excel, PowerPoint)

## Prerequisites

None

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
# Scenario: Offlining data using offline storage techniques

## Build and run the sample

To get this sample, clone or download this repository and in the command line, navigate to the **Excel.OfflineStorageAddin** folder.
```
$ git clone https://github.com/OfficeDev/PnP-OfficeAddins.git
$ cd Samples
$ cd Excel.OfflineStorageAddin
```
You can try out this sample by running the following commands in the command line:
```
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
This sample displays a table of NBA players' stats, retrieved from a local file named sampleData.js. 

## Security notes

In the webpack.config.js file, a header is set to  `"Access-Control-Allow-Origin": "*"`. This is only for development purposes. You should lock this header down to only allowed domains in production code.

You will be prompted to install self-signed certificates when you run this sample on your development computer. The certificates are intended only for running and studying this code sample. Do not reuse them in your own code solutions or in production environments.

You can install or uninstall the self-signed certificates by running the following commands in the project folder.

```cli
npx office-addin-dev-certs install
npx office-addin-dev-certs uninstall
```
<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/excel-custom-functions/storage" />

Description of the sample with additional details beyond the summary.
This sample illustrates the following concepts on top of the Office platform:

- topic 1
- topic 2
- topic 3

<img src="https://telemetry.sharepointpnp.com/officedev/samples/readme-template" />
