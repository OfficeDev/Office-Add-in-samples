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
  createdDate: 10/06/2020 1:25:00 PM
description: "This sample shows how to use web workers in custom functions to prevent blocking the UI of your Office Add-in."
---

# Custom function sample using web worker

## Summary

This sample shows how to use web workers in custom functions to prevent blocking the UI of your Office Add-in.

## Features

- Custom Functions
- Web workers

## Applies to

- Excel on Windows, Mac, and in a browser.

## Solution

Solution | Author(s)
---------|----------
Office Add-in Custom Function Using Web Workers | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.0 | 10-06-2020 | Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

----------

## Scenario

Custom functions block the UI of your Office Add-in when they run. If you have long-running custom functions, this can cause poor performance in your Office Add-in UI when the spreadsheet is calculated. For example, if someone has a table with thousands of rows, each of which is calling a long-running custom function, this can lead to the UI being blocked during a recalculation.

You can unblock the UI by using web workers to do the calculations for your custom functions.

## Run the sample

You can run this sample in Excel in a browser. The add-in web files are served from this repo on GitHub.

1. Download the **manifest.xml** file from this sample to a folder on your computer.
1. Open [Office on the web](https://office.live.com/).
1. Choose **Excel**, and then open a new document.
1. Open the **Insert** tab on the ribbon and choose **Office Add-ins**.
1. On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.
   ![The Office Add-ins dialog with a drop-down in the upper right reading "Manage my add-ins" and a drop-down below it with the option "Upload My Add-in"](images/office-add-ins-my-account.png)
1. Browse to the add-in manifest file, and then select **Upload**.
   ![The upload add-in dialog with buttons for browse, upload, and cancel.
](images/upload-add-in.png)
1. Verify that the add-in loaded successfully. You will see a **Web worker task pane** button on the **Home** tab on the ribbon.

Now you can use the following custom functions:

```
=WebWorkerSample.TEST(2)
=WebWorkerSample.TEST_PROMISE(2)
=WebWorkerSample.TEST_ERROR(2)
=WebWorkerSample.TEST_ERROR_PROMISE(2)
```

## Run the sample from Localhost

If you want to host the sample from your computer, run the following command:

```console
cd webworker-customfunction
http-server --cors .
office-addin-https-reverse-proxy --url http://localhost:8080
```

If the office-addin certificate is not found or expired, run the following command:

```console
npx office-addin-dev-certs install --days 365
```

Sideload the add-in using the the previous steps (1 - 7). Upload the `manifest-localhost.xml` file for step 6.

## Details

### Dispatch to web worker

When a custom function needs to use a web worker, we turn the calcuation into a job and dispatch it to a web worker. The **dispatchCalculationJob** function takes the function name and parameters from a custom function, and creates a job object that is posted to a web worker. For more details see the **dispatchCalculationJob** function in [functions.js](functions.js).

```JavaScript
    // Post a job to web worker to do calculation
    function dispatchCalculationJob(functionName, parameters) {
        var jobId = g_nextJobId++;
        return new Promise(function(resolve, reject) {
            // store the promise information.
            g_jobIdToPromiseInfoMap[jobId] = {resolve: resolve, reject: reject};
            var worker = getOrCreateWebWorker(jobId);
            worker.postMessage({
                jobId: jobId,
                name: functionName,
                parameters: parameters
            });
        });
    }
```

### Web worker runs the job and returns result

The web worker runs the job specified in the job object to do the actual calculation. The web worker code is in a separate file in [functions-worker.js](functions-worker.js). 

The functions-worker.js will
1. Receive a message containing the job to run.
2. Invoke a function to perform the calculation.
3. Call **postMessage** to post the result back to the main thread.

```JavaScript
self.addEventListener('message',
    function(event) {
        var data = event.data;
        if (typeof(data) == "string") {
            data = JSON.parse(data);
        }

        var jobId = data.jobId;
        try {
            var result = invokeFunction(data.name, data.parameters);
            // check whether the result is a promise
            if (typeof(result) == "function" || typeof(result) == "object" && typeof(result.then) == "function") {
                result.then(function(realResult) {
                    postMessage(
                        {
                            jobId: jobId,
                            result: realResult
                        }
                    );
                })
                .catch(function(ex) {
                    postMessage(
                        {
                            jobId: jobId,
                            error: true
                        }
                    )
                });
            }
            else {
                postMessage({
                    jobId: jobId,
                    result: result
                });
            }
        }
        catch(ex) {
            postMessage({
                jobId: jobId,
                error: true
            });
        }
    }
);

```

Most of the previous code handles the error case and Promise case.

### Process results from the web worker

In [functions.js](functions.js), when a new web worker is created, it is provided a callback function to process the result. The callback function parses the data to determine the outcome of the job. It resolves or rejects the promise as determined by the job result data.

```JavaScript
        // create a new web worker
        var webWorker = new Worker("functions-worker.js");
        webWorker.addEventListener('message', function(event) {
            var data = event.data;
            if (typeof(data) == "string") {
                data = JSON.parse(data);
            }

            if (typeof(data.jobId) == "number") {
                var jobId = data.jobId;
                // get the promise info associated with the job id
                var promiseInfo = g_jobIdToPromiseInfoMap[jobId];
                if (promiseInfo) {
                    if (data.error) {
                        // The web worker returned error
                        promiseInfo.reject(new Error());
                    }
                    else {
                        // The web worker retuned result
                        promiseInfo.resolve(data.result);
                    }
                    delete g_jobIdToPromiseInfoMap[jobId];
                }
            }
        });
```

## Copyright

Copyright (c) 2020 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://telemetry.sharepointpnp.com/officedev/samples/excel-custom-function-web-workers" />