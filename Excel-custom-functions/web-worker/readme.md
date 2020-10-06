# Custom Function Sample using Web Worker

# Purpose
This add-in is used to show how to use web worker for custom function.

The add-in contains:
- Custom Function

# Steps to run the addin
On Excel Online, insert addin using file upload. The manifest is `manifest.xml`. The agave is servered by https://officedev.github.io/testing-assets/addins/webworker-customfunction/. No need to run any "npm" commands.

Now, you could uee the following functions
```
=WebWorkerSample.TEST(2)
=WebWorkerSample.TEST_PROMISE(2)
=WebWorkerSample.TEST_ERROR(2)
=WebWorkerSample.TEST_ERROR_PROMISE(2)
```

# Steps for Maintainers
On dev machine, run the following command so that we could access the website using https://localhost/home.html
```console
cd webworker-customfunction
http-server --cors .
office-addin-https-reverse-proxy --url http://localhost:8080
```

If the office-addin certificate is not found or expired, please run
```console
npx office-addin-dev-certs install --days 365
```

Then insert addin using file uploader and the manifest is `manifest-localhost.xml`.

# Details
## Dispatch to web worker
To use web worker for custom function, we need to create web worker and then dispatch the calcuation job to the web worker. Please check the code in the [functions.js](functions.js).

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

## Web worker do calcuation and send result back
The web worker will do the real calculation. Please check the [functions-worker.js](functions-worker.js). The functions-worker.js will
1. listen to the message
2. invoke calucation
3. then call postMessage to post the result to the main page.

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
Most of the above code is to handle the error case and Promise case.

## Process results from web worker
The [functions.js](functions.js) listens to the message from the web worker and then resolve the proise or reject the promise.
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

# Maintainers
shaofengzhu
