// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

var SampleNamespace = {};

(function(SampleNamespace) {
    // The max number of web workers to be created
    var g_maxWebWorkers = 4;

    // The array of web workers
    var g_webworkers = [];
    
    // Next job id
    var g_nextJobId = 0;

    // The promise info for the job. It stores the {resolve: resolve, reject: reject} information for the job.
    var g_jobIdToPromiseInfoMap = {};

    function getOrCreateWebWorker(jobId) {
        var index = jobId % g_maxWebWorkers;
        if (g_webworkers[index]) {
            return g_webworkers[index];
        }

        // create a new web worker
        var webWorker = new Worker("functions-worker.js");
        webWorker.addEventListener('message', function(event) {
            var jobResult = event.data;
            if (typeof(jobResult) == "string") {
                jobResult = JSON.parse(jobResult);
            }

            if (typeof(jobResult.jobId) == "number") {
                var jobId = jobResult.jobId;
                // get the promise info associated with the job id
                var promiseInfo = g_jobIdToPromiseInfoMap[jobId];
                if (promiseInfo) {
                    if (jobResult.error) {
                        // The web worker returned an error
                        promiseInfo.reject(new Error());
                    }
                    else {
                        // The web worker returned a result
                        promiseInfo.resolve(jobResult.result);
                    }
                    delete g_jobIdToPromiseInfoMap[jobId];
                }
            }
        });

        g_webworkers[index] = webWorker;
        return webWorker;
    }

    // Post a job to the web worker to do the calculation
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

    SampleNamespace.dispatchCalculationJob = dispatchCalculationJob;
})(SampleNamespace);


CustomFunctions.associate("TEST", function(n) {
    return SampleNamespace.dispatchCalculationJob("TEST", [n]);
});

CustomFunctions.associate("TEST_PROMISE", function(n) {
    return SampleNamespace.dispatchCalculationJob("TEST_PROMISE", [n]);
});

CustomFunctions.associate("TEST_ERROR", function(n) {
    return SampleNamespace.dispatchCalculationJob("TEST_ERROR", [n]);
});

CustomFunctions.associate("TEST_ERROR_PROMISE", function(n) {
    return SampleNamespace.dispatchCalculationJob("TEST_ERROR_PROMISE", [n]);
});


// This function will show what happens when calculations are run on the main UI thread.
// The task pane will be blocked until this method completes.
CustomFunctions.associate("TEST_UI_THREAD", function(n) {
    var ret = 0;
    for (var i = 0; i < n; i++) {
        ret += Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(50))))))))));
        for (var l = 0; l < n; l++) {
            ret -= Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(50))))))))));     
        }
    }
    return ret;
});

