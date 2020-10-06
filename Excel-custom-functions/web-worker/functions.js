var SampleNamespace = {};

(function(SampleNamespace) {
    // The max number of web worker to be created
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

        g_webworkers[index] = webWorker;
        return webWorker;
    }

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
