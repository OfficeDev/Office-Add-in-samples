/* global clearInterval, console, setInterval */

/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
export function add(first, second) {
  return first + second;
}

/**
 * Displays the current time once a second
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
export function clock(invocation) {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
export function currentTime() {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param {number} incrementBy Amount to increment
 * @param {CustomFunctions.StreamingInvocation<number>} invocation
 */
export function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param {string} message String to write.
 * @returns String to write.
 */
export function logMessage(message) {
  console.log(message);

  return message;
}

const SampleNamespace = {};

(function(SampleNamespace) {
    // The max number of web workers to be created
    const g_maxWebWorkers = 4;

    // The array of web workers
    const g_webworkers = [];
    
    // Next job id
    let g_nextJobId = 0;

    // The promise info for the job. It stores the {resolve: resolve, reject: reject} information for the job.
    const g_jobIdToPromiseInfoMap = {};

    function getOrCreateWebWorker(jobId) {
        const index = jobId % g_maxWebWorkers;
        if (g_webworkers[index]) {
            return g_webworkers[index];
        }

        // create a new web worker
        const webWorker = new Worker("functionssWorker.js");
        webWorker.addEventListener('message', function(event) {
            let jobResult = event.data;
            if (typeof(jobResult) == "string") {
                jobResult = JSON.parse(jobResult);
            }

            if (typeof(jobResult.jobId) == "number") {
                const jobId = jobResult.jobId;
                // get the promise info associated with the job id
                const promiseInfo = g_jobIdToPromiseInfoMap[jobId];
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
        const jobId = g_nextJobId++;
        return new Promise(function(resolve, reject) {
            // store the promise information.
            g_jobIdToPromiseInfoMap[jobId] = {resolve: resolve, reject: reject};
            const worker = getOrCreateWebWorker(jobId);
            worker.postMessage({
                jobId: jobId,
                name: functionName,
                parameters: parameters
            });
        });
    }

    SampleNamespace.dispatchCalculationJob = dispatchCalculationJob;
})(SampleNamespace);

/**
 * Add two numbers
 * @customfunction
 * @param {number} n First number
 * @returns {number} The sum of the two numbers.
 */
export function TEST(n) {
    return SampleNamespace.dispatchCalculationJob("TEST", [n]);
}

/**
 * Add two numbers
 * @customfunction
 * @param {number} n First number
 * @returns {number} The sum of the two numbers.
 */
export function TEST_PROMISE(n) {
    return SampleNamespace.dispatchCalculationJob("TEST_PROMISE", [n]);
}

/**
 * Add two numbers
 * @customfunction
 * @param {number} n First number
 * @returns {number} The sum of the two numbers.
 */
export function TEST_ERROR(n) {
    return SampleNamespace.dispatchCalculationJob("TEST_ERROR", [n]);
}

/**
 * Add two numbers
 * @customfunction
 * @param {number} n First number
 * @returns {number} The sum of the two numbers.
 */
export function TEST_ERROR_PROMISE(n) {
    return SampleNamespace.dispatchCalculationJob("TEST_ERROR_PROMISE", [n]);
}

/**
 * Add two numbers
 * @customfunction
 * @param {number} n First number
 * @returns {number} The sum of the two numbers.
 */
export function TEST_UI_THREAD(n) {
    let ret = 0;
    for (let i = 0; i < n; i++) {
        ret += Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(50))))))))));
        for (let l = 0; l < n; l++) {
            ret -= Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(50))))))))));     
        }
    }
    return ret;
}

/** 
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
    let ret = 0;
    for (let i = 0; i < n; i++) {
        ret += Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(50))))))))));
        for (let l = 0; l < n; l++) {
            ret -= Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(50))))))))));     
        }
    }
    return ret;
});

*/
