const SampleNamespace = {};

(function (SampleNamespace) {
  // The max number of web workers to be created.
  const maxWebWorkers = 4;

  // The array of web workers.
  const webworkers = [];

  // Next job ID.
  let nextJobId = 0;

  // The promise info for the job. It stores the {resolve: resolve, reject: reject} information for the job.
  const jobIdToPromiseInfoMap = {};

  function getOrCreateWebWorker(jobId) {
    const index = jobId % maxWebWorkers;
    if (webworkers[index]) {
      return webworkers[index];
    }

    // Create a new web worker.
    const webWorker = new Worker("functionssWorker.js");
    webWorker.addEventListener('message', function (event) {
      let jobResult = event.data;
      if (typeof (jobResult) == "string") {
        jobResult = JSON.parse(jobResult);
      }

      if (typeof (jobResult.jobId) == "number") {
        const jobId = jobResult.jobId;
        // Get the promise info associated with the job id.
        const promiseInfo = jobIdToPromiseInfoMap[jobId];
        if (promiseInfo) {
          if (jobResult.error) {
            // The web worker returned an error.
            promiseInfo.reject(new Error());
          }
          else {
            // The web worker returned a result.
            promiseInfo.resolve(jobResult.result);
          }
          delete jobIdToPromiseInfoMap[jobId];
        }
      }
    });

    webworkers[index] = webWorker;
    return webWorker;
  }

  // Post a job to the web worker to do the calculation.
  function dispatchCalculationJob(functionName, parameters) {
    const jobId = nextJobId++;
    return new Promise(function (resolve, reject) {
      // Store the promise information.
      jobIdToPromiseInfoMap[jobId] = { resolve: resolve, reject: reject };
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
 * Dispatches a calculation job for the TEST function.
 * @customfunction
 * @param {number} n - The input number for the TEST function.
 * @returns {Promise} - A promise that resolves with the result of the calculation.
 */
export function TEST(n) {
  return SampleNamespace.dispatchCalculationJob("TEST", [n]);
}

/**
 * Dispatches a calculation job for the TEST_PROMISE function.
 * @customfunction
 * @param {number} n - The input number for the TEST_PROMISE function.
 * @returns {result} - The computing result of the calculation.
 */
export function TEST_PROMISE(n) {
  return SampleNamespace.dispatchCalculationJob("TEST_PROMISE", [n]);
}

/**
 * Dispatches a calculation job for the TEST_ERROR function.
 * @customfunction
 * @param {number} n - The input number for the TEST_ERROR function.
 * @returns {Promise} - A promise that resolves the computing result of the calculation.
 */
export function TEST_ERROR(n) {
  return SampleNamespace.dispatchCalculationJob("TEST_ERROR", [n]);
}

/**
 * Dispatches a calculation job for the TEST_ERROR_PROMISE function.
 * @customfunction
 * @param {number} n - The input number for the TEST_ERROR_PROMISE function.
 * @returns {Promise} - A promise that rejects with an error.
 */
export function TEST_ERROR_PROMISE(n) {
  return SampleNamespace.dispatchCalculationJob("TEST_ERROR_PROMISE", [n]);
}

// This function will show what happens when calculations are run on the main UI thread.
// The task pane will be blocked until this method completes.
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
