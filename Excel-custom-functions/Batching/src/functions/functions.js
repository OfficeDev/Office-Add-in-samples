// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Adds two numbers without using batching
 * @CustomFunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function addNoBatch(first, second) {
  return first + second;
}

/**
 * Divides two numbers using batching
 * @CustomFunction
 * @param dividend The number being divided
 * @param divisor The number to divide the dividend with
 * @returns The sum of the two numbers
 */
function div2(dividend, divisor) {
  console.log("div2 called");
  return _pushOperation(
    "div2",
    [dividend, divisor]
  );
}

/**
 * Multiplies two numbers together using batching
 * @CustomFunction
 * @param first First number to multiply
 * @param second Second number to multiply
 * @returns The product of the two numbers
 */
function mul2(first, second) {
  return _pushOperation(
    "mul2",
    [first, second]
  );
}

/**
 * Defines the implementation of the custom functions
 * for the function id defined in the metadata file (functions.json).
 */
CustomFunctions.associate("ADDNOBATCH", addNoBatch);
CustomFunctions.associate("DIV2", div2);
CustomFunctions.associate("MUL2", mul2);

///////////////////////////////////////

let _batch = [];
let _isBatchedRequestScheduled = false;

// This function encloses your custom functions as individual entries,
// which have some additional properties so you can keep track of whether or not
// a request has been resolved or rejected.
function _pushOperation(op, args) {
  // Create an entry for your custom function.
  console.log("pushOperation");
  const invocationEntry = {
    operation: op, // e.g. sum
    args: args,
    resolve: undefined,
    reject: undefined,
  };

  // Create a unique promise for this invocation,
  // and save its resolve and reject functions into the invocation entry.
  const promise = new Promise((resolve, reject) => {
    invocationEntry.resolve = resolve;
    invocationEntry.reject = reject;
  });

  // Push the invocation entry into the next batch.
  _batch.push(invocationEntry);

  // If a remote request hasn't been scheduled yet,
  // schedule it after a certain timeout, e.g. 100 ms.
  if (!_isBatchedRequestScheduled) {
    console.log("schedule remote request");
    _isBatchedRequestScheduled = true;
    setTimeout(_makeRemoteRequest, 100);
  }

  // Return the promise for this invocation.
  return promise;
}

// This is a private helper function, used only within your custom function add-in.
// You wouldn't call _makeRemoteRequest in Excel, for example.
// This function makes a request for remote processing of the whole batch,
// and matches the response batch to the request batch.
function _makeRemoteRequest() {
  // Copy the shared batch and allow the building of a new batch while you are waiting for a response.
  // Note the use of "splice" rather than "slice", which will modify the original _batch array
  // to empty it out.
  try{
  console.log("makeRemoteRequest");
  const batchCopy = _batch.splice(0, _batch.length);
  _isBatchedRequestScheduled = false;

  // Build a simpler request batch that only contains the arguments for each invocation.
  const requestBatch = batchCopy.map((item) => {
    return { operation: item.operation, args: item.args };
  });
  console.log("makeRemoteRequest2");
  // Make the remote request.
  _fetchFromRemoteService(requestBatch)
    .then((responseBatch) => {
      console.log("responseBatch in fetchFromRemoteService");
      // Match each value from the response batch to its corresponding invocation entry from the request batch,
      // and resolve the invocation promise with its corresponding response value.
      responseBatch.forEach((response, index) => {
        if (response.error) {
          batchCopy[index].reject(new Error(response.error));
          console.log("rejecting promise");
        } else {
          console.log("fulfilling promise");
          console.log(response);

          batchCopy[index].resolve(response.result);
        }
      });
    });
    console.log("makeRemoteRequest3");
  } catch (error) {
    console.log("error name:" + error.name);
    console.log("error message:" + error.message);
    console.log(error);
  }
}

// --------------------- A public API ------------------------------

// This function simulates the work of a remote service. Because each service
// differs, you will need to modify this function appropriately to work with the service you are using. 
// This function takes a batch of argument sets and returns a [promise of] batch of values.
// NOTE: When implementing this function on a server, also apply an appropriate authentication mechanism
//       to ensure only the correct callers can access it.
async function _fetchFromRemoteService(requestBatch) {
  // Simulate a slow network request to the server;
  console.log("_fetchFromRemoteService");
  await pause(1000);
  console.log("postpause");
  return requestBatch.map((request) => {
    console.log("requestBatch server side");
    const { operation, args } = request;

    try {
      if (operation === "div2") {
        // Divide the first argument by the second argument.
        return {
          result: args[0] / args[1]
        };
      } else if (operation === "mul2") {
        // Multiply the arguments for the given entry.
        const myResult = args[0] * args[1];
        console.log(myResult);
        return {
          result: myResult
        };
      } else {
        return {
          error: `Operation not supported: ${operation}`
        };
      }
    } catch (error) {
      return {
        error: `Operation failed: ${operation}`
      };
    }
  });
}

function pause(ms) {
  console.log("pause");
  return new Promise((resolve) => setTimeout(resolve, ms));
}
