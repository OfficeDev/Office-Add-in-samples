// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

function add(first, second){
  return first + second;
}

function increment(incrementBy, callback) {
  var result = 0;
  var timer = setInterval(function() {
    result += incrementBy;
    callback.setResult(result);
  }, 1000);

  callback.onCanceled = function() {
    clearInterval(timer);
  };
}

// Custom functions for adding and multiplying numbers.
function sum() {
  return _pushOperation(
    "sum",
    // The last argument is an InvocationContext. Skip it.
    Array.from(arguments).slice(0, -1));
}

function mul() {
  return _pushOperation(
    "mul",
    // The last argument is an InvocationContext. Skip it.
    Array.from(arguments).slice(0, -1));
}

CustomFunctions.associate("ADD", add);
CustomFunctions.associate("INCREMENT", increment);
CustomFunctions.associate("SUM", sum);
CustomFunctions.associate("MUL", mul);


// This function encloses your custom functions as individual entries, 
// which have some additional properties so you can keep track of whether or not
// a request has been resolved or rejected.  
function _pushOperation(op, args) {
  // Create an entry for your custom function.
  var invocationEntry = {
    "operation": op, // e.g. sum
    "arguments": args,
    "resolve": undefined,
    "reject": undefined
  };

  // Create a unique promise for this invocation, 
  // and save its resolve and reject functions into the invocation entry.
  var promise = new Promise((resolve, reject) => { 
    invocationEntry.resolve = resolve; 
    invocationEntry.reject = reject;
  });

  // Push the invocation entry into the next batch.
  _batch.push(invocationEntry);

  // If a remote request hasn't been scheduled yet,
  // schedule it after a certain timeout, e.g. 2 sec.
  if (!_isBatchedRequestScheduled) {
    setTimeout(_makeRemoteRequest, 2000);
    _isBatchedRequestScheduled = true;
  }

  //Return the promise for this invocation.
  return promise;
}


// Next batch
var _batch = [];
var _isBatchedRequestScheduled = false;


// This is a private function, used only within your custom function add-in. 
// You wouldn't call _makeRemoteRequest in Excel, for example. 
// This function makes a request for remote processing of the whole batch,
// and matches the response batch to the request batch.
function _makeRemoteRequest() {
  // Copy the shared batch and allow the building of a new batch while you are waiting for a response.
  var batchCopy = _batch.slice();
  _batch = [];
  _isBatchedRequestScheduled = false;

  // Build a simpler request batch that only contains the arguments for each invocation.
  var requestBatch = [];
  for (var i = 0; i < batchCopy.length; i++) {
    requestBatch[i] = {
      "operation": batchCopy[i].operation,
      "arguments": batchCopy[i].arguments
    };
  }

  // Make the remote request.
  _fetchFromRemoteService(requestBatch)
    .then(function (responseBatch) {
    // Match each value from the response batch to its corresponding invocation entry from the request batch,
    // and resolve the invocation promise with its corresponding response value.
    for (var i = 0; i < responseBatch.length; i++) {
      batchCopy[i].resolve(responseBatch[i]);
    }
  });
}


// --------------------- A public API ------------------------------

// This function simulates the work of a remote service. Because each service
// differs, you will need to modify this function appropriately to work with the service you are using. 
// This function takes a batch of argument sets and returns a [promise of] batch of values.
function _fetchFromRemoteService(requestBatch) {
  var responseBatch = [];
  for (var i = 0; i < requestBatch.length; i++) {
    var operation = requestBatch[i].operation;
    var args = requestBatch[i].arguments;
    var result;

    if (operation == "sum") {
      // Sum up the arguments for the given entry.
      result = 0;
      for (var j = 0; j < args.length; j++) {
        result += args[j];
      }
    }
    else if (operation == "mul") {
      // Multiply the arguments for the given entry.
      result = 1;
      for (var j = 0; j < args.length; j++) {
        result *= args[j];
      }
    }

    // Set the result on the responseBatch.
    responseBatch[i] = result;
  }

  // Return a promise that is resolved with the value of the response batch.
  return Promise.resolve(responseBatch);
}

