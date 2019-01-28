//Custom Functions Batching Sample with TODO items

//TODO #1: Create custom functions for adding and multiplying numbers.

function _pushOperation(op, args) {
    var invocationEntry = {
      "operation": op, // e.g. sum
      "arguments": args,
      "resolve": undefined,
      "reject": undefined
    };
  
    // Each invocation has a unique promise, which is mapped to the resolve and reject
    // properties in the invocationEntry object. 
    var promise = new Promise((resolve, reject) => { 
      invocationEntry.resolve = resolve; 
      invocationEntry.reject = reject;
    });
  
  
    _batch.push(invocationEntry);
  
    // TODO #2: If a remote request hasn't been scheduled yet,
    // schedule it after a certain timeout, e.g. 2 sec.
    
    //Return the promise for this invocation.
    return promise;
  }
  
  var _batch = [];
  var _isBatchedRequestScheduled = false;
  
  // TODO #3: Write a function which creates a copy of the batch array, empty the batch array,
  // and create a new array containing only the raw information to be processed by the API. 
  
  // TODO #4: In the same function for TODO #3, make the request to the API and match the values
  // from the returned results to the corresponding invocation entry from the request batch. 
  
  
  // --------------------- A public API ------------------------------
  
  // This function simulates the work of a remote service and is here for your reference only. Because each service
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
  
  
  