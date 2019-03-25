# Custom function batching pattern #

### Summary ###
If your custom functions call a remote service you may want to use a batching pattern to reduce the number of network calls to the remote service. This is useful when a spreadsheet recalculates and it contains many of your custom functions. Recalculate will result in many calls to your custom functions, but you can batch them into one or a few calls to the remote service. 

### Applies to ###
-  Custom functions on Excel desktop and online

### Prerequisites ###
Custom functions are currently in developer preview. To get set up and using custom functions, see [Custom functions requirements](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-requirements)

### Solution ###
Solution | Author(s)
---------|----------
Custom function batching | Microsoft

### Version history ###
Version  | Date | Comments
---------| -----| --------
1.0  | March 26th 2019 | Initial release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**


----------

# Scenario: Custom function batching #
In this scenario your custom functions call a remote service. To reduce network round trips you will batch all the calls and send them in a single call to the web service. This is ideal when the spreadsheet is recalculated. For example, if someone used your custom function in 100 cells in a spreadsheet, and then recalculates the spreadsheet, your custom function would run 100 times and make 100 network calls. By using this batching pattern, the calls can be combined to make all 100 calculations in a single network call.

Image

## Custom functions that use the pattern
The code pattern contains two custom functions named `ADD2` and `MUL2`. Instead of performing the calculation, each of them calls a `_pushOperation` function to push the operation into a batch queue to be passed to a web service.

```js
function addtwo() {
  return _pushOperation(
    "add2",
    // The last argument is an InvocationContext. Skip it.
    Array.from(arguments).slice(0, -1));
}
```

## Batching the operation
The `_pushOperation` function pushes each operation into a _batch variable. It schedules the batch call to be made within 2 seconds. You can adjust this when using the code in your own solution.

```js
if (!_isBatchedRequestScheduled) {
    setTimeout(_makeRemoteRequest, 2000);
    _isBatchedRequestScheduled = true;
  }
```

## Making the remote request
The `_makeRemoteRequest` function prepares the batch request and passes it to the `_fetchFromRemoteService` function. If you are adapting this code to your own solution you need to modify `_makeRemoteRequest` to actually call your remote service.

## The remote service
The `_fetchFromRemoteService` function processes the batch of operations, performs the operations, and then returns the results. In this sample, `_fetchFromRemoteService` is just another function to demonstrate the pattern. When adapting this code to your solution, use this method on the server-side to respond to the client call over the network.

## How to apply batching in your own solution
You can copy and paste this code into your own solution. When using this pattern, you'll need to evaluate and update the following areas of code.

### _pushOperation
Adjust the timeout value as needed. A longer time will be more noticable to the user. A shorter time may result in more calls to the remote service.

### _makeRemoteRequest
Modify this function to actually make a network call to your remote service and pass the batch operations in a single call.

### _fetchFromRemoteService
Place this function in your remote service to handle the network call from the client. You'll want to modify this to perform the actual operations of your custom functions (or call the correct methods to do so.)

