
function add(first, second) {
  return first + second;
}

function increment(incrementBy, callback) {
  var result = 0;
  var timer = setInterval(function () {
    result += incrementBy;
    callback.setResult(result);
  }, 1000);

  callback.onCanceled = function () {
    clearInterval(timer);
  };
}

// Get auth token before calling my service, a hypothetical API that will deliver a stock price based on stock ticker string, such as "MSFT"

function getStock(ticker) {
  console.log("starting");
  return new Promise(function (resolve, reject) {
      getToken2("https://localhost:8081/dialog.html")
      .then(function (token) {
        resolve(token);
      })
  });
}

//Helper
function getToken(url) {
  return new Promise(function (resolve, reject) {
    getTokenViaDialog(url)
      .then(function (result) {
        resolve(result);
      })
      .catch(function (result) {
        reject(result);
      });
  });
}

function getToken2(url) {
  return new Promise(function (resolve, reject) {
    displayDialogTest(url,200,300,false,true)
      .then(function (result) {
        resolve(result);
      })
      .catch(function (result) {
        reject(result);
      });
  });
}


function getTokenViaDialog(url) {
  return new Promise(function (resolve, reject) {
    _dialogOpen = true;
    OfficeRuntime.displayWebDialog(url, {
      height: '50%',
      width: '50%',
      onMessage: function (message, dialog) {
        _cachedToken = message;
        resolve(message);
        dialog.closeDialog();
        return;
      },
      onRuntimeError: function (error, dialog) {
        reject(error);
      }
    }).catch(function (e) {
      reject(e);
    });
  
  });
}

function displayDialogTest(url, height, width, hideTitle, closeDialog) {
  return new Promise(function (resolve) {
        OfficeRuntime.displayWebDialog(url, {
               width: width,
               height: height,
               hideTitle: hideTitle,
               onMessage: function(message, dialog) {
                      if (closeDialog) {
                            dialog.close();
                            resolve(message);
                      } else {
                            resolve(message);
                      }
               },
               onRuntimeError:function(error, dialog) {
                      if (closeDialog) {
                            dialog.close();
                      }
                     resolve(error.message);
               }
         }).catch(function(e) {
               resolve(e.message);
         });
  });
}


CustomFunctions.associate("GETSTOCK", getStock);
CustomFunctions.associate("ADD", add);
CustomFunctions.associate("INCREMENT", increment);
