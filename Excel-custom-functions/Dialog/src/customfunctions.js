
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
      getToken("https://localhost:8081/dialog.html")
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


CustomFunctions.associate("GETSTOCK", getStock);
CustomFunctions.associate("ADD", add);
CustomFunctions.associate("INCREMENT", increment);
