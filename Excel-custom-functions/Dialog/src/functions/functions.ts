/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
CustomFunctions.associate("ADD", add);

/**
 * Displays the current time once a second.
 * @customfunction 
 * @param invocation Custom function handler  
 */
function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
CustomFunctions.associate("CLOCK", clock);

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
function currentTime(): string {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction 
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler 
 */
function increment(incrementBy: number, invocation: CustomFunctions.StreamingInvocation<number>): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
CustomFunctions.associate("INCREMENT", increment);

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
function logMessage(message: string): string {
  console.log(message);

  return message;
}
CustomFunctions.associate("LOG", logMessage);

/**
 * Writes a message to console.log().
 * @customfunction
 * @param ticker String stock quote name to retrieve.
 * @returns Number Stock quote value.
 */
function getStock(ticker:string) {
  return new Promise(function (resolve, reject) {
      getToken("https://localhost:3000/dialog.html")
      .then(function (token) {
        resolve(token);
      })
  });
}
CustomFunctions.associate("GETSTOCK", getStock);

function getToken(url:string) {
  return new Promise(function (resolve, reject) {
    displayDialogTest(url,'50%','50%',false,true)
      .then(function (result) {
        resolve(result);
      })
      .catch(function (result) {
        reject(result);
      });
  });
}

function displayDialogTest(url:string, height:string, width:string, hideTitle:boolean, closeDialog:boolean) {
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

