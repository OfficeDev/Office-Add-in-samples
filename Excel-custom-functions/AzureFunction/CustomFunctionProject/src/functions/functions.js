/**
 * Add two numbers
 * @customfunction 
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  //If you publish the Azure Function online, update the following URL to use the correct URL location.
  let url = "http://localhost:7071/api/AddTwo";
 
  return new Promise(function(resolve,reject){
  
    //Note that POST uses text/plain because custom functions runtime does not support full CORS
    fetch(url, {
      method: 'POST',
      headers:{
        'Content-Type': 'text/plain'
      },
      body: JSON.stringify({"first": first ,"second": second})
    })
      .then(function (response){
        return response.text();
        }
      )
      .then(function (textanswer) {
       resolve(textanswer);
      })
      .catch(function (error) {
        console.log('error', error.message);
      });
    });  
}

/**
 * Displays the current time once a second
 * @customfunction 
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
function clock(invocation) {
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
function currentTime() {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction 
 * @param {number} incrementBy Amount to increment
 * @param {CustomFunctions.StreamingInvocation<number>} invocation
 */
function increment(incrementBy, invocation) {
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
function logMessage(message) {
  console.log(message);

  return message;
}
