// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
 function add(first, second) {
  //If you publish the Azure Function online, update the following URL to use the correct URL location.
  const url = "http://localhost:7071/api/AddTwo";

  return new Promise(async function (resolve, reject) {
    try {
      //Note that POST uses text/plain because custom functions runtime does not support full CORS
      const response = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "text/plain",
        },
        body: JSON.stringify({ first: first, second: second }),
      });
      const jsonAnswer = await response.json();
      resolve(jsonAnswer.answer);
    } catch (error) {
      console.log("error", error.message);
    }
  });
}

CustomFunctions.associate("ADDTWO", add);
