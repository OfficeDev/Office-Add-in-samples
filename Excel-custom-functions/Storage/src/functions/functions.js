// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

/**
 * @customfunction
 * @description Adds two numbers together. 
 * @param {number} first First number to be added.
 * @param {number} second Second number to be added.
 */
function add(first, second){
  return first + second;
}

/**
 * @customfunction
 * @description Stores a value in Office.storage.
 * @param {any} key Key in the key-value pair you will store. 
 * @param {any} value Value in the key-value pair you will store. 
 */
function StoreValue(key, value) {
  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}

/**
 * @customfunction
 * @description Gets value from Office.storage. 
 * @param {any} key Key of item you intend to get.
 */
function GetValue(key) {
  return OfficeRuntime.storage.getItem(key);
}

CustomFunctions.associate("ADD", add);
CustomFunctions.associate("STOREVALUE",StoreValue);
CustomFunctions.associate("GETVALUE",GetValue);
