// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

function add(first, second){
  return first + second;
}

function StoreValue(key, value) {
  return OfficeRuntime.AsyncStorage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to AsyncStorage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to AsyncStorage. " + error;
  });
}

function GetValue(key) {
  return OfficeRuntime.AsyncStorage.getItem(key);
}

CustomFunctions.associate("ADD", add);
CustomFunctions.associate("STOREVALUE",StoreValue);
CustomFunctions.associate("GETVALUE",GetValue);
