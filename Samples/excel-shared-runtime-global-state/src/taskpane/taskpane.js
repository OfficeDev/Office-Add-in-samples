/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.initialize = () => {
  // Initialize global state.
  let g = getGlobal();
  let keys = [];
  let values = [];

  // state object is used to track key/value pairs, and which storage type is in use
  g.state = {
    keys: keys,
    values: values,
    storageType: "globalvar"
  };

  // Connect handlers
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("btnStoreValue").onclick = btnStoreValue;
  document.getElementById("btnGetValue").onclick = btnGetValue;
  document.getElementById("globalvar").onclick = btnStorageChanged;
  document.getElementById("localstorage").onclick = btnStorageChanged;
};

/***
 * Handles the Store button press event and calls helper method to store the key/value pair from the user in storage.
 */
function btnStoreValue() {
  const keyElement = document.getElementById("txtKey");
  const valueElement = document.getElementById("txtValue");
  setValueForKey(keyElement.value, valueElement.value);
}

/***
 * Handles the Get button press and calls helper method to retrieve the value from storage for the given key.
 */
function btnGetValue() {
  const keyElement = document.getElementById("txtKey");
  (document.getElementById("txtValue")).value = getValueForKey(keyElement.value);
}

/***
 * Handles when the radio buttons are selected for local storage or global variable storage.
 * Updates a global variable that tracks which storage type is in use.
 */
function btnStorageChanged() {
  let g = getGlobal();
  
  if ((document.getElementById("globalvar")).checked) {
    g.state.storageType = "globalvar";
  } else {
    g.state.storageType = "localstorage";
  }
}
