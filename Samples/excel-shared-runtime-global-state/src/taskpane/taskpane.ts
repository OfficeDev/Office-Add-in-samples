/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { getValueForKey, setValueForKey } from "./helpers";
import { getGlobal } from "../commands/commands";

 /* global document,  Office  */

Office.initialize = () => {
  // Initialize global state.
  let g = getGlobal() as any;
  let keys: any[] = [];
  let values: any[] = [];

  // state object is used to track key/value pairs, and which storage type is in use
  g.state = {
    keys: keys,
    values: values,
    storageType: "globalvar"
  } as any;

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
  const keyElement = document.getElementById("txtKey") as HTMLInputElement;
  const valueElement = document.getElementById("txtValue") as HTMLInputElement;
  setValueForKey(keyElement.value, valueElement.value);
}

/***
 * Handles the Get button press and calls helper method to retrieve the value from storage for the given key.
 */
function btnGetValue() {
  const keyElement = document.getElementById("txtKey") as HTMLInputElement;
  (document.getElementById("txtValue") as HTMLInputElement).value = getValueForKey(keyElement.value);
}

/***
 * Handles when the radio buttons are selected for local storage or global variable storage.
 * Updates a global variable that tracks which storage type is in use.
 */
function btnStorageChanged() {
  let g = getGlobal() as any;
  
  if ((document.getElementById("globalvar") as HTMLInputElement).checked) {
    g.state.storageType = "globalvar";
  } else {
    g.state.storageType = "localstorage";
  }
}
