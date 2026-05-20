/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { ensureState, getValueForKey, setValueForKey } from "../shared/state";

Office.initialize = () => {
  ensureState();

  // In shared runtime, this script can execute even when task pane UI is not active.
  const sideloadMsg = document.getElementById("sideload-msg");
  const appBody = document.getElementById("app-body");
  const btnStore = document.getElementById("btnStoreValue");
  const btnGet = document.getElementById("btnGetValue");
  const globalVarRadio = document.getElementById("globalvar");
  const localStorageRadio = document.getElementById("localstorage");

  if (sideloadMsg) {
    sideloadMsg.style.display = "none";
  }

  if (appBody) {
    appBody.style.display = "flex";
  }

  if (btnStore) {
    btnStore.onclick = btnStoreValue;
  }

  if (btnGet) {
    btnGet.onclick = btnGetValue;
  }

  if (globalVarRadio) {
    globalVarRadio.onclick = btnStorageChanged;
  }

  if (localStorageRadio) {
    localStorageRadio.onclick = btnStorageChanged;
  }
};

/***
 * Handles the Store button press event and calls helper method to store the key/value pair from the user in storage.
 */
function btnStoreValue() {
  const keyElement = document.getElementById("txtKey");
  const valueElement = document.getElementById("txtValue");

  if (!keyElement || !valueElement) {
    return;
  }

  setValueForKey(keyElement.value, valueElement.value);
}

/***
 * Handles the Get button press and calls helper method to retrieve the value from storage for the given key.
 */
function btnGetValue() {
  const keyElement = document.getElementById("txtKey");
  const valueElement = document.getElementById("txtValue");

  if (!keyElement || !valueElement) {
    return;
  }

  valueElement.value = getValueForKey(keyElement.value);
}

/***
 * Handles when the radio buttons are selected for local storage or global variable storage.
 * Updates a global variable that tracks which storage type is in use.
 */
function btnStorageChanged() {
  const state = ensureState();
  const globalVarRadio = document.getElementById("globalvar");

  if (globalVarRadio && globalVarRadio.checked) {
    state.storageType = "globalvar";
  } else {
    state.storageType = "localstorage";
  }
}
