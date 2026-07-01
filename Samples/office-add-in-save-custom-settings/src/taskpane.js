// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* global console, document, Office, fabric */

/** StorageManager tracks which methods to call when
 * managing settings. When the user changes which storage
 * type to use, this object is updated to point to the corresponding
 * save and get methods. On start, the default is to use the property bag.
 */
const StorageManager = {
  mode: "PropertyBag",
  setSetting: saveToPropertyBag,
  getSetting: getFromPropertyBag,
};

Office.onReady(() => {
  // Initialize button that saves a new setting.
  document.getElementById("saveSetting").onclick = () => {
    const name = document.getElementById("setName").value;
    const value = document.getElementById("setValue").value;
    try {
      StorageManager.setSetting(name, value);
      displayStatusMessage("Saved setting KEY: " + name + " VALUE: " + value);
    } catch (err) {
      displayStatusMessage("Error saving: " + err.message);
    }
  };

  // Initialize button that gets a setting value.
  document.getElementById("getSetting").onclick = () => {
    const name = document.getElementById("getName").value;
    const value = StorageManager.getSetting(name);
    displayStatusMessage("Retrieved setting information KEY: " + name + " VALUE: " + value);
  };

  // Configure event handler for when storage options are changed.
  document.getElementById("storageOptions").onchange = setStorageMode;
});

/**
 * Sets the storage mode to match what the user chose in the drop down of options.
 * StorageManager contains two function pointers to set or get values.
 * The function pointers are updated to match the user selection.
 * For example, if the user chooses browser cookies, the methods will
 * point to saveToBrowserCookies, and getFromBrowserCookies.
 */
function setStorageMode() {
  // Get the selected option from the drop-down list.
  const selectionList = document.getElementById("storageOptions");
  const index = selectionList.selectedIndex;
  const modeSelected = selectionList.options[index];
  const mode = modeSelected.value;
  StorageManager.mode = mode;

  try {
    switch (mode) {
      case "PropertyBag":
        // Use the app for Office property bag to store and retrieve data.
        StorageManager.setSetting = saveToPropertyBag;
        StorageManager.getSetting = getFromPropertyBag;
        break;

      case "Cookies":
        // Use browser cookies to store and retrieve data.
        if (navigator.cookieEnabled == true) {
          StorageManager.setSetting = saveToBrowserCookies;
          StorageManager.getSetting = getFromBrowserCookies;
        } else {
          const browserError = { name: "Error", message: "Browser cookies are disabled. You may want to enable them." };
          throw browserError;
        }
        break;

      case "LocalStorage":
        // Use Web Storage to store and retrieve data - storage won't expire.
        if (typeof Storage !== "undefined") {
          StorageManager.setSetting = saveToLocalStorage;
          StorageManager.getSetting = getFromLocalStorage;
        } else {
          const webStorageError = { name: "Error", message: "Browser storage not available in your browser (sorry)." };
          throw webStorageError;
        }
        break;

      case "SessionStorage":
        // Use Web Storage to store and retrieve data, limited to the lifetime of the browser window.
        if (typeof Storage !== "undefined") {
          StorageManager.setSetting = saveToSessionStorage;
          StorageManager.getSetting = getFromSessionStorage;
        } else {
          const webStorageError = { name: "Error", message: "Browser storage not available in your browser (sorry)." };
          throw webStorageError;
        }
        break;

      case "Document":
        // Use a programmatically created, hidden <div> to store and retrieve data.
        StorageManager.setSetting = saveToDocument;
        StorageManager.getSetting = getFromDocument;
        break;
    }
    displayStatusMessage("Switched storage type to " + mode);
  } catch (err) {
    displayStatusMessage(err.name + ":" + err.message);
  }
}

// Stores the settings in the JavaScript APIs for Office property bag.
async function saveToPropertyBag(key, value) {
  // Note that Project does not support the settings object.
  // Need to check that the settings object is available before setting.
  if (Office.context.document.settings) {
    Office.context.document.settings.set(key, value);
    return new Promise((resolve, reject) => {
      Office.context.document.settings.saveAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
          const error = new Error('Settings save failed. Error: ' + asyncResult.error.message);
          displayStatusMessage(error.message);
          reject(error);
        } else {
          resolve();
        }
      });
    });
  } else {
    const unsupportedError = {
      name: "Error: Feature not supported",
      message: "The settings object is not supported in this host application.",
    };
    displayStatusMessage(unsupportedError.name + ": " + unsupportedError.message);
    throw unsupportedError;
  }
}

// Retrieves the specified setting value from the JavaScript APIs for Office  property bag using the specified key.
function getFromPropertyBag(key) {
  // Note that Project does not support the settings object.
  // Need to check that the settings object is available before setting.
  if (Office.context.document.settings) {
    const value = Office.context.document.settings.get(key);
    return value;
  } else {
    const unsupportedError = {
      name: "Error: Feature not supported",
      message: "The settings object is not supported in this host application.",
    };
    throw unsupportedError;
  }
}

// Stores the settings as a browser cookie.
function saveToBrowserCookies(key, value) {
  document.cookie = key + "=" + value + "; path=/; SameSite=None; Secure";
  
  // Verify the cookie was saved
  const savedValue = getFromBrowserCookies(key);
  if (savedValue === undefined) {
    throw new Error("Cookie was blocked. Browser cookies may not work in this context due to third-party cookie restrictions.");
  }
}

// Retrieves the specified setting from the browser cookies.
function getFromBrowserCookies(key) {
  const all = document.cookie;
  let value;

  if (all === "") {
    return undefined;
  } else {
    var list = all.split("; ");
    for (var i = 0; i < list.length; i++) {
      var cookie = list[i];
      var p = cookie.indexOf("=");
      var name = cookie.substring(0, p);

      if (name == key) {
        value = cookie.substring(p + 1);
        break;
      }
    }
  }
  return value;
}

// Stores the settings using local storage (Web Storage that doesn't expire).
// See http://msdn.microsoft.com/en-us/library/ie/cc197062(v=vs.85).aspx information about localStorage, sessionStorage.
function saveToLocalStorage(_key, _value) {
  localStorage.setItem(_key, _value);
}

// Retrieves the specified setting from local storage (Web Storage that doesn't expire).
function getFromLocalStorage(_key) {
  const value = localStorage.getItem(_key);
  return value;
}

// Stores the settings using session storage (Web Storage limited to the lifetime of the browser window).
function saveToSessionStorage(_key, _value) {
  sessionStorage.setItem(_key, _value);
}

// Retrieves the specified setting from session storage (Web Storage limited to the lifetime of the browser window).
function getFromSessionStorage(_key) {
  const value = sessionStorage.getItem(_key);
  return value;
}

// Stores the settings in a hidden <div> added to the document.
function saveToDocument(key, value) {
  let hiddenStorage = null;
  const hiddenName = "hiddenstorage";

  if (document.getElementById(hiddenName) == null) {
    hiddenStorage = document.createElement("div");
    hiddenStorage.setAttribute("id", hiddenName);
    hiddenStorage.setAttribute("style", "display:none;");

    document.body.appendChild(hiddenStorage);
  } else {
    hiddenStorage = document.getElementById(hiddenName);
  }

  var keyNode = document.createElement("span");
  keyNode.setAttribute("id", key);

  var valueNode = document.createTextNode(value);
  keyNode.appendChild(valueNode);

  hiddenStorage.appendChild(keyNode);
}

// Retrieves the specified setting from a hidden <div> in the document.
function getFromDocument(key) {
  let value;

  if (document.getElementById(key) != null) {
    const valueNode = document.getElementById(key);
    value = valueNode.innerHTML;
  }

  return value;
}

function displayStatusMessage(message) {
  document.getElementById("bannerText").innerText = message;
}
