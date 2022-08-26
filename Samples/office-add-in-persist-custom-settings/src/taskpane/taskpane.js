/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Office, fabric */

// StorageManager tracks which methods to call when
// managing settings. When the user changes which storage
// type to use, this object is updated to point to the corresponding
// save and get methods.
const StorageManager = {
  // Defaults to property bag to manage settings
  mode: "PropertyBag",
  setSetting: saveToPropertyBag,
  getSetting: getFromPropertyBag,
};

Office.onReady(() => {
  document.getElementById("storageOptions").onchange = setStorageMode;

  // Initialize text fields.
  const TextFieldElements = document.querySelectorAll(".ms-TextField");
  for (let i = 0; i < TextFieldElements.length; i++) {
    new fabric["TextField"](TextFieldElements[i]);
  }

  // Initialize buttons.
  let button = document.getElementById("saveSetting");
  new fabric["Button"](button, () => {
    const name = document.getElementById("setName").value;
    const value = document.getElementById("setValue").value;
    console.log(name);
    console.log(value);
    StorageManager.setSetting(name, value);
  });

  button = document.getElementById("getSetting");
  new fabric["Button"](button, () => {
    const name = document.getElementById("getName").value;
    var v = StorageManager.getSetting(name);
    showToast("Retrieved setting information KEY: " + name + " VALUE: " + v);
    console.log(v);
  });

  // Initialize dropdown.
  var DropdownHTMLElements = document.querySelectorAll(".ms-Dropdown");
  for (var i = 0; i < DropdownHTMLElements.length; ++i) {
    let Dropdown = new fabric["Dropdown"](DropdownHTMLElements[i]);
    console.log(Dropdown);
  }
});

function setStorageMode() {
  // Get the selected option fromt the drop-down list.
  var selectionList = document.getElementById("storageOptions");
  var index = selectionList.selectedIndex;
  var modeSelected = selectionList.options[index];
  var mode = modeSelected.value;
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
          var browserError = { name: "Error", message: "Browser cookies are disabled. You may want to enable them." };
          throw browserError;
        }
        break;

      case "LocalStorage":
        // Use Web Storage to store and retrieve data - storage won't expire.
        if (typeof Storage !== "undefined") {
          StorageManager.setSetting = saveToLocalStorage;
          StorageManager.getSetting = getFromLocalStorage;
        } else {
          var webStorageError = { name: "Error", message: "Browser storage not available in your browser (sorry)." };
          throw webStorageError;
        }
        break;

      case "SessionStorage":
        // Use Web Storage to store and retrieve data, limited to the lifetime of the browser window.
        if (typeof Storage !== "undefined") {
          StorageManager.setSetting = saveToSessionStorage;
          StorageManager.getSetting = getFromSessionStorage;
        } else {
          var webStorageError2 = { name: "Error", message: "Browser storage not available in your browser (sorry)." };
          throw webStorageError2;
        }
        break;

      case "Document":
        // Use a programmatically created, hidden <div> to store and retrieve data.
        StorageManager.setSetting = saveToDocument;
        StorageManager.getSetting = getFromDocument;
        break;
    }
    showToast("Switched storage type to " + mode);
  } catch (err) {
    showToast(err.name + ":" + err.message);
  }
}
// Stores the settings in the JavaScript APIs for Office property bag.
async function saveToPropertyBag(key, value) {
  // Note that Project does not support the settings object.
  // Need to check that the settings object is available before setting.
  if (Office.context.document.settings) {
    Office.context.document.settings.set(key, value);
    await Office.context.document.settings.saveAsync();
  } else {
    var unsupportedError = {
      name: "Error: Feature not supported",
      message: "The settings object is not supported in this host application.",
    };
    throw unsupportedError;
  }
}

// Retrieves the specified setting value from the JavaScript APIs for Office  property bag using the specified key.
function getFromPropertyBag(key) {
  // Note that Project does not support the settings object.
  // Need to check that the settings object is available before setting.
  if (Office.context.document.settings) {
    var value = null;
    value = Office.context.document.settings.get(key);
    return value;
  } else {
    var unsupportedError = {
      name: "Error: Feature not supported",
      message: "The settings object is not supported in this host application.",
    };
    throw unsupportedError;
  }
}

// Stores the settings as a browser cookie.
function saveToBrowserCookies(key, value) {
  document.cookie = key + "=" + value;
}

// Retrieves the specified setting from the browser cookies.
function getFromBrowserCookies(key) {
  var cookies = {};
  var all = document.cookie;
  var value = null;

  if (all === "") {
    return cookies;
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
  var value = localStorage.getItem(_key);
  return value;
}

// Stores the settings using session storage (Web Storage limited to the lifetime of the browser window).
function saveToSessionStorage(_key, _value) {
  sessionStorage.setItem(_key, _value);
}

// Retrieves the specified setting from session storage (Web Storage limited to the lifetime of the browser window).
function getFromSessionStorage(_key) {
  var value = sessionStorage.getItem(_key);
  return value;
}

// Stores the settings in a hidden <div> added to the document.
function saveToDocument(key, value) {
  var hiddenStorage = null;
  var hiddenName = "hiddenstorage";

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
  var value = null;

  if (document.getElementById(key) != null) {
    var valueNode = document.getElementById(key);
    value = valueNode.innerHTML;
  }

  return value;
}

function showToast(message) {
  document.getElementById("bannerText").innerText = message;
}
