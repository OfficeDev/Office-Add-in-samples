/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */

/* global Office */

// Create a "utility class" for the storage functions.
var StorageLibrary = (function () {
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
      var value = valueNode.innerHTML;
    }

    return value;
  }

  // 'Expose' the public members.
  return {
    saveToPropertyBag: saveToPropertyBag,
    getFromPropertyBag: getFromPropertyBag,
    saveToBrowserCookies: saveToBrowserCookies,
    getFromBrowserCookies: getFromBrowserCookies,
    saveToLocalStorage: saveToLocalStorage,
    getFromLocalStorage: getFromLocalStorage,
    saveToSessionStorage: saveToSessionStorage,
    getFromSessionStorage: getFromSessionStorage,
    saveToDocument: saveToDocument,
    getFromDocument: getFromDocument,
  };
})();
