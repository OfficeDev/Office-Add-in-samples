/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

export function ensureState() {
  // Creates a shared state bag once per runtime so task pane and custom functions read/write the same structure.
  if (!globalThis.g) {
    globalThis.g = {};
  }

  if (!globalThis.g.state) {
    globalThis.g.state = {
      keys: [],
      values: [],
      storageType: "globalvar",
    };
  }

  return globalThis.g.state;
}

export function setValueForKey(key, value) {
  // Writes to in-memory arrays by default; falls back to localStorage only when that mode is selected and available.
  const state = ensureState();

  if (state.storageType === "globalvar") {
    state.keys.push(key);
    state.values.push(value);
  } else if (typeof window !== "undefined" && window.localStorage) {
    window.localStorage.setItem(key, value);
  }
}

export function getValueForKey(key) {
  // Reads from the active storage mode and returns an empty string when a key is missing.
  const state = ensureState();

  if (state.storageType === "globalvar") {
    const index = state.keys.indexOf(key);
    return index >= 0 ? state.values[index] : "";
  }

  if (typeof window !== "undefined" && window.localStorage) {
    return window.localStorage.getItem(key) || "";
  }

  return "";
}
