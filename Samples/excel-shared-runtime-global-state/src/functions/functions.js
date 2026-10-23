/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { getValueForKey, setValueForKey } from "../shared/state";

/**
 * Get value for key
 * @customfunction
 * @param {string} key The key
 * @returns {string} The value for the key.
 */
function getValueForKeyCF(key) {
  return getValueForKey(key);
}

/**
 * Set value for key
 * @customfunction
 * @param {string} key The key
 * @param {string} value The value to store
 * @returns {string} Confirmation message
 */
function setValueForKeyCF(key, value) {
  setValueForKey(key, value);
  return "Stored key/value pair";
}

CustomFunctions.associate("GETVALUEFORKEYCF", getValueForKeyCF);
CustomFunctions.associate("SETVALUEFORKEYCF", setValueForKeyCF);
