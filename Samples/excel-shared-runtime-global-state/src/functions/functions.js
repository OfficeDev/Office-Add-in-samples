/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/**
 * Get value for key
 * @customfunction
 * @param key The key
 * @returns The value for the key.
 */
function getValueForKeyCF(key) {
  return getValueForKey(key);
}

/**
 * Get value for key
 * @customfunction
 * @param key The key
 * @returns The value for the key.
 */
function setValueForKeyCF(key, value) {
  setValueForKey(key, value);
  return "Stored key/value pair";
}

CustomFunctions.associate("GETVALUEFORKEYCF", getValueForKeyCF);
CustomFunctions.associate("SETVALUEFORKEYCF",setValueForKeyCF);
