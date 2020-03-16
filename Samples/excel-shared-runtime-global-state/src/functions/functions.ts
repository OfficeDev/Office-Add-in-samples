/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

 import { getValueForKey, setValueForKey } from "../taskpane/helpers";

/**
 * Get value for key
 * @customfunction
 * @param key The key
 * @returns The value for the key.
 */
export function getValueForKeyCF(key: string): string {
  return getValueForKey(key);
}

/**
 * Get value for key
 * @customfunction
 * @param key The key
 * @returns The value for the key.
 */
export function setValueForKeyCF(key: string, value: string): string {
  setValueForKey(key, value);
  return "Stored key/value pair";
}
