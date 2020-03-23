/* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. */

import { getGlobal } from '../commands/commands';

/***
 * Stores the key/value pair. Will use local storage or global variable to store
 * the values depending on which type the user selected.
 * 
 * @export
 * @param {string} key The key to store.
 * @param {string} value The value to store.
 */
export function setValueForKey(key: string, value: string): void {
    let g = getGlobal() as any;
    if (g.state.storageType === "globalvar") {
      g.state.keys.push(key);
      g.state.values.push(value);
    } else {
      g.window.localStorage.setItem(key, value);
    }
  }
  

  /**
   * Gets the value for the given key from storage. Will retrieve the value
   * from local storage or global variable depending on which type of storage
   * the user selected.
   *
   * @export
   * @param {string} key The key to retrieve the value for
   * @returns {string} The value
   */
  export function getValueForKey(key: string): string {
    let g = getGlobal() as any;
    let answer = "";
    if (g.state.storageType === "globalvar") {
      g.state.keys.forEach((element, index) => {
        if (element === key) {
          answer = g.state.values[index];
        }
      });
    } else {
      answer = g.window.localStorage.getItem(key);
    }
    return answer;
  }
  