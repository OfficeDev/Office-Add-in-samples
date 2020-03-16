/* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. */

import { getGlobal } from '../commands/commands';

export function setValueForKey(key: string, value: string) {
    let g = getGlobal() as any;
    if (g.state.storageType === "globalvar") {
      g.state.keys.push(key);
      g.state.values.push(value);
    } else {
      g.window.localStorage.setItem(key, value);
    }
  }
  
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
  