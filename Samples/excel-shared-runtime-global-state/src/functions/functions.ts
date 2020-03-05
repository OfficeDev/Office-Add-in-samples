import { getGlobal } from "../commands/commands";

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

/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */

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
