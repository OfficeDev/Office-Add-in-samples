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

/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
