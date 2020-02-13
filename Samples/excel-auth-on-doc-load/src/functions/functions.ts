import { cfAction } from '../../utilities/office-apis-helpers';

/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
/* global clearInterval, console, setInterval */

export function add(first: number, second: number): number {
  const myState = localStorage.getItem('loggedIn');
  if (myState !== 'yes') {
    cfAction();
    // @ts-ignore
    throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable);
  } else {
    return first + second;
  };
}

/**
 * Returns data for a given category.
 * @customfunction
 * @param category The category to filter the data with
* @returns {string[][]} A dynamic array with multiple results.
*/
export function getData(category: string): string[][] {
  console.log(category);
  return [
    ['apples', '1', 'pounds'],
    ['oranges', '3', 'pounds'],
    ['pears', '5', 'crates']
  ];
}
