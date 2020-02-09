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
