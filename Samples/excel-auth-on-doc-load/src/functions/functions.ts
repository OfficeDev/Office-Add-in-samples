import { cfAction } from '../../utilities/office-apis-helpers';

/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
/* global clearInterval, console, setInterval */

export async function add(first: number, second: number): Promise <number> {
  await cfAction();
    return first + second;
    }
