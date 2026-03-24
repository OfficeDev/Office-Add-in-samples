/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/**
 * Gets the value of a cell, evaluated synchronously with Excel's calculation.
 * @customfunction
 * @supportSync
 * @param {string} address The cell address, e.g. "A1"
 * @param {CustomFunctions.Invocation} invocation
 * @returns {Promise<any>} The cell value.
 */
export async function getCellValue(address: string, invocation: CustomFunctions.Invocation): Promise<any> {
  const context = new Excel.RequestContext();
  context.setInvocation(invocation); // The `invocation` object must be passed in the `setInvocation` method for synchronous functions.

  const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
  range.load("values");

  await context.sync();
  return range.values[0][0];
}
