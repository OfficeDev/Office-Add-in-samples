// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// The initialize function must be run each time a new page is loaded
Office.initialize = () => {
  document.getElementById('sideload-msg').style.display='none';
  document.getElementById('app-body').style.display='flex';
  document.getElementById('run').onclick = run;
};

async function run() {
  try {
    await Excel.run(async context => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}
