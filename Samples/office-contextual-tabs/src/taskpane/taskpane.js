/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
// import "../../assets/icon-16.png";
// import "../../assets/icon-32.png";
// import "../../assets/icon-80.png";
// import { getGlobal } from '../commands/commands.js';
// import { createSampleWorkSheet } from '../utilities/utilities.js';


/* global console, document, Office */

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("import").onclick = importData;

    //Create the contextual tab
    let g = getGlobal(); 
    await Office.ribbon.requestCreateControls(g.contextualTab);
  }
});

/**
 * Handles when Import data button is selected. Checks which radio button was selected (Excel or SQL)
 * and then creates the sample worksheet and sales table based on the user's choice.
 */
async function importData() {
  try {
    //Determine which data source the user selected from the radio buttons.
    let g = getGlobal(); 
    const radioExcel = document.getElementById('excelFile');
    if (radioExcel.checked) {
      g.mockDataSource = 'excelFileMockData';
    } else{
      g.mockDataSource = 'sqlMockData';
    }

    //Create the sample worksheet and sales table.
    createSampleWorkSheet(g.mockDataSource);
  } catch (error) {
    console.error(error);
  }
}

