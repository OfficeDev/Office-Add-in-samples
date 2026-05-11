/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Office */

let isTaskPaneInitialized = false;

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    // Guard against re-initialization when ShowTaskpane re-shows the pane on web.
    if (isTaskPaneInitialized) {
      return;
    }
    isTaskPaneInitialized = true;

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("import").onclick = importData;

    // Create the contextual tab (only succeeds on first call per session)
    let g = getGlobal(); 
    try {
      await Office.ribbon.requestCreateControls(g.contextualTab);
    } catch (e) {
      // requestCreateControls was already called earlier in this session.
    }
  }
});

/**
 * Handles when Import data button is selected. Checks which radio button was selected (Excel or SQL)
 * and then creates the sample worksheet and sales table based on the user's choice.
 */
 async function importData() {
  try {
    // Determine which data source the user selected from the radio buttons.
    let g = getGlobal(); 
    const radioExcel = document.getElementById('excelFile');
    if (radioExcel.checked) {
      g.mockDataSource = 'excelFileMockData';
    } else{
      g.mockDataSource = 'sqlMockData';
    }

    // Create the sample worksheet and sales table.
    await createSampleWorkSheet();
    await createSampleTable(g.mockDataSource);
  } catch (error) {
    console.error(error);
  }
}