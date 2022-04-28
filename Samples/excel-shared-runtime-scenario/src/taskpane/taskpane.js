// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    ensureStateInitialized(true);
    console.log("ensure state initialized from the office.initialize");
    isOfficeInitialized = true;
    monitorSheetChanges();

    // Associate custom functions
    CustomFunctions.associate("ADD", add);
    CustomFunctions.associate("GETDATA", getData);

    // Associate commands
    Office.actions.associate("btnConnectService", btnConnectService);
    Office.actions.associate("btnDisconnectService", btnDisconnectService);
    Office.actions.associate("btnOpenTaskpane", btnOpenTaskpane);
    Office.actions.associate("btnCloseTaskpane", btnCloseTaskpane);
    Office.actions.associate("btnEnableAddinStart", btnEnableAddinStart);
    Office.actions.associate("btnDisableAddinStart", btnDisableAddinStart);
    Office.actions.associate("btnInsertData", btnInsertData);
    Office.actions.associate("btnSumData", btnSumData);

    document.getElementById("connectService").onclick = connectService; // in office-apis-helpers.js
    document.getElementById("selectFilter").onclick = insertFilteredData;
    
    updateRibbon();
    updateTaskPaneUI();
  }
});

async function insertFilteredData() {
  try {
    //Determine which data source the user selected from the radio buttons.
    const radioExcel = document.getElementById("communicationFilter");
    if (radioExcel.checked) {
      generateCustomFunction("Communications");
    } else {
      generateCustomFunction("Groceries");
    }
  } catch (error) {
    console.error(error);
  }
}
