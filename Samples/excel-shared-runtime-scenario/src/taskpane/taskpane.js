/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, CustomFunctions, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    ensureStateInitialized(true);
    console.log("ensure state initialized from the office.initialize");
    isOfficeInitialized = true;
    monitorSheetChanges();

    
    CustomFunctions.associate("ADD", add);
    CustomFunctions.associate("GETDATA", getData);
    
    document.getElementById("connectService").onclick = connectService;
    document.getElementById("selectFilter").onclick = insertFilteredData;
    
    updateRibbon();
    updateTaskPaneUI();
  }
});

async function connectService() {
  //pop up a dialog
  let connectDialog;

  const processMessage = () => {
    const g = getGlobal();
    g.state.setConnected(true);
    g.state.isConnectInProgress = false;

    updateTaskPaneUI();

    connectDialog.close();
  };

  let g = getGlobal();
  await Office.context.ui.displayDialogAsync(
    dialogConnectUrl,
    { height: 40, width: 30, promptBeforeOpen: false },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log(`${result.error.code} ${result.error.message}`);
        g.state.setConnected(false);
      } else {
        connectDialog = result.value;
        connectDialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          processMessage
        );
      }
    }
  );
}

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
