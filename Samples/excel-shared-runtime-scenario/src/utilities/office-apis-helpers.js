// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const SetRuntimeVisibleHelper = (visible) => {
  let p;
  if (visible) {
    p = Office.addin.showAsTaskpane();
  } else {
    p = Office.addin.hide();
  }

  return p
    .then(() => {
      return visible;
    })
    .catch((error) => {
      return error.code;
    });
};

const SetStartupBehaviorHelper = (isStarting) => {
  if (isStarting) {
    Office.addin.setStartupBehavior(Office.StartupBehavior.load);
  } else {
    Office.addin.setStartupBehavior(Office.StartupBehavior.none);
  }
  let g = getGlobal();
  g.isStartOnDocOpen = isStarting;
};

function updateRibbon() {
  // Update ribbon based on state tracking
  const g = getGlobal();

  Office.ribbon.requestUpdate({
    tabs: [
      {
        id: "ShareTime",
        controls: [
          {
            id: "BtnConnectService",
            enabled: !g.state.isConnected,
          },
          {
            id: "BtnDisConnectService",
            enabled: g.state.isConnected,
          },
          {
            id: "BtnInsertData",
            enabled: g.state.isConnected,
          },
          {
            id: "BtnSumData",
            enabled: g.state.isSumEnabled,
          },
          {
            id: "BtnEnableAddinStart",
            enabled: !g.state.isStartOnDocOpen,
          },
          {
            id: "BtnDisableAddinStart",
            enabled: g.state.isStartOnDocOpen,
          },
          {
            id: "BtnOpenTaskpane",
            enabled: !g.state.isTaskpaneOpen,
          },
          {
            id: "BtnCloseTaskpane",
            enabled: g.state.isTaskpaneOpen,
          },
        ],
      },
    ],
  });
}

/*
    Managing the dialogs.
*/

async function connectService() {
  //construct url for dialog
  const path = location.path.substr(0,location.path.length-13);
  const dialogConnectUrl = location.protocol +
  "//" +
  location.hostname +
  (location.port ? ":" + location.port : "") + path +
  "/src/dialog/connect.html";

  //pop up a dialog
  let connectDialog;

  const processMessage = () => {
    const g = getGlobal();
    g.state.setConnected(true);
    g.state.isConnectInProgress = false;
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

function generateCustomFunction(selectedOption) {
  try {
    Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      let range = context.workbook.getSelectedRange();

      //let selectedOption = 'Communications';

      range.values = [['=CONTOSOSHARE.GETDATA("' + selectedOption + '")']];
      range.format.autofitColumns();
      return context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

//This will check if state is initialized, and if not, initialize it.
//Useful as there are multiple entry points that need the state and it is not clear which one will get called first.
async function ensureStateInitialized(isOfficeInitializing) {
  console.log("ensureInitialize called");
  let g = getGlobal();
  let initValue = false;
  if (isOfficeInitializing) {
    //we are being called in response to Office Initialize
    if (g.state !== undefined) {
      if (g.state.isInitialized === false) {
        g.state.isInitialized = true;
      }
    }
    if (g.state === undefined) {
      initValue = true;
    }
  }

  if (g.state === undefined) {
    g.state = {
      isStartOnDocOpen: false,
      isSignedIn: false,
      isTaskpaneOpen: false,
      isConnected: false,
      isSyncEnabled: false,
      isConnectInProgress: false,
      isFirstSyncCall: true,
      isSumEnabled: false,
      isInitialized: initValue,
      updateRct: () => {},
      setTaskpaneStatus: (opened) => {
        g.state.isTaskpaneOpen = opened;
        updateRibbon();
      },
      setConnected: (connected) => {
        g.state.isConnected = connected;

        if (connected) {
          if (g.state.updateRct !== null) {
            g.state.updateRct("true");
          }
        } else {
          if (g.state.updateRct !== null) {
            g.state.updateRct("false");
          }
        }
        updateRibbon();
      },
    };

    //track startup behavior
    if (g.state.isInitialized) {
      let addinState = await Office.addin.getStartupBehavior();
      console.log("load state is:");
      console.log("load state" + addinState);
      if (addinState === Office.StartupBehavior.load) {
        g.state.isStartOnDocOpen = true;
      }
    }

    //track sign in status
    if (localStorage.getItem("loggedIn") === "yes") {
      g.state.isSignedIn = true;
    }
  }
  if (g.state.isInitialized) {
    updateRibbon();
  }
}

async function onTableSelectionChange(event) {
  let g = getGlobal();
  return Excel.run((context) => {
    return context.sync().then(() => {
      console.log("Table section changed...");
      console.log("Change type of event: " + event.changeType);
      console.log("Address of event: " + event.address);
      console.log("Source of event: " + event.source);
      g.state.selectionAddress = event.address;
      if (event.address === "" && g.state.isSumEnabled === true) {
        g.state.isSumEnabled = false;
        updateRibbon();
      } else if (g.state.isSumEnabled === false && event.address !== "") {
        g.state.isSumEnabled = true;
        updateRibbon();
      }
    });
  });
}

async function monitorSheetChanges() {
  try {
    let g = getGlobal();
    if (g.state === undefined) {
      return;
    }
    if (g.state.isInitialized) {
      await Excel.run(async (context) => {
        let table = context.workbook.tables.getItem("ExpensesTable");
        return context
          .sync()
          .then(async () => {
            table.onSelectionChanged.add(onTableSelectionChange);
            await context.sync();
            updateRibbon();
          })
          .catch(async () => {
            g.state.isSumEnabled = false;
            updateRibbon();
          });
      });
    }
  } catch (error) {
    console.error(error);
  }
}

function updateTaskPaneUI() {
  let g = getGlobal();

  if (g.state.isConnected === true) {
    document.getElementById("filterSection").style.visibility = "visible";
    document.getElementById("connectSection").style.visibility = "hidden";
  } else {
    document.getElementById("filterSection").style.visibility = "hidden";
    document.getElementById("connectSection").style.visibility = "visible";
  }
}
