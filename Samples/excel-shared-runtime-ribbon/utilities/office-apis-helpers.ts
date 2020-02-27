export function getGlobal() {
  console.log('init globals for command buttons');
  return typeof self !== 'undefined'
    ? self
    : typeof window !== 'undefined'
    ? window
    : typeof global !== 'undefined'
    ? global
    : undefined;
}

export const SetRuntimeVisibleHelper = (visible: boolean) => {
  let p: any;
  if (visible) {
    p = Office.addin.showAsTaskpane();
  } else {
    p = Office.addin.hide();
  }

  return p
    .then(() => {
      return visible;
    })
    .catch(error => {
      return error.code;
    });
};

export const SetStartupBehaviorHelper = (isStarting: boolean) => {
  if (isStarting) {
    // @ts-ignore
    Office.addin.setStartupBehavior(Office.StartupBehavior.load);
  } else {
    // @ts-ignore
    Office.addin.setStartupBehavior(Office.StartupBehavior.none);
  }
  let g = getGlobal() as any;
  g.isStartOnDocOpen = isStarting;
};

export function updateRibbon() {
  // Update ribbon based on state tracking
  const g = getGlobal() as any;

  // @ts-ignore
  OfficeRuntime.ui
    .getRibbon()
    // @ts-ignore
    .then(ribbon => {
      ribbon.requestUpdate({
        tabs: [
          {
            id: 'ShareTime',
            // visible: 'true',
            controls: [
              {
                id: 'BtnConnectService',
                enabled: !g.state.isConnected
              },
              {
                id: 'BtnDisConnectService',
                enabled: g.state.isConnected
              },
              {
                id: 'BtnInsertData',
                enabled: g.state.isConnected
              },
              {
                id: 'BtnSyncData',
                enabled: g.state.isSyncEnabled
              },
              {
                id: 'BtnSumData',
                enabled: g.state.isSumEnabled
              },
              {
                id: 'BtnEnableAddinStart',
                enabled: !g.state.isStartOnDocOpen
              },
              {
                id: 'BtnDisableAddinStart',
                enabled: g.state.isStartOnDocOpen
              }
            ]
          }
        ]
      });
    });
}

/*
    Managing the dialogs.
*/

const dialogConnectUrl: string =
  location.protocol +
  '//' +
  location.hostname +
  (location.port ? ':' + location.port : '') +
  '/login/connect.html';

export async function connectService() {
  //pop up a dialog
  let connectDialog: Office.Dialog;

  const processMessage = () => {
    g.state.setConnected(true);
    g.state.isConnectInProgress = false;
    connectDialog.close();
  };

  let g = getGlobal() as any;
  await Office.context.ui.displayDialogAsync(
    dialogConnectUrl,
    { height: 40, width: 30 },
    result => {
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

export function generateCustomFunction(selectedOption: string) {
  try {
    Excel.run(async context => {
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
export async function ensureStateInitialized(isOfficeInitializing: boolean) {
  let g = getGlobal() as any;
  monitorSheetChanges();
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
      setTaskpaneStatus: (opened: boolean) => {
        g.state.isTaskpaneOpen = opened;
        updateRibbon();
      },
      setConnected: (connected: boolean) => {
        g.state.isConnected = connected;

        if (connected) {
          if (g.state.updateRct !== null) {
            g.state.updateRct('true');
          }
        } else {
          if (g.state.updateRct !== null) {
            g.state.updateRct('false');
          }
        }
        updateRibbon();
      }
    };

    // @ts-ignore
    let addinState = await Office.addin._getState();
    console.log('load state is:');
    console.log('load state' + addinState);
    if (addinState === 'Background') {
      g.state.isStartOnDocOpen = true;
      //run();
    }
    if (localStorage.getItem('loggedIn') === 'yes') {
      g.state.isSignedIn = true;
    }
  }
  updateRibbon();
}

async function onTableChange(event) {
  return Excel.run(context => {
    return context.sync().then(() => {
      console.log('Change type of event: ' + event.changeType);
      console.log('Address of event: ' + event.address);
      console.log('Source of event: ' + event.source);
      let g = getGlobal() as any;
      if (g.state.isConnected) {
        g.state.isSyncEnabled = true;
        updateRibbon();
      }
    });
  });
}

async function onTableSelectionChange(event) {
  let g = getGlobal() as any;
  return Excel.run(context => {
    return context.sync().then(() => {
      console.log('Table section changed...');
      console.log('Change type of event: ' + event.changeType);
      console.log('Address of event: ' + event.address);
      console.log('Source of event: ' + event.source);
      g.state.selectionAddress = event.address;
      if (event.address === '' && g.state.isSumEnabled === true) {
        g.state.isSumEnabled = false;
        updateRibbon();
      } else if (g.state.isSumEnabled === false && event.address !== '') {
        g.state.isSumEnabled = true;
        updateRibbon();
      }
    });
  });
}

export async function monitorSheetChanges() {
  try {
    let g = getGlobal() as any;
    if (g.state.isInitialized) {
      await Excel.run(async context => {
        let table = context.workbook.tables.getItem('ExpensesTable');
        if (table !== undefined) {
          table.onChanged.add(onTableChange);
          table.onSelectionChanged.add(onTableSelectionChange);
          await context.sync();
          updateRibbon();
          console.log('A handler has been registered for the onChanged event.');
        } else {
          g.state.isSumEnabled = false;
          updateRibbon();
          console.log('Expense table not present to add handler to.');
        }
      });
    }
  } catch (error) {
    console.error(error);
  }
}
