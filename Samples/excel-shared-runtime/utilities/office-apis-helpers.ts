import { AppState } from '../src/components/App';
import { AxiosResponse } from 'axios';

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

/*
     Interacting with the Office document
*/

export const writeFileNamesToWorksheet = async (
  result: AxiosResponse,
  displayError: (x: string) => void
) => {
  return Excel.run((context: Excel.RequestContext) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    const data = [
      [result.data.value[0].name],
      [result.data.value[1].name],
      [result.data.value[2].name]
    ];

    const range = sheet.getRange('B5:B7');
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
  }).catch(error => {
    displayError(error.toString());
  });
};



const processDialogEvent = (
  arg: { error: number; type: string },
  setState: (x: AppState) => void,
  displayError: (x: string) => void
) => {
  switch (arg.error) {
    case 12002:
      displayError(
        'The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.'
      );
      break;
    case 12003:
      displayError(
        'The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.'
      );
      break;
    case 12006:
      // 12006 means that the user closed the dialog instead of waiting for it to close.
      // It is not known if the user completed the login or logout, so assume the user is
      // logged out and revert to the app's starting state. It does no harm for a user to
      // press the login button again even if the user is logged in.
      setState({
        authStatus: 'notLoggedIn',
        headerMessage: 'Welcome'
      });
      break;
    default:
      displayError('Unknown error in dialog box.');
      break;
  }
};

/*
    Managing the dialogs.
*/

let loginDialog: Office.Dialog;
const dialogLoginUrl: string =
  location.protocol +
  '//' +
  location.hostname +
  (location.port ? ':' + location.port : '') +
  '/login/login.html';
const dialogConnectUrl: string =
  location.protocol +
  '//' +
  location.hostname +
  (location.port ? ':' + location.port : '') +
  '/login/connect.html';

export const signInO365 = async (
  setState: (x: AppState) => void,
  setToken: (x: string) => void,
  displayError: (x: string) => void
) => {
  setState({ authStatus: 'loginInProcess' });

  const processLoginMessage = (arg: { message: string; type: string }) => {
    let messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.status === 'success') {
      // We now have a valid access token.
      loginDialog.close();
      setToken(messageFromDialog.result);
      setState({
        authStatus: 'loggedIn',
        headerMessage: 'Get Data'
      });
    } else {
      // Something went wrong with authentication or the authorization of the web application.
      loginDialog.close();
      displayError(messageFromDialog.result);
    }
  };

  const processLoginDialogEvent = arg => {
    processDialogEvent(arg, setState, displayError);
  };
  Office.context.ui.displayDialogAsync(dialogLoginUrl, { height: 40, width: 30 }, result => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      displayError(`${result.error.code} ${result.error.message}`);
    }
    else {
      loginDialog = result.value;
      loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, processLoginMessage);
      loginDialog.addEventHandler(Office.EventType.DialogEventReceived, processLoginDialogEvent);
    }
  });
};

let logoutDialog: Office.Dialog;
const dialogLogoutUrl: string =
  location.protocol +
  '//' +
  location.hostname +
  (location.port ? ':' + location.port : '') +
  '/logout/logout.html';

export const logoutFromO365 = async (
  setState: (x: AppState) => void,
  displayError: (x: string) => void
) => {

  const processLogoutMessage = () => {
    logoutDialog.close();
    setState({
      authStatus: 'notLoggedIn',
      headerMessage: 'Welcome'
    });
  };

  const processLogoutDialogEvent = arg => {
    processDialogEvent(arg, setState, displayError);
  };

  Office.context.ui.displayDialogAsync(
    dialogLogoutUrl,
    { height: 40, width: 30 },
    result => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        displayError(`${result.error.code} ${result.error.message}`);
      } else {
        logoutDialog = result.value;
        logoutDialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          processLogoutMessage
        );
        logoutDialog.addEventHandler(
          Office.EventType.DialogEventReceived,
          processLogoutDialogEvent
        );
      }
    }
  );

};


// sign in commands (without task pane)

export class SignApp {
  appstate: AppState;
  accessToken: string;

  setToken = (accesstoken: string) => {
    this.accessToken = accesstoken;
    localStorage.setItem('mytoken', accesstoken);
    //    g.token = accesstoken;
  }

  setState = (nState: AppState) => {
    this.appstate = nState;
    //localStorage.setItem("loggedIn", "yes");
  }

  displayError = (error: string) => {
    this.setState({ errorMessage: error });
  }
}

export const SetRuntimeVisibleHelper = (visible: boolean) => {
  let p: any;
  if (visible) {
    // @ts-ignore
    p = Office.addin.showAsTaskpane();
  } else {
    // @ts-ignore
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

export async function cfAction(): Promise<void> {
  let signapp = new SignApp();
  return signInO365(
    signapp.setState,
    signapp.setToken,
    signapp.displayError
  ).then();
}

export function updateRibbon() {
  // Update ribbon based on state tracking
  const g = getGlobal() as any;

  // @ts-ignore
  OfficeRuntime.ui
    .getRibbon()
    // @ts-ignore
    .then((ribbon) => {
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

export async function connectService() {
  //pop up a dialog
  let connectDialog: Office.Dialog;
  let g = getGlobal() as any;

  const processMessage = () => {

    g.state.setConnected(true);
    g.state.isConnectInProgress = false;
    connectDialog.close();
  };


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
export async function ensureStateInitialized() {
  let g = getGlobal() as any;
  monitorSheetChanges();

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
  } catch (error) {
    console.error(error);
  }
}
