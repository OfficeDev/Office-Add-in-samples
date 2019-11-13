import { AppState } from '../src/components/App';
import { AxiosResponse } from 'axios';

/*
     Interacting with the Office document
*/

export const writeFileNamesToWorksheet = async (result: AxiosResponse,
                                          displayError: (x: string) => void) => {

        return Excel.run( (context: Excel.RequestContext) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();

            const data = [
                 [result.data.value[0].name],
                 [result.data.value[1].name],
                 [result.data.value[2].name]];

            const range = sheet.getRange('B5:B7');
            range.values = data;
            range.format.autofitColumns();

            return context.sync();
        })
        .catch( (error) => {
            displayError(error.toString());
        });
};

/*
    Managing the dialogs.
*/

let loginDialog: Office.Dialog;
const dialogLoginUrl: string = location.protocol + '//' + location.hostname + (location.port ? ':' + location.port : '') + '/login/login.html';

export const signInO365 = async (setState: (x: AppState) => void,
                                 setToken: (x: string) => void,
                                 displayError: (x: string) => void) => {

    setState({ authStatus: 'loginInProcess' });

    await Office.context.ui.displayDialogAsync(
            dialogLoginUrl,
            {height: 40, width: 30},
            (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    displayError(`${result.error.code} ${result.error.message}`);
                }
                else {
                    loginDialog = result.value;
                    loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, processLoginMessage);
                    loginDialog.addEventHandler(Office.EventType.DialogEventReceived, processLoginDialogEvent);
                }
            }
        );

    const processLoginMessage = (arg: {message: string, type: string}) => {

        let messageFromDialog = JSON.parse(arg.message);
        if (messageFromDialog.status === 'success') { 

            // We now have a valid access token.
            loginDialog.close();
            setToken(messageFromDialog.result);
            setState( { authStatus: 'loggedIn',
                        headerMessage: 'Get Data' });
        }
        else {
            // Something went wrong with authentication or the authorization of the web application.
            loginDialog.close();
            displayError(messageFromDialog.result);
        }
    };

    const processLoginDialogEvent = (arg) => {
        processDialogEvent(arg, setState, displayError);
    };
};

let logoutDialog: Office.Dialog;
const dialogLogoutUrl: string = location.protocol + '//' + location.hostname + (location.port ? ':' + location.port : '') + '/logout/logout.html';

export const logoutFromO365 = async (setState: (x: AppState) => void,
                                     displayError: (x: string) => void) => {

    Office.context.ui.displayDialogAsync(dialogLogoutUrl,
            {height: 40, width: 30},
            (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    displayError(`${result.error.code} ${result.error.message}`);
                }
                else {
                    logoutDialog = result.value;
                    logoutDialog.addEventHandler(Office.EventType.DialogMessageReceived, processLogoutMessage);
                    logoutDialog.addEventHandler(Office.EventType.DialogEventReceived, processLogoutDialogEvent);
                }
            }
        );

    const processLogoutMessage = () => {
        logoutDialog.close();
        setState({ authStatus: 'notLoggedIn',
                   headerMessage: 'Welcome' });
    };

    const processLogoutDialogEvent = (arg) => {
        processDialogEvent(arg, setState, displayError);
    };
};

const processDialogEvent = (arg: {error: number, type: string},
                            setState: (x: AppState) => void,
                            displayError: (x: string) => void) => {

    switch (arg.error) {
        case 12002:
            displayError('The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.');
            break;
        case 12003:
            displayError('The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.');
            break;
        case 12006:
            // 12006 means that the user closed the dialog instead of waiting for it to close.
            // It is not known if the user completed the login or logout, so assume the user is
            // logged out and revert to the app's starting state. It does no harm for a user to
            // press the login button again even if the user is logged in.
            setState({ authStatus: 'notLoggedIn',
                       headerMessage: 'Welcome' });
            break;
        default:
            displayError('Unknown error in dialog box.');
            break;
    }
};

