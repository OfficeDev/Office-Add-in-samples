// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file shows how to open a dialog and process any results sent back to the task pane.

const myMSALObj = new msal.PublicClientApplication(msalConfig);

let homeAccountId = null; // The home account ID of the user that signs in.

/**
 * Gets an access token to the REST API server by using MSAL (Microsoft Authentication Library for the browser.)
 * @returns A promise which if successful, returns the access token.
 */
async function getAccessTokenMSAL() {
    // Attempt to acquire token silently if user is already signed in.
    if (homeAccountId !== null) {
        const result = await myMSALObj.acquireTokenSilent(loginRequest);
        if (result !== null && result.accessToken !== null) {
            return result.accessToken;
        } else return null;
    } else {
        // Create a promise to wrap the dialog callback we need to process later in this function.
        let promise = await new Promise((resolve, reject) => {
            const url = '/dialog.html';
            var fullUrl =
                location.protocol +
                '//' +
                location.hostname +
                (location.port ? ':' + location.port : '') +
                url;

            // height and width are percentages of the size of the parent Office application, e.g., Outlook, PowerPoint, Excel, Word, etc.
            Office.context.ui.displayDialogAsync(
                fullUrl,
                { height: 60, width: 30 },
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                        console.log(
                            (result.error.code = ': ' + result.error.message)
                        );
                        reject(result.error.message);
                    } else {
                        console.log('Dialog has initialized. Wiring up events');
                        let loginDialog = result.value;

                        // Handler for the dialog box closing unexpectedly.
                        loginDialog.addEventHandler(
                            Office.EventType.DialogEventReceived,
                            (arg) => {
                                console.log(
                                    'DialogEventReceived: ' + arg.error
                                );
                                loginDialog.close();
                                // For more dialog codes, see https://learn.microsoft.com/office/dev/add-ins/develop/dialog-handle-errors-events#errors-and-events-in-the-dialog-box
                                switch (arg.error) {
                                    case 12002:
                                        reject('The auth dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.');
                                        break;
                                    case 12003:
                                        reject('The auth dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.');
                                        break;
                                    case 12006:
                                        reject('The auth dialog box was closed before the user signed in.');
                                        break;
                                    default:
                                        reject('Unknown error in auth dialog box.');
                                        break;
                                }
                            }
                        );
                        loginDialog.addEventHandler(
                            Office.EventType.DialogMessageReceived,
                            function processMessage2(arg) {
                                console.log(
                                    'Message received in processMessage'
                                );
                                let messageFromDialog = JSON.parse(arg.message);

                                if (messageFromDialog.status === 'success') {
                                    // We now have a valid access token.
                                    loginDialog.close();
                                    homeAccountId = messageFromDialog.accountId;

                                    // Set the active account so future token requests can be silent.
                                    myMSALObj.setActiveAccount(
                                        myMSALObj.getAccountByHomeId(
                                            homeAccountId
                                        )
                                    );

                                    // Return the token.
                                    resolve(messageFromDialog.result);
                                } else {
                                    // Something went wrong with authentication or the authorization of the web application.
                                    loginDialog.close();
                                    reject(messageFromDialog.error);
                                }
                            }
                        );
                    }
                }
            );
        });
        return promise;
    }
}