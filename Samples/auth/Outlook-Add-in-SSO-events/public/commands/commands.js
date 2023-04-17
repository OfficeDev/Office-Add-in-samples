/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const urlOrigin = 'https://localhost:3000'; // Change this if deploying to a different location.

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        // All console logs from events go to a runtime logging file. For more information, see https://learn.microsoft.com/office/dev/add-ins/testing/runtime-logging
        console.log('initializing ...' + info);
    }
});

// Default SSO settings for acquiring access tokens.
const defaultSSO = {
    allowSignInPrompt: false,
    allowConsentPrompt: false,
    //    forMSGraphAccess: true, // Leave commented during development testing (sideload) or you get a 13012 error from getAccessToken.
};

/**
 * Handle the OnNewMessageCompose event by calling getUserProfile.
 * This appends a signature with the user's profile to the message body on send.
 *
 * @param {Office.AddinCommands.Event} event The OnNewMessageCompose event object.
 */
function onItemComposeHandler(event) {
    callWebServerAPI('GET', urlOrigin + '/getuserprofile')
        .then((jsonResponse) => {
            let signature = `${jsonResponse.displayName} \n ${jsonResponse.mail}`;
            if (jsonResponse.jobTitle !== null) {
                signature += `\n ${jsonResponse.jobTitle}`;
            }
            if (jsonResponse.mobilePhone !== null) {
                signature += `\n ${jsonResponse.mobilePhone}`;
            }
            appendTextOnSend(signature);
            event.completed();
        })
        .catch((exception) => {
            showMessage(JSON.stringify(exception));
            event.completed();
        });
}

Office.actions.associate('onMessageComposeHandler', onItemComposeHandler);

/**
 * Calls a REST API on the server.
 * @param {*} method HTTP method to use such as GET, POST, etc...
 * @param {*} url URL of the REST API.
 * @param {*} retryRequest Indicates if this is a retry of the call.
 * @returns A promise that will return the JSON response from the REST API.
 */
function callWebServerAPI(method, url, retryRequest = false) {
    let response = null;
    // Get the access token from Office host using SSO.
    // Note that Office.auth.getAccessToken modifies the options parameter. Create a copy of the object
    // to avoid modifying the original object.
    const options = JSON.parse(JSON.stringify(defaultSSO));

    // Begin promise chain.
    return OfficeRuntime.auth
        .getAccessToken(options)
        .then((accessToken) => {
            // Call the REST API on our web server.
            return fetch(url, {
                method: method,
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': 'Bearer ' + accessToken,
                },
            });
        })
        .then((responseData) => {
            // Get the JSON body from the response.
            response = responseData;
            return response.json();
        })
        .then((jsonBody) => {
            // Check for success condition: HTTP status code 2xx.
            if (response.ok) {
                return new Promise((resolve) => {
                    resolve(jsonBody);
                });
            }
            // Check for fail condition: Did we get a Microsoft Graph API error, which is returned as bad request (403)?
            else if (
                response.status === 403 &&
                jsonBody.type === 'Microsoft Graph'
            ) {
                // Return a promise that will reject with the Microsoft Graph error details.
                return new Promise((resolve, reject) => {
                    reject('Microsoft Graph error: ' + jsonBody.errorDetails);
                });
            }
            // Check for expired token. If the token expired, retry the call which will get a refreshed token.
            else if (
                jsonBody !== null &&
                jsonBody.type === 'TokenExpiredError' &&
                !retryRequest
            ) {
                // Try the call again (and return result). The underlying call to Office JS getAccessToken will refresh the token.
                return callWebServerAPI(method, url, true); // true parameter will ensure we only do the recursion once.
            } else {
                // Handle all other errors by returning a promise that will reject with the error details.
                return new Promise((resolve, reject) => {
                    reject('Unknown error: ' + JSON.stringify(jsonBody));
                });
            }
        });
}

/**
 * Appends text to the end of the message or appointment's body once it's sent.
 * @param {*} text The text to append.
 */
function appendTextOnSend(text) {
    // It's recommended to call getTypeAsync and pass its returned value to the options.coercionType parameter of the appendOnSendAsync call.
    Office.context.mailbox.item.body.getTypeAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(
                'Action failed with error: ' + asyncResult.error.message
            );
            return;
        }

        const bodyFormat = asyncResult.value;
        Office.context.mailbox.item.body.appendOnSendAsync(
            text,
            { coercionType: bodyFormat },
            (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.log(
                        'Action failed with error: ' + asyncResult.error.message
                    );
                    return;
                }

                showMessage(
                    `"${text}" will be appended to the body once the message or appointment is sent. Send the mail item to test this feature.`
                );
            }
        );
    });
}

/**
 * Creates information bar to display a message to the user.
 */
function showMessage(text) {
    console.log(text);
    const id = 'dac64749-cb7308b6d444'; // Unique ID for the notification.
    const details = {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: text.substring(0, 150),
    };
    Office.context.mailbox.item.notificationMessages.addAsync(id, details);
}

/**
 * Handles any error returned from getAccessToken. The text shown for each error indicates next steps
 * you should take.
 * @param {*} err The error to process.
 */
function handleSSOErrors(err) {
    switch (err.code) {
        case 13001:
            // No one is signed into Office. If the add-in can't be effectively used when no one
            // is logged into Office, then the first call of getAccessToken should pass the
            // `allowSignInPrompt: true` option. Since this sample does that, you should not see
            // this error.
            showMessage(
                'No one is signed into Office. Please sign in before sending.'
            );
            break;
        case 13002:
            // The user aborted the consent prompt. If the add-in can't be effectively used when consent
            // hasn't been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
            showMessage(
                'You have not granted consent. If you want to grant consent, try sending again.'
            );
            break;
        case 13006:
            // Only seen in Office on the web.
            showMessage(
                'Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again.'
            );
            break;
        case 13008:
            // Only seen in Office on the web.
            showMessage(
                'Office is still working on the last operation. When it completes, try this operation again.'
            );
            break;
        case 13010:
            // Only seen in Office on the web.
            showMessage(
                "Follow the instructions to change your browser's zone configuration."
            );
            break;
        default:
            // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, show error message.
            showMessage('Could not sign in: ' + err.code);
            break;
    }
}
