/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady();

// Default SSO settings for acquiring access tokens.
const defaultSSO = {
    allowSignInPrompt: true,
    allowConsentPrompt: true,
    forMSGraphAccess: true,
};

/**
 * Handle the OnNewMessageCompose or OnNewAppointmentOrganizer event by calling getUserProfile.
 * This appends a signature with the user's profile to the message body on send.
 *
 * @param {Office.AddinCommands.Event} event The OnNewMessageCompose or OnNewAppointmentOrganizer event object.
 */
async function onItemComposeHandler(event) {
    await getUserProfile();
    event.completed({ allowEvent: true });
}

Office.actions.associate('onMessageComposeHandler', onItemComposeHandler);

/**
 * Call the web server API to get the user's free/busy schedule.
 * The web server will use OBO and call Microsoft Graph to get and return the schedule.
 */
async function getUserProfile() {
    try {
        // Get access token from Outlook host via SSO.
        let accessToken = await OfficeRuntime.auth.getAccessToken(defaultSSO);

        // Call web server which will make Graph call and return filename list.
        let jsonReponse = await callWebServerAPI(
            'GET',
            '/getUserProfile',
            accessToken
        );

        // Create signature from user profile.
        const signature = `${jsonReponse.displayName} \n ${jsonReponse.mail} \n ${jsonReponse.jobTitle} \n ${jsonReponse.mobilePhone}`;
        await appendTextOnSend(signature);
    } catch (exception) {
        // Exceptions are displayed in the notification bar.
        if (exception.code) {
            handleSSOErrors(exception);
            return;
        } else {
            showMessage(exception.message);
            return;
        }
    }
}

/**
 * Calls a REST API on the server.
 * @param {*} method HTTP method to use such as GET, POST, etc...
 * @param {*} url URL of the REST API.
 * @param {*} retryRequest Indicates if this is a retry of the call.
 * @returns The JSON response from the REST API.
 */
async function callWebServerAPI(method, url, retryRequest = false) {
    // Get the access token from Office host using SSO.
    // Note that Office.auth.getAccessToken modifies the options parameter. Create a copy of the object
    // to avoid modifying the original object.
    const options = JSON.parse(JSON.stringify(defaultSSO));
    const accessToken = await Office.auth.getAccessToken(options);

    const response = await fetch(url, {
        method: method,
        headers: {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + accessToken,
        },
    });

    // Check for success condition: HTTP status code 2xx.
    if (response.ok) {
        return response.json();
    }

    // Check for fail condition: Is the SSO token expired? If so, retry the call which will get a refreshed token.
    const jsonBody = await response.json();
    if (
        jsonBody !== null &&
        jsonBody.type === 'TokenExpiredError' &&
        !retryRequest
    ) {
        return callWebServerAPI(method, path, true); // Try the call again. The underlying call to Office JS getAccessToken will refresh the token.
    }

    // Check for fail condition: Did we get a Microsoft Graph API error, which is returned as bad request (403)?
    if (response.status === 403 && jsonBody.type === 'Microsoft Graph') {
        throw new Error('Microsoft Graph error: ' + jsonBody.errorDetails);
    }

    // Handle other errors.
    throw new Error(JSON.stringify(jsonBody));
}

/**
 * Handles any error returned from getAccessToken. The numbered errors are typically user actions
 * that don't require fallback auth. The text shown for each error indicates next steps
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

/**
 * Creates information bar to display a message to the user.
 */
function showMessage(text) {
    console.log(text);
    const id = 'dac64749-cb7308b6d444';
    const details = {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: text.substring(0,150),
    };
    Office.context.mailbox.item.notificationMessages.addAsync(id, details);
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
