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
 * Handle the OnNewMessageCompose or OnNewAppointmentOrganizerevent by calling getCalendarFreeBusy which will
 * append the user's free/busy schedule information to the message body.
 * 
 * @param {Office.AddinCommands.Event} event The OnNewMessageCompose or OnNewAppointmentOrganizer event object.
 */
async function onItemComposeHandler(event) {
  await getCalendarFreeBusy();
  event.completed({ allowEvent: true });
}

Office.actions.associate("onMessageComposeHandler", onItemComposeHandler);

/**
 * Call the web server API to get the user's free/busy schedule. 
 * The web server will use OBO and call Microsoft Graph to get and return the schedule.
 */
async function getCalendarFreeBusy() {
  try {
      // Get access token from Outlook host via SSO.
      let accessToken = await Office.auth.getAccessToken(defaultSSO);

      // Call web server which will make Graph call and return filename list.
      let response = await callWebServer(
          'GET',
          '/getCalendarFreeBusy',
          accessToken
      );

      if (response === null) {
          // callWebServer returns null when it handles any errors and will display messages for the user.
          // Just return in this scenario.
          return;
      }

      if (response.claims) {
          // Microsoft Graph requires an additional form of authentication. Have the Office host
          // get a new token using the Claims string, which tells AAD to prompt the user for all
          // required forms of authentication.
          let mfaMiddletierToken = await Office.auth.getAccessToken({
              authChallenge: response.claims,
          });
          response = await callWebServer(
              'GET',
              '/getCalendarFreeBusy',
              mfaMiddletierToken
          );
      }

      // AAD errors are returned to the client with HTTP code 200, so they do not trigger
      // the catch block.
      if (response.error) {
          handleAADErrors(response, callback);
          return;
      }
      
      // Write file names into the mail item body.
      const resultText = await response.text();
      await appendTextOnSend(resultText);

  } catch (exception) {
      // if handleClientSideErrors returns true then we will try to authenticate via the fallback
      // dialog rather than simply throw and error
      if (exception.code) {
          handleClientSideErrors(exception);
          return null;
      } else {
          showMessage('EXCEPTION: ' + JSON.stringify(exception));
          throw exception;
      }
  }
 
}

/**
 * Calls the REST API on the server.
 * @param {*} verb HTTP verb to use such as GET, POST, etc...
 * @param {*} url URL of the REST API.
 * @param {*} accessToken SSO access token from Office (required by the REST API).
 * @returns 
 */
async function callWebServer(verb, url, accessToken) {
  let response;
  try {
      response = await fetch(url, {
          method: verb,
          credentials: 'same-origin', // include, *same-origin, omit
          headers: {
              Authorization: 'Bearer ' + accessToken,
          },
          cache: 'no-cache',
      });
      if (!response.ok) {
          throw Error(response.statusText);
          
        }
      return response;
  } catch (error) {
      // Check for expired SSO token. Refresh and retry the call if it expired.
      if (
          response.responseJSON &&
          response.responseJSON.type === 'TokenExpiredError'
      ) {
          try {
              const refreshAccessToken = await Office.auth.getAccessToken(
                  defaultSSO
              );
              const data = await fetch(url, {
                  method: verb,
                  credentials: 'same-origin', // include, *same-origin, omit
                  headers: {
                      Authorization: 'Bearer ' + refreshAccessToken,
                  },
                  cache: false,
              });

              return data;
          } catch (error) {
              showMessage(response.responseText);
              return null;
          }
      }

      // Check for a Microsoft Graph API call error. which is returned as bad request (403)
      if (response.status === 403) {
          if (
              response.responseJSON &&
              response.responseJSON.type === 'Microsoft Graph'
          ) {
              showMessage(response.responseJSON.errorDetails);
          } else {
              showMessage(error);
          }

          return null;
      }

      // For all other error scenarios, display the message.
      showMessage('Unknown error from web server: ' + JSON.stringify(error));
      return null;
  }
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
          // No one is signed into Office. If the add-in cannot be effectively used when no one
          // is logged into Office, then the first call of getAccessToken should pass the
          // `allowSignInPrompt: true` option. Since this sample does that, you should not see
          // this error.
          showMessage(
              'No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again.'
          );
          break;
      case 13002:
          // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
          // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
          showMessage(
              'You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again.'
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
  const id = "dac64749-cb7308b6d444";
  const details =
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: text
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
      console.log("Action failed with error: " + asyncResult.error.message);
      return;
    }

    const bodyFormat = asyncResult.value;
    Office.context.mailbox.item.body.appendOnSendAsync(text, { coercionType: bodyFormat }, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
        return;
      }

      console.log(
        `"${text}" will be appended to the body once the message or appointment is sent. Send the mail item to test this feature.`
      );
    });
  });
}
