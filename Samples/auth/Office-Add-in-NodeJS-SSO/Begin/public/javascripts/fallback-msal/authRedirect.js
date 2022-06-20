// Create the main myMSALObj instance
// configuration parameters are located at authConfig.js

let username = "";
const myMSALObj = new msal.PublicClientApplication(msalConfig);

Office.initialize = function () {
  if (Office.context.ui.messageParent) {
    debugger;
    /**
     * A promise handler needs to be registered for handling the
     * response returned from redirect flow. For more information, visit:
     *
     */
    myMSALObj
      .handleRedirectPromise()
      .then(handleResponse)
      .catch((error) => {
        console.error(error);
      });
  }
};

function handleResponse(response) {
    /**
     * To see the full list of response object properties, visit:
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/request-response-object.md#response
     */
  
    if (response !== null) {
      username = response.account.username;
      Office.context.ui.messageParent( JSON.stringify({ status: 'success', result : response.accessToken }) );
      //welcomeUser(username);
      //updateTable();
    } else {
      //log in
      myMSALObj.loginRedirect(loginRequest);
    }
  }
  