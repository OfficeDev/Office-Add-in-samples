// This file copied and modified from https://github.com/Azure-Samples/ms-identity-javascript-tutorial/blob/main/1-Authentication/1-sign-in/App/authRedirect.js

// Create the main myMSALObj instance
// configuration parameters are located at authConfig.js

const myMSALObj = new msal.PublicClientApplication(msalConfig);

Office.initialize = async function () {
  if (Office.context.ui.messageParent) {
    try {
      const response = await myMSALObj.handleRedirectPromise();
      handleResponse(response);
    } catch (error) {
      console.error(error);
      Office.context.ui.messageParent(
        JSON.stringify({ status: "error", error: error.message }),
        { targetOrigin: window.location.origin }
      );
    }
  }
};

function handleResponse(response) {
  /**
   * To see the full list of response object properties, visit:
   * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/request-response-object.md#response
   */

  if (response !== null) {
    Office.context.ui.messageParent(
      JSON.stringify({ status: "success", result: response.accessToken, accountId: response.account.homeAccountId }),
      { targetOrigin: window.location.origin }
    );
  } else {
    //log in
    myMSALObj.loginRedirect(loginRequest);
  }
}
