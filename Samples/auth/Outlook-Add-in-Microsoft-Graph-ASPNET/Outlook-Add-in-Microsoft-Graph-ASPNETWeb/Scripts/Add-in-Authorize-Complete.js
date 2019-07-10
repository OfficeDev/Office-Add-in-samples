// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/* 
The pop-up dialog uses this method to tell the task pane that the user is signed in and web application has an access token.
*/

Office.initialize = function (reason) {
      $(document).ready(function () {  
              console.log("Sending auth complete message through dialog: " + oauthResult.authStatus);  
              Office.context.ui.messageParent(oauthResult.authStatus);  
          });  
}
