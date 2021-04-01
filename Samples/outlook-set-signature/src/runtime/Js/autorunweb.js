// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file contains code only used by autorunweb.html when loaded in Outlook on the web.

Office.initialize = function (reason) {};



/**
 * Set signature for current appointment
 * @param {*} signature_str Signature to insert
 * @param {*} eventObj Office event object
 */
function set_body(signature_str, eventObj) {
  Office.context.mailbox.item.body.setAsync(
    "<br/><br/>" + signature_str,
    {
      coercionType: "html",
      asyncContext: eventObj,
    },
    function (asyncResult) {
      asyncResult.asyncContext.completed();
    }
  );
}
