// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file contains code only used by autorunweb.html when loaded in Outlook on the web.

Office.initialize = function (reason) {};

/**
 * For Outlook on the web, insert signature into appointment or message.
 * Outlook on the web does not support using setSignatureAsync on appointments,
 * so this method will update the body directly.
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @param {*} user_info Information details about the user
 * @param {*} eventObj Office event object
 */
function insert_auto_signature(compose_type, user_info, eventObj) {
  let template_name = get_template_name(compose_type);
  let signature_str = get_signature_str(template_name, user_info);
  if (Office.context.mailbox.item.itemType == "appointment") {
    set_body(signature_str, eventObj);
  } else {
    set_signature(signature_str, eventObj);
  }
}

/**
 * For Outlook on the seb, set signature for current appointment
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
