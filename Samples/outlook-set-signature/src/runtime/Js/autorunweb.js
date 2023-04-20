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
  let signatureDetails = get_signature_info(template_name, user_info);
  if (Office.context.mailbox.item.itemType == "appointment") {
    set_body(signatureDetails, eventObj);
  } else {
    addTemplateSignature(signatureDetails, eventObj);
  }
}

/**
 * For Outlook on the web, set signature for current appointment
* @param {*} signatureDetails object containing:
 *  "signature": The signature HTML of the template,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 * @param {*} eventObj Office event object
 */
function set_body(signatureDetails, eventObj) {

  if (is_valid_data(signatureDetails.logoBase64) === true) {
    //If a base64 image was passed we need to attach it.
    Office.context.mailbox.item.addFileAttachmentFromBase64Async(
      signatureDetails.logoBase64,
      signatureDetails.logoFileName,
      {
        isInline: true,
      },
      function (result) { 
        Office.context.mailbox.item.body.setAsync(
        "<br/><br/>" + signatureDetails.signature,
        {
          coercionType: "html",
          asyncContext: eventObj,
        },
        function (asyncResult) {

          asyncResult.asyncContext.completed();
        }
      );
    });
  } else {
    Office.context.mailbox.item.body.setAsync(
      "<br/><br/>" + signatureDetails.signature,
      {
        coercionType: "html",
        asyncContext: eventObj,
      },
      function (asyncResult) {

        asyncResult.asyncContext.completed();
      }
    );
  }
  
}
