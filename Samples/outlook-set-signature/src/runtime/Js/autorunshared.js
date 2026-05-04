// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Contains code for event-based activation on Outlook on web, on Windows, and on Mac (new UI preview).

var CIELLOS_LOGO_URL = "https://emailsignatureciellosdev.z13.web.core.windows.net/assets/Ciellos_Logo_2Colour-Blue.png";
var CIELLOS_STYLE = "font-family:'Segoe UI',sans-serif;font-size:10pt;color:#0095FE;line-height:1.3";
var CIELLOS_STYLE_INLINE = "font-family:'Segoe UI',sans-serif;font-size:10pt;color:#0095FE";

/**
 * Checks if signature exists.
 * If not, displays the information bar for user.
 * If exists, insert signature into new item (appointment or message).
 * @param {*} eventObj Office event object
 * @returns
 */
function checkSignature(eventObj) {
  let user_info_str = Office.context.roamingSettings.get("user_info");
  if (!user_info_str) {
    display_insight_infobar();
  } else {
    let user_info = JSON.parse(user_info_str);

    if (Office.context.mailbox.item.getComposeTypeAsync) {
      //Find out if the compose type is "newEmail", "reply", or "forward" so that we can apply the correct template.
      Office.context.mailbox.item.getComposeTypeAsync(
        {
          asyncContext: {
            user_info: user_info,
            eventObj: eventObj,
          },
        },
        function (asyncResult) {
          if (asyncResult.status === "succeeded") {
            insert_auto_signature(
              asyncResult.value.composeType,
              asyncResult.asyncContext.user_info,
              asyncResult.asyncContext.eventObj
            );
          }
        }
      );
    } else {
      // Appointment item. Just use newMail pattern
      let user_info = JSON.parse(user_info_str);
      insert_auto_signature("newMail", user_info, eventObj);
    }
  }
}

/**
 * For Outlook on Windows and on Mac only. Insert signature into appointment or message.
 * Outlook on Windows and on Mac can use setSignatureAsync method on appointments and messages.
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @param {*} user_info Information details about the user
 * @param {*} eventObj Office event object
 */
function insert_auto_signature(compose_type, user_info, eventObj) {
  let template_name = get_template_name(compose_type);
  let signature_info = get_signature_info(template_name, user_info);
  addTemplateSignature(signature_info, eventObj);
}

/**
 * 
 * @param {*} signatureDetails object containing:
 *  "signature": The signature HTML of the template,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 * @param {*} eventObj 
 * @param {*} signatureImageBase64 
 */
function addTemplateSignature(signatureDetails, eventObj, signatureImageBase64) {
  if (is_valid_data(signatureDetails.logoBase64) === true) {
    //If a base64 image was passed we need to attach it.
    Office.context.mailbox.item.addFileAttachmentFromBase64Async(
      signatureDetails.logoBase64,
      signatureDetails.logoFileName,
      {
        isInline: true,
      },
      function (result) {
        //After image is attached, insert the signature
        Office.context.mailbox.item.body.setSignatureAsync(
          signatureDetails.signature,
          {
            coercionType: "html",
            asyncContext: eventObj,
          },
          function (asyncResult) {
            asyncResult.asyncContext.completed();
          }
        );
      }
    );
  } else {
    //Image is not embedded, or is referenced from template HTML
    Office.context.mailbox.item.body.setSignatureAsync(
      signatureDetails.signature,
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

/**
 * Creates information bar to display when new message or appointment is created
 */
function display_insight_infobar() {
  Office.context.mailbox.item.notificationMessages.addAsync("fd90eb33431b46f58a68720c36154b4a", {
    type: "insightMessage",
    message: "Please set your signature with the Office Add-ins sample.",
    icon: "Icon.16x16",
    actions: [
      {
        actionType: "showTaskPane",
        actionText: "Set signatures",
        commandId: get_command_id(),
        contextData: "{''}",
      },
    ],
  });
}

/**
 * Gets template name (A,B,C) mapped based on the compose type
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @returns Name of the template to use for the compose type
 */
function get_template_name(compose_type) {
  if (compose_type === "reply") return Office.context.roamingSettings.get("reply");
  if (compose_type === "forward") return Office.context.roamingSettings.get("forward");
  return Office.context.roamingSettings.get("newMail");
}

/**
 * Gets HTML signature in requested template format for given user
 * @param {\} template_name Which template format to use (A,B,C)
 * @param {*} user_info Information details about the user
 * @returns HTML signature in requested template format
 */
function get_signature_info(template_name, user_info) {
  if (template_name === "templateB") return get_template_B_info(user_info);
  if (template_name === "templateC") return get_template_C_info(user_info);
  return get_template_A_info(user_info);
}

/**
 * Gets correct command id to match to item type (appointment or message)
 * @returns The command id
 */
function get_command_id() {
  if (Office.context.mailbox.item.itemType == "appointment") {
    return "MRCS_TpBtn1";
  }
  return "MRCS_TpBtn0";
}

// Template A: New mail — logo (rowspan table), name + title, mobile + office, email, website
function get_template_A_info(user_info) {
  var nameCell = "<span style='" + CIELLOS_STYLE + "'><b>" + user_info.name;
  if (is_valid_data(user_info.pronoun)) nameCell += " " + user_info.pronoun;
  if (is_valid_data(user_info.job)) {
    nameCell += " |</b></span> <span style='" + CIELLOS_STYLE + "'>" + user_info.job + "</span>";
  } else {
    nameCell += "</b></span>";
  }

  var contactCell = "";
  if (is_valid_data(user_info.phone)) {
    contactCell += "<div style='" + CIELLOS_STYLE + "'>Mobile: " + user_info.phone + "</div>";
  }
  contactCell += "<div style='" + CIELLOS_STYLE + "'>Office: +1 (770) 799-8565</div>";
  contactCell += "<div style='" + CIELLOS_STYLE + "'>E-mail: " + user_info.email + "</div>";
  contactCell += "<div style='" + CIELLOS_STYLE + "'><a href='http://www.ciellos.com/' style='color:#0095FE'>www.ciellos.com</a></div>";

  var str = "";
  if (is_valid_data(user_info.greeting)) str += user_info.greeting + "<br/>";

  str += "<table style='text-align:left;background-color:#fff;color:#0095FE;border-collapse:collapse;border-spacing:0;border:0'><tbody>";
  str += "<tr>";
  str +=   "<td rowspan='2' style='text-align:left;padding:0;width:150px;border:0'>";
  str +=     "<img src='" + CIELLOS_LOGO_URL + "' alt='Ciellos logo' width='150' style='width:150px;height:auto;' />";
  str +=   "</td>";
  str +=   "<td style='text-align:left;padding:0 0.75pt 0.75pt 2pt;border:0'>" + nameCell + "</td>";
  str += "</tr>";
  str += "<tr>";
  str +=   "<td style='text-align:left;padding:0.75pt 0.75pt 0 2pt;vertical-align:bottom;border:0'>" + contactCell + "</td>";
  str += "</tr>";
  str += "</tbody></table>";

  return { signature: str, logoBase64: null, logoFileName: null };
}

// Template B: Reply/forward — no logo, name (pronouns) + title, mobile + office, email, website
function get_template_B_info(user_info) {
  var str = "";
  if (is_valid_data(user_info.greeting)) str += user_info.greeting + "<br/>";

  str += "<p style='margin:0'><span style='" + CIELLOS_STYLE_INLINE + "'><b>" + user_info.name;
  if (is_valid_data(user_info.pronoun)) str += " " + user_info.pronoun;
  if (is_valid_data(user_info.job)) {
    str += " |</b></span> <span style='" + CIELLOS_STYLE_INLINE + "'>" + user_info.job + "</span>";
  } else {
    str += "</b></span>";
  }
  str += "</p>";

  if (is_valid_data(user_info.phone)) {
    str += "<div style='" + CIELLOS_STYLE + "'>Mobile: " + user_info.phone + " | Office: +1 (770) 799-8565</div>";
  } else {
    str += "<div style='" + CIELLOS_STYLE + "'>Office: +1 (770) 799-8565</div>";
  }
  str += "<div style='" + CIELLOS_STYLE + "'>E-mail: <span style='color:#0095FE'>" + user_info.email + "</span></div>";
  str += "<div style='" + CIELLOS_STYLE + "'><a href='http://www.ciellos.com/' style='color:#0095FE'>www.ciellos.com</a></div>";

  return { signature: str, logoBase64: null, logoFileName: null };
}

/**
 * Gets HTML string for template C
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template C,
    "logoBase64": null since there is no image,
    "logoFileName": null since there is no image
 */
function get_template_C_info(user_info) {
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += user_info.name;

  return {
    signature: str,
    logoBase64: null,
    logoFileName: null,
  };
}

/**
 * Validates if str parameter contains text.
 * @param {*} str String to validate
 * @returns true if string is valid; otherwise, false.
 */
function is_valid_data(str) {
  return str !== null && str !== undefined && str !== "";
}

Office.actions.associate("checkSignature", checkSignature);
