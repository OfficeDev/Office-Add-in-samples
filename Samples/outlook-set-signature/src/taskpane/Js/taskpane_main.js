// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

function save_user_settings_to_roaming_settings() {
  Office.context.roamingSettings.saveAsync(function (asyncResult) {
    console.log("save_user_info_str_to_roaming_settings - " + JSON.stringify(asyncResult));
  });
}

function disable_client_signatures_if_necessary() {
  if (document.getElementById("checkbox_sig").checked === true) {
    Office.context.mailbox.item.disableClientSignatureAsync(function (asyncResult) {
      console.log("disable_client_signature_if_necessary - " + JSON.stringify(asyncResult));
    });
  }
}

function save_signature_settings() {
  let user_info_str = localStorage.getItem("user_info");

  if (user_info_str) {
    if (!window._user_info) {
      window._user_info = JSON.parse(user_info_str);
    }

    Office.context.roamingSettings.set("user_info", user_info_str);
    Office.context.roamingSettings.set("newMail", document.getElementById("new_mail").value);
    Office.context.roamingSettings.set("reply", document.getElementById("reply").value);
    Office.context.roamingSettings.set("forward", document.getElementById("forward").value);

    Office.context.roamingSettings.set("override_olk_signature", document.getElementById("checkbox_sig").checked);

    save_user_settings_to_roaming_settings();

    disable_client_signatures_if_necessary();

    document.getElementById("message").style.display = "block";
  } else {
    // TBD display an error somewhere?
  }
}

function set_body(str) {
  Office.context.mailbox.item.body.setAsync(
    get_cal_offset() + str,

    {
      coercionType: Office.CoercionType.Html,
    },

    function (asyncResult) {
      console.log("set_body - " + JSON.stringify(asyncResult));
    }
  );
}

function set_signature(str) {
  Office.context.mailbox.item.body.setSignatureAsync(
    str,

    {
      coercionType: Office.CoercionType.Html,
    },

    function (asyncResult) {
      console.log("set_signature - " + JSON.stringify(asyncResult));
    }
  );
}

function insert_signature(str) {
  if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Appointment) {
    set_body(str);
  } else {
    set_signature(str);
  }
}

function test_template_A() {
  let str = get_template_A_str(window._user_info);
  console.log("test_template_A - " + str);

  insert_signature(str);
}

function test_template_B() {
  let str = get_template_B_str(window._user_info);
  console.log("test_template_B - " + str);

  insert_signature(str);
}

function test_template_C() {
  let str = get_template_C_str(window._user_info);
  console.log("test_template_C - " + str);

  insert_signature(str);
}

function navigate_to_taskpane2() {
  window.location.href = "editsignature.html";
}

// Expose functions referenced by inline onclick handlers to global scope
window.test_template_A = test_template_A;
window.test_template_B = test_template_B;
window.test_template_C = test_template_C;
window.save_signature_settings = save_signature_settings;
window.navigate_to_taskpane2 = navigate_to_taskpane2;
