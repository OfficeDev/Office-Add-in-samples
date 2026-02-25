// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

window._user_info = null;

Office.onReady(() => {
  on_initialization_complete();
});

function on_initialization_complete() {
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initAssignPage);
  } else {
    initAssignPage();
  }
}

function initAssignPage() {
  lazy_init_user_info();
  populate_templates();
  show_signature_settings();
}

function lazy_init_user_info() {
  if (!window._user_info) {
    let user_info_str = localStorage.getItem("user_info");

    if (user_info_str) {
      window._user_info = JSON.parse(user_info_str);
    } else {
      console.log("Unable to retrieve 'user_info' from localStorage.");
    }
  }
}

function populate_templates() {
  populate_template_A();
  populate_template_B();
  populate_template_C();
}

function populate_template_A() {
  let str = get_template_A_str(window._user_info);
  document.getElementById("box_1").innerHTML = str;
}

function populate_template_B() {
  let str = get_template_B_str(window._user_info);
  document.getElementById("box_2").innerHTML = str;
}

function populate_template_C() {
  let str = get_template_C_str(window._user_info);
  document.getElementById("box_3").innerHTML = str;
}

function show_signature_settings() {
  let val = Office.context.roamingSettings.get("newMail");
  if (val) {
    document.getElementById("new_mail").value = val;
  }

  val = Office.context.roamingSettings.get("reply");
  if (val) {
    document.getElementById("reply").value = val;
  }

  val = Office.context.roamingSettings.get("forward");
  if (val) {
    document.getElementById("forward").value = val;
  }

  val = Office.context.roamingSettings.get("override_olk_signature");
  if (val != null) {
    document.getElementById("checkbox_sig").checked = val;
  }
}
