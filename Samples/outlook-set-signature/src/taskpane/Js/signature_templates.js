// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

var CIELLOS_LOGO_URL = "https://emailsignatureciellosdev.z13.web.core.windows.net/assets/Ciellos_Logo_2Colour-Blue.png";
var CIELLOS_STYLE = "font-family:'Segoe UI',sans-serif;font-size:10pt;color:#0095FE;line-height:1.3";
var CIELLOS_STYLE_INLINE = "font-family:'Segoe UI',sans-serif;font-size:10pt;color:#0095FE";

// Template A: New mail — logo (rowspan table), name + title, mobile + office, email, website
function get_template_A_str(user_info) {
  var str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

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

  return str;
}

// Template B: Reply/forward — no logo, name (pronouns) + title, mobile + office, email, website
function get_template_B_str(user_info) {
  var str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

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

  return str;
}
