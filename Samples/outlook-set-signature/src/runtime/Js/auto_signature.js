// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

function get_command_id() {
  if (Office.context.mailbox.item.itemType == "appointment") {
  	return "MRCS_TpBtn1";
  }
  return "MRCS_TpBtn0";
}
function display_insight_infobar() {
  Office.context.mailbox.item.notificationMessages.addAsync(
  	"fd90eb33431b46f58a68720c36154b4a",
	{
	  type: "insightMessage",
	  message: "Please set your signature with the PnP sample add-in.",
	  icon: "Icon.16x16",
	  actions:
	  [
		{
		"actionType" : "showTaskPane",
		"actionText" : "Set signatures",
		"commandId" : get_command_id(),
		"contextData" : "{''}"
		}
	  ]
    },
    function (asyncResult)
    {
	  // console.log("display_insight_infobar - " + JSON.stringify(asyncResult));
    }
  );
}
function is_valid_data(str)
{
  return str !== null
	&& str !== undefined
	&& str !== "";
}
function get_template_A_str(user_info)
{
  let str = "";
  if (is_valid_data(user_info.greeting))
  {
	str += user_info.greeting + "<br/>";
  }

  str += "<table>";
  str +=   "<tr>";
  str +=     "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='https://localhost:3000/assets/sample-logo.png' alt='Logo' /></td>";
  str +=     "<td style='padding-left: 5px;'>";
  str +=	   "<strong>" + user_info.name + "</strong>";
  str +=     is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str +=     "<br/>";
  str +=	   is_valid_data(user_info.job) ? user_info.job + "<br/>" : "";
  str +=	   user_info.email + "<br/>";
  str +=	   is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str +=     "</td>";
  str +=   "</tr>";
  str += "</table>";

  return str;
}
function get_template_B_str(user_info)
{
  let str = "";
  if (is_valid_data(user_info.greeting))
  {
    str += user_info.greeting + "<br/>";
  }

  str += "<table>";
  str +=   "<tr>";
  str +=     "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='https://localhost:3000/assets/sample-logo.png' alt='Logo' /></td>";
  str +=     "<td style='padding-left: 5px;'>";
  str +=	   "<strong>" + user_info.name + "</strong>";
  str +=     is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str +=     "<br/>";
  str +=	   user_info.email + "<br/>";
  str +=	   is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str +=     "</td>";
  str +=   "</tr>";
  str += "</table>";

  return str;
}
function get_template_C_str(user_info)
{
  let str = "";
  if (is_valid_data(user_info.greeting))
  {
	str += user_info.greeting + "<br/>";
  }

  str += user_info.name;

  return str;
}
function get_signature_str(template_name, user_info) {
  if (template_name === "templateB") return get_template_B_str(user_info);
  if (template_name === "templateC") return get_template_C_str(user_info);
  return get_template_A_str(user_info);
}
function get_template_name(compose_type) {
  if (compose_type === "reply") return Office.context.roamingSettings.get("reply");
  if (compose_type === "forward") return Office.context.roamingSettings.get("forward");
  return Office.context.roamingSettings.get("newMail");
}
function insert_auto_signature(compose_type, user_info, eventObj) {
  // console.log("insert_auto_signature - " + compose_type);
  // console.log("insert_auto_signature - " + user_info);
  let template_name = get_template_name(compose_type);
  let signature_str = get_signature_str(template_name, user_info);
  Office.context.mailbox.item.body.setSignatureAsync
  (
    signature_str,
    {
      "coercionType": "html",
      "asyncContext": eventObj
    },
    function (asyncResult)
    {
      asyncResult.asyncContext.completed({ "key00" : "val00" });
      // console.log("setSignatureAsync - " + JSON.stringify(asyncResult));
    }
  );
}
function checkSignature(eventObj) {
  let user_info_str = Office.context.roamingSettings.get("user_info");
  if (!user_info_str)
  {
	display_insight_infobar();
  }
  else
  {
	let user_info = JSON.parse(user_info_str);

	if (Office.context.mailbox.item.getComposeTypeAsync) {
	  Office.context.mailbox.item.getComposeTypeAsync
	  (
		{
		  "asyncContext" :
		  {
			"user_info" : user_info,
			"eventObj" : eventObj
		  }
		},
		function (asyncResult)
		{
		  // console.log("getComposeTypeAsync - " + JSON.stringify(asyncResult));
		  if (asyncResult.status === "succeeded")
		  {
			insert_auto_signature(
			asyncResult.value.composeType,
			asyncResult.asyncContext.user_info,
			asyncResult.asyncContext.eventObj);
		  }
		}
	  );
	}
	else {
	  // console.log("this is an Appointment item!");
	  let user_info = JSON.parse(user_info_str);
	  insert_auto_signature("newMail", user_info, eventObj);
	}
  }
  return [2 /*return*/];
}

Office.actions.associate("checkSignature", checkSignature);