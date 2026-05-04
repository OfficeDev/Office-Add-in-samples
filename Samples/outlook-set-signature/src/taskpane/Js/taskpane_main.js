// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Flushes in-memory roamingSettings to the server. Must be called after any .set() calls.
function save_user_settings_to_roaming_settings()
{
  Office.context.roamingSettings.saveAsync(function (asyncResult)
  {
	console.log("save_user_info_str_to_roaming_settings - " + JSON.stringify(asyncResult));
  });
}

// Tells Outlook to stop inserting its own default signature so ours takes precedence.
function disable_client_signatures_if_necessary()
{
  if ($("#checkbox_sig").prop("checked") === true)
  {
	Office.context.mailbox.item.disableClientSignatureAsync(function (asyncResult)
	{
	  console.log("disable_client_signature_if_necessary - " + JSON.stringify(asyncResult));
	});
  }
}

function save_signature_settings()
{
  // user_info is written to localStorage by editsignature.html; if it's missing the user
  // hasn't completed the profile form yet, so there's nothing to save.
  let user_info_str = localStorage.getItem('user_info');

  if (user_info_str)
  {
	if (!_user_info)
	{
	  _user_info = JSON.parse(user_info_str);
	}

	const newMail = $("#new_mail option:selected").val();
	const reply   = $("#reply option:selected").val();
	const forward = $("#forward option:selected").val();
	const override = $("#checkbox_sig").prop('checked');

	// localStorage is always available (web, new Outlook, test harness outside Outlook).
	localStorage.setItem('newMail', newMail);
	localStorage.setItem('reply', reply);
	localStorage.setItem('forward', forward);
	localStorage.setItem('override_olk_signature', override);

	// roamingSettings syncs across devices via Exchange; only available in classic Outlook.
	if (Office.context && Office.context.roamingSettings)
	{
	  Office.context.roamingSettings.set('user_info', user_info_str);
	  Office.context.roamingSettings.set('newMail', newMail);
	  Office.context.roamingSettings.set('reply', reply);
	  Office.context.roamingSettings.set('forward', forward);
	  Office.context.roamingSettings.set('override_olk_signature', override);
	  save_user_settings_to_roaming_settings();
	  disable_client_signatures_if_necessary();
	}

	$("#message").show("slow");
  }
}



function set_body(str)
{
  Office.context.mailbox.item.body.setAsync
  (
	get_cal_offset() + str,

	{
		coercionType: Office.CoercionType.Html
	},

	function (asyncResult)
	{
	  console.log("set_body - " + JSON.stringify(asyncResult));
	}
  );
}

function set_signature(str)
{
  Office.context.mailbox.item.body.setSignatureAsync
  (
	str,

	{
		coercionType: Office.CoercionType.Html
	},

	function (asyncResult)
	{
	  console.log("set_signature - " + JSON.stringify(asyncResult));
	}
  );
}

function insert_signature(str)
{
  if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Appointment)
  {
	set_body(str);
  }
  else
  {
	set_signature(str);
  }
}

function test_template_A()
{
	let str = get_template_A_str(_user_info);
	console.log("test_template_A - " + str);

	insert_signature(str);
}

function test_template_B()
{
	let str = get_template_B_str(_user_info);
	console.log("test_template_B - " + str);

	insert_signature(str);
}

function test_template_C()
{
	let str = get_template_C_str(_user_info);
	console.log("test_template_C - " + str);

	insert_signature(str);
}

function navigate_to_taskpane2()
{
  window.location.href = 'editsignature.html';
}