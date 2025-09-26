// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

let _user_info;

Office.initialize = function(reason)
{
  on_initialization_complete();
}

function on_initialization_complete()
{
	$(document).ready
	(
		async function()
		{
      await lazy_init_user_info();
      populate_templates();
      show_signature_settings();
		}
	);
}

async function lazy_init_user_info()
{
  if (!_user_info)
  {
    let user_info_str = localStorage.getItem('user_info');

    if (user_info_str)
    {
      _user_info = JSON.parse(user_info_str);
    }
    else
    {
      console.log("Unable to retrieve 'user_info' from localStorage.");
    }
  }
}


function populate_templates()
{
  populate_template_A();
  populate_template_B();
  populate_template_C();
}

async function populate_template_A()
{
  let str = await get_template_A_str(_user_info);
  $("#box_1").html(str);
}

function populate_template_B()
{
  let str = get_template_B_str(_user_info);
  $("#box_2").html(str);
}

function populate_template_C()
{
  let str = get_template_C_str(_user_info);
  $("#box_3").html(str);
}

function show_signature_settings()
{
  let val = Office.context.roamingSettings.get("newMail");
  if (val)
  {
    $("#new_mail").val(val);
  }

  val = Office.context.roamingSettings.get("reply");
  if (val)
  {
    $("#reply").val(val);
  }

  val = Office.context.roamingSettings.get("forward");
  if (val)
  {
    $("#forward").val(val);
  }

  val = Office.context.roamingSettings.get("override_olk_signature");
  if (val != null)
  {
    $("#checkbox_sig").prop('checked', val);
  }
}