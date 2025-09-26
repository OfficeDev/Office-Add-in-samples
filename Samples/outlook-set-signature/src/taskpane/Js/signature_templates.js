// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

async function get_template_A_str(user_info)
{

  let str = "";
  if (is_valid_data(user_info.greeting))
  {
    str += user_info.greeting + "<br/>";
  }

  const jpg = await (async function fetchImageAsBase64() {
    const url = "https://boringapi.com/api/v1/static/photos/300.jpeg";
    const response = await fetch(url, { cache: "no-store" });
    if (!response.ok) {
      throw new Error("Failed to fetch image: " + response.status);
    }
    const arrayBuffer = await response.arrayBuffer();
    const bytes = new Uint8Array(arrayBuffer);
    const chunkSize = 0x8000; // avoid call stack limits
    let binary = "";
    for (let i = 0; i < bytes.length; i += chunkSize) {
      const chunk = bytes.subarray(i, i + chunkSize);
      binary += String.fromCharCode.apply(null, chunk);
    }
    return btoa(binary);
  })()


  let prom = Office.context.mailbox.item.addFileAttachmentFromBase64Async(jpg, "test.jpg", { isInline: true }, function (result) {

    return new Promise( function( resolve, reject ) {
      if (result.status!=="succeeded") {
        console.error("Attaching failed:", result)
      }
      resolve(result.value);
    });
    
  });

  let cid = await prom;
  
  console.log("cid", cid);
  str += "<table>";
  str +=   "<tr>";
  str +=     `<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='cid:${cid}' alt='Logo' /></td>`;
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
  str +=     "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='https://officedev.github.io/Office-Add-in-samples/Samples/outlook-set-signature/assets/sample-logo.png' alt='Logo' /></td>";  
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