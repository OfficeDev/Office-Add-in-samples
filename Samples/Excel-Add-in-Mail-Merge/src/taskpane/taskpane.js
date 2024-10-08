/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  // Check that we loaded into Excel
  if (info.host === Office.HostType.Excel) {
    let userClientId = 'YOUR_APP_ID_HERE'; //Register your app at https://aad.portal.azure.com/
    localStorage.setItem('client-id', userClientId);

    document.getElementById("sendEmail").onclick = checkClientID;
    document.getElementById("createSampleData").onclick = createSampleData;
  }
});

class DialogAPIAuthProvider {
  async getAccessToken() {
    if (this._accessToken) {
      return this._accessToken;
    } else {
      return this.login();
    }
  }

  async login() {
    return new Promise((resolve, reject) => {
      let data = encodeURIComponent(localStorage.getItem('client-id'));
      const dialogLoginUrl = location.href.substring(0, location.href.lastIndexOf('/')) + `/consent.html?data=${data}`;
      Office.context.ui.displayDialogAsync(
        dialogLoginUrl,
        { height: 60, width: 60 },
        result => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            reject(result.error);
          }
          else {
            const loginDialog = result.value;

            loginDialog.addEventHandler(Office.EventType.DialogEventReceived, args => {
              reject(args.error);
            });

            loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, args => {
              const messageFromDialog = JSON.parse(args.message);

              loginDialog.close();

              if (messageFromDialog.status === 'success') {
                // We now have a valid access token.
                this._accessToken = messageFromDialog.result;
                resolve(this._accessToken);
              }
              else {
                // Something went wrong with authentication or the authorization of the web application.
                reject(messageFromDialog.result);
              }
            });
          }
        }
      );
    });
  }
}

const dialogAPIAuthProvider = new DialogAPIAuthProvider();

// Display a status
/**
 * @param {unknown} message
 * @param {boolean} isError
 */
function showStatus(message, isError) {
  $('.status').empty();
  $('<div/>', {
    class: `status-card ms-depth-4 ${isError ? 'error-msg' : 'success-msg'}`
  }).append($('<p/>', {
    class: 'ms-fontSize-24 ms-fontWeight-bold',
    text: isError ? 'An error occurred' : 'Success'
  })).append($('<p/>', {
    class: 'ms-fontSize-16 ms-fontWeight-regular',
    text: message
  })).appendTo('.status');
}

// Clear the status
function clearStatus() {
  $('.status').empty();
}

// Create Sample Data
async function createSampleData() {
  try {
    await Excel.run(async (context) => {
      context.workbook.worksheets.getItemOrNullObject("Sample").delete();
      const sheet = context.workbook.worksheets.add("Sample");

      let invoiceTable = sheet.tables.add("A1:E1", true);
      invoiceTable.name = "InvoiceTable";
      let range = invoiceTable.getRange();
      range.numberFormat = "@";
      invoiceTable.getHeaderRowRange().values = [["Email", "Name", "Invoice Number", "Amount", "Due Date"]];

      invoiceTable.rows.add(0, [
        ["client1@email.com", "John", "INV001", "$500", "2023-11-15"],
        ["client2@email.com", "Sarah", "INV002", "$750", "2023-11-20"],
        ["client3@microsoft.com", "Michael", "INV003", "$300", "2023-11-10"],
        ["client4@microsoft.com", "Lisa", "INV004", "$900", "2023-11-15"]
      ]);

      invoiceTable.getRange().format.autofitColumns();
      invoiceTable.getRange().format.autofitRows();

      sheet.activate();
      await context.sync();
    });
  } catch (error) {
    showStatus(`Exception when creating sample data: ${JSON.stringify(error)}`, true);
  }
}

// Use regular expressions to get text within <<...>>
/**
 * 
 * @param {*} str 
 * @returns 
 */
function getStr(str) {
  var result = str.match(/\<\<(.*?)\>\>/g);
  if (result) {
    for (var i = 0; i < result.length; i++) {
      result[i] = result[i].replace(/\<\<|\>\>/g, '');
    }
    return result[0];
  }
}

// <SendEmailSnippet>
/**
 * @param {{ preventDefault: () => void; }} evt
 */
async function checkClientID(evt) {
  evt.preventDefault();
  clearStatus();

  let userClientId = localStorage.getItem('client-id');

  if (userClientId == "YOUR_APP_ID_HERE") {
    let resultPromise = new Promise((resolve, reject) => {
      const dialogLoginUrl = location.href.substring(0, location.href.lastIndexOf('/')) + '/enterClientId.html';
      Office.context.ui.displayDialogAsync(
        dialogLoginUrl,
        { height: 25, width: 40 },
        result => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            reject(result.error);
          }
          else {
            const loginDialog = result.value;

            loginDialog.addEventHandler(Office.EventType.DialogEventReceived, args => {
              reject(args.error);
            });

            loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, args => {
              userClientId = args.message;
              localStorage.setItem('client-id', userClientId);
              loginDialog.close();

              sendEmails();
            });
          }
        }
      );
    });
  }
  else {
    sendEmails();
  }
}

async function sendEmails() {
  const graphClient = MicrosoftGraph.Client.initWithMiddleware({ authProvider: dialogAPIAuthProvider });
  await Excel.run(async (context) => {
    try {
      const subject = $('#Subject').val();
      const to = $('#ToLine').val();
      const content = $('#Content').val();

      if (to && subject && content) {
        var emailAddress = getStr(to);

        if (emailAddress != null) {
          const addressColumn = context.workbook.tables.getItem("InvoiceTable").columns.getItemOrNullObject(emailAddress);
          addressColumn.load();
          await context.sync();

          if (!addressColumn.isNullObject) {
            let i = 1;
            let addressValue = addressColumn.values;
            let subject_str = subject.toString();

            // Scan the table and send emails row by row
            while (i < addressValue.length) {
              let finalSubject = '';
              //Replace the column name in the subject textarea with the corresponding value in the table
              for (let j = 0; j < subject_str.length; j++) {
                if (subject_str[j] == "<") {
                  let replace_str = getStr(subject_str.substring(j));
                  let replaceColumn = context.workbook.tables.getItem("InvoiceTable").columns.getItemOrNullObject(replace_str);
                  replaceColumn.load();
                  await context.sync();

                  if (!replaceColumn.isNullObject) {
                    let replaceValue = replaceColumn.values;
                    finalSubject += replaceValue[i][0];

                    while (j < subject_str.length) {
                      if (subject_str[j] == ">") {
                        j += 1;
                        break;
                      }
                      j += 1;
                    }
                  }
                  else {
                    showStatus(`There is no corresponding column name as "${replace_str}" in Subject.`, true);
                    return;
                  }
                }
                else {
                  finalSubject += subject_str[j];
                }
              }

              //Replace the column name in the content textarea with the corresponding value in the table
              let content_str = content.toString();
              let finalContent = '';
              for (let k = 0; k < content_str.length; k++) {
                if (content_str[k] == "<") {
                  let replace_str = getStr(content_str.substring(k));
                  let replaceColumn = context.workbook.tables.getItem("InvoiceTable").columns.getItemOrNullObject(replace_str);
                  replaceColumn.load();
                  await context.sync();

                  if (!replaceColumn.isNullObject) {
                    let replaceValue = replaceColumn.values;
                    finalContent += replaceValue[i][0];
                    while (k < content_str.length) {
                      if (content_str[k] == ">") {
                        k += 1;
                        break;
                      }
                      k += 1;
                    }
                  }
                  else {
                    showStatus(`There is no corresponding column name as "${replace_str}" in Content.`, true);
                    return;
                  }
                }
                else {
                  finalContent += content_str[k];
                }
              }

              try {

                const sendMail =
                {
                  message: {
                    subject: finalSubject,
                    body: {
                      contentType: 'Text',
                      content: finalContent
                    },
                    toRecipients: [{
                      emailAddress: {
                        address: addressValue[i][0]
                      }
                    }]
                  }
                };

                await graphClient.api('me/SendMail')
                  .post(sendMail);
              }
              catch (error) {
                console.log(`Error: ${JSON.stringify(error)}`);
                showStatus(`Exception sending emails via Graph: ${JSON.stringify(error)}`, true);
                return;
              }

              i++;
            }
            showStatus(`Already sent ${i - 1} emails via Microsoft Graph.`, false);
          }
          else {
            showStatus(`There is no corresponding column name as "${emailAddress}" in ToLine.`, true);
          }
        }
        else {
          showStatus(`There is no corresponding column name in ToLine.`, true);
        }
      }
      else {
        showStatus(`Please fill in all the fields.`, true);
      }
    } catch (error) {
      console.log(`Error: ${JSON.stringify(error)}`);
      showStatus(`Exception sending emails via Graph: ${JSON.stringify(error)}`, true);
    }
  })
}
// </SendEmailSnippet>