/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// Helper functions

// Handles errors and complete the event with failure.
function handleError(event, errorMessage, logMessage) {
  if (logMessage) {
    console.log(logMessage);
  }
  event.completed({
    allowEvent: false,
    errorMessage: errorMessage,
  });
}

// Creates attachment details object based on attachment type.
function createAttachmentDetails(attachment) {
  const attachmentType = attachment.attachmentType;
  const attachmentName = attachment.name;

  switch (attachmentType) {
    case Office.MailboxEnums.AttachmentType.Cloud:
      return {
        attachmentType: attachmentType,
        name: attachmentName,
        path: attachment.url,
      };
    case Office.MailboxEnums.AttachmentType.File:
      return {
        attachmentType: attachmentType,
        isInline: attachment.isInline,
        contentId: attachment.contentId,
        name: attachmentName,
      };
    case Office.MailboxEnums.AttachmentType.Item:
      return {
        attachmentType: attachmentType,
        name: attachmentName,
      };
    default:
      return null;
  }
}

// Creates decrypted attachment object based on attachment details.
function createDecryptedAttachment(decryptedDetails) {
  const attachmentType = decryptedDetails.attachmentType;
  const attachmentName = decryptedDetails.name;

  switch (attachmentType) {
    case Office.MailboxEnums.AttachmentType.Cloud:
      return {
        attachmentType: attachmentType,
        name: attachmentName,
        path: decryptedDetails.path,
      };
    case Office.MailboxEnums.AttachmentType.File:
      return {
        attachmentType: attachmentType,
        content: decryptedDetails.content,
        isInline: decryptedDetails.isInline,
        contentId: decryptedDetails.contentId,
        name: attachmentName,
      };
    case Office.MailboxEnums.AttachmentType.Item:
      return {
        attachmentType: attachmentType,
        content: decryptedDetails.content,
        name: attachmentName,
      };
    default:
      return null;
  }
}

// Event handlers

// Handles the OnMessageSend event to encrypt the message body and attachments.
function onMessageSendHandler(event) {
  Office.context.mailbox.item.body.getAsync(
    Office.CoercionType.Html,
    { asyncContext: { event: event } },
    (result) => {
      const event = result.asyncContext.event;
      if (result.status === Office.AsyncResultStatus.Failed) {
        handleError(
          event,
          "Unable to encrypt the contents of this message.",
          `Failed to get body content: ${result.error.message}.`
        );
        return;
      }

      // Encrypt the content of the message body.
      const body = result.value;
      const encryptedBody = encrypt(body);
      const placeholderMessage = `
        <div style="font-family: Segoe UI, Helvetica, Arial, sans-serif; font-size: 14px; color: #333;">
          <h2 style="color: #0078d4; font-size: 18px; margin-bottom: 10px;">🔒 This message is encrypted</h2>
          <p style="margin: 10px 0;">This message has been encrypted by the <strong>Office Add-ins Sample Encryption Add-in</strong> for your security and privacy.</p>
          <p style="margin: 10px 0;">To view the contents of this message, you must install the add-in in your Outlook client.</p>
          </div>
          <p style="margin: 10px 0; color: #605e5c;">
            <a href="https://learn.microsoft.com/office/dev/add-ins/outlook/encryption-decryption" target="_blank" style="color: #0078d4; text-decoration: none;">Learn more about encryption.</a>
          </p>
        </div>
      `.trim();
      Office.context.mailbox.item.body.setAsync(
        placeholderMessage,
        { 
          asyncContext: { event: event },
          coercionType: Office.CoercionType.Html,
        },
        (result) => {
          const event = result.asyncContext.event;
          if (result.status === Office.AsyncResultStatus.Failed) {
            handleError(
              event,
              "Unable to encrypt the contents of this message.",
              `Failed to set placeholder message: ${result.error.message}.`
            );
            return;
          }

          // Add the encrypted body as an attachment.
          Office.context.mailbox.item.addFileAttachmentFromBase64Async(
            encryptedBody,
            "encrypted_body.txt",
            { asyncContext: { event: event } },
            (result) => {
              const event = result.asyncContext.event;
              if (result.status === Office.AsyncResultStatus.Failed) {
                handleError(
                  event,
                  "Unable to encrypt the contents of this message.",
                  `Failed to add attachment: ${result.error.message}.`
                );
                return;
              }

              // Encrypt all other attachments, if any.
              Office.context.mailbox.item.getAttachmentsAsync(
                { asyncContext: { event: event } },
                (result) => {
                  const event = result.asyncContext.event;
                  if (result.status === Office.AsyncResultStatus.Failed) {
                    handleError(
                      event,
                      "Unable to encrypt the contents of this message.",
                      "Failed to get attachments for encryption."
                    );
                    return;
                  }

                  const attachments = result.value;
                  const attachmentPromises = attachments
                    .filter((attachment) => attachment.name !== "encrypted_body.txt")
                    .map((attachment, i) => {
                      return new Promise((resolve, reject) => {
                        const attachmentDetails = createAttachmentDetails(attachment);

                        Office.context.mailbox.item.getAttachmentContentAsync(
                          attachment.id,
                          {
                            asyncContext: {
                              attachmentNumber: i,
                              attachmentDetails: attachmentDetails,
                              attachmentId: attachment.id,
                            },
                          },
                          (result) => {
                            if (result.status === Office.AsyncResultStatus.Failed) {
                              console.log("Failed to get attachment content for encryption.");
                              reject(result.error);
                              return;
                            }

                            const attachmentDetails = result.asyncContext.attachmentDetails;
                            const content = result.value.content;
                            attachmentDetails.content = content;
                            const attachmentStr = JSON.stringify(attachmentDetails);
                            const encryptedContent = encrypt(attachmentStr);
                            const fileIdentifier = `encrypted_attachment_${result.asyncContext.attachmentNumber}.txt`;
                            const attachmentId = result.asyncContext.attachmentId;

                            // Add encrypted attachments to the message.
                            Office.context.mailbox.item.addFileAttachmentFromBase64Async(
                              encryptedContent,
                              fileIdentifier,
                              {
                                asyncContext: { attachmentId: attachmentId },
                                isInline: false,
                              },
                              (result) => {
                                if (result.status === Office.AsyncResultStatus.Failed) {
                                  console.log("Failed to add encrypted attachment.");
                                  reject(result.error);
                                  return;
                                }

                                // Remove the original attachments from the message.
                                const attachmentId = result.asyncContext.attachmentId;
                                Office.context.mailbox.item.removeAttachmentAsync(
                                  attachmentId,
                                  { asyncContext: { event: event } },
                                  (result) => {
                                    if (result.status === Office.AsyncResultStatus.Failed) {
                                      console.log("Failed to remove the attachment.");
                                      reject(result.error);
                                      return;
                                    }
                                    resolve();
                                  }
                                );
                              }
                            );
                          }
                        );
                      });
                    });

                  Promise.all(attachmentPromises)
                    .then(() => {
                      // Set an internet header to indicate that the message is encrypted.
                      Office.context.mailbox.item.internetHeaders.setAsync(
                        { "contoso-encrypted": "contoso-encrypted" },
                        { asyncContext: { event: event } },
                        (result) => {
                          const event = result.asyncContext.event;
                          if (result.status === Office.AsyncResultStatus.Failed) {
                            handleError(
                              event,
                              "Unable to encrypt the contents of this message.",
                              `Failed to set internet header: ${result.error.message}.`
                            );
                            return;
                          }
                          event.completed({ allowEvent: true });
                        }
                      );
                    })
                    .catch((error) => {
                      handleError(
                        event,
                        "Unable to encrypt the contents of this message.",
                        `Unable to encrypt attachments: ${error}`
                      );
                    });
                }
              );
            }
          );
        }
      );
    }
  );
}

// Handles the OnMessageRead event to decrypt the message body and attachments.
function onMessageReadHandler(event) {
  const attachments = Office.context.mailbox.item.attachments;
  if (attachments.length === 0) {
    console.log("No attachments found for decryption.");
    event.completed({ allowEvent: false });
    return;
  }

  let decryptedAttachments = [];
  let decryptedBody;

  const attachmentPromises = attachments.map((attachment) => {
    return new Promise((resolve, reject) => {
      const attachmentName = attachment.name;
      const attachmentId = attachment.id;
      Office.context.mailbox.item.getAttachmentContentAsync(
        attachmentId,
        { asyncContext: { attachmentName: attachmentName, event: event } },
        (result) => {
          const event = result.asyncContext.event;
          if (result.status === Office.AsyncResultStatus.Failed) {
            console.log("Failed to get content of encrypted attachments.");
            reject(result.error.message);
            event.completed({ allowEvent: false });
            return;
          }

          const attachmentContent = result.value.content;
          const decryptedData = decrypt(attachmentContent);
          const encryptedAttachmentName = result.asyncContext.attachmentName;
          if (encryptedAttachmentName.includes("encrypted_attachment")) {
            const decryptedAttachmentDetails = JSON.parse(decryptedData);

            const decryptedAttachment = createDecryptedAttachment(decryptedAttachmentDetails);
            if (decryptedAttachment) {
              decryptedAttachments.push(decryptedAttachment);
            }
          } else if (encryptedAttachmentName.includes("encrypted_body")) {
            decryptedBody = {
              coercionType: Office.CoercionType.Html,
              content: decryptedData,
            };
          } else {
            console.log("Not an encrypted attachment.");
          }

          resolve();
        }
      );
    });
  });

  Promise.all(attachmentPromises)
    .then(() => {
      if (decryptedAttachments.length === 0) {
        event.completed({
          allowEvent: true,
          emailBody: decryptedBody,
        });
      } else {
        event.completed({
          allowEvent: true,
          emailBody: decryptedBody,
          attachments: decryptedAttachments,
        });
      }
    })
    .catch((error) => {
      console.log(`Unable to decrypt attachments: ${error}`);
      event.completed({ allowEvent: false });
    });
}

// Custom Base64 encoding for compatibility with event-based add-in.
function customBtoa(str) {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  let result = "";
  let i = 0;

  while (i < str.length) {
    const a = str.charCodeAt(i++);
    const b = i < str.length ? str.charCodeAt(i++) : 0;
    const c = i < str.length ? str.charCodeAt(i++) : 0;

    const encoded = (a << 16) | (b << 8) | c;

    result += chars.charAt((encoded >> 18) & 63);
    result += chars.charAt((encoded >> 12) & 63);
    result += chars.charAt((encoded >> 6) & 63);
    result += chars.charAt(encoded & 63);
  }

  const padding = str.length % 3;
  if (padding === 1) {
    result = result.slice(0, -2) + "==";
  } else if (padding === 2) {
    result = result.slice(0, -1) + "=";
  }

  return result;
}

// Custom Base64 decoding for compatibility with event-based add-in.
function customAtob(str) {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  let result = "";

  // Remove padding
  str = str.replace(/=+$/, "");

  for (let i = 0; i < str.length; i += 4) {
    const encoded =
      (chars.indexOf(str[i]) << 18) |
      (chars.indexOf(str[i + 1]) << 12) |
      (chars.indexOf(str[i + 2] || "A") << 6) |
      chars.indexOf(str[i + 3] || "A");

    result += String.fromCharCode((encoded >> 16) & 255);
    if (str[i + 2] && str[i + 2] !== "=") {
      result += String.fromCharCode((encoded >> 8) & 255);
    }
    if (str[i + 3] && str[i + 3] !== "=") {
      result += String.fromCharCode(encoded & 255);
    }
  }

  return result;
}

// Encryption method for sample purposes only. Do NOT use in production.
function encrypt(input, key = "OfficeAddInSampleKey") {
  try {
    let result = "";
    for (let i = 0; i < input.length; i++) {
      const charCode = input.charCodeAt(i);
      const keyChar = key.charCodeAt(i % key.length);
      result += String.fromCharCode(charCode ^ keyChar);
    }

    return customBtoa(result);
  } catch (error) {
    console.error("Encryption error:", error);
    return input;
  }
}

// Decryption method for sample purposes only. Do NOT use in production.
function decrypt(encryptedText, key = "OfficeAddInSampleKey") {
  try {
    const encrypted = customAtob(encryptedText);
    let result = "";
    for (let i = 0; i < encrypted.length; i++) {
      const charCode = encrypted.charCodeAt(i);
      const keyChar = key.charCodeAt(i % key.length);
      result += String.fromCharCode(charCode ^ keyChar);
    }

    return result;
  } catch (error) {
    console.error("Decryption error:", error);
    return encryptedText;
  }
}

// IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("onMessageReadHandler", onMessageReadHandler);
