/* Copyright (c) Maarten van Stam and Eric Legault. All rights reserved. Licensed under the MIT License. */

// https://github.com/justmarks/UiPathOutlook/blob/2247bb3f5f52f8ba6763374fea3f25e1ece5d03b/src/commands/commands.js

export async function getEmailData() {

    try {
        console.log(`Reading mailbox item`);

        const item = Office.context.mailbox.item;

        if (item === null) {
            console.error(`Index.razor.js(getEmailData): Unexpected - could not get reference to Office.context.mailbox.item`);
            console.error(`Index.razor.js(getEmailData) Catch Exception: ${err}`);

            return { Subject: "No email item" };
        }

        console.log("Email Subject: " + item.subject);
        
        var listOfAttachments = [];
        var attachments = item.attachments;

        console.log(`Index.razor.js(getEmailData): Counting ${item.attachments.length} attachment(s)...`);

        if (item.attachments.length > 0) {

            console.log("Trying to fetch attachments...", new Date());

            try {
                const returnedAttachments = await getAttachments(item);
                console.log(returnedAttachments);

                for (var i = 0; i < returnedAttachments.length; i++) {

                    console.log(`Index.razor.js (processAttachmentsAsync2): Attachment - ` + i);
                    console.log(`Index.razor.js (processAttachmentsAsync2): AttachmentID - ` + item.attachments[i].id);

                    let attachmentName = item.attachments[i].name;
                    console.log(`Index.razor.js(getEmailData): Processing attachment '${attachmentName}'...`);

                    let attachmentType = item.attachments[i].attachmentType;
                    let fileinline = item.attachments[i].isInline;

                    if (attachmentType == "item") {
                        console.log(`Index.razor.js(getEmailData): Only read file type attachments`);
                        continue;
                    }

                    console.log(`Index.razor.js(getEmailData) Type: ` + attachmentType);

                    var fileExtension = "";

                    try {
                        fileExtension = getFileExtension(attachmentName);
                    } catch (e) {
                        console.log(`Index.razor.js(getEmailData): Unable to identify attachment file extension ...`);
                    }

                    if (fileExtension !== "") {
                        console.log(`Index.razor.js(getEmailData): Processing file extension '${fileExtension}'...`);
                        console.log("Attachment Value: " + returnedAttachments[i].value.content);

                        //await item.getAttachmentContentAsync(attachmentID, handleAttachmentsCallback);

                        return {
                            AttachmentId: item.attachments[i].id,
                            AttachmentName: attachmentName,
                            Subject: item.subject,
                            AttachmentType: attachmentType, 
                            Inline: fileinline,
                            AttachmentBase64Data: returnedAttachments[i].value.content
                        };
                    } else {
                        console.error(`Index.razor.js(getEmailData): Could not parse file extension for ${attachmentName}`);
                        continue;
                    }
                }
            }
            catch (e) {
                console.log("----------------------------------------------- never executed", e);
            }
        }
        else {
            return {
                AttachmentId: "",
                AttachmentName: "No Attachments",
                Subject: item.subject,
                AttachmentType: "",
                Inline: false,
                AttachmentBase64Data: ""
            };
        }
    } catch (err) {
        console.error(`Index.razor.js(getEmailData) Catch Exception: ${err}`);
        subject = `${err}`;
        return { Subject: subject };
    }
}

function getFileExtension(fileName) {
    var a = fileName.split(".");
    if (a.length === 1 || (a[0] === "" && a.length === 2)) {
        return "";
    }
    return a.pop();
}

function getAttachments(item) {

    // getAttachmentContentAsync() is only supported above 1.8
    if (!Office.context.requirements.isSetSupported("Mailbox", "1.8")) {
        return [];
    }

    const attachments = item.attachments;

    // If you need to filter, uncomment...
    // -----------------------------------
    //    = item.attachments.filter(
    //    (attachment) => attachment.attachmentType === "file" && attachment.isInline === true
    //);

    console.log("Filtered attachments:", attachments.length);

    if (!attachments) {
        return [];
    }

    return Promise.all(
        attachments.map(
            (attachment) =>
                new Promise((resolve) => {
                    console.log(`Index.razor.js (getAttachments) attachment ID returned: ` + attachment.id);
                    item.getAttachmentContentAsync(attachment.id, (result) => resolve(result));
                })
        )
    );
}

