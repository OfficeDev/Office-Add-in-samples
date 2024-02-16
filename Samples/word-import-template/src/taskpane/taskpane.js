/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

let template;

Office.onReady((info) => {
  $(document).ready(function () {
    if (info.host === Office.HostType.Word) {
      document.getElementById("app-body").style.display = "flex";
      $("#template-file").on("change", () => tryCatch(getFileContents));
      $("#import-template").on("click", () => tryCatch(importTemplate));
      $("#save-close").on("click", () => tryCatch(saveAndCloseFile));
    }
  });
});

// Gets the contents of the selected file.
async function getFileContents() {
  const myTemplate = document.getElementById("template-file");
  const reader = new FileReader();
  reader.onload = (event) => {
    // Remove the metadata before the Base64-encoded string.
    const startIndex = reader.result.toString().indexOf("base64,");
    template = reader.result.toString().substring(startIndex + 7);

    // Show the Import and Save sections.
    $("#import-section").show();
  };

  // Read the file as a data URL so we can parse the Base64-encoded string.
  reader.readAsDataURL(myTemplate.files[0]);
}

// Imports the template into this document.
async function importTemplate() {
  await Word.run(async (context) => {
    // Use the Base64-encoded string representation of the selected .docx file.
    context.document.insertFileFromBase64(template, "Replace", {
      importTheme: true,
      importStyles: true,
      importParagraphSpacing: true,
      importPageColor: true,
      importDifferentOddEvenPages: true
    });
    await context.sync();
  });
}

// Displays the save dialog if the file hasn't already been saved, then closes the file.
async function saveAndCloseFile() {
  await Word.run(async (context) => {
    context.document.save(Word.SaveBehavior.prompt);
    await context.sync();

    context.document.close();
    await context.sync();
  });
}

// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.log(error);
  }
}