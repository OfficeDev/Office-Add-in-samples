/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

const bibtexParser = require("@orcid/bibtex-parse-js");
let bibFileContent;

Office.onReady((info) => {
  $(document).ready(function () {
    if (info.host === Office.HostType.Word) {
      document.getElementById("app-body").style.display = "flex";
      $("#bib-file").change(() => tryCatch(getFileContents));
      search();
      $("#insert-citation").click(() => tryCatch(insertCitation));
      $("#clear").click(() => tryCatch(clearSelection));
    }
  });
});

// Gets the contents of the selected file.
async function getFileContents() {
  const myBibFile = document.getElementById("bib-file");
  const reader = new FileReader();
  reader.onloadend = function () {
    bibFileContent = reader.result;
    populateCitationsFromFile();
    showReferencesSection();
  };
  reader.readAsBinaryString(myBibFile.files[0]);
}

// Searches the references list for the search text.
async function search() {
  let $search = $("#search");
  let $radioButtons = $("#radio-buttons");
  $search.on("search keyup", function () {
    let searchText = $(this).val();
    if (searchText) {
      $radioButtons.children().each(function () {
        let $this = $(this);
        if ($this.text().search(new RegExp(searchText, "i")) < 0) {
          $this.hide();
        } else {
          $this.show();
        }
      });
    } else {
      $("#radio-buttons").children().each(function () {
        $(this).show();
      });
    }
    $radioButtons.change();
  });
}

// Shows the reference section.
async function showReferencesSection() {
  let $referenceSection = $("#references-section");
  $referenceSection.show();
  $referenceSection.change();
}

// Populates the radio buttons with the citations from the file.
async function populateCitationsFromFile() {
    let citationsFromFile = bibtexParser.toJSON(bibFileContent);
    console.log(citationsFromFile);

    let $populateRadio = $("#populate-radio");
    let $radioButtons = $("#radio-buttons");
    $radioButtons.empty();
    for (let citation in citationsFromFile) {
      let citationHtml = `<section><input type="radio" id="${citationsFromFile[citation].citationKey}" name="citation" value='${citationsFromFile[citation].entryTags.author}, "${citationsFromFile[citation].entryTags.title}"'>
      <label for="${citationsFromFile[citation].citationKey}"><b>${citationsFromFile[citation].entryTags.title}</b><br>${citationsFromFile[citation].entryTags.author}</label><br><br></section>`;
      $radioButtons.append(citationHtml);
    }
    $radioButtons.appendTo($populateRadio);

    $("input[name='citation'][type='radio']").on("click", function () {
      if ($(this).prop("checked")) {
        setSelected(`Current selection: ${$(this).prop("id")}`);
        enableButtons();
      } else {
        clearSelected();
        disableButtons();
      }
    });
    $populateRadio.change();
}

// Inserts the citation at the cursor location in the document.
async function insertCitation() {
  await Word.run(async (context) => {
    const citation = $("input[name='citation'][type='radio']:checked").val();
    context.document.getSelection().insertEndnote(citation);
    await context.sync();

    console.log(`Inserted citation: ${citation}`);
  });
}

// Clears the selected radio button.
async function clearSelection() {
  $("input[name='citation'][type='radio']:checked").prop("checked", false);
  clearSelected();
  disableButtons();
}

// Sets the selected item.
async function setSelected(text) {
  $("#selected").text(text);
}

// Clears the selected item.
async function clearSelected() {
  $("#selected").text("");
}

// Enables the buttons.
async function enableButtons() {
  $(".ms-Button").removeAttr("disabled");
}

// Disables the buttons.
async function disableButtons() {
  $(".ms-Button").attr("disabled", "disabled");
}

// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.log(error);
  }
}
