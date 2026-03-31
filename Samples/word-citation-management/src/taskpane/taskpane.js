/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

const bibtexParser = require("@orcid/bibtex-parse-js");
let bibFileContent;

Office.onReady((info) => {
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", function () {
      initializeApp(info);
    });
  } else {
    initializeApp(info);
  }
});

function initializeApp(info) {
  if (info.host === Office.HostType.Word) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("bib-file").addEventListener("change", () => tryCatch(getFileContents));
    search();
    document.getElementById("insert-citation").addEventListener("click", () => tryCatch(insertCitation));
    document.getElementById("clear").addEventListener("click", () => tryCatch(clearSelection));
  }
}

// Gets the contents of the selected file.
async function getFileContents() {
  const myBibFile = document.getElementById("bib-file");
  const reader = new FileReader();
  reader.onloadend = function () {
    bibFileContent = reader.result;
    populateCitationsFromFile();
    showReferencesSection();
  };
  reader.readAsText(myBibFile.files[0]);
}

// Searches the references list for the search text.
async function search() {
  const searchElement = document.getElementById("search");
  const radioButtons = document.getElementById("radio-buttons");
  
  searchElement.addEventListener("search", searchHandler);
  searchElement.addEventListener("keyup", searchHandler);
  
  function searchHandler() {
    const searchText = searchElement.value;
    const children = radioButtons.children;
    
    if (searchText) {
      Array.from(children).forEach(function (child) {
        if (child.textContent.search(new RegExp(searchText, "i")) < 0) {
          child.style.display = "none";
        } else {
          child.style.display = "block";
        }
      });
    } else {
      Array.from(children).forEach(function (child) {
        child.style.display = "block";
      });
    }
    
    // Trigger change event
    radioButtons.dispatchEvent(new Event('change', { bubbles: true }));
  }
}

// Shows the reference section.
async function showReferencesSection() {
  const referenceSection = document.getElementById("references-section");
  referenceSection.style.display = "block";
  referenceSection.dispatchEvent(new Event('change', { bubbles: true }));
}

// Populates the radio buttons with the citations from the file.
async function populateCitationsFromFile() {
    let citationsFromFile = bibtexParser.toJSON(bibFileContent);
    console.log(citationsFromFile);

    const populateRadio = document.getElementById("populate-radio");
    const radioButtons = document.getElementById("radio-buttons");
    radioButtons.innerHTML = "";
    
    for (let citation in citationsFromFile) {
      let citationHtml = `<section><input type="radio" id="${citationsFromFile[citation].citationKey}" name="citation" value='${citationsFromFile[citation].entryTags.author}, "${citationsFromFile[citation].entryTags.title}"'>
      <label for="${citationsFromFile[citation].citationKey}"><b>${citationsFromFile[citation].entryTags.title}</b><br>${citationsFromFile[citation].entryTags.author}</label><br><br></section>`;
      radioButtons.insertAdjacentHTML('beforeend', citationHtml);
    }
    populateRadio.appendChild(radioButtons);

    // Add event listeners to all radio buttons.
    const citationRadios = document.querySelectorAll("input[name='citation'][type='radio']");
    citationRadios.forEach(radio => {
      radio.addEventListener("click", function () {
        if (this.checked) {
          setSelected(`Current selection: ${this.id}`);
          enableButtons();
        } else {
          clearSelected();
          disableButtons();
        }
      });
    });
    
    populateRadio.dispatchEvent(new Event('change', { bubbles: true }));
}

// Inserts the citation at the cursor location in the document.
async function insertCitation() {
  await Word.run(async (context) => {
    const checkedRadio = document.querySelector("input[name='citation'][type='radio']:checked");
    const citation = checkedRadio ? checkedRadio.value : "";
    context.document.getSelection().insertEndnote(citation);
    await context.sync();

    console.log(`Inserted citation: ${citation}`);
  });
}

// Clears the selected radio button.
async function clearSelection() {
  const checkedRadio = document.querySelector("input[name='citation'][type='radio']:checked");
  if (checkedRadio) {
    checkedRadio.checked = false;
  }
  clearSelected();
  disableButtons();
}

// Sets the selected item.
async function setSelected(text) {
  document.getElementById("selected").textContent = text;
}

// Clears the selected item.
async function clearSelected() {
  document.getElementById("selected").textContent = "";
}

// Enables the buttons.
async function enableButtons() {
  const buttons = document.querySelectorAll(".ms-Button");
  buttons.forEach(button => {
    button.removeAttribute("disabled");
  });
}

// Disables the buttons.
async function disableButtons() {
  const buttons = document.querySelectorAll(".ms-Button");
  buttons.forEach(button => {
    button.setAttribute("disabled", "disabled");
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
