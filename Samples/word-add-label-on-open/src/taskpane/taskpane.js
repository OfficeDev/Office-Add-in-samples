/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("nonBusiness").onclick = changeToNonBusinessLevel;
    document.getElementById("public").onclick = changeToPublicLevel;
    document.getElementById("general").onclick = changeToGeneralLevel;
    document.getElementById("confidential").onclick = changeToConfidentialLevel;
    document.getElementById("highConfidential").onclick = changeToHighConfidentialLevel;
  }
});

export async function changeToNonBusinessLevel() {
  return Word.run(async (context) => {
    const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
    const firstPageHeader = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.firstPage);

    header.clear();
    firstPageHeader.clear();
    header.insertParagraph("Non-Business - The data is personal and not business related", "Start");
    firstPageHeader.insertParagraph("Non-Business - The data is personal and not business related", "Start");
    header.font.color = "#737173";
    firstPageHeader.font.color = "#737173";

    await context.sync();
  });
}

export async function changeToPublicLevel() {
  return Word.run(async (context) => {
    const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
    const firstPageHeader = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.firstPage);
    header.clear();
    firstPageHeader.clear();
    header.insertParagraph("Public - The data is for the public and shareable externally", "Start");
    firstPageHeader.insertParagraph("Public - The data is for the public and shareable externally", "Start");
    header.font.color = "#07641d";
    firstPageHeader.font.color = "#07641d";

    await context.sync();
  });
}

export async function changeToGeneralLevel() {
  return Word.run(async (context) => {
    const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
    const firstPageHeader = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.firstPage);
    header.clear();
    firstPageHeader.clear();
    header.insertParagraph("General - Business data shared with trusted individuals", "Start");
    firstPageHeader.insertParagraph("General - Business data shared with trusted individuals", "Start");
    header.font.color = "#0177d3";
    firstPageHeader.font.color = "#0177d3";

    await context.sync();
  });
}

export async function changeToConfidentialLevel() {
  return Word.run(async (context) => {
    const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
    const firstPageHeader = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.firstPage);
    header.clear();
    firstPageHeader.clear();
    header.insertParagraph("Confidential - Sensitive business data shared with trusted individuals", "Start");
    firstPageHeader.insertParagraph("Confidential - Sensitive business data shared with trusted individuals", "Start");
    header.font.color = "#ff5c3a";
    firstPageHeader.font.color = "#ff5c3a";

    await context.sync();
  });
}

export async function changeToHighConfidentialLevel() {
  return Word.run(async (context) => {
    const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
    const firstPageHeader = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.firstPage);
    header.clear();
    firstPageHeader.clear();
    header.insertParagraph("Highly Confidential - The data must be secret or in some way highly critical", "Start");
    firstPageHeader.insertParagraph(
      "Highly Confidential - The data must be secret or in some way highly critical",
      "Start"
    );
    header.font.color = "#f8334d";
    firstPageHeader.font.color = "#f8334d";

    await context.sync();
  });
}
