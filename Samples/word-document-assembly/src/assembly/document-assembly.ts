// Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.

import { createTeamOoxml, TeamMember } from "./team";

function clearHighlight(font: Word.Font): void {
  // @ts-expect-error Office.js documents null for removing highlights, but its typings declare only string.
  font.highlightColor = null;
}

async function getAsset(name: string): Promise<Response> {
  const response = await fetch(new URL(`assets/${name}`, window.location.href));
  if (!response.ok) {
    throw new Error(`Could not load ${name}: HTTP ${response.status}.`);
  }
  return response;
}

async function getBase64Asset(name: string): Promise<string> {
  const response = await getAsset(name);
  const bytes = new Uint8Array(await response.arrayBuffer());
  let binary = "";
  for (const byte of bytes) {
    binary += String.fromCharCode(byte);
  }
  return btoa(binary);
}

export async function insertTemplate(): Promise<string> {
  const [ooxml, logo] = await Promise.all([
    getAsset("statement-of-work.xml").then((response) => response.text()),
    getBase64Asset("icon-80.png"),
  ]);
  await Word.run(async (context) => {
    const body = context.document.body;
    const currentDate = new Date().toLocaleDateString("en-US", {
      year: "numeric",
      month: "long",
      day: "numeric",
    });
    body.insertOoxml(ooxml.replace("July 23, 2015", currentDate), Word.InsertLocation.replace);
    const picture = body.insertInlinePictureFromBase64(logo, Word.InsertLocation.start);
    picture.altTextDescription = "Sample company logo";
    picture.width = 48;
    picture.height = 48;
    await context.sync();
  });
  return "Statement of Work template inserted. Continue with step 2.";
}

export async function addParagraphs(): Promise<string> {
  await Word.run(async (context) => {
    const controls = context.document.contentControls.getByTitle("ProjectCosts");
    controls.load("id");
    await context.sync();
    if (controls.items.length) {
      throw new Error("Boilerplate already exists. Insert the template again to start over.");
    }
    const body = context.document.body;
    body.insertParagraph("Timeline", Word.InsertLocation.end).styleBuiltIn =
      Word.BuiltInStyleName.heading2;
    body.insertParagraph(
      "Services begin upon approval of this illustrative Statement of Work and continue for the agreed project duration.",
      Word.InsertLocation.end
    ).styleBuiltIn = Word.BuiltInStyleName.normal;
    body.insertParagraph("Project Costs by Phase", Word.InsertLocation.end).styleBuiltIn =
      Word.BuiltInStyleName.heading2;
    const costs = body.insertParagraph("<Add Project Costs Here>", Word.InsertLocation.end);
    costs.styleBuiltIn = Word.BuiltInStyleName.normal;
    costs.font.highlightColor = "#FFFF00";
    const control = costs.insertContentControl();
    control.title = "ProjectCosts";
    control.tag = "project-costs";
    body.insertParagraph("Project Team", Word.InsertLocation.end).styleBuiltIn =
      Word.BuiltInStyleName.heading2;
    body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
    body.insertParagraph("Terms of Work", Word.InsertLocation.end).styleBuiltIn =
      Word.BuiltInStyleName.heading1;
    body.insertParagraph(
      "Contractor shall provide the agreed services and deliverables. This document contains synthetic demo terms, not legal advice.",
      Word.InsertLocation.end
    ).styleBuiltIn = Word.BuiltInStyleName.normal;
    body.insertParagraph(
      "Out-of-Pocket Expenses / Invoice Procedures",
      Word.InsertLocation.end
    ).styleBuiltIn = Word.BuiltInStyleName.heading2;
    body.insertParagraph(
      "Client will be invoiced monthly for approved services and expenses. Each invoice will include supporting time sheets and expense receipts.",
      Word.InsertLocation.end
    ).styleBuiltIn = Word.BuiltInStyleName.normal;
    await context.sync();
  });
  return "Boilerplate, a ProjectCosts content control, and a page break added.";
}

export async function changeLogo(): Promise<string> {
  const logo = await getBase64Asset("replacement-logo.png");
  await Word.run(async (context) => {
    const pictures = context.document.body.inlinePictures;
    pictures.load("width");
    await context.sync();
    if (!pictures.items.length) {
      throw new Error("Insert the template before replacing its logo.");
    }
    const picture = pictures.items[0].insertInlinePictureFromBase64(
      logo,
      Word.InsertLocation.replace
    );
    picture.altTextDescription = "Updated sample company logo";
    picture.width = 48;
    picture.height = 48;
    await context.sync();
  });
  return "The first inline picture was replaced with the bundled logo.";
}

export async function addCustomerControls(): Promise<string> {
  return Word.run(async (context) => {
    const existing = context.document.contentControls.getByTag("customer");
    existing.load("id");
    const results = context.document.body.search("Contoso", {
      matchCase: true,
      matchWholeWord: true,
    });
    results.load("text");
    await context.sync();
    if (existing.items.length) {
      throw new Error("Customer controls already exist. Use Change Customer to update them.");
    }
    if (!results.items.length) {
      throw new Error('No "Contoso" text found. Insert the template first.');
    }
    for (const range of results.items) {
      range.font.color = "#FF0000";
      range.font.highlightColor = "#FFFF00";
      range.font.bold = true;
      const control = range.insertContentControl();
      control.tag = "customer";
      control.title = "Customer Name";
    }
    await context.sync();
    return `Created ${results.items.length} customer content controls.`;
  });
}

export async function changeCustomer(name: string): Promise<string> {
  const customer = name.trim();
  if (!customer || customer.length > 100) {
    throw new Error("Enter a customer name between 1 and 100 characters.");
  }
  return Word.run(async (context) => {
    const controls = context.document.contentControls.getByTag("customer");
    controls.load("id");
    await context.sync();
    if (!controls.items.length) {
      throw new Error("Add customer content controls with step 4 first.");
    }
    for (const control of controls.items) {
      control.insertText(customer, Word.InsertLocation.replace);
      clearHighlight(control.font);
    }
    await context.sync();
    return `Updated ${controls.items.length} customer content controls to ${customer}.`;
  });
}

export async function insertCosts(): Promise<string> {
  const base64 = await getBase64Asset("project-costs.docx");
  return Word.run(async (context) => {
    const controls = context.document.contentControls.getByTitle("ProjectCosts");
    controls.load("id");
    await context.sync();
    if (!controls.items.length) {
      throw new Error("Add the ProjectCosts content control with step 2 first.");
    }
    for (const control of controls.items) {
      control.insertFileFromBase64(base64, Word.InsertLocation.replace);
      clearHighlight(control.font);
    }
    await context.sync();
    return `Inserted the project-costs document into ${controls.items.length} content control(s).`;
  });
}

export async function highlightCurrency(): Promise<string> {
  return Word.run(async (context) => {
    // Restrict matches to digits and grouping commas instead of spanning intervening prose with "*".
    const results = context.document.body.search("$[0-9,]@.[0-9]{2}", { matchWildcards: true });
    results.load("text");
    await context.sync();
    if (!results.items.length) {
      throw new Error("No USD amounts found. Insert project costs or a team table first.");
    }
    for (const range of results.items) {
      range.font.color = "#FF0000";
      range.font.highlightColor = "#FFFF00";
      range.font.bold = true;
    }
    results.items[results.items.length - 1].select();
    await context.sync();
    return `Highlighted ${results.items.length} USD amounts and selected the last match.`;
  });
}

export async function insertFooter(): Promise<string> {
  await Word.run(async (context) => {
    const footer = context.document.sections.getFirst().getFooter(Word.HeaderFooterType.primary);
    footer.insertText("Confidential", Word.InsertLocation.replace);
    await context.sync();
  });
  return 'The first section\'s primary footer now reads "Confidential".';
}

export async function insertTeam(team: readonly TeamMember[], hours: number): Promise<string> {
  const ooxml = createTeamOoxml(team, hours);
  await new Promise<void>((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      ooxml,
      { coercionType: Office.CoercionType.Ooxml },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(result.error.message));
        } else {
          resolve();
        }
      }
    );
  });
  return "Project team table inserted at the current selection.";
}
