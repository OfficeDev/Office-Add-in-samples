// Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.

import {
  addCustomerControls,
  addParagraphs,
  changeCustomer,
  changeLogo,
  highlightCurrency,
  insertCosts,
  insertFooter,
  insertTeam,
  insertTemplate,
} from "../assembly/document-assembly";
import { calculateTeam, currency, members, roles, TeamMember } from "../assembly/team";

function input(id: string): HTMLInputElement {
  const element = document.getElementById(id);
  if (!(element instanceof HTMLInputElement)) {
    throw new Error(`Missing input: ${id}`);
  }
  return element;
}

function showStatus(message: string, error = false): void {
  const status = document.getElementById("status");
  if (!status) {
    throw new Error("Missing status element.");
  }
  status.textContent = message;
  status.classList.toggle("error", error);
}

function reportError(error: unknown): void {
  console.error(error);
  showStatus(error instanceof Error ? error.message : String(error), true);
}

const selections: HTMLSelectElement[] = [];

function getTeam(): TeamMember[] {
  return selections.map((select) => {
    const member = members[Number(select.value)];
    if (!member) {
      throw new Error("Choose a valid consultant for every role.");
    }
    return member;
  });
}

function updateTotal(): void {
  const output = document.getElementById("team-total");
  if (!output) {
    throw new Error("Missing team total.");
  }
  try {
    output.textContent = currency(calculateTeam(getTeam(), input("hours").valueAsNumber).total);
  } catch (error) {
    output.textContent = error instanceof Error ? error.message : String(error);
  }
}

function initialize(): void {
  const actions = document.getElementById("actions");
  const fields = document.getElementById("team-fields");
  if (!(actions instanceof HTMLFieldSetElement) || !fields) {
    throw new Error("Missing task-pane controls.");
  }
  const actionFieldset = actions;
  roles.forEach((role, index) => {
    const label = document.createElement("label");
    const select = document.createElement("select");
    select.id = `role-${index}`;
    label.htmlFor = select.id;
    label.textContent = role;
    members.forEach((member, memberIndex) => {
      select.add(new Option(`${member.name} (${currency(member.rate)}/hour)`, String(memberIndex)));
    });
    select.value = String(index);
    select.addEventListener("change", updateTotal);
    selections.push(select);
    fields.append(label, select);
  });
  input("hours").addEventListener("input", updateTotal);
  updateTotal();

  function bind(id: string, action: () => Promise<string>): void {
    const button = document.getElementById(id);
    if (!(button instanceof HTMLButtonElement)) {
      throw new Error(`Missing button: ${id}`);
    }
    button.addEventListener("click", async () => {
      actionFieldset.disabled = true;
      showStatus("Working...");
      try {
        showStatus(await action());
      } catch (error) {
        reportError(error);
      } finally {
        actionFieldset.disabled = false;
      }
    });
  }

  bind("insert-template", async () => {
    const confirmation = input("replace-document");
    if (!confirmation.checked) {
      throw new Error(
        "Confirm that replacing the document body is OK before inserting the template."
      );
    }
    const message = await insertTemplate();
    confirmation.checked = false;
    return message;
  });
  bind("add-paragraphs", addParagraphs);
  bind("change-logo", changeLogo);
  bind("add-customer-controls", addCustomerControls);
  bind("change-customer", () => changeCustomer(input("customer-name").value));
  bind("insert-costs", insertCosts);
  bind("highlight-currency", highlightCurrency);
  bind("insert-footer", insertFooter);
  bind("insert-team", () => insertTeam(getTeam(), input("hours").valueAsNumber));
  actions.disabled = false;
  showStatus("Ready. Use a disposable document and start with step 1.");
}

Office.onReady((info) => {
  if (
    info.host !== Office.HostType.Word ||
    !Office.context.requirements.isSetSupported("WordApi", "1.3")
  ) {
    showStatus("This sample requires Word with WordApi 1.3 or later.", true);
    return;
  }
  initialize();
}).catch(reportError);
