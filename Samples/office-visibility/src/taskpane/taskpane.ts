/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import "../commands/commands";

/* global Office console document HTMLElement HTMLButtonElement HTMLSelectElement */

declare global {
  namespace Office {
    interface Control {
      visible?: boolean;
    }

    interface Group {
      visible?: boolean;
    }
  }
}

const TAB_ID = "ContosoVisibilityTab";
const PRIMARY_GROUP_ID = "PrimaryGroup";
const SECONDARY_GROUP_ID = "SecondaryGroup";

interface RibbonControlDefinition {
  id: string;
  label: string;
  groupId: string;
  initiallyVisible: boolean;
}

interface RibbonGroupDefinition {
  id: string;
  label: string;
  initiallyVisible: boolean;
}

const controls: readonly RibbonControlDefinition[] = [
  {
    id: "Button1",
    label: "Button 1",
    groupId: PRIMARY_GROUP_ID,
    initiallyVisible: true,
  },
  {
    id: "Button2",
    label: "Button 2",
    groupId: PRIMARY_GROUP_ID,
    initiallyVisible: false,
  },
  {
    id: "SampleMenu",
    label: "Sample menu",
    groupId: PRIMARY_GROUP_ID,
    initiallyVisible: true,
  },
  {
    id: "SecondaryButtonOne",
    label: "Secondary button one",
    groupId: SECONDARY_GROUP_ID,
    initiallyVisible: true,
  },
  {
    id: "SecondaryButtonTwo",
    label: "Secondary button two",
    groupId: SECONDARY_GROUP_ID,
    initiallyVisible: true,
  },
];

const groups: readonly RibbonGroupDefinition[] = [
  {
    id: PRIMARY_GROUP_ID,
    label: "Primary commands",
    initiallyVisible: true,
  },
  {
    id: SECONDARY_GROUP_ID,
    label: "Secondary commands",
    initiallyVisible: false,
  },
];

const selectableControls = controls;

function getElement(id: string): HTMLElement {
  const element = document.getElementById(id);
  if (!element) {
    throw new Error(`The task pane element "${id}" wasn't found.`);
  }

  return element;
}

function getButton(id: string): HTMLButtonElement {
  const element = getElement(id);
  if (!(element instanceof HTMLButtonElement)) {
    throw new Error(`The task pane element "${id}" isn't a button.`);
  }

  return element;
}

function getSelect(id: string): HTMLSelectElement {
  const element = getElement(id);
  if (!(element instanceof HTMLSelectElement)) {
    throw new Error(`The task pane element "${id}" isn't a select element.`);
  }

  return element;
}

function setStatus(message: string, kind: "info" | "success" | "error" = "info") {
  const status = getElement("status");
  status.textContent = message;
  status.className = kind === "info" ? "status" : `status ${kind}`;
}

function log(message: string, isError = false) {
  if (isError) {
    console.error(message);
  } else {
    console.log(message);
  }

  const entry = document.createElement("div");
  entry.textContent = message;
  if (isError) {
    entry.className = "log-error";
  }

  const logElement = getElement("log");
  logElement.appendChild(entry);
  logElement.scrollTop = logElement.scrollHeight;
}

function formatError(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function setRibbonButtonsDisabled(disabled: boolean) {
  const buttons = document.querySelectorAll<HTMLButtonElement>("button[data-ribbon-action]");
  buttons.forEach((button) => {
    button.disabled = disabled;
  });
}

function populateSelect(
  select: HTMLSelectElement,
  definitions: readonly { id: string; label: string; initiallyVisible: boolean }[]
) {
  for (const definition of definitions) {
    const option = document.createElement("option");
    option.value = definition.id;
    option.textContent = `${definition.label} (initially ${definition.initiallyVisible ? "visible" : "hidden"})`;
    select.appendChild(option);
  }
}

function getSelectedControl(): RibbonControlDefinition {
  const selectedId = getSelect("control-select").value;
  const selectedControl = controls.find((control) => control.id === selectedId);
  if (!selectedControl) {
    throw new Error(`Unknown ribbon control ID: ${selectedId}`);
  }

  return selectedControl;
}

function getSelectedGroup(): RibbonGroupDefinition {
  const selectedId = getSelect("group-select").value;
  const selectedGroup = groups.find((group) => group.id === selectedId);
  if (!selectedGroup) {
    throw new Error(`Unknown ribbon group ID: ${selectedId}`);
  }

  return selectedGroup;
}

async function requestRibbonUpdate(label: string, update: Office.RibbonUpdaterData) {
  log(`${label}: ${JSON.stringify(update)}`);
  await Office.ribbon.requestUpdate(update);
}

async function setControlVisibility(control: RibbonControlDefinition, visible: boolean) {
  await requestRibbonUpdate(`${visible ? "Showing" : "Hiding"} ${control.label}`, {
    tabs: [
      {
        id: TAB_ID,
        groups: [
          {
            id: control.groupId,
            controls: [
              {
                id: control.id,
                visible,
              },
            ],
          },
        ],
      },
    ],
  });
}

async function setGroupVisibility(group: RibbonGroupDefinition, visible: boolean) {
  await requestRibbonUpdate(`${visible ? "Showing" : "Hiding"} ${group.label}`, {
    tabs: [
      {
        id: TAB_ID,
        groups: [
          {
            id: group.id,
            visible,
          },
        ],
      },
    ],
  });
}

function createGroupUpdate(
  group: RibbonGroupDefinition,
  groupVisible: boolean,
  getControlVisibility: (control: RibbonControlDefinition) => boolean
): Office.Group {
  return {
    id: group.id,
    visible: groupVisible,
    controls: controls
      .filter((control) => control.groupId === group.id)
      .map((control) => ({
        id: control.id,
        visible: getControlVisibility(control),
      })),
  };
}

async function setAllSampleContentVisibility(visible: boolean) {
  await requestRibbonUpdate(`${visible ? "Showing" : "Hiding"} all sample content`, {
    tabs: [
      {
        id: TAB_ID,
        groups: groups.map((group) => createGroupUpdate(group, visible, () => visible)),
      },
    ],
  });
}

async function restoreManifestDefaults() {
  await requestRibbonUpdate("Restoring manifest defaults", {
    tabs: [
      {
        id: TAB_ID,
        groups: groups.map((group) =>
          createGroupUpdate(group, group.initiallyVisible, (control) => control.initiallyVisible)
        ),
      },
    ],
  });
}

async function runOperation(description: string, operation: () => Promise<void>) {
  setRibbonButtonsDisabled(true);
  setStatus(description);

  try {
    await operation();
    setStatus(`${description} completed.`, "success");
    log(`${description} completed.`);
  } catch (error) {
    const message = formatError(error);
    setStatus(message, "error");
    log(`Error: ${message}`, true);
  } finally {
    setRibbonButtonsDisabled(false);
  }
}

function wireTaskPane() {
  getButton("show-control").addEventListener("click", () => {
    void runOperation("Show control", () => setControlVisibility(getSelectedControl(), true));
  });
  getButton("hide-control").addEventListener("click", () => {
    void runOperation("Hide control", () => setControlVisibility(getSelectedControl(), false));
  });
  getButton("show-group").addEventListener("click", () => {
    void runOperation("Show group", () => setGroupVisibility(getSelectedGroup(), true));
  });
  getButton("hide-group").addEventListener("click", () => {
    void runOperation("Hide group", () => setGroupVisibility(getSelectedGroup(), false));
  });
  getButton("show-all").addEventListener("click", () => {
    void runOperation("Show all sample content", () => setAllSampleContentVisibility(true));
  });
  getButton("hide-all").addEventListener("click", () => {
    void runOperation("Hide all sample content", () => setAllSampleContentVisibility(false));
  });
  getButton("restore-defaults").addEventListener("click", () => {
    void runOperation("Restore manifest defaults", restoreManifestDefaults);
  });
  getButton("clear-log").addEventListener("click", () => {
    getElement("log").replaceChildren();
  });
}

Office.onReady((info) => {
  populateSelect(getSelect("control-select"), selectableControls);
  populateSelect(getSelect("group-select"), groups);
  wireTaskPane();

  setRibbonButtonsDisabled(false);

  const host = info.host ? String(info.host) : "Office";

  setStatus(`Ready to test ribbon visibility in ${host}.`);
  log(`Ready in ${host}.`);
});
