// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* global window, document */

/**
 * Constructs a local URL for the web page for the given path.
 * @param path The path to construct a local URL for.
 * @returns
 */
export function createLocalUrl(path: string) {
  return `${window.location.origin}/${path}`;
}

/**
 * Makes the Sign out button visible or invisible on the task pane.
 *
 * @param visible true if the sign out button should be visible; otherwise, false.
 * @returns
 */
export function setSignOutButtonVisibility(visible: boolean) {
  const signOutButton = document.getElementById("signOutButton");
  if (!signOutButton) return;
  if (visible) {
    signOutButton.classList.remove("is-disabled");
  } else {
    signOutButton.classList.add("is-disabled");
  }
}
