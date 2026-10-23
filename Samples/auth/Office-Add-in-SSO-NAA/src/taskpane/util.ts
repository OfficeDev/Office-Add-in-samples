// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* global window */

/**
 * Constructs a local URL for the web page for the given path.
 * @param path The path to construct a local URL for.
 * @returns
 */
export function createLocalUrl(path: string) {
  return `${window.location.origin}/${path}`;
}
