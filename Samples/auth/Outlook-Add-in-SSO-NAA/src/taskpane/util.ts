// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* global window */

export function createLocalUrl(path: string) {
  return `${window.location.origin}/${path}`;
}
