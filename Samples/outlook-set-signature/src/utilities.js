// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

function is_valid_data(str) {
  return str !== null && str !== undefined && str !== "";
}

function get_cal_offset() {
  return "<br/><br/>";
}

// Expose for cross-module access in webpack bundles
window.is_valid_data = is_valid_data;
window.get_cal_offset = get_cal_offset;
