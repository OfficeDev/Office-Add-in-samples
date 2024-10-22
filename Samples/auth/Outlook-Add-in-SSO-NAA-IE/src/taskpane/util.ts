/* global window, Office */

export function createLocalUrl(path: string) {
  return `${window.location.origin}/${path}`;
}

export function isInternetExplorer() {
  return /MSIE|Trident/.test(window.navigator.userAgent);
}

export async function sendDialogMessage(message: string) {
  await Office.onReady();
  Office.context.ui.messageParent(message);
}

export function shouldCloseDialog() {
  return window.location.search.indexOf("close=1") !== -1;
}

export function getCurrentPageUrl(queryParams?: { [key: string]: string }) {
  let querystring = "";
  for (const key in queryParams) {
    if (Object.prototype.hasOwnProperty.call(queryParams, key)) {
      if (!querystring) {
        querystring += "?";
      }
      querystring += `${key}=${queryParams[key]}&`;
    }
  }
  return window.location.origin + window.location.pathname + querystring;
}
