import { showMessage } from "./message-helper";

export function handleClientSideErrors(error: any): boolean {
  let invokeFallBackDialog: boolean = false;
  switch (error.code) {
    case 13001:
      // No one is signed into Office. If the add-in cannot be effectively used when no one
      // is logged into Office, then the first call of getAccessToken should pass the
      // `allowSignInPrompt: true` option.
      showMessage(
        "No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again."
      );
      return invokeFallBackDialog;
    case 13002:
      // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
      // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
      showMessage(
        "You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."
      );
      return invokeFallBackDialog;
    case 13006:
      // Only seen in Office on the Web.
      showMessage(
        "Office on the Web is experiencing a problem. Please sign out of Office, close the browser, and then start again."
      );
      return invokeFallBackDialog;
    case 13008:
      // Only seen in Office on the Web.
      showMessage("Office is still working on the last operation. When it completes, try this operation again.");
      return invokeFallBackDialog;
    case 13010:
      // Only seen in Office on the Web.
      showMessage("Follow the instructions to change your browser's zone configuration.");
      return invokeFallBackDialog;
    default:
      // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, fall back
      // to non-SSO sign-in.
      invokeFallBackDialog = true;
      return invokeFallBackDialog;
  }
}
