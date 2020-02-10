import { SignApp, signInO365 } from '../../utilities/office-apis-helpers';


function getGlobal() {
    console.log('init globals for command buttons');
    return typeof self !== "undefined"
        ? self
        : typeof window !== "undefined"
            ? window
            : typeof global !== "undefined"
                ? global
                : undefined;
}

const g = getGlobal() as any;


// the add-in command functions need to be available in global scope
g.btnsignin = btnSignIn;

export function btnSignIn(event: Office.AddinCommands.Event) {
    console.log('sign in button pressed');
    // Your code goes here

    let signapp = new SignApp();
    signInO365(signapp.setState, signapp.setToken, signapp.displayError);
    //SetRuntimeVisibleHelper(true);
    // Be sure to indicate when the add-in command function is complete
    event.completed();
}