import {AppState} from '../src/components/app';
import {signInO365} from '../utilities/office-apis-helpers';

// sign in commands (without task pane)

class SignApp {
    appstate: AppState;
    accessToken: string;

   setToken = (accesstoken: string) => {
    this.accessToken = accesstoken;
   }

   setState = (nState: AppState) => {
       this.appstate = nState;
   }

   displayError = (error: string) => {
    this.setState({ errorMessage: error });
   }
}

export function appCmdSignIn(){
    console.log('starting');
    signincallme();
}

async function signincallme()
{
    let signapp = new SignApp();
    await signInO365(signapp.setState, signapp.setToken, signapp.displayError);
}

