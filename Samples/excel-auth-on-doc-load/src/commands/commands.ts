import { SignApp, signInO365, SetStartupBehaviorHelper, SetRuntimeVisibleHelper, updateRibbon, getGlobal } from '../../utilities/office-apis-helpers';

const g = getGlobal() as any;

// the add-in command functions need to be available in global scope
g.btnsignin = btnSignIn;
g.btnsignout = btnSignOut;
g.btnenableaddinstart = btnEnableAddinStart;
g.btndisableaddinstart = btnDisableAddinStart;
g.btninsertdata = btnInsertData;
g.btnopentaskpane = btnOpenTaskpane;
g.btnclosetaskpane = btnCloseTaskpane;

export function btnOpenTaskpane(event: Office.AddinCommands.Event) {
    console.log('Open task pane button pressed');
    // Your code goes here
    SetRuntimeVisibleHelper(true);
    updateRibbon();
    event.completed();
}

export function btnCloseTaskpane(event: Office.AddinCommands.Event) {
    console.log('Open task pane button pressed');
    // Your code goes here
    SetRuntimeVisibleHelper(false);
    event.completed();
}

export function btnSignIn(event: Office.AddinCommands.Event) {
    console.log('sign in button pressed');
    // Your code goes here

    let signapp = new SignApp();
    signInO365(signapp.setState, signapp.setToken, signapp.displayError);
    //SetRuntimeVisibleHelper(true);
    // Be sure to indicate when the add-in command function is complete
    event.completed();
}

export function btnSignOut(event: Office.AddinCommands.Event) {
    console.log('sign out button pressed');
    // Your code goes here

    event.completed();
}

export function btnEnableAddinStart(event: Office.AddinCommands.Event) {
    console.log('Enable add-in start button pressed');
    // Your code goes here
    SetStartupBehaviorHelper(true)
    event.completed();
}

export function btnDisableAddinStart(event: Office.AddinCommands.Event) {
    console.log('Disable add-in start button pressed');
    // Your code goes here
    SetStartupBehaviorHelper(false);
    event.completed();
}

export function btnInsertData(event: Office.AddinCommands.Event) {
    console.log('Insert data button pressed');
    // Mock code that pretends to insert data from a data source
    insertData();
    event.completed();
}



async function insertData() {
    try {
      await Excel.run(async context => {
        /**
         * Insert your Excel code here
         */
        const ws = context.workbook.worksheets.getActiveWorksheet();
        let range = ws.getRange('A1');
        range.load('values');
        return context.sync(range).then( (range) => {
            let v = range.values[0][0];
            v += 1;
            range.values = [[ v ]];
            range.format.autofitColumns();

            return context.sync();
        });
    });
    } catch (error) {
      console.error(error);
    }
  }