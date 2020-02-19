import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { AppContainer } from 'react-hot-loader';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { getGlobal, updateRibbon } from '../utilities/office-apis-helpers';

import App from './components/App';
import { add, getData } from './functions/functions';

import './styles.less';
import 'office-ui-fabric-react/dist/css/fabric.min.css';
//import { registerOnThemeChangeCallback } from 'office-ui-fabric-react';

initializeIcons();

let isOfficeInitialized = false;

const title = 'Office-Add-in-Microsoft-Graph-React';

const render = (Component) => {
    ReactDOM.render(
        <AppContainer>
            <Component title={title} isOfficeInitialized={isOfficeInitialized} />
        </AppContainer>,
        document.getElementById('container')
    );
};

/* Render application after Office initializes */
Office.initialize = async () => {
    let g = getGlobal() as any;
    g.state = {
        'isStartOnDocOpen': false,
        'isSignedIn': false,
        'isTaskpaneOpen': false,
        'isConnected': false,
        'isSyncEnabled': false,
        'isConnectInProgress': false,
        'isFirstSyncCall': true,
        updateRct: () => { },
        setTaskpaneStatus: (opened: boolean) => {
            g.state.isTaskpaneOpen = opened;
            updateRibbon();
        },
        setConnected: (connected: boolean) => {
            g.state.isConnected = connected;


            if (connected) {
                if (g.state.updateRct !== null) {
                    g.state.updateRct('true');
                }
            } else {
                if (g.state.updateRct !== null) {
                    g.state.updateRct('false');
                }
            }
            updateRibbon();
        }
    };
    //    g.isStartOnDocOpen = false;
    //  g.isSignedIn = false;

    // @ts-ignore
    let addinState = await Office.addin._getState();
    console.log("load state is:");
    console.log("load state" + addinState);
    if (addinState === 'Background') {
        g.state.isStartOnDocOpen = true;
        //run();
    }
    if (localStorage.getItem('loggedIn') === 'yes') {
        g.state.isSignedIn = true;
    }

    isOfficeInitialized = true;
    // SetRuntimeVisibleHelper(true);
    // @ts-ignore
    //SetStartupBehaviorHelper(Office.StartupBehavior.load);


    console.log('task pane running');
    CustomFunctions.associate('ADD', add);
    CustomFunctions.associate('GETDATA', getData);
    monitorSheetChanges();
    render(App);
};

async function onChange(event) {
    return Excel.run((context) => {
        return context.sync()
            .then(() => {
                console.log("Change type of event: " + event.changeType);
                console.log("Address of event: " + event.address);
                console.log("Source of event: " + event.source);
                let g = getGlobal() as any;
                if (g.state.isConnected && !g.state.isFirstSyncCall) {
                    g.state.isSyncEnabled = true;

                    updateRibbon();

                }
                g.state.isFirstSyncCall = false;
            });

    });
}

async function monitorSheetChanges() {
    try {
        await Excel.run(async (context) => {
            let sheet = context.workbook.worksheets.getActiveWorksheet();
            sheet.onChanged.add(onChange);

            await context.sync();
            console.log("A handler has been registered for the onChanged event.");
        });

    } catch (error) {
        console.error(error);
    }
}
/* Initial render showing a progress bar */
render(App);

if ((module as any).hot) {
    (module as any).hot.accept('./components/App', () => {
        const NextApp = require('./components/App').default;
        render(NextApp);
    });
}

