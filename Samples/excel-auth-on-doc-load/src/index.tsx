import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { AppContainer } from 'react-hot-loader';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { ensureStateInitialized, updateRibbon } from '../utilities/office-apis-helpers';

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
    ensureStateInitialized();

    isOfficeInitialized = true;
    // SetRuntimeVisibleHelper(true);
    // @ts-ignore
    //SetStartupBehaviorHelper(Office.StartupBehavior.load);


    console.log('task pane running');
    CustomFunctions.associate('ADD', add);
    CustomFunctions.associate('GETDATA', getData);
    render(App);
    updateRibbon();
};

/* Initial render showing a progress bar */
render(App);

if ((module as any).hot) {
    (module as any).hot.accept('./components/App', () => {
        const NextApp = require('./components/App').default;
        render(NextApp);
    });
}

