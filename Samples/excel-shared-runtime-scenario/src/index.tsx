import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { AppContainer } from 'react-hot-loader';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { ensureStateInitialized, updateRibbon, monitorSheetChanges } from '../utilities/office-apis-helpers';

import App from './components/App';

import './styles.less';
import 'office-ui-fabric-react/dist/css/fabric.min.css';

initializeIcons();

let isOfficeInitialized = false;

const title = 'Office-Add-in-Contoso-Data-Importer';

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
    ensureStateInitialized(true);
    console.log("ensure state initialized from the office.initialize");
    isOfficeInitialized = true;
    monitorSheetChanges();

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

