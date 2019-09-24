/*
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

/// <reference types="office-js" />
/// <reference types="node" />
/* global Office */  //Required for this to be found.  see: https://github.com/OfficeDev/office-js-docs-pr/issues/691

import 'react-app-polyfill/ie11';
import 'react-app-polyfill/stable';
import 'office-ui-fabric-react/dist/css/fabric.min.css';
import App from './components/App';
import { AppContainer } from 'react-hot-loader';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import * as React from 'react';
import * as ReactDOM from 'react-dom';

initializeIcons();


let isOfficeInitialized = false;

const title = 'Contoso Task Pane Add-in TypeScript2';

const render = (Component) => {
    ReactDOM.render(
        <AppContainer>
            <Component title={title} isOfficeInitialized={isOfficeInitialized} />
        </AppContainer>,
        document.getElementById('container')
    );
};

/* Render application after Office initializes */
Office.initialize = () => {
    isOfficeInitialized = true;
    render(App);
};

/* Initial render showing a progress bar */
render(App);

if ((module as any).hot) {
    (module as any).hot.accept('./components/App', () => {
        const NextApp = require('./components/App').default;
        render(NextApp);
    });
}