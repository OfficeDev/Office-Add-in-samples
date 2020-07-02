"use strict";
/*
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */
Object.defineProperty(exports, "__esModule", { value: true });
/// <reference types="office-js" />
/// <reference types="node" />
/* global Office */ //Required for this to be found.  see: https://github.com/OfficeDev/office-js-docs-pr/issues/691
require("react-app-polyfill/ie11");
require("react-app-polyfill/stable");
require("office-ui-fabric-react/dist/css/fabric.min.css");
var App_1 = require("./components/App");
var react_hot_loader_1 = require("react-hot-loader");
var Icons_1 = require("office-ui-fabric-react/lib/Icons");
var React = require("react");
var ReactDOM = require("react-dom");
Icons_1.initializeIcons();
var isOfficeInitialized = false;
var title = 'Contoso Task Pane Add-in TypeScript and .NET Core 3.1';
var render = function (Component) {
    ReactDOM.render(React.createElement(react_hot_loader_1.AppContainer, null,
        React.createElement(Component, { title: title, isOfficeInitialized: isOfficeInitialized })), document.getElementById('container'));
};
/* Render application after Office initializes */
Office.initialize = function () {
    isOfficeInitialized = true;
    render(App_1.default);
};
/* Initial render showing a progress bar */
render(App_1.default);
if (module.hot) {
    module.hot.accept('./components/App', function () {
        var NextApp = require('./components/App').default;
        render(NextApp);
    });
}
//# sourceMappingURL=index.js.map