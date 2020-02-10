import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { AppContainer } from 'react-hot-loader';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
//import { SetRuntimeVisibleHelper,SetStartupBehaviorHelper } from '../utilities/office-apis-helpers';

import App from './components/App';
import {add} from './functions/functions';

import './styles.less';
import 'office-ui-fabric-react/dist/css/fabric.min.css';

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
Office.initialize = () => {
    isOfficeInitialized = true;
   // SetRuntimeVisibleHelper(true);
    // @ts-ignore
    //SetStartupBehaviorHelper(Office.StartupBehavior.load);

    // init the xls
    run();


    console.log('task pane running');
    CustomFunctions.associate('ADD',add);
    render(App);
};


async function run() {
    try {
      await Excel.run(async context => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();
  
        // Read the range address
        range.load("address");
  
        // Update the fill color
        range.format.fill.color = "yellow";
  
        await context.sync();
        console.log(`The range address was ${range.address}.`);
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

