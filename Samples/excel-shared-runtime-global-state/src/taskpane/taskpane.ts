import { add, clock, currentTime, logMessage } from '../functions/functions';
import { getGlobal } from '../commands/commands';

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.initialize = () => {
  let g = getGlobal() as any;
  let keys: any[] = [];
  let values: any[] = [];
  g.state = {
    "keys": keys,
    "values": values
  } as any
  

  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("btnStoreValue").onclick = btnStoreValue;
  document.getElementById("btnGetValue").onclick = btnGetValue;
  
  // eslint-disable-next-line no-undef
  CustomFunctions.associate('ADD', add);

  // eslint-disable-next-line no-undef
  CustomFunctions.associate('CLOCK', clock);

  // eslint-disable-next-line no-undef
  CustomFunctions.associate('CURRENTTIME', currentTime);

  // eslint-disable-next-line no-undef
  CustomFunctions.associate('LOGMESSAGE', logMessage);
  
};

function btnStoreValue() {
  // @ts-ignore
  let key = document.getElementById("txtKey").value;
  // @ts-ignore
  let value = document.getElementById("txtValue").value;
  let g = getGlobal() as any;
  g.state.keys.push(key);
  g.state.values.push(value);
}

function btnGetValue() {
  let g = getGlobal() as any;
  // @ts-ignore
  let key = document.getElementById("txtKey").value;
  g.state.keys.forEach((element, index) => {
    if (element === key)
    {
      // @ts-ignore
      document.getElementById("txtValue").value = g.state.values[index];
    }
  });
}