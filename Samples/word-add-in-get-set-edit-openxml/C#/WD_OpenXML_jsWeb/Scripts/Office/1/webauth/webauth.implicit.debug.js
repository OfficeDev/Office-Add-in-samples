(function webpackUniversalModuleDefinition(root, factory) {
	if(typeof exports === 'object' && typeof module === 'object')
		module.exports = factory();
	else if(typeof define === 'function' && define.amd)
		define("Implicit", [], factory);
	else if(typeof exports === 'object')
		exports["Implicit"] = factory();
	else
		root["Implicit"] = factory();
})(window, function() {
return /******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 0);
/******/ })
/************************************************************************/
/******/ ({

/***/ "./packages/Microsoft.Office.WebAuth.Implicit/lib/api.js":
/***/ (function(module, exports, __webpack_require__) {

"use strict";


Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.logUserAction = logUserAction;
exports.logActivity = logActivity;
exports.sendTelemetryEvent = sendTelemetryEvent;
exports.sendActivityEvent = sendActivityEvent;
exports.sendOtelEvent = sendOtelEvent;
exports.sendUserActionEvent = sendUserActionEvent;
exports.addNamespaceMapping = addNamespaceMapping;
exports.setEnabledState = setEnabledState;
exports.shutdown = shutdown;
exports.registerEventHandler = registerEventHandler; // Assume that telemetry is disabled and simply drop events on the floor unless the developer called initialize(true /*enabled*/).
// This should work well for component / unittest environments since nobody will end up listening to the events.
// The alternative is to cache them, but given that nobody will ever process them it might cause issues since the events would be cached forever.

var telemetryEnabled = false;
var events = [];
var eventHandler;
var numberOfDroppedEvents = 0;
var maxQueueSize = 20000;
var unknownStr = 'Unknown'; // Primary consumer public API
// ========================================================================================================================
// Call LogUserAction for logging a user action to Otel.
// This is similar to the bSqm actions that used to be logged earlier (deprecated now).
// Make sure you read the documentation below for userActionName and the Kusto table name implications.
// userActionName: Name of the user action, this should come from your app's commands,
//     for example: OneNoteCommands in office-online-ui\packages\onenote-online-ux\src\store\OneNoteCommands.ts (https://office.visualstudio.com/OC/_git/office-online-ui?path=%2Fpackages%2Fonenote-online-ux%2Fsrc%2Fstore%2FOneNoteCommands.ts&version=GBmaster)
//     Note that the userActionName will be the name of your table in Aria Kusto. So if 'ABC' is passed in for userActionName, the table in Kusto will be called Office_OneNote_Online_UserAction_ABC (or generically speaking Office_{AppName}_Online_UserAction_ABC )
//     Look at Kusto connection https://kusto.aria.microsoft.com and databases Office Word Online or Office OneNote Online, etc. and look at *UserAction* tables.
// success: Status of the user action (success is true, failure is false).
// parentNameStr: parent surface of the user action (example, tabView, tabHelp, Layout, etc).
// inputMethod: how the user action was performed (for example, via keyboard, or mouse, touch, etc.)
//             See the enum in /packages/app-commanding-ui/src/UISurfaces/controls/InputMethod.ts
//             Pass in this param as:  InputMethod.Keyboard.toString() instead of passing in "Keyboard"
// uiLocation: the surface where the user action was initiated from (example, ribbon, FileMenu, TellMe, etc).
//             See enum in /packages/app-commanding-ui/src/UISurfaces/controls/UILocation.ts
//             Pass in this param as:  UILocation.SingleLineRibbon.toString() instead of passing in "SingleLineRibbon"
// durationMsec: the time taken by the action (if relevant to the action)
// dataFieldArr: These are custom fields that you may want to add for your user action.
//               Example: InsertTable action may log custom data fields such as rowSize and colSize of the table inserted.
//                      Or in Excel, a cell related action may log the x and y coordinates of the cell.
// Note that things such as sessionID, data center, etc will be added to all user action logs.

function logUserAction(userActionName, success, parentNameStr, inputMethod, uiLocation, durationMsec, dataFieldArr) {
  if (success === void 0) {
    success = true;
  }

  if (parentNameStr === void 0) {
    parentNameStr = unknownStr;
  }

  if (inputMethod === void 0) {
    inputMethod = unknownStr;
  }

  if (uiLocation === void 0) {
    uiLocation = unknownStr;
  }

  if (durationMsec === void 0) {
    durationMsec = 0;
  }

  if (dataFieldArr === void 0) {
    dataFieldArr = [];
  } // passing null for 'name' field, which is the event table name. We will determine that in sendUserAction in full\api.ts as there we know what app we are, and hence what the event table name is


  sendUserActionEvent({
    name: null,
    actionName: userActionName,
    commandSurface: uiLocation,
    parentName: parentNameStr,
    triggerMethod: inputMethod,
    durationMs: durationMsec,
    succeeded: success,
    dataFields: dataFieldArr
  });
} //////////////////////////////////////////////////////////////////////////////////////////////////
// Call logActivity for logging an activity to Otel.
// This will be logged under Office {App} Online Data tenant
// For example, if your activity name is "ABC",
// it will go to a table called "Office_Word_Online_Data_Activity_ABC" for Word or "Office_OneNote_Online_Data_Activity_ABC" for OneNote.
// activityName: name of activity being logged
// success: Status of the activity (success is true, failure is false).
// durationMsec: the time taken by the action (if relevant to the action)
// dataFieldArr: These are custom fields that you may want to add for your activity, and will be added as columns to the activity table.
//               Example: dataFields has typingSpeedPerSec (integer) and dayOfWeek (string) in it, the activity table for this particular activity will contain these two custom fields.
// Note that things such as sessionID, data center, etc will be added to all user action logs.


function logActivity(activityName, success, durationMsec, dataFieldArr) {
  if (success === void 0) {
    success = true;
  }

  if (durationMsec === void 0) {
    durationMsec = 0;
  }

  if (dataFieldArr === void 0) {
    dataFieldArr = [];
  }

  sendActivityEvent({
    name: activityName,
    succeeded: success,
    durationMs: durationMsec,
    dataFields: dataFieldArr
  });
} // Call LogNonUserAction for logging a non user action to Otel.
// This is
// activityName: Name of the action (non user)
//     Note that the userActionName will be what your table will be named in Aria Kusto. So if 'ABC' is passed in for non user action, the table in Kusto will be called Office_OneNote_Online_NonUser_ABC
//     Look at Kusto connection https://kusto.aria.microsoft.com and databases Office Word Online or Office OneNote Online, etc. and look at *NonUser* tables.
// succeeded: Status of the user action (success is true, failure is false).
// parentName: parent surface of the user action (example, tabView, tabHelp, Layout, etc).
// inputMethod: how the user action was performed (for example, via keyboard, or mouse, touch, etc.)
// uiLocation: the surface where the user action was initiated from (example, ribbon, FileMenu, TellMe, etc)
// startTime: start time of the activity
// endTime: end time of the activity
// dataFields: These are custom fields that you may want to add for your user action.
//             Example: InsertTable action may log custom data fields such as rowSize and colSize of the table inserted.
//                      Or in Excel, a cell related action may log the x and y coordinates of the cell.
//             Note that things such as sessionID, data center, etc will be added to all user action logs.

/*
Being commented out as we dont think we should expose this API. But code is here in case someone educates us on why this should be exposed (it is being used historically in scriptsharp in OtelActionListener.cs (WsaListener.cs))

export const nonUserActionPrefix = 'non_user_action_'; // used in full\api.ts sendActivityEvent to determine if an activity is non user action or a regular activity

If this code is reinstated, then we need to add the following in full\api.ts:

import { nonUserActionPrefix } from '../core';
nonUserActionEventName = 'Office.Online.NonUserAction';
nonUserActionEventName = `Office.${settings.alwaysOnMetadata.name}.Online.NonUser.`;

function ContainsNonUserActionPrefix(eventName: string): boolean {
  return eventName.indexOf(nonUserActionPrefix) == 0;
}

And these lines in sendActivity function in full\api.ts

  if (event.name != null) {
    if (ContainsNonUserActionPrefix(event.name)) {
      event.name = nonUserActionEventName + event.name.substring(nonUserActionPrefix.length);
    }
  }

export function LogNonUserAction(
  activityName: string,
  succeeded: boolean = true,
  parentName: string = unknownStr,
  inputMethod: InputMethod = InputMethod.Unknown,
  uiLocation: UILocation | null = null,
  startTime: number = 0,
  endTime: number = 0,
  dataFields: DataField[] = []
) {
  let durationMs: number = Math.max(endTime - startTime, 0);

  dataFields!.push({ name: 'ParentName', string: parentName != null ? parentName : unknownStr });
  dataFields!.push({ name: 'TriggerMethod', string: inputMethod != null ? inputMethod.toString() : unknownStr });
  dataFields!.push({ name: 'CommandSurface', string: uiLocation != null ? uiLocation.toString() : unknownStr });
  dataFields!.push({ name: 'ActionName', string: activityName });
  dataFields!.push({ name: 'StartTime', double: startTime });
  dataFields!.push({ name: 'EndTime', double: endTime });
  dataFields!.push({ name: 'Succeeded', bool: succeeded });

  // add a sentinel prefix to activity name such that we know we need to add the non user activity event table name (instead of a regular activity event table name) in  sendActivityEvent in full\api.ts as there we know what app we are, and hence what the event table name is
  let activityNameWithPrefix = nonUserActionPrefix;
  if (activityName != null) {
    activityNameWithPrefix = activityNameWithPrefix.concat(activityName);
  }

  sendActivityEvent({
    name: activityNameWithPrefix,
    dataFields: dataFields,
    durationMs: durationMs,
    succeeded: succeeded
  });
}
*/


function sendTelemetryEvent(event) {
  raiseEvent({
    kind: 'event',
    event: event,
    timestamp: new Date().getTime()
  });
}

function sendActivityEvent(event) {
  raiseEvent({
    kind: 'activity',
    event: event,
    timestamp: new Date().getTime()
  });
}

function sendOtelEvent(event) {
  raiseEvent({
    kind: 'otel',
    event: event
  });
}

function sendUserActionEvent(event) {
  raiseEvent({
    kind: 'action',
    event: event,
    timestamp: new Date().getTime()
  });
}

function addNamespaceMapping(namespace, ariaTenantToken) {
  raiseEvent({
    kind: 'addNamespaceMapping',
    namespace: namespace,
    ariaTenantToken: ariaTenantToken
  });
} // Initialization / Shutdown
// ========================================================================================================================


function setEnabledState(enabled) {
  telemetryEnabled = enabled; // If the caller disables the queue, be sure to drop all of the outstanding events.
  // This can happen in cases where the slice with event processor functionality failed to load.

  if (!telemetryEnabled) {
    events = [];
  }
}

function shutdown() {
  raiseEvent({
    kind: 'shutdown'
  });
  return events.length + numberOfDroppedEvents;
}

function registerEventHandler(handler) {
  eventHandler = handler; // Then go through the queue and process the events in the order in which they were received
  // VSO.2533164: Push batch event processing to otelFull and add a lightweight queue

  events.forEach(function (event) {
    return raiseEvent(event);
  });
  events = [];
}

function raiseEvent(event) {
  if (!telemetryEnabled) {
    return;
  }

  if (eventHandler) {
    eventHandler(event);
  } else {
    if (events.length <= maxQueueSize) {
      events.push(event);
    } else {
      numberOfDroppedEvents += 1;
    }
  }
}


/***/ }),

/***/ "./packages/Microsoft.Office.WebAuth.Implicit/lib/msal.js":
/***/ (function(module, exports, __webpack_require__) {

"use strict";
/*! msal v1.3.3 2020-07-14 */

(function webpackUniversalModuleDefinition(root, factory) {
	if(true)
		module.exports = factory();
	else {}
})(window, function() {
return /******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 29);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*! *****************************************************************************
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the Apache License, Version 2.0 (the "License"); you may not use
this file except in compliance with the License. You may obtain a copy of the
License at http://www.apache.org/licenses/LICENSE-2.0

THIS CODE IS PROVIDED ON AN *AS IS* BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
KIND, EITHER EXPRESS OR IMPLIED, INCLUDING WITHOUT LIMITATION ANY IMPLIED
WARRANTIES OR CONDITIONS OF TITLE, FITNESS FOR A PARTICULAR PURPOSE,
MERCHANTABLITY OR NON-INFRINGEMENT.

See the Apache Version 2.0 License for specific language governing permissions
and limitations under the License.
***************************************************************************** */
/* global Reflect, Promise */
Object.defineProperty(exports, "__esModule", { value: true });
var extendStatics = function (d, b) {
    extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b)
            if (b.hasOwnProperty(p))
                d[p] = b[p]; };
    return extendStatics(d, b);
};
function __extends(d, b) {
    extendStatics(d, b);
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
}
exports.__extends = __extends;
exports.__assign = function () {
    exports.__assign = Object.assign || function __assign(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s)
                if (Object.prototype.hasOwnProperty.call(s, p))
                    t[p] = s[p];
        }
        return t;
    };
    return exports.__assign.apply(this, arguments);
};
function __rest(s, e) {
    var t = {};
    for (var p in s)
        if (Object.prototype.hasOwnProperty.call(s, p) && e.indexOf(p) < 0)
            t[p] = s[p];
    if (s != null && typeof Object.getOwnPropertySymbols === "function")
        for (var i = 0, p = Object.getOwnPropertySymbols(s); i < p.length; i++) {
            if (e.indexOf(p[i]) < 0 && Object.prototype.propertyIsEnumerable.call(s, p[i]))
                t[p[i]] = s[p[i]];
        }
    return t;
}
exports.__rest = __rest;
function __decorate(decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function")
        r = Reflect.decorate(decorators, target, key, desc);
    else
        for (var i = decorators.length - 1; i >= 0; i--)
            if (d = decorators[i])
                r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
}
exports.__decorate = __decorate;
function __param(paramIndex, decorator) {
    return function (target, key) { decorator(target, key, paramIndex); };
}
exports.__param = __param;
function __metadata(metadataKey, metadataValue) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function")
        return Reflect.metadata(metadataKey, metadataValue);
}
exports.__metadata = __metadata;
function __awaiter(thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try {
            step(generator.next(value));
        }
        catch (e) {
            reject(e);
        } }
        function rejected(value) { try {
            step(generator["throw"](value));
        }
        catch (e) {
            reject(e);
        } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
}
exports.__awaiter = __awaiter;
function __generator(thisArg, body) {
    var _ = { label: 0, sent: function () { if (t[0] & 1)
            throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function () { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f)
            throw new TypeError("Generator is already executing.");
        while (_)
            try {
                if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done)
                    return t;
                if (y = 0, t)
                    op = [op[0] & 2, t.value];
                switch (op[0]) {
                    case 0:
                    case 1:
                        t = op;
                        break;
                    case 4:
                        _.label++;
                        return { value: op[1], done: false };
                    case 5:
                        _.label++;
                        y = op[1];
                        op = [0];
                        continue;
                    case 7:
                        op = _.ops.pop();
                        _.trys.pop();
                        continue;
                    default:
                        if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) {
                            _ = 0;
                            continue;
                        }
                        if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) {
                            _.label = op[1];
                            break;
                        }
                        if (op[0] === 6 && _.label < t[1]) {
                            _.label = t[1];
                            t = op;
                            break;
                        }
                        if (t && _.label < t[2]) {
                            _.label = t[2];
                            _.ops.push(op);
                            break;
                        }
                        if (t[2])
                            _.ops.pop();
                        _.trys.pop();
                        continue;
                }
                op = body.call(thisArg, _);
            }
            catch (e) {
                op = [6, e];
                y = 0;
            }
            finally {
                f = t = 0;
            }
        if (op[0] & 5)
            throw op[1];
        return { value: op[0] ? op[1] : void 0, done: true };
    }
}
exports.__generator = __generator;
function __exportStar(m, exports) {
    for (var p in m)
        if (!exports.hasOwnProperty(p))
            exports[p] = m[p];
}
exports.__exportStar = __exportStar;
function __values(o) {
    var m = typeof Symbol === "function" && o[Symbol.iterator], i = 0;
    if (m)
        return m.call(o);
    return {
        next: function () {
            if (o && i >= o.length)
                o = void 0;
            return { value: o && o[i++], done: !o };
        }
    };
}
exports.__values = __values;
function __read(o, n) {
    var m = typeof Symbol === "function" && o[Symbol.iterator];
    if (!m)
        return o;
    var i = m.call(o), r, ar = [], e;
    try {
        while ((n === void 0 || n-- > 0) && !(r = i.next()).done)
            ar.push(r.value);
    }
    catch (error) {
        e = { error: error };
    }
    finally {
        try {
            if (r && !r.done && (m = i["return"]))
                m.call(i);
        }
        finally {
            if (e)
                throw e.error;
        }
    }
    return ar;
}
exports.__read = __read;
function __spread() {
    for (var ar = [], i = 0; i < arguments.length; i++)
        ar = ar.concat(__read(arguments[i]));
    return ar;
}
exports.__spread = __spread;
function __spreadArrays() {
    for (var s = 0, i = 0, il = arguments.length; i < il; i++)
        s += arguments[i].length;
    for (var r = Array(s), k = 0, i = 0; i < il; i++)
        for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++)
            r[k] = a[j];
    return r;
}
exports.__spreadArrays = __spreadArrays;
;
function __await(v) {
    return this instanceof __await ? (this.v = v, this) : new __await(v);
}
exports.__await = __await;
function __asyncGenerator(thisArg, _arguments, generator) {
    if (!Symbol.asyncIterator)
        throw new TypeError("Symbol.asyncIterator is not defined.");
    var g = generator.apply(thisArg, _arguments || []), i, q = [];
    return i = {}, verb("next"), verb("throw"), verb("return"), i[Symbol.asyncIterator] = function () { return this; }, i;
    function verb(n) { if (g[n])
        i[n] = function (v) { return new Promise(function (a, b) { q.push([n, v, a, b]) > 1 || resume(n, v); }); }; }
    function resume(n, v) { try {
        step(g[n](v));
    }
    catch (e) {
        settle(q[0][3], e);
    } }
    function step(r) { r.value instanceof __await ? Promise.resolve(r.value.v).then(fulfill, reject) : settle(q[0][2], r); }
    function fulfill(value) { resume("next", value); }
    function reject(value) { resume("throw", value); }
    function settle(f, v) { if (f(v), q.shift(), q.length)
        resume(q[0][0], q[0][1]); }
}
exports.__asyncGenerator = __asyncGenerator;
function __asyncDelegator(o) {
    var i, p;
    return i = {}, verb("next"), verb("throw", function (e) { throw e; }), verb("return"), i[Symbol.iterator] = function () { return this; }, i;
    function verb(n, f) { i[n] = o[n] ? function (v) { return (p = !p) ? { value: __await(o[n](v)), done: n === "return" } : f ? f(v) : v; } : f; }
}
exports.__asyncDelegator = __asyncDelegator;
function __asyncValues(o) {
    if (!Symbol.asyncIterator)
        throw new TypeError("Symbol.asyncIterator is not defined.");
    var m = o[Symbol.asyncIterator], i;
    return m ? m.call(o) : (o = typeof __values === "function" ? __values(o) : o[Symbol.iterator](), i = {}, verb("next"), verb("throw"), verb("return"), i[Symbol.asyncIterator] = function () { return this; }, i);
    function verb(n) { i[n] = o[n] && function (v) { return new Promise(function (resolve, reject) { v = o[n](v), settle(resolve, reject, v.done, v.value); }); }; }
    function settle(resolve, reject, d, v) { Promise.resolve(v).then(function (v) { resolve({ value: v, done: d }); }, reject); }
}
exports.__asyncValues = __asyncValues;
function __makeTemplateObject(cooked, raw) {
    if (Object.defineProperty) {
        Object.defineProperty(cooked, "raw", { value: raw });
    }
    else {
        cooked.raw = raw;
    }
    return cooked;
}
exports.__makeTemplateObject = __makeTemplateObject;
;
function __importStar(mod) {
    if (mod && mod.__esModule)
        return mod;
    var result = {};
    if (mod != null)
        for (var k in mod)
            if (Object.hasOwnProperty.call(mod, k))
                result[k] = mod[k];
    result.default = mod;
    return result;
}
exports.__importStar = __importStar;
function __importDefault(mod) {
    return (mod && mod.__esModule) ? mod : { default: mod };
}
exports.__importDefault = __importDefault;


/***/ }),
/* 1 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
/**
 * @hidden
 * Constants
 */
var Constants = /** @class */ (function () {
    function Constants() {
    }
    Object.defineProperty(Constants, "libraryName", {
        get: function () { return "Msal.js"; } // used in telemetry sdkName
        ,
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "claims", {
        get: function () { return "claims"; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "clientId", {
        get: function () { return "clientId"; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "adalIdToken", {
        get: function () { return "adal.idtoken"; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "cachePrefix", {
        get: function () { return "msal"; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "scopes", {
        get: function () { return "scopes"; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "no_account", {
        get: function () { return "NO_ACCOUNT"; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "upn", {
        get: function () { return "upn"; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "domain_hint", {
        get: function () { return "domain_hint"; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "prompt_select_account", {
        get: function () { return "&prompt=select_account"; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "prompt_none", {
        get: function () { return "&prompt=none"; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "prompt", {
        get: function () { return "prompt"; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "response_mode_fragment", {
        get: function () { return "&response_mode=fragment"; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "resourceDelimiter", {
        get: function () { return "|"; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "cacheDelimiter", {
        get: function () { return "."; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "popUpWidth", {
        get: function () { return this._popUpWidth; },
        set: function (width) {
            this._popUpWidth = width;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "popUpHeight", {
        get: function () { return this._popUpHeight; },
        set: function (height) {
            this._popUpHeight = height;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "login", {
        get: function () { return "LOGIN"; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "renewToken", {
        get: function () { return "RENEW_TOKEN"; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "unknown", {
        get: function () { return "UNKNOWN"; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "ADFS", {
        get: function () { return "adfs"; },
        enumerable: true,
        configurable: true
    });
    ;
    Object.defineProperty(Constants, "homeAccountIdentifier", {
        get: function () { return "homeAccountIdentifier"; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "common", {
        get: function () { return "common"; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "openidScope", {
        get: function () { return "openid"; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "profileScope", {
        get: function () { return "profile"; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "interactionTypeRedirect", {
        get: function () { return "redirectInteraction"; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "interactionTypePopup", {
        get: function () { return "popupInteraction"; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "interactionTypeSilent", {
        get: function () { return "silentInteraction"; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Constants, "inProgress", {
        get: function () { return "inProgress"; },
        enumerable: true,
        configurable: true
    });
    Constants._popUpWidth = 483;
    Constants._popUpHeight = 600;
    return Constants;
}());
exports.Constants = Constants;
/**
 * Keys in the hashParams
 */
var ServerHashParamKeys;
(function (ServerHashParamKeys) {
    ServerHashParamKeys["SCOPE"] = "scope";
    ServerHashParamKeys["STATE"] = "state";
    ServerHashParamKeys["ERROR"] = "error";
    ServerHashParamKeys["ERROR_DESCRIPTION"] = "error_description";
    ServerHashParamKeys["ACCESS_TOKEN"] = "access_token";
    ServerHashParamKeys["ID_TOKEN"] = "id_token";
    ServerHashParamKeys["EXPIRES_IN"] = "expires_in";
    ServerHashParamKeys["SESSION_STATE"] = "session_state";
    ServerHashParamKeys["CLIENT_INFO"] = "client_info";
})(ServerHashParamKeys = exports.ServerHashParamKeys || (exports.ServerHashParamKeys = {}));
;
/**
 * @hidden
 * CacheKeys for MSAL
 */
var TemporaryCacheKeys;
(function (TemporaryCacheKeys) {
    TemporaryCacheKeys["AUTHORITY"] = "authority";
    TemporaryCacheKeys["ACQUIRE_TOKEN_ACCOUNT"] = "acquireTokenAccount";
    TemporaryCacheKeys["SESSION_STATE"] = "session.state";
    TemporaryCacheKeys["STATE_LOGIN"] = "state.login";
    TemporaryCacheKeys["STATE_ACQ_TOKEN"] = "state.acquireToken";
    TemporaryCacheKeys["STATE_RENEW"] = "state.renew";
    TemporaryCacheKeys["NONCE_IDTOKEN"] = "nonce.idtoken";
    TemporaryCacheKeys["LOGIN_REQUEST"] = "login.request";
    TemporaryCacheKeys["RENEW_STATUS"] = "token.renew.status";
    TemporaryCacheKeys["URL_HASH"] = "urlHash";
    TemporaryCacheKeys["INTERACTION_STATUS"] = "interaction_status";
    TemporaryCacheKeys["REDIRECT_REQUEST"] = "redirect_request";
})(TemporaryCacheKeys = exports.TemporaryCacheKeys || (exports.TemporaryCacheKeys = {}));
var PersistentCacheKeys;
(function (PersistentCacheKeys) {
    PersistentCacheKeys["IDTOKEN"] = "idtoken";
    PersistentCacheKeys["CLIENT_INFO"] = "client.info";
})(PersistentCacheKeys = exports.PersistentCacheKeys || (exports.PersistentCacheKeys = {}));
var ErrorCacheKeys;
(function (ErrorCacheKeys) {
    ErrorCacheKeys["LOGIN_ERROR"] = "login.error";
    ErrorCacheKeys["ERROR"] = "error";
    ErrorCacheKeys["ERROR_DESC"] = "error.description";
})(ErrorCacheKeys = exports.ErrorCacheKeys || (exports.ErrorCacheKeys = {}));
exports.DEFAULT_AUTHORITY = "https://login.microsoftonline.com/common/";
exports.AAD_INSTANCE_DISCOVERY_ENDPOINT = exports.DEFAULT_AUTHORITY + "/discovery/instance?api-version=1.1&authorization_endpoint=";
/**
 * @hidden
 * SSO Types - generated to populate hints
 */
var SSOTypes;
(function (SSOTypes) {
    SSOTypes["ACCOUNT"] = "account";
    SSOTypes["SID"] = "sid";
    SSOTypes["LOGIN_HINT"] = "login_hint";
    SSOTypes["ID_TOKEN"] = "id_token";
    SSOTypes["ACCOUNT_ID"] = "accountIdentifier";
    SSOTypes["HOMEACCOUNT_ID"] = "homeAccountIdentifier";
})(SSOTypes = exports.SSOTypes || (exports.SSOTypes = {}));
;
/**
 * @hidden
 */
exports.BlacklistedEQParams = [
    SSOTypes.SID,
    SSOTypes.LOGIN_HINT
];
exports.NetworkRequestType = {
    GET: "GET",
    POST: "POST"
};
/**
 * we considered making this "enum" in the request instead of string, however it looks like the allowed list of
 * prompt values kept changing over past couple of years. There are some undocumented prompt values for some
 * internal partners too, hence the choice of generic "string" type instead of the "enum"
 * @hidden
 */
exports.PromptState = {
    LOGIN: "login",
    SELECT_ACCOUNT: "select_account",
    CONSENT: "consent",
    NONE: "none"
};
/**
 * Frame name prefixes for the hidden iframe created in silent frames
 */
exports.FramePrefix = {
    ID_TOKEN_FRAME: "msalIdTokenFrame",
    TOKEN_FRAME: "msalRenewFrame"
};
/**
 * MSAL JS Library Version
 */
function libraryVersion() {
    return "1.3.3";
}
exports.libraryVersion = libraryVersion;


/***/ }),
/* 2 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
/**
 * @hidden
 */
var CryptoUtils = /** @class */ (function () {
    function CryptoUtils() {
    }
    /**
     * Creates a new random GUID
     * @returns string (GUID)
     */
    CryptoUtils.createNewGuid = function () {
        /*
         * RFC4122: The version 4 UUID is meant for generating UUIDs from truly-random or
         * pseudo-random numbers.
         * The algorithm is as follows:
         *     Set the two most significant bits (bits 6 and 7) of the
         *        clock_seq_hi_and_reserved to zero and one, respectively.
         *     Set the four most significant bits (bits 12 through 15) of the
         *        time_hi_and_version field to the 4-bit version number from
         *        Section 4.1.3. Version4
         *     Set all the other bits to randomly (or pseudo-randomly) chosen
         *     values.
         * UUID                   = time-low "-" time-mid "-"time-high-and-version "-"clock-seq-reserved and low(2hexOctet)"-" node
         * time-low               = 4hexOctet
         * time-mid               = 2hexOctet
         * time-high-and-version  = 2hexOctet
         * clock-seq-and-reserved = hexOctet:
         * clock-seq-low          = hexOctet
         * node                   = 6hexOctet
         * Format: xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx
         * y could be 1000, 1001, 1010, 1011 since most significant two bits needs to be 10
         * y values are 8, 9, A, B
         */
        var cryptoObj = window.crypto; // for IE 11
        if (cryptoObj && cryptoObj.getRandomValues) {
            var buffer = new Uint8Array(16);
            cryptoObj.getRandomValues(buffer);
            // buffer[6] and buffer[7] represents the time_hi_and_version field. We will set the four most significant bits (4 through 7) of buffer[6] to represent decimal number 4 (UUID version number).
            buffer[6] |= 0x40; // buffer[6] | 01000000 will set the 6 bit to 1.
            buffer[6] &= 0x4f; // buffer[6] & 01001111 will set the 4, 5, and 7 bit to 0 such that bits 4-7 == 0100 = "4".
            // buffer[8] represents the clock_seq_hi_and_reserved field. We will set the two most significant bits (6 and 7) of the clock_seq_hi_and_reserved to zero and one, respectively.
            buffer[8] |= 0x80; // buffer[8] | 10000000 will set the 7 bit to 1.
            buffer[8] &= 0xbf; // buffer[8] & 10111111 will set the 6 bit to 0.
            return CryptoUtils.decimalToHex(buffer[0]) + CryptoUtils.decimalToHex(buffer[1])
                + CryptoUtils.decimalToHex(buffer[2]) + CryptoUtils.decimalToHex(buffer[3])
                + "-" + CryptoUtils.decimalToHex(buffer[4]) + CryptoUtils.decimalToHex(buffer[5])
                + "-" + CryptoUtils.decimalToHex(buffer[6]) + CryptoUtils.decimalToHex(buffer[7])
                + "-" + CryptoUtils.decimalToHex(buffer[8]) + CryptoUtils.decimalToHex(buffer[9])
                + "-" + CryptoUtils.decimalToHex(buffer[10]) + CryptoUtils.decimalToHex(buffer[11])
                + CryptoUtils.decimalToHex(buffer[12]) + CryptoUtils.decimalToHex(buffer[13])
                + CryptoUtils.decimalToHex(buffer[14]) + CryptoUtils.decimalToHex(buffer[15]);
        }
        else {
            var guidHolder = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx";
            var hex = "0123456789abcdef";
            var r = 0;
            var guidResponse = "";
            for (var i = 0; i < 36; i++) {
                if (guidHolder[i] !== "-" && guidHolder[i] !== "4") {
                    // each x and y needs to be random
                    r = Math.random() * 16 | 0;
                }
                if (guidHolder[i] === "x") {
                    guidResponse += hex[r];
                }
                else if (guidHolder[i] === "y") {
                    // clock-seq-and-reserved first hex is filtered and remaining hex values are random
                    r &= 0x3; // bit and with 0011 to set pos 2 to zero ?0??
                    r |= 0x8; // set pos 3 to 1 as 1???
                    guidResponse += hex[r];
                }
                else {
                    guidResponse += guidHolder[i];
                }
            }
            return guidResponse;
        }
    };
    /**
     * verifies if a string is  GUID
     * @param guid
     */
    CryptoUtils.isGuid = function (guid) {
        var regexGuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
        return regexGuid.test(guid);
    };
    /**
     * Decimal to Hex
     *
     * @param num
     */
    CryptoUtils.decimalToHex = function (num) {
        var hex = num.toString(16);
        while (hex.length < 2) {
            hex = "0" + hex;
        }
        return hex;
    };
    // See: https://developer.mozilla.org/en-US/docs/Web/API/WindowBase64/Base64_encoding_and_decoding#Solution_4_%E2%80%93_escaping_the_string_before_encoding_it
    /**
     * encoding string to base64 - platform specific check
     *
     * @param input
     */
    CryptoUtils.base64Encode = function (input) {
        return btoa(encodeURIComponent(input).replace(/%([0-9A-F]{2})/g, function toSolidBytes(match, p1) {
            return String.fromCharCode(Number("0x" + p1));
        }));
    };
    /**
     * Decodes a base64 encoded string.
     *
     * @param input
     */
    CryptoUtils.base64Decode = function (input) {
        var encodedString = input.replace(/-/g, "+").replace(/_/g, "/");
        switch (encodedString.length % 4) {
            case 0:
                break;
            case 2:
                encodedString += "==";
                break;
            case 3:
                encodedString += "=";
                break;
            default:
                throw new Error("Invalid base64 string");
        }
        return decodeURIComponent(atob(encodedString).split("").map(function (c) {
            return "%" + ("00" + c.charCodeAt(0).toString(16)).slice(-2);
        }).join(""));
    };
    /**
     * deserialize a string
     *
     * @param query
     */
    CryptoUtils.deserialize = function (query) {
        var match; // Regex for replacing addition symbol with a space
        var pl = /\+/g;
        var search = /([^&=]+)=([^&]*)/g;
        var decode = function (s) { return decodeURIComponent(decodeURIComponent(s.replace(pl, " "))); }; // Some values (e.g. state) may need to be decoded twice
        var obj = {};
        match = search.exec(query);
        while (match) {
            obj[decode(match[1])] = decode(match[2]);
            match = search.exec(query);
        }
        return obj;
    };
    return CryptoUtils;
}());
exports.CryptoUtils = CryptoUtils;


/***/ }),
/* 3 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
/**
 * @hidden
 */
var StringUtils = /** @class */ (function () {
    function StringUtils() {
    }
    /**
     * Check if a string is empty
     *
     * @param str
     */
    StringUtils.isEmpty = function (str) {
        return (typeof str === "undefined" || !str || 0 === str.length);
    };
    return StringUtils;
}());
exports.StringUtils = StringUtils;


/***/ }),
/* 4 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
var Constants_1 = __webpack_require__(1);
var ScopeSet_1 = __webpack_require__(9);
var StringUtils_1 = __webpack_require__(3);
var CryptoUtils_1 = __webpack_require__(2);
/**
 * @hidden
 */
var UrlUtils = /** @class */ (function () {
    function UrlUtils() {
    }
    /**
     * generates the URL with QueryString Parameters
     * @param scopes
     */
    UrlUtils.createNavigateUrl = function (serverRequestParams) {
        var str = this.createNavigationUrlString(serverRequestParams);
        var authEndpoint = serverRequestParams.authorityInstance.AuthorizationEndpoint;
        // if the endpoint already has queryparams, lets add to it, otherwise add the first one
        if (authEndpoint.indexOf("?") < 0) {
            authEndpoint += "?";
        }
        else {
            authEndpoint += "&";
        }
        var requestUrl = "" + authEndpoint + str.join("&");
        return requestUrl;
    };
    /**
     * Generate the array of all QueryStringParams to be sent to the server
     * @param scopes
     */
    UrlUtils.createNavigationUrlString = function (serverRequestParams) {
        var scopes = serverRequestParams.scopes;
        if (scopes.indexOf(serverRequestParams.clientId) === -1) {
            scopes.push(serverRequestParams.clientId);
        }
        var str = [];
        str.push("response_type=" + serverRequestParams.responseType);
        this.translateclientIdUsedInScope(scopes, serverRequestParams.clientId);
        str.push("scope=" + encodeURIComponent(ScopeSet_1.ScopeSet.parseScope(scopes)));
        str.push("client_id=" + encodeURIComponent(serverRequestParams.clientId));
        str.push("redirect_uri=" + encodeURIComponent(serverRequestParams.redirectUri));
        str.push("state=" + encodeURIComponent(serverRequestParams.state));
        str.push("nonce=" + encodeURIComponent(serverRequestParams.nonce));
        str.push("client_info=1");
        str.push("x-client-SKU=" + serverRequestParams.xClientSku);
        str.push("x-client-Ver=" + serverRequestParams.xClientVer);
        if (serverRequestParams.promptValue) {
            str.push("prompt=" + encodeURIComponent(serverRequestParams.promptValue));
        }
        if (serverRequestParams.claimsValue) {
            str.push("claims=" + encodeURIComponent(serverRequestParams.claimsValue));
        }
        if (serverRequestParams.queryParameters) {
            str.push(serverRequestParams.queryParameters);
        }
        if (serverRequestParams.extraQueryParameters) {
            str.push(serverRequestParams.extraQueryParameters);
        }
        str.push("client-request-id=" + encodeURIComponent(serverRequestParams.correlationId));
        return str;
    };
    /**
     * append the required scopes: https://openid.net/specs/openid-connect-basic-1_0.html#Scopes
     * @param scopes
     */
    UrlUtils.translateclientIdUsedInScope = function (scopes, clientId) {
        var clientIdIndex = scopes.indexOf(clientId);
        if (clientIdIndex >= 0) {
            scopes.splice(clientIdIndex, 1);
            if (scopes.indexOf("openid") === -1) {
                scopes.push("openid");
            }
            if (scopes.indexOf("profile") === -1) {
                scopes.push("profile");
            }
        }
    };
    /**
     * Returns current window URL as redirect uri
     */
    UrlUtils.getCurrentUrl = function () {
        return window.location.href.split("?")[0].split("#")[0];
    };
    /**
     * Returns given URL with query string removed
     */
    UrlUtils.removeHashFromUrl = function (url) {
        return url.split("#")[0];
    };
    /**
     * Given a url like https://a:b/common/d?e=f#g, and a tenantId, returns https://a:b/tenantId/d
     * @param href The url
     * @param tenantId The tenant id to replace
     */
    UrlUtils.replaceTenantPath = function (url, tenantId) {
        url = url.toLowerCase();
        var urlObject = this.GetUrlComponents(url);
        var pathArray = urlObject.PathSegments;
        if (tenantId && (pathArray.length !== 0 && pathArray[0] === Constants_1.Constants.common)) {
            pathArray[0] = tenantId;
        }
        return this.constructAuthorityUriFromObject(urlObject, pathArray);
    };
    UrlUtils.constructAuthorityUriFromObject = function (urlObject, pathArray) {
        return this.CanonicalizeUri(urlObject.Protocol + "//" + urlObject.HostNameAndPort + "/" + pathArray.join("/"));
    };
    /**
     * Parses out the components from a url string.
     * @returns An object with the various components. Please cache this value insted of calling this multiple times on the same url.
     */
    UrlUtils.GetUrlComponents = function (url) {
        if (!url) {
            throw "Url required";
        }
        // https://gist.github.com/curtisz/11139b2cfcaef4a261e0
        var regEx = RegExp("^(([^:/?#]+):)?(//([^/?#]*))?([^?#]*)(\\?([^#]*))?(#(.*))?");
        var match = url.match(regEx);
        if (!match || match.length < 6) {
            throw "Valid url required";
        }
        var urlComponents = {
            Protocol: match[1],
            HostNameAndPort: match[4],
            AbsolutePath: match[5]
        };
        var pathSegments = urlComponents.AbsolutePath.split("/");
        pathSegments = pathSegments.filter(function (val) { return val && val.length > 0; }); // remove empty elements
        urlComponents.PathSegments = pathSegments;
        if (match[6]) {
            urlComponents.Search = match[6];
        }
        if (match[8]) {
            urlComponents.Hash = match[8];
        }
        return urlComponents;
    };
    /**
     * Given a url or path, append a trailing slash if one doesnt exist
     *
     * @param url
     */
    UrlUtils.CanonicalizeUri = function (url) {
        if (url) {
            url = url.toLowerCase();
        }
        if (url && !UrlUtils.endsWith(url, "/")) {
            url += "/";
        }
        return url;
    };
    /**
     * Checks to see if the url ends with the suffix
     * Required because we are compiling for es5 instead of es6
     * @param url
     * @param str
     */
    // TODO: Rename this, not clear what it is supposed to do
    UrlUtils.endsWith = function (url, suffix) {
        if (!url || !suffix) {
            return false;
        }
        return url.indexOf(suffix, url.length - suffix.length) !== -1;
    };
    /**
     * Utils function to remove the login_hint and domain_hint from the i/p extraQueryParameters
     * @param url
     * @param name
     */
    UrlUtils.urlRemoveQueryStringParameter = function (url, name) {
        if (StringUtils_1.StringUtils.isEmpty(url)) {
            return url;
        }
        var regex = new RegExp("(\\&" + name + "=)[^\&]+");
        url = url.replace(regex, "");
        // name=value&
        regex = new RegExp("(" + name + "=)[^\&]+&");
        url = url.replace(regex, "");
        // name=value
        regex = new RegExp("(" + name + "=)[^\&]+");
        url = url.replace(regex, "");
        return url;
    };
    /**
     * @hidden
     * @ignore
     *
     * Returns the anchor part(#) of the URL
     */
    UrlUtils.getHashFromUrl = function (urlStringOrFragment) {
        var hashIndex1 = urlStringOrFragment.indexOf("#");
        var hashIndex2 = urlStringOrFragment.indexOf("#/");
        if (hashIndex2 > -1) {
            return urlStringOrFragment.substring(hashIndex2 + 2);
        }
        else if (hashIndex1 > -1) {
            return urlStringOrFragment.substring(hashIndex1 + 1);
        }
        return urlStringOrFragment;
    };
    /**
     * @hidden
     * Check if the url contains a hash with known properties
     * @ignore
     */
    UrlUtils.urlContainsHash = function (urlString) {
        var parameters = UrlUtils.deserializeHash(urlString);
        return (parameters.hasOwnProperty(Constants_1.ServerHashParamKeys.ERROR_DESCRIPTION) ||
            parameters.hasOwnProperty(Constants_1.ServerHashParamKeys.ERROR) ||
            parameters.hasOwnProperty(Constants_1.ServerHashParamKeys.ACCESS_TOKEN) ||
            parameters.hasOwnProperty(Constants_1.ServerHashParamKeys.ID_TOKEN));
    };
    /**
     * @hidden
     * Returns deserialized portion of URL hash
     * @ignore
     */
    UrlUtils.deserializeHash = function (urlFragment) {
        var hash = UrlUtils.getHashFromUrl(urlFragment);
        return CryptoUtils_1.CryptoUtils.deserialize(hash);
    };
    /**
     * @ignore
     * @param {string} URI
     * @returns {string} host from the URI
     *
     * extract URI from the host
     */
    UrlUtils.getHostFromUri = function (uri) {
        // remove http:// or https:// from uri
        var extractedUri = String(uri).replace(/^(https?:)\/\//, "");
        extractedUri = extractedUri.split("/")[0];
        return extractedUri;
    };
    return UrlUtils;
}());
exports.UrlUtils = UrlUtils;


/***/ }),
/* 5 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = __webpack_require__(0);
var ClientAuthError_1 = __webpack_require__(6);
;
exports.ClientConfigurationErrorMessage = {
    configurationNotSet: {
        code: "no_config_set",
        desc: "Configuration has not been set. Please call the UserAgentApplication constructor with a valid Configuration object."
    },
    storageNotSupported: {
        code: "storage_not_supported",
        desc: "The value for the cacheLocation is not supported."
    },
    noRedirectCallbacksSet: {
        code: "no_redirect_callbacks",
        desc: "No redirect callbacks have been set. Please call handleRedirectCallback() with the appropriate function arguments before continuing. " +
            "More information is available here: https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki/MSAL-basics."
    },
    invalidCallbackObject: {
        code: "invalid_callback_object",
        desc: "The object passed for the callback was invalid. " +
            "More information is available here: https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki/MSAL-basics."
    },
    scopesRequired: {
        code: "scopes_required",
        desc: "Scopes are required to obtain an access token."
    },
    emptyScopes: {
        code: "empty_input_scopes_error",
        desc: "Scopes cannot be passed as empty array."
    },
    nonArrayScopes: {
        code: "nonarray_input_scopes_error",
        desc: "Scopes cannot be passed as non-array."
    },
    clientScope: {
        code: "clientid_input_scopes_error",
        desc: "Client ID can only be provided as a single scope."
    },
    invalidPrompt: {
        code: "invalid_prompt_value",
        desc: "Supported prompt values are 'login', 'select_account', 'consent' and 'none'",
    },
    invalidAuthorityType: {
        code: "invalid_authority_type",
        desc: "The given authority is not a valid type of authority supported by MSAL. Please see here for valid authorities: <insert URL here>."
    },
    authorityUriInsecure: {
        code: "authority_uri_insecure",
        desc: "Authority URIs must use https."
    },
    authorityUriInvalidPath: {
        code: "authority_uri_invalid_path",
        desc: "Given authority URI is invalid."
    },
    unsupportedAuthorityValidation: {
        code: "unsupported_authority_validation",
        desc: "The authority validation is not supported for this authority type."
    },
    untrustedAuthority: {
        code: "untrusted_authority",
        desc: "The provided authority is not a trusted authority. Please include this authority in the knownAuthorities config parameter or set validateAuthority=false."
    },
    b2cAuthorityUriInvalidPath: {
        code: "b2c_authority_uri_invalid_path",
        desc: "The given URI for the B2C authority is invalid."
    },
    b2cKnownAuthoritiesNotSet: {
        code: "b2c_known_authorities_not_set",
        desc: "Must set known authorities when validateAuthority is set to True and using B2C"
    },
    claimsRequestParsingError: {
        code: "claims_request_parsing_error",
        desc: "Could not parse the given claims request object."
    },
    emptyRequestError: {
        code: "empty_request_error",
        desc: "Request object is required."
    },
    invalidCorrelationIdError: {
        code: "invalid_guid_sent_as_correlationId",
        desc: "Please set the correlationId as a valid guid"
    },
    telemetryConfigError: {
        code: "telemetry_config_error",
        desc: "Telemetry config is not configured with required values"
    },
    ssoSilentError: {
        code: "sso_silent_error",
        desc: "request must contain either sid or login_hint"
    },
    invalidAuthorityMetadataError: {
        code: "authority_metadata_error",
        desc: "Invalid authorityMetadata. Must be a JSON object containing authorization_endpoint, end_session_endpoint, and issuer fields."
    }
};
/**
 * Error thrown when there is an error in configuration of the .js library.
 */
var ClientConfigurationError = /** @class */ (function (_super) {
    tslib_1.__extends(ClientConfigurationError, _super);
    function ClientConfigurationError(errorCode, errorMessage) {
        var _this = _super.call(this, errorCode, errorMessage) || this;
        _this.name = "ClientConfigurationError";
        Object.setPrototypeOf(_this, ClientConfigurationError.prototype);
        return _this;
    }
    ClientConfigurationError.createNoSetConfigurationError = function () {
        return new ClientConfigurationError(exports.ClientConfigurationErrorMessage.configurationNotSet.code, "" + exports.ClientConfigurationErrorMessage.configurationNotSet.desc);
    };
    ClientConfigurationError.createStorageNotSupportedError = function (givenCacheLocation) {
        return new ClientConfigurationError(exports.ClientConfigurationErrorMessage.storageNotSupported.code, exports.ClientConfigurationErrorMessage.storageNotSupported.desc + " Given location: " + givenCacheLocation);
    };
    ClientConfigurationError.createRedirectCallbacksNotSetError = function () {
        return new ClientConfigurationError(exports.ClientConfigurationErrorMessage.noRedirectCallbacksSet.code, exports.ClientConfigurationErrorMessage.noRedirectCallbacksSet.desc);
    };
    ClientConfigurationError.createInvalidCallbackObjectError = function (callbackObject) {
        return new ClientConfigurationError(exports.ClientConfigurationErrorMessage.invalidCallbackObject.code, exports.ClientConfigurationErrorMessage.invalidCallbackObject.desc + " Given value for callback function: " + callbackObject);
    };
    ClientConfigurationError.createEmptyScopesArrayError = function (scopesValue) {
        return new ClientConfigurationError(exports.ClientConfigurationErrorMessage.emptyScopes.code, exports.ClientConfigurationErrorMessage.emptyScopes.desc + " Given value: " + scopesValue + ".");
    };
    ClientConfigurationError.createScopesNonArrayError = function (scopesValue) {
        return new ClientConfigurationError(exports.ClientConfigurationErrorMessage.nonArrayScopes.code, exports.ClientConfigurationErrorMessage.nonArrayScopes.desc + " Given value: " + scopesValue + ".");
    };
    ClientConfigurationError.createClientIdSingleScopeError = function (scopesValue) {
        return new ClientConfigurationError(exports.ClientConfigurationErrorMessage.clientScope.code, exports.ClientConfigurationErrorMessage.clientScope.desc + " Given value: " + scopesValue + ".");
    };
    ClientConfigurationError.createScopesRequiredError = function (scopesValue) {
        return new ClientConfigurationError(exports.ClientConfigurationErrorMessage.scopesRequired.code, exports.ClientConfigurationErrorMessage.scopesRequired.desc + " Given value: " + scopesValue);
    };
    ClientConfigurationError.createInvalidPromptError = function (promptValue) {
        return new ClientConfigurationError(exports.ClientConfigurationErrorMessage.invalidPrompt.code, exports.ClientConfigurationErrorMessage.invalidPrompt.desc + " Given value: " + promptValue);
    };
    ClientConfigurationError.createClaimsRequestParsingError = function (claimsRequestParseError) {
        return new ClientConfigurationError(exports.ClientConfigurationErrorMessage.claimsRequestParsingError.code, exports.ClientConfigurationErrorMessage.claimsRequestParsingError.desc + " Given value: " + claimsRequestParseError);
    };
    ClientConfigurationError.createEmptyRequestError = function () {
        var _a = exports.ClientConfigurationErrorMessage.emptyRequestError, code = _a.code, desc = _a.desc;
        return new ClientConfigurationError(code, desc);
    };
    ClientConfigurationError.createInvalidCorrelationIdError = function () {
        return new ClientConfigurationError(exports.ClientConfigurationErrorMessage.invalidCorrelationIdError.code, exports.ClientConfigurationErrorMessage.invalidCorrelationIdError.desc);
    };
    ClientConfigurationError.createKnownAuthoritiesNotSetError = function () {
        return new ClientConfigurationError(exports.ClientConfigurationErrorMessage.b2cKnownAuthoritiesNotSet.code, exports.ClientConfigurationErrorMessage.b2cKnownAuthoritiesNotSet.desc);
    };
    ClientConfigurationError.createInvalidAuthorityTypeError = function () {
        return new ClientConfigurationError(exports.ClientConfigurationErrorMessage.invalidAuthorityType.code, exports.ClientConfigurationErrorMessage.invalidAuthorityType.desc);
    };
    ClientConfigurationError.createUntrustedAuthorityError = function (host) {
        return new ClientConfigurationError(exports.ClientConfigurationErrorMessage.untrustedAuthority.code, exports.ClientConfigurationErrorMessage.untrustedAuthority.desc + " Provided Authority: " + host);
    };
    ClientConfigurationError.createTelemetryConfigError = function (config) {
        var _a = exports.ClientConfigurationErrorMessage.telemetryConfigError, code = _a.code, desc = _a.desc;
        var requiredKeys = {
            applicationName: "string",
            applicationVersion: "string",
            telemetryEmitter: "function"
        };
        var missingKeys = Object.keys(requiredKeys)
            .reduce(function (keys, key) {
            return config[key] ? keys : keys.concat([key + " (" + requiredKeys[key] + ")"]);
        }, []);
        return new ClientConfigurationError(code, desc + " mising values: " + missingKeys.join(","));
    };
    ClientConfigurationError.createSsoSilentError = function () {
        return new ClientConfigurationError(exports.ClientConfigurationErrorMessage.ssoSilentError.code, exports.ClientConfigurationErrorMessage.ssoSilentError.desc);
    };
    ClientConfigurationError.createInvalidAuthorityMetadataError = function () {
        return new ClientConfigurationError(exports.ClientConfigurationErrorMessage.invalidAuthorityMetadataError.code, exports.ClientConfigurationErrorMessage.invalidAuthorityMetadataError.desc);
    };
    return ClientConfigurationError;
}(ClientAuthError_1.ClientAuthError));
exports.ClientConfigurationError = ClientConfigurationError;


/***/ }),
/* 6 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = __webpack_require__(0);
var AuthError_1 = __webpack_require__(7);
var StringUtils_1 = __webpack_require__(3);
exports.ClientAuthErrorMessage = {
    multipleMatchingTokens: {
        code: "multiple_matching_tokens",
        desc: "The cache contains multiple tokens satisfying the requirements. " +
            "Call AcquireToken again providing more requirements like authority."
    },
    multipleCacheAuthorities: {
        code: "multiple_authorities",
        desc: "Multiple authorities found in the cache. Pass authority in the API overload."
    },
    endpointResolutionError: {
        code: "endpoints_resolution_error",
        desc: "Error: could not resolve endpoints. Please check network and try again."
    },
    popUpWindowError: {
        code: "popup_window_error",
        desc: "Error opening popup window. This can happen if you are using IE or if popups are blocked in the browser."
    },
    tokenRenewalError: {
        code: "token_renewal_error",
        desc: "Token renewal operation failed due to timeout."
    },
    invalidIdToken: {
        code: "invalid_id_token",
        desc: "Invalid ID token format."
    },
    invalidStateError: {
        code: "invalid_state_error",
        desc: "Invalid state."
    },
    nonceMismatchError: {
        code: "nonce_mismatch_error",
        desc: "Nonce is not matching, Nonce received: "
    },
    loginProgressError: {
        code: "login_progress_error",
        desc: "Login_In_Progress: Error during login call - login is already in progress."
    },
    acquireTokenProgressError: {
        code: "acquiretoken_progress_error",
        desc: "AcquireToken_In_Progress: Error during login call - login is already in progress."
    },
    userCancelledError: {
        code: "user_cancelled",
        desc: "User cancelled the flow."
    },
    callbackError: {
        code: "callback_error",
        desc: "Error occurred in token received callback function."
    },
    userLoginRequiredError: {
        code: "user_login_error",
        desc: "User login is required. For silent calls, request must contain either sid or login_hint"
    },
    userDoesNotExistError: {
        code: "user_non_existent",
        desc: "User object does not exist. Please call a login API."
    },
    clientInfoDecodingError: {
        code: "client_info_decoding_error",
        desc: "The client info could not be parsed/decoded correctly. Please review the trace to determine the root cause."
    },
    clientInfoNotPopulatedError: {
        code: "client_info_not_populated_error",
        desc: "The service did not populate client_info in the response, Please verify with the service team"
    },
    nullOrEmptyIdToken: {
        code: "null_or_empty_id_token",
        desc: "The idToken is null or empty. Please review the trace to determine the root cause."
    },
    idTokenNotParsed: {
        code: "id_token_parsing_error",
        desc: "ID token cannot be parsed. Please review stack trace to determine root cause."
    },
    tokenEncodingError: {
        code: "token_encoding_error",
        desc: "The token to be decoded is not encoded correctly."
    },
    invalidInteractionType: {
        code: "invalid_interaction_type",
        desc: "The interaction type passed to the handler was incorrect or unknown"
    },
    cacheParseError: {
        code: "cannot_parse_cache",
        desc: "The cached token key is not a valid JSON and cannot be parsed"
    },
    blockTokenRequestsInHiddenIframe: {
        code: "block_token_requests",
        desc: "Token calls are blocked in hidden iframes"
    }
};
/**
 * Error thrown when there is an error in the client code running on the browser.
 */
var ClientAuthError = /** @class */ (function (_super) {
    tslib_1.__extends(ClientAuthError, _super);
    function ClientAuthError(errorCode, errorMessage) {
        var _this = _super.call(this, errorCode, errorMessage) || this;
        _this.name = "ClientAuthError";
        Object.setPrototypeOf(_this, ClientAuthError.prototype);
        return _this;
    }
    ClientAuthError.createEndpointResolutionError = function (errDetail) {
        var errorMessage = exports.ClientAuthErrorMessage.endpointResolutionError.desc;
        if (errDetail && !StringUtils_1.StringUtils.isEmpty(errDetail)) {
            errorMessage += " Details: " + errDetail;
        }
        return new ClientAuthError(exports.ClientAuthErrorMessage.endpointResolutionError.code, errorMessage);
    };
    ClientAuthError.createMultipleMatchingTokensInCacheError = function (scope) {
        return new ClientAuthError(exports.ClientAuthErrorMessage.multipleMatchingTokens.code, "Cache error for scope " + scope + ": " + exports.ClientAuthErrorMessage.multipleMatchingTokens.desc + ".");
    };
    ClientAuthError.createMultipleAuthoritiesInCacheError = function (scope) {
        return new ClientAuthError(exports.ClientAuthErrorMessage.multipleCacheAuthorities.code, "Cache error for scope " + scope + ": " + exports.ClientAuthErrorMessage.multipleCacheAuthorities.desc + ".");
    };
    ClientAuthError.createPopupWindowError = function (errDetail) {
        var errorMessage = exports.ClientAuthErrorMessage.popUpWindowError.desc;
        if (errDetail && !StringUtils_1.StringUtils.isEmpty(errDetail)) {
            errorMessage += " Details: " + errDetail;
        }
        return new ClientAuthError(exports.ClientAuthErrorMessage.popUpWindowError.code, errorMessage);
    };
    ClientAuthError.createTokenRenewalTimeoutError = function () {
        return new ClientAuthError(exports.ClientAuthErrorMessage.tokenRenewalError.code, exports.ClientAuthErrorMessage.tokenRenewalError.desc);
    };
    ClientAuthError.createInvalidIdTokenError = function (idToken) {
        return new ClientAuthError(exports.ClientAuthErrorMessage.invalidIdToken.code, exports.ClientAuthErrorMessage.invalidIdToken.desc + " Given token: " + idToken);
    };
    // TODO: Is this not a security flaw to send the user the state expected??
    ClientAuthError.createInvalidStateError = function (invalidState, actualState) {
        return new ClientAuthError(exports.ClientAuthErrorMessage.invalidStateError.code, exports.ClientAuthErrorMessage.invalidStateError.desc + " " + invalidState + ", state expected : " + actualState + ".");
    };
    // TODO: Is this not a security flaw to send the user the Nonce expected??
    ClientAuthError.createNonceMismatchError = function (invalidNonce, actualNonce) {
        return new ClientAuthError(exports.ClientAuthErrorMessage.nonceMismatchError.code, exports.ClientAuthErrorMessage.nonceMismatchError.desc + " " + invalidNonce + ", nonce expected : " + actualNonce + ".");
    };
    ClientAuthError.createLoginInProgressError = function () {
        return new ClientAuthError(exports.ClientAuthErrorMessage.loginProgressError.code, exports.ClientAuthErrorMessage.loginProgressError.desc);
    };
    ClientAuthError.createAcquireTokenInProgressError = function () {
        return new ClientAuthError(exports.ClientAuthErrorMessage.acquireTokenProgressError.code, exports.ClientAuthErrorMessage.acquireTokenProgressError.desc);
    };
    ClientAuthError.createUserCancelledError = function () {
        return new ClientAuthError(exports.ClientAuthErrorMessage.userCancelledError.code, exports.ClientAuthErrorMessage.userCancelledError.desc);
    };
    ClientAuthError.createErrorInCallbackFunction = function (errorDesc) {
        return new ClientAuthError(exports.ClientAuthErrorMessage.callbackError.code, exports.ClientAuthErrorMessage.callbackError.desc + " " + errorDesc + ".");
    };
    ClientAuthError.createUserLoginRequiredError = function () {
        return new ClientAuthError(exports.ClientAuthErrorMessage.userLoginRequiredError.code, exports.ClientAuthErrorMessage.userLoginRequiredError.desc);
    };
    ClientAuthError.createUserDoesNotExistError = function () {
        return new ClientAuthError(exports.ClientAuthErrorMessage.userDoesNotExistError.code, exports.ClientAuthErrorMessage.userDoesNotExistError.desc);
    };
    ClientAuthError.createClientInfoDecodingError = function (caughtError) {
        return new ClientAuthError(exports.ClientAuthErrorMessage.clientInfoDecodingError.code, exports.ClientAuthErrorMessage.clientInfoDecodingError.desc + " Failed with error: " + caughtError);
    };
    ClientAuthError.createClientInfoNotPopulatedError = function (caughtError) {
        return new ClientAuthError(exports.ClientAuthErrorMessage.clientInfoNotPopulatedError.code, exports.ClientAuthErrorMessage.clientInfoNotPopulatedError.desc + " Failed with error: " + caughtError);
    };
    ClientAuthError.createIdTokenNullOrEmptyError = function (invalidRawTokenString) {
        return new ClientAuthError(exports.ClientAuthErrorMessage.nullOrEmptyIdToken.code, exports.ClientAuthErrorMessage.nullOrEmptyIdToken.desc + " Raw ID Token Value: " + invalidRawTokenString);
    };
    ClientAuthError.createIdTokenParsingError = function (caughtParsingError) {
        return new ClientAuthError(exports.ClientAuthErrorMessage.idTokenNotParsed.code, exports.ClientAuthErrorMessage.idTokenNotParsed.desc + " Failed with error: " + caughtParsingError);
    };
    ClientAuthError.createTokenEncodingError = function (incorrectlyEncodedToken) {
        return new ClientAuthError(exports.ClientAuthErrorMessage.tokenEncodingError.code, exports.ClientAuthErrorMessage.tokenEncodingError.desc + " Attempted to decode: " + incorrectlyEncodedToken);
    };
    ClientAuthError.createInvalidInteractionTypeError = function () {
        return new ClientAuthError(exports.ClientAuthErrorMessage.invalidInteractionType.code, exports.ClientAuthErrorMessage.invalidInteractionType.desc);
    };
    ClientAuthError.createCacheParseError = function (key) {
        var errorMessage = "invalid key: " + key + ", " + exports.ClientAuthErrorMessage.cacheParseError.desc;
        return new ClientAuthError(exports.ClientAuthErrorMessage.cacheParseError.code, errorMessage);
    };
    ClientAuthError.createBlockTokenRequestsInHiddenIframeError = function () {
        return new ClientAuthError(exports.ClientAuthErrorMessage.blockTokenRequestsInHiddenIframe.code, exports.ClientAuthErrorMessage.blockTokenRequestsInHiddenIframe.desc);
    };
    return ClientAuthError;
}(AuthError_1.AuthError));
exports.ClientAuthError = ClientAuthError;


/***/ }),
/* 7 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = __webpack_require__(0);
exports.AuthErrorMessage = {
    unexpectedError: {
        code: "unexpected_error",
        desc: "Unexpected error in authentication."
    },
    noWindowObjectError: {
        code: "no_window_object",
        desc: "No window object available. Details:"
    }
};
/**
 * General error class thrown by the MSAL.js library.
 */
var AuthError = /** @class */ (function (_super) {
    tslib_1.__extends(AuthError, _super);
    function AuthError(errorCode, errorMessage) {
        var _this = _super.call(this, errorMessage) || this;
        Object.setPrototypeOf(_this, AuthError.prototype);
        _this.errorCode = errorCode;
        _this.errorMessage = errorMessage;
        _this.name = "AuthError";
        return _this;
    }
    AuthError.createUnexpectedError = function (errDesc) {
        return new AuthError(exports.AuthErrorMessage.unexpectedError.code, exports.AuthErrorMessage.unexpectedError.desc + ": " + errDesc);
    };
    AuthError.createNoWindowObjectError = function (errDesc) {
        return new AuthError(exports.AuthErrorMessage.noWindowObjectError.code, exports.AuthErrorMessage.noWindowObjectError.desc + " " + errDesc);
    };
    return AuthError;
}(Error));
exports.AuthError = AuthError;


/***/ }),
/* 8 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
exports.EVENT_NAME_PREFIX = "msal.";
exports.EVENT_NAME_KEY = "event_name";
exports.START_TIME_KEY = "start_time";
exports.ELAPSED_TIME_KEY = "elapsed_time";
exports.TELEMETRY_BLOB_EVENT_NAMES = {
    MsalCorrelationIdConstStrKey: "Microsoft.MSAL.correlation_id",
    ApiTelemIdConstStrKey: "msal.api_telem_id",
    ApiIdConstStrKey: "msal.api_id",
    BrokerAppConstStrKey: "Microsoft_MSAL_broker_app",
    CacheEventCountConstStrKey: "Microsoft_MSAL_cache_event_count",
    HttpEventCountTelemetryBatchKey: "Microsoft_MSAL_http_event_count",
    IdpConstStrKey: "Microsoft_MSAL_idp",
    IsSilentTelemetryBatchKey: "",
    IsSuccessfulConstStrKey: "Microsoft_MSAL_is_successful",
    ResponseTimeConstStrKey: "Microsoft_MSAL_response_time",
    TenantIdConstStrKey: "Microsoft_MSAL_tenant_id",
    UiEventCountTelemetryBatchKey: "Microsoft_MSAL_ui_event_count"
};
// This is used to replace the real tenant in telemetry info
exports.TENANT_PLACEHOLDER = "<tenant>";


/***/ }),
/* 9 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
var ClientConfigurationError_1 = __webpack_require__(5);
var Constants_1 = __webpack_require__(1);
var ScopeSet = /** @class */ (function () {
    function ScopeSet() {
    }
    /**
     * Check if there are dup scopes in a given request
     *
     * @param cachedScopes
     * @param scopes
     */
    // TODO: Rename this, intersecting scopes isn't a great name for duplicate checker
    ScopeSet.isIntersectingScopes = function (cachedScopes, scopes) {
        var convertedCachedScopes = this.trimAndConvertArrayToLowerCase(cachedScopes.slice());
        var requestScopes = this.trimAndConvertArrayToLowerCase(scopes.slice());
        for (var i = 0; i < requestScopes.length; i++) {
            if (convertedCachedScopes.indexOf(requestScopes[i].toLowerCase()) > -1) {
                return true;
            }
        }
        return false;
    };
    /**
     * Check if a given scope is present in the request
     *
     * @param cachedScopes
     * @param scopes
     */
    ScopeSet.containsScope = function (cachedScopes, scopes) {
        var convertedCachedScopes = this.trimAndConvertArrayToLowerCase(cachedScopes.slice());
        var requestScopes = this.trimAndConvertArrayToLowerCase(scopes.slice());
        return requestScopes.every(function (value) { return convertedCachedScopes.indexOf(value.toString().toLowerCase()) >= 0; });
    };
    /**
     *  Trims and converts string to lower case
     *
     * @param scopes
     */
    // TODO: Rename this, too generic name for a function that only deals with scopes
    ScopeSet.trimAndConvertToLowerCase = function (scope) {
        return scope.trim().toLowerCase();
    };
    /**
     * Performs trimeAndConvertToLowerCase on string array
     * @param scopes
     */
    ScopeSet.trimAndConvertArrayToLowerCase = function (scopes) {
        var _this = this;
        return scopes.map(function (scope) { return _this.trimAndConvertToLowerCase(scope); });
    };
    /**
     * remove one element from a scope array
     *
     * @param scopes
     * @param scope
     */
    // TODO: Rename this, too generic name for a function that only deals with scopes
    ScopeSet.removeElement = function (scopes, scope) {
        var scopeVal = this.trimAndConvertToLowerCase(scope);
        return scopes.filter(function (value) { return value !== scopeVal; });
    };
    /**
     * Parse the scopes into a formatted scopeList
     * @param scopes
     */
    ScopeSet.parseScope = function (scopes) {
        var scopeList = "";
        if (scopes) {
            for (var i = 0; i < scopes.length; ++i) {
                scopeList += (i !== scopes.length - 1) ? scopes[i] + " " : scopes[i];
            }
        }
        return scopeList;
    };
    /**
     * @hidden
     *
     * Used to validate the scopes input parameter requested  by the developer.
     * @param {Array<string>} scopes - Developer requested permissions. Not all scopes are guaranteed to be included in the access token returned.
     * @param {boolean} scopesRequired - Boolean indicating whether the scopes array is required or not
     * @ignore
     */
    ScopeSet.validateInputScope = function (scopes, scopesRequired, clientId) {
        if (!scopes) {
            if (scopesRequired) {
                throw ClientConfigurationError_1.ClientConfigurationError.createScopesRequiredError(scopes);
            }
            else {
                return;
            }
        }
        // Check that scopes is an array object (also throws error if scopes == null)
        if (!Array.isArray(scopes)) {
            throw ClientConfigurationError_1.ClientConfigurationError.createScopesNonArrayError(scopes);
        }
        // Check that scopes is not an empty array
        if (scopes.length < 1) {
            throw ClientConfigurationError_1.ClientConfigurationError.createEmptyScopesArrayError(scopes.toString());
        }
        // Check that clientId is passed as single scope
        if (scopes.indexOf(clientId) > -1) {
            if (scopes.length > 1) {
                throw ClientConfigurationError_1.ClientConfigurationError.createClientIdSingleScopeError(scopes.toString());
            }
        }
    };
    /**
     * @hidden
     *
     * Extracts scope value from the state sent with the authentication request.
     * @param {string} state
     * @returns {string} scope.
     * @ignore
     */
    ScopeSet.getScopeFromState = function (state) {
        if (state) {
            var splitIndex = state.indexOf(Constants_1.Constants.resourceDelimiter);
            if (splitIndex > -1 && splitIndex + 1 < state.length) {
                return state.substring(splitIndex + 1);
            }
        }
        return "";
    };
    /**
     * @ignore
     * Appends extraScopesToConsent if passed
     * @param {@link AuthenticationParameters}
     */
    ScopeSet.appendScopes = function (reqScopes, reqExtraScopesToConsent) {
        if (reqScopes) {
            var convertedExtraScopes = reqExtraScopesToConsent ? this.trimAndConvertArrayToLowerCase(reqExtraScopesToConsent.slice()) : null;
            var convertedReqScopes = this.trimAndConvertArrayToLowerCase(reqScopes.slice());
            return convertedExtraScopes ? convertedReqScopes.concat(convertedExtraScopes) : convertedReqScopes;
        }
        return null;
    };
    return ScopeSet;
}());
exports.ScopeSet = ScopeSet;


/***/ }),
/* 10 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var TelemetryConstants_1 = __webpack_require__(8);
var CryptoUtils_1 = __webpack_require__(2);
var UrlUtils_1 = __webpack_require__(4);
var AuthorityFactory_1 = __webpack_require__(21);
exports.scrubTenantFromUri = function (uri) {
    var url = UrlUtils_1.UrlUtils.GetUrlComponents(uri);
    // validate trusted host
    if (AuthorityFactory_1.AuthorityFactory.isAdfs(uri)) {
        /**
         * returning what was passed because the library needs to work with uris that are non
         * AAD trusted but passed by users such as B2C or others.
         * HTTP Events for instance can take a url to the open id config endpoint
         */
        return uri;
    }
    var pathParams = url.PathSegments;
    if (pathParams && pathParams.length >= 2) {
        var tenantPosition = pathParams[1] === "tfp" ? 2 : 1;
        if (tenantPosition < pathParams.length) {
            pathParams[tenantPosition] = TelemetryConstants_1.TENANT_PLACEHOLDER;
        }
    }
    return url.Protocol + "//" + url.HostNameAndPort + "/" + pathParams.join("/");
};
exports.hashPersonalIdentifier = function (valueToHash) {
    /*
     * TODO sha256 this
     * Current test runner is being funny with node libs that are webpacked anyway
     * need a different solution
     */
    return CryptoUtils_1.CryptoUtils.base64Encode(valueToHash);
};
exports.prependEventNamePrefix = function (suffix) { return "" + TelemetryConstants_1.EVENT_NAME_PREFIX + (suffix || ""); };
exports.supportsBrowserPerformance = function () { return !!(typeof window !== "undefined" &&
    "performance" in window &&
    window.performance.mark &&
    window.performance.measure); };
exports.endBrowserPerformanceMeasurement = function (measureName, startMark, endMark) {
    if (exports.supportsBrowserPerformance()) {
        window.performance.mark(endMark);
        window.performance.measure(measureName, startMark, endMark);
        window.performance.clearMeasures(measureName);
        window.performance.clearMarks(startMark);
        window.performance.clearMarks(endMark);
    }
};
exports.startBrowserPerformanceMeasurement = function (startMark) {
    if (exports.supportsBrowserPerformance()) {
        window.performance.mark(startMark);
    }
};


/***/ }),
/* 11 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
/**
 * @hidden
 */
var TimeUtils = /** @class */ (function () {
    function TimeUtils() {
    }
    /**
     * Returns time in seconds for expiration based on string value passed in.
     *
     * @param expiresIn
     */
    TimeUtils.parseExpiresIn = function (expiresIn) {
        // if AAD did not send "expires_in" property, use default expiration of 3599 seconds, for some reason AAD sends 3599 as "expires_in" value instead of 3600
        if (!expiresIn) {
            expiresIn = "3599";
        }
        return parseInt(expiresIn, 10);
    };
    /**
     * Return the current time in Unix time (seconds). Date.getTime() returns in milliseconds.
     */
    TimeUtils.now = function () {
        return Math.round(new Date().getTime() / 1000.0);
    };
    /**
     * Returns the amount of time in milliseconds since the page loaded.
     */
    TimeUtils.relativeNowMs = function () {
        return window.performance.now();
    };
    return TimeUtils;
}());
exports.TimeUtils = TimeUtils;


/***/ }),
/* 12 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
var StringUtils_1 = __webpack_require__(3);
var Constants_1 = __webpack_require__(1);
var LogLevel;
(function (LogLevel) {
    LogLevel[LogLevel["Error"] = 0] = "Error";
    LogLevel[LogLevel["Warning"] = 1] = "Warning";
    LogLevel[LogLevel["Info"] = 2] = "Info";
    LogLevel[LogLevel["Verbose"] = 3] = "Verbose";
})(LogLevel = exports.LogLevel || (exports.LogLevel = {}));
var Logger = /** @class */ (function () {
    function Logger(localCallback, options) {
        if (options === void 0) { options = {}; }
        /**
         * @hidden
         */
        this.level = LogLevel.Info;
        var _a = options.correlationId, correlationId = _a === void 0 ? "" : _a, _b = options.level, level = _b === void 0 ? LogLevel.Info : _b, _c = options.piiLoggingEnabled, piiLoggingEnabled = _c === void 0 ? false : _c;
        this.localCallback = localCallback;
        this.correlationId = correlationId;
        this.level = level;
        this.piiLoggingEnabled = piiLoggingEnabled;
    }
    /**
     * @hidden
     */
    Logger.prototype.logMessage = function (logLevel, logMessage, containsPii) {
        if ((logLevel > this.level) || (!this.piiLoggingEnabled && containsPii)) {
            return;
        }
        var timestamp = new Date().toUTCString();
        var log;
        if (!StringUtils_1.StringUtils.isEmpty(this.correlationId)) {
            log = timestamp + ":" + this.correlationId + "-" + Constants_1.libraryVersion() + "-" + LogLevel[logLevel] + (containsPii ? "-pii" : "") + " " + logMessage;
        }
        else {
            log = timestamp + ":" + Constants_1.libraryVersion() + "-" + LogLevel[logLevel] + (containsPii ? "-pii" : "") + " " + logMessage;
        }
        this.executeCallback(logLevel, log, containsPii);
    };
    /**
     * @hidden
     */
    Logger.prototype.executeCallback = function (level, message, containsPii) {
        if (this.localCallback) {
            this.localCallback(level, message, containsPii);
        }
    };
    /**
     * @hidden
     */
    Logger.prototype.error = function (message) {
        this.logMessage(LogLevel.Error, message, false);
    };
    /**
     * @hidden
     */
    Logger.prototype.errorPii = function (message) {
        this.logMessage(LogLevel.Error, message, true);
    };
    /**
     * @hidden
     */
    Logger.prototype.warning = function (message) {
        this.logMessage(LogLevel.Warning, message, false);
    };
    /**
     * @hidden
     */
    Logger.prototype.warningPii = function (message) {
        this.logMessage(LogLevel.Warning, message, true);
    };
    /**
     * @hidden
     */
    Logger.prototype.info = function (message) {
        this.logMessage(LogLevel.Info, message, false);
    };
    /**
     * @hidden
     */
    Logger.prototype.infoPii = function (message) {
        this.logMessage(LogLevel.Info, message, true);
    };
    /**
     * @hidden
     */
    Logger.prototype.verbose = function (message) {
        this.logMessage(LogLevel.Verbose, message, false);
    };
    /**
     * @hidden
     */
    Logger.prototype.verbosePii = function (message) {
        this.logMessage(LogLevel.Verbose, message, true);
    };
    Logger.prototype.isPiiLoggingEnabled = function () {
        return this.piiLoggingEnabled;
    };
    return Logger;
}());
exports.Logger = Logger;


/***/ }),
/* 13 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = __webpack_require__(0);
var AuthError_1 = __webpack_require__(7);
exports.ServerErrorMessage = {
    serverUnavailable: {
        code: "server_unavailable",
        desc: "Server is temporarily unavailable."
    },
    unknownServerError: {
        code: "unknown_server_error"
    },
};
/**
 * Error thrown when there is an error with the server code, for example, unavailability.
 */
var ServerError = /** @class */ (function (_super) {
    tslib_1.__extends(ServerError, _super);
    function ServerError(errorCode, errorMessage) {
        var _this = _super.call(this, errorCode, errorMessage) || this;
        _this.name = "ServerError";
        Object.setPrototypeOf(_this, ServerError.prototype);
        return _this;
    }
    ServerError.createServerUnavailableError = function () {
        return new ServerError(exports.ServerErrorMessage.serverUnavailable.code, exports.ServerErrorMessage.serverUnavailable.desc);
    };
    ServerError.createUnknownServerError = function (errorDesc) {
        return new ServerError(exports.ServerErrorMessage.unknownServerError.code, errorDesc);
    };
    return ServerError;
}(AuthError_1.AuthError));
exports.ServerError = ServerError;


/***/ }),
/* 14 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = __webpack_require__(0);
var TelemetryConstants_1 = __webpack_require__(8);
var TelemetryConstants_2 = __webpack_require__(8);
var TelemetryUtils_1 = __webpack_require__(10);
var CryptoUtils_1 = __webpack_require__(2);
var TelemetryEvent = /** @class */ (function () {
    function TelemetryEvent(eventName, correlationId, eventLabel) {
        var _a;
        this.eventId = CryptoUtils_1.CryptoUtils.createNewGuid();
        this.label = eventLabel;
        this.event = (_a = {},
            _a[TelemetryUtils_1.prependEventNamePrefix(TelemetryConstants_2.EVENT_NAME_KEY)] = eventName,
            _a[TelemetryUtils_1.prependEventNamePrefix(TelemetryConstants_2.ELAPSED_TIME_KEY)] = -1,
            _a["" + TelemetryConstants_1.TELEMETRY_BLOB_EVENT_NAMES.MsalCorrelationIdConstStrKey] = correlationId,
            _a);
    }
    TelemetryEvent.prototype.setElapsedTime = function (time) {
        this.event[TelemetryUtils_1.prependEventNamePrefix(TelemetryConstants_2.ELAPSED_TIME_KEY)] = time;
    };
    TelemetryEvent.prototype.stop = function () {
        // Set duration of event
        this.setElapsedTime(+Date.now() - +this.startTimestamp);
        TelemetryUtils_1.endBrowserPerformanceMeasurement(this.displayName, this.perfStartMark, this.perfEndMark);
    };
    TelemetryEvent.prototype.start = function () {
        this.startTimestamp = Date.now();
        this.event[TelemetryUtils_1.prependEventNamePrefix(TelemetryConstants_2.START_TIME_KEY)] = this.startTimestamp;
        TelemetryUtils_1.startBrowserPerformanceMeasurement(this.perfStartMark);
    };
    Object.defineProperty(TelemetryEvent.prototype, "telemetryCorrelationId", {
        get: function () {
            return this.event["" + TelemetryConstants_1.TELEMETRY_BLOB_EVENT_NAMES.MsalCorrelationIdConstStrKey];
        },
        set: function (value) {
            this.event["" + TelemetryConstants_1.TELEMETRY_BLOB_EVENT_NAMES.MsalCorrelationIdConstStrKey] = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(TelemetryEvent.prototype, "eventName", {
        get: function () {
            return this.event[TelemetryUtils_1.prependEventNamePrefix(TelemetryConstants_2.EVENT_NAME_KEY)];
        },
        enumerable: true,
        configurable: true
    });
    TelemetryEvent.prototype.get = function () {
        return tslib_1.__assign({}, this.event, { eventId: this.eventId });
    };
    Object.defineProperty(TelemetryEvent.prototype, "key", {
        get: function () {
            return this.telemetryCorrelationId + "_" + this.eventId + "-" + this.eventName;
        },
        enumerable: true,
        configurable: true
    });
    ;
    Object.defineProperty(TelemetryEvent.prototype, "displayName", {
        get: function () {
            return "Msal-" + this.label + "-" + this.telemetryCorrelationId;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(TelemetryEvent.prototype, "perfStartMark", {
        get: function () {
            return "start-" + this.key;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(TelemetryEvent.prototype, "perfEndMark", {
        get: function () {
            return "end-" + this.key;
        },
        enumerable: true,
        configurable: true
    });
    return TelemetryEvent;
}());
exports.default = TelemetryEvent;


/***/ }),
/* 15 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = __webpack_require__(0);
var AccessTokenKey_1 = __webpack_require__(31);
var AccessTokenValue_1 = __webpack_require__(32);
var ServerRequestParameters_1 = __webpack_require__(16);
var ClientInfo_1 = __webpack_require__(33);
var IdToken_1 = __webpack_require__(34);
var AuthCache_1 = __webpack_require__(35);
var Account_1 = __webpack_require__(19);
var ScopeSet_1 = __webpack_require__(9);
var StringUtils_1 = __webpack_require__(3);
var WindowUtils_1 = __webpack_require__(20);
var TokenUtils_1 = __webpack_require__(17);
var TimeUtils_1 = __webpack_require__(11);
var UrlUtils_1 = __webpack_require__(4);
var RequestUtils_1 = __webpack_require__(18);
var ResponseUtils_1 = __webpack_require__(38);
var AuthorityFactory_1 = __webpack_require__(21);
var Configuration_1 = __webpack_require__(25);
var ClientConfigurationError_1 = __webpack_require__(5);
var AuthError_1 = __webpack_require__(7);
var ClientAuthError_1 = __webpack_require__(6);
var ServerError_1 = __webpack_require__(13);
var InteractionRequiredAuthError_1 = __webpack_require__(26);
var AuthResponse_1 = __webpack_require__(27);
var TelemetryManager_1 = tslib_1.__importDefault(__webpack_require__(39));
var ApiEvent_1 = __webpack_require__(28);
var Constants_1 = __webpack_require__(1);
var CryptoUtils_1 = __webpack_require__(2);
var TrustedAuthority_1 = __webpack_require__(24);
// default authority
var DEFAULT_AUTHORITY = "https://login.microsoftonline.com/common";
/**
 * @hidden
 * @ignore
 * response_type from OpenIDConnect
 * References: https://openid.net/specs/oauth-v2-multiple-response-types-1_0.html & https://tools.ietf.org/html/rfc6749#section-4.2.1
 * Since we support only implicit flow in this library, we restrict the response_type support to only 'token' and 'id_token'
 *
 */
var ResponseTypes = {
    id_token: "id_token",
    token: "token",
    id_token_token: "id_token token"
};
/**
 * UserAgentApplication class
 *
 * Object Instance that the developer can use to make loginXX OR acquireTokenXX functions
 */
var UserAgentApplication = /** @class */ (function () {
    /**
     * @constructor
     * Constructor for the UserAgentApplication used to instantiate the UserAgentApplication object
     *
     * Important attributes in the Configuration object for auth are:
     * - clientID: the application ID of your application.
     * You can obtain one by registering your application with our Application registration portal : https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredAppsPreview
     * - authority: the authority URL for your application.
     *
     * In Azure AD, authority is a URL indicating the Azure active directory that MSAL uses to obtain tokens.
     * It is of the form https://login.microsoftonline.com/&lt;Enter_the_Tenant_Info_Here&gt;.
     * If your application supports Accounts in one organizational directory, replace "Enter_the_Tenant_Info_Here" value with the Tenant Id or Tenant name (for example, contoso.microsoft.com).
     * If your application supports Accounts in any organizational directory, replace "Enter_the_Tenant_Info_Here" value with organizations.
     * If your application supports Accounts in any organizational directory and personal Microsoft accounts, replace "Enter_the_Tenant_Info_Here" value with common.
     * To restrict support to Personal Microsoft accounts only, replace "Enter_the_Tenant_Info_Here" value with consumers.
     *
     *
     * In Azure B2C, authority is of the form https://&lt;instance&gt;/tfp/&lt;tenant&gt;/&lt;policyName&gt;/
     *
     * @param {@link (Configuration:type)} configuration object for the MSAL UserAgentApplication instance
     */
    function UserAgentApplication(configuration) {
        // callbacks for token/error
        this.authResponseCallback = null;
        this.tokenReceivedCallback = null;
        this.errorReceivedCallback = null;
        // Set the Configuration
        this.config = Configuration_1.buildConfiguration(configuration);
        this.logger = this.config.system.logger;
        this.clientId = this.config.auth.clientId;
        this.inCookie = this.config.cache.storeAuthStateInCookie;
        this.telemetryManager = this.getTelemetryManagerFromConfig(this.config.system.telemetry, this.clientId);
        TrustedAuthority_1.TrustedAuthority.setTrustedAuthoritiesFromConfig(this.config.auth.validateAuthority, this.config.auth.knownAuthorities);
        AuthorityFactory_1.AuthorityFactory.saveMetadataFromConfig(this.config.auth.authority, this.config.auth.authorityMetadata);
        // if no authority is passed, set the default: "https://login.microsoftonline.com/common"
        this.authority = this.config.auth.authority || DEFAULT_AUTHORITY;
        // cache keys msal - typescript throws an error if any value other than "localStorage" or "sessionStorage" is passed
        this.cacheStorage = new AuthCache_1.AuthCache(this.clientId, this.config.cache.cacheLocation, this.inCookie);
        // Initialize window handling code
        window.activeRenewals = {};
        window.renewStates = [];
        window.callbackMappedToRenewStates = {};
        window.promiseMappedToRenewStates = {};
        window.msal = this;
        var urlHash = window.location.hash;
        var urlContainsHash = UrlUtils_1.UrlUtils.urlContainsHash(urlHash);
        // check if back button is pressed
        WindowUtils_1.WindowUtils.checkIfBackButtonIsPressed(this.cacheStorage);
        // On the server 302 - Redirect, handle this
        if (urlContainsHash) {
            var stateInfo = this.getResponseState(urlHash);
            if (stateInfo.method === Constants_1.Constants.interactionTypeRedirect) {
                this.handleRedirectAuthenticationResponse(urlHash);
            }
        }
    }
    Object.defineProperty(UserAgentApplication.prototype, "authority", {
        /**
         * Method to manage the authority URL.
         *
         * @returns {string} authority
         */
        get: function () {
            return this.authorityInstance.CanonicalAuthority;
        },
        /**
         * setter for the authority URL
         * @param {string} authority
         */
        // If the developer passes an authority, create an instance
        set: function (val) {
            this.authorityInstance = AuthorityFactory_1.AuthorityFactory.CreateInstance(val, this.config.auth.validateAuthority);
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Get the current authority instance from the MSAL configuration object
     *
     * @returns {@link Authority} authority instance
     */
    UserAgentApplication.prototype.getAuthorityInstance = function () {
        return this.authorityInstance;
    };
    UserAgentApplication.prototype.handleRedirectCallback = function (authOrTokenCallback, errorReceivedCallback) {
        if (!authOrTokenCallback) {
            throw ClientConfigurationError_1.ClientConfigurationError.createInvalidCallbackObjectError(authOrTokenCallback);
        }
        // Set callbacks
        if (errorReceivedCallback) {
            this.tokenReceivedCallback = authOrTokenCallback;
            this.errorReceivedCallback = errorReceivedCallback;
            this.logger.warning("This overload for callback is deprecated - please change the format of the callbacks to a single callback as shown: (err: AuthError, response: AuthResponse).");
        }
        else {
            this.authResponseCallback = authOrTokenCallback;
        }
        if (this.redirectError) {
            this.authErrorHandler(Constants_1.Constants.interactionTypeRedirect, this.redirectError, this.redirectResponse);
        }
        else if (this.redirectResponse) {
            this.authResponseHandler(Constants_1.Constants.interactionTypeRedirect, this.redirectResponse);
        }
    };
    /**
     * Public API to verify if the URL contains the hash with known properties
     * @param hash
     */
    UserAgentApplication.prototype.urlContainsHash = function (hash) {
        this.logger.verbose("UrlContainsHash has been called");
        return UrlUtils_1.UrlUtils.urlContainsHash(hash);
    };
    UserAgentApplication.prototype.authResponseHandler = function (interactionType, response, resolve) {
        this.logger.verbose("AuthResponseHandler has been called");
        if (interactionType === Constants_1.Constants.interactionTypeRedirect) {
            this.logger.verbose("Interaction type is redirect");
            if (this.errorReceivedCallback) {
                this.logger.verbose("Two callbacks were provided to handleRedirectCallback, calling success callback with response");
                this.tokenReceivedCallback(response);
            }
            else if (this.authResponseCallback) {
                this.logger.verbose("One callback was provided to handleRedirectCallback, calling authResponseCallback with response");
                this.authResponseCallback(null, response);
            }
        }
        else if (interactionType === Constants_1.Constants.interactionTypePopup) {
            this.logger.verbose("Interaction type is popup, resolving");
            resolve(response);
        }
        else {
            throw ClientAuthError_1.ClientAuthError.createInvalidInteractionTypeError();
        }
    };
    UserAgentApplication.prototype.authErrorHandler = function (interactionType, authErr, response, reject) {
        this.logger.verbose("AuthErrorHandler has been called");
        // set interaction_status to complete
        this.cacheStorage.removeItem(Constants_1.TemporaryCacheKeys.INTERACTION_STATUS);
        if (interactionType === Constants_1.Constants.interactionTypeRedirect) {
            this.logger.verbose("Interaction type is redirect");
            if (this.errorReceivedCallback) {
                this.logger.verbose("Two callbacks were provided to handleRedirectCallback, calling error callback");
                this.errorReceivedCallback(authErr, response.accountState);
            }
            else {
                this.logger.verbose("One callback was provided to handleRedirectCallback, calling authResponseCallback with error");
                this.authResponseCallback(authErr, response);
            }
        }
        else if (interactionType === Constants_1.Constants.interactionTypePopup) {
            this.logger.verbose("Interaction type is popup, rejecting");
            reject(authErr);
        }
        else {
            throw ClientAuthError_1.ClientAuthError.createInvalidInteractionTypeError();
        }
    };
    // #endregion
    /**
     * Use when initiating the login process by redirecting the user's browser to the authorization endpoint.
     * @param {@link (AuthenticationParameters:type)}
     */
    UserAgentApplication.prototype.loginRedirect = function (userRequest) {
        this.logger.verbose("LoginRedirect has been called");
        // validate request
        var request = RequestUtils_1.RequestUtils.validateRequest(userRequest, true, this.clientId, Constants_1.Constants.interactionTypeRedirect);
        this.acquireTokenInteractive(Constants_1.Constants.interactionTypeRedirect, true, request, null, null);
    };
    /**
     * Use when you want to obtain an access_token for your API by redirecting the user's browser window to the authorization endpoint.
     * @param {@link (AuthenticationParameters:type)}
     *
     * To renew idToken, please pass clientId as the only scope in the Authentication Parameters
     */
    UserAgentApplication.prototype.acquireTokenRedirect = function (userRequest) {
        this.logger.verbose("AcquireTokenRedirect has been called");
        // validate request
        var request = RequestUtils_1.RequestUtils.validateRequest(userRequest, false, this.clientId, Constants_1.Constants.interactionTypeRedirect);
        this.acquireTokenInteractive(Constants_1.Constants.interactionTypeRedirect, false, request, null, null);
    };
    /**
     * Use when initiating the login process via opening a popup window in the user's browser
     *
     * @param {@link (AuthenticationParameters:type)}
     *
     * @returns {Promise.<AuthResponse>} - a promise that is fulfilled when this function has completed, or rejected if an error was raised. Returns the {@link AuthResponse} object
     */
    UserAgentApplication.prototype.loginPopup = function (userRequest) {
        var _this = this;
        this.logger.verbose("LoginPopup has been called");
        // validate request
        var request = RequestUtils_1.RequestUtils.validateRequest(userRequest, true, this.clientId, Constants_1.Constants.interactionTypePopup);
        var apiEvent = this.telemetryManager.createAndStartApiEvent(request.correlationId, ApiEvent_1.API_EVENT_IDENTIFIER.LoginPopup);
        return new Promise(function (resolve, reject) {
            _this.acquireTokenInteractive(Constants_1.Constants.interactionTypePopup, true, request, resolve, reject);
        })
            .then(function (resp) {
            _this.logger.verbose("Successfully logged in");
            _this.telemetryManager.stopAndFlushApiEvent(request.correlationId, apiEvent, true);
            return resp;
        })
            .catch(function (error) {
            _this.cacheStorage.resetTempCacheItems(request.state);
            _this.telemetryManager.stopAndFlushApiEvent(request.correlationId, apiEvent, false, error.errorCode);
            throw error;
        });
    };
    /**
     * Use when you want to obtain an access_token for your API via opening a popup window in the user's browser
     * @param {@link AuthenticationParameters}
     *
     * To renew idToken, please pass clientId as the only scope in the Authentication Parameters
     * @returns {Promise.<AuthResponse>} - a promise that is fulfilled when this function has completed, or rejected if an error was raised. Returns the {@link AuthResponse} object
     */
    UserAgentApplication.prototype.acquireTokenPopup = function (userRequest) {
        var _this = this;
        this.logger.verbose("AcquireTokenPopup has been called");
        // validate request
        var request = RequestUtils_1.RequestUtils.validateRequest(userRequest, false, this.clientId, Constants_1.Constants.interactionTypePopup);
        var apiEvent = this.telemetryManager.createAndStartApiEvent(request.correlationId, ApiEvent_1.API_EVENT_IDENTIFIER.AcquireTokenPopup);
        return new Promise(function (resolve, reject) {
            _this.acquireTokenInteractive(Constants_1.Constants.interactionTypePopup, false, request, resolve, reject);
        })
            .then(function (resp) {
            _this.logger.verbose("Successfully acquired token");
            _this.telemetryManager.stopAndFlushApiEvent(request.correlationId, apiEvent, true);
            return resp;
        })
            .catch(function (error) {
            _this.cacheStorage.resetTempCacheItems(request.state);
            _this.telemetryManager.stopAndFlushApiEvent(request.correlationId, apiEvent, false, error.errorCode);
            throw error;
        });
    };
    // #region Acquire Token
    /**
     * Use when initiating the login process or when you want to obtain an access_token for your API,
     * either by redirecting the user's browser window to the authorization endpoint or via opening a popup window in the user's browser.
     * @param {@link (AuthenticationParameters:type)}
     *
     * To renew idToken, please pass clientId as the only scope in the Authentication Parameters
     */
    UserAgentApplication.prototype.acquireTokenInteractive = function (interactionType, isLoginCall, request, resolve, reject) {
        var _this = this;
        this.logger.verbose("AcquireTokenInteractive has been called");
        // block the request if made from the hidden iframe
        WindowUtils_1.WindowUtils.blockReloadInHiddenIframes();
        var interactionProgress = this.cacheStorage.getItem(Constants_1.TemporaryCacheKeys.INTERACTION_STATUS);
        if (interactionType === Constants_1.Constants.interactionTypeRedirect) {
            this.cacheStorage.setItem(Constants_1.TemporaryCacheKeys.REDIRECT_REQUEST, "" + Constants_1.Constants.inProgress + Constants_1.Constants.resourceDelimiter + request.state);
        }
        // If already in progress, do not proceed
        if (interactionProgress === Constants_1.Constants.inProgress) {
            var thrownError = isLoginCall ? ClientAuthError_1.ClientAuthError.createLoginInProgressError() : ClientAuthError_1.ClientAuthError.createAcquireTokenInProgressError();
            var stateOnlyResponse = AuthResponse_1.buildResponseStateOnly(this.getAccountState(request.state));
            this.cacheStorage.resetTempCacheItems(request.state);
            this.authErrorHandler(interactionType, thrownError, stateOnlyResponse, reject);
            return;
        }
        // Get the account object if a session exists
        var account;
        if (request && request.account && !isLoginCall) {
            account = request.account;
            this.logger.verbose("Account set from request");
        }
        else {
            account = this.getAccount();
            this.logger.verbose("Account set from MSAL Cache");
        }
        // If no session exists, prompt the user to login.
        if (!account && !ServerRequestParameters_1.ServerRequestParameters.isSSOParam(request)) {
            if (isLoginCall) {
                // extract ADAL id_token if exists
                var adalIdToken = this.extractADALIdToken();
                // silent login if ADAL id_token is retrieved successfully - SSO
                if (adalIdToken && !request.scopes) {
                    this.logger.info("ADAL's idToken exists. Extracting login information from ADAL's idToken");
                    var tokenRequest = this.buildIDTokenRequest(request);
                    this.silentLogin = true;
                    this.acquireTokenSilent(tokenRequest).then(function (response) {
                        _this.silentLogin = false;
                        _this.logger.info("Unified cache call is successful");
                        _this.authResponseHandler(interactionType, response, resolve);
                        return;
                    }, function (error) {
                        _this.silentLogin = false;
                        _this.logger.error("Error occurred during unified cache ATS: " + error);
                        // proceed to login since ATS failed
                        _this.acquireTokenHelper(null, interactionType, isLoginCall, request, resolve, reject);
                    });
                }
                // No ADAL token found, proceed to login
                else {
                    this.logger.verbose("Login call but no token found, proceed to login");
                    this.acquireTokenHelper(null, interactionType, isLoginCall, request, resolve, reject);
                }
            }
            // AcquireToken call, but no account or context given, so throw error
            else {
                this.logger.verbose("AcquireToken call, no context or account given");
                this.logger.info("User login is required");
                var stateOnlyResponse = AuthResponse_1.buildResponseStateOnly(this.getAccountState(request.state));
                this.cacheStorage.resetTempCacheItems(request.state);
                this.authErrorHandler(interactionType, ClientAuthError_1.ClientAuthError.createUserLoginRequiredError(), stateOnlyResponse, reject);
                return;
            }
        }
        // User session exists
        else {
            this.logger.verbose("User session exists, login not required");
            this.acquireTokenHelper(account, interactionType, isLoginCall, request, resolve, reject);
        }
    };
    /**
     * @hidden
     * @ignore
     * Helper function to acquireToken
     *
     */
    UserAgentApplication.prototype.acquireTokenHelper = function (account, interactionType, isLoginCall, request, resolve, reject) {
        return tslib_1.__awaiter(this, void 0, Promise, function () {
            var scope, serverAuthenticationRequest, acquireTokenAuthority, popUpWindow, responseType, loginStartPage, urlNavigate, hash, error_1, navigate, err_1;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.logger.verbose("AcquireTokenHelper has been called");
                        this.logger.verbose("Interaction type: " + interactionType + ". isLoginCall: " + isLoginCall);
                        // Track the acquireToken progress
                        this.cacheStorage.setItem(Constants_1.TemporaryCacheKeys.INTERACTION_STATUS, Constants_1.Constants.inProgress);
                        scope = request.scopes ? request.scopes.join(" ").toLowerCase() : this.clientId.toLowerCase();
                        this.logger.verbosePii("Serialized scopes: " + scope);
                        acquireTokenAuthority = (request && request.authority) ? AuthorityFactory_1.AuthorityFactory.CreateInstance(request.authority, this.config.auth.validateAuthority, request.authorityMetadata) : this.authorityInstance;
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 11, , 12]);
                        if (!!acquireTokenAuthority.hasCachedMetadata()) return [3 /*break*/, 3];
                        this.logger.verbose("No cached metadata for authority");
                        return [4 /*yield*/, AuthorityFactory_1.AuthorityFactory.saveMetadataFromNetwork(acquireTokenAuthority, this.telemetryManager, request.correlationId)];
                    case 2:
                        _a.sent();
                        return [3 /*break*/, 4];
                    case 3:
                        this.logger.verbose("Cached metadata found for authority");
                        _a.label = 4;
                    case 4:
                        responseType = isLoginCall ? ResponseTypes.id_token : this.getTokenType(account, request.scopes, false);
                        loginStartPage = request.redirectStartPage || window.location.href;
                        serverAuthenticationRequest = new ServerRequestParameters_1.ServerRequestParameters(acquireTokenAuthority, this.clientId, responseType, this.getRedirectUri(request && request.redirectUri), request.scopes, request.state, request.correlationId);
                        this.logger.verbose("Finished building server authentication request");
                        this.updateCacheEntries(serverAuthenticationRequest, account, isLoginCall, loginStartPage);
                        this.logger.verbose("Updating cache entries");
                        // populate QueryParameters (sid/login_hint) and any other extraQueryParameters set by the developer
                        serverAuthenticationRequest.populateQueryParams(account, request);
                        this.logger.verbose("Query parameters populated from account");
                        urlNavigate = UrlUtils_1.UrlUtils.createNavigateUrl(serverAuthenticationRequest) + Constants_1.Constants.response_mode_fragment;
                        // set state in cache
                        if (interactionType === Constants_1.Constants.interactionTypeRedirect) {
                            if (!isLoginCall) {
                                this.cacheStorage.setItem(AuthCache_1.AuthCache.generateTemporaryCacheKey(Constants_1.TemporaryCacheKeys.STATE_ACQ_TOKEN, request.state), serverAuthenticationRequest.state, this.inCookie);
                                this.logger.verbose("State cached for redirect");
                                this.logger.verbosePii("State cached: " + serverAuthenticationRequest.state);
                            }
                            else {
                                this.logger.verbose("Interaction type redirect but login call is true. State not cached");
                            }
                        }
                        else if (interactionType === Constants_1.Constants.interactionTypePopup) {
                            window.renewStates.push(serverAuthenticationRequest.state);
                            window.requestType = isLoginCall ? Constants_1.Constants.login : Constants_1.Constants.renewToken;
                            this.logger.verbose("State saved to window");
                            this.logger.verbosePii("State saved: " + serverAuthenticationRequest.state);
                            // Register callback to capture results from server
                            this.registerCallback(serverAuthenticationRequest.state, scope, resolve, reject);
                        }
                        else {
                            this.logger.verbose("Invalid interaction error. State not cached");
                            throw ClientAuthError_1.ClientAuthError.createInvalidInteractionTypeError();
                        }
                        if (!(interactionType === Constants_1.Constants.interactionTypePopup)) return [3 /*break*/, 9];
                        this.logger.verbose("Interaction type is popup. Generating popup window");
                        // Generate a popup window
                        try {
                            popUpWindow = this.openPopup(urlNavigate, "msal", Constants_1.Constants.popUpWidth, Constants_1.Constants.popUpHeight);
                            // Push popup window handle onto stack for tracking
                            WindowUtils_1.WindowUtils.trackPopup(popUpWindow);
                        }
                        catch (e) {
                            this.logger.info(ClientAuthError_1.ClientAuthErrorMessage.popUpWindowError.code + ":" + ClientAuthError_1.ClientAuthErrorMessage.popUpWindowError.desc);
                            this.cacheStorage.setItem(Constants_1.ErrorCacheKeys.ERROR, ClientAuthError_1.ClientAuthErrorMessage.popUpWindowError.code);
                            this.cacheStorage.setItem(Constants_1.ErrorCacheKeys.ERROR_DESC, ClientAuthError_1.ClientAuthErrorMessage.popUpWindowError.desc);
                            if (reject) {
                                reject(ClientAuthError_1.ClientAuthError.createPopupWindowError());
                                return [2 /*return*/];
                            }
                        }
                        if (!popUpWindow) return [3 /*break*/, 8];
                        _a.label = 5;
                    case 5:
                        _a.trys.push([5, 7, , 8]);
                        return [4 /*yield*/, WindowUtils_1.WindowUtils.monitorPopupForHash(popUpWindow, this.config.system.loadFrameTimeout, urlNavigate, this.logger)];
                    case 6:
                        hash = _a.sent();
                        this.handleAuthenticationResponse(hash);
                        // Request completed successfully, set to completed
                        this.cacheStorage.removeItem(Constants_1.TemporaryCacheKeys.INTERACTION_STATUS);
                        this.logger.info("Closing popup window");
                        // TODO: Check how this can be extracted for any framework specific code?
                        if (this.config.framework.isAngular) {
                            this.broadcast("msal:popUpHashChanged", hash);
                            WindowUtils_1.WindowUtils.closePopups();
                        }
                        return [3 /*break*/, 8];
                    case 7:
                        error_1 = _a.sent();
                        if (reject) {
                            reject(error_1);
                        }
                        if (this.config.framework.isAngular) {
                            this.broadcast("msal:popUpClosed", error_1.errorCode + Constants_1.Constants.resourceDelimiter + error_1.errorMessage);
                        }
                        else {
                            // Request failed, set to canceled
                            this.cacheStorage.removeItem(Constants_1.TemporaryCacheKeys.INTERACTION_STATUS);
                            popUpWindow.close();
                        }
                        return [3 /*break*/, 8];
                    case 8: return [3 /*break*/, 10];
                    case 9:
                        // If onRedirectNavigate is implemented, invoke it and provide urlNavigate
                        if (request.onRedirectNavigate) {
                            this.logger.verbose("Invoking onRedirectNavigate callback");
                            navigate = request.onRedirectNavigate(urlNavigate);
                            // Returning false from onRedirectNavigate will stop navigation
                            if (navigate !== false) {
                                this.logger.verbose("onRedirectNavigate did not return false, navigating");
                                this.navigateWindow(urlNavigate);
                            }
                            else {
                                this.logger.verbose("onRedirectNavigate returned false, stopping navigation");
                            }
                        }
                        else {
                            // Otherwise, perform navigation
                            this.logger.verbose("Navigating window to urlNavigate");
                            this.navigateWindow(urlNavigate);
                        }
                        _a.label = 10;
                    case 10: return [3 /*break*/, 12];
                    case 11:
                        err_1 = _a.sent();
                        this.logger.error(err_1);
                        this.cacheStorage.resetTempCacheItems(request.state);
                        this.authErrorHandler(interactionType, ClientAuthError_1.ClientAuthError.createEndpointResolutionError(err_1.toString), AuthResponse_1.buildResponseStateOnly(request.state), reject);
                        if (popUpWindow) {
                            popUpWindow.close();
                        }
                        return [3 /*break*/, 12];
                    case 12: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * API interfacing idToken request when applications already have a session/hint acquired by authorization client applications
     * @param request
     */
    UserAgentApplication.prototype.ssoSilent = function (request) {
        this.logger.verbose("ssoSilent has been called");
        // throw an error on an empty request
        if (!request) {
            throw ClientConfigurationError_1.ClientConfigurationError.createEmptyRequestError();
        }
        // throw an error on no hints passed
        if (!request.sid && !request.loginHint) {
            throw ClientConfigurationError_1.ClientConfigurationError.createSsoSilentError();
        }
        return this.acquireTokenSilent(tslib_1.__assign({}, request, { scopes: [this.clientId] }));
    };
    /**
     * Use this function to obtain a token before every call to the API / resource provider
     *
     * MSAL return's a cached token when available
     * Or it send's a request to the STS to obtain a new token using a hidden iframe.
     *
     * @param {@link AuthenticationParameters}
     *
     * To renew idToken, please pass clientId as the only scope in the Authentication Parameters
     * @returns {Promise.<AuthResponse>} - a promise that is fulfilled when this function has completed, or rejected if an error was raised. Returns the {@link AuthResponse} object
     *
     */
    UserAgentApplication.prototype.acquireTokenSilent = function (userRequest) {
        var _this = this;
        this.logger.verbose("AcquireTokenSilent has been called");
        // validate the request
        var request = RequestUtils_1.RequestUtils.validateRequest(userRequest, false, this.clientId, Constants_1.Constants.interactionTypeSilent);
        var apiEvent = this.telemetryManager.createAndStartApiEvent(request.correlationId, ApiEvent_1.API_EVENT_IDENTIFIER.AcquireTokenSilent);
        var requestSignature = RequestUtils_1.RequestUtils.createRequestSignature(request);
        return new Promise(function (resolve, reject) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var scope, account, adalIdToken, responseType, serverAuthenticationRequest, adalIdTokenObject, userContainedClaims, authErr, cacheResultResponse, logMessage, err_2;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        // block the request if made from the hidden iframe
                        WindowUtils_1.WindowUtils.blockReloadInHiddenIframes();
                        scope = request.scopes.join(" ").toLowerCase();
                        this.logger.verbosePii("Serialized scopes: " + scope);
                        if (request.account) {
                            account = request.account;
                            this.logger.verbose("Account set from request");
                        }
                        else {
                            account = this.getAccount();
                            this.logger.verbose("Account set from MSAL Cache");
                        }
                        adalIdToken = this.cacheStorage.getItem(Constants_1.Constants.adalIdToken);
                        // In the event of no account being passed in the config, no session id, and no pre-existing adalIdToken, user will need to log in
                        if (!account && !(request.sid || request.loginHint) && StringUtils_1.StringUtils.isEmpty(adalIdToken)) {
                            this.logger.info("User login is required");
                            // The promise rejects with a UserLoginRequiredError, which should be caught and user should be prompted to log in interactively
                            return [2 /*return*/, reject(ClientAuthError_1.ClientAuthError.createUserLoginRequiredError())];
                        }
                        responseType = this.getTokenType(account, request.scopes, true);
                        this.logger.verbose("Response type: " + responseType);
                        serverAuthenticationRequest = new ServerRequestParameters_1.ServerRequestParameters(AuthorityFactory_1.AuthorityFactory.CreateInstance(request.authority, this.config.auth.validateAuthority, request.authorityMetadata), this.clientId, responseType, this.getRedirectUri(request.redirectUri), request.scopes, request.state, request.correlationId);
                        this.logger.verbose("Finished building server authentication request");
                        // populate QueryParameters (sid/login_hint) and any other extraQueryParameters set by the developer
                        if (ServerRequestParameters_1.ServerRequestParameters.isSSOParam(request) || account) {
                            serverAuthenticationRequest.populateQueryParams(account, request, null, true);
                            this.logger.verbose("Query parameters populated from existing SSO or account");
                        }
                        // if user didn't pass login_hint/sid and adal's idtoken is present, extract the login_hint from the adalIdToken
                        else if (!account && !StringUtils_1.StringUtils.isEmpty(adalIdToken)) {
                            adalIdTokenObject = TokenUtils_1.TokenUtils.extractIdToken(adalIdToken);
                            this.logger.verbose("ADAL's idToken exists. Extracting login information from ADAL's idToken to populate query parameters");
                            serverAuthenticationRequest.populateQueryParams(account, null, adalIdTokenObject, true);
                        }
                        else {
                            this.logger.verbose("No additional query parameters added");
                        }
                        userContainedClaims = request.claimsRequest || serverAuthenticationRequest.claimsValue;
                        // If request.forceRefresh is set to true, force a request for a new token instead of getting it from the cache
                        if (!userContainedClaims && !request.forceRefresh) {
                            try {
                                cacheResultResponse = this.getCachedToken(serverAuthenticationRequest, account);
                            }
                            catch (e) {
                                authErr = e;
                            }
                        }
                        if (!cacheResultResponse) return [3 /*break*/, 1];
                        this.logger.verbose("Token found in cache lookup");
                        this.logger.verbosePii("Scopes found: " + JSON.stringify(cacheResultResponse.scopes));
                        resolve(cacheResultResponse);
                        return [2 /*return*/, null];
                    case 1:
                        if (!authErr) return [3 /*break*/, 2];
                        this.logger.infoPii(authErr.errorCode + ":" + authErr.errorMessage);
                        reject(authErr);
                        return [2 /*return*/, null];
                    case 2:
                        logMessage = void 0;
                        if (userContainedClaims) {
                            logMessage = "Skipped cache lookup since claims were given";
                        }
                        else if (request.forceRefresh) {
                            logMessage = "Skipped cache lookup since request.forceRefresh option was set to true";
                        }
                        else {
                            logMessage = "No token found in cache lookup";
                        }
                        this.logger.verbose(logMessage);
                        // Cache result can return null if cache is empty. In that case, set authority to default value if no authority is passed to the API.
                        if (!serverAuthenticationRequest.authorityInstance) {
                            serverAuthenticationRequest.authorityInstance = request.authority ? AuthorityFactory_1.AuthorityFactory.CreateInstance(request.authority, this.config.auth.validateAuthority, request.authorityMetadata) : this.authorityInstance;
                        }
                        this.logger.verbosePii("Authority instance: " + serverAuthenticationRequest.authority);
                        _a.label = 3;
                    case 3:
                        _a.trys.push([3, 7, , 8]);
                        if (!!serverAuthenticationRequest.authorityInstance.hasCachedMetadata()) return [3 /*break*/, 5];
                        this.logger.verbose("No cached metadata for authority");
                        return [4 /*yield*/, AuthorityFactory_1.AuthorityFactory.saveMetadataFromNetwork(serverAuthenticationRequest.authorityInstance, this.telemetryManager, request.correlationId)];
                    case 4:
                        _a.sent();
                        this.logger.verbose("Authority has been updated with endpoint discovery response");
                        return [3 /*break*/, 6];
                    case 5:
                        this.logger.verbose("Cached metadata found for authority");
                        _a.label = 6;
                    case 6:
                        /*
                         * refresh attempt with iframe
                         * Already renewing for this scope, callback when we get the token.
                         */
                        if (window.activeRenewals[requestSignature]) {
                            this.logger.verbose("Renewing token in progress. Registering callback");
                            // Active renewals contains the state for each renewal.
                            this.registerCallback(window.activeRenewals[requestSignature], requestSignature, resolve, reject);
                        }
                        else {
                            if (request.scopes && request.scopes.indexOf(this.clientId) > -1 && request.scopes.length === 1) {
                                /*
                                 * App uses idToken to send to api endpoints
                                 * Default scope is tracked as clientId to store this token
                                 */
                                this.logger.verbose("ClientId is the only scope, renewing idToken");
                                this.silentLogin = true;
                                this.renewIdToken(requestSignature, resolve, reject, account, serverAuthenticationRequest);
                            }
                            else {
                                // renew access token
                                this.logger.verbose("Renewing access token");
                                this.renewToken(requestSignature, resolve, reject, account, serverAuthenticationRequest);
                            }
                        }
                        return [3 /*break*/, 8];
                    case 7:
                        err_2 = _a.sent();
                        this.logger.error(err_2);
                        reject(ClientAuthError_1.ClientAuthError.createEndpointResolutionError(err_2.toString()));
                        return [2 /*return*/, null];
                    case 8: return [2 /*return*/];
                }
            });
        }); })
            .then(function (res) {
            _this.logger.verbose("Successfully acquired token");
            _this.telemetryManager.stopAndFlushApiEvent(request.correlationId, apiEvent, true);
            return res;
        })
            .catch(function (error) {
            _this.cacheStorage.resetTempCacheItems(request.state);
            _this.telemetryManager.stopAndFlushApiEvent(request.correlationId, apiEvent, false, error.errorCode);
            throw error;
        });
    };
    // #endregion
    // #region Popup Window Creation
    /**
     * @hidden
     *
     * Configures popup window for login.
     *
     * @param urlNavigate
     * @param title
     * @param popUpWidth
     * @param popUpHeight
     * @ignore
     * @hidden
     */
    UserAgentApplication.prototype.openPopup = function (urlNavigate, title, popUpWidth, popUpHeight) {
        this.logger.verbose("OpenPopup has been called");
        try {
            /**
             * adding winLeft and winTop to account for dual monitor
             * using screenLeft and screenTop for IE8 and earlier
             */
            var winLeft = window.screenLeft ? window.screenLeft : window.screenX;
            var winTop = window.screenTop ? window.screenTop : window.screenY;
            /**
             * window.innerWidth displays browser window"s height and width excluding toolbars
             * using document.documentElement.clientWidth for IE8 and earlier
             */
            var width = window.innerWidth || document.documentElement.clientWidth || document.body.clientWidth;
            var height = window.innerHeight || document.documentElement.clientHeight || document.body.clientHeight;
            var left = ((width / 2) - (popUpWidth / 2)) + winLeft;
            var top = ((height / 2) - (popUpHeight / 2)) + winTop;
            // open the window
            var popupWindow = window.open(urlNavigate, title, "width=" + popUpWidth + ", height=" + popUpHeight + ", top=" + top + ", left=" + left + ", scrollbars=yes");
            if (!popupWindow) {
                throw ClientAuthError_1.ClientAuthError.createPopupWindowError();
            }
            if (popupWindow.focus) {
                popupWindow.focus();
            }
            return popupWindow;
        }
        catch (e) {
            this.cacheStorage.removeItem(Constants_1.TemporaryCacheKeys.INTERACTION_STATUS);
            throw ClientAuthError_1.ClientAuthError.createPopupWindowError(e.toString());
        }
    };
    // #endregion
    // #region Iframe Management
    /**
     * @hidden
     * Calling _loadFrame but with a timeout to signal failure in loadframeStatus. Callbacks are left.
     * registered when network errors occur and subsequent token requests for same resource are registered to the pending request.
     * @ignore
     */
    UserAgentApplication.prototype.loadIframeTimeout = function (urlNavigate, frameName, requestSignature) {
        return tslib_1.__awaiter(this, void 0, Promise, function () {
            var expectedState, iframe, _a, hash, error_2;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        expectedState = window.activeRenewals[requestSignature];
                        this.logger.verbosePii("Set loading state to pending for: " + requestSignature + ":" + expectedState);
                        this.cacheStorage.setItem(AuthCache_1.AuthCache.generateTemporaryCacheKey(Constants_1.TemporaryCacheKeys.RENEW_STATUS, expectedState), Constants_1.Constants.inProgress);
                        if (!this.config.system.navigateFrameWait) return [3 /*break*/, 2];
                        return [4 /*yield*/, WindowUtils_1.WindowUtils.loadFrame(urlNavigate, frameName, this.config.system.navigateFrameWait, this.logger)];
                    case 1:
                        _a = _b.sent();
                        return [3 /*break*/, 3];
                    case 2:
                        _a = WindowUtils_1.WindowUtils.loadFrameSync(urlNavigate, frameName, this.logger);
                        _b.label = 3;
                    case 3:
                        iframe = _a;
                        _b.label = 4;
                    case 4:
                        _b.trys.push([4, 6, , 7]);
                        return [4 /*yield*/, WindowUtils_1.WindowUtils.monitorIframeForHash(iframe.contentWindow, this.config.system.loadFrameTimeout, urlNavigate, this.logger)];
                    case 5:
                        hash = _b.sent();
                        if (hash) {
                            this.handleAuthenticationResponse(hash);
                        }
                        return [3 /*break*/, 7];
                    case 6:
                        error_2 = _b.sent();
                        if (this.cacheStorage.getItem(AuthCache_1.AuthCache.generateTemporaryCacheKey(Constants_1.TemporaryCacheKeys.RENEW_STATUS, expectedState)) === Constants_1.Constants.inProgress) {
                            // fail the iframe session if it's in pending state
                            this.logger.verbose("Loading frame has timed out after: " + (this.config.system.loadFrameTimeout / 1000) + " seconds for scope/authority " + requestSignature + ":" + expectedState);
                            // Error after timeout
                            if (expectedState && window.callbackMappedToRenewStates[expectedState]) {
                                window.callbackMappedToRenewStates[expectedState](null, error_2);
                            }
                            this.cacheStorage.removeItem(AuthCache_1.AuthCache.generateTemporaryCacheKey(Constants_1.TemporaryCacheKeys.RENEW_STATUS, expectedState));
                        }
                        WindowUtils_1.WindowUtils.removeHiddenIframe(iframe);
                        throw error_2;
                    case 7:
                        WindowUtils_1.WindowUtils.removeHiddenIframe(iframe);
                        return [2 /*return*/];
                }
            });
        });
    };
    // #endregion
    // #region General Helpers
    /**
     * @hidden
     * Used to redirect the browser to the STS authorization endpoint
     * @param {string} urlNavigate - URL of the authorization endpoint
     */
    UserAgentApplication.prototype.navigateWindow = function (urlNavigate, popupWindow) {
        // Navigate if valid URL
        if (urlNavigate && !StringUtils_1.StringUtils.isEmpty(urlNavigate)) {
            var navigateWindow = popupWindow ? popupWindow : window;
            var logMessage = popupWindow ? "Navigated Popup window to:" + urlNavigate : "Navigate to:" + urlNavigate;
            this.logger.infoPii(logMessage);
            navigateWindow.location.assign(urlNavigate);
        }
        else {
            this.logger.info("Navigate url is empty");
            throw AuthError_1.AuthError.createUnexpectedError("Navigate url is empty");
        }
    };
    /**
     * @hidden
     * Used to add the developer requested callback to the array of callbacks for the specified scopes. The updated array is stored on the window object
     * @param {string} expectedState - Unique state identifier (guid).
     * @param {string} scope - Developer requested permissions. Not all scopes are guaranteed to be included in the access token returned.
     * @param {Function} resolve - The resolve function of the promise object.
     * @param {Function} reject - The reject function of the promise object.
     * @ignore
     */
    UserAgentApplication.prototype.registerCallback = function (expectedState, requestSignature, resolve, reject) {
        var _this = this;
        // track active renewals
        window.activeRenewals[requestSignature] = expectedState;
        // initialize callbacks mapped array
        if (!window.promiseMappedToRenewStates[expectedState]) {
            window.promiseMappedToRenewStates[expectedState] = [];
        }
        // indexing on the current state, push the callback params to callbacks mapped
        window.promiseMappedToRenewStates[expectedState].push({ resolve: resolve, reject: reject });
        // Store the server response in the current window??
        if (!window.callbackMappedToRenewStates[expectedState]) {
            window.callbackMappedToRenewStates[expectedState] = function (response, error) {
                // reset active renewals
                window.activeRenewals[requestSignature] = null;
                // for all promiseMappedtoRenewStates for a given 'state' - call the reject/resolve with error/token respectively
                for (var i = 0; i < window.promiseMappedToRenewStates[expectedState].length; ++i) {
                    try {
                        if (error) {
                            window.promiseMappedToRenewStates[expectedState][i].reject(error);
                        }
                        else if (response) {
                            window.promiseMappedToRenewStates[expectedState][i].resolve(response);
                        }
                        else {
                            _this.cacheStorage.resetTempCacheItems(expectedState);
                            throw AuthError_1.AuthError.createUnexpectedError("Error and response are both null");
                        }
                    }
                    catch (e) {
                        _this.logger.warning(e);
                    }
                }
                // reset
                window.promiseMappedToRenewStates[expectedState] = null;
                window.callbackMappedToRenewStates[expectedState] = null;
            };
        }
    };
    // #endregion
    // #region Logout
    /**
     * Use to log out the current user, and redirect the user to the postLogoutRedirectUri.
     * Default behaviour is to redirect the user to `window.location.href`.
     */
    UserAgentApplication.prototype.logout = function (correlationId) {
        this.logger.verbose("Logout has been called");
        this.logoutAsync(correlationId);
    };
    /**
     * Async version of logout(). Use to log out the current user.
     * @param correlationId Request correlationId
     */
    UserAgentApplication.prototype.logoutAsync = function (correlationId) {
        return tslib_1.__awaiter(this, void 0, Promise, function () {
            var requestCorrelationId, apiEvent, correlationIdParam, postLogoutQueryParam, urlNavigate, error_3;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        requestCorrelationId = correlationId || CryptoUtils_1.CryptoUtils.createNewGuid();
                        apiEvent = this.telemetryManager.createAndStartApiEvent(requestCorrelationId, ApiEvent_1.API_EVENT_IDENTIFIER.Logout);
                        this.clearCache();
                        this.account = null;
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 5, , 6]);
                        if (!!this.authorityInstance.hasCachedMetadata()) return [3 /*break*/, 3];
                        this.logger.verbose("No cached metadata for authority");
                        return [4 /*yield*/, AuthorityFactory_1.AuthorityFactory.saveMetadataFromNetwork(this.authorityInstance, this.telemetryManager, correlationId)];
                    case 2:
                        _a.sent();
                        return [3 /*break*/, 4];
                    case 3:
                        this.logger.verbose("Cached metadata found for authority");
                        _a.label = 4;
                    case 4:
                        correlationIdParam = "client-request-id=" + requestCorrelationId;
                        postLogoutQueryParam = void 0;
                        if (this.getPostLogoutRedirectUri()) {
                            postLogoutQueryParam = "&post_logout_redirect_uri=" + encodeURIComponent(this.getPostLogoutRedirectUri());
                            this.logger.verbose("redirectUri found and set");
                        }
                        else {
                            postLogoutQueryParam = "";
                            this.logger.verbose("No redirectUri set for app. postLogoutQueryParam is empty");
                        }
                        urlNavigate = void 0;
                        if (this.authorityInstance.EndSessionEndpoint) {
                            urlNavigate = this.authorityInstance.EndSessionEndpoint + "?" + correlationIdParam + postLogoutQueryParam;
                            this.logger.verbose("EndSessionEndpoint found and urlNavigate set");
                            this.logger.verbosePii("urlNavigate set to: " + this.authorityInstance.EndSessionEndpoint);
                        }
                        else {
                            urlNavigate = this.authority + "oauth2/v2.0/logout?" + correlationIdParam + postLogoutQueryParam;
                            this.logger.verbose("No endpoint, urlNavigate set to default");
                        }
                        this.telemetryManager.stopAndFlushApiEvent(requestCorrelationId, apiEvent, true);
                        this.logger.verbose("Navigating window to urlNavigate");
                        this.navigateWindow(urlNavigate);
                        return [3 /*break*/, 6];
                    case 5:
                        error_3 = _a.sent();
                        this.telemetryManager.stopAndFlushApiEvent(requestCorrelationId, apiEvent, false, error_3.errorCode);
                        return [3 /*break*/, 6];
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * @hidden
     * Clear all access tokens in the cache.
     * @ignore
     */
    UserAgentApplication.prototype.clearCache = function () {
        this.logger.verbose("Clearing cache");
        window.renewStates = [];
        var accessTokenItems = this.cacheStorage.getAllAccessTokens(Constants_1.Constants.clientId, Constants_1.Constants.homeAccountIdentifier);
        for (var i = 0; i < accessTokenItems.length; i++) {
            this.cacheStorage.removeItem(JSON.stringify(accessTokenItems[i].key));
        }
        this.cacheStorage.resetCacheItems();
        this.cacheStorage.clearMsalCookie();
        this.logger.verbose("Cache cleared");
    };
    /**
     * @hidden
     * Clear a given access token from the cache.
     *
     * @param accessToken
     */
    UserAgentApplication.prototype.clearCacheForScope = function (accessToken) {
        this.logger.verbose("Clearing access token from cache");
        var accessTokenItems = this.cacheStorage.getAllAccessTokens(Constants_1.Constants.clientId, Constants_1.Constants.homeAccountIdentifier);
        for (var i = 0; i < accessTokenItems.length; i++) {
            var token = accessTokenItems[i];
            if (token.value.accessToken === accessToken) {
                this.cacheStorage.removeItem(JSON.stringify(token.key));
                this.logger.verbosePii("Access token removed: " + token.key);
            }
        }
    };
    // #endregion
    // #region Response
    /**
     * @hidden
     * @ignore
     * Checks if the redirect response is received from the STS. In case of redirect, the url fragment has either id_token, access_token or error.
     * @param {string} hash - Hash passed from redirect page.
     * @returns {Boolean} - true if response contains id_token, access_token or error, false otherwise.
     */
    UserAgentApplication.prototype.isCallback = function (hash) {
        this.logger.info("isCallback will be deprecated in favor of urlContainsHash in MSAL.js v2.0.");
        this.logger.verbose("isCallback has been called");
        return UrlUtils_1.UrlUtils.urlContainsHash(hash);
    };
    /**
     * @hidden
     * Used to call the constructor callback with the token/error
     * @param {string} [hash=window.location.hash] - Hash fragment of Url.
     */
    UserAgentApplication.prototype.processCallBack = function (hash, stateInfo, parentCallback) {
        this.logger.info("ProcessCallBack has been called. Processing callback from redirect response");
        // get the state info from the hash
        if (!stateInfo) {
            this.logger.verbose("StateInfo is null, getting stateInfo from hash");
            stateInfo = this.getResponseState(hash);
        }
        var response;
        var authErr;
        // Save the token info from the hash
        try {
            response = this.saveTokenFromHash(hash, stateInfo);
        }
        catch (err) {
            authErr = err;
        }
        try {
            // Clear the cookie in the hash
            this.cacheStorage.clearMsalCookie(stateInfo.state);
            var accountState = this.getAccountState(stateInfo.state);
            if (response) {
                if ((stateInfo.requestType === Constants_1.Constants.renewToken) || response.accessToken) {
                    if (window.parent !== window) {
                        this.logger.verbose("Window is in iframe, acquiring token silently");
                    }
                    else {
                        this.logger.verbose("Acquiring token interactive in progress");
                    }
                    this.logger.verbose("Response tokenType set to " + Constants_1.ServerHashParamKeys.ACCESS_TOKEN);
                    response.tokenType = Constants_1.ServerHashParamKeys.ACCESS_TOKEN;
                }
                else if (stateInfo.requestType === Constants_1.Constants.login) {
                    this.logger.verbose("Response tokenType set to " + Constants_1.ServerHashParamKeys.ID_TOKEN);
                    response.tokenType = Constants_1.ServerHashParamKeys.ID_TOKEN;
                }
                if (!parentCallback) {
                    this.logger.verbose("Setting redirectResponse");
                    this.redirectResponse = response;
                    return;
                }
            }
            else if (!parentCallback) {
                this.logger.verbose("Response is null, setting redirectResponse with state");
                this.redirectResponse = AuthResponse_1.buildResponseStateOnly(accountState);
                this.redirectError = authErr;
                this.cacheStorage.resetTempCacheItems(stateInfo.state);
                return;
            }
            this.logger.verbose("Calling callback provided to processCallback");
            parentCallback(response, authErr);
        }
        catch (err) {
            this.logger.error("Error occurred in token received callback function: " + err);
            throw ClientAuthError_1.ClientAuthError.createErrorInCallbackFunction(err.toString());
        }
    };
    /**
     * @hidden
     * This method must be called for processing the response received from the STS if using popups or iframes. It extracts the hash, processes the token or error
     * information and saves it in the cache. It then resolves the promises with the result.
     * @param {string} [hash=window.location.hash] - Hash fragment of Url.
     */
    UserAgentApplication.prototype.handleAuthenticationResponse = function (hash) {
        this.logger.verbose("HandleAuthenticationResponse has been called");
        // retrieve the hash
        var locationHash = hash || window.location.hash;
        // if (window.parent !== window), by using self, window.parent becomes equal to window in getResponseState method specifically
        var stateInfo = this.getResponseState(locationHash);
        this.logger.verbose("Obtained state from response");
        var tokenResponseCallback = window.callbackMappedToRenewStates[stateInfo.state];
        this.processCallBack(locationHash, stateInfo, tokenResponseCallback);
        // If current window is opener, close all windows
        WindowUtils_1.WindowUtils.closePopups();
    };
    /**
     * @hidden
     * This method must be called for processing the response received from the STS when using redirect flows. It extracts the hash, processes the token or error
     * information and saves it in the cache. The result can then be accessed by user registered callbacks.
     * @param {string} [hash=window.location.hash] - Hash fragment of Url.
     */
    UserAgentApplication.prototype.handleRedirectAuthenticationResponse = function (hash) {
        this.logger.info("Returned from redirect url");
        this.logger.verbose("HandleRedirectAuthenticationResponse has been called");
        // clear hash from window
        window.location.hash = "";
        this.logger.verbose("Window.location.hash cleared");
        // if (window.parent !== window), by using self, window.parent becomes equal to window in getResponseState method specifically
        var stateInfo = this.getResponseState(hash);
        // if set to navigate to loginRequest page post login
        if (this.config.auth.navigateToLoginRequestUrl && window.parent === window) {
            this.logger.verbose("Window.parent is equal to window, not in popup or iframe. Navigation to login request url after login turned on");
            var loginRequestUrl = this.cacheStorage.getItem(AuthCache_1.AuthCache.generateTemporaryCacheKey(Constants_1.TemporaryCacheKeys.LOGIN_REQUEST, stateInfo.state), this.inCookie);
            // Redirect to home page if login request url is null (real null or the string null)
            if (!loginRequestUrl || loginRequestUrl === "null") {
                this.logger.error("Unable to get valid login request url from cache, redirecting to home page");
                window.location.assign("/");
                return;
            }
            else {
                this.logger.verbose("Valid login request url obtained from cache");
                var currentUrl = UrlUtils_1.UrlUtils.removeHashFromUrl(window.location.href);
                var finalRedirectUrl = UrlUtils_1.UrlUtils.removeHashFromUrl(loginRequestUrl);
                if (currentUrl !== finalRedirectUrl) {
                    this.logger.verbose("Current url is not login request url, navigating");
                    this.logger.verbosePii("CurrentUrl: " + currentUrl + ", finalRedirectUrl: " + finalRedirectUrl);
                    window.location.assign("" + finalRedirectUrl + hash);
                    return;
                }
                else {
                    this.logger.verbose("Current url matches login request url");
                    var loginRequestUrlComponents = UrlUtils_1.UrlUtils.GetUrlComponents(loginRequestUrl);
                    if (loginRequestUrlComponents.Hash) {
                        this.logger.verbose("Login request url contains hash, resetting non-msal hash");
                        window.location.hash = loginRequestUrlComponents.Hash;
                    }
                }
            }
        }
        else if (!this.config.auth.navigateToLoginRequestUrl) {
            this.logger.verbose("Default navigation to start page after login turned off");
        }
        this.processCallBack(hash, stateInfo, null);
    };
    /**
     * @hidden
     * Creates a stateInfo object from the URL fragment and returns it.
     * @param {string} hash  -  Hash passed from redirect page
     * @returns {TokenResponse} an object created from the redirect response from AAD comprising of the keys - parameters, requestType, stateMatch, stateResponse and valid.
     * @ignore
     */
    UserAgentApplication.prototype.getResponseState = function (hash) {
        this.logger.verbose("GetResponseState has been called");
        var parameters = UrlUtils_1.UrlUtils.deserializeHash(hash);
        var stateResponse;
        if (!parameters) {
            throw AuthError_1.AuthError.createUnexpectedError("Hash was not parsed correctly.");
        }
        if (parameters.hasOwnProperty(Constants_1.ServerHashParamKeys.STATE)) {
            this.logger.verbose("Hash contains state. Creating stateInfo object");
            var parsedState = RequestUtils_1.RequestUtils.parseLibraryState(parameters.state);
            stateResponse = {
                requestType: Constants_1.Constants.unknown,
                state: parameters.state,
                timestamp: parsedState.ts,
                method: parsedState.method,
                stateMatch: false
            };
        }
        else {
            throw AuthError_1.AuthError.createUnexpectedError("Hash does not contain state.");
        }
        /*
         * async calls can fire iframe and login request at the same time if developer does not use the API as expected
         * incoming callback needs to be looked up to find the request type
         */
        // loginRedirect
        if (stateResponse.state === this.cacheStorage.getItem(AuthCache_1.AuthCache.generateTemporaryCacheKey(Constants_1.TemporaryCacheKeys.STATE_LOGIN, stateResponse.state), this.inCookie) || stateResponse.state === this.silentAuthenticationState) {
            this.logger.verbose("State matches cached state, setting requestType to login");
            stateResponse.requestType = Constants_1.Constants.login;
            stateResponse.stateMatch = true;
            return stateResponse;
        }
        // acquireTokenRedirect
        else if (stateResponse.state === this.cacheStorage.getItem(AuthCache_1.AuthCache.generateTemporaryCacheKey(Constants_1.TemporaryCacheKeys.STATE_ACQ_TOKEN, stateResponse.state), this.inCookie)) {
            this.logger.verbose("State matches cached state, setting requestType to renewToken");
            stateResponse.requestType = Constants_1.Constants.renewToken;
            stateResponse.stateMatch = true;
            return stateResponse;
        }
        // external api requests may have many renewtoken requests for different resource
        if (!stateResponse.stateMatch) {
            this.logger.verbose("State does not match cached state, setting requestType to type from window");
            stateResponse.requestType = window.requestType;
            var statesInParentContext = window.renewStates;
            for (var i = 0; i < statesInParentContext.length; i++) {
                if (statesInParentContext[i] === stateResponse.state) {
                    this.logger.verbose("Matching state found for request");
                    stateResponse.stateMatch = true;
                    break;
                }
            }
            if (!stateResponse.stateMatch) {
                this.logger.verbose("Matching state not found for request");
            }
        }
        return stateResponse;
    };
    // #endregion
    // #region Token Processing (Extract to TokenProcessing.ts)
    /**
     * @hidden
     * Used to get token for the specified set of scopes from the cache
     * @param {@link ServerRequestParameters} - Request sent to the STS to obtain an id_token/access_token
     * @param {Account} account - Account for which the scopes were requested
     */
    UserAgentApplication.prototype.getCachedToken = function (serverAuthenticationRequest, account) {
        this.logger.verbose("GetCachedToken has been called");
        var accessTokenCacheItem = null;
        var scopes = serverAuthenticationRequest.scopes;
        // filter by clientId and account
        var tokenCacheItems = this.cacheStorage.getAllAccessTokens(this.clientId, account ? account.homeAccountIdentifier : null);
        this.logger.verbose("Getting all cached access tokens");
        // No match found after initial filtering
        if (tokenCacheItems.length === 0) {
            this.logger.verbose("No matching tokens found when filtered by clientId and account");
            return null;
        }
        var filteredItems = [];
        // if no authority passed
        if (!serverAuthenticationRequest.authority) {
            this.logger.verbose("No authority passed, filtering tokens by scope");
            // filter by scope
            for (var i = 0; i < tokenCacheItems.length; i++) {
                var cacheItem = tokenCacheItems[i];
                var cachedScopes = cacheItem.key.scopes.split(" ");
                if (ScopeSet_1.ScopeSet.containsScope(cachedScopes, scopes)) {
                    filteredItems.push(cacheItem);
                }
            }
            // if only one cached token found
            if (filteredItems.length === 1) {
                this.logger.verbose("One matching token found, setting authorityInstance");
                accessTokenCacheItem = filteredItems[0];
                serverAuthenticationRequest.authorityInstance = AuthorityFactory_1.AuthorityFactory.CreateInstance(accessTokenCacheItem.key.authority, this.config.auth.validateAuthority);
            }
            // if more than one cached token is found
            else if (filteredItems.length > 1) {
                throw ClientAuthError_1.ClientAuthError.createMultipleMatchingTokensInCacheError(scopes.toString());
            }
            // if no match found, check if there was a single authority used
            else {
                this.logger.verbose("No matching token found when filtering by scope");
                var authorityList = this.getUniqueAuthority(tokenCacheItems, "authority");
                if (authorityList.length > 1) {
                    throw ClientAuthError_1.ClientAuthError.createMultipleAuthoritiesInCacheError(scopes.toString());
                }
                this.logger.verbose("Single authority used, setting authorityInstance");
                serverAuthenticationRequest.authorityInstance = AuthorityFactory_1.AuthorityFactory.CreateInstance(authorityList[0], this.config.auth.validateAuthority);
            }
        }
        // if an authority is passed in the API
        else {
            this.logger.verbose("Authority passed, filtering by authority and scope");
            // filter by authority and scope
            for (var i = 0; i < tokenCacheItems.length; i++) {
                var cacheItem = tokenCacheItems[i];
                var cachedScopes = cacheItem.key.scopes.split(" ");
                if (ScopeSet_1.ScopeSet.containsScope(cachedScopes, scopes) && UrlUtils_1.UrlUtils.CanonicalizeUri(cacheItem.key.authority) === serverAuthenticationRequest.authority) {
                    filteredItems.push(cacheItem);
                }
            }
            // no match
            if (filteredItems.length === 0) {
                this.logger.verbose("No matching tokens found");
                return null;
            }
            // if only one cachedToken Found
            else if (filteredItems.length === 1) {
                this.logger.verbose("Single token found");
                accessTokenCacheItem = filteredItems[0];
            }
            else {
                // if more than one cached token is found
                throw ClientAuthError_1.ClientAuthError.createMultipleMatchingTokensInCacheError(scopes.toString());
            }
        }
        if (accessTokenCacheItem != null) {
            this.logger.verbose("Evaluating access token found");
            var expired = Number(accessTokenCacheItem.value.expiresIn);
            // If expiration is within offset, it will force renew
            var offset = this.config.system.tokenRenewalOffsetSeconds || 300;
            if (expired && (expired > TimeUtils_1.TimeUtils.now() + offset)) {
                this.logger.verbose("Token expiration is within offset, renewing token");
                var idTokenObj = new IdToken_1.IdToken(accessTokenCacheItem.value.idToken);
                if (!account) {
                    account = this.getAccount();
                    if (!account) {
                        throw AuthError_1.AuthError.createUnexpectedError("Account should not be null here.");
                    }
                }
                var aState = this.getAccountState(serverAuthenticationRequest.state);
                var response = {
                    uniqueId: "",
                    tenantId: "",
                    tokenType: (accessTokenCacheItem.value.idToken === accessTokenCacheItem.value.accessToken) ? Constants_1.ServerHashParamKeys.ID_TOKEN : Constants_1.ServerHashParamKeys.ACCESS_TOKEN,
                    idToken: idTokenObj,
                    idTokenClaims: idTokenObj.claims,
                    accessToken: accessTokenCacheItem.value.accessToken,
                    scopes: accessTokenCacheItem.key.scopes.split(" "),
                    expiresOn: new Date(expired * 1000),
                    account: account,
                    accountState: aState,
                    fromCache: true
                };
                ResponseUtils_1.ResponseUtils.setResponseIdToken(response, idTokenObj);
                this.logger.verbose("Response generated and token set");
                return response;
            }
            else {
                this.logger.verbose("Token expired, removing from cache");
                this.cacheStorage.removeItem(JSON.stringify(filteredItems[0].key));
                return null;
            }
        }
        else {
            this.logger.verbose("No tokens found");
            return null;
        }
    };
    /**
     * @hidden
     * Used to get a unique list of authorities from the cache
     * @param {Array<AccessTokenCacheItem>}  accessTokenCacheItems - accessTokenCacheItems saved in the cache
     * @ignore
     */
    UserAgentApplication.prototype.getUniqueAuthority = function (accessTokenCacheItems, property) {
        this.logger.verbose("GetUniqueAuthority has been called");
        var authorityList = [];
        var flags = [];
        accessTokenCacheItems.forEach(function (element) {
            if (element.key.hasOwnProperty(property) && (flags.indexOf(element.key[property]) === -1)) {
                flags.push(element.key[property]);
                authorityList.push(element.key[property]);
            }
        });
        return authorityList;
    };
    /**
     * @hidden
     * Check if ADAL id_token exists and return if exists.
     *
     */
    UserAgentApplication.prototype.extractADALIdToken = function () {
        this.logger.verbose("ExtractADALIdToken has been called");
        var adalIdToken = this.cacheStorage.getItem(Constants_1.Constants.adalIdToken);
        return (!StringUtils_1.StringUtils.isEmpty(adalIdToken)) ? TokenUtils_1.TokenUtils.extractIdToken(adalIdToken) : null;
    };
    /**
     * @hidden
     * Acquires access token using a hidden iframe.
     * @ignore
     */
    UserAgentApplication.prototype.renewToken = function (requestSignature, resolve, reject, account, serverAuthenticationRequest) {
        this.logger.verbose("RenewToken has been called");
        this.logger.verbosePii("RenewToken scope and authority: " + requestSignature);
        var frameName = WindowUtils_1.WindowUtils.generateFrameName(Constants_1.FramePrefix.TOKEN_FRAME, requestSignature);
        WindowUtils_1.WindowUtils.addHiddenIFrame(frameName, this.logger);
        this.updateCacheEntries(serverAuthenticationRequest, account, false);
        this.logger.verbosePii("RenewToken expected state: " + serverAuthenticationRequest.state);
        // Build urlNavigate with "prompt=none" and navigate to URL in hidden iFrame
        var urlNavigate = UrlUtils_1.UrlUtils.urlRemoveQueryStringParameter(UrlUtils_1.UrlUtils.createNavigateUrl(serverAuthenticationRequest), Constants_1.Constants.prompt) + Constants_1.Constants.prompt_none + Constants_1.Constants.response_mode_fragment;
        window.renewStates.push(serverAuthenticationRequest.state);
        window.requestType = Constants_1.Constants.renewToken;
        this.logger.verbose("Set window.renewState and requestType");
        this.registerCallback(serverAuthenticationRequest.state, requestSignature, resolve, reject);
        this.logger.infoPii("Navigate to: " + urlNavigate);
        this.loadIframeTimeout(urlNavigate, frameName, requestSignature).catch(function (error) { return reject(error); });
    };
    /**
     * @hidden
     * Renews idtoken for app's own backend when clientId is passed as a single scope in the scopes array.
     * @ignore
     */
    UserAgentApplication.prototype.renewIdToken = function (requestSignature, resolve, reject, account, serverAuthenticationRequest) {
        this.logger.info("RenewIdToken has been called");
        var frameName = WindowUtils_1.WindowUtils.generateFrameName(Constants_1.FramePrefix.ID_TOKEN_FRAME, requestSignature);
        WindowUtils_1.WindowUtils.addHiddenIFrame(frameName, this.logger);
        this.updateCacheEntries(serverAuthenticationRequest, account, false);
        this.logger.verbose("RenewIdToken expected state: " + serverAuthenticationRequest.state);
        // Build urlNavigate with "prompt=none" and navigate to URL in hidden iFrame
        var urlNavigate = UrlUtils_1.UrlUtils.urlRemoveQueryStringParameter(UrlUtils_1.UrlUtils.createNavigateUrl(serverAuthenticationRequest), Constants_1.Constants.prompt) + Constants_1.Constants.prompt_none + Constants_1.Constants.response_mode_fragment;
        if (this.silentLogin) {
            this.logger.verbose("Silent login is true, set silentAuthenticationState");
            window.requestType = Constants_1.Constants.login;
            this.silentAuthenticationState = serverAuthenticationRequest.state;
        }
        else {
            this.logger.verbose("Not silent login, set window.renewState and requestType");
            window.requestType = Constants_1.Constants.renewToken;
            window.renewStates.push(serverAuthenticationRequest.state);
        }
        // note: scope here is clientId
        this.registerCallback(serverAuthenticationRequest.state, requestSignature, resolve, reject);
        this.logger.infoPii("Navigate to:\" " + urlNavigate);
        this.loadIframeTimeout(urlNavigate, frameName, requestSignature).catch(function (error) { return reject(error); });
    };
    /**
     * @hidden
     *
     * This method must be called for processing the response received from AAD. It extracts the hash, processes the token or error, saves it in the cache and calls the registered callbacks with the result.
     * @param {string} authority authority received in the redirect response from AAD.
     * @param {TokenResponse} requestInfo an object created from the redirect response from AAD comprising of the keys - parameters, requestType, stateMatch, stateResponse and valid.
     * @param {Account} account account object for which scopes are consented for. The default account is the logged in account.
     * @param {ClientInfo} clientInfo clientInfo received as part of the response comprising of fields uid and utid.
     * @param {IdToken} idToken idToken received as part of the response.
     * @ignore
     * @private
     */
    /* tslint:disable:no-string-literal */
    UserAgentApplication.prototype.saveAccessToken = function (response, authority, parameters, clientInfo, idTokenObj) {
        this.logger.verbose("SaveAccessToken has been called");
        var scope;
        var accessTokenResponse = tslib_1.__assign({}, response);
        var clientObj = new ClientInfo_1.ClientInfo(clientInfo);
        var expiration;
        // if the response contains "scope"
        if (parameters.hasOwnProperty(Constants_1.ServerHashParamKeys.SCOPE)) {
            this.logger.verbose("Response parameters contains scope");
            // read the scopes
            scope = parameters[Constants_1.ServerHashParamKeys.SCOPE];
            var consentedScopes = scope.split(" ");
            // retrieve all access tokens from the cache, remove the dup scores
            var accessTokenCacheItems = this.cacheStorage.getAllAccessTokens(this.clientId, authority);
            this.logger.verbose("Retrieving all access tokens from cache and removing duplicates");
            for (var i = 0; i < accessTokenCacheItems.length; i++) {
                var accessTokenCacheItem = accessTokenCacheItems[i];
                if (accessTokenCacheItem.key.homeAccountIdentifier === response.account.homeAccountIdentifier) {
                    var cachedScopes = accessTokenCacheItem.key.scopes.split(" ");
                    if (ScopeSet_1.ScopeSet.isIntersectingScopes(cachedScopes, consentedScopes)) {
                        this.cacheStorage.removeItem(JSON.stringify(accessTokenCacheItem.key));
                    }
                }
            }
            // Generate and cache accessTokenKey and accessTokenValue
            var expiresIn = TimeUtils_1.TimeUtils.parseExpiresIn(parameters[Constants_1.ServerHashParamKeys.EXPIRES_IN]);
            var parsedState = RequestUtils_1.RequestUtils.parseLibraryState(parameters[Constants_1.ServerHashParamKeys.STATE]);
            expiration = parsedState.ts + expiresIn;
            var accessTokenKey = new AccessTokenKey_1.AccessTokenKey(authority, this.clientId, scope, clientObj.uid, clientObj.utid);
            var accessTokenValue = new AccessTokenValue_1.AccessTokenValue(parameters[Constants_1.ServerHashParamKeys.ACCESS_TOKEN], idTokenObj.rawIdToken, expiration.toString(), clientInfo);
            this.cacheStorage.setItem(JSON.stringify(accessTokenKey), JSON.stringify(accessTokenValue));
            this.logger.verbose("Saving token to cache");
            accessTokenResponse.accessToken = parameters[Constants_1.ServerHashParamKeys.ACCESS_TOKEN];
            accessTokenResponse.scopes = consentedScopes;
        }
        // if the response does not contain "scope" - scope is usually client_id and the token will be id_token
        else {
            this.logger.verbose("Response parameters does not contain scope, clientId set as scope");
            scope = this.clientId;
            // Generate and cache accessTokenKey and accessTokenValue
            var accessTokenKey = new AccessTokenKey_1.AccessTokenKey(authority, this.clientId, scope, clientObj.uid, clientObj.utid);
            expiration = Number(idTokenObj.expiration);
            var accessTokenValue = new AccessTokenValue_1.AccessTokenValue(parameters[Constants_1.ServerHashParamKeys.ID_TOKEN], parameters[Constants_1.ServerHashParamKeys.ID_TOKEN], expiration.toString(), clientInfo);
            this.cacheStorage.setItem(JSON.stringify(accessTokenKey), JSON.stringify(accessTokenValue));
            this.logger.verbose("Saving token to cache");
            accessTokenResponse.scopes = [scope];
            accessTokenResponse.accessToken = parameters[Constants_1.ServerHashParamKeys.ID_TOKEN];
        }
        if (expiration) {
            this.logger.verbose("New expiration set");
            accessTokenResponse.expiresOn = new Date(expiration * 1000);
        }
        else {
            this.logger.error("Could not parse expiresIn parameter");
        }
        return accessTokenResponse;
    };
    /**
     * @hidden
     * Saves token or error received in the response from AAD in the cache. In case of id_token, it also creates the account object.
     * @ignore
     */
    UserAgentApplication.prototype.saveTokenFromHash = function (hash, stateInfo) {
        this.logger.verbose("SaveTokenFromHash has been called");
        this.logger.info("State status: " + stateInfo.stateMatch + "; Request type: " + stateInfo.requestType);
        var response = {
            uniqueId: "",
            tenantId: "",
            tokenType: "",
            idToken: null,
            idTokenClaims: null,
            accessToken: null,
            scopes: [],
            expiresOn: null,
            account: null,
            accountState: "",
            fromCache: false
        };
        var error;
        var hashParams = UrlUtils_1.UrlUtils.deserializeHash(hash);
        var authorityKey = "";
        var acquireTokenAccountKey = "";
        var idTokenObj = null;
        // If server returns an error
        if (hashParams.hasOwnProperty(Constants_1.ServerHashParamKeys.ERROR_DESCRIPTION) || hashParams.hasOwnProperty(Constants_1.ServerHashParamKeys.ERROR)) {
            this.logger.verbose("Server returned an error");
            this.logger.infoPii("Error : " + hashParams[Constants_1.ServerHashParamKeys.ERROR] + "; Error description: " + hashParams[Constants_1.ServerHashParamKeys.ERROR_DESCRIPTION]);
            this.cacheStorage.setItem(Constants_1.ErrorCacheKeys.ERROR, hashParams[Constants_1.ServerHashParamKeys.ERROR]);
            this.cacheStorage.setItem(Constants_1.ErrorCacheKeys.ERROR_DESC, hashParams[Constants_1.ServerHashParamKeys.ERROR_DESCRIPTION]);
            // login
            if (stateInfo.requestType === Constants_1.Constants.login) {
                this.logger.verbose("RequestType is login, caching login error, generating authorityKey");
                this.cacheStorage.setItem(Constants_1.ErrorCacheKeys.LOGIN_ERROR, hashParams[Constants_1.ServerHashParamKeys.ERROR_DESCRIPTION] + ":" + hashParams[Constants_1.ServerHashParamKeys.ERROR]);
                authorityKey = AuthCache_1.AuthCache.generateAuthorityKey(stateInfo.state);
            }
            // acquireToken
            if (stateInfo.requestType === Constants_1.Constants.renewToken) {
                this.logger.verbose("RequestType is renewToken, generating acquireTokenAccountKey");
                authorityKey = AuthCache_1.AuthCache.generateAuthorityKey(stateInfo.state);
                var account = this.getAccount();
                var accountId = void 0;
                if (account && !StringUtils_1.StringUtils.isEmpty(account.homeAccountIdentifier)) {
                    accountId = account.homeAccountIdentifier;
                    this.logger.verbose("AccountId is set");
                }
                else {
                    accountId = Constants_1.Constants.no_account;
                    this.logger.verbose("AccountId is set as no_account");
                }
                acquireTokenAccountKey = AuthCache_1.AuthCache.generateAcquireTokenAccountKey(accountId, stateInfo.state);
            }
            var _a = Constants_1.ServerHashParamKeys.ERROR, hashErr = hashParams[_a], _b = Constants_1.ServerHashParamKeys.ERROR_DESCRIPTION, hashErrDesc = hashParams[_b];
            if (InteractionRequiredAuthError_1.InteractionRequiredAuthError.isInteractionRequiredError(hashErr) ||
                InteractionRequiredAuthError_1.InteractionRequiredAuthError.isInteractionRequiredError(hashErrDesc)) {
                error = new InteractionRequiredAuthError_1.InteractionRequiredAuthError(hashParams[Constants_1.ServerHashParamKeys.ERROR], hashParams[Constants_1.ServerHashParamKeys.ERROR_DESCRIPTION]);
            }
            else {
                error = new ServerError_1.ServerError(hashParams[Constants_1.ServerHashParamKeys.ERROR], hashParams[Constants_1.ServerHashParamKeys.ERROR_DESCRIPTION]);
            }
        }
        // If the server returns "Success"
        else {
            this.logger.verbose("Server returns success");
            // Verify the state from redirect and record tokens to storage if exists
            if (stateInfo.stateMatch) {
                this.logger.info("State is right");
                if (hashParams.hasOwnProperty(Constants_1.ServerHashParamKeys.SESSION_STATE)) {
                    this.logger.verbose("Fragment has session state, caching");
                    this.cacheStorage.setItem(AuthCache_1.AuthCache.generateTemporaryCacheKey(Constants_1.TemporaryCacheKeys.SESSION_STATE, stateInfo.state), hashParams[Constants_1.ServerHashParamKeys.SESSION_STATE]);
                }
                response.accountState = this.getAccountState(stateInfo.state);
                var clientInfo = "";
                // Process access_token
                if (hashParams.hasOwnProperty(Constants_1.ServerHashParamKeys.ACCESS_TOKEN)) {
                    this.logger.info("Fragment has access token");
                    response.accessToken = hashParams[Constants_1.ServerHashParamKeys.ACCESS_TOKEN];
                    if (hashParams.hasOwnProperty(Constants_1.ServerHashParamKeys.SCOPE)) {
                        response.scopes = hashParams[Constants_1.ServerHashParamKeys.SCOPE].split(" ");
                    }
                    // retrieve the id_token from response if present
                    if (hashParams.hasOwnProperty(Constants_1.ServerHashParamKeys.ID_TOKEN)) {
                        this.logger.verbose("Fragment has id_token");
                        idTokenObj = new IdToken_1.IdToken(hashParams[Constants_1.ServerHashParamKeys.ID_TOKEN]);
                        response.idToken = idTokenObj;
                        response.idTokenClaims = idTokenObj.claims;
                    }
                    else {
                        this.logger.verbose("No idToken on fragment, getting idToken from cache");
                        idTokenObj = new IdToken_1.IdToken(this.cacheStorage.getItem(Constants_1.PersistentCacheKeys.IDTOKEN));
                        response = ResponseUtils_1.ResponseUtils.setResponseIdToken(response, idTokenObj);
                    }
                    // set authority
                    var authority = this.populateAuthority(stateInfo.state, this.inCookie, this.cacheStorage, idTokenObj);
                    this.logger.verbose("Got authority from cache");
                    // retrieve client_info - if it is not found, generate the uid and utid from idToken
                    if (hashParams.hasOwnProperty(Constants_1.ServerHashParamKeys.CLIENT_INFO)) {
                        this.logger.verbose("Fragment has clientInfo");
                        clientInfo = hashParams[Constants_1.ServerHashParamKeys.CLIENT_INFO];
                    }
                    else {
                        this.logger.warning("ClientInfo not received in the response from AAD");
                        throw ClientAuthError_1.ClientAuthError.createClientInfoNotPopulatedError("ClientInfo not received in the response from the server");
                    }
                    response.account = Account_1.Account.createAccount(idTokenObj, new ClientInfo_1.ClientInfo(clientInfo));
                    this.logger.verbose("Account object created from response");
                    var accountKey = void 0;
                    if (response.account && !StringUtils_1.StringUtils.isEmpty(response.account.homeAccountIdentifier)) {
                        this.logger.verbose("AccountKey set");
                        accountKey = response.account.homeAccountIdentifier;
                    }
                    else {
                        this.logger.verbose("AccountKey set as no_account");
                        accountKey = Constants_1.Constants.no_account;
                    }
                    acquireTokenAccountKey = AuthCache_1.AuthCache.generateAcquireTokenAccountKey(accountKey, stateInfo.state);
                    var acquireTokenAccountKey_noaccount = AuthCache_1.AuthCache.generateAcquireTokenAccountKey(Constants_1.Constants.no_account, stateInfo.state);
                    this.logger.verbose("AcquireTokenAccountKey generated");
                    var cachedAccount = this.cacheStorage.getItem(acquireTokenAccountKey);
                    var acquireTokenAccount = void 0;
                    // Check with the account in the Cache
                    if (!StringUtils_1.StringUtils.isEmpty(cachedAccount)) {
                        acquireTokenAccount = JSON.parse(cachedAccount);
                        this.logger.verbose("AcquireToken request account retrieved from cache");
                        if (response.account && acquireTokenAccount && Account_1.Account.compareAccounts(response.account, acquireTokenAccount)) {
                            response = this.saveAccessToken(response, authority, hashParams, clientInfo, idTokenObj);
                            this.logger.info("The user object received in the response is the same as the one passed in the acquireToken request");
                        }
                        else {
                            this.logger.warning("The account object created from the response is not the same as the one passed in the acquireToken request");
                        }
                    }
                    else if (!StringUtils_1.StringUtils.isEmpty(this.cacheStorage.getItem(acquireTokenAccountKey_noaccount))) {
                        this.logger.verbose("No acquireToken account retrieved from cache");
                        response = this.saveAccessToken(response, authority, hashParams, clientInfo, idTokenObj);
                    }
                }
                // Process id_token
                if (hashParams.hasOwnProperty(Constants_1.ServerHashParamKeys.ID_TOKEN)) {
                    this.logger.info("Fragment has idToken");
                    // set the idToken
                    idTokenObj = new IdToken_1.IdToken(hashParams[Constants_1.ServerHashParamKeys.ID_TOKEN]);
                    response = ResponseUtils_1.ResponseUtils.setResponseIdToken(response, idTokenObj);
                    if (hashParams.hasOwnProperty(Constants_1.ServerHashParamKeys.CLIENT_INFO)) {
                        this.logger.verbose("Fragment has clientInfo");
                        clientInfo = hashParams[Constants_1.ServerHashParamKeys.CLIENT_INFO];
                    }
                    else {
                        this.logger.warning("ClientInfo not received in the response from AAD");
                    }
                    // set authority
                    var authority = this.populateAuthority(stateInfo.state, this.inCookie, this.cacheStorage, idTokenObj);
                    this.account = Account_1.Account.createAccount(idTokenObj, new ClientInfo_1.ClientInfo(clientInfo));
                    response.account = this.account;
                    this.logger.verbose("Account object created from response");
                    if (idTokenObj && idTokenObj.nonce) {
                        this.logger.verbose("IdToken has nonce");
                        // check nonce integrity if idToken has nonce - throw an error if not matched
                        var cachedNonce = this.cacheStorage.getItem(AuthCache_1.AuthCache.generateTemporaryCacheKey(Constants_1.TemporaryCacheKeys.NONCE_IDTOKEN, stateInfo.state), this.inCookie);
                        if (idTokenObj.nonce !== cachedNonce) {
                            this.account = null;
                            this.cacheStorage.setItem(Constants_1.ErrorCacheKeys.LOGIN_ERROR, "Nonce Mismatch. Expected Nonce: " + cachedNonce + "," + "Actual Nonce: " + idTokenObj.nonce);
                            this.logger.error("Nonce Mismatch. Expected Nonce: " + cachedNonce + ", Actual Nonce: " + idTokenObj.nonce);
                            error = ClientAuthError_1.ClientAuthError.createNonceMismatchError(cachedNonce, idTokenObj.nonce);
                        }
                        // Save the token
                        else {
                            this.logger.verbose("Nonce matches, saving idToken to cache");
                            this.cacheStorage.setItem(Constants_1.PersistentCacheKeys.IDTOKEN, hashParams[Constants_1.ServerHashParamKeys.ID_TOKEN], this.inCookie);
                            this.cacheStorage.setItem(Constants_1.PersistentCacheKeys.CLIENT_INFO, clientInfo, this.inCookie);
                            // Save idToken as access token for app itself
                            this.saveAccessToken(response, authority, hashParams, clientInfo, idTokenObj);
                        }
                    }
                    else {
                        this.logger.verbose("No idToken or no nonce. Cache key for Authority set as state");
                        authorityKey = stateInfo.state;
                        acquireTokenAccountKey = stateInfo.state;
                        this.logger.error("Invalid id_token received in the response");
                        error = ClientAuthError_1.ClientAuthError.createInvalidIdTokenError(idTokenObj);
                        this.cacheStorage.setItem(Constants_1.ErrorCacheKeys.ERROR, error.errorCode);
                        this.cacheStorage.setItem(Constants_1.ErrorCacheKeys.ERROR_DESC, error.errorMessage);
                    }
                }
            }
            // State mismatch - unexpected/invalid state
            else {
                this.logger.verbose("State mismatch");
                authorityKey = stateInfo.state;
                acquireTokenAccountKey = stateInfo.state;
                var expectedState = this.cacheStorage.getItem(AuthCache_1.AuthCache.generateTemporaryCacheKey(Constants_1.TemporaryCacheKeys.STATE_LOGIN, stateInfo.state), this.inCookie);
                this.logger.error("State Mismatch. Expected State: " + expectedState + ", Actual State: " + stateInfo.state);
                error = ClientAuthError_1.ClientAuthError.createInvalidStateError(stateInfo.state, expectedState);
                this.cacheStorage.setItem(Constants_1.ErrorCacheKeys.ERROR, error.errorCode);
                this.cacheStorage.setItem(Constants_1.ErrorCacheKeys.ERROR_DESC, error.errorMessage);
            }
        }
        // Set status to completed
        this.cacheStorage.removeItem(AuthCache_1.AuthCache.generateTemporaryCacheKey(Constants_1.TemporaryCacheKeys.RENEW_STATUS, stateInfo.state));
        this.cacheStorage.resetTempCacheItems(stateInfo.state);
        this.logger.verbose("Status set to complete, temporary cache cleared");
        // this is required if navigateToLoginRequestUrl=false
        if (this.inCookie) {
            this.logger.verbose("InCookie is true, setting authorityKey in cookie");
            this.cacheStorage.setItemCookie(authorityKey, "", -1);
            this.cacheStorage.clearMsalCookie(stateInfo.state);
        }
        if (error) {
            // Error case, set status to cancelled
            throw error;
        }
        if (!response) {
            throw AuthError_1.AuthError.createUnexpectedError("Response is null");
        }
        return response;
    };
    /**
     * Set Authority when saving Token from the hash
     * @param state
     * @param inCookie
     * @param cacheStorage
     * @param idTokenObj
     * @param response
     */
    UserAgentApplication.prototype.populateAuthority = function (state, inCookie, cacheStorage, idTokenObj) {
        this.logger.verbose("PopulateAuthority has been called");
        var authorityKey = AuthCache_1.AuthCache.generateAuthorityKey(state);
        var cachedAuthority = cacheStorage.getItem(authorityKey, inCookie);
        // retrieve the authority from cache and replace with tenantID
        return StringUtils_1.StringUtils.isEmpty(cachedAuthority) ? cachedAuthority : UrlUtils_1.UrlUtils.replaceTenantPath(cachedAuthority, idTokenObj.tenantId);
    };
    /* tslint:enable:no-string-literal */
    // #endregion
    // #region Account
    /**
     * Returns the signed in account
     * (the account object is created at the time of successful login)
     * or null when no state is found
     * @returns {@link Account} - the account object stored in MSAL
     */
    UserAgentApplication.prototype.getAccount = function () {
        // if a session already exists, get the account from the session
        if (this.account) {
            return this.account;
        }
        // frame is used to get idToken and populate the account for the given session
        var rawIdToken = this.cacheStorage.getItem(Constants_1.PersistentCacheKeys.IDTOKEN, this.inCookie);
        var rawClientInfo = this.cacheStorage.getItem(Constants_1.PersistentCacheKeys.CLIENT_INFO, this.inCookie);
        if (!StringUtils_1.StringUtils.isEmpty(rawIdToken) && !StringUtils_1.StringUtils.isEmpty(rawClientInfo)) {
            var idToken = new IdToken_1.IdToken(rawIdToken);
            var clientInfo = new ClientInfo_1.ClientInfo(rawClientInfo);
            this.account = Account_1.Account.createAccount(idToken, clientInfo);
            return this.account;
        }
        // if login not yet done, return null
        return null;
    };
    /**
     * @hidden
     *
     * Extracts state value from the accountState sent with the authentication request.
     * @returns {string} scope.
     * @ignore
     */
    UserAgentApplication.prototype.getAccountState = function (state) {
        if (state) {
            var splitIndex = state.indexOf(Constants_1.Constants.resourceDelimiter);
            if (splitIndex > -1 && splitIndex + 1 < state.length) {
                return state.substring(splitIndex + 1);
            }
        }
        return state;
    };
    /**
     * Use to get a list of unique accounts in MSAL cache based on homeAccountIdentifier.
     *
     * @param {@link Array<Account>} Account - all unique accounts in MSAL cache.
     */
    UserAgentApplication.prototype.getAllAccounts = function () {
        var accounts = [];
        var accessTokenCacheItems = this.cacheStorage.getAllAccessTokens(Constants_1.Constants.clientId, Constants_1.Constants.homeAccountIdentifier);
        for (var i = 0; i < accessTokenCacheItems.length; i++) {
            var idToken = new IdToken_1.IdToken(accessTokenCacheItems[i].value.idToken);
            var clientInfo = new ClientInfo_1.ClientInfo(accessTokenCacheItems[i].value.homeAccountIdentifier);
            var account = Account_1.Account.createAccount(idToken, clientInfo);
            accounts.push(account);
        }
        return this.getUniqueAccounts(accounts);
    };
    /**
     * @hidden
     *
     * Used to filter accounts based on homeAccountIdentifier
     * @param {Array<Account>}  Accounts - accounts saved in the cache
     * @ignore
     */
    UserAgentApplication.prototype.getUniqueAccounts = function (accounts) {
        if (!accounts || accounts.length <= 1) {
            return accounts;
        }
        var flags = [];
        var uniqueAccounts = [];
        for (var index = 0; index < accounts.length; ++index) {
            if (accounts[index].homeAccountIdentifier && flags.indexOf(accounts[index].homeAccountIdentifier) === -1) {
                flags.push(accounts[index].homeAccountIdentifier);
                uniqueAccounts.push(accounts[index]);
            }
        }
        return uniqueAccounts;
    };
    // #endregion
    // #region Angular
    /**
     * @hidden
     *
     * Broadcast messages - Used only for Angular?  *
     * @param eventName
     * @param data
     */
    UserAgentApplication.prototype.broadcast = function (eventName, data) {
        var evt = new CustomEvent(eventName, { detail: data });
        window.dispatchEvent(evt);
    };
    /**
     * @hidden
     *
     * Helper function to retrieve the cached token
     *
     * @param scopes
     * @param {@link Account} account
     * @param state
     * @return {@link AuthResponse} AuthResponse
     */
    UserAgentApplication.prototype.getCachedTokenInternal = function (scopes, account, state, correlationId) {
        // Get the current session's account object
        var accountObject = account || this.getAccount();
        if (!accountObject) {
            return null;
        }
        // Construct AuthenticationRequest based on response type; set "redirectUri" from the "request" which makes this call from Angular - for this.getRedirectUri()
        var newAuthority = this.authorityInstance ? this.authorityInstance : AuthorityFactory_1.AuthorityFactory.CreateInstance(this.authority, this.config.auth.validateAuthority);
        var responseType = this.getTokenType(accountObject, scopes, true);
        var serverAuthenticationRequest = new ServerRequestParameters_1.ServerRequestParameters(newAuthority, this.clientId, responseType, this.getRedirectUri(), scopes, state, correlationId);
        // get cached token
        return this.getCachedToken(serverAuthenticationRequest, account);
    };
    /**
     * @hidden
     *
     * Get scopes for the Endpoint - Used in Angular to track protected and unprotected resources without interaction from the developer app
     * Note: Please check if we need to set the "redirectUri" from the "request" which makes this call from Angular - for this.getRedirectUri()
     *
     * @param endpoint
     */
    UserAgentApplication.prototype.getScopesForEndpoint = function (endpoint) {
        // if user specified list of unprotectedResources, no need to send token to these endpoints, return null.
        if (this.config.framework.unprotectedResources.length > 0) {
            for (var i = 0; i < this.config.framework.unprotectedResources.length; i++) {
                if (endpoint.indexOf(this.config.framework.unprotectedResources[i]) > -1) {
                    return null;
                }
            }
        }
        // process all protected resources and send the matched one
        if (this.config.framework.protectedResourceMap.size > 0) {
            for (var _i = 0, _a = Array.from(this.config.framework.protectedResourceMap.keys()); _i < _a.length; _i++) {
                var key = _a[_i];
                // configEndpoint is like /api/Todo requested endpoint can be /api/Todo/1
                if (endpoint.indexOf(key) > -1) {
                    return this.config.framework.protectedResourceMap.get(key);
                }
            }
        }
        /*
         * default resource will be clientid if nothing specified
         * App will use idtoken for calls to itself
         * check if it's staring from http or https, needs to match with app host
         */
        if (endpoint.indexOf("http://") > -1 || endpoint.indexOf("https://") > -1) {
            if (UrlUtils_1.UrlUtils.getHostFromUri(endpoint) === UrlUtils_1.UrlUtils.getHostFromUri(this.getRedirectUri())) {
                return new Array(this.clientId);
            }
        }
        else {
            /*
             * in angular level, the url for $http interceptor call could be relative url,
             * if it's relative call, we'll treat it as app backend call.
             */
            return new Array(this.clientId);
        }
        // if not the app's own backend or not a domain listed in the endpoints structure
        return null;
    };
    /**
     * Return boolean flag to developer to help inform if login is in progress
     * @returns {boolean} true/false
     */
    UserAgentApplication.prototype.getLoginInProgress = function () {
        return this.cacheStorage.getItem(Constants_1.TemporaryCacheKeys.INTERACTION_STATUS) === Constants_1.Constants.inProgress;
    };
    /**
     * @hidden
     * @ignore
     *
     * @param loginInProgress
     */
    UserAgentApplication.prototype.setInteractionInProgress = function (inProgress) {
        if (inProgress) {
            this.cacheStorage.setItem(Constants_1.TemporaryCacheKeys.INTERACTION_STATUS, Constants_1.Constants.inProgress);
        }
        else {
            this.cacheStorage.removeItem(Constants_1.TemporaryCacheKeys.INTERACTION_STATUS);
        }
    };
    /**
     * @hidden
     * @ignore
     *
     * @param loginInProgress
     */
    UserAgentApplication.prototype.setloginInProgress = function (loginInProgress) {
        this.setInteractionInProgress(loginInProgress);
    };
    /**
     * @hidden
     * @ignore
     *
     * returns the status of acquireTokenInProgress
     */
    UserAgentApplication.prototype.getAcquireTokenInProgress = function () {
        return this.cacheStorage.getItem(Constants_1.TemporaryCacheKeys.INTERACTION_STATUS) === Constants_1.Constants.inProgress;
    };
    /**
     * @hidden
     * @ignore
     *
     * @param acquireTokenInProgress
     */
    UserAgentApplication.prototype.setAcquireTokenInProgress = function (acquireTokenInProgress) {
        this.setInteractionInProgress(acquireTokenInProgress);
    };
    /**
     * @hidden
     * @ignore
     *
     * returns the logger handle
     */
    UserAgentApplication.prototype.getLogger = function () {
        return this.logger;
    };
    /**
     * Sets the logger callback.
     * @param logger Logger callback
     */
    UserAgentApplication.prototype.setLogger = function (logger) {
        this.logger = logger;
    };
    // #endregion
    // #region Getters and Setters
    /**
     * Use to get the redirect uri configured in MSAL or null.
     * Evaluates redirectUri if its a function, otherwise simply returns its value.
     *
     * @returns {string} redirect URL
     */
    UserAgentApplication.prototype.getRedirectUri = function (reqRedirectUri) {
        if (reqRedirectUri) {
            return reqRedirectUri;
        }
        else if (typeof this.config.auth.redirectUri === "function") {
            return this.config.auth.redirectUri();
        }
        return this.config.auth.redirectUri;
    };
    /**
     * Use to get the post logout redirect uri configured in MSAL or null.
     * Evaluates postLogoutredirectUri if its a function, otherwise simply returns its value.
     *
     * @returns {string} post logout redirect URL
     */
    UserAgentApplication.prototype.getPostLogoutRedirectUri = function () {
        if (typeof this.config.auth.postLogoutRedirectUri === "function") {
            return this.config.auth.postLogoutRedirectUri();
        }
        return this.config.auth.postLogoutRedirectUri;
    };
    /**
     * Use to get the current {@link Configuration} object in MSAL
     *
     * @returns {@link Configuration}
     */
    UserAgentApplication.prototype.getCurrentConfiguration = function () {
        if (!this.config) {
            throw ClientConfigurationError_1.ClientConfigurationError.createNoSetConfigurationError();
        }
        return this.config;
    };
    /**
     * @ignore
     *
     * Utils function to create the Authentication
     * @param {@link account} account object
     * @param scopes
     * @param silentCall
     *
     * @returns {string} token type: id_token or access_token
     *
     */
    UserAgentApplication.prototype.getTokenType = function (accountObject, scopes, silentCall) {
        /*
         * if account is passed and matches the account object/or set to getAccount() from cache
         * if client-id is passed as scope, get id_token else token/id_token_token (in case no session exists)
         */
        var tokenType;
        // acquireTokenSilent
        if (silentCall) {
            if (Account_1.Account.compareAccounts(accountObject, this.getAccount())) {
                tokenType = (scopes.indexOf(this.config.auth.clientId) > -1) ? ResponseTypes.id_token : ResponseTypes.token;
            }
            else {
                tokenType = (scopes.indexOf(this.config.auth.clientId) > -1) ? ResponseTypes.id_token : ResponseTypes.id_token_token;
            }
            return tokenType;
        }
        // all other cases
        else {
            if (!Account_1.Account.compareAccounts(accountObject, this.getAccount())) {
                tokenType = ResponseTypes.id_token_token;
            }
            else {
                tokenType = (scopes.indexOf(this.clientId) > -1) ? ResponseTypes.id_token : ResponseTypes.token;
            }
            return tokenType;
        }
    };
    /**
     * @hidden
     * @ignore
     *
     * Sets the cachekeys for and stores the account information in cache
     * @param account
     * @param state
     * @hidden
     */
    UserAgentApplication.prototype.setAccountCache = function (account, state) {
        // Cache acquireTokenAccountKey
        var accountId = account ? this.getAccountId(account) : Constants_1.Constants.no_account;
        var acquireTokenAccountKey = AuthCache_1.AuthCache.generateAcquireTokenAccountKey(accountId, state);
        this.cacheStorage.setItem(acquireTokenAccountKey, JSON.stringify(account));
    };
    /**
     * @hidden
     * @ignore
     *
     * Sets the cacheKey for and stores the authority information in cache
     * @param state
     * @param authority
     * @hidden
     */
    UserAgentApplication.prototype.setAuthorityCache = function (state, authority) {
        // Cache authorityKey
        var authorityKey = AuthCache_1.AuthCache.generateAuthorityKey(state);
        this.cacheStorage.setItem(authorityKey, UrlUtils_1.UrlUtils.CanonicalizeUri(authority), this.inCookie);
    };
    /**
     * Updates account, authority, and nonce in cache
     * @param serverAuthenticationRequest
     * @param account
     * @hidden
     * @ignore
     */
    UserAgentApplication.prototype.updateCacheEntries = function (serverAuthenticationRequest, account, isLoginCall, loginStartPage) {
        // Cache Request Originator Page
        if (loginStartPage) {
            this.cacheStorage.setItem(AuthCache_1.AuthCache.generateTemporaryCacheKey(Constants_1.TemporaryCacheKeys.LOGIN_REQUEST, serverAuthenticationRequest.state), loginStartPage, this.inCookie);
        }
        // Cache account and authority
        if (isLoginCall) {
            // Cache the state
            this.cacheStorage.setItem(AuthCache_1.AuthCache.generateTemporaryCacheKey(Constants_1.TemporaryCacheKeys.STATE_LOGIN, serverAuthenticationRequest.state), serverAuthenticationRequest.state, this.inCookie);
        }
        else {
            this.setAccountCache(account, serverAuthenticationRequest.state);
        }
        // Cache authorityKey
        this.setAuthorityCache(serverAuthenticationRequest.state, serverAuthenticationRequest.authority);
        // Cache nonce
        this.cacheStorage.setItem(AuthCache_1.AuthCache.generateTemporaryCacheKey(Constants_1.TemporaryCacheKeys.NONCE_IDTOKEN, serverAuthenticationRequest.state), serverAuthenticationRequest.nonce, this.inCookie);
    };
    /**
     * Returns the unique identifier for the logged in account
     * @param account
     * @hidden
     * @ignore
     */
    UserAgentApplication.prototype.getAccountId = function (account) {
        // return `${account.accountIdentifier}` + Constants.resourceDelimiter + `${account.homeAccountIdentifier}`;
        var accountId;
        if (!StringUtils_1.StringUtils.isEmpty(account.homeAccountIdentifier)) {
            accountId = account.homeAccountIdentifier;
        }
        else {
            accountId = Constants_1.Constants.no_account;
        }
        return accountId;
    };
    /**
     * @ignore
     * @param extraQueryParameters
     *
     * Construct 'tokenRequest' from the available data in adalIdToken
     */
    UserAgentApplication.prototype.buildIDTokenRequest = function (request) {
        var tokenRequest = {
            scopes: [this.clientId],
            authority: this.authority,
            account: this.getAccount(),
            extraQueryParameters: request.extraQueryParameters,
            correlationId: request.correlationId
        };
        return tokenRequest;
    };
    /**
     * @ignore
     * @param config
     * @param clientId
     *
     * Construct TelemetryManager from Configuration
     */
    UserAgentApplication.prototype.getTelemetryManagerFromConfig = function (config, clientId) {
        if (!config) { // if unset
            return TelemetryManager_1.default.getTelemetrymanagerStub(clientId, this.logger);
        }
        // if set then validate
        var applicationName = config.applicationName, applicationVersion = config.applicationVersion, telemetryEmitter = config.telemetryEmitter;
        if (!applicationName || !applicationVersion || !telemetryEmitter) {
            throw ClientConfigurationError_1.ClientConfigurationError.createTelemetryConfigError(config);
        }
        // if valid then construct
        var telemetryPlatform = {
            applicationName: applicationName,
            applicationVersion: applicationVersion
        };
        var telemetryManagerConfig = {
            platform: telemetryPlatform,
            clientId: clientId
        };
        return new TelemetryManager_1.default(telemetryManagerConfig, telemetryEmitter, this.logger);
    };
    return UserAgentApplication;
}());
exports.UserAgentApplication = UserAgentApplication;


/***/ }),
/* 16 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
var CryptoUtils_1 = __webpack_require__(2);
var Constants_1 = __webpack_require__(1);
var StringUtils_1 = __webpack_require__(3);
var ScopeSet_1 = __webpack_require__(9);
/**
 * Nonce: OIDC Nonce definition: https://openid.net/specs/openid-connect-core-1_0.html#IDToken
 * State: OAuth Spec: https://tools.ietf.org/html/rfc6749#section-10.12
 * @hidden
 */
var ServerRequestParameters = /** @class */ (function () {
    /**
     * Constructor
     * @param authority
     * @param clientId
     * @param scope
     * @param responseType
     * @param redirectUri
     * @param state
     */
    function ServerRequestParameters(authority, clientId, responseType, redirectUri, scopes, state, correlationId) {
        this.authorityInstance = authority;
        this.clientId = clientId;
        this.nonce = CryptoUtils_1.CryptoUtils.createNewGuid();
        // set scope to clientId if null
        this.scopes = scopes ? scopes.slice() : [clientId];
        this.scopes = ScopeSet_1.ScopeSet.trimAndConvertArrayToLowerCase(this.scopes);
        // set state (already set at top level)
        this.state = state;
        // set correlationId
        this.correlationId = correlationId;
        // telemetry information
        this.xClientSku = "MSAL.JS";
        this.xClientVer = Constants_1.libraryVersion();
        this.responseType = responseType;
        this.redirectUri = redirectUri;
    }
    Object.defineProperty(ServerRequestParameters.prototype, "authority", {
        get: function () {
            return this.authorityInstance ? this.authorityInstance.CanonicalAuthority : null;
        },
        enumerable: true,
        configurable: true
    });
    /**
     * @hidden
     * @ignore
     *
     * Utility to populate QueryParameters and ExtraQueryParameters to ServerRequestParamerers
     * @param request
     * @param serverAuthenticationRequest
     */
    ServerRequestParameters.prototype.populateQueryParams = function (account, request, adalIdTokenObject, silentCall) {
        var queryParameters = {};
        if (request) {
            // add the prompt parameter to serverRequestParameters if passed
            if (request.prompt) {
                this.promptValue = request.prompt;
            }
            // Add claims challenge to serverRequestParameters if passed
            if (request.claimsRequest) {
                this.claimsValue = request.claimsRequest;
            }
            // if the developer provides one of these, give preference to developer choice
            if (ServerRequestParameters.isSSOParam(request)) {
                queryParameters = this.constructUnifiedCacheQueryParameter(request, null);
            }
        }
        if (adalIdTokenObject) {
            queryParameters = this.constructUnifiedCacheQueryParameter(null, adalIdTokenObject);
        }
        /*
         * adds sid/login_hint if not populated
         * this.logger.verbose("Calling addHint parameters");
         */
        queryParameters = this.addHintParameters(account, queryParameters);
        // sanity check for developer passed extraQueryParameters
        var eQParams = request ? request.extraQueryParameters : null;
        // Populate the extraQueryParameters to be sent to the server
        this.queryParameters = ServerRequestParameters.generateQueryParametersString(queryParameters);
        this.extraQueryParameters = ServerRequestParameters.generateQueryParametersString(eQParams, silentCall);
    };
    // #region QueryParam helpers
    /**
     * Constructs extraQueryParameters to be sent to the server for the AuthenticationParameters set by the developer
     * in any login() or acquireToken() calls
     * @param idTokenObject
     * @param extraQueryParameters
     * @param sid
     * @param loginHint
     */
    // TODO: check how this behaves when domain_hint only is sent in extraparameters and idToken has no upn.
    ServerRequestParameters.prototype.constructUnifiedCacheQueryParameter = function (request, idTokenObject) {
        // preference order: account > sid > login_hint
        var ssoType;
        var ssoData;
        var serverReqParam = {};
        // if account info is passed, account.sid > account.login_hint
        if (request) {
            if (request.account) {
                var account = request.account;
                if (account.sid) {
                    ssoType = Constants_1.SSOTypes.SID;
                    ssoData = account.sid;
                }
                else if (account.userName) {
                    ssoType = Constants_1.SSOTypes.LOGIN_HINT;
                    ssoData = account.userName;
                }
            }
            // sid from request
            else if (request.sid) {
                ssoType = Constants_1.SSOTypes.SID;
                ssoData = request.sid;
            }
            // loginHint from request
            else if (request.loginHint) {
                ssoType = Constants_1.SSOTypes.LOGIN_HINT;
                ssoData = request.loginHint;
            }
        }
        // adalIdToken retrieved from cache
        else if (idTokenObject) {
            if (idTokenObject.hasOwnProperty(Constants_1.Constants.upn)) {
                ssoType = Constants_1.SSOTypes.ID_TOKEN;
                ssoData = idTokenObject.upn;
            }
        }
        serverReqParam = this.addSSOParameter(ssoType, ssoData);
        return serverReqParam;
    };
    /**
     * @hidden
     *
     * Adds login_hint to authorization URL which is used to pre-fill the username field of sign in page for the user if known ahead of time
     * domain_hint if added skips the email based discovery process of the user - only supported for interactive calls in implicit_flow
     * domain_req utid received as part of the clientInfo
     * login_req uid received as part of clientInfo
     * Also does a sanity check for extraQueryParameters passed by the user to ensure no repeat queryParameters
     *
     * @param {@link Account} account - Account for which the token is requested
     * @param queryparams
     * @param {@link ServerRequestParameters}
     * @ignore
     */
    ServerRequestParameters.prototype.addHintParameters = function (account, qParams) {
        /*
         * This is a final check for all queryParams added so far; preference order: sid > login_hint
         * sid cannot be passed along with login_hint or domain_hint, hence we check both are not populated yet in queryParameters
         */
        if (account && !qParams[Constants_1.SSOTypes.SID]) {
            // sid - populate only if login_hint is not already populated and the account has sid
            var populateSID = !qParams[Constants_1.SSOTypes.LOGIN_HINT] && account.sid && this.promptValue === Constants_1.PromptState.NONE;
            if (populateSID) {
                qParams = this.addSSOParameter(Constants_1.SSOTypes.SID, account.sid, qParams);
            }
            // login_hint - account.userName
            else {
                var populateLoginHint = !qParams[Constants_1.SSOTypes.LOGIN_HINT] && account.userName && !StringUtils_1.StringUtils.isEmpty(account.userName);
                if (populateLoginHint) {
                    qParams = this.addSSOParameter(Constants_1.SSOTypes.LOGIN_HINT, account.userName, qParams);
                }
            }
        }
        return qParams;
    };
    /**
     * Add SID to extraQueryParameters
     * @param sid
     */
    ServerRequestParameters.prototype.addSSOParameter = function (ssoType, ssoData, ssoParam) {
        if (!ssoParam) {
            ssoParam = {};
        }
        if (!ssoData) {
            return ssoParam;
        }
        switch (ssoType) {
            case Constants_1.SSOTypes.SID: {
                ssoParam[Constants_1.SSOTypes.SID] = ssoData;
                break;
            }
            case Constants_1.SSOTypes.ID_TOKEN: {
                ssoParam[Constants_1.SSOTypes.LOGIN_HINT] = ssoData;
                break;
            }
            case Constants_1.SSOTypes.LOGIN_HINT: {
                ssoParam[Constants_1.SSOTypes.LOGIN_HINT] = ssoData;
                break;
            }
        }
        return ssoParam;
    };
    /**
     * Utility to generate a QueryParameterString from a Key-Value mapping of extraQueryParameters passed
     * @param extraQueryParameters
     */
    ServerRequestParameters.generateQueryParametersString = function (queryParameters, silentCall) {
        var paramsString = null;
        if (queryParameters) {
            Object.keys(queryParameters).forEach(function (key) {
                // sid cannot be passed along with login_hint or domain_hint
                if (key === Constants_1.Constants.domain_hint && (silentCall || queryParameters[Constants_1.SSOTypes.SID])) {
                    return;
                }
                if (paramsString == null) {
                    paramsString = key + "=" + encodeURIComponent(queryParameters[key]);
                }
                else {
                    paramsString += "&" + key + "=" + encodeURIComponent(queryParameters[key]);
                }
            });
        }
        return paramsString;
    };
    // #endregion
    /**
     * Check to see if there are SSO params set in the Request
     * @param request
     */
    ServerRequestParameters.isSSOParam = function (request) {
        return request && (request.account || request.sid || request.loginHint);
    };
    return ServerRequestParameters;
}());
exports.ServerRequestParameters = ServerRequestParameters;


/***/ }),
/* 17 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
var CryptoUtils_1 = __webpack_require__(2);
var StringUtils_1 = __webpack_require__(3);
/**
 * @hidden
 */
var TokenUtils = /** @class */ (function () {
    function TokenUtils() {
    }
    /**
     * decode a JWT
     *
     * @param jwtToken
     */
    TokenUtils.decodeJwt = function (jwtToken) {
        if (StringUtils_1.StringUtils.isEmpty(jwtToken)) {
            return null;
        }
        var idTokenPartsRegex = /^([^\.\s]*)\.([^\.\s]+)\.([^\.\s]*)$/;
        var matches = idTokenPartsRegex.exec(jwtToken);
        if (!matches || matches.length < 4) {
            // this._requestContext.logger.warn("The returned id_token is not parseable.");
            return null;
        }
        var crackedToken = {
            header: matches[1],
            JWSPayload: matches[2],
            JWSSig: matches[3]
        };
        return crackedToken;
    };
    /**
     * Extract IdToken by decoding the RAWIdToken
     *
     * @param encodedIdToken
     */
    TokenUtils.extractIdToken = function (encodedIdToken) {
        // id token will be decoded to get the username
        var decodedToken = this.decodeJwt(encodedIdToken);
        if (!decodedToken) {
            return null;
        }
        try {
            var base64IdToken = decodedToken.JWSPayload;
            var base64Decoded = CryptoUtils_1.CryptoUtils.base64Decode(base64IdToken);
            if (!base64Decoded) {
                // this._requestContext.logger.info("The returned id_token could not be base64 url safe decoded.");
                return null;
            }
            // ECMA script has JSON built-in support
            return JSON.parse(base64Decoded);
        }
        catch (err) {
            // this._requestContext.logger.error("The returned id_token could not be decoded" + err);
        }
        return null;
    };
    return TokenUtils;
}());
exports.TokenUtils = TokenUtils;


/***/ }),
/* 18 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = __webpack_require__(0);
var Constants_1 = __webpack_require__(1);
var ClientConfigurationError_1 = __webpack_require__(5);
var ScopeSet_1 = __webpack_require__(9);
var StringUtils_1 = __webpack_require__(3);
var CryptoUtils_1 = __webpack_require__(2);
var TimeUtils_1 = __webpack_require__(11);
var ClientAuthError_1 = __webpack_require__(6);
/**
 * @hidden
 */
var RequestUtils = /** @class */ (function () {
    function RequestUtils() {
    }
    /**
     * @ignore
     *
     * @param request
     * @param isLoginCall
     * @param cacheStorage
     * @param clientId
     *
     * validates all request parameters and generates a consumable request object
     */
    RequestUtils.validateRequest = function (request, isLoginCall, clientId, interactionType) {
        // Throw error if request is empty for acquire * calls
        if (!isLoginCall && !request) {
            throw ClientConfigurationError_1.ClientConfigurationError.createEmptyRequestError();
        }
        var scopes;
        var extraQueryParameters;
        if (request) {
            // if extraScopesToConsent is passed in loginCall, append them to the login request; Validate and filter scopes (the validate function will throw if validation fails)
            scopes = isLoginCall ? ScopeSet_1.ScopeSet.appendScopes(request.scopes, request.extraScopesToConsent) : request.scopes;
            ScopeSet_1.ScopeSet.validateInputScope(scopes, !isLoginCall, clientId);
            // validate prompt parameter
            this.validatePromptParameter(request.prompt);
            // validate extraQueryParameters
            extraQueryParameters = this.validateEQParameters(request.extraQueryParameters, request.claimsRequest);
            // validate claimsRequest
            this.validateClaimsRequest(request.claimsRequest);
        }
        // validate and generate state and correlationId
        var state = this.validateAndGenerateState(request && request.state, interactionType);
        var correlationId = this.validateAndGenerateCorrelationId(request && request.correlationId);
        var validatedRequest = tslib_1.__assign({}, request, { extraQueryParameters: extraQueryParameters,
            scopes: scopes,
            state: state,
            correlationId: correlationId });
        return validatedRequest;
    };
    /**
     * @ignore
     *
     * Utility to test if valid prompt value is passed in the request
     * @param request
     */
    RequestUtils.validatePromptParameter = function (prompt) {
        if (prompt) {
            if ([Constants_1.PromptState.LOGIN, Constants_1.PromptState.SELECT_ACCOUNT, Constants_1.PromptState.CONSENT, Constants_1.PromptState.NONE].indexOf(prompt) < 0) {
                throw ClientConfigurationError_1.ClientConfigurationError.createInvalidPromptError(prompt);
            }
        }
    };
    /**
     * @ignore
     *
     * Removes unnecessary or duplicate query parameters from extraQueryParameters
     * @param request
     */
    RequestUtils.validateEQParameters = function (extraQueryParameters, claimsRequest) {
        var eQParams = tslib_1.__assign({}, extraQueryParameters);
        if (!eQParams) {
            return null;
        }
        if (claimsRequest) {
            // this.logger.warning("Removed duplicate claims from extraQueryParameters. Please use either the claimsRequest field OR pass as extraQueryParameter - not both.");
            delete eQParams[Constants_1.Constants.claims];
        }
        Constants_1.BlacklistedEQParams.forEach(function (param) {
            if (eQParams[param]) {
                // this.logger.warning("Removed duplicate " + param + " from extraQueryParameters. Please use the " + param + " field in request object.");
                delete eQParams[param];
            }
        });
        return eQParams;
    };
    /**
     * @ignore
     *
     * Validates the claims passed in request is a JSON
     * TODO: More validation will be added when the server team tells us how they have actually implemented claims
     * @param claimsRequest
     */
    RequestUtils.validateClaimsRequest = function (claimsRequest) {
        if (!claimsRequest) {
            return;
        }
        var claims;
        try {
            claims = JSON.parse(claimsRequest);
        }
        catch (e) {
            throw ClientConfigurationError_1.ClientConfigurationError.createClaimsRequestParsingError(e);
        }
    };
    /**
     * @ignore
     *
     * generate unique state per request
     * @param userState User-provided state value
     * @returns State string include library state and user state
     */
    RequestUtils.validateAndGenerateState = function (userState, interactionType) {
        return !StringUtils_1.StringUtils.isEmpty(userState) ? "" + RequestUtils.generateLibraryState(interactionType) + Constants_1.Constants.resourceDelimiter + userState : RequestUtils.generateLibraryState(interactionType);
    };
    /**
     * Generates the state value used by the library.
     *
     * @returns Base64 encoded string representing the state
     */
    RequestUtils.generateLibraryState = function (interactionType) {
        var stateObject = {
            id: CryptoUtils_1.CryptoUtils.createNewGuid(),
            ts: TimeUtils_1.TimeUtils.now(),
            method: interactionType
        };
        var stateString = JSON.stringify(stateObject);
        return CryptoUtils_1.CryptoUtils.base64Encode(stateString);
    };
    /**
     * Decodes the state value into a StateObject
     *
     * @param state State value returned in the request
     * @returns Parsed values from the encoded state value
     */
    RequestUtils.parseLibraryState = function (state) {
        var libraryState = decodeURIComponent(state).split(Constants_1.Constants.resourceDelimiter)[0];
        if (CryptoUtils_1.CryptoUtils.isGuid(libraryState)) {
            // If state is guid, assume timestamp is now and is redirect, as redirect should be only method where this can happen.
            return {
                id: libraryState,
                ts: TimeUtils_1.TimeUtils.now(),
                method: Constants_1.Constants.interactionTypeRedirect
            };
        }
        try {
            var stateString = CryptoUtils_1.CryptoUtils.base64Decode(libraryState);
            var stateObject = JSON.parse(stateString);
            return stateObject;
        }
        catch (e) {
            throw ClientAuthError_1.ClientAuthError.createInvalidStateError(state, null);
        }
    };
    /**
     * @ignore
     *
     * validate correlationId and generate if not valid or not set by the user
     * @param correlationId
     */
    RequestUtils.validateAndGenerateCorrelationId = function (correlationId) {
        // validate user set correlationId or set one for the user if null
        if (correlationId && !CryptoUtils_1.CryptoUtils.isGuid(correlationId)) {
            throw ClientConfigurationError_1.ClientConfigurationError.createInvalidCorrelationIdError();
        }
        return CryptoUtils_1.CryptoUtils.isGuid(correlationId) ? correlationId : CryptoUtils_1.CryptoUtils.createNewGuid();
    };
    /**
     * Create a request signature
     * @param request
     */
    RequestUtils.createRequestSignature = function (request) {
        return "" + request.scopes.join(" ").toLowerCase() + Constants_1.Constants.resourceDelimiter + request.authority;
    };
    return RequestUtils;
}());
exports.RequestUtils = RequestUtils;


/***/ }),
/* 19 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
var CryptoUtils_1 = __webpack_require__(2);
var StringUtils_1 = __webpack_require__(3);
/**
 * accountIdentifier       combination of idToken.uid and idToken.utid
 * homeAccountIdentifier   combination of clientInfo.uid and clientInfo.utid
 * userName                idToken.preferred_username
 * name                    idToken.name
 * idToken                 idToken
 * sid                     idToken.sid - session identifier
 * environment             idtoken.issuer (the authority that issues the token)
 */
var Account = /** @class */ (function () {
    /**
     * Creates an Account Object
     * @praram accountIdentifier
     * @param homeAccountIdentifier
     * @param userName
     * @param name
     * @param idToken
     * @param sid
     * @param environment
     */
    function Account(accountIdentifier, homeAccountIdentifier, userName, name, idTokenClaims, sid, environment) {
        this.accountIdentifier = accountIdentifier;
        this.homeAccountIdentifier = homeAccountIdentifier;
        this.userName = userName;
        this.name = name;
        // will be deprecated soon
        this.idToken = idTokenClaims;
        this.idTokenClaims = idTokenClaims;
        this.sid = sid;
        this.environment = environment;
    }
    /**
     * @hidden
     * @param idToken
     * @param clientInfo
     */
    Account.createAccount = function (idToken, clientInfo) {
        // create accountIdentifier
        var accountIdentifier = idToken.objectId || idToken.subject;
        // create homeAccountIdentifier
        var uid = clientInfo ? clientInfo.uid : "";
        var utid = clientInfo ? clientInfo.utid : "";
        var homeAccountIdentifier;
        if (!StringUtils_1.StringUtils.isEmpty(uid) && !StringUtils_1.StringUtils.isEmpty(utid)) {
            homeAccountIdentifier = CryptoUtils_1.CryptoUtils.base64Encode(uid) + "." + CryptoUtils_1.CryptoUtils.base64Encode(utid);
        }
        return new Account(accountIdentifier, homeAccountIdentifier, idToken.preferredName, idToken.name, idToken.claims, idToken.sid, idToken.issuer);
    };
    /**
     * Utils function to compare two Account objects - used to check if the same user account is logged in
     *
     * @param a1: Account object
     * @param a2: Account object
     */
    Account.compareAccounts = function (a1, a2) {
        if (!a1 || !a2) {
            return false;
        }
        if (a1.homeAccountIdentifier && a2.homeAccountIdentifier) {
            if (a1.homeAccountIdentifier === a2.homeAccountIdentifier) {
                return true;
            }
        }
        return false;
    };
    return Account;
}());
exports.Account = Account;


/***/ }),
/* 20 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var ClientAuthError_1 = __webpack_require__(6);
var UrlUtils_1 = __webpack_require__(4);
var Constants_1 = __webpack_require__(1);
var TimeUtils_1 = __webpack_require__(11);
var WindowUtils = /** @class */ (function () {
    function WindowUtils() {
    }
    /**
     * @hidden
     * Checks if the current page is running in an iframe.
     * @ignore
     */
    WindowUtils.isInIframe = function () {
        return window.parent !== window;
    };
    /**
     * @hidden
     * Check if the current page is running in a popup.
     * @ignore
     */
    WindowUtils.isInPopup = function () {
        return !!(window.opener && window.opener !== window);
    };
    /**
     * @hidden
     * @param prefix
     * @param scopes
     * @param authority
     */
    WindowUtils.generateFrameName = function (prefix, requestSignature) {
        return "" + prefix + Constants_1.Constants.resourceDelimiter + requestSignature;
    };
    /**
     * @hidden
     * Polls an iframe until it loads a url with a hash
     * @ignore
     */
    WindowUtils.monitorIframeForHash = function (contentWindow, timeout, urlNavigate, logger) {
        return new Promise(function (resolve, reject) {
            /*
             * Polling for iframes can be purely timing based,
             * since we don't need to account for interaction.
             */
            var nowMark = TimeUtils_1.TimeUtils.relativeNowMs();
            var timeoutMark = nowMark + timeout;
            logger.verbose("monitorWindowForIframe polling started");
            var intervalId = setInterval(function () {
                if (TimeUtils_1.TimeUtils.relativeNowMs() > timeoutMark) {
                    logger.error("monitorIframeForHash unable to find hash in url, timing out");
                    logger.errorPii("monitorIframeForHash polling timed out for url: " + urlNavigate);
                    clearInterval(intervalId);
                    reject(ClientAuthError_1.ClientAuthError.createTokenRenewalTimeoutError());
                    return;
                }
                var href;
                try {
                    /*
                     * Will throw if cross origin,
                     * which should be caught and ignored
                     * since we need the interval to keep running while on STS UI.
                     */
                    href = contentWindow.location.href;
                }
                catch (e) { }
                if (href && UrlUtils_1.UrlUtils.urlContainsHash(href)) {
                    logger.verbose("monitorIframeForHash found url in hash");
                    clearInterval(intervalId);
                    resolve(contentWindow.location.hash);
                }
            }, WindowUtils.POLLING_INTERVAL_MS);
        });
    };
    /**
     * @hidden
     * Polls a popup until it loads a url with a hash
     * @ignore
     */
    WindowUtils.monitorPopupForHash = function (contentWindow, timeout, urlNavigate, logger) {
        return new Promise(function (resolve, reject) {
            /*
             * Polling for popups needs to be tick-based,
             * since a non-trivial amount of time can be spent on interaction (which should not count against the timeout).
             */
            var maxTicks = timeout / WindowUtils.POLLING_INTERVAL_MS;
            var ticks = 0;
            logger.verbose("monitorWindowForHash polling started");
            var intervalId = setInterval(function () {
                if (contentWindow.closed) {
                    logger.error("monitorWindowForHash window closed");
                    clearInterval(intervalId);
                    reject(ClientAuthError_1.ClientAuthError.createUserCancelledError());
                    return;
                }
                var href;
                try {
                    /*
                     * Will throw if cross origin,
                     * which should be caught and ignored
                     * since we need the interval to keep running while on STS UI.
                     */
                    href = contentWindow.location.href;
                }
                catch (e) { }
                // Don't process blank pages or cross domain
                if (!href || href === "about:blank") {
                    return;
                }
                /*
                 * Only run clock when we are on same domain for popups
                 * as popup operations can take a long time.
                 */
                ticks++;
                if (href && UrlUtils_1.UrlUtils.urlContainsHash(href)) {
                    logger.verbose("monitorPopupForHash found url in hash");
                    clearInterval(intervalId);
                    resolve(contentWindow.location.hash);
                }
                else if (ticks > maxTicks) {
                    logger.error("monitorPopupForHash unable to find hash in url, timing out");
                    logger.errorPii("monitorPopupForHash polling timed out for url: " + urlNavigate);
                    clearInterval(intervalId);
                    reject(ClientAuthError_1.ClientAuthError.createTokenRenewalTimeoutError());
                }
            }, WindowUtils.POLLING_INTERVAL_MS);
        });
    };
    /**
     * @hidden
     * Loads iframe with authorization endpoint URL
     * @ignore
     */
    WindowUtils.loadFrame = function (urlNavigate, frameName, timeoutMs, logger) {
        var _this = this;
        /*
         * This trick overcomes iframe navigation in IE
         * IE does not load the page consistently in iframe
         */
        logger.infoPii("LoadFrame: " + frameName);
        return new Promise(function (resolve, reject) {
            setTimeout(function () {
                var frameHandle = _this.loadFrameSync(urlNavigate, frameName, logger);
                if (!frameHandle) {
                    reject("Unable to load iframe with name: " + frameName);
                    return;
                }
                resolve(frameHandle);
            }, timeoutMs);
        });
    };
    /**
     * @hidden
     * Loads the iframe synchronously when the navigateTimeFrame is set to `0`
     * @param urlNavigate
     * @param frameName
     * @param logger
     */
    WindowUtils.loadFrameSync = function (urlNavigate, frameName, logger) {
        var frameHandle = WindowUtils.addHiddenIFrame(frameName, logger);
        // returning to handle null in loadFrame, also to avoid null object access errors
        if (!frameHandle) {
            return null;
        }
        else if (frameHandle.src === "" || frameHandle.src === "about:blank") {
            frameHandle.src = urlNavigate;
            logger.infoPii("Frame Name : " + frameName + " Navigated to: " + urlNavigate);
        }
        return frameHandle;
    };
    /**
     * @hidden
     * Adds the hidden iframe for silent token renewal.
     * @ignore
     */
    WindowUtils.addHiddenIFrame = function (iframeId, logger) {
        if (typeof iframeId === "undefined") {
            return null;
        }
        logger.infoPii("Add msal frame to document:" + iframeId);
        var adalFrame = document.getElementById(iframeId);
        if (!adalFrame) {
            if (document.createElement &&
                document.documentElement &&
                (window.navigator.userAgent.indexOf("MSIE 5.0") === -1)) {
                var ifr = document.createElement("iframe");
                ifr.setAttribute("id", iframeId);
                ifr.setAttribute("aria-hidden", "true");
                ifr.style.visibility = "hidden";
                ifr.style.position = "absolute";
                ifr.style.width = ifr.style.height = "0";
                ifr.style.border = "0";
                ifr.setAttribute("sandbox", "allow-scripts allow-same-origin allow-forms");
                adalFrame = document.getElementsByTagName("body")[0].appendChild(ifr);
            }
            else if (document.body && document.body.insertAdjacentHTML) {
                document.body.insertAdjacentHTML("beforeend", "<iframe name='" + iframeId + "' id='" + iframeId + "' style='display:none'></iframe>");
            }
            if (window.frames && window.frames[iframeId]) {
                adalFrame = window.frames[iframeId];
            }
        }
        return adalFrame;
    };
    /**
     * @hidden
     * Removes a hidden iframe from the page.
     * @ignore
     */
    WindowUtils.removeHiddenIframe = function (iframe) {
        if (document.body === iframe.parentNode) {
            document.body.removeChild(iframe);
        }
    };
    /**
     * @hidden
     * Find and return the iframe element with the given hash
     * @ignore
     */
    WindowUtils.getIframeWithHash = function (hash) {
        var iframes = document.getElementsByTagName("iframe");
        var iframeArray = Array.apply(null, Array(iframes.length)).map(function (iframe, index) { return iframes.item(index); }); // eslint-disable-line prefer-spread
        return iframeArray.filter(function (iframe) {
            try {
                return iframe.contentWindow.location.hash === hash;
            }
            catch (e) {
                return false;
            }
        })[0];
    };
    /**
     * @hidden
     * Returns an array of all the popups opened by MSAL
     * @ignore
     */
    WindowUtils.getPopups = function () {
        if (!window.openedWindows) {
            window.openedWindows = [];
        }
        return window.openedWindows;
    };
    /**
     * @hidden
     * Find and return the popup with the given hash
     * @ignore
     */
    WindowUtils.getPopUpWithHash = function (hash) {
        return WindowUtils.getPopups().filter(function (popup) {
            try {
                return popup.location.hash === hash;
            }
            catch (e) {
                return false;
            }
        })[0];
    };
    /**
     * @hidden
     * Add the popup to the known list of popups
     * @ignore
     */
    WindowUtils.trackPopup = function (popup) {
        WindowUtils.getPopups().push(popup);
    };
    /**
     * @hidden
     * Close all popups
     * @ignore
     */
    WindowUtils.closePopups = function () {
        WindowUtils.getPopups().forEach(function (popup) { return popup.close(); });
    };
    /**
     * @ignore
     *
     * blocks any login/acquireToken calls to reload from within a hidden iframe (generated for silent calls)
     */
    WindowUtils.blockReloadInHiddenIframes = function () {
        // return an error if called from the hidden iframe created by the msal js silent calls
        if (UrlUtils_1.UrlUtils.urlContainsHash(window.location.hash) && WindowUtils.isInIframe()) {
            throw ClientAuthError_1.ClientAuthError.createBlockTokenRequestsInHiddenIframeError();
        }
    };
    /**
     *
     * @param cacheStorage
     */
    WindowUtils.checkIfBackButtonIsPressed = function (cacheStorage) {
        var redirectCache = cacheStorage.getItem(Constants_1.TemporaryCacheKeys.REDIRECT_REQUEST);
        // if redirect request is set and there is no hash
        if (redirectCache && !UrlUtils_1.UrlUtils.urlContainsHash(window.location.hash)) {
            var splitCache = redirectCache.split(Constants_1.Constants.resourceDelimiter);
            var state = splitCache.length > 1 ? splitCache[splitCache.length - 1] : null;
            cacheStorage.resetTempCacheItems(state);
        }
    };
    /**
     * @hidden
     * Interval in milliseconds that we poll a window
     * @ignore
     */
    WindowUtils.POLLING_INTERVAL_MS = 50;
    return WindowUtils;
}());
exports.WindowUtils = WindowUtils;


/***/ }),
/* 21 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = __webpack_require__(0);
/**
 * @hidden
 */
var Authority_1 = __webpack_require__(22);
var StringUtils_1 = __webpack_require__(3);
var ClientConfigurationError_1 = __webpack_require__(5);
var Constants_1 = __webpack_require__(1);
var UrlUtils_1 = __webpack_require__(4);
var AuthorityFactory = /** @class */ (function () {
    function AuthorityFactory() {
    }
    AuthorityFactory.saveMetadataFromNetwork = function (authorityInstance, telemetryManager, correlationId) {
        return tslib_1.__awaiter(this, void 0, Promise, function () {
            var metadata;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, authorityInstance.resolveEndpointsAsync(telemetryManager, correlationId)];
                    case 1:
                        metadata = _a.sent();
                        this.metadataMap.set(authorityInstance.CanonicalAuthority, metadata);
                        return [2 /*return*/, metadata];
                }
            });
        });
    };
    AuthorityFactory.getMetadata = function (authorityUrl) {
        return this.metadataMap.get(authorityUrl);
    };
    AuthorityFactory.saveMetadataFromConfig = function (authorityUrl, authorityMetadataJson) {
        try {
            if (authorityMetadataJson) {
                var parsedMetadata = JSON.parse(authorityMetadataJson);
                if (!parsedMetadata.authorization_endpoint || !parsedMetadata.end_session_endpoint || !parsedMetadata.issuer) {
                    throw ClientConfigurationError_1.ClientConfigurationError.createInvalidAuthorityMetadataError();
                }
                this.metadataMap.set(authorityUrl, {
                    AuthorizationEndpoint: parsedMetadata.authorization_endpoint,
                    EndSessionEndpoint: parsedMetadata.end_session_endpoint,
                    Issuer: parsedMetadata.issuer
                });
            }
        }
        catch (e) {
            throw ClientConfigurationError_1.ClientConfigurationError.createInvalidAuthorityMetadataError();
        }
    };
    /**
     * Create an authority object of the correct type based on the url
     * Performs basic authority validation - checks to see if the authority is of a valid type (eg aad, b2c)
     */
    AuthorityFactory.CreateInstance = function (authorityUrl, validateAuthority, authorityMetadata) {
        if (StringUtils_1.StringUtils.isEmpty(authorityUrl)) {
            return null;
        }
        if (authorityMetadata) {
            // todo: log statements
            this.saveMetadataFromConfig(authorityUrl, authorityMetadata);
        }
        return new Authority_1.Authority(authorityUrl, validateAuthority, this.metadataMap.get(authorityUrl));
    };
    AuthorityFactory.isAdfs = function (authorityUrl) {
        var components = UrlUtils_1.UrlUtils.GetUrlComponents(authorityUrl);
        var pathSegments = components.PathSegments;
        if (pathSegments.length && pathSegments[0].toLowerCase() === Constants_1.Constants.ADFS) {
            return true;
        }
        return false;
    };
    AuthorityFactory.metadataMap = new Map();
    return AuthorityFactory;
}());
exports.AuthorityFactory = AuthorityFactory;


/***/ }),
/* 22 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = __webpack_require__(0);
var ClientConfigurationError_1 = __webpack_require__(5);
var XHRClient_1 = __webpack_require__(23);
var UrlUtils_1 = __webpack_require__(4);
var TrustedAuthority_1 = __webpack_require__(24);
var Constants_1 = __webpack_require__(1);
/**
 * @hidden
 */
var AuthorityType;
(function (AuthorityType) {
    AuthorityType[AuthorityType["Default"] = 0] = "Default";
    AuthorityType[AuthorityType["Adfs"] = 1] = "Adfs";
})(AuthorityType = exports.AuthorityType || (exports.AuthorityType = {}));
/**
 * @hidden
 */
var Authority = /** @class */ (function () {
    function Authority(authority, validateAuthority, authorityMetadata) {
        this.IsValidationEnabled = validateAuthority;
        this.CanonicalAuthority = authority;
        this.validateAsUri();
        this.tenantDiscoveryResponse = authorityMetadata;
    }
    Object.defineProperty(Authority.prototype, "Tenant", {
        get: function () {
            return this.CanonicalAuthorityUrlComponents.PathSegments[0];
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Authority.prototype, "AuthorizationEndpoint", {
        get: function () {
            this.validateResolved();
            return this.tenantDiscoveryResponse.AuthorizationEndpoint.replace(/{tenant}|{tenantid}/g, this.Tenant);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Authority.prototype, "EndSessionEndpoint", {
        get: function () {
            this.validateResolved();
            return this.tenantDiscoveryResponse.EndSessionEndpoint.replace(/{tenant}|{tenantid}/g, this.Tenant);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Authority.prototype, "SelfSignedJwtAudience", {
        get: function () {
            this.validateResolved();
            return this.tenantDiscoveryResponse.Issuer.replace(/{tenant}|{tenantid}/g, this.Tenant);
        },
        enumerable: true,
        configurable: true
    });
    Authority.prototype.validateResolved = function () {
        if (!this.hasCachedMetadata()) {
            throw "Please call ResolveEndpointsAsync first";
        }
    };
    Object.defineProperty(Authority.prototype, "CanonicalAuthority", {
        /**
         * A URL that is the authority set by the developer
         */
        get: function () {
            return this.canonicalAuthority;
        },
        set: function (url) {
            this.canonicalAuthority = UrlUtils_1.UrlUtils.CanonicalizeUri(url);
            this.canonicalAuthorityUrlComponents = null;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Authority.prototype, "CanonicalAuthorityUrlComponents", {
        get: function () {
            if (!this.canonicalAuthorityUrlComponents) {
                this.canonicalAuthorityUrlComponents = UrlUtils_1.UrlUtils.GetUrlComponents(this.CanonicalAuthority);
            }
            return this.canonicalAuthorityUrlComponents;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Authority.prototype, "DefaultOpenIdConfigurationEndpoint", {
        // http://openid.net/specs/openid-connect-discovery-1_0.html#ProviderMetadata
        get: function () {
            return this.CanonicalAuthority + "v2.0/.well-known/openid-configuration";
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Given a string, validate that it is of the form https://domain/path
     */
    Authority.prototype.validateAsUri = function () {
        var components;
        try {
            components = this.CanonicalAuthorityUrlComponents;
        }
        catch (e) {
            throw ClientConfigurationError_1.ClientConfigurationErrorMessage.invalidAuthorityType;
        }
        if (!components.Protocol || components.Protocol.toLowerCase() !== "https:") {
            throw ClientConfigurationError_1.ClientConfigurationErrorMessage.authorityUriInsecure;
        }
        if (!components.PathSegments || components.PathSegments.length < 1) {
            throw ClientConfigurationError_1.ClientConfigurationErrorMessage.authorityUriInvalidPath;
        }
    };
    /**
     * Calls the OIDC endpoint and returns the response
     */
    Authority.prototype.DiscoverEndpoints = function (openIdConfigurationEndpoint, telemetryManager, correlationId) {
        var client = new XHRClient_1.XhrClient();
        var httpMethod = Constants_1.NetworkRequestType.GET;
        var httpEvent = telemetryManager.createAndStartHttpEvent(correlationId, httpMethod, openIdConfigurationEndpoint, "openIdConfigurationEndpoint");
        return client.sendRequestAsync(openIdConfigurationEndpoint, httpMethod, /* enableCaching: */ true)
            .then(function (response) {
            httpEvent.httpResponseStatus = response.statusCode;
            telemetryManager.stopEvent(httpEvent);
            return {
                AuthorizationEndpoint: response.body.authorization_endpoint,
                EndSessionEndpoint: response.body.end_session_endpoint,
                Issuer: response.body.issuer
            };
        })
            .catch(function (err) {
            httpEvent.serverErrorCode = err;
            telemetryManager.stopEvent(httpEvent);
            throw err;
        });
    };
    /**
     * Returns a promise.
     * Checks to see if the authority is in the cache
     * Discover endpoints via openid-configuration
     * If successful, caches the endpoint for later use in OIDC
     */
    Authority.prototype.resolveEndpointsAsync = function (telemetryManager, correlationId) {
        return tslib_1.__awaiter(this, void 0, Promise, function () {
            var host, openIdConfigurationEndpointResponse, _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!this.IsValidationEnabled) return [3 /*break*/, 3];
                        host = this.canonicalAuthorityUrlComponents.HostNameAndPort;
                        if (!(TrustedAuthority_1.TrustedAuthority.getTrustedHostList().length === 0)) return [3 /*break*/, 2];
                        return [4 /*yield*/, TrustedAuthority_1.TrustedAuthority.setTrustedAuthoritiesFromNetwork(this.canonicalAuthority, telemetryManager, correlationId)];
                    case 1:
                        _b.sent();
                        _b.label = 2;
                    case 2:
                        if (!TrustedAuthority_1.TrustedAuthority.IsInTrustedHostList(host)) {
                            throw ClientConfigurationError_1.ClientConfigurationError.createUntrustedAuthorityError(host);
                        }
                        _b.label = 3;
                    case 3:
                        openIdConfigurationEndpointResponse = this.GetOpenIdConfigurationEndpoint();
                        _a = this;
                        return [4 /*yield*/, this.DiscoverEndpoints(openIdConfigurationEndpointResponse, telemetryManager, correlationId)];
                    case 4:
                        _a.tenantDiscoveryResponse = _b.sent();
                        return [2 /*return*/, this.tenantDiscoveryResponse];
                }
            });
        });
    };
    /**
     * Checks if there is a cached tenant discovery response with required fields.
     */
    Authority.prototype.hasCachedMetadata = function () {
        return !!(this.tenantDiscoveryResponse &&
            this.tenantDiscoveryResponse.AuthorizationEndpoint &&
            this.tenantDiscoveryResponse.EndSessionEndpoint &&
            this.tenantDiscoveryResponse.Issuer);
    };
    /**
     * Returns a promise which resolves to the OIDC endpoint
     * Only responds with the endpoint
     */
    Authority.prototype.GetOpenIdConfigurationEndpoint = function () {
        return this.DefaultOpenIdConfigurationEndpoint;
    };
    return Authority;
}());
exports.Authority = Authority;


/***/ }),
/* 23 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
var Constants_1 = __webpack_require__(1);
/**
 * XHR client for JSON endpoints
 * https://www.npmjs.com/package/async-promise
 * @hidden
 */
var XhrClient = /** @class */ (function () {
    function XhrClient() {
    }
    XhrClient.prototype.sendRequestAsync = function (url, method, enableCaching) {
        var _this = this;
        return new Promise(function (resolve, reject) {
            var xhr = new XMLHttpRequest();
            xhr.open(method, url, /* async: */ true);
            if (enableCaching) {
                /*
                 * TODO: (shivb) ensure that this can be cached
                 * xhr.setRequestHeader("Cache-Control", "Public");
                 */
            }
            xhr.onload = function (ev) {
                if (xhr.status < 200 || xhr.status >= 300) {
                    reject(_this.handleError(xhr.responseText));
                }
                var jsonResponse;
                try {
                    jsonResponse = JSON.parse(xhr.responseText);
                }
                catch (e) {
                    reject(_this.handleError(xhr.responseText));
                }
                var response = {
                    statusCode: xhr.status,
                    body: jsonResponse
                };
                resolve(response);
            };
            xhr.onerror = function (ev) {
                reject(xhr.status);
            };
            if (method === Constants_1.NetworkRequestType.GET) {
                xhr.send();
            }
            else {
                throw "not implemented";
            }
        });
    };
    XhrClient.prototype.handleError = function (responseText) {
        var jsonResponse;
        try {
            jsonResponse = JSON.parse(responseText);
            if (jsonResponse.error) {
                return jsonResponse.error;
            }
            else {
                throw responseText;
            }
        }
        catch (e) {
            return responseText;
        }
    };
    return XhrClient;
}());
exports.XhrClient = XhrClient;


/***/ }),
/* 24 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = __webpack_require__(0);
var XHRClient_1 = __webpack_require__(23);
var Constants_1 = __webpack_require__(1);
var UrlUtils_1 = __webpack_require__(4);
var TrustedAuthority = /** @class */ (function () {
    function TrustedAuthority() {
    }
    /**
     *
     * @param validateAuthority
     * @param knownAuthorities
     */
    TrustedAuthority.setTrustedAuthoritiesFromConfig = function (validateAuthority, knownAuthorities) {
        if (validateAuthority && !this.getTrustedHostList().length) {
            knownAuthorities.forEach(function (authority) {
                TrustedAuthority.TrustedHostList.push(authority.toLowerCase());
            });
        }
    };
    /**
     *
     * @param telemetryManager
     * @param correlationId
     */
    TrustedAuthority.getAliases = function (authorityToVerify, telemetryManager, correlationId) {
        return tslib_1.__awaiter(this, void 0, Promise, function () {
            var client, httpMethod, instanceDiscoveryEndpoint, httpEvent;
            return tslib_1.__generator(this, function (_a) {
                client = new XHRClient_1.XhrClient();
                httpMethod = Constants_1.NetworkRequestType.GET;
                instanceDiscoveryEndpoint = "" + Constants_1.AAD_INSTANCE_DISCOVERY_ENDPOINT + authorityToVerify + "oauth2/v2.0/authorize";
                httpEvent = telemetryManager.createAndStartHttpEvent(correlationId, httpMethod, instanceDiscoveryEndpoint, "getAliases");
                return [2 /*return*/, client.sendRequestAsync(instanceDiscoveryEndpoint, httpMethod, true)
                        .then(function (response) {
                        httpEvent.httpResponseStatus = response.statusCode;
                        telemetryManager.stopEvent(httpEvent);
                        return response.body.metadata;
                    })
                        .catch(function (err) {
                        httpEvent.serverErrorCode = err;
                        telemetryManager.stopEvent(httpEvent);
                        throw err;
                    })];
            });
        });
    };
    /**
     *
     * @param telemetryManager
     * @param correlationId
     */
    TrustedAuthority.setTrustedAuthoritiesFromNetwork = function (authorityToVerify, telemetryManager, correlationId) {
        return tslib_1.__awaiter(this, void 0, Promise, function () {
            var metadata, host;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.getAliases(authorityToVerify, telemetryManager, correlationId)];
                    case 1:
                        metadata = _a.sent();
                        metadata.forEach(function (entry) {
                            var authorities = entry.aliases;
                            authorities.forEach(function (authority) {
                                TrustedAuthority.TrustedHostList.push(authority.toLowerCase());
                            });
                        });
                        host = UrlUtils_1.UrlUtils.GetUrlComponents(authorityToVerify).HostNameAndPort;
                        if (TrustedAuthority.getTrustedHostList().length && !TrustedAuthority.IsInTrustedHostList(host)) {
                            // Custom Domain scenario, host is trusted because Instance Discovery call succeeded
                            TrustedAuthority.TrustedHostList.push(host.toLowerCase());
                        }
                        return [2 /*return*/];
                }
            });
        });
    };
    TrustedAuthority.getTrustedHostList = function () {
        return this.TrustedHostList;
    };
    /**
     * Checks to see if the host is in a list of trusted hosts
     * @param host
     */
    TrustedAuthority.IsInTrustedHostList = function (host) {
        return this.TrustedHostList.indexOf(host.toLowerCase()) > -1;
    };
    TrustedAuthority.TrustedHostList = [];
    return TrustedAuthority;
}());
exports.TrustedAuthority = TrustedAuthority;


/***/ }),
/* 25 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = __webpack_require__(0);
var Logger_1 = __webpack_require__(12);
var UrlUtils_1 = __webpack_require__(4);
/**
 * Defaults for the Configuration Options
 */
var FRAME_TIMEOUT = 6000;
var OFFSET = 300;
var NAVIGATE_FRAME_WAIT = 500;
var DEFAULT_AUTH_OPTIONS = {
    clientId: "",
    authority: null,
    validateAuthority: true,
    authorityMetadata: "",
    knownAuthorities: [],
    redirectUri: function () { return UrlUtils_1.UrlUtils.getCurrentUrl(); },
    postLogoutRedirectUri: function () { return UrlUtils_1.UrlUtils.getCurrentUrl(); },
    navigateToLoginRequestUrl: true
};
var DEFAULT_CACHE_OPTIONS = {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false
};
var DEFAULT_SYSTEM_OPTIONS = {
    logger: new Logger_1.Logger(null),
    loadFrameTimeout: FRAME_TIMEOUT,
    tokenRenewalOffsetSeconds: OFFSET,
    navigateFrameWait: NAVIGATE_FRAME_WAIT
};
var DEFAULT_FRAMEWORK_OPTIONS = {
    isAngular: false,
    unprotectedResources: new Array(),
    protectedResourceMap: new Map()
};
/**
 * MSAL function that sets the default options when not explicitly configured from app developer
 *
 * @param TAuthOptions
 * @param TCacheOptions
 * @param TSystemOptions
 * @param TFrameworkOptions
 * @param TAuthorityDataOptions
 *
 * @returns TConfiguration object
 */
function buildConfiguration(_a) {
    var auth = _a.auth, _b = _a.cache, cache = _b === void 0 ? {} : _b, _c = _a.system, system = _c === void 0 ? {} : _c, _d = _a.framework, framework = _d === void 0 ? {} : _d;
    var overlayedConfig = {
        auth: tslib_1.__assign({}, DEFAULT_AUTH_OPTIONS, auth),
        cache: tslib_1.__assign({}, DEFAULT_CACHE_OPTIONS, cache),
        system: tslib_1.__assign({}, DEFAULT_SYSTEM_OPTIONS, system),
        framework: tslib_1.__assign({}, DEFAULT_FRAMEWORK_OPTIONS, framework)
    };
    return overlayedConfig;
}
exports.buildConfiguration = buildConfiguration;


/***/ }),
/* 26 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = __webpack_require__(0);
var ServerError_1 = __webpack_require__(13);
exports.InteractionRequiredAuthErrorMessage = {
    interactionRequired: {
        code: "interaction_required"
    },
    consentRequired: {
        code: "consent_required"
    },
    loginRequired: {
        code: "login_required"
    },
};
/**
 * Error thrown when the user is required to perform an interactive token request.
 */
var InteractionRequiredAuthError = /** @class */ (function (_super) {
    tslib_1.__extends(InteractionRequiredAuthError, _super);
    function InteractionRequiredAuthError(errorCode, errorMessage) {
        var _this = _super.call(this, errorCode, errorMessage) || this;
        _this.name = "InteractionRequiredAuthError";
        Object.setPrototypeOf(_this, InteractionRequiredAuthError.prototype);
        return _this;
    }
    InteractionRequiredAuthError.isInteractionRequiredError = function (errorString) {
        var interactionRequiredCodes = [
            exports.InteractionRequiredAuthErrorMessage.interactionRequired.code,
            exports.InteractionRequiredAuthErrorMessage.consentRequired.code,
            exports.InteractionRequiredAuthErrorMessage.loginRequired.code
        ];
        return errorString && interactionRequiredCodes.indexOf(errorString) > -1;
    };
    InteractionRequiredAuthError.createLoginRequiredAuthError = function (errorDesc) {
        return new InteractionRequiredAuthError(exports.InteractionRequiredAuthErrorMessage.loginRequired.code, errorDesc);
    };
    InteractionRequiredAuthError.createInteractionRequiredAuthError = function (errorDesc) {
        return new InteractionRequiredAuthError(exports.InteractionRequiredAuthErrorMessage.interactionRequired.code, errorDesc);
    };
    InteractionRequiredAuthError.createConsentRequiredAuthError = function (errorDesc) {
        return new InteractionRequiredAuthError(exports.InteractionRequiredAuthErrorMessage.consentRequired.code, errorDesc);
    };
    return InteractionRequiredAuthError;
}(ServerError_1.ServerError));
exports.InteractionRequiredAuthError = InteractionRequiredAuthError;


/***/ }),
/* 27 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
function buildResponseStateOnly(state) {
    return {
        uniqueId: "",
        tenantId: "",
        tokenType: "",
        idToken: null,
        idTokenClaims: null,
        accessToken: "",
        scopes: null,
        expiresOn: null,
        account: null,
        accountState: state,
        fromCache: false
    };
}
exports.buildResponseStateOnly = buildResponseStateOnly;


/***/ }),
/* 28 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var _a;
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = __webpack_require__(0);
var TelemetryEvent_1 = tslib_1.__importDefault(__webpack_require__(14));
var TelemetryConstants_1 = __webpack_require__(8);
var TelemetryUtils_1 = __webpack_require__(10);
exports.EVENT_KEYS = {
    AUTHORITY: TelemetryUtils_1.prependEventNamePrefix("authority"),
    AUTHORITY_TYPE: TelemetryUtils_1.prependEventNamePrefix("authority_type"),
    PROMPT: TelemetryUtils_1.prependEventNamePrefix("ui_behavior"),
    TENANT_ID: TelemetryUtils_1.prependEventNamePrefix("tenant_id"),
    USER_ID: TelemetryUtils_1.prependEventNamePrefix("user_id"),
    WAS_SUCESSFUL: TelemetryUtils_1.prependEventNamePrefix("was_successful"),
    API_ERROR_CODE: TelemetryUtils_1.prependEventNamePrefix("api_error_code"),
    LOGIN_HINT: TelemetryUtils_1.prependEventNamePrefix("login_hint")
};
var API_CODE;
(function (API_CODE) {
    API_CODE[API_CODE["AcquireTokenRedirect"] = 2001] = "AcquireTokenRedirect";
    API_CODE[API_CODE["AcquireTokenSilent"] = 2002] = "AcquireTokenSilent";
    API_CODE[API_CODE["AcquireTokenPopup"] = 2003] = "AcquireTokenPopup";
    API_CODE[API_CODE["LoginRedirect"] = 2004] = "LoginRedirect";
    API_CODE[API_CODE["LoginPopup"] = 2005] = "LoginPopup";
    API_CODE[API_CODE["Logout"] = 2006] = "Logout";
})(API_CODE = exports.API_CODE || (exports.API_CODE = {}));
var API_EVENT_IDENTIFIER;
(function (API_EVENT_IDENTIFIER) {
    API_EVENT_IDENTIFIER["AcquireTokenRedirect"] = "AcquireTokenRedirect";
    API_EVENT_IDENTIFIER["AcquireTokenSilent"] = "AcquireTokenSilent";
    API_EVENT_IDENTIFIER["AcquireTokenPopup"] = "AcquireTokenPopup";
    API_EVENT_IDENTIFIER["LoginRedirect"] = "LoginRedirect";
    API_EVENT_IDENTIFIER["LoginPopup"] = "LoginPopup";
    API_EVENT_IDENTIFIER["Logout"] = "Logout";
})(API_EVENT_IDENTIFIER = exports.API_EVENT_IDENTIFIER || (exports.API_EVENT_IDENTIFIER = {}));
var mapEventIdentiferToCode = (_a = {},
    _a[API_EVENT_IDENTIFIER.AcquireTokenSilent] = API_CODE.AcquireTokenSilent,
    _a[API_EVENT_IDENTIFIER.AcquireTokenPopup] = API_CODE.AcquireTokenPopup,
    _a[API_EVENT_IDENTIFIER.AcquireTokenRedirect] = API_CODE.AcquireTokenRedirect,
    _a[API_EVENT_IDENTIFIER.LoginPopup] = API_CODE.LoginPopup,
    _a[API_EVENT_IDENTIFIER.LoginRedirect] = API_CODE.LoginRedirect,
    _a[API_EVENT_IDENTIFIER.Logout] = API_CODE.Logout,
    _a);
var ApiEvent = /** @class */ (function (_super) {
    tslib_1.__extends(ApiEvent, _super);
    function ApiEvent(correlationId, piiEnabled, apiEventIdentifier) {
        var _this = _super.call(this, TelemetryUtils_1.prependEventNamePrefix("api_event"), correlationId, apiEventIdentifier) || this;
        if (apiEventIdentifier) {
            _this.apiCode = mapEventIdentiferToCode[apiEventIdentifier];
            _this.apiEventIdentifier = apiEventIdentifier;
        }
        _this.piiEnabled = piiEnabled;
        return _this;
    }
    Object.defineProperty(ApiEvent.prototype, "apiEventIdentifier", {
        set: function (apiEventIdentifier) {
            this.event[TelemetryConstants_1.TELEMETRY_BLOB_EVENT_NAMES.ApiTelemIdConstStrKey] = apiEventIdentifier;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ApiEvent.prototype, "apiCode", {
        set: function (apiCode) {
            this.event[TelemetryConstants_1.TELEMETRY_BLOB_EVENT_NAMES.ApiIdConstStrKey] = apiCode;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ApiEvent.prototype, "authority", {
        set: function (uri) {
            this.event[exports.EVENT_KEYS.AUTHORITY] = TelemetryUtils_1.scrubTenantFromUri(uri).toLowerCase();
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ApiEvent.prototype, "apiErrorCode", {
        set: function (errorCode) {
            this.event[exports.EVENT_KEYS.API_ERROR_CODE] = errorCode;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ApiEvent.prototype, "tenantId", {
        set: function (tenantId) {
            this.event[exports.EVENT_KEYS.TENANT_ID] = this.piiEnabled && tenantId ?
                TelemetryUtils_1.hashPersonalIdentifier(tenantId)
                : null;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ApiEvent.prototype, "accountId", {
        set: function (accountId) {
            this.event[exports.EVENT_KEYS.USER_ID] = this.piiEnabled && accountId ?
                TelemetryUtils_1.hashPersonalIdentifier(accountId)
                : null;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ApiEvent.prototype, "wasSuccessful", {
        get: function () {
            return this.event[exports.EVENT_KEYS.WAS_SUCESSFUL] === true;
        },
        set: function (wasSuccessful) {
            this.event[exports.EVENT_KEYS.WAS_SUCESSFUL] = wasSuccessful;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ApiEvent.prototype, "loginHint", {
        set: function (loginHint) {
            this.event[exports.EVENT_KEYS.LOGIN_HINT] = this.piiEnabled && loginHint ?
                TelemetryUtils_1.hashPersonalIdentifier(loginHint)
                : null;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ApiEvent.prototype, "authorityType", {
        set: function (authorityType) {
            this.event[exports.EVENT_KEYS.AUTHORITY_TYPE] = authorityType.toLowerCase();
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ApiEvent.prototype, "promptType", {
        set: function (promptType) {
            this.event[exports.EVENT_KEYS.PROMPT] = promptType.toLowerCase();
        },
        enumerable: true,
        configurable: true
    });
    return ApiEvent;
}(TelemetryEvent_1.default));
exports.default = ApiEvent;


/***/ }),
/* 29 */
/***/ (function(module, exports, __webpack_require__) {

module.exports = __webpack_require__(30);


/***/ }),
/* 30 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var UserAgentApplication_1 = __webpack_require__(15);
exports.UserAgentApplication = UserAgentApplication_1.UserAgentApplication;
exports.authResponseCallback = UserAgentApplication_1.authResponseCallback;
exports.errorReceivedCallback = UserAgentApplication_1.errorReceivedCallback;
exports.tokenReceivedCallback = UserAgentApplication_1.tokenReceivedCallback;
var Logger_1 = __webpack_require__(12);
exports.Logger = Logger_1.Logger;
var Logger_2 = __webpack_require__(12);
exports.LogLevel = Logger_2.LogLevel;
var Account_1 = __webpack_require__(19);
exports.Account = Account_1.Account;
var Constants_1 = __webpack_require__(1);
exports.Constants = Constants_1.Constants;
exports.ServerHashParamKeys = Constants_1.ServerHashParamKeys;
var Authority_1 = __webpack_require__(22);
exports.Authority = Authority_1.Authority;
var UserAgentApplication_2 = __webpack_require__(15);
exports.CacheResult = UserAgentApplication_2.CacheResult;
var Configuration_1 = __webpack_require__(25);
exports.CacheLocation = Configuration_1.CacheLocation;
exports.Configuration = Configuration_1.Configuration;
var AuthenticationParameters_1 = __webpack_require__(42);
exports.AuthenticationParameters = AuthenticationParameters_1.AuthenticationParameters;
var AuthResponse_1 = __webpack_require__(27);
exports.AuthResponse = AuthResponse_1.AuthResponse;
var CryptoUtils_1 = __webpack_require__(2);
exports.CryptoUtils = CryptoUtils_1.CryptoUtils;
var UrlUtils_1 = __webpack_require__(4);
exports.UrlUtils = UrlUtils_1.UrlUtils;
var WindowUtils_1 = __webpack_require__(20);
exports.WindowUtils = WindowUtils_1.WindowUtils;
// Errors
var AuthError_1 = __webpack_require__(7);
exports.AuthError = AuthError_1.AuthError;
var ClientAuthError_1 = __webpack_require__(6);
exports.ClientAuthError = ClientAuthError_1.ClientAuthError;
var ServerError_1 = __webpack_require__(13);
exports.ServerError = ServerError_1.ServerError;
var ClientConfigurationError_1 = __webpack_require__(5);
exports.ClientConfigurationError = ClientConfigurationError_1.ClientConfigurationError;
var InteractionRequiredAuthError_1 = __webpack_require__(26);
exports.InteractionRequiredAuthError = InteractionRequiredAuthError_1.InteractionRequiredAuthError;


/***/ }),
/* 31 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
var CryptoUtils_1 = __webpack_require__(2);
var UrlUtils_1 = __webpack_require__(4);
/**
 * @hidden
 */
var AccessTokenKey = /** @class */ (function () {
    function AccessTokenKey(authority, clientId, scopes, uid, utid) {
        this.authority = UrlUtils_1.UrlUtils.CanonicalizeUri(authority);
        this.clientId = clientId;
        this.scopes = scopes;
        this.homeAccountIdentifier = CryptoUtils_1.CryptoUtils.base64Encode(uid) + "." + CryptoUtils_1.CryptoUtils.base64Encode(utid);
    }
    return AccessTokenKey;
}());
exports.AccessTokenKey = AccessTokenKey;


/***/ }),
/* 32 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
/**
 * @hidden
 */
var AccessTokenValue = /** @class */ (function () {
    function AccessTokenValue(accessToken, idToken, expiresIn, homeAccountIdentifier) {
        this.accessToken = accessToken;
        this.idToken = idToken;
        this.expiresIn = expiresIn;
        this.homeAccountIdentifier = homeAccountIdentifier;
    }
    return AccessTokenValue;
}());
exports.AccessTokenValue = AccessTokenValue;


/***/ }),
/* 33 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
var CryptoUtils_1 = __webpack_require__(2);
var ClientAuthError_1 = __webpack_require__(6);
var StringUtils_1 = __webpack_require__(3);
/**
 * @hidden
 */
var ClientInfo = /** @class */ (function () {
    function ClientInfo(rawClientInfo) {
        if (!rawClientInfo || StringUtils_1.StringUtils.isEmpty(rawClientInfo)) {
            this.uid = "";
            this.utid = "";
            return;
        }
        try {
            var decodedClientInfo = CryptoUtils_1.CryptoUtils.base64Decode(rawClientInfo);
            var clientInfo = JSON.parse(decodedClientInfo);
            if (clientInfo) {
                if (clientInfo.hasOwnProperty("uid")) {
                    this.uid = clientInfo.uid;
                }
                if (clientInfo.hasOwnProperty("utid")) {
                    this.utid = clientInfo.utid;
                }
            }
        }
        catch (e) {
            throw ClientAuthError_1.ClientAuthError.createClientInfoDecodingError(e);
        }
    }
    Object.defineProperty(ClientInfo.prototype, "uid", {
        get: function () {
            return this._uid ? this._uid : "";
        },
        set: function (uid) {
            this._uid = uid;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ClientInfo.prototype, "utid", {
        get: function () {
            return this._utid ? this._utid : "";
        },
        set: function (utid) {
            this._utid = utid;
        },
        enumerable: true,
        configurable: true
    });
    return ClientInfo;
}());
exports.ClientInfo = ClientInfo;


/***/ }),
/* 34 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
var ClientAuthError_1 = __webpack_require__(6);
var TokenUtils_1 = __webpack_require__(17);
var StringUtils_1 = __webpack_require__(3);
/**
 * @hidden
 */
var IdToken = /** @class */ (function () {
    /* tslint:disable:no-string-literal */
    function IdToken(rawIdToken) {
        if (StringUtils_1.StringUtils.isEmpty(rawIdToken)) {
            throw ClientAuthError_1.ClientAuthError.createIdTokenNullOrEmptyError(rawIdToken);
        }
        try {
            this.rawIdToken = rawIdToken;
            this.claims = TokenUtils_1.TokenUtils.extractIdToken(rawIdToken);
            if (this.claims) {
                if (this.claims.hasOwnProperty("iss")) {
                    this.issuer = this.claims["iss"];
                }
                if (this.claims.hasOwnProperty("oid")) {
                    this.objectId = this.claims["oid"];
                }
                if (this.claims.hasOwnProperty("sub")) {
                    this.subject = this.claims["sub"];
                }
                if (this.claims.hasOwnProperty("tid")) {
                    this.tenantId = this.claims["tid"];
                }
                if (this.claims.hasOwnProperty("ver")) {
                    this.version = this.claims["ver"];
                }
                if (this.claims.hasOwnProperty("preferred_username")) {
                    this.preferredName = this.claims["preferred_username"];
                }
                if (this.claims.hasOwnProperty("name")) {
                    this.name = this.claims["name"];
                }
                if (this.claims.hasOwnProperty("nonce")) {
                    this.nonce = this.claims["nonce"];
                }
                if (this.claims.hasOwnProperty("exp")) {
                    this.expiration = this.claims["exp"];
                }
                if (this.claims.hasOwnProperty("home_oid")) {
                    this.homeObjectId = this.claims["home_oid"];
                }
                if (this.claims.hasOwnProperty("sid")) {
                    this.sid = this.claims["sid"];
                }
                if (this.claims.hasOwnProperty("cloud_instance_host_name")) {
                    this.cloudInstance = this.claims["cloud_instance_host_name"];
                }
                /* tslint:enable:no-string-literal */
            }
        }
        catch (e) {
            /*
             * TODO: This error here won't really every be thrown, since extractIdToken() returns null if the decodeJwt() fails.
             * Need to add better error handling here to account for being unable to decode jwts.
             */
            throw ClientAuthError_1.ClientAuthError.createIdTokenParsingError(e);
        }
    }
    return IdToken;
}());
exports.IdToken = IdToken;


/***/ }),
/* 35 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = __webpack_require__(0);
var Constants_1 = __webpack_require__(1);
var AccessTokenCacheItem_1 = __webpack_require__(36);
var BrowserStorage_1 = __webpack_require__(37);
var ClientAuthError_1 = __webpack_require__(6);
var RequestUtils_1 = __webpack_require__(18);
/**
 * @hidden
 */
var AuthCache = /** @class */ (function (_super) {
    tslib_1.__extends(AuthCache, _super);
    function AuthCache(clientId, cacheLocation, storeAuthStateInCookie) {
        var _this = _super.call(this, cacheLocation) || this;
        _this.clientId = clientId;
        // This is hardcoded to true for now. We may make this configurable in the future
        _this.rollbackEnabled = true;
        _this.migrateCacheEntries(storeAuthStateInCookie);
        return _this;
    }
    /**
     * Support roll back to old cache schema until the next major release: true by default now
     * @param storeAuthStateInCookie
     */
    AuthCache.prototype.migrateCacheEntries = function (storeAuthStateInCookie) {
        var _this = this;
        var idTokenKey = Constants_1.Constants.cachePrefix + "." + Constants_1.PersistentCacheKeys.IDTOKEN;
        var clientInfoKey = Constants_1.Constants.cachePrefix + "." + Constants_1.PersistentCacheKeys.CLIENT_INFO;
        var errorKey = Constants_1.Constants.cachePrefix + "." + Constants_1.ErrorCacheKeys.ERROR;
        var errorDescKey = Constants_1.Constants.cachePrefix + "." + Constants_1.ErrorCacheKeys.ERROR_DESC;
        var idTokenValue = _super.prototype.getItem.call(this, idTokenKey);
        var clientInfoValue = _super.prototype.getItem.call(this, clientInfoKey);
        var errorValue = _super.prototype.getItem.call(this, errorKey);
        var errorDescValue = _super.prototype.getItem.call(this, errorDescKey);
        var values = [idTokenValue, clientInfoValue, errorValue, errorDescValue];
        var keysToMigrate = [Constants_1.PersistentCacheKeys.IDTOKEN, Constants_1.PersistentCacheKeys.CLIENT_INFO, Constants_1.ErrorCacheKeys.ERROR, Constants_1.ErrorCacheKeys.ERROR_DESC];
        keysToMigrate.forEach(function (cacheKey, index) { return _this.duplicateCacheEntry(cacheKey, values[index], storeAuthStateInCookie); });
    };
    /**
     * Utility function to help with roll back keys
     * @param newKey
     * @param value
     * @param storeAuthStateInCookie
     */
    AuthCache.prototype.duplicateCacheEntry = function (newKey, value, storeAuthStateInCookie) {
        if (value) {
            this.setItem(newKey, value, storeAuthStateInCookie);
        }
    };
    /**
     * Prepend msal.<client-id> to each key; Skip for any JSON object as Key (defined schemas do not need the key appended: AccessToken Keys or the upcoming schema)
     * @param key
     * @param addInstanceId
     */
    AuthCache.prototype.generateCacheKey = function (key, addInstanceId) {
        try {
            // Defined schemas do not need the key appended
            JSON.parse(key);
            return key;
        }
        catch (e) {
            if (key.indexOf("" + Constants_1.Constants.cachePrefix) === 0 || key.indexOf(Constants_1.Constants.adalIdToken) === 0) {
                return key;
            }
            return addInstanceId ? Constants_1.Constants.cachePrefix + "." + this.clientId + "." + key : Constants_1.Constants.cachePrefix + "." + key;
        }
    };
    /**
     * add value to storage
     * @param key
     * @param value
     * @param enableCookieStorage
     */
    AuthCache.prototype.setItem = function (key, value, enableCookieStorage) {
        _super.prototype.setItem.call(this, this.generateCacheKey(key, true), value, enableCookieStorage);
        // Values stored in cookies will have rollback disabled to minimize cookie length
        if (this.rollbackEnabled && !enableCookieStorage) {
            _super.prototype.setItem.call(this, this.generateCacheKey(key, false), value, enableCookieStorage);
        }
    };
    /**
     * get one item by key from storage
     * @param key
     * @param enableCookieStorage
     */
    AuthCache.prototype.getItem = function (key, enableCookieStorage) {
        return _super.prototype.getItem.call(this, this.generateCacheKey(key, true), enableCookieStorage);
    };
    /**
     * remove value from storage
     * @param key
     */
    AuthCache.prototype.removeItem = function (key) {
        _super.prototype.removeItem.call(this, this.generateCacheKey(key, true));
        if (this.rollbackEnabled) {
            _super.prototype.removeItem.call(this, this.generateCacheKey(key, false));
        }
    };
    /**
     * Reset the cache items
     */
    AuthCache.prototype.resetCacheItems = function () {
        var storage = window[this.cacheLocation];
        var key;
        for (key in storage) {
            // Check if key contains msal prefix; For now, we are clearing all cache items created by MSAL.js
            if (storage.hasOwnProperty(key) && (key.indexOf(Constants_1.Constants.cachePrefix) !== -1)) {
                _super.prototype.removeItem.call(this, key);
                // TODO: Clear cache based on client id (clarify use cases where this is needed)
            }
        }
    };
    /**
     * Reset all temporary cache items
     */
    AuthCache.prototype.resetTempCacheItems = function (state) {
        var _this = this;
        var stateId = state && RequestUtils_1.RequestUtils.parseLibraryState(state).id;
        var isTokenRenewalInProgress = this.tokenRenewalInProgress(state);
        var storage = window[this.cacheLocation];
        var key;
        // check state and remove associated cache
        if (stateId && !isTokenRenewalInProgress) {
            Object.keys(storage).forEach(function (key) {
                if (key.indexOf(stateId) !== -1) {
                    _this.removeItem(key);
                    _super.prototype.clearItemCookie.call(_this, key);
                }
            });
        }
        // delete the interaction status cache
        this.removeItem(Constants_1.TemporaryCacheKeys.INTERACTION_STATUS);
        this.removeItem(Constants_1.TemporaryCacheKeys.REDIRECT_REQUEST);
    };
    /**
     * Set cookies for IE
     * @param cName
     * @param cValue
     * @param expires
     */
    AuthCache.prototype.setItemCookie = function (cName, cValue, expires) {
        _super.prototype.setItemCookie.call(this, this.generateCacheKey(cName, true), cValue, expires);
        if (this.rollbackEnabled) {
            _super.prototype.setItemCookie.call(this, this.generateCacheKey(cName, false), cValue, expires);
        }
    };
    AuthCache.prototype.clearItemCookie = function (cName) {
        _super.prototype.clearItemCookie.call(this, this.generateCacheKey(cName, true));
        if (this.rollbackEnabled) {
            _super.prototype.clearItemCookie.call(this, this.generateCacheKey(cName, false));
        }
    };
    /**
     * get one item by key from cookies
     * @param cName
     */
    AuthCache.prototype.getItemCookie = function (cName) {
        return _super.prototype.getItemCookie.call(this, this.generateCacheKey(cName, true));
    };
    /**
     * Get all access tokens in the cache
     * @param clientId
     * @param homeAccountIdentifier
     */
    AuthCache.prototype.getAllAccessTokens = function (clientId, homeAccountIdentifier) {
        var _this = this;
        var results = Object.keys(window[this.cacheLocation]).reduce(function (tokens, key) {
            var keyMatches = key.match(clientId) && key.match(homeAccountIdentifier) && key.match(Constants_1.Constants.scopes);
            if (keyMatches) {
                var value = _this.getItem(key);
                if (value) {
                    try {
                        var parseAtKey = JSON.parse(key);
                        var newAccessTokenCacheItem = new AccessTokenCacheItem_1.AccessTokenCacheItem(parseAtKey, JSON.parse(value));
                        return tokens.concat([newAccessTokenCacheItem]);
                    }
                    catch (e) {
                        throw ClientAuthError_1.ClientAuthError.createCacheParseError(key);
                    }
                }
            }
            return tokens;
        }, []);
        return results;
    };
    /**
     * Return if the token renewal is still in progress
     * @param stateValue
     */
    AuthCache.prototype.tokenRenewalInProgress = function (stateValue) {
        var renewStatus = this.getItem(AuthCache.generateTemporaryCacheKey(Constants_1.TemporaryCacheKeys.RENEW_STATUS, stateValue));
        return !!(renewStatus && renewStatus === Constants_1.Constants.inProgress);
    };
    /**
     * Clear all cookies
     */
    AuthCache.prototype.clearMsalCookie = function (state) {
        var _this = this;
        /*
         * If state is truthy, remove values associated with that request.
         * Otherwise, remove all MSAL cookies.
         */
        if (state) {
            this.clearItemCookie(AuthCache.generateTemporaryCacheKey(Constants_1.TemporaryCacheKeys.NONCE_IDTOKEN, state));
            this.clearItemCookie(AuthCache.generateTemporaryCacheKey(Constants_1.TemporaryCacheKeys.STATE_LOGIN, state));
            this.clearItemCookie(AuthCache.generateTemporaryCacheKey(Constants_1.TemporaryCacheKeys.LOGIN_REQUEST, state));
            this.clearItemCookie(AuthCache.generateTemporaryCacheKey(Constants_1.TemporaryCacheKeys.STATE_ACQ_TOKEN, state));
        }
        else {
            var cookies = document.cookie.split(";");
            cookies.forEach(function (cookieString) {
                var cookieName = cookieString.trim().split("=")[0];
                if (cookieName.indexOf(Constants_1.Constants.cachePrefix) > -1) {
                    _super.prototype.clearItemCookie.call(_this, cookieName);
                }
            });
        }
    };
    /**
     * Create acquireTokenAccountKey to cache account object
     * @param accountId
     * @param state
     */
    AuthCache.generateAcquireTokenAccountKey = function (accountId, state) {
        var stateId = RequestUtils_1.RequestUtils.parseLibraryState(state).id;
        return "" + Constants_1.TemporaryCacheKeys.ACQUIRE_TOKEN_ACCOUNT + Constants_1.Constants.resourceDelimiter + accountId + Constants_1.Constants.resourceDelimiter + stateId;
    };
    /**
     * Create authorityKey to cache authority
     * @param state
     */
    AuthCache.generateAuthorityKey = function (state) {
        return AuthCache.generateTemporaryCacheKey(Constants_1.TemporaryCacheKeys.AUTHORITY, state);
    };
    /**
     * Generates the cache key for temporary cache items, using request state
     * @param tempCacheKey Cache key prefix
     * @param state Request state value
     */
    AuthCache.generateTemporaryCacheKey = function (tempCacheKey, state) {
        // Use the state id (a guid), in the interest of shorter key names, which is important for cookies.
        var stateId = RequestUtils_1.RequestUtils.parseLibraryState(state).id;
        return "" + tempCacheKey + Constants_1.Constants.resourceDelimiter + stateId;
    };
    return AuthCache;
}(BrowserStorage_1.BrowserStorage));
exports.AuthCache = AuthCache;


/***/ }),
/* 36 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
/**
 * @hidden
 */
var AccessTokenCacheItem = /** @class */ (function () {
    function AccessTokenCacheItem(key, value) {
        this.key = key;
        this.value = value;
    }
    return AccessTokenCacheItem;
}());
exports.AccessTokenCacheItem = AccessTokenCacheItem;


/***/ }),
/* 37 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
var ClientConfigurationError_1 = __webpack_require__(5);
var AuthError_1 = __webpack_require__(7);
/**
 * @hidden
 */
var BrowserStorage = /** @class */ (function () {
    function BrowserStorage(cacheLocation) {
        if (!window) {
            throw AuthError_1.AuthError.createNoWindowObjectError("Browser storage class could not find window object");
        }
        var storageSupported = typeof window[cacheLocation] !== "undefined" && window[cacheLocation] != null;
        if (!storageSupported) {
            throw ClientConfigurationError_1.ClientConfigurationError.createStorageNotSupportedError(cacheLocation);
        }
        this.cacheLocation = cacheLocation;
    }
    /**
     * add value to storage
     * @param key
     * @param value
     * @param enableCookieStorage
     */
    BrowserStorage.prototype.setItem = function (key, value, enableCookieStorage) {
        window[this.cacheLocation].setItem(key, value);
        if (enableCookieStorage) {
            this.setItemCookie(key, value);
        }
    };
    /**
     * get one item by key from storage
     * @param key
     * @param enableCookieStorage
     */
    BrowserStorage.prototype.getItem = function (key, enableCookieStorage) {
        if (enableCookieStorage && this.getItemCookie(key)) {
            return this.getItemCookie(key);
        }
        return window[this.cacheLocation].getItem(key);
    };
    /**
     * remove value from storage
     * @param key
     */
    BrowserStorage.prototype.removeItem = function (key) {
        return window[this.cacheLocation].removeItem(key);
    };
    /**
     * clear storage (remove all items from it)
     */
    BrowserStorage.prototype.clear = function () {
        return window[this.cacheLocation].clear();
    };
    /**
     * add value to cookies
     * @param cName
     * @param cValue
     * @param expires
     */
    BrowserStorage.prototype.setItemCookie = function (cName, cValue, expires) {
        var cookieStr = cName + "=" + cValue + ";path=/;";
        if (expires) {
            var expireTime = this.getCookieExpirationTime(expires);
            cookieStr += "expires=" + expireTime + ";";
        }
        document.cookie = cookieStr;
    };
    /**
     * get one item by key from cookies
     * @param cName
     */
    BrowserStorage.prototype.getItemCookie = function (cName) {
        var name = cName + "=";
        var ca = document.cookie.split(";");
        for (var i = 0; i < ca.length; i++) {
            var c = ca[i];
            while (c.charAt(0) === " ") {
                c = c.substring(1);
            }
            if (c.indexOf(name) === 0) {
                return c.substring(name.length, c.length);
            }
        }
        return "";
    };
    /**
     * Clear an item in the cookies by key
     * @param cName
     */
    BrowserStorage.prototype.clearItemCookie = function (cName) {
        this.setItemCookie(cName, "", -1);
    };
    /**
     * Get cookie expiration time
     * @param cookieLifeDays
     */
    BrowserStorage.prototype.getCookieExpirationTime = function (cookieLifeDays) {
        var today = new Date();
        var expr = new Date(today.getTime() + cookieLifeDays * 24 * 60 * 60 * 1000);
        return expr.toUTCString();
    };
    return BrowserStorage;
}());
exports.BrowserStorage = BrowserStorage;


/***/ }),
/* 38 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = __webpack_require__(0);
/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
/**
 * @hidden
 */
var ResponseUtils = /** @class */ (function () {
    function ResponseUtils() {
    }
    ResponseUtils.setResponseIdToken = function (originalResponse, idTokenObj) {
        if (!originalResponse) {
            return null;
        }
        else if (!idTokenObj) {
            return originalResponse;
        }
        var exp = Number(idTokenObj.expiration);
        if (exp && !originalResponse.expiresOn) {
            originalResponse.expiresOn = new Date(exp * 1000);
        }
        return tslib_1.__assign({}, originalResponse, { idToken: idTokenObj, idTokenClaims: idTokenObj.claims, uniqueId: idTokenObj.objectId || idTokenObj.subject, tenantId: idTokenObj.tenantId });
    };
    return ResponseUtils;
}());
exports.ResponseUtils = ResponseUtils;


/***/ }),
/* 39 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = __webpack_require__(0);
var DefaultEvent_1 = tslib_1.__importDefault(__webpack_require__(40));
var Constants_1 = __webpack_require__(1);
var ApiEvent_1 = tslib_1.__importDefault(__webpack_require__(28));
var HttpEvent_1 = tslib_1.__importDefault(__webpack_require__(41));
// for use in cache events
var MSAL_CACHE_EVENT_VALUE_PREFIX = "msal.token";
var MSAL_CACHE_EVENT_NAME = "msal.cache_event";
var TelemetryManager = /** @class */ (function () {
    function TelemetryManager(config, telemetryEmitter, logger) {
        // correlation Id to list of events
        this.completedEvents = {};
        // event key to event
        this.inProgressEvents = {};
        // correlation id to map of eventname to count
        this.eventCountByCorrelationId = {};
        // Implement after API EVENT
        this.onlySendFailureTelemetry = false;
        // TODO THROW if bad options
        this.telemetryPlatform = tslib_1.__assign({ sdk: Constants_1.Constants.libraryName, sdkVersion: Constants_1.libraryVersion(), networkInformation: {
                // @ts-ignore
                connectionSpeed: typeof navigator !== "undefined" && navigator.connection && navigator.connection.effectiveType
            } }, config.platform);
        this.clientId = config.clientId;
        this.onlySendFailureTelemetry = config.onlySendFailureTelemetry;
        /*
         * TODO, when i get to wiring this through, think about what it means if
         * a developer does not implement telem at all, we still instrument, but telemetryEmitter can be
         * optional?
         */
        this.telemetryEmitter = telemetryEmitter;
        this.logger = logger;
    }
    TelemetryManager.getTelemetrymanagerStub = function (clientId, logger) {
        var applicationName = "UnSetStub";
        var applicationVersion = "0.0";
        var telemetryEmitter = function () { };
        var telemetryPlatform = {
            applicationName: applicationName,
            applicationVersion: applicationVersion
        };
        var telemetryManagerConfig = {
            platform: telemetryPlatform,
            clientId: clientId
        };
        return new this(telemetryManagerConfig, telemetryEmitter, logger);
    };
    TelemetryManager.prototype.startEvent = function (event) {
        this.logger.verbose("Telemetry Event started: " + event.key);
        if (!this.telemetryEmitter) {
            return;
        }
        event.start();
        this.inProgressEvents[event.key] = event;
    };
    TelemetryManager.prototype.stopEvent = function (event) {
        this.logger.verbose("Telemetry Event stopped: " + event.key);
        if (!this.telemetryEmitter || !this.inProgressEvents[event.key]) {
            return;
        }
        event.stop();
        this.incrementEventCount(event);
        var completedEvents = this.completedEvents[event.telemetryCorrelationId];
        this.completedEvents[event.telemetryCorrelationId] = (completedEvents || []).concat([event]);
        delete this.inProgressEvents[event.key];
    };
    TelemetryManager.prototype.flush = function (correlationId) {
        var _this = this;
        this.logger.verbose("Flushing telemetry events: " + correlationId);
        // If there is only unfinished events should this still return them?
        if (!this.telemetryEmitter || !this.completedEvents[correlationId]) {
            return;
        }
        var orphanedEvents = this.getOrphanedEvents(correlationId);
        orphanedEvents.forEach(function (event) { return _this.incrementEventCount(event); });
        var eventsToFlush = this.completedEvents[correlationId].concat(orphanedEvents);
        delete this.completedEvents[correlationId];
        var eventCountsToFlush = this.eventCountByCorrelationId[correlationId];
        delete this.eventCountByCorrelationId[correlationId];
        // TODO add funcitonality for onlyFlushFailures after implementing api event? ??
        if (!eventsToFlush || !eventsToFlush.length) {
            return;
        }
        var defaultEvent = new DefaultEvent_1.default(this.telemetryPlatform, correlationId, this.clientId, eventCountsToFlush);
        var eventsWithDefaultEvent = eventsToFlush.concat([defaultEvent]);
        this.telemetryEmitter(eventsWithDefaultEvent.map(function (e) { return e.get(); }));
    };
    TelemetryManager.prototype.createAndStartApiEvent = function (correlationId, apiEventIdentifier) {
        var apiEvent = new ApiEvent_1.default(correlationId, this.logger.isPiiLoggingEnabled(), apiEventIdentifier);
        this.startEvent(apiEvent);
        return apiEvent;
    };
    TelemetryManager.prototype.stopAndFlushApiEvent = function (correlationId, apiEvent, wasSuccessful, errorCode) {
        apiEvent.wasSuccessful = wasSuccessful;
        if (errorCode) {
            apiEvent.apiErrorCode = errorCode;
        }
        this.stopEvent(apiEvent);
        this.flush(correlationId);
    };
    TelemetryManager.prototype.createAndStartHttpEvent = function (correlation, httpMethod, url, eventLabel) {
        var httpEvent = new HttpEvent_1.default(correlation, eventLabel);
        httpEvent.url = url;
        httpEvent.httpMethod = httpMethod;
        this.startEvent(httpEvent);
        return httpEvent;
    };
    TelemetryManager.prototype.incrementEventCount = function (event) {
        var _a;
        /*
         * TODO, name cache event different?
         * if type is cache event, change name
         */
        var eventName = event.eventName;
        var eventCount = this.eventCountByCorrelationId[event.telemetryCorrelationId];
        if (!eventCount) {
            this.eventCountByCorrelationId[event.telemetryCorrelationId] = (_a = {},
                _a[eventName] = 1,
                _a);
        }
        else {
            eventCount[eventName] = eventCount[eventName] ? eventCount[eventName] + 1 : 1;
        }
    };
    TelemetryManager.prototype.getOrphanedEvents = function (correlationId) {
        var _this = this;
        return Object.keys(this.inProgressEvents)
            .reduce(function (memo, eventKey) {
            if (eventKey.indexOf(correlationId) !== -1) {
                var event = _this.inProgressEvents[eventKey];
                delete _this.inProgressEvents[eventKey];
                return memo.concat([event]);
            }
            return memo;
        }, []);
    };
    return TelemetryManager;
}());
exports.default = TelemetryManager;


/***/ }),
/* 40 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = __webpack_require__(0);
var TelemetryConstants_1 = __webpack_require__(8);
var TelemetryEvent_1 = tslib_1.__importDefault(__webpack_require__(14));
var TelemetryUtils_1 = __webpack_require__(10);
var DefaultEvent = /** @class */ (function (_super) {
    tslib_1.__extends(DefaultEvent, _super);
    // TODO Platform Type
    function DefaultEvent(platform, correlationId, clientId, eventCount) {
        var _this = _super.call(this, TelemetryUtils_1.prependEventNamePrefix("default_event"), correlationId, "DefaultEvent") || this;
        _this.event[TelemetryUtils_1.prependEventNamePrefix("client_id")] = clientId;
        _this.event[TelemetryUtils_1.prependEventNamePrefix("sdk_plaform")] = platform.sdk;
        _this.event[TelemetryUtils_1.prependEventNamePrefix("sdk_version")] = platform.sdkVersion;
        _this.event[TelemetryUtils_1.prependEventNamePrefix("application_name")] = platform.applicationName;
        _this.event[TelemetryUtils_1.prependEventNamePrefix("application_version")] = platform.applicationVersion;
        _this.event[TelemetryUtils_1.prependEventNamePrefix("effective_connection_speed")] = platform.networkInformation && platform.networkInformation.connectionSpeed;
        _this.event["" + TelemetryConstants_1.TELEMETRY_BLOB_EVENT_NAMES.UiEventCountTelemetryBatchKey] = _this.getEventCount(TelemetryUtils_1.prependEventNamePrefix("ui_event"), eventCount);
        _this.event["" + TelemetryConstants_1.TELEMETRY_BLOB_EVENT_NAMES.HttpEventCountTelemetryBatchKey] = _this.getEventCount(TelemetryUtils_1.prependEventNamePrefix("http_event"), eventCount);
        _this.event["" + TelemetryConstants_1.TELEMETRY_BLOB_EVENT_NAMES.CacheEventCountConstStrKey] = _this.getEventCount(TelemetryUtils_1.prependEventNamePrefix("cache_event"), eventCount);
        return _this;
        // / Device id?
    }
    DefaultEvent.prototype.getEventCount = function (eventName, eventCount) {
        if (!eventCount[eventName]) {
            return 0;
        }
        return eventCount[eventName];
    };
    return DefaultEvent;
}(TelemetryEvent_1.default));
exports.default = DefaultEvent;


/***/ }),
/* 41 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = __webpack_require__(0);
var TelemetryEvent_1 = tslib_1.__importDefault(__webpack_require__(14));
var TelemetryUtils_1 = __webpack_require__(10);
var ServerRequestParameters_1 = __webpack_require__(16);
exports.EVENT_KEYS = {
    HTTP_PATH: TelemetryUtils_1.prependEventNamePrefix("http_path"),
    USER_AGENT: TelemetryUtils_1.prependEventNamePrefix("user_agent"),
    QUERY_PARAMETERS: TelemetryUtils_1.prependEventNamePrefix("query_parameters"),
    API_VERSION: TelemetryUtils_1.prependEventNamePrefix("api_version"),
    RESPONSE_CODE: TelemetryUtils_1.prependEventNamePrefix("response_code"),
    O_AUTH_ERROR_CODE: TelemetryUtils_1.prependEventNamePrefix("oauth_error_code"),
    HTTP_METHOD: TelemetryUtils_1.prependEventNamePrefix("http_method"),
    REQUEST_ID_HEADER: TelemetryUtils_1.prependEventNamePrefix("request_id_header"),
    SPE_INFO: TelemetryUtils_1.prependEventNamePrefix("spe_info"),
    SERVER_ERROR_CODE: TelemetryUtils_1.prependEventNamePrefix("server_error_code"),
    SERVER_SUB_ERROR_CODE: TelemetryUtils_1.prependEventNamePrefix("server_sub_error_code"),
    URL: TelemetryUtils_1.prependEventNamePrefix("url")
};
var HttpEvent = /** @class */ (function (_super) {
    tslib_1.__extends(HttpEvent, _super);
    function HttpEvent(correlationId, eventLabel) {
        return _super.call(this, TelemetryUtils_1.prependEventNamePrefix("http_event"), correlationId, eventLabel) || this;
    }
    Object.defineProperty(HttpEvent.prototype, "url", {
        set: function (url) {
            var scrubbedUri = TelemetryUtils_1.scrubTenantFromUri(url);
            this.event[exports.EVENT_KEYS.URL] = scrubbedUri && scrubbedUri.toLowerCase();
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(HttpEvent.prototype, "httpPath", {
        set: function (httpPath) {
            this.event[exports.EVENT_KEYS.HTTP_PATH] = TelemetryUtils_1.scrubTenantFromUri(httpPath).toLowerCase();
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(HttpEvent.prototype, "userAgent", {
        set: function (userAgent) {
            this.event[exports.EVENT_KEYS.USER_AGENT] = userAgent;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(HttpEvent.prototype, "queryParams", {
        set: function (queryParams) {
            this.event[exports.EVENT_KEYS.QUERY_PARAMETERS] = ServerRequestParameters_1.ServerRequestParameters.generateQueryParametersString(queryParams);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(HttpEvent.prototype, "apiVersion", {
        set: function (apiVersion) {
            this.event[exports.EVENT_KEYS.API_VERSION] = apiVersion.toLowerCase();
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(HttpEvent.prototype, "httpResponseStatus", {
        set: function (statusCode) {
            this.event[exports.EVENT_KEYS.RESPONSE_CODE] = statusCode;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(HttpEvent.prototype, "oAuthErrorCode", {
        set: function (errorCode) {
            this.event[exports.EVENT_KEYS.O_AUTH_ERROR_CODE] = errorCode;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(HttpEvent.prototype, "httpMethod", {
        set: function (httpMethod) {
            this.event[exports.EVENT_KEYS.HTTP_METHOD] = httpMethod;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(HttpEvent.prototype, "requestIdHeader", {
        set: function (requestIdHeader) {
            this.event[exports.EVENT_KEYS.REQUEST_ID_HEADER] = requestIdHeader;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(HttpEvent.prototype, "speInfo", {
        /**
         * Indicates whether the request was executed on a ring serving SPE traffic.
         * An empty string indicates this occurred on an outer ring, and the string "I"
         * indicates the request occurred on the inner ring
         */
        set: function (speInfo) {
            this.event[exports.EVENT_KEYS.SPE_INFO] = speInfo;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(HttpEvent.prototype, "serverErrorCode", {
        set: function (errorCode) {
            this.event[exports.EVENT_KEYS.SERVER_ERROR_CODE] = errorCode;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(HttpEvent.prototype, "serverSubErrorCode", {
        set: function (subErrorCode) {
            this.event[exports.EVENT_KEYS.SERVER_SUB_ERROR_CODE] = subErrorCode;
        },
        enumerable: true,
        configurable: true
    });
    return HttpEvent;
}(TelemetryEvent_1.default));
exports.default = HttpEvent;


/***/ }),
/* 42 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
Object.defineProperty(exports, "__esModule", { value: true });
var ClientConfigurationError_1 = __webpack_require__(5);
function validateClaimsRequest(request) {
    if (!request.claimsRequest) {
        return;
    }
    var claims;
    try {
        claims = JSON.parse(request.claimsRequest);
    }
    catch (e) {
        throw ClientConfigurationError_1.ClientConfigurationError.createClaimsRequestParsingError(e);
    }
    // TODO: More validation will be added when the server team tells us how they have actually implemented claims
}
exports.validateClaimsRequest = validateClaimsRequest;


/***/ })
/******/ ]);
});
//# sourceMappingURL=msal.js.map

/***/ }),

/***/ "./packages/Microsoft.Office.WebAuth.Implicit/lib/msal.min.js":
/***/ (function(module, exports, __webpack_require__) {

"use strict";
/*! msal v1.3.3 2020-07-14 */
!function(e,t){ true?module.exports=t():undefined}(window,function(){return o={},n.m=r=[function(e,t,r){
/*! *****************************************************************************
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the Apache License, Version 2.0 (the "License"); you may not use
this file except in compliance with the License. You may obtain a copy of the
License at http://www.apache.org/licenses/LICENSE-2.0

THIS CODE IS PROVIDED ON AN *AS IS* BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
KIND, EITHER EXPRESS OR IMPLIED, INCLUDING WITHOUT LIMITATION ANY IMPLIED
WARRANTIES OR CONDITIONS OF TITLE, FITNESS FOR A PARTICULAR PURPOSE,
MERCHANTABLITY OR NON-INFRINGEMENT.

See the Apache Version 2.0 License for specific language governing permissions
and limitations under the License.
***************************************************************************** */
Object.defineProperty(t,"__esModule",{value:!0});var o=function(e,t){return(o=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(e,t){e.__proto__=t}||function(e,t){for(var r in t)t.hasOwnProperty(r)&&(e[r]=t[r])})(e,t)};function i(e){var t="function"==typeof Symbol&&e[Symbol.iterator],r=0;return t?t.call(e):{next:function(){return e&&r>=e.length&&(e=void 0),{value:e&&e[r++],done:!e}}}}function n(e,t){var r="function"==typeof Symbol&&e[Symbol.iterator];if(!r)return e;var o,n,i=r.call(e),a=[];try{for(;(void 0===t||0<t--)&&!(o=i.next()).done;)a.push(o.value)}catch(e){n={error:e}}finally{try{o&&!o.done&&(r=i.return)&&r.call(i)}finally{if(n)throw n.error}}return a}function h(e){return this instanceof h?(this.v=e,this):new h(e)}t.__extends=function(e,t){function r(){this.constructor=e}o(e,t),e.prototype=null===t?Object.create(t):(r.prototype=t.prototype,new r)},t.__assign=function(){return t.__assign=Object.assign||function(e){for(var t,r=1,o=arguments.length;r<o;r++)for(var n in t=arguments[r])Object.prototype.hasOwnProperty.call(t,n)&&(e[n]=t[n]);return e},t.__assign.apply(this,arguments)},t.__rest=function(e,t){var r={};for(var o in e)Object.prototype.hasOwnProperty.call(e,o)&&t.indexOf(o)<0&&(r[o]=e[o]);if(null!=e&&"function"==typeof Object.getOwnPropertySymbols){var n=0;for(o=Object.getOwnPropertySymbols(e);n<o.length;n++)t.indexOf(o[n])<0&&Object.prototype.propertyIsEnumerable.call(e,o[n])&&(r[o[n]]=e[o[n]])}return r},t.__decorate=function(e,t,r,o){var n,i=arguments.length,a=i<3?t:null===o?o=Object.getOwnPropertyDescriptor(t,r):o;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)a=Reflect.decorate(e,t,r,o);else for(var s=e.length-1;0<=s;s--)(n=e[s])&&(a=(i<3?n(a):3<i?n(t,r,a):n(t,r))||a);return 3<i&&a&&Object.defineProperty(t,r,a),a},t.__param=function(r,o){return function(e,t){o(e,t,r)}},t.__metadata=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)},t.__awaiter=function(i,a,s,c){return new(s=s||Promise)(function(e,t){function r(e){try{n(c.next(e))}catch(e){t(e)}}function o(e){try{n(c.throw(e))}catch(e){t(e)}}function n(t){t.done?e(t.value):new s(function(e){e(t.value)}).then(r,o)}n((c=c.apply(i,a||[])).next())})},t.__generator=function(r,o){var n,i,a,e,s={label:0,sent:function(){if(1&a[0])throw a[1];return a[1]},trys:[],ops:[]};return e={next:t(0),throw:t(1),return:t(2)},"function"==typeof Symbol&&(e[Symbol.iterator]=function(){return this}),e;function t(t){return function(e){return function(t){if(n)throw new TypeError("Generator is already executing.");for(;s;)try{if(n=1,i&&(a=2&t[0]?i.return:t[0]?i.throw||((a=i.return)&&a.call(i),0):i.next)&&!(a=a.call(i,t[1])).done)return a;switch(i=0,a&&(t=[2&t[0],a.value]),t[0]){case 0:case 1:a=t;break;case 4:return s.label++,{value:t[1],done:!1};case 5:s.label++,i=t[1],t=[0];continue;case 7:t=s.ops.pop(),s.trys.pop();continue;default:if(!(a=0<(a=s.trys).length&&a[a.length-1])&&(6===t[0]||2===t[0])){s=0;continue}if(3===t[0]&&(!a||t[1]>a[0]&&t[1]<a[3])){s.label=t[1];break}if(6===t[0]&&s.label<a[1]){s.label=a[1],a=t;break}if(a&&s.label<a[2]){s.label=a[2],s.ops.push(t);break}a[2]&&s.ops.pop(),s.trys.pop();continue}t=o.call(r,s)}catch(e){t=[6,e],i=0}finally{n=a=0}if(5&t[0])throw t[1];return{value:t[0]?t[1]:void 0,done:!0}}([t,e])}}},t.__exportStar=function(e,t){for(var r in e)t.hasOwnProperty(r)||(t[r]=e[r])},t.__values=i,t.__read=n,t.__spread=function(){for(var e=[],t=0;t<arguments.length;t++)e=e.concat(n(arguments[t]));return e},t.__spreadArrays=function(){for(var e=0,t=0,r=arguments.length;t<r;t++)e+=arguments[t].length;var o=Array(e),n=0;for(t=0;t<r;t++)for(var i=arguments[t],a=0,s=i.length;a<s;a++,n++)o[n]=i[a];return o},t.__await=h,t.__asyncGenerator=function(e,t,r){if(!Symbol.asyncIterator)throw new TypeError("Symbol.asyncIterator is not defined.");var n,i=r.apply(e,t||[]),a=[];return n={},o("next"),o("throw"),o("return"),n[Symbol.asyncIterator]=function(){return this},n;function o(o){i[o]&&(n[o]=function(r){return new Promise(function(e,t){1<a.push([o,r,e,t])||s(o,r)})})}function s(e,t){try{!function(e){e.value instanceof h?Promise.resolve(e.value.v).then(c,u):l(a[0][2],e)}(i[e](t))}catch(e){l(a[0][3],e)}}function c(e){s("next",e)}function u(e){s("throw",e)}function l(e,t){e(t),a.shift(),a.length&&s(a[0][0],a[0][1])}},t.__asyncDelegator=function(o){var e,n;return e={},t("next"),t("throw",function(e){throw e}),t("return"),e[Symbol.iterator]=function(){return this},e;function t(t,r){e[t]=o[t]?function(e){return(n=!n)?{value:h(o[t](e)),done:"return"===t}:r?r(e):e}:r}},t.__asyncValues=function(n){if(!Symbol.asyncIterator)throw new TypeError("Symbol.asyncIterator is not defined.");var e,t=n[Symbol.asyncIterator];return t?t.call(n):(n=i(n),e={},r("next"),r("throw"),r("return"),e[Symbol.asyncIterator]=function(){return this},e);function r(o){e[o]=n[o]&&function(r){return new Promise(function(e,t){(function(t,e,r,o){Promise.resolve(o).then(function(e){t({value:e,done:r})},e)})(e,t,(r=n[o](r)).done,r.value)})}}},t.__makeTemplateObject=function(e,t){return Object.defineProperty?Object.defineProperty(e,"raw",{value:t}):e.raw=t,e},t.__importStar=function(e){if(e&&e.__esModule)return e;var t={};if(null!=e)for(var r in e)Object.hasOwnProperty.call(e,r)&&(t[r]=e[r]);return t.default=e,t},t.__importDefault=function(e){return e&&e.__esModule?e:{default:e}}},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var o,n,i,a,s,c,u=(Object.defineProperty(l,"libraryName",{get:function(){return"Msal.js"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"claims",{get:function(){return"claims"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"clientId",{get:function(){return"clientId"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"adalIdToken",{get:function(){return"adal.idtoken"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"cachePrefix",{get:function(){return"msal"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"scopes",{get:function(){return"scopes"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"no_account",{get:function(){return"NO_ACCOUNT"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"upn",{get:function(){return"upn"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"domain_hint",{get:function(){return"domain_hint"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"prompt_select_account",{get:function(){return"&prompt=select_account"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"prompt_none",{get:function(){return"&prompt=none"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"prompt",{get:function(){return"prompt"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"response_mode_fragment",{get:function(){return"&response_mode=fragment"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"resourceDelimiter",{get:function(){return"|"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"cacheDelimiter",{get:function(){return"."},enumerable:!0,configurable:!0}),Object.defineProperty(l,"popUpWidth",{get:function(){return this._popUpWidth},set:function(e){this._popUpWidth=e},enumerable:!0,configurable:!0}),Object.defineProperty(l,"popUpHeight",{get:function(){return this._popUpHeight},set:function(e){this._popUpHeight=e},enumerable:!0,configurable:!0}),Object.defineProperty(l,"login",{get:function(){return"LOGIN"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"renewToken",{get:function(){return"RENEW_TOKEN"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"unknown",{get:function(){return"UNKNOWN"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"ADFS",{get:function(){return"adfs"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"homeAccountIdentifier",{get:function(){return"homeAccountIdentifier"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"common",{get:function(){return"common"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"openidScope",{get:function(){return"openid"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"profileScope",{get:function(){return"profile"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"interactionTypeRedirect",{get:function(){return"redirectInteraction"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"interactionTypePopup",{get:function(){return"popupInteraction"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"interactionTypeSilent",{get:function(){return"silentInteraction"},enumerable:!0,configurable:!0}),Object.defineProperty(l,"inProgress",{get:function(){return"inProgress"},enumerable:!0,configurable:!0}),l._popUpWidth=483,l._popUpHeight=600,l);function l(){}t.Constants=u,(o=t.ServerHashParamKeys||(t.ServerHashParamKeys={})).SCOPE="scope",o.STATE="state",o.ERROR="error",o.ERROR_DESCRIPTION="error_description",o.ACCESS_TOKEN="access_token",o.ID_TOKEN="id_token",o.EXPIRES_IN="expires_in",o.SESSION_STATE="session_state",o.CLIENT_INFO="client_info",(n=t.TemporaryCacheKeys||(t.TemporaryCacheKeys={})).AUTHORITY="authority",n.ACQUIRE_TOKEN_ACCOUNT="acquireTokenAccount",n.SESSION_STATE="session.state",n.STATE_LOGIN="state.login",n.STATE_ACQ_TOKEN="state.acquireToken",n.STATE_RENEW="state.renew",n.NONCE_IDTOKEN="nonce.idtoken",n.LOGIN_REQUEST="login.request",n.RENEW_STATUS="token.renew.status",n.URL_HASH="urlHash",n.INTERACTION_STATUS="interaction_status",n.REDIRECT_REQUEST="redirect_request",(i=t.PersistentCacheKeys||(t.PersistentCacheKeys={})).IDTOKEN="idtoken",i.CLIENT_INFO="client.info",(a=t.ErrorCacheKeys||(t.ErrorCacheKeys={})).LOGIN_ERROR="login.error",a.ERROR="error",a.ERROR_DESC="error.description",t.DEFAULT_AUTHORITY="https://login.microsoftonline.com/common/",t.AAD_INSTANCE_DISCOVERY_ENDPOINT=t.DEFAULT_AUTHORITY+"/discovery/instance?api-version=1.1&authorization_endpoint=",(c=s=t.SSOTypes||(t.SSOTypes={})).ACCOUNT="account",c.SID="sid",c.LOGIN_HINT="login_hint",c.ID_TOKEN="id_token",c.ACCOUNT_ID="accountIdentifier",c.HOMEACCOUNT_ID="homeAccountIdentifier",t.BlacklistedEQParams=[s.SID,s.LOGIN_HINT],t.NetworkRequestType={GET:"GET",POST:"POST"},t.PromptState={LOGIN:"login",SELECT_ACCOUNT:"select_account",CONSENT:"consent",NONE:"none"},t.FramePrefix={ID_TOKEN_FRAME:"msalIdTokenFrame",TOKEN_FRAME:"msalRenewFrame"},t.libraryVersion=function(){return"1.3.3"}},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var o=(s.createNewGuid=function(){var e=window.crypto;if(e&&e.getRandomValues){var t=new Uint8Array(16);return e.getRandomValues(t),t[6]|=64,t[6]&=79,t[8]|=128,t[8]&=191,s.decimalToHex(t[0])+s.decimalToHex(t[1])+s.decimalToHex(t[2])+s.decimalToHex(t[3])+"-"+s.decimalToHex(t[4])+s.decimalToHex(t[5])+"-"+s.decimalToHex(t[6])+s.decimalToHex(t[7])+"-"+s.decimalToHex(t[8])+s.decimalToHex(t[9])+"-"+s.decimalToHex(t[10])+s.decimalToHex(t[11])+s.decimalToHex(t[12])+s.decimalToHex(t[13])+s.decimalToHex(t[14])+s.decimalToHex(t[15])}for(var r="xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx",o="0123456789abcdef",n=0,i="",a=0;a<36;a++)"-"!==r[a]&&"4"!==r[a]&&(n=16*Math.random()|0),"x"===r[a]?i+=o[n]:"y"===r[a]?(n&=3,i+=o[n|=8]):i+=r[a];return i},s.isGuid=function(e){return/^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(e)},s.decimalToHex=function(e){for(var t=e.toString(16);t.length<2;)t="0"+t;return t},s.base64Encode=function(e){return btoa(encodeURIComponent(e).replace(/%([0-9A-F]{2})/g,function(e,t){return String.fromCharCode(Number("0x"+t))}))},s.base64Decode=function(e){var t=e.replace(/-/g,"+").replace(/_/g,"/");switch(t.length%4){case 0:break;case 2:t+="==";break;case 3:t+="=";break;default:throw new Error("Invalid base64 string")}return decodeURIComponent(atob(t).split("").map(function(e){return"%"+("00"+e.charCodeAt(0).toString(16)).slice(-2)}).join(""))},s.deserialize=function(e){function t(e){return decodeURIComponent(decodeURIComponent(e.replace(o," ")))}var r,o=/\+/g,n=/([^&=]+)=([^&]*)/g,i={};for(r=n.exec(e);r;)i[t(r[1])]=t(r[2]),r=n.exec(e);return i},s);function s(){}t.CryptoUtils=o},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var o=(n.isEmpty=function(e){return void 0===e||!e||0===e.length},n);function n(){}t.StringUtils=o},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var n=r(1),o=r(9),i=r(3),a=r(2),s=(c.createNavigateUrl=function(e){var t=this.createNavigationUrlString(e),r=e.authorityInstance.AuthorizationEndpoint;return r.indexOf("?")<0?r+="?":r+="&",""+r+t.join("&")},c.createNavigationUrlString=function(e){var t=e.scopes;-1===t.indexOf(e.clientId)&&t.push(e.clientId);var r=[];return r.push("response_type="+e.responseType),this.translateclientIdUsedInScope(t,e.clientId),r.push("scope="+encodeURIComponent(o.ScopeSet.parseScope(t))),r.push("client_id="+encodeURIComponent(e.clientId)),r.push("redirect_uri="+encodeURIComponent(e.redirectUri)),r.push("state="+encodeURIComponent(e.state)),r.push("nonce="+encodeURIComponent(e.nonce)),r.push("client_info=1"),r.push("x-client-SKU="+e.xClientSku),r.push("x-client-Ver="+e.xClientVer),e.promptValue&&r.push("prompt="+encodeURIComponent(e.promptValue)),e.claimsValue&&r.push("claims="+encodeURIComponent(e.claimsValue)),e.queryParameters&&r.push(e.queryParameters),e.extraQueryParameters&&r.push(e.extraQueryParameters),r.push("client-request-id="+encodeURIComponent(e.correlationId)),r},c.translateclientIdUsedInScope=function(e,t){var r=e.indexOf(t);0<=r&&(e.splice(r,1),-1===e.indexOf("openid")&&e.push("openid"),-1===e.indexOf("profile")&&e.push("profile"))},c.getCurrentUrl=function(){return window.location.href.split("?")[0].split("#")[0]},c.removeHashFromUrl=function(e){return e.split("#")[0]},c.replaceTenantPath=function(e,t){e=e.toLowerCase();var r=this.GetUrlComponents(e),o=r.PathSegments;return t&&0!==o.length&&o[0]===n.Constants.common&&(o[0]=t),this.constructAuthorityUriFromObject(r,o)},c.constructAuthorityUriFromObject=function(e,t){return this.CanonicalizeUri(e.Protocol+"//"+e.HostNameAndPort+"/"+t.join("/"))},c.GetUrlComponents=function(e){if(!e)throw"Url required";var t=RegExp("^(([^:/?#]+):)?(//([^/?#]*))?([^?#]*)(\\?([^#]*))?(#(.*))?"),r=e.match(t);if(!r||r.length<6)throw"Valid url required";var o={Protocol:r[1],HostNameAndPort:r[4],AbsolutePath:r[5]},n=o.AbsolutePath.split("/");return n=n.filter(function(e){return e&&0<e.length}),o.PathSegments=n,r[6]&&(o.Search=r[6]),r[8]&&(o.Hash=r[8]),o},c.CanonicalizeUri=function(e){return(e=e&&e.toLowerCase())&&!c.endsWith(e,"/")&&(e+="/"),e},c.endsWith=function(e,t){return!(!e||!t)&&-1!==e.indexOf(t,e.length-t.length)},c.urlRemoveQueryStringParameter=function(e,t){if(i.StringUtils.isEmpty(e))return e;var r=new RegExp("(\\&"+t+"=)[^&]+");return e=e.replace(r,""),r=new RegExp("("+t+"=)[^&]+&"),e=e.replace(r,""),r=new RegExp("("+t+"=)[^&]+"),e=e.replace(r,"")},c.getHashFromUrl=function(e){var t=e.indexOf("#"),r=e.indexOf("#/");return-1<r?e.substring(r+2):-1<t?e.substring(t+1):e},c.urlContainsHash=function(e){var t=c.deserializeHash(e);return t.hasOwnProperty(n.ServerHashParamKeys.ERROR_DESCRIPTION)||t.hasOwnProperty(n.ServerHashParamKeys.ERROR)||t.hasOwnProperty(n.ServerHashParamKeys.ACCESS_TOKEN)||t.hasOwnProperty(n.ServerHashParamKeys.ID_TOKEN)},c.deserializeHash=function(e){var t=c.getHashFromUrl(e);return a.CryptoUtils.deserialize(t)},c.getHostFromUri=function(e){var t=String(e).replace(/^(https?:)\/\//,"");return t=t.split("/")[0]},c);function c(){}t.UrlUtils=s},function(e,i,t){Object.defineProperty(i,"__esModule",{value:!0});var r=t(0),o=t(6);i.ClientConfigurationErrorMessage={configurationNotSet:{code:"no_config_set",desc:"Configuration has not been set. Please call the UserAgentApplication constructor with a valid Configuration object."},storageNotSupported:{code:"storage_not_supported",desc:"The value for the cacheLocation is not supported."},noRedirectCallbacksSet:{code:"no_redirect_callbacks",desc:"No redirect callbacks have been set. Please call handleRedirectCallback() with the appropriate function arguments before continuing. More information is available here: https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki/MSAL-basics."},invalidCallbackObject:{code:"invalid_callback_object",desc:"The object passed for the callback was invalid. More information is available here: https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki/MSAL-basics."},scopesRequired:{code:"scopes_required",desc:"Scopes are required to obtain an access token."},emptyScopes:{code:"empty_input_scopes_error",desc:"Scopes cannot be passed as empty array."},nonArrayScopes:{code:"nonarray_input_scopes_error",desc:"Scopes cannot be passed as non-array."},clientScope:{code:"clientid_input_scopes_error",desc:"Client ID can only be provided as a single scope."},invalidPrompt:{code:"invalid_prompt_value",desc:"Supported prompt values are 'login', 'select_account', 'consent' and 'none'"},invalidAuthorityType:{code:"invalid_authority_type",desc:"The given authority is not a valid type of authority supported by MSAL. Please see here for valid authorities: <insert URL here>."},authorityUriInsecure:{code:"authority_uri_insecure",desc:"Authority URIs must use https."},authorityUriInvalidPath:{code:"authority_uri_invalid_path",desc:"Given authority URI is invalid."},unsupportedAuthorityValidation:{code:"unsupported_authority_validation",desc:"The authority validation is not supported for this authority type."},untrustedAuthority:{code:"untrusted_authority",desc:"The provided authority is not a trusted authority. Please include this authority in the knownAuthorities config parameter or set validateAuthority=false."},b2cAuthorityUriInvalidPath:{code:"b2c_authority_uri_invalid_path",desc:"The given URI for the B2C authority is invalid."},b2cKnownAuthoritiesNotSet:{code:"b2c_known_authorities_not_set",desc:"Must set known authorities when validateAuthority is set to True and using B2C"},claimsRequestParsingError:{code:"claims_request_parsing_error",desc:"Could not parse the given claims request object."},emptyRequestError:{code:"empty_request_error",desc:"Request object is required."},invalidCorrelationIdError:{code:"invalid_guid_sent_as_correlationId",desc:"Please set the correlationId as a valid guid"},telemetryConfigError:{code:"telemetry_config_error",desc:"Telemetry config is not configured with required values"},ssoSilentError:{code:"sso_silent_error",desc:"request must contain either sid or login_hint"},invalidAuthorityMetadataError:{code:"authority_metadata_error",desc:"Invalid authorityMetadata. Must be a JSON object containing authorization_endpoint, end_session_endpoint, and issuer fields."}};var n,a=(n=o.ClientAuthError,r.__extends(s,n),s.createNoSetConfigurationError=function(){return new s(i.ClientConfigurationErrorMessage.configurationNotSet.code,""+i.ClientConfigurationErrorMessage.configurationNotSet.desc)},s.createStorageNotSupportedError=function(e){return new s(i.ClientConfigurationErrorMessage.storageNotSupported.code,i.ClientConfigurationErrorMessage.storageNotSupported.desc+" Given location: "+e)},s.createRedirectCallbacksNotSetError=function(){return new s(i.ClientConfigurationErrorMessage.noRedirectCallbacksSet.code,i.ClientConfigurationErrorMessage.noRedirectCallbacksSet.desc)},s.createInvalidCallbackObjectError=function(e){return new s(i.ClientConfigurationErrorMessage.invalidCallbackObject.code,i.ClientConfigurationErrorMessage.invalidCallbackObject.desc+" Given value for callback function: "+e)},s.createEmptyScopesArrayError=function(e){return new s(i.ClientConfigurationErrorMessage.emptyScopes.code,i.ClientConfigurationErrorMessage.emptyScopes.desc+" Given value: "+e+".")},s.createScopesNonArrayError=function(e){return new s(i.ClientConfigurationErrorMessage.nonArrayScopes.code,i.ClientConfigurationErrorMessage.nonArrayScopes.desc+" Given value: "+e+".")},s.createClientIdSingleScopeError=function(e){return new s(i.ClientConfigurationErrorMessage.clientScope.code,i.ClientConfigurationErrorMessage.clientScope.desc+" Given value: "+e+".")},s.createScopesRequiredError=function(e){return new s(i.ClientConfigurationErrorMessage.scopesRequired.code,i.ClientConfigurationErrorMessage.scopesRequired.desc+" Given value: "+e)},s.createInvalidPromptError=function(e){return new s(i.ClientConfigurationErrorMessage.invalidPrompt.code,i.ClientConfigurationErrorMessage.invalidPrompt.desc+" Given value: "+e)},s.createClaimsRequestParsingError=function(e){return new s(i.ClientConfigurationErrorMessage.claimsRequestParsingError.code,i.ClientConfigurationErrorMessage.claimsRequestParsingError.desc+" Given value: "+e)},s.createEmptyRequestError=function(){var e=i.ClientConfigurationErrorMessage.emptyRequestError;return new s(e.code,e.desc)},s.createInvalidCorrelationIdError=function(){return new s(i.ClientConfigurationErrorMessage.invalidCorrelationIdError.code,i.ClientConfigurationErrorMessage.invalidCorrelationIdError.desc)},s.createKnownAuthoritiesNotSetError=function(){return new s(i.ClientConfigurationErrorMessage.b2cKnownAuthoritiesNotSet.code,i.ClientConfigurationErrorMessage.b2cKnownAuthoritiesNotSet.desc)},s.createInvalidAuthorityTypeError=function(){return new s(i.ClientConfigurationErrorMessage.invalidAuthorityType.code,i.ClientConfigurationErrorMessage.invalidAuthorityType.desc)},s.createUntrustedAuthorityError=function(e){return new s(i.ClientConfigurationErrorMessage.untrustedAuthority.code,i.ClientConfigurationErrorMessage.untrustedAuthority.desc+" Provided Authority: "+e)},s.createTelemetryConfigError=function(r){var e=i.ClientConfigurationErrorMessage.telemetryConfigError,t=e.code,o=e.desc,n={applicationName:"string",applicationVersion:"string",telemetryEmitter:"function"};return new s(t,o+" mising values: "+Object.keys(n).reduce(function(e,t){return r[t]?e:e.concat([t+" ("+n[t]+")"])},[]).join(","))},s.createSsoSilentError=function(){return new s(i.ClientConfigurationErrorMessage.ssoSilentError.code,i.ClientConfigurationErrorMessage.ssoSilentError.desc)},s.createInvalidAuthorityMetadataError=function(){return new s(i.ClientConfigurationErrorMessage.invalidAuthorityMetadataError.code,i.ClientConfigurationErrorMessage.invalidAuthorityMetadataError.desc)},s);function s(e,t){var r=n.call(this,e,t)||this;return r.name="ClientConfigurationError",Object.setPrototypeOf(r,s.prototype),r}i.ClientConfigurationError=a},function(e,r,t){Object.defineProperty(r,"__esModule",{value:!0});var o=t(0),n=t(7),i=t(3);r.ClientAuthErrorMessage={multipleMatchingTokens:{code:"multiple_matching_tokens",desc:"The cache contains multiple tokens satisfying the requirements. Call AcquireToken again providing more requirements like authority."},multipleCacheAuthorities:{code:"multiple_authorities",desc:"Multiple authorities found in the cache. Pass authority in the API overload."},endpointResolutionError:{code:"endpoints_resolution_error",desc:"Error: could not resolve endpoints. Please check network and try again."},popUpWindowError:{code:"popup_window_error",desc:"Error opening popup window. This can happen if you are using IE or if popups are blocked in the browser."},tokenRenewalError:{code:"token_renewal_error",desc:"Token renewal operation failed due to timeout."},invalidIdToken:{code:"invalid_id_token",desc:"Invalid ID token format."},invalidStateError:{code:"invalid_state_error",desc:"Invalid state."},nonceMismatchError:{code:"nonce_mismatch_error",desc:"Nonce is not matching, Nonce received: "},loginProgressError:{code:"login_progress_error",desc:"Login_In_Progress: Error during login call - login is already in progress."},acquireTokenProgressError:{code:"acquiretoken_progress_error",desc:"AcquireToken_In_Progress: Error during login call - login is already in progress."},userCancelledError:{code:"user_cancelled",desc:"User cancelled the flow."},callbackError:{code:"callback_error",desc:"Error occurred in token received callback function."},userLoginRequiredError:{code:"user_login_error",desc:"User login is required. For silent calls, request must contain either sid or login_hint"},userDoesNotExistError:{code:"user_non_existent",desc:"User object does not exist. Please call a login API."},clientInfoDecodingError:{code:"client_info_decoding_error",desc:"The client info could not be parsed/decoded correctly. Please review the trace to determine the root cause."},clientInfoNotPopulatedError:{code:"client_info_not_populated_error",desc:"The service did not populate client_info in the response, Please verify with the service team"},nullOrEmptyIdToken:{code:"null_or_empty_id_token",desc:"The idToken is null or empty. Please review the trace to determine the root cause."},idTokenNotParsed:{code:"id_token_parsing_error",desc:"ID token cannot be parsed. Please review stack trace to determine root cause."},tokenEncodingError:{code:"token_encoding_error",desc:"The token to be decoded is not encoded correctly."},invalidInteractionType:{code:"invalid_interaction_type",desc:"The interaction type passed to the handler was incorrect or unknown"},cacheParseError:{code:"cannot_parse_cache",desc:"The cached token key is not a valid JSON and cannot be parsed"},blockTokenRequestsInHiddenIframe:{code:"block_token_requests",desc:"Token calls are blocked in hidden iframes"}};var a,s=(a=n.AuthError,o.__extends(c,a),c.createEndpointResolutionError=function(e){var t=r.ClientAuthErrorMessage.endpointResolutionError.desc;return e&&!i.StringUtils.isEmpty(e)&&(t+=" Details: "+e),new c(r.ClientAuthErrorMessage.endpointResolutionError.code,t)},c.createMultipleMatchingTokensInCacheError=function(e){return new c(r.ClientAuthErrorMessage.multipleMatchingTokens.code,"Cache error for scope "+e+": "+r.ClientAuthErrorMessage.multipleMatchingTokens.desc+".")},c.createMultipleAuthoritiesInCacheError=function(e){return new c(r.ClientAuthErrorMessage.multipleCacheAuthorities.code,"Cache error for scope "+e+": "+r.ClientAuthErrorMessage.multipleCacheAuthorities.desc+".")},c.createPopupWindowError=function(e){var t=r.ClientAuthErrorMessage.popUpWindowError.desc;return e&&!i.StringUtils.isEmpty(e)&&(t+=" Details: "+e),new c(r.ClientAuthErrorMessage.popUpWindowError.code,t)},c.createTokenRenewalTimeoutError=function(){return new c(r.ClientAuthErrorMessage.tokenRenewalError.code,r.ClientAuthErrorMessage.tokenRenewalError.desc)},c.createInvalidIdTokenError=function(e){return new c(r.ClientAuthErrorMessage.invalidIdToken.code,r.ClientAuthErrorMessage.invalidIdToken.desc+" Given token: "+e)},c.createInvalidStateError=function(e,t){return new c(r.ClientAuthErrorMessage.invalidStateError.code,r.ClientAuthErrorMessage.invalidStateError.desc+" "+e+", state expected : "+t+".")},c.createNonceMismatchError=function(e,t){return new c(r.ClientAuthErrorMessage.nonceMismatchError.code,r.ClientAuthErrorMessage.nonceMismatchError.desc+" "+e+", nonce expected : "+t+".")},c.createLoginInProgressError=function(){return new c(r.ClientAuthErrorMessage.loginProgressError.code,r.ClientAuthErrorMessage.loginProgressError.desc)},c.createAcquireTokenInProgressError=function(){return new c(r.ClientAuthErrorMessage.acquireTokenProgressError.code,r.ClientAuthErrorMessage.acquireTokenProgressError.desc)},c.createUserCancelledError=function(){return new c(r.ClientAuthErrorMessage.userCancelledError.code,r.ClientAuthErrorMessage.userCancelledError.desc)},c.createErrorInCallbackFunction=function(e){return new c(r.ClientAuthErrorMessage.callbackError.code,r.ClientAuthErrorMessage.callbackError.desc+" "+e+".")},c.createUserLoginRequiredError=function(){return new c(r.ClientAuthErrorMessage.userLoginRequiredError.code,r.ClientAuthErrorMessage.userLoginRequiredError.desc)},c.createUserDoesNotExistError=function(){return new c(r.ClientAuthErrorMessage.userDoesNotExistError.code,r.ClientAuthErrorMessage.userDoesNotExistError.desc)},c.createClientInfoDecodingError=function(e){return new c(r.ClientAuthErrorMessage.clientInfoDecodingError.code,r.ClientAuthErrorMessage.clientInfoDecodingError.desc+" Failed with error: "+e)},c.createClientInfoNotPopulatedError=function(e){return new c(r.ClientAuthErrorMessage.clientInfoNotPopulatedError.code,r.ClientAuthErrorMessage.clientInfoNotPopulatedError.desc+" Failed with error: "+e)},c.createIdTokenNullOrEmptyError=function(e){return new c(r.ClientAuthErrorMessage.nullOrEmptyIdToken.code,r.ClientAuthErrorMessage.nullOrEmptyIdToken.desc+" Raw ID Token Value: "+e)},c.createIdTokenParsingError=function(e){return new c(r.ClientAuthErrorMessage.idTokenNotParsed.code,r.ClientAuthErrorMessage.idTokenNotParsed.desc+" Failed with error: "+e)},c.createTokenEncodingError=function(e){return new c(r.ClientAuthErrorMessage.tokenEncodingError.code,r.ClientAuthErrorMessage.tokenEncodingError.desc+" Attempted to decode: "+e)},c.createInvalidInteractionTypeError=function(){return new c(r.ClientAuthErrorMessage.invalidInteractionType.code,r.ClientAuthErrorMessage.invalidInteractionType.desc)},c.createCacheParseError=function(e){var t="invalid key: "+e+", "+r.ClientAuthErrorMessage.cacheParseError.desc;return new c(r.ClientAuthErrorMessage.cacheParseError.code,t)},c.createBlockTokenRequestsInHiddenIframeError=function(){return new c(r.ClientAuthErrorMessage.blockTokenRequestsInHiddenIframe.code,r.ClientAuthErrorMessage.blockTokenRequestsInHiddenIframe.desc)},c);function c(e,t){var r=a.call(this,e,t)||this;return r.name="ClientAuthError",Object.setPrototypeOf(r,c.prototype),r}r.ClientAuthError=s},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var o=r(0);t.AuthErrorMessage={unexpectedError:{code:"unexpected_error",desc:"Unexpected error in authentication."},noWindowObjectError:{code:"no_window_object",desc:"No window object available. Details:"}};var n,i=(n=Error,o.__extends(a,n),a.createUnexpectedError=function(e){return new a(t.AuthErrorMessage.unexpectedError.code,t.AuthErrorMessage.unexpectedError.desc+": "+e)},a.createNoWindowObjectError=function(e){return new a(t.AuthErrorMessage.noWindowObjectError.code,t.AuthErrorMessage.noWindowObjectError.desc+" "+e)},a);function a(e,t){var r=n.call(this,t)||this;return Object.setPrototypeOf(r,a.prototype),r.errorCode=e,r.errorMessage=t,r.name="AuthError",r}t.AuthError=i},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0}),t.EVENT_NAME_PREFIX="msal.",t.EVENT_NAME_KEY="event_name",t.START_TIME_KEY="start_time",t.ELAPSED_TIME_KEY="elapsed_time",t.TELEMETRY_BLOB_EVENT_NAMES={MsalCorrelationIdConstStrKey:"Microsoft.MSAL.correlation_id",ApiTelemIdConstStrKey:"msal.api_telem_id",ApiIdConstStrKey:"msal.api_id",BrokerAppConstStrKey:"Microsoft_MSAL_broker_app",CacheEventCountConstStrKey:"Microsoft_MSAL_cache_event_count",HttpEventCountTelemetryBatchKey:"Microsoft_MSAL_http_event_count",IdpConstStrKey:"Microsoft_MSAL_idp",IsSilentTelemetryBatchKey:"",IsSuccessfulConstStrKey:"Microsoft_MSAL_is_successful",ResponseTimeConstStrKey:"Microsoft_MSAL_response_time",TenantIdConstStrKey:"Microsoft_MSAL_tenant_id",UiEventCountTelemetryBatchKey:"Microsoft_MSAL_ui_event_count"},t.TENANT_PLACEHOLDER="<tenant>"},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var o=r(5),n=r(1),i=(a.isIntersectingScopes=function(e,t){for(var r=this.trimAndConvertArrayToLowerCase(e.slice()),o=this.trimAndConvertArrayToLowerCase(t.slice()),n=0;n<o.length;n++)if(-1<r.indexOf(o[n].toLowerCase()))return!0;return!1},a.containsScope=function(e,t){var r=this.trimAndConvertArrayToLowerCase(e.slice());return this.trimAndConvertArrayToLowerCase(t.slice()).every(function(e){return 0<=r.indexOf(e.toString().toLowerCase())})},a.trimAndConvertToLowerCase=function(e){return e.trim().toLowerCase()},a.trimAndConvertArrayToLowerCase=function(e){var t=this;return e.map(function(e){return t.trimAndConvertToLowerCase(e)})},a.removeElement=function(e,t){var r=this.trimAndConvertToLowerCase(t);return e.filter(function(e){return e!==r})},a.parseScope=function(e){var t="";if(e)for(var r=0;r<e.length;++r)t+=r!==e.length-1?e[r]+" ":e[r];return t},a.validateInputScope=function(e,t,r){if(e){if(!Array.isArray(e))throw o.ClientConfigurationError.createScopesNonArrayError(e);if(e.length<1)throw o.ClientConfigurationError.createEmptyScopesArrayError(e.toString());if(-1<e.indexOf(r)&&1<e.length)throw o.ClientConfigurationError.createClientIdSingleScopeError(e.toString())}else if(t)throw o.ClientConfigurationError.createScopesRequiredError(e)},a.getScopeFromState=function(e){if(e){var t=e.indexOf(n.Constants.resourceDelimiter);if(-1<t&&t+1<e.length)return e.substring(t+1)}return""},a.appendScopes=function(e,t){if(e){var r=t?this.trimAndConvertArrayToLowerCase(t.slice()):null,o=this.trimAndConvertArrayToLowerCase(e.slice());return r?o.concat(r):o}return null},a);function a(){}t.ScopeSet=i},function(e,o,t){Object.defineProperty(o,"__esModule",{value:!0});var n=t(8),r=t(2),i=t(4),a=t(21);o.scrubTenantFromUri=function(e){var t=i.UrlUtils.GetUrlComponents(e);if(a.AuthorityFactory.isAdfs(e))return e;var r=t.PathSegments;if(r&&2<=r.length){var o="tfp"===r[1]?2:1;o<r.length&&(r[o]=n.TENANT_PLACEHOLDER)}return t.Protocol+"//"+t.HostNameAndPort+"/"+r.join("/")},o.hashPersonalIdentifier=function(e){return r.CryptoUtils.base64Encode(e)},o.prependEventNamePrefix=function(e){return""+n.EVENT_NAME_PREFIX+(e||"")},o.supportsBrowserPerformance=function(){return!!("undefined"!=typeof window&&"performance"in window&&window.performance.mark&&window.performance.measure)},o.endBrowserPerformanceMeasurement=function(e,t,r){o.supportsBrowserPerformance()&&(window.performance.mark(r),window.performance.measure(e,t,r),window.performance.clearMeasures(e),window.performance.clearMarks(t),window.performance.clearMarks(r))},o.startBrowserPerformanceMeasurement=function(e){o.supportsBrowserPerformance()&&window.performance.mark(e)}},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var o=(n.parseExpiresIn=function(e){return e=e||"3599",parseInt(e,10)},n.now=function(){return Math.round((new Date).getTime()/1e3)},n.relativeNowMs=function(){return window.performance.now()},n);function n(){}t.TimeUtils=o},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var c,o,i=r(3),a=r(1);(o=c=t.LogLevel||(t.LogLevel={}))[o.Error=0]="Error",o[o.Warning=1]="Warning",o[o.Info=2]="Info",o[o.Verbose=3]="Verbose";var n=(s.prototype.logMessage=function(e,t,r){if(!(e>this.level||!this.piiLoggingEnabled&&r)){var o,n=(new Date).toUTCString();o=i.StringUtils.isEmpty(this.correlationId)?n+":"+a.libraryVersion()+"-"+c[e]+(r?"-pii":"")+" "+t:n+":"+this.correlationId+"-"+a.libraryVersion()+"-"+c[e]+(r?"-pii":"")+" "+t,this.executeCallback(e,o,r)}},s.prototype.executeCallback=function(e,t,r){this.localCallback&&this.localCallback(e,t,r)},s.prototype.error=function(e){this.logMessage(c.Error,e,!1)},s.prototype.errorPii=function(e){this.logMessage(c.Error,e,!0)},s.prototype.warning=function(e){this.logMessage(c.Warning,e,!1)},s.prototype.warningPii=function(e){this.logMessage(c.Warning,e,!0)},s.prototype.info=function(e){this.logMessage(c.Info,e,!1)},s.prototype.infoPii=function(e){this.logMessage(c.Info,e,!0)},s.prototype.verbose=function(e){this.logMessage(c.Verbose,e,!1)},s.prototype.verbosePii=function(e){this.logMessage(c.Verbose,e,!0)},s.prototype.isPiiLoggingEnabled=function(){return this.piiLoggingEnabled},s);function s(e,t){void 0===t&&(t={}),this.level=c.Info;var r=t.correlationId,o=void 0===r?"":r,n=t.level,i=void 0===n?c.Info:n,a=t.piiLoggingEnabled,s=void 0!==a&&a;this.localCallback=e,this.correlationId=o,this.level=i,this.piiLoggingEnabled=s}t.Logger=n},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var o=r(0),n=r(7);t.ServerErrorMessage={serverUnavailable:{code:"server_unavailable",desc:"Server is temporarily unavailable."},unknownServerError:{code:"unknown_server_error"}};var i,a=(i=n.AuthError,o.__extends(s,i),s.createServerUnavailableError=function(){return new s(t.ServerErrorMessage.serverUnavailable.code,t.ServerErrorMessage.serverUnavailable.desc)},s.createUnknownServerError=function(e){return new s(t.ServerErrorMessage.unknownServerError.code,e)},s);function s(e,t){var r=i.call(this,e,t)||this;return r.name="ServerError",Object.setPrototypeOf(r,s.prototype),r}t.ServerError=a},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var o=r(0),n=r(8),i=r(8),a=r(10),s=r(2),c=(u.prototype.setElapsedTime=function(e){this.event[a.prependEventNamePrefix(i.ELAPSED_TIME_KEY)]=e},u.prototype.stop=function(){this.setElapsedTime(+Date.now()-+this.startTimestamp),a.endBrowserPerformanceMeasurement(this.displayName,this.perfStartMark,this.perfEndMark)},u.prototype.start=function(){this.startTimestamp=Date.now(),this.event[a.prependEventNamePrefix(i.START_TIME_KEY)]=this.startTimestamp,a.startBrowserPerformanceMeasurement(this.perfStartMark)},Object.defineProperty(u.prototype,"telemetryCorrelationId",{get:function(){return this.event[""+n.TELEMETRY_BLOB_EVENT_NAMES.MsalCorrelationIdConstStrKey]},set:function(e){this.event[""+n.TELEMETRY_BLOB_EVENT_NAMES.MsalCorrelationIdConstStrKey]=e},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"eventName",{get:function(){return this.event[a.prependEventNamePrefix(i.EVENT_NAME_KEY)]},enumerable:!0,configurable:!0}),u.prototype.get=function(){return o.__assign({},this.event,{eventId:this.eventId})},Object.defineProperty(u.prototype,"key",{get:function(){return this.telemetryCorrelationId+"_"+this.eventId+"-"+this.eventName},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"displayName",{get:function(){return"Msal-"+this.label+"-"+this.telemetryCorrelationId},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"perfStartMark",{get:function(){return"start-"+this.key},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"perfEndMark",{get:function(){return"end-"+this.key},enumerable:!0,configurable:!0}),u);function u(e,t,r){var o;this.eventId=s.CryptoUtils.createNewGuid(),this.label=r,this.event=((o={})[a.prependEventNamePrefix(i.EVENT_NAME_KEY)]=e,o[a.prependEventNamePrefix(i.ELAPSED_TIME_KEY)]=-1,o[""+n.TELEMETRY_BLOB_EVENT_NAMES.MsalCorrelationIdConstStrKey]=t,o)}t.default=c},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var m=r(0),v=r(31),E=r(32),C=r(16),T=r(33),b=r(34),I=r(35),S=r(19),_=r(9),A=r(3),w=r(20),y=r(17),P=r(11),R=r(4),k=r(18),O=r(38),N=r(21),o=r(25),a=r(5),U=r(7),M=r(6),K=r(13),x=r(26),q=r(27),s=m.__importDefault(r(39)),c=r(28),L=r(1),u=r(2),n=r(24),H="id_token",i="token",l="id_token token",h=(Object.defineProperty(d.prototype,"authority",{get:function(){return this.authorityInstance.CanonicalAuthority},set:function(e){this.authorityInstance=N.AuthorityFactory.CreateInstance(e,this.config.auth.validateAuthority)},enumerable:!0,configurable:!0}),d.prototype.getAuthorityInstance=function(){return this.authorityInstance},d.prototype.handleRedirectCallback=function(e,t){if(!e)throw a.ClientConfigurationError.createInvalidCallbackObjectError(e);t?(this.tokenReceivedCallback=e,this.errorReceivedCallback=t,this.logger.warning("This overload for callback is deprecated - please change the format of the callbacks to a single callback as shown: (err: AuthError, response: AuthResponse).")):this.authResponseCallback=e,this.redirectError?this.authErrorHandler(L.Constants.interactionTypeRedirect,this.redirectError,this.redirectResponse):this.redirectResponse&&this.authResponseHandler(L.Constants.interactionTypeRedirect,this.redirectResponse)},d.prototype.urlContainsHash=function(e){return this.logger.verbose("UrlContainsHash has been called"),R.UrlUtils.urlContainsHash(e)},d.prototype.authResponseHandler=function(e,t,r){if(this.logger.verbose("AuthResponseHandler has been called"),e===L.Constants.interactionTypeRedirect)this.logger.verbose("Interaction type is redirect"),this.errorReceivedCallback?(this.logger.verbose("Two callbacks were provided to handleRedirectCallback, calling success callback with response"),this.tokenReceivedCallback(t)):this.authResponseCallback&&(this.logger.verbose("One callback was provided to handleRedirectCallback, calling authResponseCallback with response"),this.authResponseCallback(null,t));else{if(e!==L.Constants.interactionTypePopup)throw M.ClientAuthError.createInvalidInteractionTypeError();this.logger.verbose("Interaction type is popup, resolving"),r(t)}},d.prototype.authErrorHandler=function(e,t,r,o){if(this.logger.verbose("AuthErrorHandler has been called"),this.cacheStorage.removeItem(L.TemporaryCacheKeys.INTERACTION_STATUS),e===L.Constants.interactionTypeRedirect)this.logger.verbose("Interaction type is redirect"),this.errorReceivedCallback?(this.logger.verbose("Two callbacks were provided to handleRedirectCallback, calling error callback"),this.errorReceivedCallback(t,r.accountState)):(this.logger.verbose("One callback was provided to handleRedirectCallback, calling authResponseCallback with error"),this.authResponseCallback(t,r));else{if(e!==L.Constants.interactionTypePopup)throw M.ClientAuthError.createInvalidInteractionTypeError();this.logger.verbose("Interaction type is popup, rejecting"),o(t)}},d.prototype.loginRedirect=function(e){this.logger.verbose("LoginRedirect has been called");var t=k.RequestUtils.validateRequest(e,!0,this.clientId,L.Constants.interactionTypeRedirect);this.acquireTokenInteractive(L.Constants.interactionTypeRedirect,!0,t,null,null)},d.prototype.acquireTokenRedirect=function(e){this.logger.verbose("AcquireTokenRedirect has been called");var t=k.RequestUtils.validateRequest(e,!1,this.clientId,L.Constants.interactionTypeRedirect);this.acquireTokenInteractive(L.Constants.interactionTypeRedirect,!1,t,null,null)},d.prototype.loginPopup=function(e){var r=this;this.logger.verbose("LoginPopup has been called");var o=k.RequestUtils.validateRequest(e,!0,this.clientId,L.Constants.interactionTypePopup),t=this.telemetryManager.createAndStartApiEvent(o.correlationId,c.API_EVENT_IDENTIFIER.LoginPopup);return new Promise(function(e,t){r.acquireTokenInteractive(L.Constants.interactionTypePopup,!0,o,e,t)}).then(function(e){return r.logger.verbose("Successfully logged in"),r.telemetryManager.stopAndFlushApiEvent(o.correlationId,t,!0),e}).catch(function(e){throw r.cacheStorage.resetTempCacheItems(o.state),r.telemetryManager.stopAndFlushApiEvent(o.correlationId,t,!1,e.errorCode),e})},d.prototype.acquireTokenPopup=function(e){var r=this;this.logger.verbose("AcquireTokenPopup has been called");var o=k.RequestUtils.validateRequest(e,!1,this.clientId,L.Constants.interactionTypePopup),t=this.telemetryManager.createAndStartApiEvent(o.correlationId,c.API_EVENT_IDENTIFIER.AcquireTokenPopup);return new Promise(function(e,t){r.acquireTokenInteractive(L.Constants.interactionTypePopup,!1,o,e,t)}).then(function(e){return r.logger.verbose("Successfully acquired token"),r.telemetryManager.stopAndFlushApiEvent(o.correlationId,t,!0),e}).catch(function(e){throw r.cacheStorage.resetTempCacheItems(o.state),r.telemetryManager.stopAndFlushApiEvent(o.correlationId,t,!1,e.errorCode),e})},d.prototype.acquireTokenInteractive=function(t,r,o,n,i){var a=this;this.logger.verbose("AcquireTokenInteractive has been called"),w.WindowUtils.blockReloadInHiddenIframes();var e,s=this.cacheStorage.getItem(L.TemporaryCacheKeys.INTERACTION_STATUS);if(t===L.Constants.interactionTypeRedirect&&this.cacheStorage.setItem(L.TemporaryCacheKeys.REDIRECT_REQUEST,""+L.Constants.inProgress+L.Constants.resourceDelimiter+o.state),s===L.Constants.inProgress){var c=r?M.ClientAuthError.createLoginInProgressError():M.ClientAuthError.createAcquireTokenInProgressError(),u=q.buildResponseStateOnly(this.getAccountState(o.state));return this.cacheStorage.resetTempCacheItems(o.state),void this.authErrorHandler(t,c,u,i)}if(o&&o.account&&!r?(e=o.account,this.logger.verbose("Account set from request")):(e=this.getAccount(),this.logger.verbose("Account set from MSAL Cache")),e||C.ServerRequestParameters.isSSOParam(o))this.logger.verbose("User session exists, login not required"),this.acquireTokenHelper(e,t,r,o,n,i);else{if(!r)return this.logger.verbose("AcquireToken call, no context or account given"),this.logger.info("User login is required"),u=q.buildResponseStateOnly(this.getAccountState(o.state)),this.cacheStorage.resetTempCacheItems(o.state),void this.authErrorHandler(t,M.ClientAuthError.createUserLoginRequiredError(),u,i);if(this.extractADALIdToken()&&!o.scopes){this.logger.info("ADAL's idToken exists. Extracting login information from ADAL's idToken");var l=this.buildIDTokenRequest(o);this.silentLogin=!0,this.acquireTokenSilent(l).then(function(e){a.silentLogin=!1,a.logger.info("Unified cache call is successful"),a.authResponseHandler(t,e,n)},function(e){a.silentLogin=!1,a.logger.error("Error occurred during unified cache ATS: "+e),a.acquireTokenHelper(null,t,r,o,n,i)})}else this.logger.verbose("Login call but no token found, proceed to login"),this.acquireTokenHelper(null,t,r,o,n,i)}},d.prototype.acquireTokenHelper=function(h,d,p,g,f,y){return m.__awaiter(this,void 0,Promise,function(){var t,r,o,n,i,a,s,c,u,l;return m.__generator(this,function(e){switch(e.label){case 0:this.logger.verbose("AcquireTokenHelper has been called"),this.logger.verbose("Interaction type: "+d+". isLoginCall: "+p),this.cacheStorage.setItem(L.TemporaryCacheKeys.INTERACTION_STATUS,L.Constants.inProgress),t=g.scopes?g.scopes.join(" ").toLowerCase():this.clientId.toLowerCase(),this.logger.verbosePii("Serialized scopes: "+t),o=g&&g.authority?N.AuthorityFactory.CreateInstance(g.authority,this.config.auth.validateAuthority,g.authorityMetadata):this.authorityInstance,e.label=1;case 1:return e.trys.push([1,11,,12]),o.hasCachedMetadata()?[3,3]:(this.logger.verbose("No cached metadata for authority"),[4,N.AuthorityFactory.saveMetadataFromNetwork(o,this.telemetryManager,g.correlationId)]);case 2:return e.sent(),[3,4];case 3:this.logger.verbose("Cached metadata found for authority"),e.label=4;case 4:if(i=p?H:this.getTokenType(h,g.scopes,!1),a=g.redirectStartPage||window.location.href,r=new C.ServerRequestParameters(o,this.clientId,i,this.getRedirectUri(g&&g.redirectUri),g.scopes,g.state,g.correlationId),this.logger.verbose("Finished building server authentication request"),this.updateCacheEntries(r,h,p,a),this.logger.verbose("Updating cache entries"),r.populateQueryParams(h,g),this.logger.verbose("Query parameters populated from account"),s=R.UrlUtils.createNavigateUrl(r)+L.Constants.response_mode_fragment,d===L.Constants.interactionTypeRedirect)p?this.logger.verbose("Interaction type redirect but login call is true. State not cached"):(this.cacheStorage.setItem(I.AuthCache.generateTemporaryCacheKey(L.TemporaryCacheKeys.STATE_ACQ_TOKEN,g.state),r.state,this.inCookie),this.logger.verbose("State cached for redirect"),this.logger.verbosePii("State cached: "+r.state));else{if(d!==L.Constants.interactionTypePopup)throw this.logger.verbose("Invalid interaction error. State not cached"),M.ClientAuthError.createInvalidInteractionTypeError();window.renewStates.push(r.state),window.requestType=p?L.Constants.login:L.Constants.renewToken,this.logger.verbose("State saved to window"),this.logger.verbosePii("State saved: "+r.state),this.registerCallback(r.state,t,f,y)}if(d!==L.Constants.interactionTypePopup)return[3,9];this.logger.verbose("Interaction type is popup. Generating popup window");try{n=this.openPopup(s,"msal",L.Constants.popUpWidth,L.Constants.popUpHeight),w.WindowUtils.trackPopup(n)}catch(e){if(this.logger.info(M.ClientAuthErrorMessage.popUpWindowError.code+":"+M.ClientAuthErrorMessage.popUpWindowError.desc),this.cacheStorage.setItem(L.ErrorCacheKeys.ERROR,M.ClientAuthErrorMessage.popUpWindowError.code),this.cacheStorage.setItem(L.ErrorCacheKeys.ERROR_DESC,M.ClientAuthErrorMessage.popUpWindowError.desc),y)return y(M.ClientAuthError.createPopupWindowError()),[2]}if(!n)return[3,8];e.label=5;case 5:return e.trys.push([5,7,,8]),[4,w.WindowUtils.monitorPopupForHash(n,this.config.system.loadFrameTimeout,s,this.logger)];case 6:return c=e.sent(),this.handleAuthenticationResponse(c),this.cacheStorage.removeItem(L.TemporaryCacheKeys.INTERACTION_STATUS),this.logger.info("Closing popup window"),this.config.framework.isAngular&&(this.broadcast("msal:popUpHashChanged",c),w.WindowUtils.closePopups()),[3,8];case 7:return u=e.sent(),y&&y(u),this.config.framework.isAngular?this.broadcast("msal:popUpClosed",u.errorCode+L.Constants.resourceDelimiter+u.errorMessage):(this.cacheStorage.removeItem(L.TemporaryCacheKeys.INTERACTION_STATUS),n.close()),[3,8];case 8:return[3,10];case 9:g.onRedirectNavigate?(this.logger.verbose("Invoking onRedirectNavigate callback"),!1!==g.onRedirectNavigate(s)?(this.logger.verbose("onRedirectNavigate did not return false, navigating"),this.navigateWindow(s)):this.logger.verbose("onRedirectNavigate returned false, stopping navigation")):(this.logger.verbose("Navigating window to urlNavigate"),this.navigateWindow(s)),e.label=10;case 10:return[3,12];case 11:return l=e.sent(),this.logger.error(l),this.cacheStorage.resetTempCacheItems(g.state),this.authErrorHandler(d,M.ClientAuthError.createEndpointResolutionError(l.toString),q.buildResponseStateOnly(g.state),y),n&&n.close(),[3,12];case 12:return[2]}})})},d.prototype.ssoSilent=function(e){if(this.logger.verbose("ssoSilent has been called"),!e)throw a.ClientConfigurationError.createEmptyRequestError();if(!e.sid&&!e.loginHint)throw a.ClientConfigurationError.createSsoSilentError();return this.acquireTokenSilent(m.__assign({},e,{scopes:[this.clientId]}))},d.prototype.acquireTokenSilent=function(e){var t=this;this.logger.verbose("AcquireTokenSilent has been called");var g=k.RequestUtils.validateRequest(e,!1,this.clientId,L.Constants.interactionTypeSilent),r=this.telemetryManager.createAndStartApiEvent(g.correlationId,c.API_EVENT_IDENTIFIER.AcquireTokenSilent),f=k.RequestUtils.createRequestSignature(g);return new Promise(function(d,p){return m.__awaiter(t,void 0,void 0,function(){var t,r,o,n,i,a,s,c,u,l,h;return m.__generator(this,function(e){switch(e.label){case 0:if(w.WindowUtils.blockReloadInHiddenIframes(),t=g.scopes.join(" ").toLowerCase(),this.logger.verbosePii("Serialized scopes: "+t),g.account?(r=g.account,this.logger.verbose("Account set from request")):(r=this.getAccount(),this.logger.verbose("Account set from MSAL Cache")),o=this.cacheStorage.getItem(L.Constants.adalIdToken),!r&&!g.sid&&!g.loginHint&&A.StringUtils.isEmpty(o))return this.logger.info("User login is required"),[2,p(M.ClientAuthError.createUserLoginRequiredError())];if(n=this.getTokenType(r,g.scopes,!0),this.logger.verbose("Response type: "+n),i=new C.ServerRequestParameters(N.AuthorityFactory.CreateInstance(g.authority,this.config.auth.validateAuthority,g.authorityMetadata),this.clientId,n,this.getRedirectUri(g.redirectUri),g.scopes,g.state,g.correlationId),this.logger.verbose("Finished building server authentication request"),C.ServerRequestParameters.isSSOParam(g)||r?(i.populateQueryParams(r,g,null,!0),this.logger.verbose("Query parameters populated from existing SSO or account")):r||A.StringUtils.isEmpty(o)?this.logger.verbose("No additional query parameters added"):(a=y.TokenUtils.extractIdToken(o),this.logger.verbose("ADAL's idToken exists. Extracting login information from ADAL's idToken to populate query parameters"),i.populateQueryParams(r,null,a,!0)),!(s=g.claimsRequest||i.claimsValue)&&!g.forceRefresh)try{u=this.getCachedToken(i,r)}catch(e){c=e}return u?(this.logger.verbose("Token found in cache lookup"),this.logger.verbosePii("Scopes found: "+JSON.stringify(u.scopes)),d(u),[2,null]):[3,1];case 1:return c?(this.logger.infoPii(c.errorCode+":"+c.errorMessage),p(c),[2,null]):[3,2];case 2:l=void 0,l=s?"Skipped cache lookup since claims were given":g.forceRefresh?"Skipped cache lookup since request.forceRefresh option was set to true":"No token found in cache lookup",this.logger.verbose(l),i.authorityInstance||(i.authorityInstance=g.authority?N.AuthorityFactory.CreateInstance(g.authority,this.config.auth.validateAuthority,g.authorityMetadata):this.authorityInstance),this.logger.verbosePii("Authority instance: "+i.authority),e.label=3;case 3:return e.trys.push([3,7,,8]),i.authorityInstance.hasCachedMetadata()?[3,5]:(this.logger.verbose("No cached metadata for authority"),[4,N.AuthorityFactory.saveMetadataFromNetwork(i.authorityInstance,this.telemetryManager,g.correlationId)]);case 4:return e.sent(),this.logger.verbose("Authority has been updated with endpoint discovery response"),[3,6];case 5:this.logger.verbose("Cached metadata found for authority"),e.label=6;case 6:return window.activeRenewals[f]?(this.logger.verbose("Renewing token in progress. Registering callback"),this.registerCallback(window.activeRenewals[f],f,d,p)):g.scopes&&-1<g.scopes.indexOf(this.clientId)&&1===g.scopes.length?(this.logger.verbose("ClientId is the only scope, renewing idToken"),this.silentLogin=!0,this.renewIdToken(f,d,p,r,i)):(this.logger.verbose("Renewing access token"),this.renewToken(f,d,p,r,i)),[3,8];case 7:return h=e.sent(),this.logger.error(h),p(M.ClientAuthError.createEndpointResolutionError(h.toString())),[2,null];case 8:return[2]}})})}).then(function(e){return t.logger.verbose("Successfully acquired token"),t.telemetryManager.stopAndFlushApiEvent(g.correlationId,r,!0),e}).catch(function(e){throw t.cacheStorage.resetTempCacheItems(g.state),t.telemetryManager.stopAndFlushApiEvent(g.correlationId,r,!1,e.errorCode),e})},d.prototype.openPopup=function(e,t,r,o){this.logger.verbose("OpenPopup has been called");try{var n=window.screenLeft?window.screenLeft:window.screenX,i=window.screenTop?window.screenTop:window.screenY,a=window.innerWidth||document.documentElement.clientWidth||document.body.clientWidth,s=window.innerHeight||document.documentElement.clientHeight||document.body.clientHeight,c=a/2-r/2+n,u=s/2-o/2+i,l=window.open(e,t,"width="+r+", height="+o+", top="+u+", left="+c+", scrollbars=yes");if(!l)throw M.ClientAuthError.createPopupWindowError();return l.focus&&l.focus(),l}catch(e){throw this.cacheStorage.removeItem(L.TemporaryCacheKeys.INTERACTION_STATUS),M.ClientAuthError.createPopupWindowError(e.toString())}},d.prototype.loadIframeTimeout=function(a,s,c){return m.__awaiter(this,void 0,Promise,function(){var t,r,o,n,i;return m.__generator(this,function(e){switch(e.label){case 0:return t=window.activeRenewals[c],this.logger.verbosePii("Set loading state to pending for: "+c+":"+t),this.cacheStorage.setItem(I.AuthCache.generateTemporaryCacheKey(L.TemporaryCacheKeys.RENEW_STATUS,t),L.Constants.inProgress),this.config.system.navigateFrameWait?[4,w.WindowUtils.loadFrame(a,s,this.config.system.navigateFrameWait,this.logger)]:[3,2];case 1:return o=e.sent(),[3,3];case 2:o=w.WindowUtils.loadFrameSync(a,s,this.logger),e.label=3;case 3:r=o,e.label=4;case 4:return e.trys.push([4,6,,7]),[4,w.WindowUtils.monitorIframeForHash(r.contentWindow,this.config.system.loadFrameTimeout,a,this.logger)];case 5:return(n=e.sent())&&this.handleAuthenticationResponse(n),[3,7];case 6:throw i=e.sent(),this.cacheStorage.getItem(I.AuthCache.generateTemporaryCacheKey(L.TemporaryCacheKeys.RENEW_STATUS,t))===L.Constants.inProgress&&(this.logger.verbose("Loading frame has timed out after: "+this.config.system.loadFrameTimeout/1e3+" seconds for scope/authority "+c+":"+t),t&&window.callbackMappedToRenewStates[t]&&window.callbackMappedToRenewStates[t](null,i),this.cacheStorage.removeItem(I.AuthCache.generateTemporaryCacheKey(L.TemporaryCacheKeys.RENEW_STATUS,t))),w.WindowUtils.removeHiddenIframe(r),i;case 7:return w.WindowUtils.removeHiddenIframe(r),[2]}})})},d.prototype.navigateWindow=function(e,t){if(!e||A.StringUtils.isEmpty(e))throw this.logger.info("Navigate url is empty"),U.AuthError.createUnexpectedError("Navigate url is empty");var r=t||window,o=t?"Navigated Popup window to:"+e:"Navigate to:"+e;this.logger.infoPii(o),r.location.assign(e)},d.prototype.registerCallback=function(o,n,e,t){var i=this;window.activeRenewals[n]=o,window.promiseMappedToRenewStates[o]||(window.promiseMappedToRenewStates[o]=[]),window.promiseMappedToRenewStates[o].push({resolve:e,reject:t}),window.callbackMappedToRenewStates[o]||(window.callbackMappedToRenewStates[o]=function(e,t){window.activeRenewals[n]=null;for(var r=0;r<window.promiseMappedToRenewStates[o].length;++r)try{if(t)window.promiseMappedToRenewStates[o][r].reject(t);else{if(!e)throw i.cacheStorage.resetTempCacheItems(o),U.AuthError.createUnexpectedError("Error and response are both null");window.promiseMappedToRenewStates[o][r].resolve(e)}}catch(e){i.logger.warning(e)}window.promiseMappedToRenewStates[o]=null,window.callbackMappedToRenewStates[o]=null})},d.prototype.logout=function(e){this.logger.verbose("Logout has been called"),this.logoutAsync(e)},d.prototype.logoutAsync=function(s){return m.__awaiter(this,void 0,Promise,function(){var t,r,o,n,i,a;return m.__generator(this,function(e){switch(e.label){case 0:t=s||u.CryptoUtils.createNewGuid(),r=this.telemetryManager.createAndStartApiEvent(t,c.API_EVENT_IDENTIFIER.Logout),this.clearCache(),this.account=null,e.label=1;case 1:return e.trys.push([1,5,,6]),this.authorityInstance.hasCachedMetadata()?[3,3]:(this.logger.verbose("No cached metadata for authority"),[4,N.AuthorityFactory.saveMetadataFromNetwork(this.authorityInstance,this.telemetryManager,s)]);case 2:return e.sent(),[3,4];case 3:this.logger.verbose("Cached metadata found for authority"),e.label=4;case 4:return o="client-request-id="+t,n=void 0,this.getPostLogoutRedirectUri()?(n="&post_logout_redirect_uri="+encodeURIComponent(this.getPostLogoutRedirectUri()),this.logger.verbose("redirectUri found and set")):(n="",this.logger.verbose("No redirectUri set for app. postLogoutQueryParam is empty")),i=void 0,this.authorityInstance.EndSessionEndpoint?(i=this.authorityInstance.EndSessionEndpoint+"?"+o+n,this.logger.verbose("EndSessionEndpoint found and urlNavigate set"),this.logger.verbosePii("urlNavigate set to: "+this.authorityInstance.EndSessionEndpoint)):(i=this.authority+"oauth2/v2.0/logout?"+o+n,this.logger.verbose("No endpoint, urlNavigate set to default")),this.telemetryManager.stopAndFlushApiEvent(t,r,!0),this.logger.verbose("Navigating window to urlNavigate"),this.navigateWindow(i),[3,6];case 5:return a=e.sent(),this.telemetryManager.stopAndFlushApiEvent(t,r,!1,a.errorCode),[3,6];case 6:return[2]}})})},d.prototype.clearCache=function(){this.logger.verbose("Clearing cache"),window.renewStates=[];for(var e=this.cacheStorage.getAllAccessTokens(L.Constants.clientId,L.Constants.homeAccountIdentifier),t=0;t<e.length;t++)this.cacheStorage.removeItem(JSON.stringify(e[t].key));this.cacheStorage.resetCacheItems(),this.cacheStorage.clearMsalCookie(),this.logger.verbose("Cache cleared")},d.prototype.clearCacheForScope=function(e){this.logger.verbose("Clearing access token from cache");for(var t=this.cacheStorage.getAllAccessTokens(L.Constants.clientId,L.Constants.homeAccountIdentifier),r=0;r<t.length;r++){var o=t[r];o.value.accessToken===e&&(this.cacheStorage.removeItem(JSON.stringify(o.key)),this.logger.verbosePii("Access token removed: "+o.key))}},d.prototype.isCallback=function(e){return this.logger.info("isCallback will be deprecated in favor of urlContainsHash in MSAL.js v2.0."),this.logger.verbose("isCallback has been called"),R.UrlUtils.urlContainsHash(e)},d.prototype.processCallBack=function(e,t,r){var o,n;this.logger.info("ProcessCallBack has been called. Processing callback from redirect response"),t||(this.logger.verbose("StateInfo is null, getting stateInfo from hash"),t=this.getResponseState(e));try{o=this.saveTokenFromHash(e,t)}catch(e){n=e}try{this.cacheStorage.clearMsalCookie(t.state);var i=this.getAccountState(t.state);if(o){if(t.requestType===L.Constants.renewToken||o.accessToken?(window.parent!==window?this.logger.verbose("Window is in iframe, acquiring token silently"):this.logger.verbose("Acquiring token interactive in progress"),this.logger.verbose("Response tokenType set to "+L.ServerHashParamKeys.ACCESS_TOKEN),o.tokenType=L.ServerHashParamKeys.ACCESS_TOKEN):t.requestType===L.Constants.login&&(this.logger.verbose("Response tokenType set to "+L.ServerHashParamKeys.ID_TOKEN),o.tokenType=L.ServerHashParamKeys.ID_TOKEN),!r)return this.logger.verbose("Setting redirectResponse"),void(this.redirectResponse=o)}else if(!r)return this.logger.verbose("Response is null, setting redirectResponse with state"),this.redirectResponse=q.buildResponseStateOnly(i),this.redirectError=n,void this.cacheStorage.resetTempCacheItems(t.state);this.logger.verbose("Calling callback provided to processCallback"),r(o,n)}catch(e){throw this.logger.error("Error occurred in token received callback function: "+e),M.ClientAuthError.createErrorInCallbackFunction(e.toString())}},d.prototype.handleAuthenticationResponse=function(e){this.logger.verbose("HandleAuthenticationResponse has been called");var t=e||window.location.hash,r=this.getResponseState(t);this.logger.verbose("Obtained state from response");var o=window.callbackMappedToRenewStates[r.state];this.processCallBack(t,r,o),w.WindowUtils.closePopups()},d.prototype.handleRedirectAuthenticationResponse=function(e){this.logger.info("Returned from redirect url"),this.logger.verbose("HandleRedirectAuthenticationResponse has been called"),window.location.hash="",this.logger.verbose("Window.location.hash cleared");var t=this.getResponseState(e);if(this.config.auth.navigateToLoginRequestUrl&&window.parent===window){this.logger.verbose("Window.parent is equal to window, not in popup or iframe. Navigation to login request url after login turned on");var r=this.cacheStorage.getItem(I.AuthCache.generateTemporaryCacheKey(L.TemporaryCacheKeys.LOGIN_REQUEST,t.state),this.inCookie);if(!r||"null"===r)return this.logger.error("Unable to get valid login request url from cache, redirecting to home page"),void window.location.assign("/");this.logger.verbose("Valid login request url obtained from cache");var o=R.UrlUtils.removeHashFromUrl(window.location.href),n=R.UrlUtils.removeHashFromUrl(r);if(o!==n)return this.logger.verbose("Current url is not login request url, navigating"),this.logger.verbosePii("CurrentUrl: "+o+", finalRedirectUrl: "+n),void window.location.assign(""+n+e);this.logger.verbose("Current url matches login request url");var i=R.UrlUtils.GetUrlComponents(r);i.Hash&&(this.logger.verbose("Login request url contains hash, resetting non-msal hash"),window.location.hash=i.Hash)}else this.config.auth.navigateToLoginRequestUrl||this.logger.verbose("Default navigation to start page after login turned off");this.processCallBack(e,t,null)},d.prototype.getResponseState=function(e){this.logger.verbose("GetResponseState has been called");var t,r=R.UrlUtils.deserializeHash(e);if(!r)throw U.AuthError.createUnexpectedError("Hash was not parsed correctly.");if(!r.hasOwnProperty(L.ServerHashParamKeys.STATE))throw U.AuthError.createUnexpectedError("Hash does not contain state.");this.logger.verbose("Hash contains state. Creating stateInfo object");var o=k.RequestUtils.parseLibraryState(r.state);if((t={requestType:L.Constants.unknown,state:r.state,timestamp:o.ts,method:o.method,stateMatch:!1}).state===this.cacheStorage.getItem(I.AuthCache.generateTemporaryCacheKey(L.TemporaryCacheKeys.STATE_LOGIN,t.state),this.inCookie)||t.state===this.silentAuthenticationState)return this.logger.verbose("State matches cached state, setting requestType to login"),t.requestType=L.Constants.login,t.stateMatch=!0,t;if(t.state===this.cacheStorage.getItem(I.AuthCache.generateTemporaryCacheKey(L.TemporaryCacheKeys.STATE_ACQ_TOKEN,t.state),this.inCookie))return this.logger.verbose("State matches cached state, setting requestType to renewToken"),t.requestType=L.Constants.renewToken,t.stateMatch=!0,t;if(!t.stateMatch){this.logger.verbose("State does not match cached state, setting requestType to type from window"),t.requestType=window.requestType;for(var n=window.renewStates,i=0;i<n.length;i++)if(n[i]===t.state){this.logger.verbose("Matching state found for request"),t.stateMatch=!0;break}t.stateMatch||this.logger.verbose("Matching state not found for request")}return t},d.prototype.getCachedToken=function(e,t){this.logger.verbose("GetCachedToken has been called");var r=null,o=e.scopes,n=this.cacheStorage.getAllAccessTokens(this.clientId,t?t.homeAccountIdentifier:null);if(this.logger.verbose("Getting all cached access tokens"),0===n.length)return this.logger.verbose("No matching tokens found when filtered by clientId and account"),null;var i=[];if(e.authority){for(this.logger.verbose("Authority passed, filtering by authority and scope"),a=0;a<n.length;a++)c=(s=n[a]).key.scopes.split(" "),_.ScopeSet.containsScope(c,o)&&R.UrlUtils.CanonicalizeUri(s.key.authority)===e.authority&&i.push(s);if(0===i.length)return this.logger.verbose("No matching tokens found"),null;if(1!==i.length)throw M.ClientAuthError.createMultipleMatchingTokensInCacheError(o.toString());this.logger.verbose("Single token found"),r=i[0]}else{this.logger.verbose("No authority passed, filtering tokens by scope");for(var a=0;a<n.length;a++){var s,c=(s=n[a]).key.scopes.split(" ");_.ScopeSet.containsScope(c,o)&&i.push(s)}if(1===i.length)this.logger.verbose("One matching token found, setting authorityInstance"),r=i[0],e.authorityInstance=N.AuthorityFactory.CreateInstance(r.key.authority,this.config.auth.validateAuthority);else{if(1<i.length)throw M.ClientAuthError.createMultipleMatchingTokensInCacheError(o.toString());this.logger.verbose("No matching token found when filtering by scope");var u=this.getUniqueAuthority(n,"authority");if(1<u.length)throw M.ClientAuthError.createMultipleAuthoritiesInCacheError(o.toString());this.logger.verbose("Single authority used, setting authorityInstance"),e.authorityInstance=N.AuthorityFactory.CreateInstance(u[0],this.config.auth.validateAuthority)}}if(null==r)return this.logger.verbose("No tokens found"),null;this.logger.verbose("Evaluating access token found");var l=Number(r.value.expiresIn),h=this.config.system.tokenRenewalOffsetSeconds||300;if(l&&l>P.TimeUtils.now()+h){this.logger.verbose("Token expiration is within offset, renewing token");var d=new b.IdToken(r.value.idToken);if(!t&&!(t=this.getAccount()))throw U.AuthError.createUnexpectedError("Account should not be null here.");var p=this.getAccountState(e.state),g={uniqueId:"",tenantId:"",tokenType:r.value.idToken===r.value.accessToken?L.ServerHashParamKeys.ID_TOKEN:L.ServerHashParamKeys.ACCESS_TOKEN,idToken:d,idTokenClaims:d.claims,accessToken:r.value.accessToken,scopes:r.key.scopes.split(" "),expiresOn:new Date(1e3*l),account:t,accountState:p,fromCache:!0};return O.ResponseUtils.setResponseIdToken(g,d),this.logger.verbose("Response generated and token set"),g}return this.logger.verbose("Token expired, removing from cache"),this.cacheStorage.removeItem(JSON.stringify(i[0].key)),null},d.prototype.getUniqueAuthority=function(e,t){this.logger.verbose("GetUniqueAuthority has been called");var r=[],o=[];return e.forEach(function(e){e.key.hasOwnProperty(t)&&-1===o.indexOf(e.key[t])&&(o.push(e.key[t]),r.push(e.key[t]))}),r},d.prototype.extractADALIdToken=function(){this.logger.verbose("ExtractADALIdToken has been called");var e=this.cacheStorage.getItem(L.Constants.adalIdToken);return A.StringUtils.isEmpty(e)?null:y.TokenUtils.extractIdToken(e)},d.prototype.renewToken=function(e,t,r,o,n){this.logger.verbose("RenewToken has been called"),this.logger.verbosePii("RenewToken scope and authority: "+e);var i=w.WindowUtils.generateFrameName(L.FramePrefix.TOKEN_FRAME,e);w.WindowUtils.addHiddenIFrame(i,this.logger),this.updateCacheEntries(n,o,!1),this.logger.verbosePii("RenewToken expected state: "+n.state);var a=R.UrlUtils.urlRemoveQueryStringParameter(R.UrlUtils.createNavigateUrl(n),L.Constants.prompt)+L.Constants.prompt_none+L.Constants.response_mode_fragment;window.renewStates.push(n.state),window.requestType=L.Constants.renewToken,this.logger.verbose("Set window.renewState and requestType"),this.registerCallback(n.state,e,t,r),this.logger.infoPii("Navigate to: "+a),this.loadIframeTimeout(a,i,e).catch(function(e){return r(e)})},d.prototype.renewIdToken=function(e,t,r,o,n){this.logger.info("RenewIdToken has been called");var i=w.WindowUtils.generateFrameName(L.FramePrefix.ID_TOKEN_FRAME,e);w.WindowUtils.addHiddenIFrame(i,this.logger),this.updateCacheEntries(n,o,!1),this.logger.verbose("RenewIdToken expected state: "+n.state);var a=R.UrlUtils.urlRemoveQueryStringParameter(R.UrlUtils.createNavigateUrl(n),L.Constants.prompt)+L.Constants.prompt_none+L.Constants.response_mode_fragment;this.silentLogin?(this.logger.verbose("Silent login is true, set silentAuthenticationState"),window.requestType=L.Constants.login,this.silentAuthenticationState=n.state):(this.logger.verbose("Not silent login, set window.renewState and requestType"),window.requestType=L.Constants.renewToken,window.renewStates.push(n.state)),this.registerCallback(n.state,e,t,r),this.logger.infoPii('Navigate to:" '+a),this.loadIframeTimeout(a,i,e).catch(function(e){return r(e)})},d.prototype.saveAccessToken=function(e,t,r,o,n){var i;this.logger.verbose("SaveAccessToken has been called");var a,s=m.__assign({},e),c=new T.ClientInfo(o);if(r.hasOwnProperty(L.ServerHashParamKeys.SCOPE)){this.logger.verbose("Response parameters contains scope");var u=(i=r[L.ServerHashParamKeys.SCOPE]).split(" "),l=this.cacheStorage.getAllAccessTokens(this.clientId,t);this.logger.verbose("Retrieving all access tokens from cache and removing duplicates");for(var h=0;h<l.length;h++){var d=l[h];if(d.key.homeAccountIdentifier===e.account.homeAccountIdentifier){var p=d.key.scopes.split(" ");_.ScopeSet.isIntersectingScopes(p,u)&&this.cacheStorage.removeItem(JSON.stringify(d.key))}}var g=P.TimeUtils.parseExpiresIn(r[L.ServerHashParamKeys.EXPIRES_IN]);a=k.RequestUtils.parseLibraryState(r[L.ServerHashParamKeys.STATE]).ts+g;var f=new v.AccessTokenKey(t,this.clientId,i,c.uid,c.utid),y=new E.AccessTokenValue(r[L.ServerHashParamKeys.ACCESS_TOKEN],n.rawIdToken,a.toString(),o);this.cacheStorage.setItem(JSON.stringify(f),JSON.stringify(y)),this.logger.verbose("Saving token to cache"),s.accessToken=r[L.ServerHashParamKeys.ACCESS_TOKEN],s.scopes=u}else this.logger.verbose("Response parameters does not contain scope, clientId set as scope"),i=this.clientId,f=new v.AccessTokenKey(t,this.clientId,i,c.uid,c.utid),a=Number(n.expiration),y=new E.AccessTokenValue(r[L.ServerHashParamKeys.ID_TOKEN],r[L.ServerHashParamKeys.ID_TOKEN],a.toString(),o),this.cacheStorage.setItem(JSON.stringify(f),JSON.stringify(y)),this.logger.verbose("Saving token to cache"),s.scopes=[i],s.accessToken=r[L.ServerHashParamKeys.ID_TOKEN];return a?(this.logger.verbose("New expiration set"),s.expiresOn=new Date(1e3*a)):this.logger.error("Could not parse expiresIn parameter"),s},d.prototype.saveTokenFromHash=function(e,t){this.logger.verbose("SaveTokenFromHash has been called"),this.logger.info("State status: "+t.stateMatch+"; Request type: "+t.requestType);var r,o={uniqueId:"",tenantId:"",tokenType:"",idToken:null,idTokenClaims:null,accessToken:null,scopes:[],expiresOn:null,account:null,accountState:"",fromCache:!1},n=R.UrlUtils.deserializeHash(e),i="",a="",s=null;if(n.hasOwnProperty(L.ServerHashParamKeys.ERROR_DESCRIPTION)||n.hasOwnProperty(L.ServerHashParamKeys.ERROR)){if(this.logger.verbose("Server returned an error"),this.logger.infoPii("Error : "+n[L.ServerHashParamKeys.ERROR]+"; Error description: "+n[L.ServerHashParamKeys.ERROR_DESCRIPTION]),this.cacheStorage.setItem(L.ErrorCacheKeys.ERROR,n[L.ServerHashParamKeys.ERROR]),this.cacheStorage.setItem(L.ErrorCacheKeys.ERROR_DESC,n[L.ServerHashParamKeys.ERROR_DESCRIPTION]),t.requestType===L.Constants.login&&(this.logger.verbose("RequestType is login, caching login error, generating authorityKey"),this.cacheStorage.setItem(L.ErrorCacheKeys.LOGIN_ERROR,n[L.ServerHashParamKeys.ERROR_DESCRIPTION]+":"+n[L.ServerHashParamKeys.ERROR]),i=I.AuthCache.generateAuthorityKey(t.state)),t.requestType===L.Constants.renewToken){this.logger.verbose("RequestType is renewToken, generating acquireTokenAccountKey"),i=I.AuthCache.generateAuthorityKey(t.state);var c=this.getAccount(),u=void 0;c&&!A.StringUtils.isEmpty(c.homeAccountIdentifier)?(u=c.homeAccountIdentifier,this.logger.verbose("AccountId is set")):(u=L.Constants.no_account,this.logger.verbose("AccountId is set as no_account")),a=I.AuthCache.generateAcquireTokenAccountKey(u,t.state)}var l=n[L.ServerHashParamKeys.ERROR],h=n[L.ServerHashParamKeys.ERROR_DESCRIPTION];r=x.InteractionRequiredAuthError.isInteractionRequiredError(l)||x.InteractionRequiredAuthError.isInteractionRequiredError(h)?new x.InteractionRequiredAuthError(n[L.ServerHashParamKeys.ERROR],n[L.ServerHashParamKeys.ERROR_DESCRIPTION]):new K.ServerError(n[L.ServerHashParamKeys.ERROR],n[L.ServerHashParamKeys.ERROR_DESCRIPTION])}else if(this.logger.verbose("Server returns success"),t.stateMatch){this.logger.info("State is right"),n.hasOwnProperty(L.ServerHashParamKeys.SESSION_STATE)&&(this.logger.verbose("Fragment has session state, caching"),this.cacheStorage.setItem(I.AuthCache.generateTemporaryCacheKey(L.TemporaryCacheKeys.SESSION_STATE,t.state),n[L.ServerHashParamKeys.SESSION_STATE])),o.accountState=this.getAccountState(t.state);var d="";if(n.hasOwnProperty(L.ServerHashParamKeys.ACCESS_TOKEN)){this.logger.info("Fragment has access token"),o.accessToken=n[L.ServerHashParamKeys.ACCESS_TOKEN],n.hasOwnProperty(L.ServerHashParamKeys.SCOPE)&&(o.scopes=n[L.ServerHashParamKeys.SCOPE].split(" ")),n.hasOwnProperty(L.ServerHashParamKeys.ID_TOKEN)?(this.logger.verbose("Fragment has id_token"),s=new b.IdToken(n[L.ServerHashParamKeys.ID_TOKEN]),o.idToken=s,o.idTokenClaims=s.claims):(this.logger.verbose("No idToken on fragment, getting idToken from cache"),s=new b.IdToken(this.cacheStorage.getItem(L.PersistentCacheKeys.IDTOKEN)),o=O.ResponseUtils.setResponseIdToken(o,s));var p=this.populateAuthority(t.state,this.inCookie,this.cacheStorage,s);if(this.logger.verbose("Got authority from cache"),!n.hasOwnProperty(L.ServerHashParamKeys.CLIENT_INFO))throw this.logger.warning("ClientInfo not received in the response from AAD"),M.ClientAuthError.createClientInfoNotPopulatedError("ClientInfo not received in the response from the server");this.logger.verbose("Fragment has clientInfo"),d=n[L.ServerHashParamKeys.CLIENT_INFO],o.account=S.Account.createAccount(s,new T.ClientInfo(d)),this.logger.verbose("Account object created from response");var g=void 0;g=o.account&&!A.StringUtils.isEmpty(o.account.homeAccountIdentifier)?(this.logger.verbose("AccountKey set"),o.account.homeAccountIdentifier):(this.logger.verbose("AccountKey set as no_account"),L.Constants.no_account),a=I.AuthCache.generateAcquireTokenAccountKey(g,t.state);var f=I.AuthCache.generateAcquireTokenAccountKey(L.Constants.no_account,t.state);this.logger.verbose("AcquireTokenAccountKey generated");var y=this.cacheStorage.getItem(a),m=void 0;A.StringUtils.isEmpty(y)?A.StringUtils.isEmpty(this.cacheStorage.getItem(f))||(this.logger.verbose("No acquireToken account retrieved from cache"),o=this.saveAccessToken(o,p,n,d,s)):(m=JSON.parse(y),this.logger.verbose("AcquireToken request account retrieved from cache"),o.account&&m&&S.Account.compareAccounts(o.account,m)?(o=this.saveAccessToken(o,p,n,d,s),this.logger.info("The user object received in the response is the same as the one passed in the acquireToken request")):this.logger.warning("The account object created from the response is not the same as the one passed in the acquireToken request"))}if(n.hasOwnProperty(L.ServerHashParamKeys.ID_TOKEN))if(this.logger.info("Fragment has idToken"),s=new b.IdToken(n[L.ServerHashParamKeys.ID_TOKEN]),o=O.ResponseUtils.setResponseIdToken(o,s),n.hasOwnProperty(L.ServerHashParamKeys.CLIENT_INFO)?(this.logger.verbose("Fragment has clientInfo"),d=n[L.ServerHashParamKeys.CLIENT_INFO]):this.logger.warning("ClientInfo not received in the response from AAD"),p=this.populateAuthority(t.state,this.inCookie,this.cacheStorage,s),this.account=S.Account.createAccount(s,new T.ClientInfo(d)),o.account=this.account,this.logger.verbose("Account object created from response"),s&&s.nonce){this.logger.verbose("IdToken has nonce");var v=this.cacheStorage.getItem(I.AuthCache.generateTemporaryCacheKey(L.TemporaryCacheKeys.NONCE_IDTOKEN,t.state),this.inCookie);s.nonce!==v?(this.account=null,this.cacheStorage.setItem(L.ErrorCacheKeys.LOGIN_ERROR,"Nonce Mismatch. Expected Nonce: "+v+",Actual Nonce: "+s.nonce),this.logger.error("Nonce Mismatch. Expected Nonce: "+v+", Actual Nonce: "+s.nonce),r=M.ClientAuthError.createNonceMismatchError(v,s.nonce)):(this.logger.verbose("Nonce matches, saving idToken to cache"),this.cacheStorage.setItem(L.PersistentCacheKeys.IDTOKEN,n[L.ServerHashParamKeys.ID_TOKEN],this.inCookie),this.cacheStorage.setItem(L.PersistentCacheKeys.CLIENT_INFO,d,this.inCookie),this.saveAccessToken(o,p,n,d,s))}else this.logger.verbose("No idToken or no nonce. Cache key for Authority set as state"),i=t.state,a=t.state,this.logger.error("Invalid id_token received in the response"),r=M.ClientAuthError.createInvalidIdTokenError(s),this.cacheStorage.setItem(L.ErrorCacheKeys.ERROR,r.errorCode),this.cacheStorage.setItem(L.ErrorCacheKeys.ERROR_DESC,r.errorMessage)}else{this.logger.verbose("State mismatch"),i=t.state,a=t.state;var E=this.cacheStorage.getItem(I.AuthCache.generateTemporaryCacheKey(L.TemporaryCacheKeys.STATE_LOGIN,t.state),this.inCookie);this.logger.error("State Mismatch. Expected State: "+E+", Actual State: "+t.state),r=M.ClientAuthError.createInvalidStateError(t.state,E),this.cacheStorage.setItem(L.ErrorCacheKeys.ERROR,r.errorCode),this.cacheStorage.setItem(L.ErrorCacheKeys.ERROR_DESC,r.errorMessage)}if(this.cacheStorage.removeItem(I.AuthCache.generateTemporaryCacheKey(L.TemporaryCacheKeys.RENEW_STATUS,t.state)),this.cacheStorage.resetTempCacheItems(t.state),this.logger.verbose("Status set to complete, temporary cache cleared"),this.inCookie&&(this.logger.verbose("InCookie is true, setting authorityKey in cookie"),this.cacheStorage.setItemCookie(i,"",-1),this.cacheStorage.clearMsalCookie(t.state)),r)throw r;if(!o)throw U.AuthError.createUnexpectedError("Response is null");return o},d.prototype.populateAuthority=function(e,t,r,o){this.logger.verbose("PopulateAuthority has been called");var n=I.AuthCache.generateAuthorityKey(e),i=r.getItem(n,t);return A.StringUtils.isEmpty(i)?i:R.UrlUtils.replaceTenantPath(i,o.tenantId)},d.prototype.getAccount=function(){if(this.account)return this.account;var e=this.cacheStorage.getItem(L.PersistentCacheKeys.IDTOKEN,this.inCookie),t=this.cacheStorage.getItem(L.PersistentCacheKeys.CLIENT_INFO,this.inCookie);if(A.StringUtils.isEmpty(e)||A.StringUtils.isEmpty(t))return null;var r=new b.IdToken(e),o=new T.ClientInfo(t);return this.account=S.Account.createAccount(r,o),this.account},d.prototype.getAccountState=function(e){if(e){var t=e.indexOf(L.Constants.resourceDelimiter);if(-1<t&&t+1<e.length)return e.substring(t+1)}return e},d.prototype.getAllAccounts=function(){for(var e=[],t=this.cacheStorage.getAllAccessTokens(L.Constants.clientId,L.Constants.homeAccountIdentifier),r=0;r<t.length;r++){var o=new b.IdToken(t[r].value.idToken),n=new T.ClientInfo(t[r].value.homeAccountIdentifier),i=S.Account.createAccount(o,n);e.push(i)}return this.getUniqueAccounts(e)},d.prototype.getUniqueAccounts=function(e){if(!e||e.length<=1)return e;for(var t=[],r=[],o=0;o<e.length;++o)e[o].homeAccountIdentifier&&-1===t.indexOf(e[o].homeAccountIdentifier)&&(t.push(e[o].homeAccountIdentifier),r.push(e[o]));return r},d.prototype.broadcast=function(e,t){var r=new CustomEvent(e,{detail:t});window.dispatchEvent(r)},d.prototype.getCachedTokenInternal=function(e,t,r,o){var n=t||this.getAccount();if(!n)return null;var i=this.authorityInstance?this.authorityInstance:N.AuthorityFactory.CreateInstance(this.authority,this.config.auth.validateAuthority),a=this.getTokenType(n,e,!0),s=new C.ServerRequestParameters(i,this.clientId,a,this.getRedirectUri(),e,r,o);return this.getCachedToken(s,t)},d.prototype.getScopesForEndpoint=function(e){if(0<this.config.framework.unprotectedResources.length)for(var t=0;t<this.config.framework.unprotectedResources.length;t++)if(-1<e.indexOf(this.config.framework.unprotectedResources[t]))return null;if(0<this.config.framework.protectedResourceMap.size)for(var r=0,o=Array.from(this.config.framework.protectedResourceMap.keys());r<o.length;r++){var n=o[r];if(-1<e.indexOf(n))return this.config.framework.protectedResourceMap.get(n)}return-1<e.indexOf("http://")||-1<e.indexOf("https://")?R.UrlUtils.getHostFromUri(e)===R.UrlUtils.getHostFromUri(this.getRedirectUri())?new Array(this.clientId):null:new Array(this.clientId)},d.prototype.getLoginInProgress=function(){return this.cacheStorage.getItem(L.TemporaryCacheKeys.INTERACTION_STATUS)===L.Constants.inProgress},d.prototype.setInteractionInProgress=function(e){e?this.cacheStorage.setItem(L.TemporaryCacheKeys.INTERACTION_STATUS,L.Constants.inProgress):this.cacheStorage.removeItem(L.TemporaryCacheKeys.INTERACTION_STATUS)},d.prototype.setloginInProgress=function(e){this.setInteractionInProgress(e)},d.prototype.getAcquireTokenInProgress=function(){return this.cacheStorage.getItem(L.TemporaryCacheKeys.INTERACTION_STATUS)===L.Constants.inProgress},d.prototype.setAcquireTokenInProgress=function(e){this.setInteractionInProgress(e)},d.prototype.getLogger=function(){return this.logger},d.prototype.setLogger=function(e){this.logger=e},d.prototype.getRedirectUri=function(e){return e||("function"==typeof this.config.auth.redirectUri?this.config.auth.redirectUri():this.config.auth.redirectUri)},d.prototype.getPostLogoutRedirectUri=function(){return"function"==typeof this.config.auth.postLogoutRedirectUri?this.config.auth.postLogoutRedirectUri():this.config.auth.postLogoutRedirectUri},d.prototype.getCurrentConfiguration=function(){if(!this.config)throw a.ClientConfigurationError.createNoSetConfigurationError();return this.config},d.prototype.getTokenType=function(e,t,r){return r?S.Account.compareAccounts(e,this.getAccount())?-1<t.indexOf(this.config.auth.clientId)?H:i:-1<t.indexOf(this.config.auth.clientId)?H:l:S.Account.compareAccounts(e,this.getAccount())?-1<t.indexOf(this.clientId)?H:i:l},d.prototype.setAccountCache=function(e,t){var r=e?this.getAccountId(e):L.Constants.no_account,o=I.AuthCache.generateAcquireTokenAccountKey(r,t);this.cacheStorage.setItem(o,JSON.stringify(e))},d.prototype.setAuthorityCache=function(e,t){var r=I.AuthCache.generateAuthorityKey(e);this.cacheStorage.setItem(r,R.UrlUtils.CanonicalizeUri(t),this.inCookie)},d.prototype.updateCacheEntries=function(e,t,r,o){o&&this.cacheStorage.setItem(I.AuthCache.generateTemporaryCacheKey(L.TemporaryCacheKeys.LOGIN_REQUEST,e.state),o,this.inCookie),r?this.cacheStorage.setItem(I.AuthCache.generateTemporaryCacheKey(L.TemporaryCacheKeys.STATE_LOGIN,e.state),e.state,this.inCookie):this.setAccountCache(t,e.state),this.setAuthorityCache(e.state,e.authority),this.cacheStorage.setItem(I.AuthCache.generateTemporaryCacheKey(L.TemporaryCacheKeys.NONCE_IDTOKEN,e.state),e.nonce,this.inCookie)},d.prototype.getAccountId=function(e){return A.StringUtils.isEmpty(e.homeAccountIdentifier)?L.Constants.no_account:e.homeAccountIdentifier},d.prototype.buildIDTokenRequest=function(e){return{scopes:[this.clientId],authority:this.authority,account:this.getAccount(),extraQueryParameters:e.extraQueryParameters,correlationId:e.correlationId}},d.prototype.getTelemetryManagerFromConfig=function(e,t){if(!e)return s.default.getTelemetrymanagerStub(t,this.logger);var r=e.applicationName,o=e.applicationVersion,n=e.telemetryEmitter;if(!r||!o||!n)throw a.ClientConfigurationError.createTelemetryConfigError(e);var i={platform:{applicationName:r,applicationVersion:o},clientId:t};return new s.default(i,n,this.logger)},d);function d(e){this.authResponseCallback=null,this.tokenReceivedCallback=null,this.errorReceivedCallback=null,this.config=o.buildConfiguration(e),this.logger=this.config.system.logger,this.clientId=this.config.auth.clientId,this.inCookie=this.config.cache.storeAuthStateInCookie,this.telemetryManager=this.getTelemetryManagerFromConfig(this.config.system.telemetry,this.clientId),n.TrustedAuthority.setTrustedAuthoritiesFromConfig(this.config.auth.validateAuthority,this.config.auth.knownAuthorities),N.AuthorityFactory.saveMetadataFromConfig(this.config.auth.authority,this.config.auth.authorityMetadata),this.authority=this.config.auth.authority||"https://login.microsoftonline.com/common",this.cacheStorage=new I.AuthCache(this.clientId,this.config.cache.cacheLocation,this.inCookie),window.activeRenewals={},window.renewStates=[],window.callbackMappedToRenewStates={},window.promiseMappedToRenewStates={},window.msal=this;var t=window.location.hash,r=R.UrlUtils.urlContainsHash(t);w.WindowUtils.checkIfBackButtonIsPressed(this.cacheStorage),!r||this.getResponseState(t).method===L.Constants.interactionTypeRedirect&&this.handleRedirectAuthenticationResponse(t)}t.UserAgentApplication=h},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var s=r(2),c=r(1),o=r(3),u=r(9),n=(Object.defineProperty(a.prototype,"authority",{get:function(){return this.authorityInstance?this.authorityInstance.CanonicalAuthority:null},enumerable:!0,configurable:!0}),a.prototype.populateQueryParams=function(e,t,r,o){var n={};t&&(t.prompt&&(this.promptValue=t.prompt),t.claimsRequest&&(this.claimsValue=t.claimsRequest),a.isSSOParam(t)&&(n=this.constructUnifiedCacheQueryParameter(t,null))),r&&(n=this.constructUnifiedCacheQueryParameter(null,r)),n=this.addHintParameters(e,n);var i=t?t.extraQueryParameters:null;this.queryParameters=a.generateQueryParametersString(n),this.extraQueryParameters=a.generateQueryParametersString(i,o)},a.prototype.constructUnifiedCacheQueryParameter=function(e,t){var r,o;if(e)if(e.account){var n=e.account;n.sid?(r=c.SSOTypes.SID,o=n.sid):n.userName&&(r=c.SSOTypes.LOGIN_HINT,o=n.userName)}else e.sid?(r=c.SSOTypes.SID,o=e.sid):e.loginHint&&(r=c.SSOTypes.LOGIN_HINT,o=e.loginHint);else t&&t.hasOwnProperty(c.Constants.upn)&&(r=c.SSOTypes.ID_TOKEN,o=t.upn);return this.addSSOParameter(r,o)},a.prototype.addHintParameters=function(e,t){return e&&!t[c.SSOTypes.SID]&&(!t[c.SSOTypes.LOGIN_HINT]&&e.sid&&this.promptValue===c.PromptState.NONE?t=this.addSSOParameter(c.SSOTypes.SID,e.sid,t):t[c.SSOTypes.LOGIN_HINT]||!e.userName||o.StringUtils.isEmpty(e.userName)||(t=this.addSSOParameter(c.SSOTypes.LOGIN_HINT,e.userName,t))),t},a.prototype.addSSOParameter=function(e,t,r){if(r=r||{},!t)return r;switch(e){case c.SSOTypes.SID:r[c.SSOTypes.SID]=t;break;case c.SSOTypes.ID_TOKEN:case c.SSOTypes.LOGIN_HINT:r[c.SSOTypes.LOGIN_HINT]=t}return r},a.generateQueryParametersString=function(t,r){var o=null;return t&&Object.keys(t).forEach(function(e){e===c.Constants.domain_hint&&(r||t[c.SSOTypes.SID])||(null==o?o=e+"="+encodeURIComponent(t[e]):o+="&"+e+"="+encodeURIComponent(t[e]))}),o},a.isSSOParam=function(e){return e&&(e.account||e.sid||e.loginHint)},a);function a(e,t,r,o,n,i,a){this.authorityInstance=e,this.clientId=t,this.nonce=s.CryptoUtils.createNewGuid(),this.scopes=n?n.slice():[t],this.scopes=u.ScopeSet.trimAndConvertArrayToLowerCase(this.scopes),this.state=i,this.correlationId=a,this.xClientSku="MSAL.JS",this.xClientVer=c.libraryVersion(),this.responseType=r,this.redirectUri=o}t.ServerRequestParameters=n},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var n=r(2),o=r(3),i=(a.decodeJwt=function(e){if(o.StringUtils.isEmpty(e))return null;var t=/^([^\.\s]*)\.([^\.\s]+)\.([^\.\s]*)$/.exec(e);return!t||t.length<4?null:{header:t[1],JWSPayload:t[2],JWSSig:t[3]}},a.extractIdToken=function(e){var t=this.decodeJwt(e);if(!t)return null;try{var r=t.JWSPayload,o=n.CryptoUtils.base64Decode(r);return o?JSON.parse(o):null}catch(e){}return null},a);function a(){}t.TokenUtils=i},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var c=r(0),o=r(1),u=r(5),l=r(9),n=r(3),i=r(2),a=r(11),s=r(6),h=(d.validateRequest=function(e,t,r,o){if(!t&&!e)throw u.ClientConfigurationError.createEmptyRequestError();var n,i;e&&(n=t?l.ScopeSet.appendScopes(e.scopes,e.extraScopesToConsent):e.scopes,l.ScopeSet.validateInputScope(n,!t,r),this.validatePromptParameter(e.prompt),i=this.validateEQParameters(e.extraQueryParameters,e.claimsRequest),this.validateClaimsRequest(e.claimsRequest));var a=this.validateAndGenerateState(e&&e.state,o),s=this.validateAndGenerateCorrelationId(e&&e.correlationId);return c.__assign({},e,{extraQueryParameters:i,scopes:n,state:a,correlationId:s})},d.validatePromptParameter=function(e){if(e&&[o.PromptState.LOGIN,o.PromptState.SELECT_ACCOUNT,o.PromptState.CONSENT,o.PromptState.NONE].indexOf(e)<0)throw u.ClientConfigurationError.createInvalidPromptError(e)},d.validateEQParameters=function(e,t){var r=c.__assign({},e);return r?(t&&delete r[o.Constants.claims],o.BlacklistedEQParams.forEach(function(e){r[e]&&delete r[e]}),r):null},d.validateClaimsRequest=function(e){if(e)try{JSON.parse(e)}catch(e){throw u.ClientConfigurationError.createClaimsRequestParsingError(e)}},d.validateAndGenerateState=function(e,t){return n.StringUtils.isEmpty(e)?d.generateLibraryState(t):""+d.generateLibraryState(t)+o.Constants.resourceDelimiter+e},d.generateLibraryState=function(e){var t={id:i.CryptoUtils.createNewGuid(),ts:a.TimeUtils.now(),method:e},r=JSON.stringify(t);return i.CryptoUtils.base64Encode(r)},d.parseLibraryState=function(t){var e=decodeURIComponent(t).split(o.Constants.resourceDelimiter)[0];if(i.CryptoUtils.isGuid(e))return{id:e,ts:a.TimeUtils.now(),method:o.Constants.interactionTypeRedirect};try{var r=i.CryptoUtils.base64Decode(e);return JSON.parse(r)}catch(e){throw s.ClientAuthError.createInvalidStateError(t,null)}},d.validateAndGenerateCorrelationId=function(e){if(e&&!i.CryptoUtils.isGuid(e))throw u.ClientConfigurationError.createInvalidCorrelationIdError();return i.CryptoUtils.isGuid(e)?e:i.CryptoUtils.createNewGuid()},d.createRequestSignature=function(e){return""+e.scopes.join(" ").toLowerCase()+o.Constants.resourceDelimiter+e.authority},d);function d(){}t.RequestUtils=h},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var a=r(2),s=r(3),o=(c.createAccount=function(e,t){var r,o=e.objectId||e.subject,n=t?t.uid:"",i=t?t.utid:"";return s.StringUtils.isEmpty(n)||s.StringUtils.isEmpty(i)||(r=a.CryptoUtils.base64Encode(n)+"."+a.CryptoUtils.base64Encode(i)),new c(o,r,e.preferredName,e.name,e.claims,e.sid,e.issuer)},c.compareAccounts=function(e,t){return!!(e&&t&&e.homeAccountIdentifier&&t.homeAccountIdentifier&&e.homeAccountIdentifier===t.homeAccountIdentifier)},c);function c(e,t,r,o,n,i,a){this.accountIdentifier=e,this.homeAccountIdentifier=t,this.userName=r,this.name=o,this.idToken=n,this.idTokenClaims=n,this.sid=i,this.environment=a}t.Account=o},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var u=r(6),l=r(4),n=r(1),c=r(11),o=(h.isInIframe=function(){return window.parent!==window},h.isInPopup=function(){return!(!window.opener||window.opener===window)},h.generateFrameName=function(e,t){return""+e+n.Constants.resourceDelimiter+t},h.monitorIframeForHash=function(i,e,a,s){return new Promise(function(t,r){var o=c.TimeUtils.relativeNowMs()+e;s.verbose("monitorWindowForIframe polling started");var n=setInterval(function(){if(c.TimeUtils.relativeNowMs()>o)return s.error("monitorIframeForHash unable to find hash in url, timing out"),s.errorPii("monitorIframeForHash polling timed out for url: "+a),clearInterval(n),void r(u.ClientAuthError.createTokenRenewalTimeoutError());var e;try{e=i.location.href}catch(e){}e&&l.UrlUtils.urlContainsHash(e)&&(s.verbose("monitorIframeForHash found url in hash"),clearInterval(n),t(i.location.hash))},h.POLLING_INTERVAL_MS)})},h.monitorPopupForHash=function(a,e,s,c){return new Promise(function(t,r){var o=e/h.POLLING_INTERVAL_MS,n=0;c.verbose("monitorWindowForHash polling started");var i=setInterval(function(){if(a.closed)return c.error("monitorWindowForHash window closed"),clearInterval(i),void r(u.ClientAuthError.createUserCancelledError());var e;try{e=a.location.href}catch(e){}e&&"about:blank"!==e&&(n++,e&&l.UrlUtils.urlContainsHash(e)?(c.verbose("monitorPopupForHash found url in hash"),clearInterval(i),t(a.location.hash)):o<n&&(c.error("monitorPopupForHash unable to find hash in url, timing out"),c.errorPii("monitorPopupForHash polling timed out for url: "+s),clearInterval(i),r(u.ClientAuthError.createTokenRenewalTimeoutError())))},h.POLLING_INTERVAL_MS)})},h.loadFrame=function(o,n,e,i){var a=this;return i.infoPii("LoadFrame: "+n),new Promise(function(t,r){setTimeout(function(){var e=a.loadFrameSync(o,n,i);e?t(e):r("Unable to load iframe with name: "+n)},e)})},h.loadFrameSync=function(e,t,r){var o=h.addHiddenIFrame(t,r);return o?(""!==o.src&&"about:blank"!==o.src||(o.src=e,r.infoPii("Frame Name : "+t+" Navigated to: "+e)),o):null},h.addHiddenIFrame=function(e,t){if(void 0===e)return null;t.infoPii("Add msal frame to document:"+e);var r=document.getElementById(e);if(!r){if(document.createElement&&document.documentElement&&-1===window.navigator.userAgent.indexOf("MSIE 5.0")){var o=document.createElement("iframe");o.setAttribute("id",e),o.setAttribute("aria-hidden","true"),o.style.visibility="hidden",o.style.position="absolute",o.style.width=o.style.height="0",o.style.border="0",o.setAttribute("sandbox","allow-scripts allow-same-origin allow-forms"),r=document.getElementsByTagName("body")[0].appendChild(o)}else document.body&&document.body.insertAdjacentHTML&&document.body.insertAdjacentHTML("beforeend","<iframe name='"+e+"' id='"+e+"' style='display:none'></iframe>");window.frames&&window.frames[e]&&(r=window.frames[e])}return r},h.removeHiddenIframe=function(e){document.body===e.parentNode&&document.body.removeChild(e)},h.getIframeWithHash=function(t){var r=document.getElementsByTagName("iframe");return Array.apply(null,Array(r.length)).map(function(e,t){return r.item(t)}).filter(function(e){try{return e.contentWindow.location.hash===t}catch(e){return!1}})[0]},h.getPopups=function(){return window.openedWindows||(window.openedWindows=[]),window.openedWindows},h.getPopUpWithHash=function(t){return h.getPopups().filter(function(e){try{return e.location.hash===t}catch(e){return!1}})[0]},h.trackPopup=function(e){h.getPopups().push(e)},h.closePopups=function(){h.getPopups().forEach(function(e){return e.close()})},h.blockReloadInHiddenIframes=function(){if(l.UrlUtils.urlContainsHash(window.location.hash)&&h.isInIframe())throw u.ClientAuthError.createBlockTokenRequestsInHiddenIframeError()},h.checkIfBackButtonIsPressed=function(e){var t=e.getItem(n.TemporaryCacheKeys.REDIRECT_REQUEST);if(t&&!l.UrlUtils.urlContainsHash(window.location.hash)){var r=t.split(n.Constants.resourceDelimiter),o=1<r.length?r[r.length-1]:null;e.resetTempCacheItems(o)}},h.POLLING_INTERVAL_MS=50,h);function h(){}t.WindowUtils=o},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var i=r(0),o=r(22),n=r(3),a=r(5),s=r(1),c=r(4),u=(l.saveMetadataFromNetwork=function(r,o,n){return i.__awaiter(this,void 0,Promise,function(){var t;return i.__generator(this,function(e){switch(e.label){case 0:return[4,r.resolveEndpointsAsync(o,n)];case 1:return t=e.sent(),this.metadataMap.set(r.CanonicalAuthority,t),[2,t]}})})},l.getMetadata=function(e){return this.metadataMap.get(e)},l.saveMetadataFromConfig=function(e,t){try{if(t){var r=JSON.parse(t);if(!r.authorization_endpoint||!r.end_session_endpoint||!r.issuer)throw a.ClientConfigurationError.createInvalidAuthorityMetadataError();this.metadataMap.set(e,{AuthorizationEndpoint:r.authorization_endpoint,EndSessionEndpoint:r.end_session_endpoint,Issuer:r.issuer})}}catch(e){throw a.ClientConfigurationError.createInvalidAuthorityMetadataError()}},l.CreateInstance=function(e,t,r){return n.StringUtils.isEmpty(e)?null:(r&&this.saveMetadataFromConfig(e,r),new o.Authority(e,t,this.metadataMap.get(e)))},l.isAdfs=function(e){var t=c.UrlUtils.GetUrlComponents(e).PathSegments;return!(!t.length||t[0].toLowerCase()!==s.Constants.ADFS)},l.metadataMap=new Map,l);function l(){}t.AuthorityFactory=u},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var o,a=r(0),s=r(5),c=r(23),n=r(4),u=r(24),l=r(1);(o=t.AuthorityType||(t.AuthorityType={}))[o.Default=0]="Default",o[o.Adfs=1]="Adfs";var i=(Object.defineProperty(h.prototype,"Tenant",{get:function(){return this.CanonicalAuthorityUrlComponents.PathSegments[0]},enumerable:!0,configurable:!0}),Object.defineProperty(h.prototype,"AuthorizationEndpoint",{get:function(){return this.validateResolved(),this.tenantDiscoveryResponse.AuthorizationEndpoint.replace(/{tenant}|{tenantid}/g,this.Tenant)},enumerable:!0,configurable:!0}),Object.defineProperty(h.prototype,"EndSessionEndpoint",{get:function(){return this.validateResolved(),this.tenantDiscoveryResponse.EndSessionEndpoint.replace(/{tenant}|{tenantid}/g,this.Tenant)},enumerable:!0,configurable:!0}),Object.defineProperty(h.prototype,"SelfSignedJwtAudience",{get:function(){return this.validateResolved(),this.tenantDiscoveryResponse.Issuer.replace(/{tenant}|{tenantid}/g,this.Tenant)},enumerable:!0,configurable:!0}),h.prototype.validateResolved=function(){if(!this.hasCachedMetadata())throw"Please call ResolveEndpointsAsync first"},Object.defineProperty(h.prototype,"CanonicalAuthority",{get:function(){return this.canonicalAuthority},set:function(e){this.canonicalAuthority=n.UrlUtils.CanonicalizeUri(e),this.canonicalAuthorityUrlComponents=null},enumerable:!0,configurable:!0}),Object.defineProperty(h.prototype,"CanonicalAuthorityUrlComponents",{get:function(){return this.canonicalAuthorityUrlComponents||(this.canonicalAuthorityUrlComponents=n.UrlUtils.GetUrlComponents(this.CanonicalAuthority)),this.canonicalAuthorityUrlComponents},enumerable:!0,configurable:!0}),Object.defineProperty(h.prototype,"DefaultOpenIdConfigurationEndpoint",{get:function(){return this.CanonicalAuthority+"v2.0/.well-known/openid-configuration"},enumerable:!0,configurable:!0}),h.prototype.validateAsUri=function(){var e;try{e=this.CanonicalAuthorityUrlComponents}catch(e){throw s.ClientConfigurationErrorMessage.invalidAuthorityType}if(!e.Protocol||"https:"!==e.Protocol.toLowerCase())throw s.ClientConfigurationErrorMessage.authorityUriInsecure;if(!e.PathSegments||e.PathSegments.length<1)throw s.ClientConfigurationErrorMessage.authorityUriInvalidPath},h.prototype.DiscoverEndpoints=function(e,t,r){var o=new c.XhrClient,n=l.NetworkRequestType.GET,i=t.createAndStartHttpEvent(r,n,e,"openIdConfigurationEndpoint");return o.sendRequestAsync(e,n,!0).then(function(e){return i.httpResponseStatus=e.statusCode,t.stopEvent(i),{AuthorizationEndpoint:e.body.authorization_endpoint,EndSessionEndpoint:e.body.end_session_endpoint,Issuer:e.body.issuer}}).catch(function(e){throw i.serverErrorCode=e,t.stopEvent(i),e})},h.prototype.resolveEndpointsAsync=function(n,i){return a.__awaiter(this,void 0,Promise,function(){var t,r,o;return a.__generator(this,function(e){switch(e.label){case 0:return this.IsValidationEnabled?(t=this.canonicalAuthorityUrlComponents.HostNameAndPort,0!==u.TrustedAuthority.getTrustedHostList().length?[3,2]:[4,u.TrustedAuthority.setTrustedAuthoritiesFromNetwork(this.canonicalAuthority,n,i)]):[3,3];case 1:e.sent(),e.label=2;case 2:if(!u.TrustedAuthority.IsInTrustedHostList(t))throw s.ClientConfigurationError.createUntrustedAuthorityError(t);e.label=3;case 3:return r=this.GetOpenIdConfigurationEndpoint(),[4,(o=this).DiscoverEndpoints(r,n,i)];case 4:return o.tenantDiscoveryResponse=e.sent(),[2,this.tenantDiscoveryResponse]}})})},h.prototype.hasCachedMetadata=function(){return!!(this.tenantDiscoveryResponse&&this.tenantDiscoveryResponse.AuthorizationEndpoint&&this.tenantDiscoveryResponse.EndSessionEndpoint&&this.tenantDiscoveryResponse.Issuer)},h.prototype.GetOpenIdConfigurationEndpoint=function(){return this.DefaultOpenIdConfigurationEndpoint},h);function h(e,t,r){this.IsValidationEnabled=t,this.CanonicalAuthority=e,this.validateAsUri(),this.tenantDiscoveryResponse=r}t.Authority=i},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var s=r(1),o=(n.prototype.sendRequestAsync=function(e,t,r){var a=this;return new Promise(function(o,n){var i=new XMLHttpRequest;if(i.open(t,e,!0),i.onload=function(e){var t;(i.status<200||300<=i.status)&&n(a.handleError(i.responseText));try{t=JSON.parse(i.responseText)}catch(e){n(a.handleError(i.responseText))}var r={statusCode:i.status,body:t};o(r)},i.onerror=function(e){n(i.status)},t!==s.NetworkRequestType.GET)throw"not implemented";i.send()})},n.prototype.handleError=function(t){var e;try{if((e=JSON.parse(t)).error)return e.error;throw t}catch(e){return t}},n);function n(){}t.XhrClient=o},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var c=r(0),u=r(23),l=r(1),i=r(4),o=(a.setTrustedAuthoritiesFromConfig=function(e,t){e&&!this.getTrustedHostList().length&&t.forEach(function(e){a.TrustedHostList.push(e.toLowerCase())})},a.getAliases=function(i,a,s){return c.__awaiter(this,void 0,Promise,function(){var t,r,o,n;return c.__generator(this,function(e){return t=new u.XhrClient,r=l.NetworkRequestType.GET,o=""+l.AAD_INSTANCE_DISCOVERY_ENDPOINT+i+"oauth2/v2.0/authorize",n=a.createAndStartHttpEvent(s,r,o,"getAliases"),[2,t.sendRequestAsync(o,r,!0).then(function(e){return n.httpResponseStatus=e.statusCode,a.stopEvent(n),e.body.metadata}).catch(function(e){throw n.serverErrorCode=e,a.stopEvent(n),e})]})})},a.setTrustedAuthoritiesFromNetwork=function(r,o,n){return c.__awaiter(this,void 0,Promise,function(){var t;return c.__generator(this,function(e){switch(e.label){case 0:return[4,this.getAliases(r,o,n)];case 1:return e.sent().forEach(function(e){e.aliases.forEach(function(e){a.TrustedHostList.push(e.toLowerCase())})}),t=i.UrlUtils.GetUrlComponents(r).HostNameAndPort,a.getTrustedHostList().length&&!a.IsInTrustedHostList(t)&&a.TrustedHostList.push(t.toLowerCase()),[2]}})})},a.getTrustedHostList=function(){return this.TrustedHostList},a.IsInTrustedHostList=function(e){return-1<this.TrustedHostList.indexOf(e.toLowerCase())},a.TrustedHostList=[],a);function a(){}t.TrustedAuthority=o},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var c=r(0),o=r(12),n=r(4),u={clientId:"",authority:null,validateAuthority:!0,authorityMetadata:"",knownAuthorities:[],redirectUri:function(){return n.UrlUtils.getCurrentUrl()},postLogoutRedirectUri:function(){return n.UrlUtils.getCurrentUrl()},navigateToLoginRequestUrl:!0},l={cacheLocation:"sessionStorage",storeAuthStateInCookie:!1},h={logger:new o.Logger(null),loadFrameTimeout:6e3,tokenRenewalOffsetSeconds:300,navigateFrameWait:500},d={isAngular:!1,unprotectedResources:new Array,protectedResourceMap:new Map};t.buildConfiguration=function(e){var t=e.auth,r=e.cache,o=void 0===r?{}:r,n=e.system,i=void 0===n?{}:n,a=e.framework,s=void 0===a?{}:a;return{auth:c.__assign({},u,t),cache:c.__assign({},l,o),system:c.__assign({},h,i),framework:c.__assign({},d,s)}}},function(e,r,t){Object.defineProperty(r,"__esModule",{value:!0});var o=t(0),n=t(13);r.InteractionRequiredAuthErrorMessage={interactionRequired:{code:"interaction_required"},consentRequired:{code:"consent_required"},loginRequired:{code:"login_required"}};var i,a=(i=n.ServerError,o.__extends(s,i),s.isInteractionRequiredError=function(e){var t=[r.InteractionRequiredAuthErrorMessage.interactionRequired.code,r.InteractionRequiredAuthErrorMessage.consentRequired.code,r.InteractionRequiredAuthErrorMessage.loginRequired.code];return e&&-1<t.indexOf(e)},s.createLoginRequiredAuthError=function(e){return new s(r.InteractionRequiredAuthErrorMessage.loginRequired.code,e)},s.createInteractionRequiredAuthError=function(e){return new s(r.InteractionRequiredAuthErrorMessage.interactionRequired.code,e)},s.createConsentRequiredAuthError=function(e){return new s(r.InteractionRequiredAuthErrorMessage.consentRequired.code,e)},s);function s(e,t){var r=i.call(this,e,t)||this;return r.name="InteractionRequiredAuthError",Object.setPrototypeOf(r,s.prototype),r}r.InteractionRequiredAuthError=a},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0}),t.buildResponseStateOnly=function(e){return{uniqueId:"",tenantId:"",tokenType:"",idToken:null,idTokenClaims:null,accessToken:"",scopes:null,expiresOn:null,account:null,accountState:e,fromCache:!1}}},function(e,t,r){var o;Object.defineProperty(t,"__esModule",{value:!0});var n,i,a,s,c=r(0),u=c.__importDefault(r(14)),l=r(8),h=r(10);t.EVENT_KEYS={AUTHORITY:h.prependEventNamePrefix("authority"),AUTHORITY_TYPE:h.prependEventNamePrefix("authority_type"),PROMPT:h.prependEventNamePrefix("ui_behavior"),TENANT_ID:h.prependEventNamePrefix("tenant_id"),USER_ID:h.prependEventNamePrefix("user_id"),WAS_SUCESSFUL:h.prependEventNamePrefix("was_successful"),API_ERROR_CODE:h.prependEventNamePrefix("api_error_code"),LOGIN_HINT:h.prependEventNamePrefix("login_hint")},(i=n=t.API_CODE||(t.API_CODE={}))[i.AcquireTokenRedirect=2001]="AcquireTokenRedirect",i[i.AcquireTokenSilent=2002]="AcquireTokenSilent",i[i.AcquireTokenPopup=2003]="AcquireTokenPopup",i[i.LoginRedirect=2004]="LoginRedirect",i[i.LoginPopup=2005]="LoginPopup",i[i.Logout=2006]="Logout",(s=a=t.API_EVENT_IDENTIFIER||(t.API_EVENT_IDENTIFIER={})).AcquireTokenRedirect="AcquireTokenRedirect",s.AcquireTokenSilent="AcquireTokenSilent",s.AcquireTokenPopup="AcquireTokenPopup",s.LoginRedirect="LoginRedirect",s.LoginPopup="LoginPopup",s.Logout="Logout";var d,p=((o={})[a.AcquireTokenSilent]=n.AcquireTokenSilent,o[a.AcquireTokenPopup]=n.AcquireTokenPopup,o[a.AcquireTokenRedirect]=n.AcquireTokenRedirect,o[a.LoginPopup]=n.LoginPopup,o[a.LoginRedirect]=n.LoginRedirect,o[a.Logout]=n.Logout,o),g=(d=u.default,c.__extends(f,d),Object.defineProperty(f.prototype,"apiEventIdentifier",{set:function(e){this.event[l.TELEMETRY_BLOB_EVENT_NAMES.ApiTelemIdConstStrKey]=e},enumerable:!0,configurable:!0}),Object.defineProperty(f.prototype,"apiCode",{set:function(e){this.event[l.TELEMETRY_BLOB_EVENT_NAMES.ApiIdConstStrKey]=e},enumerable:!0,configurable:!0}),Object.defineProperty(f.prototype,"authority",{set:function(e){this.event[t.EVENT_KEYS.AUTHORITY]=h.scrubTenantFromUri(e).toLowerCase()},enumerable:!0,configurable:!0}),Object.defineProperty(f.prototype,"apiErrorCode",{set:function(e){this.event[t.EVENT_KEYS.API_ERROR_CODE]=e},enumerable:!0,configurable:!0}),Object.defineProperty(f.prototype,"tenantId",{set:function(e){this.event[t.EVENT_KEYS.TENANT_ID]=this.piiEnabled&&e?h.hashPersonalIdentifier(e):null},enumerable:!0,configurable:!0}),Object.defineProperty(f.prototype,"accountId",{set:function(e){this.event[t.EVENT_KEYS.USER_ID]=this.piiEnabled&&e?h.hashPersonalIdentifier(e):null},enumerable:!0,configurable:!0}),Object.defineProperty(f.prototype,"wasSuccessful",{get:function(){return!0===this.event[t.EVENT_KEYS.WAS_SUCESSFUL]},set:function(e){this.event[t.EVENT_KEYS.WAS_SUCESSFUL]=e},enumerable:!0,configurable:!0}),Object.defineProperty(f.prototype,"loginHint",{set:function(e){this.event[t.EVENT_KEYS.LOGIN_HINT]=this.piiEnabled&&e?h.hashPersonalIdentifier(e):null},enumerable:!0,configurable:!0}),Object.defineProperty(f.prototype,"authorityType",{set:function(e){this.event[t.EVENT_KEYS.AUTHORITY_TYPE]=e.toLowerCase()},enumerable:!0,configurable:!0}),Object.defineProperty(f.prototype,"promptType",{set:function(e){this.event[t.EVENT_KEYS.PROMPT]=e.toLowerCase()},enumerable:!0,configurable:!0}),f);function f(e,t,r){var o=d.call(this,h.prependEventNamePrefix("api_event"),e,r)||this;return r&&(o.apiCode=p[r],o.apiEventIdentifier=r),o.piiEnabled=t,o}t.default=g},function(e,t,r){e.exports=r(30)},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var o=r(15);t.UserAgentApplication=o.UserAgentApplication,t.authResponseCallback=o.authResponseCallback,t.errorReceivedCallback=o.errorReceivedCallback,t.tokenReceivedCallback=o.tokenReceivedCallback;var n=r(12);t.Logger=n.Logger;var i=r(12);t.LogLevel=i.LogLevel;var a=r(19);t.Account=a.Account;var s=r(1);t.Constants=s.Constants,t.ServerHashParamKeys=s.ServerHashParamKeys;var c=r(22);t.Authority=c.Authority;var u=r(15);t.CacheResult=u.CacheResult;var l=r(25);t.CacheLocation=l.CacheLocation,t.Configuration=l.Configuration;var h=r(42);t.AuthenticationParameters=h.AuthenticationParameters;var d=r(27);t.AuthResponse=d.AuthResponse;var p=r(2);t.CryptoUtils=p.CryptoUtils;var g=r(4);t.UrlUtils=g.UrlUtils;var f=r(20);t.WindowUtils=f.WindowUtils;var y=r(7);t.AuthError=y.AuthError;var m=r(6);t.ClientAuthError=m.ClientAuthError;var v=r(13);t.ServerError=v.ServerError;var E=r(5);t.ClientConfigurationError=E.ClientConfigurationError;var C=r(26);t.InteractionRequiredAuthError=C.InteractionRequiredAuthError},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});function o(e,t,r,o,n){this.authority=a.UrlUtils.CanonicalizeUri(e),this.clientId=t,this.scopes=r,this.homeAccountIdentifier=i.CryptoUtils.base64Encode(o)+"."+i.CryptoUtils.base64Encode(n)}var i=r(2),a=r(4);t.AccessTokenKey=o},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});function o(e,t,r,o){this.accessToken=e,this.idToken=t,this.expiresIn=r,this.homeAccountIdentifier=o}t.AccessTokenValue=o},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var o=r(2),n=r(6),i=r(3),a=(Object.defineProperty(s.prototype,"uid",{get:function(){return this._uid?this._uid:""},set:function(e){this._uid=e},enumerable:!0,configurable:!0}),Object.defineProperty(s.prototype,"utid",{get:function(){return this._utid?this._utid:""},set:function(e){this._utid=e},enumerable:!0,configurable:!0}),s);function s(e){if(!e||i.StringUtils.isEmpty(e))return this.uid="",void(this.utid="");try{var t=o.CryptoUtils.base64Decode(e),r=JSON.parse(t);r&&(r.hasOwnProperty("uid")&&(this.uid=r.uid),r.hasOwnProperty("utid")&&(this.utid=r.utid))}catch(e){throw n.ClientAuthError.createClientInfoDecodingError(e)}}t.ClientInfo=a},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});function o(e){if(a.StringUtils.isEmpty(e))throw n.ClientAuthError.createIdTokenNullOrEmptyError(e);try{this.rawIdToken=e,this.claims=i.TokenUtils.extractIdToken(e),this.claims&&(this.claims.hasOwnProperty("iss")&&(this.issuer=this.claims.iss),this.claims.hasOwnProperty("oid")&&(this.objectId=this.claims.oid),this.claims.hasOwnProperty("sub")&&(this.subject=this.claims.sub),this.claims.hasOwnProperty("tid")&&(this.tenantId=this.claims.tid),this.claims.hasOwnProperty("ver")&&(this.version=this.claims.ver),this.claims.hasOwnProperty("preferred_username")&&(this.preferredName=this.claims.preferred_username),this.claims.hasOwnProperty("name")&&(this.name=this.claims.name),this.claims.hasOwnProperty("nonce")&&(this.nonce=this.claims.nonce),this.claims.hasOwnProperty("exp")&&(this.expiration=this.claims.exp),this.claims.hasOwnProperty("home_oid")&&(this.homeObjectId=this.claims.home_oid),this.claims.hasOwnProperty("sid")&&(this.sid=this.claims.sid),this.claims.hasOwnProperty("cloud_instance_host_name")&&(this.cloudInstance=this.claims.cloud_instance_host_name))}catch(e){throw n.ClientAuthError.createIdTokenParsingError(e)}}var n=r(6),i=r(17),a=r(3);t.IdToken=o},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var s,o=r(0),c=r(1),u=r(36),n=r(37),l=r(6),i=r(18),a=(s=n.BrowserStorage,o.__extends(h,s),h.prototype.migrateCacheEntries=function(r){var o=this,e=c.Constants.cachePrefix+"."+c.PersistentCacheKeys.IDTOKEN,t=c.Constants.cachePrefix+"."+c.PersistentCacheKeys.CLIENT_INFO,n=c.Constants.cachePrefix+"."+c.ErrorCacheKeys.ERROR,i=c.Constants.cachePrefix+"."+c.ErrorCacheKeys.ERROR_DESC,a=[s.prototype.getItem.call(this,e),s.prototype.getItem.call(this,t),s.prototype.getItem.call(this,n),s.prototype.getItem.call(this,i)];[c.PersistentCacheKeys.IDTOKEN,c.PersistentCacheKeys.CLIENT_INFO,c.ErrorCacheKeys.ERROR,c.ErrorCacheKeys.ERROR_DESC].forEach(function(e,t){return o.duplicateCacheEntry(e,a[t],r)})},h.prototype.duplicateCacheEntry=function(e,t,r){t&&this.setItem(e,t,r)},h.prototype.generateCacheKey=function(t,r){try{return JSON.parse(t),t}catch(e){return 0===t.indexOf(""+c.Constants.cachePrefix)||0===t.indexOf(c.Constants.adalIdToken)?t:r?c.Constants.cachePrefix+"."+this.clientId+"."+t:c.Constants.cachePrefix+"."+t}},h.prototype.setItem=function(e,t,r){s.prototype.setItem.call(this,this.generateCacheKey(e,!0),t,r),this.rollbackEnabled&&!r&&s.prototype.setItem.call(this,this.generateCacheKey(e,!1),t,r)},h.prototype.getItem=function(e,t){return s.prototype.getItem.call(this,this.generateCacheKey(e,!0),t)},h.prototype.removeItem=function(e){s.prototype.removeItem.call(this,this.generateCacheKey(e,!0)),this.rollbackEnabled&&s.prototype.removeItem.call(this,this.generateCacheKey(e,!1))},h.prototype.resetCacheItems=function(){var e,t=window[this.cacheLocation];for(e in t)t.hasOwnProperty(e)&&-1!==e.indexOf(c.Constants.cachePrefix)&&s.prototype.removeItem.call(this,e)},h.prototype.resetTempCacheItems=function(e){var t=this,r=e&&i.RequestUtils.parseLibraryState(e).id,o=this.tokenRenewalInProgress(e),n=window[this.cacheLocation];r&&!o&&Object.keys(n).forEach(function(e){-1!==e.indexOf(r)&&(t.removeItem(e),s.prototype.clearItemCookie.call(t,e))}),this.removeItem(c.TemporaryCacheKeys.INTERACTION_STATUS),this.removeItem(c.TemporaryCacheKeys.REDIRECT_REQUEST)},h.prototype.setItemCookie=function(e,t,r){s.prototype.setItemCookie.call(this,this.generateCacheKey(e,!0),t,r),this.rollbackEnabled&&s.prototype.setItemCookie.call(this,this.generateCacheKey(e,!1),t,r)},h.prototype.clearItemCookie=function(e){s.prototype.clearItemCookie.call(this,this.generateCacheKey(e,!0)),this.rollbackEnabled&&s.prototype.clearItemCookie.call(this,this.generateCacheKey(e,!1))},h.prototype.getItemCookie=function(e){return s.prototype.getItemCookie.call(this,this.generateCacheKey(e,!0))},h.prototype.getAllAccessTokens=function(i,a){var s=this;return Object.keys(window[this.cacheLocation]).reduce(function(e,t){if(t.match(i)&&t.match(a)&&t.match(c.Constants.scopes)){var r=s.getItem(t);if(r)try{var o=JSON.parse(t),n=new u.AccessTokenCacheItem(o,JSON.parse(r));return e.concat([n])}catch(e){throw l.ClientAuthError.createCacheParseError(t)}}return e},[])},h.prototype.tokenRenewalInProgress=function(e){var t=this.getItem(h.generateTemporaryCacheKey(c.TemporaryCacheKeys.RENEW_STATUS,e));return!(!t||t!==c.Constants.inProgress)},h.prototype.clearMsalCookie=function(e){var r=this;e?(this.clearItemCookie(h.generateTemporaryCacheKey(c.TemporaryCacheKeys.NONCE_IDTOKEN,e)),this.clearItemCookie(h.generateTemporaryCacheKey(c.TemporaryCacheKeys.STATE_LOGIN,e)),this.clearItemCookie(h.generateTemporaryCacheKey(c.TemporaryCacheKeys.LOGIN_REQUEST,e)),this.clearItemCookie(h.generateTemporaryCacheKey(c.TemporaryCacheKeys.STATE_ACQ_TOKEN,e))):document.cookie.split(";").forEach(function(e){var t=e.trim().split("=")[0];-1<t.indexOf(c.Constants.cachePrefix)&&s.prototype.clearItemCookie.call(r,t)})},h.generateAcquireTokenAccountKey=function(e,t){var r=i.RequestUtils.parseLibraryState(t).id;return""+c.TemporaryCacheKeys.ACQUIRE_TOKEN_ACCOUNT+c.Constants.resourceDelimiter+e+c.Constants.resourceDelimiter+r},h.generateAuthorityKey=function(e){return h.generateTemporaryCacheKey(c.TemporaryCacheKeys.AUTHORITY,e)},h.generateTemporaryCacheKey=function(e,t){var r=i.RequestUtils.parseLibraryState(t).id;return""+e+c.Constants.resourceDelimiter+r},h);function h(e,t,r){var o=s.call(this,t)||this;return o.clientId=e,o.rollbackEnabled=!0,o.migrateCacheEntries(r),o}t.AuthCache=a},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});function o(e,t){this.key=e,this.value=t}t.AccessTokenCacheItem=o},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var o=r(5),n=r(7),i=(a.prototype.setItem=function(e,t,r){window[this.cacheLocation].setItem(e,t),r&&this.setItemCookie(e,t)},a.prototype.getItem=function(e,t){return t&&this.getItemCookie(e)?this.getItemCookie(e):window[this.cacheLocation].getItem(e)},a.prototype.removeItem=function(e){return window[this.cacheLocation].removeItem(e)},a.prototype.clear=function(){return window[this.cacheLocation].clear()},a.prototype.setItemCookie=function(e,t,r){var o=e+"="+t+";path=/;";r&&(o+="expires="+this.getCookieExpirationTime(r)+";"),document.cookie=o},a.prototype.getItemCookie=function(e){for(var t=e+"=",r=document.cookie.split(";"),o=0;o<r.length;o++){for(var n=r[o];" "===n.charAt(0);)n=n.substring(1);if(0===n.indexOf(t))return n.substring(t.length,n.length)}return""},a.prototype.clearItemCookie=function(e){this.setItemCookie(e,"",-1)},a.prototype.getCookieExpirationTime=function(e){var t=new Date;return new Date(t.getTime()+24*e*60*60*1e3).toUTCString()},a);function a(e){if(!window)throw n.AuthError.createNoWindowObjectError("Browser storage class could not find window object");if(!(void 0!==window[e]&&null!=window[e]))throw o.ClientConfigurationError.createStorageNotSupportedError(e);this.cacheLocation=e}t.BrowserStorage=i},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var o=r(0),n=(i.setResponseIdToken=function(e,t){if(!e)return null;if(!t)return e;var r=Number(t.expiration);return r&&!e.expiresOn&&(e.expiresOn=new Date(1e3*r)),o.__assign({},e,{idToken:t,idTokenClaims:t.claims,uniqueId:t.objectId||t.subject,tenantId:t.tenantId})},i);function i(){}t.ResponseUtils=n},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var o=r(0),s=o.__importDefault(r(40)),n=r(1),i=o.__importDefault(r(28)),a=o.__importDefault(r(41)),c=(u.getTelemetrymanagerStub=function(e,t){return new this({platform:{applicationName:"UnSetStub",applicationVersion:"0.0"},clientId:e},function(){},t)},u.prototype.startEvent=function(e){this.logger.verbose("Telemetry Event started: "+e.key),this.telemetryEmitter&&(e.start(),this.inProgressEvents[e.key]=e)},u.prototype.stopEvent=function(e){if(this.logger.verbose("Telemetry Event stopped: "+e.key),this.telemetryEmitter&&this.inProgressEvents[e.key]){e.stop(),this.incrementEventCount(e);var t=this.completedEvents[e.telemetryCorrelationId];this.completedEvents[e.telemetryCorrelationId]=(t||[]).concat([e]),delete this.inProgressEvents[e.key]}},u.prototype.flush=function(e){var t=this;if(this.logger.verbose("Flushing telemetry events: "+e),this.telemetryEmitter&&this.completedEvents[e]){var r=this.getOrphanedEvents(e);r.forEach(function(e){return t.incrementEventCount(e)});var o=this.completedEvents[e].concat(r);delete this.completedEvents[e];var n=this.eventCountByCorrelationId[e];if(delete this.eventCountByCorrelationId[e],o&&o.length){var i=new s.default(this.telemetryPlatform,e,this.clientId,n),a=o.concat([i]);this.telemetryEmitter(a.map(function(e){return e.get()}))}}},u.prototype.createAndStartApiEvent=function(e,t){var r=new i.default(e,this.logger.isPiiLoggingEnabled(),t);return this.startEvent(r),r},u.prototype.stopAndFlushApiEvent=function(e,t,r,o){t.wasSuccessful=r,o&&(t.apiErrorCode=o),this.stopEvent(t),this.flush(e)},u.prototype.createAndStartHttpEvent=function(e,t,r,o){var n=new a.default(e,o);return n.url=r,n.httpMethod=t,this.startEvent(n),n},u.prototype.incrementEventCount=function(e){var t,r=e.eventName,o=this.eventCountByCorrelationId[e.telemetryCorrelationId];o?o[r]=o[r]?o[r]+1:1:this.eventCountByCorrelationId[e.telemetryCorrelationId]=((t={})[r]=1,t)},u.prototype.getOrphanedEvents=function(o){var n=this;return Object.keys(this.inProgressEvents).reduce(function(e,t){if(-1===t.indexOf(o))return e;var r=n.inProgressEvents[t];return delete n.inProgressEvents[t],e.concat([r])},[])},u);function u(e,t,r){this.completedEvents={},this.inProgressEvents={},this.eventCountByCorrelationId={},this.onlySendFailureTelemetry=!1,this.telemetryPlatform=o.__assign({sdk:n.Constants.libraryName,sdkVersion:n.libraryVersion(),networkInformation:{connectionSpeed:"undefined"!=typeof navigator&&navigator.connection&&navigator.connection.effectiveType}},e.platform),this.clientId=e.clientId,this.onlySendFailureTelemetry=e.onlySendFailureTelemetry,this.telemetryEmitter=t,this.logger=r}t.default=c},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var i,o=r(0),a=r(8),n=o.__importDefault(r(14)),s=r(10),c=(i=n.default,o.__extends(u,i),u.prototype.getEventCount=function(e,t){return t[e]?t[e]:0},u);function u(e,t,r,o){var n=i.call(this,s.prependEventNamePrefix("default_event"),t,"DefaultEvent")||this;return n.event[s.prependEventNamePrefix("client_id")]=r,n.event[s.prependEventNamePrefix("sdk_plaform")]=e.sdk,n.event[s.prependEventNamePrefix("sdk_version")]=e.sdkVersion,n.event[s.prependEventNamePrefix("application_name")]=e.applicationName,n.event[s.prependEventNamePrefix("application_version")]=e.applicationVersion,n.event[s.prependEventNamePrefix("effective_connection_speed")]=e.networkInformation&&e.networkInformation.connectionSpeed,n.event[""+a.TELEMETRY_BLOB_EVENT_NAMES.UiEventCountTelemetryBatchKey]=n.getEventCount(s.prependEventNamePrefix("ui_event"),o),n.event[""+a.TELEMETRY_BLOB_EVENT_NAMES.HttpEventCountTelemetryBatchKey]=n.getEventCount(s.prependEventNamePrefix("http_event"),o),n.event[""+a.TELEMETRY_BLOB_EVENT_NAMES.CacheEventCountConstStrKey]=n.getEventCount(s.prependEventNamePrefix("cache_event"),o),n}t.default=c},function(e,r,t){Object.defineProperty(r,"__esModule",{value:!0});var o=t(0),n=o.__importDefault(t(14)),i=t(10),a=t(16);r.EVENT_KEYS={HTTP_PATH:i.prependEventNamePrefix("http_path"),USER_AGENT:i.prependEventNamePrefix("user_agent"),QUERY_PARAMETERS:i.prependEventNamePrefix("query_parameters"),API_VERSION:i.prependEventNamePrefix("api_version"),RESPONSE_CODE:i.prependEventNamePrefix("response_code"),O_AUTH_ERROR_CODE:i.prependEventNamePrefix("oauth_error_code"),HTTP_METHOD:i.prependEventNamePrefix("http_method"),REQUEST_ID_HEADER:i.prependEventNamePrefix("request_id_header"),SPE_INFO:i.prependEventNamePrefix("spe_info"),SERVER_ERROR_CODE:i.prependEventNamePrefix("server_error_code"),SERVER_SUB_ERROR_CODE:i.prependEventNamePrefix("server_sub_error_code"),URL:i.prependEventNamePrefix("url")};var s,c=(s=n.default,o.__extends(u,s),Object.defineProperty(u.prototype,"url",{set:function(e){var t=i.scrubTenantFromUri(e);this.event[r.EVENT_KEYS.URL]=t&&t.toLowerCase()},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"httpPath",{set:function(e){this.event[r.EVENT_KEYS.HTTP_PATH]=i.scrubTenantFromUri(e).toLowerCase()},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"userAgent",{set:function(e){this.event[r.EVENT_KEYS.USER_AGENT]=e},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"queryParams",{set:function(e){this.event[r.EVENT_KEYS.QUERY_PARAMETERS]=a.ServerRequestParameters.generateQueryParametersString(e)},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"apiVersion",{set:function(e){this.event[r.EVENT_KEYS.API_VERSION]=e.toLowerCase()},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"httpResponseStatus",{set:function(e){this.event[r.EVENT_KEYS.RESPONSE_CODE]=e},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"oAuthErrorCode",{set:function(e){this.event[r.EVENT_KEYS.O_AUTH_ERROR_CODE]=e},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"httpMethod",{set:function(e){this.event[r.EVENT_KEYS.HTTP_METHOD]=e},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"requestIdHeader",{set:function(e){this.event[r.EVENT_KEYS.REQUEST_ID_HEADER]=e},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"speInfo",{set:function(e){this.event[r.EVENT_KEYS.SPE_INFO]=e},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"serverErrorCode",{set:function(e){this.event[r.EVENT_KEYS.SERVER_ERROR_CODE]=e},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"serverSubErrorCode",{set:function(e){this.event[r.EVENT_KEYS.SERVER_SUB_ERROR_CODE]=e},enumerable:!0,configurable:!0}),u);function u(e,t){return s.call(this,i.prependEventNamePrefix("http_event"),e,t)||this}r.default=c},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var o=r(5);t.validateClaimsRequest=function(e){if(e.claimsRequest)try{JSON.parse(e.claimsRequest)}catch(e){throw o.ClientConfigurationError.createClaimsRequestParsingError(e)}}}],n.c=o,n.d=function(e,t,r){n.o(e,t)||Object.defineProperty(e,t,{enumerable:!0,get:r})},n.r=function(e){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})},n.t=function(t,e){if(1&e&&(t=n(t)),8&e)return t;if(4&e&&"object"==typeof t&&t&&t.__esModule)return t;var r=Object.create(null);if(n.r(r),Object.defineProperty(r,"default",{enumerable:!0,value:t}),2&e&&"string"!=typeof t)for(var o in t)n.d(r,o,function(e){return t[e]}.bind(null,o));return r},n.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return n.d(t,"a",t),t},n.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},n.p="",n(n.s=29);function n(e){if(o[e])return o[e].exports;var t=o[e]={i:e,l:!1,exports:{}};return r[e].call(t.exports,t,t.exports,n),t.l=!0,t.exports}var r,o});
//# sourceMappingURL=msal.min.js.map

/***/ }),

/***/ "./packages/Microsoft.Office.WebAuth.Implicit/scripts/Implicit.ts":
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/**
 * Copyright (c) Microsoft. All rights reserved.
 */
/// <reference path="./Definitions/IImplicitAuthConfig.d.ts" />
/// <reference path="./Definitions/IImplicitAuthResult.d.ts" />
// Above references are needed for ts-node
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
Object.defineProperty(exports, "__esModule", { value: true });
var Msal = __webpack_require__("./packages/Microsoft.Office.WebAuth.Implicit/lib/msal.js");
var api_js_1 = __webpack_require__("./packages/Microsoft.Office.WebAuth.Implicit/lib/api.js");
api_js_1.addNamespaceMapping('Office.Identity.WebAuth.Implicit', '5c65bbc4edbf480d9637ace04d62bd98-12844893-8ab9-4dde-b850-5612cb12e0f2-7822');
var applications;
// Export properties and functions so we can test. But do not open them up 
// in case partners create a dependency or use them improperly.
/////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////
var authConfig;
var enableConsoleLog = false;
/**
 * Internal Module includes definition of different constants
 */
var Constants;
(function (Constants) {
    var IdentityProvider = /** @class */ (function () {
        function IdentityProvider() {
        }
        /**
         * AAD
         */
        IdentityProvider.Aad = "aad";
        /**
         * MSA
         */
        IdentityProvider.Msa = "msa";
        return IdentityProvider;
    }());
    Constants.IdentityProvider = IdentityProvider;
    var PostMessageType = /** @class */ (function () {
        function PostMessageType() {
        }
        PostMessageType.RequestAuthConfig = "RequestAuthConfig";
        PostMessageType.ResponseAuthConfig = "ResponseAuthConfig";
        PostMessageType.iFramePrefix = "msalRenewFrame";
        PostMessageType.iFrameIdTokenPrefix = "msalIdTokenFrame";
        return PostMessageType;
    }());
    Constants.PostMessageType = PostMessageType;
    var Authority = /** @class */ (function () {
        function Authority() {
        }
        /**
         * Prod
         */
        Authority.Prod = "https://login.microsoftonline.com/";
        /**
         * Legacy Prod
         */
        Authority.ProdLegacy = "https://login.windows.net/";
        /**
         * PPE
         */
        Authority.Ppe = "https://login.windows-ppe.net/";
        Authority.AadSuffix = "common";
        Authority.MsaSuffix = "consumers";
        return Authority;
    }());
    Constants.Authority = Authority;
    var Telemetry = /** @class */ (function () {
        function Telemetry() {
        }
        Telemetry.OtelInstance = "otel";
        Telemetry.LoadTelemetryName = "Office.Identity.WebAuth.Implicit.Load";
        Telemetry.GetTokenTelemetryName = "Office.Identity.WebAuth.Implicit.GetToken";
        Telemetry.CheckUpnTelemetryName = "Office.Identity.WebAuth.Implicit.CheckUpn";
        Telemetry.Duration = "Duration";
        Telemetry.Succeeded = "Succeeded";
        Telemetry.IdentityProvider = "IdentityProvider";
        Telemetry.ApplicationId = "ApplicationId";
        Telemetry.TokenScope = "TokenScope";
        Telemetry.CorrelationId = "CorrelationId";
        Telemetry.loadedApplicationCount = "LoadedApplicationCount";
        Telemetry.ErrorCodeForGetToken = "ErrorCodeForGetToken";
        Telemetry.ErrorMessageForGetToken = "ErrorMessageForGetToken";
        Telemetry.ErrorCodeForCheckUpn = "ErrorCodeForCheckUpn";
        Telemetry.ErrorMessageForCheckUpn = "ErrorMessageForCheckUpn";
        return Telemetry;
    }());
    Constants.Telemetry = Telemetry;
})(Constants || (Constants = {}));
/**
 * Internal Module contains utility methods for logging
 */
var LoggingUtils;
(function (LoggingUtils) {
    /**
     * Returns if we should proceed with logging or not
     * @param shouldLog - should this message be logged
     */
    function shouldProceed(shouldLog) {
        if (shouldLog != null && shouldLog !== undefined && !shouldLog) {
            return false;
        }
        return true;
    }
    /**
     * Logs message to the console
     * @param message - message which was passed in
     * @param shouldLog - should this message be logged
     */
    function log(message, shouldLog) {
        if (!shouldProceed(shouldLog)) {
            return false;
        }
        console.log(message);
        return true;
    }
    LoggingUtils.log = log;
    /**
     * Logs warning message to the console
     * @param message - message which was passed in
     * @param shouldLog - should this message be logged
     */
    function warn(message, shouldLog) {
        if (!shouldProceed(shouldLog)) {
            return false;
        }
        console.warn(message);
        return true;
    }
    LoggingUtils.warn = warn;
    /**
     * Logs error message to the console
     * @param message - message which was passed in
     * @param shouldLog - should this message be logged
     */
    function error(message, shouldLog) {
        if (!shouldProceed(shouldLog)) {
            return false;
        }
        console.error(message);
        return true;
    }
    LoggingUtils.error = error;
})(LoggingUtils || (LoggingUtils = {}));
/**
 * Internal Module contains utility methods for extracting tokens
 */
var ExtractUtils;
(function (ExtractUtils) {
    /**
     * Extract AccessToken by decoding the RAWAccessToken
     *
     * @param encodedAccessToken
     */
    function extractAccessToken(encodedAccessToken) {
        // access token will be decoded to get the username
        var decodedToken = decodeJwt(encodedAccessToken);
        if (!decodedToken) {
            return null;
        }
        try {
            var base64AccessToken = decodedToken.JWSPayload;
            var base64Decoded = base64Decode(base64AccessToken);
            if (!base64Decoded) {
                LoggingUtils.log("The returned access_token could not be base64 url safe decoded.");
                return null;
            }
            // ECMA script has JSON built-in support
            return JSON.parse(base64Decoded);
        }
        catch (err) {
            LoggingUtils.error("The returned access_token could not be decoded" + err);
        }
        return null;
    }
    ExtractUtils.extractAccessToken = extractAccessToken;
    ;
    /**
     * Decodes a base64 encoded string.
     *
     * @param input
     */
    function base64Decode(input) {
        var encodedString = input.replace(/-/g, "+").replace(/_/g, "/");
        switch (encodedString.length % 4) {
            case 0:
                break;
            case 2:
                encodedString += "==";
                break;
            case 3:
                encodedString += "=";
                break;
            default:
                throw new Error("Invalid base64 string");
        }
        return decodeURIComponent(atob(encodedString).split("").map(function (c) {
            return "%" + ("00" + c.charCodeAt(0).toString(16)).slice(-2);
        }).join(""));
    }
    ;
    /**
     * decode a JWT
     *
     * @param jwtToken
     */
    function decodeJwt(jwtToken) {
        if (jwtToken === "undefined" || !jwtToken || 0 === jwtToken.length) {
            return null;
        }
        var idTokenPartsRegex = /^([^\.\s]*)\.([^\.\s]+)\.([^\.\s]*)$/;
        var matches = idTokenPartsRegex.exec(jwtToken);
        if (!matches || matches.length < 4) {
            LoggingUtils.warn("The returned access_token is not parseable.");
            return null;
        }
        var crackedToken = {
            header: matches[1],
            JWSPayload: matches[2],
            JWSSig: matches[3]
        };
        return crackedToken;
    }
    ;
})(ExtractUtils || (ExtractUtils = {}));
/**
 * Module includes timer related utilities
 */
var TimerUtils;
(function (TimerUtils) {
    /**
     * Timer function
     */
    function timer() {
        var timeStart = new Date().getTime();
        return {
            /**
             * Returns time in seconds (example: 500)
             */
            get seconds() {
                var seconds = Math.ceil((new Date().getTime() - timeStart) / 1000);
                return seconds;
            },
            /**
             * Returns time in Milliseconds (example: 2000)
             */
            get ms() {
                var ms = (new Date().getTime() - timeStart);
                return ms;
            },
            /**
             * Returns formatted time in seconds (example: 500s)
             */
            get formattedSeconds() {
                var seconds = Math.ceil(this.seconds / 1000) + "s";
                return seconds;
            },
            /**
             * Returns formatted time in Milliseconds (example: 2000ms)
             */
            get formattedMs() {
                var ms = this.ms + "ms";
                return ms;
            }
        };
    }
    TimerUtils.timer = timer;
})(TimerUtils || (TimerUtils = {}));
/**
 * Load implicit auth module
 * @param configurations - auth configs
 * @param correlationId - the same correlation ID exists under the caller's context, the same correlation ID will passed on to MSAL.js for unified experience
 * @returns the {@link IImplicitLoadResult} object
 */
function Load(configurations, correlationId) {
    var timerClock = TimerUtils.timer();
    authConfig = configurations;
    if (authConfig.enableConsoleLogging) {
        enableConsoleLog = authConfig.enableConsoleLogging;
    }
    // Implicitly swap login.windows.net to login.microsoftonline.com
    if (authConfig.authority) {
        authConfig.authority = authConfig.authority.replace(Constants.Authority.ProdLegacy, Constants.Authority.Prod);
    }
    // Hidden iframe --> { request: 'RequestAuthConfig', iframe: id } --> Return iframe --> {request: 'ResponseAuthConfig', config:authConfig }
    // Only add an event listener once
    if (!applications) {
        window.addEventListener('message', function (e) {
            if (e.origin && e.origin == location.origin && e.data && e.data.iframe && e.data.request && e.data.request == Constants.PostMessageType.RequestAuthConfig) {
                var targetIframeName = Constants.PostMessageType.iFramePrefix + e.data.iframe;
                var targetiFrame = document.getElementById(targetIframeName.replace('+', ' '));
                if (targetiFrame === null) {
                    targetIframeName = Constants.PostMessageType.iFrameIdTokenPrefix;
                    var targetiFrame = document.getElementById(targetIframeName.replace('+', ' '));
                }
                LoggingUtils.log('received from ' + targetIframeName, enableConsoleLog);
                if (targetiFrame === null) {
                    LoggingUtils.log('targetiFrame is null', enableConsoleLog);
                    return;
                }
                var targetContentWindow = targetiFrame.contentWindow;
                if (targetContentWindow != null) {
                    LoggingUtils.log('returning to ' + e.data.iframe, enableConsoleLog);
                    targetContentWindow.postMessage({ request: Constants.PostMessageType.ResponseAuthConfig, config: authConfig }, location.origin);
                }
            }
        }, false);
    }
    if (authConfig.idp.toLowerCase() === Constants.IdentityProvider.Msa.toLowerCase()) {
        // There is no guarantee that our callers will send MSAL.js supported authority to us,
        // hard code the authority to be login.microsoftonline.com since there is no PPE tenant for MSA.
        authConfig.authority = Constants.Authority.Prod + Constants.Authority.MsaSuffix;
    }
    else {
        if (!authConfig.authority) {
            authConfig.authority = Constants.Authority.Prod + Constants.Authority.AadSuffix;
        }
        // Add common suffix if it doesn't have
        if (authConfig.authority.indexOf(Constants.Authority.AadSuffix) < 0) {
            if (authConfig.authority.charAt(authConfig.authority.length - 1) == "/")
                authConfig.authority += Constants.Authority.AadSuffix;
            else
                authConfig.authority += "/" + Constants.Authority.AadSuffix;
        }
    }
    applications = new Array();
    for (var _i = 0, _a = authConfig.appIds; _i < _a.length; _i++) {
        var appId = _a[_i];
        if (!appId || !IsGuid(appId)) {
            continue;
        }
        var application = new Msal.UserAgentApplication({
            auth: {
                clientId: appId,
                authority: authConfig.authority,
                redirectUri: (authConfig.redirectUri) ? authConfig.redirectUri.split("?")[0] : location.href.split("?")[0],
                navigateToLoginRequestUrl: (authConfig.navigateToLoginRequestUrl) ? authConfig.navigateToLoginRequestUrl : true,
            },
            cache: {
                cacheLocation: 'localStorage',
                // Store auth state in cookies can make the request too big and fail the request sometimes, need to keep it as false.
                storeAuthStateInCookie: false
            },
            system: {
                loadFrameTimeout: (authConfig.timeout) ? authConfig.timeout : 6000,
            },
        });
        var entry = { applicationId: appId, application: application };
        applications.push(entry);
        HandleFragment(application);
    }
    ;
    // For data fields that are null, blank or empty, the value is set to "unknown" at this point
    // based on office-online-otel documentation: https://office.visualstudio.com/OC/_git/office-online-ui?path=%2Fpackages%2Foffice-online-otel%2FREADME.md&version=GBmaster
    // Null or undefined value will cause exceptions down the line.
    var dataFields = [
        { name: Constants.Telemetry.Duration, int64: timerClock.ms },
        { name: Constants.Telemetry.Succeeded, bool: true },
        { name: Constants.Telemetry.IdentityProvider, string: authConfig.idp.toLowerCase() },
        { name: Constants.Telemetry.CorrelationId, string: correlationId ? correlationId : 'unknown' },
        { name: Constants.Telemetry.loadedApplicationCount, int64: applications.length }
    ];
    if (!authConfig.telemetryInstance && typeof OTel === "undefined") {
        api_js_1.sendTelemetryEvent({
            name: Constants.Telemetry.LoadTelemetryName,
            dataFields: dataFields
        });
    }
    return {
        Telemetry: {
            timeToLoad: timerClock.ms,
            succeeded: true,
            idp: authConfig.idp.toLowerCase(),
            correlationId: correlationId ? correlationId : '',
            loadedApplicationCount: applications.length
        }
    };
}
exports.Load = Load;
/**
 * Saves data included in the hash fragment to storage.
 * @param application - The calling application
 */
function HandleFragment(application) {
    // isCallback will be deprecated in favor of urlContainsHash in MSAL.js v2.0.
    if (application.isCallback(window.location.hash)) {
        LoggingUtils.log("Hash: " + window.location.hash, enableConsoleLog);
        application.handleAuthenticationResponse();
        LoggingUtils.log("Completed Hash Handling", enableConsoleLog);
    }
    else {
        // Clear the existing cache if it's not a callback
        application.clearCache();
    }
}
/**
 * Acquire an access token by given target
 * @param target - resource for V1 token, scope for V2 token
 * @param applicationId - the application ID which needs access token
 * @param correlationId - the same correlation ID exists under the caller's context, the same correlation ID will passed on to MSAL.js for unified experience
 * @param login - If true, shows a login dialog. If false, skips login.
 * @param popup - If true, popsup a dialog for interactive flow. If false, acquires token silently.
 * @param forThirdParty - If true, treats the caller as third-party and avoids sending PII telemetry.
 * @param claims - Claims from AAD to be used in scenarios such as credentials change or MFA
 * @returns {Promise.<IImplicitAuthResult>} - a promise that is fulfilled when this function has completed, or rejected if an error was raised. Returns the {@link IImplicitAuthResult} object
 */
function GetToken(target, applicationId, correlationId, login, popup, forThirdParty, claims) {
    if (forThirdParty === void 0) { forThirdParty = false; }
    var timerClock = TimerUtils.timer();
    var application = GetApplication(applicationId);
    // Wrong format of correlation ID or blank are not valid in MSAL.js
    // With an invalid correlation ID in the request, the access token acquiring request will be rejected by MSAL.js with exceptions.
    // Correlation ID will be set to undefined in those cases and MSAL.js will generate a new correlation ID if it is undefined.
    if (!correlationId || !IsGuid(correlationId)) {
        correlationId = undefined;
    }
    var result = {};
    if (!target) {
        result.ErrorCode = 'missing_target';
        result.ErrorMessage = 'The provided target for Implicit.GetToken is null, blank or empty';
        LogTelemetryDataFieldsForGetToken(result, correlationId, applicationId, forThirdParty, timerClock, undefined, result.ErrorCode, result.ErrorMessage);
        return Promise.reject(result);
    }
    var scopes = [GetScope(target)];
    if (!applicationId || !IsGuid(applicationId)) {
        result.ErrorCode = 'invalid_application_ID';
        result.ErrorMessage = 'The provided application ID for Implicit.GetToken is null, blank, empty or with invalid format';
        LogTelemetryDataFieldsForGetToken(result, correlationId, applicationId, forThirdParty, timerClock, scopes, result.ErrorCode, result.ErrorMessage);
        return Promise.reject(result);
    }
    var isMsa = authConfig.idp.toLowerCase() === Constants.IdentityProvider.Msa.toLowerCase();
    // When popup = true, login will be attempted.
    var attemptLogin = function () {
        LoggingUtils.log("Logging in", enableConsoleLog);
        application.authResponseCallback = function () { HandleFragment(application); }; // callback must be set directly on the application.
        return application.loginPopup({
            // Prefill the UPN so that the user just needs to enter the password.
            scopes: scopes,
            loginHint: authConfig.upn,
            correlationId: correlationId,
        }).catch(function (error) {
            application.clearCache();
            return Promise.reject(error);
        });
    };
    var acquireToken = function () {
        return new Promise(function (resolve, reject) {
            LoggingUtils.log("Config: " + JSON.stringify(application), enableConsoleLog);
            LoggingUtils.log("application calls acquireTokenSilent", enableConsoleLog);
            var tokenRequest;
            var tokenParams = __assign({ scopes: scopes, loginHint: authConfig.upn, correlationId: correlationId }, claims && { claimsRequest: claims });
            if (popup) {
                tokenRequest = application.acquireTokenPopup(tokenParams);
            }
            else {
                tokenRequest = application.acquireTokenSilent(tokenParams);
            }
            tokenRequest.then(function (authResponse) {
                var token = authResponse.accessToken;
                if (token && UpnMatchesUpnFromIdToken(authResponse.account) && (forThirdParty || isMsa || UpnMatchesUpnFromAccessToken(token))) {
                    LoggingUtils.log("acquireToken->token: " + token, enableConsoleLog);
                    result.Token = token;
                    LogTelemetryDataFieldsForGetToken(result, correlationId, applicationId, forThirdParty, timerClock, scopes);
                    return resolve(result);
                }
                else {
                    result.ErrorCode = "upn_mismatch";
                    result.ErrorMessage = "upn doesn't match with given upn in config";
                    LogTelemetryDataFieldsForGetToken(result, correlationId, applicationId, forThirdParty, timerClock, scopes, result.ErrorCode, result.ErrorMessage);
                    LoggingUtils.log("acquireToken->error: " + result.ErrorMessage, enableConsoleLog);
                    application.clearCache();
                    return reject(result);
                }
            }).catch(function (error) {
                result.ErrorCode = error.errorCode;
                result.ErrorMessage = error.errorMessage;
                LogTelemetryDataFieldsForGetToken(result, correlationId, applicationId, forThirdParty, timerClock, scopes, error.errorCode, error.errorMessage);
                LoggingUtils.log("acquireToken->error: " + result.ErrorMessage, enableConsoleLog);
                application.clearCache();
                return reject(result);
            });
        });
    };
    if (login) {
        return attemptLogin()
            .catch(function (loginError) {
            LoggingUtils.log("loginPopup->error: " + loginError.errorMessage, enableConsoleLog);
            result.ErrorCode = loginError.errorCode;
            result.ErrorMessage = loginError.errorMessage;
            LogTelemetryDataFieldsForGetToken(result, correlationId, applicationId, forThirdParty, timerClock, scopes, loginError.errorCode, loginError.errorMessage);
            return Promise.reject(result);
        })
            .then(acquireToken);
    }
    return acquireToken();
}
exports.GetToken = GetToken;
/**
 * Log the telemetry data points in the provided {@link IImplicitAuthResult} or Otel pipeline.
 * @param result - the provided {@link IImplicitAuthResult} to log the telemetry data points into
 * @param correlationId - the same correlation ID exists under the caller's context, the same correlation ID will passed on to MSAL.js for unified experience
 * @param applicationId - the application ID in the access token request
 * @param forThirdParty - If true, treats the caller as third-party and avoids sending PII telemetry
 * @param timerClock - the timerClock to log the time duration
 * @param scopes - the scopes in the acquire token request
 * @param errorCode - the error code included in the exception, if any
 * @param errorMessage - the error message included in the exception, if any
 */
function LogTelemetryDataFieldsForGetToken(result, correlationId, applicationId, forThirdParty, timerClock, scopes, errorCode, errorMessage) {
    // For data fields that are null, blank or empty, the value is set to "unknown" at this point
    // based on office-online-otel documentation: https://office.visualstudio.com/OC/_git/office-online-ui?path=%2Fpackages%2Foffice-online-otel%2FREADME.md&version=GBmaster
    // Null or undefined value will cause exceptions down the line.
    var dataFields = [
        { name: Constants.Telemetry.Duration, int64: timerClock.ms },
        { name: Constants.Telemetry.Succeeded, bool: errorCode ? false : true },
        { name: Constants.Telemetry.IdentityProvider, string: authConfig.idp.toLowerCase() },
        { name: Constants.Telemetry.ApplicationId, string: applicationId },
        { name: Constants.Telemetry.TokenScope, string: (scopes && !forThirdParty) ? scopes.toString() : 'unknown' },
        { name: Constants.Telemetry.CorrelationId, string: correlationId ? correlationId : 'unknown' },
        { name: Constants.Telemetry.ErrorCodeForGetToken, string: errorCode ? errorCode : 'unknown' },
        { name: Constants.Telemetry.ErrorMessageForGetToken, string: (errorMessage && !forThirdParty) ? errorMessage : 'unknown' },
    ];
    if (!authConfig.telemetryInstance && typeof OTel === "undefined") {
        api_js_1.sendTelemetryEvent({
            name: Constants.Telemetry.GetTokenTelemetryName,
            dataFields: dataFields
        });
    }
    result.Telemetry = {
        timeToGetToken: timerClock.ms,
        succeeded: errorCode ? false : true,
        idp: authConfig.idp.toLowerCase(),
        applicationId: applicationId,
        tokenScope: scopes ? scopes.toString() : undefined,
        correlationId: correlationId,
        errorCodeForGetToken: errorCode ? errorCode : undefined,
        errorMessageForGetToken: errorMessage ? errorMessage : undefined
    };
}
/**
 * Construct the Msal.UserAgentApplication instance for V2 endpoint calls.
 * @param applicationId - the application ID used to find or construct the MSAL instance.
 * @returns the Msal.UserAgentApplication instance to make calls to V2 endpoint.
 */
function GetApplication(applicationId) {
    var application = undefined;
    applications.some(function (value) {
        if (applicationId && value.applicationId && applicationId.toUpperCase() === value.applicationId.toUpperCase()) {
            application = value.application;
            return true;
        }
        return false;
    });
    if (!application) {
        application = new Msal.UserAgentApplication({
            auth: {
                clientId: applicationId,
                authority: authConfig.authority,
                redirectUri: (authConfig.redirectUri) ? authConfig.redirectUri.split("?")[0] : location.href.split("?")[0],
                navigateToLoginRequestUrl: (authConfig.navigateToLoginRequestUrl) ? authConfig.navigateToLoginRequestUrl : true,
            },
            cache: {
                cacheLocation: 'localStorage',
                // Store auth state in cookies can make the request too big and fail the request sometimes, need to keep it as false.
                storeAuthStateInCookie: false
            },
            system: {
                loadFrameTimeout: (authConfig.timeout) ? authConfig.timeout : 6000,
            },
        });
        var entry = { applicationId: applicationId, application: application };
        applications.push(entry);
    }
    return application;
}
/**
 * Construct the scope for V2 endpoint calls.
 * @param target - resource for V1 token, scope for V2 token
 * @returns the right scope to make calls to V2 endpoint.
 */
function GetScope(target) {
    // To consume V2 endpoint, "/.default" needs to be added for given resources.
    var resourcePrefix = ["HTTPS:", "API:"];
    if (resourcePrefix.some(function (prefix) { return target.toLocaleUpperCase().startsWith(prefix); }) || IsGuid(target)) {
        return target + "/.default";
    }
    // Other cases could be that it is acquiring V2 tokens with scopes "ConnectedServices.ReadWrite" etc
    // or wl.skydrive
    return target;
}
/**
 * Check whether the given string is in GUID format or not.
 * @param str - provided string for format checking.
 * @returns true if the string is in GUID format, returns false otherwise.
 */
function IsGuid(str) {
    // Checking GUID based on the GUID format
    var regexGuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
    return regexGuid.test(str);
}
/**
 * Acquire an Id token by given upn/target, and check if the upn matches context
 * @param correlationId - the same correlation ID exists under the caller's context, the same correlation ID will passed on to MSAL.js for unified experience
 * @returns {Promise.<IImplicitUpnResult>} - a promise that is fulfilled when this function has completed, or rejected if an error was raised. Returns the {@link IImplicitUpnResult} object
 */
function CheckUpnMatchIdToken(applicationId, correlationId) {
    var timerClock = TimerUtils.timer();
    // Wrong format of correlation ID or blank are not valid in MSAL.js
    // With an invalid correlation ID in the request, the access token acquiring request will be rejected by MSAL.js with exceptions.
    // Correlation ID will be set to undefined in those cases and MSAL.js will generate a new correlation ID if it is undefined.
    if (!correlationId || !IsGuid(correlationId)) {
        correlationId = undefined;
    }
    var result = {};
    if (!applicationId || !IsGuid(applicationId)) {
        result.ErrorCode = 'invalid_application_ID';
        result.ErrorMessage = 'The provided application ID for Implicit.CheckUpnMatchIdToken is null, blank, empty or with invalid format';
        LogTelemetryDataFieldsForCheckUpn(result, correlationId, applicationId, timerClock, undefined, result.ErrorCode, result.ErrorMessage);
        return Promise.reject(result);
    }
    var application = GetApplication(applicationId);
    var scopes = [application.config.auth.clientId]; /*It will acquire id token instead of access token if resource/scope is clientId*/
    return new Promise(function (resolve, reject) {
        LoggingUtils.log("Config: " + JSON.stringify(application), enableConsoleLog);
        LoggingUtils.log("application calls acquireTokenSilent", enableConsoleLog);
        application.acquireTokenSilent({
            scopes: scopes,
            loginHint: authConfig.upn,
            correlationId: correlationId,
        }).then(function (authResponse) {
            result.IsUpnMatch = UpnMatchesUpnFromIdToken(authResponse.account);
            if (result.IsUpnMatch) {
                LogTelemetryDataFieldsForCheckUpn(result, correlationId, applicationId, timerClock, scopes);
                return resolve(result);
            }
        }).catch(function (error) {
            result.ErrorCode = error.errorCode;
            result.ErrorMessage = error.errorMessage;
            LoggingUtils.log("acquireToken->error: " + result.ErrorMessage, enableConsoleLog);
            result.IsUpnMatch = false;
            LogTelemetryDataFieldsForCheckUpn(result, correlationId, applicationId, timerClock, scopes, error.errorCode, error.errorMessage);
            return reject(result);
        });
    });
}
exports.CheckUpnMatchIdToken = CheckUpnMatchIdToken;
/**
 * * Log the telemetry data points in the provided {@link IImplicitUpnResult} or Otel pipeline.
 * @param result - the provided {@link IImplicitUpnResult} to log the telemetry data points into
 * @param correlationId - the same correlation ID exists under the caller's context, the same correlation ID will passed on to MSAL.js for unified experience
 * @param applicationId - the application ID in the access token request
 * @param timerClock - the timerClock to log the time duration
 * @param scopes - the scopes in the acquire token request
 * @param errorCode - the error code included in the exception, if any
 * @param errorMessage - the error message included in the exception, if any
 */
function LogTelemetryDataFieldsForCheckUpn(result, correlationId, applicationId, timerClock, scopes, errorCode, errorMessage) {
    // For data fields that are null, blank or empty, the value is set to "unknown" at this point
    // based on office-online-otel documentation: https://office.visualstudio.com/OC/_git/office-online-ui?path=%2Fpackages%2Foffice-online-otel%2FREADME.md&version=GBmaster
    // Null or undefined value will cause exceptions down the line.
    var dataFields = [
        { name: Constants.Telemetry.Duration, int64: timerClock.ms },
        { name: Constants.Telemetry.Succeeded, bool: errorCode ? false : true },
        { name: Constants.Telemetry.IdentityProvider, string: authConfig.idp.toLowerCase() },
        { name: Constants.Telemetry.ApplicationId, string: applicationId },
        { name: Constants.Telemetry.TokenScope, string: scopes ? scopes.toString() : 'unknown' },
        { name: Constants.Telemetry.CorrelationId, string: correlationId ? correlationId : 'unknown' },
        { name: Constants.Telemetry.ErrorCodeForCheckUpn, string: errorCode ? errorCode : 'unknown' },
        { name: Constants.Telemetry.ErrorMessageForCheckUpn, string: errorMessage ? errorMessage : 'unknown' }
    ];
    if (!authConfig.telemetryInstance && typeof OTel === "undefined") {
        api_js_1.sendTelemetryEvent({
            name: Constants.Telemetry.CheckUpnTelemetryName,
            dataFields: dataFields
        });
    }
    result.Telemetry = {
        timeToCheckUPN: timerClock.ms,
        succeeded: errorCode ? false : true,
        idp: authConfig.idp.toLowerCase(),
        applicationId: applicationId,
        tokenScope: scopes ? scopes.toString() : undefined,
        correlationId: correlationId,
        errorCodeForCheckUPN: errorCode,
        errorMessageForCheckUPN: errorMessage
    };
}
/**
 * Get Authentication Config from parent.
 * This must be called if page is created by MSAL.js
 * @returns {Promise.<IImplicitAuthConfig>} - a promise that is fulfilled when this function has completed, or rejected if an error was raised. Returns the {@link IImplicitAuthConfig} object
 */
function GetAuthConfig() {
    return new Promise(function (resolve, reject) {
        HashHasState().then(function (result) {
            // Receiver ResponseAuthConfig
            window.addEventListener('message', function (e) {
                if (e.origin && e.origin == location.origin && e.data && e.data.config && e.data.request && e.data.request == Constants.PostMessageType.ResponseAuthConfig) {
                    resolve(e.data.config);
                }
            }, false);
            // Requester RequestAuthContext
            parent.postMessage({ request: Constants.PostMessageType.RequestAuthConfig, iframe: result }, location.origin);
        }, function () { reject({}); });
    });
}
exports.GetAuthConfig = GetAuthConfig;
/**
 * Verify the upn in the config matches the upn for the cached user
 * @param account - the account information included in ID token.
 * @returns true if there is a match or there is no upn in the config
 */
function UpnMatchesUpnFromIdToken(account) {
    if (!authConfig || !authConfig.upn) {
        LoggingUtils.log('Upn does not exist in the configuration, returning true', enableConsoleLog);
        return true;
    }
    if (account && account.userName && account.userName.toLowerCase() === authConfig.upn.toLowerCase()) {
        return true;
    }
    LoggingUtils.log('Upn in config does not match cached user upn', enableConsoleLog);
    return false;
}
/**
 * Verify the upn in the id token matches the Upn in the access token
 * @param token - the token to extract the upn from
 * @returns true if there is a match or there is no upn in the id token or access token
 */
function UpnMatchesUpnFromAccessToken(token) {
    if (!authConfig || !authConfig.upn) {
        LoggingUtils.log('Upn does not exist in the configuration, returning true', enableConsoleLog);
        return true;
    }
    var accessToken = ExtractUtils.extractAccessToken(token);
    // AccessToken extraction would not work for future encrypted JWE tokens,
    // If cannot be extracted, also return true
    if (!accessToken || (accessToken && accessToken.upn && accessToken.upn.toLowerCase() === authConfig.upn.toLowerCase())) {
        return true;
    }
    LoggingUtils.log('provided Upn does not match Upn extracted from token', enableConsoleLog);
    return false;
}
/**
 * Verify the state is in hash or not
 */
function HashHasState() {
    return new Promise(function (resolve, reject) {
        if (window.location.hash) {
            // Get hash
            var hash = window.location.hash;
            if (hash.indexOf('#/') > -1) {
                hash = hash.substring(hash.indexOf('#/') + 2);
            }
            else if (hash.indexOf('#') > -1) {
                hash = hash.substring(1);
            }
            // Get state from hash
            var arrHash = hash.split('&');
            for (var i = 0; i < arrHash.length; i++) {
                var keyvalue = arrHash[i].split('=');
                if (decodeURIComponent(keyvalue[0]) == "state") {
                    // State information or format can be changed by msal.js over time,
                    // but the last item is always the user provided state, which is the parent iframe url in this case.
                    var state = decodeURIComponent(keyvalue[keyvalue.length - 1]).split('|');
                    if (state.length == 2) {
                        resolve(state[1]);
                    }
                }
            }
        }
        reject();
    });
}
/**
 * Logout current user from all registered applications
 */
function Logout() {
    applications.forEach(function (entry) {
        LoggingUtils.log("application calls logOut", enableConsoleLog);
        entry.application.logOut();
    });
    // Also clear registered applications
    applications = new Array();
}
exports.Logout = Logout;


/***/ }),

/***/ 0:
/***/ (function(module, exports, __webpack_require__) {

__webpack_require__("./packages/Microsoft.Office.WebAuth.Implicit/lib/msal.min.js");
__webpack_require__("./packages/Microsoft.Office.WebAuth.Implicit/lib/api.js");
module.exports = __webpack_require__("./packages/Microsoft.Office.WebAuth.Implicit/scripts/Implicit.ts");


/***/ })

/******/ });
});