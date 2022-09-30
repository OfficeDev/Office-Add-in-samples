/* Outlook Web specific API library */
/* osfweb version: 16.0.14516.10000 */
/* office-js-api version: 20210901.L */
/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/
/*
    Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.

    This file also contains the following Promise implementation (with a few small modifications):
        * @overview es6-promise - a tiny implementation of Promises/A+.
        * @copyright Copyright (c) 2014 Yehuda Katz, Tom Dale, Stefan Penner and contributors (Conversion to ES6 API by Jake Archibald)
        * @license   Licensed under MIT license
        *            See https://raw.githubusercontent.com/jakearchibald/es6-promise/master/LICENSE
        * @version   2.3.0
*/
if (typeof OSFPerformance !== "undefined") {
    OSFPerformance.hostInitializationStart = OSFPerformance.now();
}

/* Outlook OWA specific API library */
/* Version: 16.0.14516.10000 */
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    }
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var OfficeExt;
(function (OfficeExt) {
    var MicrosoftAjaxFactory = (function () {
        function MicrosoftAjaxFactory() {
        }
        MicrosoftAjaxFactory.prototype.isMsAjaxLoaded = function () {
            if (typeof (Sys) !== 'undefined' && typeof (Type) !== 'undefined' &&
                Sys.StringBuilder && typeof (Sys.StringBuilder) === "function" &&
                Type.registerNamespace && typeof (Type.registerNamespace) === "function" &&
                Type.registerClass && typeof (Type.registerClass) === "function" &&
                typeof (Function._validateParams) === "function" &&
                Sys.Serialization && Sys.Serialization.JavaScriptSerializer && typeof (Sys.Serialization.JavaScriptSerializer.serialize) === "function") {
                return true;
            }
            else {
                return false;
            }
        };
        MicrosoftAjaxFactory.prototype.loadMsAjaxFull = function (callback) {
            var msAjaxCDNPath = (window.location.protocol.toLowerCase() === 'https:' ? 'https:' : 'http:') + '//ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js';
            OSF.OUtil.loadScript(msAjaxCDNPath, callback);
        };
        Object.defineProperty(MicrosoftAjaxFactory.prototype, "msAjaxError", {
            get: function () {
                if (this._msAjaxError == null && this.isMsAjaxLoaded()) {
                    this._msAjaxError = Error;
                }
                return this._msAjaxError;
            },
            set: function (errorClass) {
                this._msAjaxError = errorClass;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(MicrosoftAjaxFactory.prototype, "msAjaxString", {
            get: function () {
                if (this._msAjaxString == null && this.isMsAjaxLoaded()) {
                    this._msAjaxString = String;
                }
                return this._msAjaxString;
            },
            set: function (stringClass) {
                this._msAjaxString = stringClass;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(MicrosoftAjaxFactory.prototype, "msAjaxDebug", {
            get: function () {
                if (this._msAjaxDebug == null && this.isMsAjaxLoaded()) {
                    this._msAjaxDebug = Sys.Debug;
                }
                return this._msAjaxDebug;
            },
            set: function (debugClass) {
                this._msAjaxDebug = debugClass;
            },
            enumerable: true,
            configurable: true
        });
        return MicrosoftAjaxFactory;
    }());
    OfficeExt.MicrosoftAjaxFactory = MicrosoftAjaxFactory;
})(OfficeExt || (OfficeExt = {}));
var OsfMsAjaxFactory = new OfficeExt.MicrosoftAjaxFactory();
var OSF = OSF || {};
(function (OfficeExt) {
    var SafeStorage = (function () {
        function SafeStorage(_internalStorage) {
            this._internalStorage = _internalStorage;
        }
        SafeStorage.prototype.getItem = function (key) {
            try {
                return this._internalStorage && this._internalStorage.getItem(key);
            }
            catch (e) {
                return null;
            }
        };
        SafeStorage.prototype.setItem = function (key, data) {
            try {
                this._internalStorage && this._internalStorage.setItem(key, data);
            }
            catch (e) {
            }
        };
        SafeStorage.prototype.clear = function () {
            try {
                this._internalStorage && this._internalStorage.clear();
            }
            catch (e) {
            }
        };
        SafeStorage.prototype.removeItem = function (key) {
            try {
                this._internalStorage && this._internalStorage.removeItem(key);
            }
            catch (e) {
            }
        };
        SafeStorage.prototype.getKeysWithPrefix = function (keyPrefix) {
            var keyList = [];
            try {
                var len = this._internalStorage && this._internalStorage.length || 0;
                for (var i = 0; i < len; i++) {
                    var key = this._internalStorage.key(i);
                    if (key.indexOf(keyPrefix) === 0) {
                        keyList.push(key);
                    }
                }
            }
            catch (e) {
            }
            return keyList;
        };
        SafeStorage.prototype.isLocalStorageAvailable = function () {
            return (this._internalStorage != null);
        };
        return SafeStorage;
    }());
    OfficeExt.SafeStorage = SafeStorage;
})(OfficeExt || (OfficeExt = {}));
OSF.XdmFieldName = {
    ConversationUrl: "ConversationUrl",
    AppId: "AppId"
};
OSF.TestFlightStart = 1000;
OSF.TestFlightEnd = 1009;
OSF.FlightNames = {
    UseOriginNotUrl: 0,
    AddinEnforceHttps: 2,
    FirstPartyAnonymousProxyReadyCheckTimeout: 6,
    AddinRibbonIdAllowUnknown: 9,
    ManifestParserDevConsoleLog: 15,
    AddinActionDefinitionHybridMode: 18,
    UseActionIdForUILessCommand: 20,
    RequirementSetRibbonApiOnePointTwo: 21,
    SetFocusToTaskpaneIsEnabled: 22,
    ShortcutInfoArrayInUserPreferenceData: 23,
    OSFTestFlight1000: OSF.TestFlightStart,
    OSFTestFlight1001: OSF.TestFlightStart + 1,
    OSFTestFlight1002: OSF.TestFlightStart + 2,
    OSFTestFlight1003: OSF.TestFlightStart + 3,
    OSFTestFlight1004: OSF.TestFlightStart + 4,
    OSFTestFlight1005: OSF.TestFlightStart + 5,
    OSFTestFlight1006: OSF.TestFlightStart + 6,
    OSFTestFlight1007: OSF.TestFlightStart + 7,
    OSFTestFlight1008: OSF.TestFlightStart + 8,
    OSFTestFlight1009: OSF.TestFlightEnd
};
OSF.FlightTreatmentNames = {
    AllowStorageAccessByUserActivationOnIFrame: "Microsoft.Office.SharedOnline.AllowStorageAccessByUserActivationOnIFrame",
    IsPrivateAddin: "Microsoft.Office.SharedOnline.IsPrivateAddin",
    LogAllAddinsAsPublic: "Microsoft.Office.SharedOnline.LogAllAddinsAsPublic",
    AddinCommandRibbonCacheFixEnabled: "Microsoft.Office.SharedOnline.AddinCommandRibbonCacheFixEnabled",
    OSFSolutionRefactor: "Microsoft.Office.SharedOnline.OSFSolutionRefactor"
};
OSF.Flights = [];
OSF.Settings = {};
OSF.WindowNameItemKeys = {
    BaseFrameName: "baseFrameName",
    HostInfo: "hostInfo",
    XdmInfo: "xdmInfo",
    SerializerVersion: "serializerVersion",
    AppContext: "appContext",
    Flights: "flights"
};
OSF.OUtil = (function () {
    var _uniqueId = -1;
    var _xdmInfoKey = '&_xdm_Info=';
    var _serializerVersionKey = '&_serializer_version=';
    var _flightsKey = '&_flights=';
    var _xdmSessionKeyPrefix = '_xdm_';
    var _serializerVersionKeyPrefix = '_serializer_version=';
    var _flightsKeyPrefix = '_flights=';
    var _fragmentSeparator = '#';
    var _fragmentInfoDelimiter = '&';
    var _classN = "class";
    var _loadedScripts = {};
    var _defaultScriptLoadingTimeout = 30000;
    var _safeSessionStorage = null;
    var _safeLocalStorage = null;
    var _rndentropy = new Date().getTime();
    function _random() {
        var nextrand = 0x7fffffff * (Math.random());
        nextrand ^= _rndentropy ^ ((new Date().getMilliseconds()) << Math.floor(Math.random() * (31 - 10)));
        return nextrand.toString(16);
    }
    ;
    function _getSessionStorage() {
        if (!_safeSessionStorage) {
            try {
                var sessionStorage = window.sessionStorage;
            }
            catch (ex) {
                sessionStorage = null;
            }
            _safeSessionStorage = new OfficeExt.SafeStorage(sessionStorage);
        }
        return _safeSessionStorage;
    }
    ;
    function _reOrderTabbableElements(elements) {
        var bucket0 = [];
        var bucketPositive = [];
        var i;
        var len = elements.length;
        var ele;
        for (i = 0; i < len; i++) {
            ele = elements[i];
            if (ele.tabIndex) {
                if (ele.tabIndex > 0) {
                    bucketPositive.push(ele);
                }
                else if (ele.tabIndex === 0) {
                    bucket0.push(ele);
                }
            }
            else {
                bucket0.push(ele);
            }
        }
        bucketPositive = bucketPositive.sort(function (left, right) {
            var diff = left.tabIndex - right.tabIndex;
            if (diff === 0) {
                diff = bucketPositive.indexOf(left) - bucketPositive.indexOf(right);
            }
            return diff;
        });
        return [].concat(bucketPositive, bucket0);
    }
    ;
    return {
        set_entropy: function OSF_OUtil$set_entropy(entropy) {
            if (typeof entropy == "string") {
                for (var i = 0; i < entropy.length; i += 4) {
                    var temp = 0;
                    for (var j = 0; j < 4 && i + j < entropy.length; j++) {
                        temp = (temp << 8) + entropy.charCodeAt(i + j);
                    }
                    _rndentropy ^= temp;
                }
            }
            else if (typeof entropy == "number") {
                _rndentropy ^= entropy;
            }
            else {
                _rndentropy ^= 0x7fffffff * Math.random();
            }
            _rndentropy &= 0x7fffffff;
        },
        extend: function OSF_OUtil$extend(child, parent) {
            var F = function () { };
            F.prototype = parent.prototype;
            child.prototype = new F();
            child.prototype.constructor = child;
            child.uber = parent.prototype;
            if (parent.prototype.constructor === Object.prototype.constructor) {
                parent.prototype.constructor = parent;
            }
        },
        setNamespace: function OSF_OUtil$setNamespace(name, parent) {
            if (parent && name && !parent[name]) {
                parent[name] = {};
            }
        },
        unsetNamespace: function OSF_OUtil$unsetNamespace(name, parent) {
            if (parent && name && parent[name]) {
                delete parent[name];
            }
        },
        serializeSettings: function OSF_OUtil$serializeSettings(settingsCollection) {
            var ret = {};
            for (var key in settingsCollection) {
                var value = settingsCollection[key];
                try {
                    if (JSON) {
                        value = JSON.stringify(value, function dateReplacer(k, v) {
                            return OSF.OUtil.isDate(this[k]) ? OSF.DDA.SettingsManager.DateJSONPrefix + this[k].getTime() + OSF.DDA.SettingsManager.DataJSONSuffix : v;
                        });
                    }
                    else {
                        value = Sys.Serialization.JavaScriptSerializer.serialize(value);
                    }
                    ret[key] = value;
                }
                catch (ex) {
                }
            }
            return ret;
        },
        deserializeSettings: function OSF_OUtil$deserializeSettings(serializedSettings) {
            var ret = {};
            serializedSettings = serializedSettings || {};
            for (var key in serializedSettings) {
                var value = serializedSettings[key];
                try {
                    if (JSON) {
                        value = JSON.parse(value, function dateReviver(k, v) {
                            var d;
                            if (typeof v === 'string' && v && v.length > 6 && v.slice(0, 5) === OSF.DDA.SettingsManager.DateJSONPrefix && v.slice(-1) === OSF.DDA.SettingsManager.DataJSONSuffix) {
                                d = new Date(parseInt(v.slice(5, -1)));
                                if (d) {
                                    return d;
                                }
                            }
                            return v;
                        });
                    }
                    else {
                        value = Sys.Serialization.JavaScriptSerializer.deserialize(value, true);
                    }
                    ret[key] = value;
                }
                catch (ex) {
                }
            }
            return ret;
        },
        loadScript: function OSF_OUtil$loadScript(url, callback, timeoutInMs) {
            if (url && callback) {
                var doc = window.document;
                var _loadedScriptEntry = _loadedScripts[url];
                if (!_loadedScriptEntry) {
                    var script = doc.createElement("script");
                    script.type = "text/javascript";
                    _loadedScriptEntry = { loaded: false, pendingCallbacks: [callback], timer: null };
                    _loadedScripts[url] = _loadedScriptEntry;
                    var onLoadCallback = function OSF_OUtil_loadScript$onLoadCallback() {
                        if (_loadedScriptEntry.timer != null) {
                            clearTimeout(_loadedScriptEntry.timer);
                            delete _loadedScriptEntry.timer;
                        }
                        _loadedScriptEntry.loaded = true;
                        var pendingCallbackCount = _loadedScriptEntry.pendingCallbacks.length;
                        for (var i = 0; i < pendingCallbackCount; i++) {
                            var currentCallback = _loadedScriptEntry.pendingCallbacks.shift();
                            currentCallback();
                        }
                    };
                    var onLoadTimeOut = function OSF_OUtil_loadScript$onLoadTimeOut() {
                        if (window.navigator.userAgent.indexOf("Trident") > 0) {
                            onLoadError(null);
                        }
                        else {
                            onLoadError(new Event("Script load timed out"));
                        }
                    };
                    var onLoadError = function OSF_OUtil_loadScript$onLoadError(errorEvent) {
                        delete _loadedScripts[url];
                        if (_loadedScriptEntry.timer != null) {
                            clearTimeout(_loadedScriptEntry.timer);
                            delete _loadedScriptEntry.timer;
                        }
                        var pendingCallbackCount = _loadedScriptEntry.pendingCallbacks.length;
                        for (var i = 0; i < pendingCallbackCount; i++) {
                            var currentCallback = _loadedScriptEntry.pendingCallbacks.shift();
                            currentCallback();
                        }
                    };
                    if (script.readyState) {
                        script.onreadystatechange = function () {
                            if (script.readyState == "loaded" || script.readyState == "complete") {
                                script.onreadystatechange = null;
                                onLoadCallback();
                            }
                        };
                    }
                    else {
                        script.onload = onLoadCallback;
                    }
                    script.onerror = onLoadError;
                    timeoutInMs = timeoutInMs || _defaultScriptLoadingTimeout;
                    _loadedScriptEntry.timer = setTimeout(onLoadTimeOut, timeoutInMs);
                    script.setAttribute("crossOrigin", "anonymous");
                    script.src = url;
                    doc.getElementsByTagName("head")[0].appendChild(script);
                }
                else if (_loadedScriptEntry.loaded) {
                    callback();
                }
                else {
                    _loadedScriptEntry.pendingCallbacks.push(callback);
                }
            }
        },
        loadCSS: function OSF_OUtil$loadCSS(url) {
            if (url) {
                var doc = window.document;
                var link = doc.createElement("link");
                link.type = "text/css";
                link.rel = "stylesheet";
                link.href = url;
                doc.getElementsByTagName("head")[0].appendChild(link);
            }
        },
        parseEnum: function OSF_OUtil$parseEnum(str, enumObject) {
            var parsed = enumObject[str.trim()];
            if (typeof (parsed) == 'undefined') {
                OsfMsAjaxFactory.msAjaxDebug.trace("invalid enumeration string:" + str);
                throw OsfMsAjaxFactory.msAjaxError.argument("str");
            }
            return parsed;
        },
        delayExecutionAndCache: function OSF_OUtil$delayExecutionAndCache() {
            var obj = { calc: arguments[0] };
            return function () {
                if (obj.calc) {
                    obj.val = obj.calc.apply(this, arguments);
                    delete obj.calc;
                }
                return obj.val;
            };
        },
        getUniqueId: function OSF_OUtil$getUniqueId() {
            _uniqueId = _uniqueId + 1;
            return _uniqueId.toString();
        },
        formatString: function OSF_OUtil$formatString() {
            var args = arguments;
            var source = args[0];
            return source.replace(/{(\d+)}/gm, function (match, number) {
                var index = parseInt(number, 10) + 1;
                return args[index] === undefined ? '{' + number + '}' : args[index];
            });
        },
        generateConversationId: function OSF_OUtil$generateConversationId() {
            return [_random(), _random(), (new Date()).getTime().toString()].join('_');
        },
        getFrameName: function OSF_OUtil$getFrameName(cacheKey) {
            return _xdmSessionKeyPrefix + cacheKey + this.generateConversationId();
        },
        addXdmInfoAsHash: function OSF_OUtil$addXdmInfoAsHash(url, xdmInfoValue) {
            return OSF.OUtil.addInfoAsHash(url, _xdmInfoKey, xdmInfoValue, false);
        },
        addSerializerVersionAsHash: function OSF_OUtil$addSerializerVersionAsHash(url, serializerVersion) {
            return OSF.OUtil.addInfoAsHash(url, _serializerVersionKey, serializerVersion, true);
        },
        addFlightsAsHash: function OSF_OUtil$addFlightsAsHash(url, flights) {
            return OSF.OUtil.addInfoAsHash(url, _flightsKey, flights, true);
        },
        addInfoAsHash: function OSF_OUtil$addInfoAsHash(url, keyName, infoValue, encodeInfo) {
            url = url.trim() || '';
            var urlParts = url.split(_fragmentSeparator);
            var urlWithoutFragment = urlParts.shift();
            var fragment = urlParts.join(_fragmentSeparator);
            var newFragment;
            if (encodeInfo) {
                newFragment = [keyName, encodeURIComponent(infoValue), fragment].join('');
            }
            else {
                newFragment = [fragment, keyName, infoValue].join('');
            }
            return [urlWithoutFragment, _fragmentSeparator, newFragment].join('');
        },
        parseHostInfoFromWindowName: function OSF_OUtil$parseHostInfoFromWindowName(skipSessionStorage, windowName) {
            return OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.HostInfo);
        },
        parseXdmInfo: function OSF_OUtil$parseXdmInfo(skipSessionStorage) {
            var xdmInfoValue = OSF.OUtil.parseXdmInfoWithGivenFragment(skipSessionStorage, window.location.hash);
            if (!xdmInfoValue) {
                xdmInfoValue = OSF.OUtil.parseXdmInfoFromWindowName(skipSessionStorage, window.name);
            }
            return xdmInfoValue;
        },
        parseXdmInfoFromWindowName: function OSF_OUtil$parseXdmInfoFromWindowName(skipSessionStorage, windowName) {
            return OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.XdmInfo);
        },
        parseXdmInfoWithGivenFragment: function OSF_OUtil$parseXdmInfoWithGivenFragment(skipSessionStorage, fragment) {
            return OSF.OUtil.parseInfoWithGivenFragment(_xdmInfoKey, _xdmSessionKeyPrefix, false, skipSessionStorage, fragment);
        },
        parseSerializerVersion: function OSF_OUtil$parseSerializerVersion(skipSessionStorage) {
            var serializerVersion = OSF.OUtil.parseSerializerVersionWithGivenFragment(skipSessionStorage, window.location.hash);
            if (isNaN(serializerVersion)) {
                serializerVersion = OSF.OUtil.parseSerializerVersionFromWindowName(skipSessionStorage, window.name);
            }
            return serializerVersion;
        },
        parseSerializerVersionFromWindowName: function OSF_OUtil$parseSerializerVersionFromWindowName(skipSessionStorage, windowName) {
            return parseInt(OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.SerializerVersion));
        },
        parseSerializerVersionWithGivenFragment: function OSF_OUtil$parseSerializerVersionWithGivenFragment(skipSessionStorage, fragment) {
            return parseInt(OSF.OUtil.parseInfoWithGivenFragment(_serializerVersionKey, _serializerVersionKeyPrefix, true, skipSessionStorage, fragment));
        },
        parseFlights: function OSF_OUtil$parseFlights(skipSessionStorage) {
            var flights = OSF.OUtil.parseFlightsWithGivenFragment(skipSessionStorage, window.location.hash);
            if (flights.length == 0) {
                flights = OSF.OUtil.parseFlightsFromWindowName(skipSessionStorage, window.name);
            }
            return flights;
        },
        checkFlight: function OSF_OUtil$checkFlightEnabled(flight) {
            return OSF.Flights && OSF.Flights.indexOf(flight) >= 0;
        },
        pushFlight: function OSF_OUtil$pushFlight(flight) {
            if (OSF.Flights.indexOf(flight) < 0) {
                OSF.Flights.push(flight);
                return true;
            }
            return false;
        },
        getBooleanSetting: function OSF_OUtil$getSetting(settingName) {
            return OSF.OUtil.getBooleanFromDictionary(OSF.Settings, settingName);
        },
        getBooleanFromDictionary: function OSF_OUtil$getBooleanFromDictionary(settings, settingName) {
            var result = (settings && settingName && settings[settingName] !== undefined && settings[settingName] &&
                ((typeof (settings[settingName]) === "string" && settings[settingName].toUpperCase() === 'TRUE') ||
                    (typeof (settings[settingName]) === "boolean" && settings[settingName])));
            return result !== undefined ? result : false;
        },
        parseFlightsFromWindowName: function OSF_OUtil$parseFlightsFromWindowName(skipSessionStorage, windowName) {
            return OSF.OUtil.parseArrayWithDefault(OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.Flights));
        },
        parseFlightsWithGivenFragment: function OSF_OUtil$parseFlightsWithGivenFragment(skipSessionStorage, fragment) {
            return OSF.OUtil.parseArrayWithDefault(OSF.OUtil.parseInfoWithGivenFragment(_flightsKey, _flightsKeyPrefix, true, skipSessionStorage, fragment));
        },
        parseArrayWithDefault: function OSF_OUtil$parseArrayWithDefault(jsonString) {
            var array = [];
            try {
                array = JSON.parse(jsonString);
            }
            catch (ex) { }
            if (!Array.isArray(array)) {
                array = [];
            }
            return array;
        },
        parseInfoFromWindowName: function OSF_OUtil$parseInfoFromWindowName(skipSessionStorage, windowName, infoKey) {
            try {
                var windowNameObj = JSON.parse(windowName);
                var infoValue = windowNameObj != null ? windowNameObj[infoKey] : null;
                var osfSessionStorage = _getSessionStorage();
                if (!skipSessionStorage && osfSessionStorage && windowNameObj != null) {
                    var sessionKey = windowNameObj[OSF.WindowNameItemKeys.BaseFrameName] + infoKey;
                    if (infoValue) {
                        osfSessionStorage.setItem(sessionKey, infoValue);
                    }
                    else {
                        infoValue = osfSessionStorage.getItem(sessionKey);
                    }
                }
                return infoValue;
            }
            catch (Exception) {
                return null;
            }
        },
        parseInfoWithGivenFragment: function OSF_OUtil$parseInfoWithGivenFragment(infoKey, infoKeyPrefix, decodeInfo, skipSessionStorage, fragment) {
            var fragmentParts = fragment.split(infoKey);
            var infoValue = fragmentParts.length > 1 ? fragmentParts[fragmentParts.length - 1] : null;
            if (decodeInfo && infoValue != null) {
                if (infoValue.indexOf(_fragmentInfoDelimiter) >= 0) {
                    infoValue = infoValue.split(_fragmentInfoDelimiter)[0];
                }
                infoValue = decodeURIComponent(infoValue);
            }
            var osfSessionStorage = _getSessionStorage();
            if (!skipSessionStorage && osfSessionStorage) {
                var sessionKeyStart = window.name.indexOf(infoKeyPrefix);
                if (sessionKeyStart > -1) {
                    var sessionKeyEnd = window.name.indexOf(";", sessionKeyStart);
                    if (sessionKeyEnd == -1) {
                        sessionKeyEnd = window.name.length;
                    }
                    var sessionKey = window.name.substring(sessionKeyStart, sessionKeyEnd);
                    if (infoValue) {
                        osfSessionStorage.setItem(sessionKey, infoValue);
                    }
                    else {
                        infoValue = osfSessionStorage.getItem(sessionKey);
                    }
                }
            }
            return infoValue;
        },
        getConversationId: function OSF_OUtil$getConversationId() {
            var searchString = window.location.search;
            var conversationId = null;
            if (searchString) {
                var index = searchString.indexOf("&");
                conversationId = index > 0 ? searchString.substring(1, index) : searchString.substr(1);
                if (conversationId && conversationId.charAt(conversationId.length - 1) === '=') {
                    conversationId = conversationId.substring(0, conversationId.length - 1);
                    if (conversationId) {
                        conversationId = decodeURIComponent(conversationId);
                    }
                }
            }
            return conversationId;
        },
        getInfoItems: function OSF_OUtil$getInfoItems(strInfo) {
            var items = strInfo.split("$");
            if (typeof items[1] == "undefined") {
                items = strInfo.split("|");
            }
            if (typeof items[1] == "undefined") {
                items = strInfo.split("%7C");
            }
            return items;
        },
        getXdmFieldValue: function OSF_OUtil$getXdmFieldValue(xdmFieldName, skipSessionStorage) {
            var fieldValue = '';
            var xdmInfoValue = OSF.OUtil.parseXdmInfo(skipSessionStorage);
            if (xdmInfoValue) {
                var items = OSF.OUtil.getInfoItems(xdmInfoValue);
                if (items != undefined && items.length >= 3) {
                    switch (xdmFieldName) {
                        case OSF.XdmFieldName.ConversationUrl:
                            fieldValue = items[2];
                            break;
                        case OSF.XdmFieldName.AppId:
                            fieldValue = items[1];
                            break;
                    }
                }
            }
            return fieldValue;
        },
        validateParamObject: function OSF_OUtil$validateParamObject(params, expectedProperties, callback) {
            var e = Function._validateParams(arguments, [{ name: "params", type: Object, mayBeNull: false },
                { name: "expectedProperties", type: Object, mayBeNull: false },
                { name: "callback", type: Function, mayBeNull: true }
            ]);
            if (e)
                throw e;
            for (var p in expectedProperties) {
                e = Function._validateParameter(params[p], expectedProperties[p], p);
                if (e)
                    throw e;
            }
        },
        writeProfilerMark: function OSF_OUtil$writeProfilerMark(text) {
            if (window.msWriteProfilerMark) {
                window.msWriteProfilerMark(text);
                OsfMsAjaxFactory.msAjaxDebug.trace(text);
            }
        },
        outputDebug: function OSF_OUtil$outputDebug(text) {
            if (typeof (OsfMsAjaxFactory) !== 'undefined' && OsfMsAjaxFactory.msAjaxDebug && OsfMsAjaxFactory.msAjaxDebug.trace) {
                OsfMsAjaxFactory.msAjaxDebug.trace(text);
            }
        },
        defineNondefaultProperty: function OSF_OUtil$defineNondefaultProperty(obj, prop, descriptor, attributes) {
            descriptor = descriptor || {};
            for (var nd in attributes) {
                var attribute = attributes[nd];
                if (descriptor[attribute] == undefined) {
                    descriptor[attribute] = true;
                }
            }
            Object.defineProperty(obj, prop, descriptor);
            return obj;
        },
        defineNondefaultProperties: function OSF_OUtil$defineNondefaultProperties(obj, descriptors, attributes) {
            descriptors = descriptors || {};
            for (var prop in descriptors) {
                OSF.OUtil.defineNondefaultProperty(obj, prop, descriptors[prop], attributes);
            }
            return obj;
        },
        defineEnumerableProperty: function OSF_OUtil$defineEnumerableProperty(obj, prop, descriptor) {
            return OSF.OUtil.defineNondefaultProperty(obj, prop, descriptor, ["enumerable"]);
        },
        defineEnumerableProperties: function OSF_OUtil$defineEnumerableProperties(obj, descriptors) {
            return OSF.OUtil.defineNondefaultProperties(obj, descriptors, ["enumerable"]);
        },
        defineMutableProperty: function OSF_OUtil$defineMutableProperty(obj, prop, descriptor) {
            return OSF.OUtil.defineNondefaultProperty(obj, prop, descriptor, ["writable", "enumerable", "configurable"]);
        },
        defineMutableProperties: function OSF_OUtil$defineMutableProperties(obj, descriptors) {
            return OSF.OUtil.defineNondefaultProperties(obj, descriptors, ["writable", "enumerable", "configurable"]);
        },
        finalizeProperties: function OSF_OUtil$finalizeProperties(obj, descriptor) {
            descriptor = descriptor || {};
            var props = Object.getOwnPropertyNames(obj);
            var propsLength = props.length;
            for (var i = 0; i < propsLength; i++) {
                var prop = props[i];
                var desc = Object.getOwnPropertyDescriptor(obj, prop);
                if (!desc.get && !desc.set) {
                    desc.writable = descriptor.writable || false;
                }
                desc.configurable = descriptor.configurable || false;
                desc.enumerable = descriptor.enumerable || true;
                Object.defineProperty(obj, prop, desc);
            }
            return obj;
        },
        mapList: function OSF_OUtil$MapList(list, mapFunction) {
            var ret = [];
            if (list) {
                for (var item in list) {
                    ret.push(mapFunction(list[item]));
                }
            }
            return ret;
        },
        listContainsKey: function OSF_OUtil$listContainsKey(list, key) {
            for (var item in list) {
                if (key == item) {
                    return true;
                }
            }
            return false;
        },
        listContainsValue: function OSF_OUtil$listContainsElement(list, value) {
            for (var item in list) {
                if (value == list[item]) {
                    return true;
                }
            }
            return false;
        },
        augmentList: function OSF_OUtil$augmentList(list, addenda) {
            var add = list.push ? function (key, value) { list.push(value); } : function (key, value) { list[key] = value; };
            for (var key in addenda) {
                add(key, addenda[key]);
            }
        },
        redefineList: function OSF_Outil$redefineList(oldList, newList) {
            for (var key1 in oldList) {
                delete oldList[key1];
            }
            for (var key2 in newList) {
                oldList[key2] = newList[key2];
            }
        },
        isArray: function OSF_OUtil$isArray(obj) {
            return Object.prototype.toString.apply(obj) === "[object Array]";
        },
        isFunction: function OSF_OUtil$isFunction(obj) {
            return Object.prototype.toString.apply(obj) === "[object Function]";
        },
        isDate: function OSF_OUtil$isDate(obj) {
            return Object.prototype.toString.apply(obj) === "[object Date]";
        },
        addEventListener: function OSF_OUtil$addEventListener(element, eventName, listener) {
            if (element.addEventListener) {
                element.addEventListener(eventName, listener, false);
            }
            else if ((Sys.Browser.agent === Sys.Browser.InternetExplorer) && element.attachEvent) {
                element.attachEvent("on" + eventName, listener);
            }
            else {
                element["on" + eventName] = listener;
            }
        },
        removeEventListener: function OSF_OUtil$removeEventListener(element, eventName, listener) {
            if (element.removeEventListener) {
                element.removeEventListener(eventName, listener, false);
            }
            else if ((Sys.Browser.agent === Sys.Browser.InternetExplorer) && element.detachEvent) {
                element.detachEvent("on" + eventName, listener);
            }
            else {
                element["on" + eventName] = null;
            }
        },
        getCookieValue: function OSF_OUtil$getCookieValue(cookieName) {
            var tmpCookieString = RegExp(cookieName + "[^;]+").exec(document.cookie);
            return tmpCookieString.toString().replace(/^[^=]+./, "");
        },
        xhrGet: function OSF_OUtil$xhrGet(url, onSuccess, onError) {
            var xmlhttp;
            try {
                xmlhttp = new XMLHttpRequest();
                xmlhttp.onreadystatechange = function () {
                    if (xmlhttp.readyState == 4) {
                        if (xmlhttp.status == 200) {
                            onSuccess(xmlhttp.responseText);
                        }
                        else {
                            onError(xmlhttp.status);
                        }
                    }
                };
                xmlhttp.open("GET", url, true);
                xmlhttp.send();
            }
            catch (ex) {
                onError(ex);
            }
        },
        xhrGetFull: function OSF_OUtil$xhrGetFull(url, oneDriveFileName, onSuccess, onError) {
            var xmlhttp;
            var requestedFileName = oneDriveFileName;
            try {
                xmlhttp = new XMLHttpRequest();
                xmlhttp.onreadystatechange = function () {
                    if (xmlhttp.readyState == 4) {
                        if (xmlhttp.status == 200) {
                            onSuccess(xmlhttp, requestedFileName);
                        }
                        else {
                            onError(xmlhttp.status);
                        }
                    }
                };
                xmlhttp.open("GET", url, true);
                xmlhttp.send();
            }
            catch (ex) {
                onError(ex);
            }
        },
        encodeBase64: function OSF_Outil$encodeBase64(input) {
            if (!input)
                return input;
            var codex = "ABCDEFGHIJKLMNOP" + "QRSTUVWXYZabcdef" + "ghijklmnopqrstuv" + "wxyz0123456789+/=";
            var output = [];
            var temp = [];
            var index = 0;
            var c1, c2, c3, a, b, c;
            var i;
            var length = input.length;
            do {
                c1 = input.charCodeAt(index++);
                c2 = input.charCodeAt(index++);
                c3 = input.charCodeAt(index++);
                i = 0;
                a = c1 & 255;
                b = c1 >> 8;
                c = c2 & 255;
                temp[i++] = a >> 2;
                temp[i++] = ((a & 3) << 4) | (b >> 4);
                temp[i++] = ((b & 15) << 2) | (c >> 6);
                temp[i++] = c & 63;
                if (!isNaN(c2)) {
                    a = c2 >> 8;
                    b = c3 & 255;
                    c = c3 >> 8;
                    temp[i++] = a >> 2;
                    temp[i++] = ((a & 3) << 4) | (b >> 4);
                    temp[i++] = ((b & 15) << 2) | (c >> 6);
                    temp[i++] = c & 63;
                }
                if (isNaN(c2)) {
                    temp[i - 1] = 64;
                }
                else if (isNaN(c3)) {
                    temp[i - 2] = 64;
                    temp[i - 1] = 64;
                }
                for (var t = 0; t < i; t++) {
                    output.push(codex.charAt(temp[t]));
                }
            } while (index < length);
            return output.join("");
        },
        getSessionStorage: function OSF_Outil$getSessionStorage() {
            return _getSessionStorage();
        },
        getLocalStorage: function OSF_Outil$getLocalStorage() {
            if (!_safeLocalStorage) {
                try {
                    var localStorage = window.localStorage;
                }
                catch (ex) {
                    localStorage = null;
                }
                _safeLocalStorage = new OfficeExt.SafeStorage(localStorage);
            }
            return _safeLocalStorage;
        },
        convertIntToCssHexColor: function OSF_Outil$convertIntToCssHexColor(val) {
            var hex = "#" + (Number(val) + 0x1000000).toString(16).slice(-6);
            return hex;
        },
        attachClickHandler: function OSF_Outil$attachClickHandler(element, handler) {
            element.onclick = function (e) {
                handler();
            };
            element.ontouchend = function (e) {
                handler();
                e.preventDefault();
            };
        },
        getQueryStringParamValue: function OSF_Outil$getQueryStringParamValue(queryString, paramName) {
            var e = Function._validateParams(arguments, [{ name: "queryString", type: String, mayBeNull: false },
                { name: "paramName", type: String, mayBeNull: false }
            ]);
            if (e) {
                OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: Parameters cannot be null.");
                return "";
            }
            var queryExp = new RegExp("[\\?&]" + paramName + "=([^&#]*)", "i");
            if (!queryExp.test(queryString)) {
                OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: The parameter is not found.");
                return "";
            }
            return queryExp.exec(queryString)[1];
        },
        getHostnamePortionForLogging: function OSF_Outil$getHostnamePortionForLogging(hostname) {
            var e = Function._validateParams(arguments, [{ name: "hostname", type: String, mayBeNull: false }
            ]);
            if (e) {
                return "";
            }
            var hostnameSubstrings = hostname.split('.');
            var len = hostnameSubstrings.length;
            if (len >= 2) {
                return hostnameSubstrings[len - 2] + "." + hostnameSubstrings[len - 1];
            }
            else if (len == 1) {
                return hostnameSubstrings[0];
            }
        },
        isiOS: function OSF_Outil$isiOS() {
            return (window.navigator.userAgent.match(/(iPad|iPhone|iPod)/g) ? true : false);
        },
        isChrome: function OSF_Outil$isChrome() {
            return (window.navigator.userAgent.indexOf("Chrome") > 0) && !OSF.OUtil.isEdge();
        },
        isEdge: function OSF_Outil$isEdge() {
            return window.navigator.userAgent.indexOf("Edge") > 0;
        },
        isIE: function OSF_Outil$isIE() {
            return window.navigator.userAgent.indexOf("Trident") > 0;
        },
        isFirefox: function OSF_Outil$isFirefox() {
            return window.navigator.userAgent.indexOf("Firefox") > 0;
        },
        startsWith: function OSF_Outil$startsWith(originalString, patternToCheck, browserIsIE) {
            if (browserIsIE) {
                return originalString.substr(0, patternToCheck.length) === patternToCheck;
            }
            else {
                return originalString.startsWith(patternToCheck);
            }
        },
        containsPort: function OSF_Outil$containsPort(url, protocol, hostname, portNumber) {
            return this.startsWith(url, protocol + "//" + hostname + ":" + portNumber, true) || this.startsWith(url, hostname + ":" + portNumber, true);
        },
        getRedundandPortString: function OSF_Outil$getRedundandPortString(url, parser) {
            if (!url || !parser)
                return "";
            if (parser.protocol == "https:" && this.containsPort(url, "https:", parser.hostname, "443"))
                return ":443";
            else if (parser.protocol == "http:" && this.containsPort(url, "http:", parser.hostname, "80"))
                return ":80";
            return "";
        },
        removeChar: function OSF_Outil$removeChar(url, indexOfCharToRemove) {
            if (indexOfCharToRemove < url.length - 1)
                return url.substring(0, indexOfCharToRemove) + url.substring(indexOfCharToRemove + 1);
            else if (indexOfCharToRemove == url.length - 1)
                return url.substring(0, url.length - 1);
            else
                return url;
        },
        cleanUrlOfChar: function OSF_Outil$cleanUrlOfChar(url, charToClean) {
            var i;
            for (i = 0; i < url.length; i++) {
                if (url.charAt(i) === charToClean) {
                    if (i + 1 >= url.length) {
                        return this.removeChar(url, i);
                    }
                    else if (charToClean === '/') {
                        if (url.charAt(i + 1) === '?' || url.charAt(i + 1) === '#') {
                            return this.removeChar(url, i);
                        }
                    }
                    else if (charToClean === '?') {
                        if (url.charAt(i + 1) === '#') {
                            return this.removeChar(url, i);
                        }
                    }
                }
            }
            return url;
        },
        cleanUrl: function OSF_Outil$cleanUrl(url) {
            url = this.cleanUrlOfChar(url, '/');
            url = this.cleanUrlOfChar(url, '?');
            url = this.cleanUrlOfChar(url, '#');
            if (url.substr(0, 8) == "https://") {
                var portIndex = url.indexOf(":443");
                if (portIndex != -1) {
                    if (portIndex == url.length - 4 || url.charAt(portIndex + 4) == "/" || url.charAt(portIndex + 4) == "?" || url.charAt(portIndex + 4) == "#") {
                        url = url.substring(0, portIndex) + url.substring(portIndex + 4);
                    }
                }
            }
            else if (url.substr(0, 7) == "http://") {
                var portIndex = url.indexOf(":80");
                if (portIndex != -1) {
                    if (portIndex == url.length - 3 || url.charAt(portIndex + 3) == "/" || url.charAt(portIndex + 3) == "?" || url.charAt(portIndex + 3) == "#") {
                        url = url.substring(0, portIndex) + url.substring(portIndex + 3);
                    }
                }
            }
            return url;
        },
        parseUrl: function OSF_Outil$parseUrl(url, enforceHttps) {
            if (enforceHttps === void 0) { enforceHttps = false; }
            if (typeof url === "undefined" || !url) {
                return undefined;
            }
            var notHttpsErrorMessage = "NotHttps";
            var invalidUrlErrorMessage = "InvalidUrl";
            var isIEBoolean = this.isIE();
            var parsedUrlObj = {
                protocol: undefined,
                hostname: undefined,
                host: undefined,
                port: undefined,
                pathname: undefined,
                search: undefined,
                hash: undefined,
                isPortPartOfUrl: undefined
            };
            try {
                if (isIEBoolean) {
                    var parser = document.createElement("a");
                    parser.href = url;
                    if (!parser || !parser.protocol || !parser.host || !parser.hostname || !parser.href
                        || this.cleanUrl(parser.href).toLowerCase() !== this.cleanUrl(url).toLowerCase()) {
                        throw invalidUrlErrorMessage;
                    }
                    if (OSF.OUtil.checkFlight(OSF.FlightNames.AddinEnforceHttps)) {
                        if (enforceHttps && parser.protocol != "https:")
                            throw new Error(notHttpsErrorMessage);
                    }
                    var redundandPortString = this.getRedundandPortString(url, parser);
                    parsedUrlObj.protocol = parser.protocol;
                    parsedUrlObj.hostname = parser.hostname;
                    parsedUrlObj.port = (redundandPortString == "") ? parser.port : "";
                    parsedUrlObj.host = (redundandPortString != "") ? parser.hostname : parser.host;
                    parsedUrlObj.pathname = (isIEBoolean ? "/" : "") + parser.pathname;
                    parsedUrlObj.search = parser.search;
                    parsedUrlObj.hash = parser.hash;
                    parsedUrlObj.isPortPartOfUrl = this.containsPort(url, parser.protocol, parser.hostname, parser.port);
                }
                else {
                    var urlObj = new URL(url);
                    if (urlObj && urlObj.protocol && urlObj.host && urlObj.hostname) {
                        if (OSF.OUtil.checkFlight(OSF.FlightNames.AddinEnforceHttps)) {
                            if (enforceHttps && urlObj.protocol != "https:")
                                throw new Error(notHttpsErrorMessage);
                        }
                        parsedUrlObj.protocol = urlObj.protocol;
                        parsedUrlObj.hostname = urlObj.hostname;
                        parsedUrlObj.port = urlObj.port;
                        parsedUrlObj.host = urlObj.host;
                        parsedUrlObj.pathname = urlObj.pathname;
                        parsedUrlObj.search = urlObj.search;
                        parsedUrlObj.hash = urlObj.hash;
                        parsedUrlObj.isPortPartOfUrl = urlObj.host.lastIndexOf(":" + urlObj.port) == (urlObj.host.length - urlObj.port.length - 1);
                    }
                }
            }
            catch (err) {
                if (err.message === notHttpsErrorMessage)
                    throw err;
            }
            return parsedUrlObj;
        },
        shallowCopy: function OSF_Outil$shallowCopy(sourceObj) {
            if (sourceObj == null) {
                return null;
            }
            else if (!(sourceObj instanceof Object)) {
                return sourceObj;
            }
            else if (Array.isArray(sourceObj)) {
                var copyArr = [];
                for (var i = 0; i < sourceObj.length; i++) {
                    copyArr.push(sourceObj[i]);
                }
                return copyArr;
            }
            else {
                var copyObj = sourceObj.constructor();
                for (var property in sourceObj) {
                    if (sourceObj.hasOwnProperty(property)) {
                        copyObj[property] = sourceObj[property];
                    }
                }
                return copyObj;
            }
        },
        createObject: function OSF_Outil$createObject(properties) {
            var obj = null;
            if (properties) {
                obj = {};
                var len = properties.length;
                for (var i = 0; i < len; i++) {
                    obj[properties[i].name] = properties[i].value;
                }
            }
            return obj;
        },
        addClass: function OSF_OUtil$addClass(elmt, val) {
            if (!OSF.OUtil.hasClass(elmt, val)) {
                var className = elmt.getAttribute(_classN);
                if (className) {
                    elmt.setAttribute(_classN, className + " " + val);
                }
                else {
                    elmt.setAttribute(_classN, val);
                }
            }
        },
        removeClass: function OSF_OUtil$removeClass(elmt, val) {
            if (OSF.OUtil.hasClass(elmt, val)) {
                var className = elmt.getAttribute(_classN);
                var reg = new RegExp('(\\s|^)' + val + '(\\s|$)');
                className = className.replace(reg, '');
                elmt.setAttribute(_classN, className);
            }
        },
        hasClass: function OSF_OUtil$hasClass(elmt, clsName) {
            var className = elmt.getAttribute(_classN);
            return className && className.match(new RegExp('(\\s|^)' + clsName + '(\\s|$)'));
        },
        focusToFirstTabbable: function OSF_OUtil$focusToFirstTabbable(all, backward) {
            var next;
            var focused = false;
            var candidate;
            var setFlag = function (e) {
                focused = true;
            };
            var findNextPos = function (allLen, currPos, backward) {
                if (currPos < 0 || currPos > allLen) {
                    return -1;
                }
                else if (currPos === 0 && backward) {
                    return -1;
                }
                else if (currPos === allLen - 1 && !backward) {
                    return -1;
                }
                if (backward) {
                    return currPos - 1;
                }
                else {
                    return currPos + 1;
                }
            };
            all = _reOrderTabbableElements(all);
            next = backward ? all.length - 1 : 0;
            if (all.length === 0) {
                return null;
            }
            while (!focused && next >= 0 && next < all.length) {
                candidate = all[next];
                window.focus();
                candidate.addEventListener('focus', setFlag);
                candidate.focus();
                candidate.removeEventListener('focus', setFlag);
                next = findNextPos(all.length, next, backward);
                if (!focused && candidate === document.activeElement) {
                    focused = true;
                }
            }
            if (focused) {
                return candidate;
            }
            else {
                return null;
            }
        },
        focusToNextTabbable: function OSF_OUtil$focusToNextTabbable(all, curr, shift) {
            var currPos;
            var next;
            var focused = false;
            var candidate;
            var setFlag = function (e) {
                focused = true;
            };
            var findCurrPos = function (all, curr) {
                var i = 0;
                for (; i < all.length; i++) {
                    if (all[i] === curr) {
                        return i;
                    }
                }
                return -1;
            };
            var findNextPos = function (allLen, currPos, shift) {
                if (currPos < 0 || currPos > allLen) {
                    return -1;
                }
                else if (currPos === 0 && shift) {
                    return -1;
                }
                else if (currPos === allLen - 1 && !shift) {
                    return -1;
                }
                if (shift) {
                    return currPos - 1;
                }
                else {
                    return currPos + 1;
                }
            };
            all = _reOrderTabbableElements(all);
            currPos = findCurrPos(all, curr);
            next = findNextPos(all.length, currPos, shift);
            if (next < 0) {
                return null;
            }
            while (!focused && next >= 0 && next < all.length) {
                candidate = all[next];
                candidate.addEventListener('focus', setFlag);
                candidate.focus();
                candidate.removeEventListener('focus', setFlag);
                next = findNextPos(all.length, next, shift);
                if (!focused && candidate === document.activeElement) {
                    focused = true;
                }
            }
            if (focused) {
                return candidate;
            }
            else {
                return null;
            }
        },
        isNullOrUndefined: function OSF_OUtil$isNullOrUndefined(value) {
            if (typeof (value) === "undefined") {
                return true;
            }
            if (value === null) {
                return true;
            }
            return false;
        },
        stringEndsWith: function OSF_OUtil$stringEndsWith(value, subString) {
            if (!OSF.OUtil.isNullOrUndefined(value) && !OSF.OUtil.isNullOrUndefined(subString)) {
                if (subString.length > value.length) {
                    return false;
                }
                if (value.substr(value.length - subString.length) === subString) {
                    return true;
                }
            }
            return false;
        },
        hashCode: function OSF_OUtil$hashCode(str) {
            var hash = 0;
            if (!OSF.OUtil.isNullOrUndefined(str)) {
                var i = 0;
                var len = str.length;
                while (i < len) {
                    hash = (hash << 5) - hash + str.charCodeAt(i++) | 0;
                }
            }
            return hash;
        },
        getValue: function OSF_OUtil$getValue(value, defaultValue) {
            if (OSF.OUtil.isNullOrUndefined(value)) {
                return defaultValue;
            }
            return value;
        },
        externalNativeFunctionExists: function OSF_OUtil$externalNativeFunctionExists(type) {
            return type === 'unknown' || type !== 'undefined';
        }
    };
})();
OSF.OUtil.Guid = (function () {
    var hexCode = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f"];
    return {
        generateNewGuid: function OSF_Outil_Guid$generateNewGuid() {
            var result = "";
            var tick = (new Date()).getTime();
            var index = 0;
            for (; index < 32 && tick > 0; index++) {
                if (index == 8 || index == 12 || index == 16 || index == 20) {
                    result += "-";
                }
                result += hexCode[tick % 16];
                tick = Math.floor(tick / 16);
            }
            for (; index < 32; index++) {
                if (index == 8 || index == 12 || index == 16 || index == 20) {
                    result += "-";
                }
                result += hexCode[Math.floor(Math.random() * 16)];
            }
            return result;
        }
    };
})();
try {
    (function () {
        OSF.Flights = OSF.OUtil.parseFlights(true);
    })();
}
catch (ex) { }
window.OSF = OSF;
OSF.OUtil.setNamespace("OSF", window);
OSF.MessageIDs = {
    "FetchBundleUrl": 0,
    "LoadReactBundle": 1,
    "LoadBundleSuccess": 2,
    "LoadBundleError": 3
};
OSF.AppName = {
    Unsupported: 0,
    Excel: 1,
    Word: 2,
    PowerPoint: 4,
    Outlook: 8,
    ExcelWebApp: 16,
    WordWebApp: 32,
    OutlookWebApp: 64,
    Project: 128,
    AccessWebApp: 256,
    PowerpointWebApp: 512,
    ExcelIOS: 1024,
    Sway: 2048,
    WordIOS: 4096,
    PowerPointIOS: 8192,
    Access: 16384,
    Lync: 32768,
    OutlookIOS: 65536,
    OneNoteWebApp: 131072,
    OneNote: 262144,
    ExcelWinRT: 524288,
    WordWinRT: 1048576,
    PowerpointWinRT: 2097152,
    OutlookAndroid: 4194304,
    OneNoteWinRT: 8388608,
    ExcelAndroid: 8388609,
    VisioWebApp: 8388610,
    OneNoteIOS: 8388611,
    WordAndroid: 8388613,
    PowerpointAndroid: 8388614,
    Visio: 8388615,
    OneNoteAndroid: 4194305
};
OSF.InternalPerfMarker = {
    DataCoercionBegin: "Agave.HostCall.CoerceDataStart",
    DataCoercionEnd: "Agave.HostCall.CoerceDataEnd"
};
OSF.HostCallPerfMarker = {
    IssueCall: "Agave.HostCall.IssueCall",
    ReceiveResponse: "Agave.HostCall.ReceiveResponse",
    RuntimeExceptionRaised: "Agave.HostCall.RuntimeExecptionRaised"
};
OSF.AgaveHostAction = {
    "Select": 0,
    "UnSelect": 1,
    "CancelDialog": 2,
    "InsertAgave": 3,
    "CtrlF6In": 4,
    "CtrlF6Exit": 5,
    "CtrlF6ExitShift": 6,
    "SelectWithError": 7,
    "NotifyHostError": 8,
    "RefreshAddinCommands": 9,
    "PageIsReady": 10,
    "TabIn": 11,
    "TabInShift": 12,
    "TabExit": 13,
    "TabExitShift": 14,
    "EscExit": 15,
    "F2Exit": 16,
    "ExitNoFocusable": 17,
    "ExitNoFocusableShift": 18,
    "MouseEnter": 19,
    "MouseLeave": 20,
    "UpdateTargetUrl": 21,
    "InstallCustomFunctions": 22,
    "SendTelemetryEvent": 23,
    "UninstallCustomFunctions": 24,
    "SendMessage": 25,
    "LaunchExtensionComponent": 26,
    "StopExtensionComponent": 27,
    "RestartExtensionComponent": 28,
    "EnableTaskPaneHeaderButton": 29,
    "DisableTaskPaneHeaderButton": 30,
    "TaskPaneHeaderButtonClicked": 31,
    "RemoveAppCommandsAddin": 32,
    "RefreshRibbonGallery": 33,
    "GetOriginalControlId": 34,
    "OfficeJsReady": 35,
    "InsertDevManifest": 36,
    "InsertDevManifestError": 37,
    "SendCustomerContent": 38,
    "KeyboardShortcuts": 39
};
OSF.SharedConstants = {
    "NotificationConversationIdSuffix": '_ntf'
};
OSF.DialogMessageType = {
    DialogMessageReceived: 0,
    DialogParentMessageReceived: 1,
    DialogClosed: 12006
};
OSF.OfficeAppContext = function OSF_OfficeAppContext(id, appName, appVersion, appUILocale, dataLocale, docUrl, clientMode, settings, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, appMinorVersion, requirementMatrix, hostCustomMessage, hostFullVersion, clientWindowHeight, clientWindowWidth, addinName, appDomains, dialogRequirementMatrix, featureGates, officeTheme, initialDisplayMode) {
    this._id = id;
    this._appName = appName;
    this._appVersion = appVersion;
    this._appUILocale = appUILocale;
    this._dataLocale = dataLocale;
    this._docUrl = docUrl;
    this._clientMode = clientMode;
    this._settings = settings;
    this._reason = reason;
    this._osfControlType = osfControlType;
    this._eToken = eToken;
    this._correlationId = correlationId;
    this._appInstanceId = appInstanceId;
    this._touchEnabled = touchEnabled;
    this._commerceAllowed = commerceAllowed;
    this._appMinorVersion = appMinorVersion;
    this._requirementMatrix = requirementMatrix;
    this._hostCustomMessage = hostCustomMessage;
    this._hostFullVersion = hostFullVersion;
    this._isDialog = false;
    this._clientWindowHeight = clientWindowHeight;
    this._clientWindowWidth = clientWindowWidth;
    this._addinName = addinName;
    this._appDomains = appDomains;
    this._dialogRequirementMatrix = dialogRequirementMatrix;
    this._featureGates = featureGates;
    this._officeTheme = officeTheme;
    this._initialDisplayMode = initialDisplayMode;
    this.get_id = function get_id() { return this._id; };
    this.get_appName = function get_appName() { return this._appName; };
    this.get_appVersion = function get_appVersion() { return this._appVersion; };
    this.get_appUILocale = function get_appUILocale() { return this._appUILocale; };
    this.get_dataLocale = function get_dataLocale() { return this._dataLocale; };
    this.get_docUrl = function get_docUrl() { return this._docUrl; };
    this.get_clientMode = function get_clientMode() { return this._clientMode; };
    this.get_bindings = function get_bindings() { return this._bindings; };
    this.get_settings = function get_settings() { return this._settings; };
    this.get_reason = function get_reason() { return this._reason; };
    this.get_osfControlType = function get_osfControlType() { return this._osfControlType; };
    this.get_eToken = function get_eToken() { return this._eToken; };
    this.get_correlationId = function get_correlationId() { return this._correlationId; };
    this.get_appInstanceId = function get_appInstanceId() { return this._appInstanceId; };
    this.get_touchEnabled = function get_touchEnabled() { return this._touchEnabled; };
    this.get_commerceAllowed = function get_commerceAllowed() { return this._commerceAllowed; };
    this.get_appMinorVersion = function get_appMinorVersion() { return this._appMinorVersion; };
    this.get_requirementMatrix = function get_requirementMatrix() { return this._requirementMatrix; };
    this.get_dialogRequirementMatrix = function get_dialogRequirementMatrix() { return this._dialogRequirementMatrix; };
    this.get_hostCustomMessage = function get_hostCustomMessage() { return this._hostCustomMessage; };
    this.get_hostFullVersion = function get_hostFullVersion() { return this._hostFullVersion; };
    this.get_isDialog = function get_isDialog() { return this._isDialog; };
    this.get_clientWindowHeight = function get_clientWindowHeight() { return this._clientWindowHeight; };
    this.get_clientWindowWidth = function get_clientWindowWidth() { return this._clientWindowWidth; };
    this.get_addinName = function get_addinName() { return this._addinName; };
    this.get_appDomains = function get_appDomains() { return this._appDomains; };
    this.get_featureGates = function get_featureGates() { return this._featureGates; };
    this.get_officeTheme = function get_officeTheme() { return this._officeTheme; };
    this.get_initialDisplayMode = function get_initialDisplayMode() { return this._initialDisplayMode ? this._initialDisplayMode : 0; };
};
OSF.OsfControlType = {
    DocumentLevel: 0,
    ContainerLevel: 1
};
OSF.ClientMode = {
    ReadOnly: 0,
    ReadWrite: 1
};
OSF.OUtil.setNamespace("Microsoft", window);
OSF.OUtil.setNamespace("Office", Microsoft);
OSF.OUtil.setNamespace("Client", Microsoft.Office);
OSF.OUtil.setNamespace("WebExtension", Microsoft.Office);
Microsoft.Office.WebExtension.InitializationReason = {
    Inserted: "inserted",
    DocumentOpened: "documentOpened",
    ControlActivation: "controlActivation"
};
Microsoft.Office.WebExtension.ValueFormat = {
    Unformatted: "unformatted",
    Formatted: "formatted"
};
Microsoft.Office.WebExtension.FilterType = {
    All: "all"
};
Microsoft.Office.WebExtension.Parameters = {
    BindingType: "bindingType",
    CoercionType: "coercionType",
    ValueFormat: "valueFormat",
    FilterType: "filterType",
    Columns: "columns",
    SampleData: "sampleData",
    GoToType: "goToType",
    SelectionMode: "selectionMode",
    Id: "id",
    PromptText: "promptText",
    ItemName: "itemName",
    FailOnCollision: "failOnCollision",
    StartRow: "startRow",
    StartColumn: "startColumn",
    RowCount: "rowCount",
    ColumnCount: "columnCount",
    Callback: "callback",
    AsyncContext: "asyncContext",
    Data: "data",
    Rows: "rows",
    OverwriteIfStale: "overwriteIfStale",
    FileType: "fileType",
    EventType: "eventType",
    Handler: "handler",
    SliceSize: "sliceSize",
    SliceIndex: "sliceIndex",
    ActiveView: "activeView",
    Status: "status",
    PlatformType: "platformType",
    HostType: "hostType",
    ForceConsent: "forceConsent",
    ForceAddAccount: "forceAddAccount",
    AuthChallenge: "authChallenge",
    AllowConsentPrompt: "allowConsentPrompt",
    ForMSGraphAccess: "forMSGraphAccess",
    AllowSignInPrompt: "allowSignInPrompt",
    JsonPayload: "jsonPayload",
    EnableNewHosts: "enableNewHosts",
    AccountTypeFilter: "accountTypeFilter",
    AddinTrustId: "addinTrustId",
    Reserved: "reserved",
    Tcid: "tcid",
    Xml: "xml",
    Namespace: "namespace",
    Prefix: "prefix",
    XPath: "xPath",
    Text: "text",
    ImageLeft: "imageLeft",
    ImageTop: "imageTop",
    ImageWidth: "imageWidth",
    ImageHeight: "imageHeight",
    TaskId: "taskId",
    FieldId: "fieldId",
    FieldValue: "fieldValue",
    ServerUrl: "serverUrl",
    ListName: "listName",
    ResourceId: "resourceId",
    ViewType: "viewType",
    ViewName: "viewName",
    GetRawValue: "getRawValue",
    CellFormat: "cellFormat",
    TableOptions: "tableOptions",
    TaskIndex: "taskIndex",
    ResourceIndex: "resourceIndex",
    CustomFieldId: "customFieldId",
    Url: "url",
    MessageHandler: "messageHandler",
    Width: "width",
    Height: "height",
    RequireHTTPs: "requireHTTPS",
    MessageToParent: "messageToParent",
    DisplayInIframe: "displayInIframe",
    MessageContent: "messageContent",
    HideTitle: "hideTitle",
    UseDeviceIndependentPixels: "useDeviceIndependentPixels",
    PromptBeforeOpen: "promptBeforeOpen",
    EnforceAppDomain: "enforceAppDomain",
    UrlNoHostInfo: "urlNoHostInfo",
    TargetOrigin: "targetOrigin",
    AppCommandInvocationCompletedData: "appCommandInvocationCompletedData",
    Base64: "base64",
    FormId: "formId"
};
OSF.OUtil.setNamespace("DDA", OSF);
OSF.DDA.DocumentMode = {
    ReadOnly: 1,
    ReadWrite: 0
};
OSF.DDA.PropertyDescriptors = {
    AsyncResultStatus: "AsyncResultStatus"
};
OSF.DDA.EventDescriptors = {};
OSF.DDA.ListDescriptors = {};
OSF.DDA.UI = {};
OSF.DDA.getXdmEventName = function OSF_DDA$GetXdmEventName(id, eventType) {
    if (eventType == Microsoft.Office.WebExtension.EventType.BindingSelectionChanged ||
        eventType == Microsoft.Office.WebExtension.EventType.BindingDataChanged ||
        eventType == Microsoft.Office.WebExtension.EventType.DataNodeDeleted ||
        eventType == Microsoft.Office.WebExtension.EventType.DataNodeInserted ||
        eventType == Microsoft.Office.WebExtension.EventType.DataNodeReplaced) {
        return id + "_" + eventType;
    }
    else {
        return eventType;
    }
};
OSF.DDA.MethodDispId = {
    dispidMethodMin: 64,
    dispidGetSelectedDataMethod: 64,
    dispidSetSelectedDataMethod: 65,
    dispidAddBindingFromSelectionMethod: 66,
    dispidAddBindingFromPromptMethod: 67,
    dispidGetBindingMethod: 68,
    dispidReleaseBindingMethod: 69,
    dispidGetBindingDataMethod: 70,
    dispidSetBindingDataMethod: 71,
    dispidAddRowsMethod: 72,
    dispidClearAllRowsMethod: 73,
    dispidGetAllBindingsMethod: 74,
    dispidLoadSettingsMethod: 75,
    dispidSaveSettingsMethod: 76,
    dispidGetDocumentCopyMethod: 77,
    dispidAddBindingFromNamedItemMethod: 78,
    dispidAddColumnsMethod: 79,
    dispidGetDocumentCopyChunkMethod: 80,
    dispidReleaseDocumentCopyMethod: 81,
    dispidNavigateToMethod: 82,
    dispidGetActiveViewMethod: 83,
    dispidGetDocumentThemeMethod: 84,
    dispidGetOfficeThemeMethod: 85,
    dispidGetFilePropertiesMethod: 86,
    dispidClearFormatsMethod: 87,
    dispidSetTableOptionsMethod: 88,
    dispidSetFormatsMethod: 89,
    dispidExecuteRichApiRequestMethod: 93,
    dispidAppCommandInvocationCompletedMethod: 94,
    dispidCloseContainerMethod: 97,
    dispidGetAccessTokenMethod: 98,
    dispidGetAuthContextMethod: 99,
    dispidOpenBrowserWindow: 102,
    dispidCreateDocumentMethod: 105,
    dispidInsertFormMethod: 106,
    dispidDisplayRibbonCalloutAsyncMethod: 109,
    dispidGetSelectedTaskMethod: 110,
    dispidGetSelectedResourceMethod: 111,
    dispidGetTaskMethod: 112,
    dispidGetResourceFieldMethod: 113,
    dispidGetWSSUrlMethod: 114,
    dispidGetTaskFieldMethod: 115,
    dispidGetProjectFieldMethod: 116,
    dispidGetSelectedViewMethod: 117,
    dispidGetTaskByIndexMethod: 118,
    dispidGetResourceByIndexMethod: 119,
    dispidSetTaskFieldMethod: 120,
    dispidSetResourceFieldMethod: 121,
    dispidGetMaxTaskIndexMethod: 122,
    dispidGetMaxResourceIndexMethod: 123,
    dispidCreateTaskMethod: 124,
    dispidAddDataPartMethod: 128,
    dispidGetDataPartByIdMethod: 129,
    dispidGetDataPartsByNamespaceMethod: 130,
    dispidGetDataPartXmlMethod: 131,
    dispidGetDataPartNodesMethod: 132,
    dispidDeleteDataPartMethod: 133,
    dispidGetDataNodeValueMethod: 134,
    dispidGetDataNodeXmlMethod: 135,
    dispidGetDataNodesMethod: 136,
    dispidSetDataNodeValueMethod: 137,
    dispidSetDataNodeXmlMethod: 138,
    dispidAddDataNamespaceMethod: 139,
    dispidGetDataUriByPrefixMethod: 140,
    dispidGetDataPrefixByUriMethod: 141,
    dispidGetDataNodeTextMethod: 142,
    dispidSetDataNodeTextMethod: 143,
    dispidMessageParentMethod: 144,
    dispidSendMessageMethod: 145,
    dispidExecuteFeature: 146,
    dispidQueryFeature: 147,
    dispidMethodMax: 147
};
OSF.DDA.EventDispId = {
    dispidEventMin: 0,
    dispidInitializeEvent: 0,
    dispidSettingsChangedEvent: 1,
    dispidDocumentSelectionChangedEvent: 2,
    dispidBindingSelectionChangedEvent: 3,
    dispidBindingDataChangedEvent: 4,
    dispidDocumentOpenEvent: 5,
    dispidDocumentCloseEvent: 6,
    dispidActiveViewChangedEvent: 7,
    dispidDocumentThemeChangedEvent: 8,
    dispidOfficeThemeChangedEvent: 9,
    dispidDialogMessageReceivedEvent: 10,
    dispidDialogNotificationShownInAddinEvent: 11,
    dispidDialogParentMessageReceivedEvent: 12,
    dispidObjectDeletedEvent: 13,
    dispidObjectSelectionChangedEvent: 14,
    dispidObjectDataChangedEvent: 15,
    dispidContentControlAddedEvent: 16,
    dispidActivationStatusChangedEvent: 32,
    dispidRichApiMessageEvent: 33,
    dispidAppCommandInvokedEvent: 39,
    dispidOlkItemSelectedChangedEvent: 46,
    dispidOlkRecipientsChangedEvent: 47,
    dispidOlkAppointmentTimeChangedEvent: 48,
    dispidOlkRecurrenceChangedEvent: 49,
    dispidOlkAttachmentsChangedEvent: 50,
    dispidOlkEnhancedLocationsChangedEvent: 51,
    dispidOlkInfobarClickedEvent: 52,
    dispidTaskSelectionChangedEvent: 56,
    dispidResourceSelectionChangedEvent: 57,
    dispidViewSelectionChangedEvent: 58,
    dispidDataNodeAddedEvent: 60,
    dispidDataNodeReplacedEvent: 61,
    dispidDataNodeDeletedEvent: 62,
    dispidEventMax: 63
};
OSF.DDA.ErrorCodeManager = (function () {
    var _errorMappings = {};
    return {
        getErrorArgs: function OSF_DDA_ErrorCodeManager$getErrorArgs(errorCode) {
            var errorArgs = _errorMappings[errorCode];
            if (!errorArgs) {
                errorArgs = _errorMappings[this.errorCodes.ooeInternalError];
            }
            else {
                if (!errorArgs.name) {
                    errorArgs.name = _errorMappings[this.errorCodes.ooeInternalError].name;
                }
                if (!errorArgs.message) {
                    errorArgs.message = _errorMappings[this.errorCodes.ooeInternalError].message;
                }
            }
            return errorArgs;
        },
        addErrorMessage: function OSF_DDA_ErrorCodeManager$addErrorMessage(errorCode, errorNameMessage) {
            _errorMappings[errorCode] = errorNameMessage;
        },
        errorCodes: {
            ooeSuccess: 0,
            ooeChunkResult: 1,
            ooeCoercionTypeNotSupported: 1000,
            ooeGetSelectionNotMatchDataType: 1001,
            ooeCoercionTypeNotMatchBinding: 1002,
            ooeInvalidGetRowColumnCounts: 1003,
            ooeSelectionNotSupportCoercionType: 1004,
            ooeInvalidGetStartRowColumn: 1005,
            ooeNonUniformPartialGetNotSupported: 1006,
            ooeGetDataIsTooLarge: 1008,
            ooeFileTypeNotSupported: 1009,
            ooeGetDataParametersConflict: 1010,
            ooeInvalidGetColumns: 1011,
            ooeInvalidGetRows: 1012,
            ooeInvalidReadForBlankRow: 1013,
            ooeUnsupportedDataObject: 2000,
            ooeCannotWriteToSelection: 2001,
            ooeDataNotMatchSelection: 2002,
            ooeOverwriteWorksheetData: 2003,
            ooeDataNotMatchBindingSize: 2004,
            ooeInvalidSetStartRowColumn: 2005,
            ooeInvalidDataFormat: 2006,
            ooeDataNotMatchCoercionType: 2007,
            ooeDataNotMatchBindingType: 2008,
            ooeSetDataIsTooLarge: 2009,
            ooeNonUniformPartialSetNotSupported: 2010,
            ooeInvalidSetColumns: 2011,
            ooeInvalidSetRows: 2012,
            ooeSetDataParametersConflict: 2013,
            ooeCellDataAmountBeyondLimits: 2014,
            ooeSelectionCannotBound: 3000,
            ooeBindingNotExist: 3002,
            ooeBindingToMultipleSelection: 3003,
            ooeInvalidSelectionForBindingType: 3004,
            ooeOperationNotSupportedOnThisBindingType: 3005,
            ooeNamedItemNotFound: 3006,
            ooeMultipleNamedItemFound: 3007,
            ooeInvalidNamedItemForBindingType: 3008,
            ooeUnknownBindingType: 3009,
            ooeOperationNotSupportedOnMatrixData: 3010,
            ooeInvalidColumnsForBinding: 3011,
            ooeSettingNameNotExist: 4000,
            ooeSettingsCannotSave: 4001,
            ooeSettingsAreStale: 4002,
            ooeOperationNotSupported: 5000,
            ooeInternalError: 5001,
            ooeDocumentReadOnly: 5002,
            ooeEventHandlerNotExist: 5003,
            ooeInvalidApiCallInContext: 5004,
            ooeShuttingDown: 5005,
            ooeUnsupportedEnumeration: 5007,
            ooeIndexOutOfRange: 5008,
            ooeBrowserAPINotSupported: 5009,
            ooeInvalidParam: 5010,
            ooeRequestTimeout: 5011,
            ooeInvalidOrTimedOutSession: 5012,
            ooeInvalidApiArguments: 5013,
            ooeOperationCancelled: 5014,
            ooeWorkbookHidden: 5015,
            ooeWriteNotSupportedWhenModalDialogOpen: 5016,
            ooeTooManyIncompleteRequests: 5100,
            ooeRequestTokenUnavailable: 5101,
            ooeActivityLimitReached: 5102,
            ooeRequestPayloadSizeLimitExceeded: 5103,
            ooeResponsePayloadSizeLimitExceeded: 5104,
            ooeCustomXmlNodeNotFound: 6000,
            ooeCustomXmlError: 6100,
            ooeCustomXmlExceedQuota: 6101,
            ooeCustomXmlOutOfDate: 6102,
            ooeNoCapability: 7000,
            ooeCannotNavTo: 7001,
            ooeSpecifiedIdNotExist: 7002,
            ooeNavOutOfBound: 7004,
            ooeElementMissing: 8000,
            ooeProtectedError: 8001,
            ooeInvalidCellsValue: 8010,
            ooeInvalidTableOptionValue: 8011,
            ooeInvalidFormatValue: 8012,
            ooeRowIndexOutOfRange: 8020,
            ooeColIndexOutOfRange: 8021,
            ooeFormatValueOutOfRange: 8022,
            ooeCellFormatAmountBeyondLimits: 8023,
            ooeMemoryFileLimit: 11000,
            ooeNetworkProblemRetrieveFile: 11001,
            ooeInvalidSliceSize: 11002,
            ooeInvalidCallback: 11101,
            ooeInvalidWidth: 12000,
            ooeInvalidHeight: 12001,
            ooeNavigationError: 12002,
            ooeInvalidScheme: 12003,
            ooeAppDomains: 12004,
            ooeRequireHTTPS: 12005,
            ooeWebDialogClosed: 12006,
            ooeDialogAlreadyOpened: 12007,
            ooeEndUserAllow: 12008,
            ooeEndUserIgnore: 12009,
            ooeNotUILessDialog: 12010,
            ooeCrossZone: 12011,
            ooeNotSSOAgave: 13000,
            ooeSSOUserNotSignedIn: 13001,
            ooeSSOUserAborted: 13002,
            ooeSSOUnsupportedUserIdentity: 13003,
            ooeSSOInvalidResourceUrl: 13004,
            ooeSSOInvalidGrant: 13005,
            ooeSSOClientError: 13006,
            ooeSSOServerError: 13007,
            ooeAddinIsAlreadyRequestingToken: 13008,
            ooeSSOUserConsentNotSupportedByCurrentAddinCategory: 13009,
            ooeSSOConnectionLost: 13010,
            ooeResourceNotAllowed: 13011,
            ooeSSOUnsupportedPlatform: 13012,
            ooeSSOCallThrottled: 13013,
            ooeAccessDenied: 13990,
            ooeGeneralException: 13991
        },
        initializeErrorMessages: function OSF_DDA_ErrorCodeManager$initializeErrorMessages(stringNS) {
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotSupported] = { name: stringNS.L_InvalidCoercion, message: stringNS.L_CoercionTypeNotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetSelectionNotMatchDataType] = { name: stringNS.L_DataReadError, message: stringNS.L_GetSelectionNotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding] = { name: stringNS.L_InvalidCoercion, message: stringNS.L_CoercionTypeNotMatchBinding };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRowColumnCounts] = { name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetRowColumnCounts };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionNotSupportCoercionType] = { name: stringNS.L_DataReadError, message: stringNS.L_SelectionNotSupportCoercionType };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetStartRowColumn] = { name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetStartRowColumn };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialGetNotSupported] = { name: stringNS.L_DataReadError, message: stringNS.L_NonUniformPartialGetNotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataIsTooLarge] = { name: stringNS.L_DataReadError, message: stringNS.L_GetDataIsTooLarge };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeFileTypeNotSupported] = { name: stringNS.L_DataReadError, message: stringNS.L_FileTypeNotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataParametersConflict] = { name: stringNS.L_DataReadError, message: stringNS.L_GetDataParametersConflict };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetColumns] = { name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetColumns };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRows] = { name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetRows };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidReadForBlankRow] = { name: stringNS.L_DataReadError, message: stringNS.L_InvalidReadForBlankRow };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedDataObject] = { name: stringNS.L_DataWriteError, message: stringNS.L_UnsupportedDataObject };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotWriteToSelection] = { name: stringNS.L_DataWriteError, message: stringNS.L_CannotWriteToSelection };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchSelection] = { name: stringNS.L_DataWriteError, message: stringNS.L_DataNotMatchSelection };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOverwriteWorksheetData] = { name: stringNS.L_DataWriteError, message: stringNS.L_OverwriteWorksheetData };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingSize] = { name: stringNS.L_DataWriteError, message: stringNS.L_DataNotMatchBindingSize };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetStartRowColumn] = { name: stringNS.L_DataWriteError, message: stringNS.L_InvalidSetStartRowColumn };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidDataFormat] = { name: stringNS.L_InvalidFormat, message: stringNS.L_InvalidDataFormat };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchCoercionType] = { name: stringNS.L_InvalidDataObject, message: stringNS.L_DataNotMatchCoercionType };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingType] = { name: stringNS.L_InvalidDataObject, message: stringNS.L_DataNotMatchBindingType };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataIsTooLarge] = { name: stringNS.L_DataWriteError, message: stringNS.L_SetDataIsTooLarge };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialSetNotSupported] = { name: stringNS.L_DataWriteError, message: stringNS.L_NonUniformPartialSetNotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetColumns] = { name: stringNS.L_DataWriteError, message: stringNS.L_InvalidSetColumns };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetRows] = { name: stringNS.L_DataWriteError, message: stringNS.L_InvalidSetRows };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataParametersConflict] = { name: stringNS.L_DataWriteError, message: stringNS.L_SetDataParametersConflict };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionCannotBound] = { name: stringNS.L_BindingCreationError, message: stringNS.L_SelectionCannotBound };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingNotExist] = { name: stringNS.L_InvalidBindingError, message: stringNS.L_BindingNotExist };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingToMultipleSelection] = { name: stringNS.L_BindingCreationError, message: stringNS.L_BindingToMultipleSelection };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSelectionForBindingType] = { name: stringNS.L_BindingCreationError, message: stringNS.L_InvalidSelectionForBindingType };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnThisBindingType] = { name: stringNS.L_InvalidBindingOperation, message: stringNS.L_OperationNotSupportedOnThisBindingType };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNamedItemNotFound] = { name: stringNS.L_BindingCreationError, message: stringNS.L_NamedItemNotFound };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeMultipleNamedItemFound] = { name: stringNS.L_BindingCreationError, message: stringNS.L_MultipleNamedItemFound };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidNamedItemForBindingType] = { name: stringNS.L_BindingCreationError, message: stringNS.L_InvalidNamedItemForBindingType };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnknownBindingType] = { name: stringNS.L_InvalidBinding, message: stringNS.L_UnknownBindingType };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnMatrixData] = { name: stringNS.L_InvalidBindingOperation, message: stringNS.L_OperationNotSupportedOnMatrixData };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidColumnsForBinding] = { name: stringNS.L_InvalidBinding, message: stringNS.L_InvalidColumnsForBinding };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingNameNotExist] = { name: stringNS.L_ReadSettingsError, message: stringNS.L_SettingNameNotExist };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsCannotSave] = { name: stringNS.L_SaveSettingsError, message: stringNS.L_SettingsCannotSave };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsAreStale] = { name: stringNS.L_SettingsStaleError, message: stringNS.L_SettingsAreStale };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupported] = { name: stringNS.L_HostError, message: stringNS.L_OperationNotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError] = { name: stringNS.L_InternalError, message: stringNS.L_InternalErrorDescription };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDocumentReadOnly] = { name: stringNS.L_PermissionDenied, message: stringNS.L_DocumentReadOnly };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist] = { name: stringNS.L_EventRegistrationError, message: stringNS.L_EventHandlerNotExist };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext] = { name: stringNS.L_InvalidAPICall, message: stringNS.L_InvalidApiCallInContext };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeShuttingDown] = { name: stringNS.L_ShuttingDown, message: stringNS.L_ShuttingDown };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration] = { name: stringNS.L_UnsupportedEnumeration, message: stringNS.L_UnsupportedEnumerationMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeIndexOutOfRange] = { name: stringNS.L_IndexOutOfRange, message: stringNS.L_IndexOutOfRange };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBrowserAPINotSupported] = { name: stringNS.L_APINotSupported, message: stringNS.L_BrowserAPINotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTimeout] = { name: stringNS.L_APICallFailed, message: stringNS.L_RequestTimeout };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidOrTimedOutSession] = { name: stringNS.L_InvalidOrTimedOutSession, message: stringNS.L_InvalidOrTimedOutSessionMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiArguments] = { name: stringNS.L_APICallFailed, message: stringNS.L_InvalidApiArgumentsMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeWorkbookHidden] = { name: stringNS.L_APICallFailed, message: stringNS.L_WorkbookHiddenMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeWriteNotSupportedWhenModalDialogOpen] = { name: stringNS.L_APICallFailed, message: stringNS.L_WriteNotSupportedWhenModalDialogOpen };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeTooManyIncompleteRequests] = { name: stringNS.L_APICallFailed, message: stringNS.L_TooManyIncompleteRequests };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTokenUnavailable] = { name: stringNS.L_APICallFailed, message: stringNS.L_RequestTokenUnavailable };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeActivityLimitReached] = { name: stringNS.L_APICallFailed, message: stringNS.L_ActivityLimitReached };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestPayloadSizeLimitExceeded] = { name: stringNS.L_APICallFailed, message: stringNS.L_RequestPayloadSizeLimitExceededMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeResponsePayloadSizeLimitExceeded] = { name: stringNS.L_APICallFailed, message: stringNS.L_ResponsePayloadSizeLimitExceededMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlNodeNotFound] = { name: stringNS.L_InvalidNode, message: stringNS.L_CustomXmlNodeNotFound };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlError] = { name: stringNS.L_CustomXmlError, message: stringNS.L_CustomXmlError };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlExceedQuota] = { name: stringNS.L_CustomXmlExceedQuotaName, message: stringNS.L_CustomXmlExceedQuotaMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlOutOfDate] = { name: stringNS.L_CustomXmlOutOfDateName, message: stringNS.L_CustomXmlOutOfDateMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability] = { name: stringNS.L_PermissionDenied, message: stringNS.L_NoCapability };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotNavTo] = { name: stringNS.L_CannotNavigateTo, message: stringNS.L_CannotNavigateTo };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSpecifiedIdNotExist] = { name: stringNS.L_SpecifiedIdNotExist, message: stringNS.L_SpecifiedIdNotExist };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavOutOfBound] = { name: stringNS.L_NavOutOfBound, message: stringNS.L_NavOutOfBound };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellDataAmountBeyondLimits] = { name: stringNS.L_DataWriteReminder, message: stringNS.L_CellDataAmountBeyondLimits };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeElementMissing] = { name: stringNS.L_MissingParameter, message: stringNS.L_ElementMissing };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeProtectedError] = { name: stringNS.L_PermissionDenied, message: stringNS.L_NoCapability };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidCellsValue] = { name: stringNS.L_InvalidValue, message: stringNS.L_InvalidCellsValue };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidTableOptionValue] = { name: stringNS.L_InvalidValue, message: stringNS.L_InvalidTableOptionValue };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidFormatValue] = { name: stringNS.L_InvalidValue, message: stringNS.L_InvalidFormatValue };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRowIndexOutOfRange] = { name: stringNS.L_OutOfRange, message: stringNS.L_RowIndexOutOfRange };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeColIndexOutOfRange] = { name: stringNS.L_OutOfRange, message: stringNS.L_ColIndexOutOfRange };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeFormatValueOutOfRange] = { name: stringNS.L_OutOfRange, message: stringNS.L_FormatValueOutOfRange };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellFormatAmountBeyondLimits] = { name: stringNS.L_FormattingReminder, message: stringNS.L_CellFormatAmountBeyondLimits };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeMemoryFileLimit] = { name: stringNS.L_MemoryLimit, message: stringNS.L_CloseFileBeforeRetrieve };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNetworkProblemRetrieveFile] = { name: stringNS.L_NetworkProblem, message: stringNS.L_NetworkProblemRetrieveFile };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSliceSize] = { name: stringNS.L_InvalidValue, message: stringNS.L_SliceSizeNotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogAlreadyOpened] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_DialogAlreadyOpened };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidWidth] = { name: stringNS.L_IndexOutOfRange, message: stringNS.L_IndexOutOfRange };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidHeight] = { name: stringNS.L_IndexOutOfRange, message: stringNS.L_IndexOutOfRange };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavigationError] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_NetworkProblem };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidScheme] = { name: stringNS.L_DialogNavigateError, message: stringNS.L_DialogInvalidScheme };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeAppDomains] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_DialogAddressNotTrusted };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequireHTTPS] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_DialogRequireHTTPS };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeEndUserIgnore] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_UserClickIgnore };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCrossZone] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_NewWindowCrossZoneErrorString };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNotSSOAgave] = { name: stringNS.L_APINotSupported, message: stringNS.L_InvalidSSOAddinMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserNotSignedIn] = { name: stringNS.L_UserNotSignedIn, message: stringNS.L_UserNotSignedIn };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserAborted] = { name: stringNS.L_UserAborted, message: stringNS.L_UserAbortedMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUnsupportedUserIdentity] = { name: stringNS.L_UnsupportedUserIdentity, message: stringNS.L_UnsupportedUserIdentityMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOInvalidResourceUrl] = { name: stringNS.L_InvalidResourceUrl, message: stringNS.L_InvalidResourceUrlMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOInvalidGrant] = { name: stringNS.L_InvalidGrant, message: stringNS.L_InvalidGrantMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOClientError] = { name: stringNS.L_SSOClientError, message: stringNS.L_SSOClientErrorMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOServerError] = { name: stringNS.L_SSOServerError, message: stringNS.L_SSOServerErrorMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeAddinIsAlreadyRequestingToken] = { name: stringNS.L_AddinIsAlreadyRequestingToken, message: stringNS.L_AddinIsAlreadyRequestingTokenMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserConsentNotSupportedByCurrentAddinCategory] = { name: stringNS.L_SSOUserConsentNotSupportedByCurrentAddinCategory, message: stringNS.L_SSOUserConsentNotSupportedByCurrentAddinCategoryMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOConnectionLost] = { name: stringNS.L_SSOConnectionLostError, message: stringNS.L_SSOConnectionLostErrorMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUnsupportedPlatform] = { name: stringNS.L_APINotSupported, message: stringNS.L_SSOUnsupportedPlatform };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOCallThrottled] = { name: stringNS.L_APICallFailed, message: stringNS.L_RequestTokenUnavailable };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationCancelled] = { name: stringNS.L_OperationCancelledError, message: stringNS.L_OperationCancelledErrorMessage };
        }
    };
})();
(function (OfficeExt) {
    var Requirement;
    (function (Requirement) {
        var RequirementVersion = (function () {
            function RequirementVersion() {
            }
            return RequirementVersion;
        }());
        Requirement.RequirementVersion = RequirementVersion;
        var RequirementMatrix = (function () {
            function RequirementMatrix(_setMap) {
                this.isSetSupported = function _isSetSupported(name, minVersion) {
                    if (name == undefined) {
                        return false;
                    }
                    if (minVersion == undefined) {
                        minVersion = 0;
                    }
                    var setSupportArray = this._setMap;
                    var sets = setSupportArray._sets;
                    if (sets.hasOwnProperty(name.toLowerCase())) {
                        var setMaxVersion = sets[name.toLowerCase()];
                        try {
                            var setMaxVersionNum = this._getVersion(setMaxVersion);
                            minVersion = minVersion + "";
                            var minVersionNum = this._getVersion(minVersion);
                            if (setMaxVersionNum.major > 0 && setMaxVersionNum.major > minVersionNum.major) {
                                return true;
                            }
                            if (setMaxVersionNum.major > 0 &&
                                setMaxVersionNum.minor >= 0 &&
                                setMaxVersionNum.major == minVersionNum.major &&
                                setMaxVersionNum.minor >= minVersionNum.minor) {
                                return true;
                            }
                        }
                        catch (e) {
                            return false;
                        }
                    }
                    return false;
                };
                this._getVersion = function (version) {
                    version = version + "";
                    var temp = version.split(".");
                    var major = 0;
                    var minor = 0;
                    if (temp.length < 2 && isNaN(Number(version))) {
                        throw "version format incorrect";
                    }
                    else {
                        major = Number(temp[0]);
                        if (temp.length >= 2) {
                            minor = Number(temp[1]);
                        }
                        if (isNaN(major) || isNaN(minor)) {
                            throw "version format incorrect";
                        }
                    }
                    var result = { "minor": minor, "major": major };
                    return result;
                };
                this._setMap = _setMap;
                this.isSetSupported = this.isSetSupported.bind(this);
            }
            return RequirementMatrix;
        }());
        Requirement.RequirementMatrix = RequirementMatrix;
        var DefaultSetRequirement = (function () {
            function DefaultSetRequirement(setMap) {
                this._addSetMap = function DefaultSetRequirement_addSetMap(addedSet) {
                    for (var name in addedSet) {
                        this._sets[name] = addedSet[name];
                    }
                };
                this._sets = setMap;
            }
            return DefaultSetRequirement;
        }());
        Requirement.DefaultSetRequirement = DefaultSetRequirement;
        var DefaultRequiredDialogSetRequirement = (function (_super) {
            __extends(DefaultRequiredDialogSetRequirement, _super);
            function DefaultRequiredDialogSetRequirement() {
                return _super.call(this, {
                    "dialogapi": 1.1
                }) || this;
            }
            return DefaultRequiredDialogSetRequirement;
        }(DefaultSetRequirement));
        Requirement.DefaultRequiredDialogSetRequirement = DefaultRequiredDialogSetRequirement;
        var DefaultOptionalDialogSetRequirement = (function (_super) {
            __extends(DefaultOptionalDialogSetRequirement, _super);
            function DefaultOptionalDialogSetRequirement() {
                return _super.call(this, {
                    "dialogorigin": 1.1
                }) || this;
            }
            return DefaultOptionalDialogSetRequirement;
        }(DefaultSetRequirement));
        Requirement.DefaultOptionalDialogSetRequirement = DefaultOptionalDialogSetRequirement;
        var ExcelClientDefaultSetRequirement = (function (_super) {
            __extends(ExcelClientDefaultSetRequirement, _super);
            function ExcelClientDefaultSetRequirement() {
                return _super.call(this, {
                    "bindingevents": 1.1,
                    "documentevents": 1.1,
                    "excelapi": 1.1,
                    "matrixbindings": 1.1,
                    "matrixcoercion": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "tablebindings": 1.1,
                    "tablecoercion": 1.1,
                    "textbindings": 1.1,
                    "textcoercion": 1.1
                }) || this;
            }
            return ExcelClientDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.ExcelClientDefaultSetRequirement = ExcelClientDefaultSetRequirement;
        var ExcelClientV1DefaultSetRequirement = (function (_super) {
            __extends(ExcelClientV1DefaultSetRequirement, _super);
            function ExcelClientV1DefaultSetRequirement() {
                var _this = _super.call(this) || this;
                _this._addSetMap({
                    "imagecoercion": 1.1
                });
                return _this;
            }
            return ExcelClientV1DefaultSetRequirement;
        }(ExcelClientDefaultSetRequirement));
        Requirement.ExcelClientV1DefaultSetRequirement = ExcelClientV1DefaultSetRequirement;
        var OutlookClientDefaultSetRequirement = (function (_super) {
            __extends(OutlookClientDefaultSetRequirement, _super);
            function OutlookClientDefaultSetRequirement() {
                return _super.call(this, {
                    "mailbox": 1.3
                }) || this;
            }
            return OutlookClientDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.OutlookClientDefaultSetRequirement = OutlookClientDefaultSetRequirement;
        var WordClientDefaultSetRequirement = (function (_super) {
            __extends(WordClientDefaultSetRequirement, _super);
            function WordClientDefaultSetRequirement() {
                return _super.call(this, {
                    "bindingevents": 1.1,
                    "compressedfile": 1.1,
                    "customxmlparts": 1.1,
                    "documentevents": 1.1,
                    "file": 1.1,
                    "htmlcoercion": 1.1,
                    "matrixbindings": 1.1,
                    "matrixcoercion": 1.1,
                    "ooxmlcoercion": 1.1,
                    "pdffile": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "tablebindings": 1.1,
                    "tablecoercion": 1.1,
                    "textbindings": 1.1,
                    "textcoercion": 1.1,
                    "textfile": 1.1,
                    "wordapi": 1.1
                }) || this;
            }
            return WordClientDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.WordClientDefaultSetRequirement = WordClientDefaultSetRequirement;
        var WordClientV1DefaultSetRequirement = (function (_super) {
            __extends(WordClientV1DefaultSetRequirement, _super);
            function WordClientV1DefaultSetRequirement() {
                var _this = _super.call(this) || this;
                _this._addSetMap({
                    "customxmlparts": 1.2,
                    "wordapi": 1.2,
                    "imagecoercion": 1.1
                });
                return _this;
            }
            return WordClientV1DefaultSetRequirement;
        }(WordClientDefaultSetRequirement));
        Requirement.WordClientV1DefaultSetRequirement = WordClientV1DefaultSetRequirement;
        var PowerpointClientDefaultSetRequirement = (function (_super) {
            __extends(PowerpointClientDefaultSetRequirement, _super);
            function PowerpointClientDefaultSetRequirement() {
                return _super.call(this, {
                    "activeview": 1.1,
                    "compressedfile": 1.1,
                    "documentevents": 1.1,
                    "file": 1.1,
                    "pdffile": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "textcoercion": 1.1
                }) || this;
            }
            return PowerpointClientDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.PowerpointClientDefaultSetRequirement = PowerpointClientDefaultSetRequirement;
        var PowerpointClientV1DefaultSetRequirement = (function (_super) {
            __extends(PowerpointClientV1DefaultSetRequirement, _super);
            function PowerpointClientV1DefaultSetRequirement() {
                var _this = _super.call(this) || this;
                _this._addSetMap({
                    "imagecoercion": 1.1
                });
                return _this;
            }
            return PowerpointClientV1DefaultSetRequirement;
        }(PowerpointClientDefaultSetRequirement));
        Requirement.PowerpointClientV1DefaultSetRequirement = PowerpointClientV1DefaultSetRequirement;
        var ProjectClientDefaultSetRequirement = (function (_super) {
            __extends(ProjectClientDefaultSetRequirement, _super);
            function ProjectClientDefaultSetRequirement() {
                return _super.call(this, {
                    "selection": 1.1,
                    "textcoercion": 1.1
                }) || this;
            }
            return ProjectClientDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.ProjectClientDefaultSetRequirement = ProjectClientDefaultSetRequirement;
        var ExcelWebDefaultSetRequirement = (function (_super) {
            __extends(ExcelWebDefaultSetRequirement, _super);
            function ExcelWebDefaultSetRequirement() {
                return _super.call(this, {
                    "bindingevents": 1.1,
                    "documentevents": 1.1,
                    "matrixbindings": 1.1,
                    "matrixcoercion": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "tablebindings": 1.1,
                    "tablecoercion": 1.1,
                    "textbindings": 1.1,
                    "textcoercion": 1.1,
                    "file": 1.1
                }) || this;
            }
            return ExcelWebDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.ExcelWebDefaultSetRequirement = ExcelWebDefaultSetRequirement;
        var WordWebDefaultSetRequirement = (function (_super) {
            __extends(WordWebDefaultSetRequirement, _super);
            function WordWebDefaultSetRequirement() {
                return _super.call(this, {
                    "compressedfile": 1.1,
                    "documentevents": 1.1,
                    "file": 1.1,
                    "imagecoercion": 1.1,
                    "matrixcoercion": 1.1,
                    "ooxmlcoercion": 1.1,
                    "pdffile": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "tablecoercion": 1.1,
                    "textcoercion": 1.1,
                    "textfile": 1.1
                }) || this;
            }
            return WordWebDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.WordWebDefaultSetRequirement = WordWebDefaultSetRequirement;
        var PowerpointWebDefaultSetRequirement = (function (_super) {
            __extends(PowerpointWebDefaultSetRequirement, _super);
            function PowerpointWebDefaultSetRequirement() {
                return _super.call(this, {
                    "activeview": 1.1,
                    "settings": 1.1
                }) || this;
            }
            return PowerpointWebDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.PowerpointWebDefaultSetRequirement = PowerpointWebDefaultSetRequirement;
        var OutlookWebDefaultSetRequirement = (function (_super) {
            __extends(OutlookWebDefaultSetRequirement, _super);
            function OutlookWebDefaultSetRequirement() {
                return _super.call(this, {
                    "mailbox": 1.3
                }) || this;
            }
            return OutlookWebDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.OutlookWebDefaultSetRequirement = OutlookWebDefaultSetRequirement;
        var SwayWebDefaultSetRequirement = (function (_super) {
            __extends(SwayWebDefaultSetRequirement, _super);
            function SwayWebDefaultSetRequirement() {
                return _super.call(this, {
                    "activeview": 1.1,
                    "documentevents": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "textcoercion": 1.1
                }) || this;
            }
            return SwayWebDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.SwayWebDefaultSetRequirement = SwayWebDefaultSetRequirement;
        var AccessWebDefaultSetRequirement = (function (_super) {
            __extends(AccessWebDefaultSetRequirement, _super);
            function AccessWebDefaultSetRequirement() {
                return _super.call(this, {
                    "bindingevents": 1.1,
                    "partialtablebindings": 1.1,
                    "settings": 1.1,
                    "tablebindings": 1.1,
                    "tablecoercion": 1.1
                }) || this;
            }
            return AccessWebDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.AccessWebDefaultSetRequirement = AccessWebDefaultSetRequirement;
        var ExcelIOSDefaultSetRequirement = (function (_super) {
            __extends(ExcelIOSDefaultSetRequirement, _super);
            function ExcelIOSDefaultSetRequirement() {
                return _super.call(this, {
                    "bindingevents": 1.1,
                    "documentevents": 1.1,
                    "matrixbindings": 1.1,
                    "matrixcoercion": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "tablebindings": 1.1,
                    "tablecoercion": 1.1,
                    "textbindings": 1.1,
                    "textcoercion": 1.1
                }) || this;
            }
            return ExcelIOSDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.ExcelIOSDefaultSetRequirement = ExcelIOSDefaultSetRequirement;
        var WordIOSDefaultSetRequirement = (function (_super) {
            __extends(WordIOSDefaultSetRequirement, _super);
            function WordIOSDefaultSetRequirement() {
                return _super.call(this, {
                    "bindingevents": 1.1,
                    "compressedfile": 1.1,
                    "customxmlparts": 1.1,
                    "documentevents": 1.1,
                    "file": 1.1,
                    "htmlcoercion": 1.1,
                    "matrixbindings": 1.1,
                    "matrixcoercion": 1.1,
                    "ooxmlcoercion": 1.1,
                    "pdffile": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "tablebindings": 1.1,
                    "tablecoercion": 1.1,
                    "textbindings": 1.1,
                    "textcoercion": 1.1,
                    "textfile": 1.1
                }) || this;
            }
            return WordIOSDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.WordIOSDefaultSetRequirement = WordIOSDefaultSetRequirement;
        var WordIOSV1DefaultSetRequirement = (function (_super) {
            __extends(WordIOSV1DefaultSetRequirement, _super);
            function WordIOSV1DefaultSetRequirement() {
                var _this = _super.call(this) || this;
                _this._addSetMap({
                    "customxmlparts": 1.2,
                    "wordapi": 1.2
                });
                return _this;
            }
            return WordIOSV1DefaultSetRequirement;
        }(WordIOSDefaultSetRequirement));
        Requirement.WordIOSV1DefaultSetRequirement = WordIOSV1DefaultSetRequirement;
        var PowerpointIOSDefaultSetRequirement = (function (_super) {
            __extends(PowerpointIOSDefaultSetRequirement, _super);
            function PowerpointIOSDefaultSetRequirement() {
                return _super.call(this, {
                    "activeview": 1.1,
                    "compressedfile": 1.1,
                    "documentevents": 1.1,
                    "file": 1.1,
                    "pdffile": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "textcoercion": 1.1
                }) || this;
            }
            return PowerpointIOSDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.PowerpointIOSDefaultSetRequirement = PowerpointIOSDefaultSetRequirement;
        var OutlookIOSDefaultSetRequirement = (function (_super) {
            __extends(OutlookIOSDefaultSetRequirement, _super);
            function OutlookIOSDefaultSetRequirement() {
                return _super.call(this, {
                    "mailbox": 1.1
                }) || this;
            }
            return OutlookIOSDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.OutlookIOSDefaultSetRequirement = OutlookIOSDefaultSetRequirement;
        var RequirementsMatrixFactory = (function () {
            function RequirementsMatrixFactory() {
            }
            RequirementsMatrixFactory.initializeOsfDda = function () {
                OSF.OUtil.setNamespace("Requirement", OSF.DDA);
            };
            RequirementsMatrixFactory.getDefaultRequirementMatrix = function (appContext) {
                this.initializeDefaultSetMatrix();
                var defaultRequirementMatrix = undefined;
                var clientRequirement = appContext.get_requirementMatrix();
                if (clientRequirement != undefined && clientRequirement.length > 0 && typeof (JSON) !== "undefined") {
                    var matrixItem = JSON.parse(appContext.get_requirementMatrix().toLowerCase());
                    try {
                        var setName = "dialogorigin";
                        if (!matrixItem.hasOwnProperty(setName)) {
                            matrixItem[setName] = 1.1;
                        }
                    }
                    catch (ex) { }
                    defaultRequirementMatrix = new RequirementMatrix(new DefaultSetRequirement(matrixItem));
                }
                else {
                    var appLocator = RequirementsMatrixFactory.getClientFullVersionString(appContext);
                    if (RequirementsMatrixFactory.DefaultSetArrayMatrix != undefined && RequirementsMatrixFactory.DefaultSetArrayMatrix[appLocator] != undefined) {
                        defaultRequirementMatrix = new RequirementMatrix(RequirementsMatrixFactory.DefaultSetArrayMatrix[appLocator]);
                    }
                    else {
                        defaultRequirementMatrix = new RequirementMatrix(new DefaultSetRequirement({}));
                    }
                }
                return defaultRequirementMatrix;
            };
            RequirementsMatrixFactory.getDefaultDialogRequirementMatrix = function (appContext) {
                var setRequirements = undefined;
                var clientRequirement = appContext.get_dialogRequirementMatrix();
                if (clientRequirement != undefined && clientRequirement.length > 0 && typeof (JSON) !== "undefined") {
                    var matrixItem = JSON.parse(appContext.get_requirementMatrix().toLowerCase());
                    setRequirements = new DefaultSetRequirement(matrixItem);
                }
                else {
                    setRequirements = new DefaultRequiredDialogSetRequirement();
                    var mainRequirement = appContext.get_requirementMatrix();
                    if (mainRequirement != undefined && mainRequirement.length > 0 && typeof (JSON) !== "undefined") {
                        var matrixItem = JSON.parse(mainRequirement.toLowerCase());
                        for (var name in setRequirements._sets) {
                            if (matrixItem.hasOwnProperty(name)) {
                                setRequirements._sets[name] = matrixItem[name];
                            }
                        }
                        var dialogOptionalSetRequirement = new DefaultOptionalDialogSetRequirement();
                        for (var name in dialogOptionalSetRequirement._sets) {
                            if (matrixItem.hasOwnProperty(name)) {
                                setRequirements._sets[name] = matrixItem[name];
                            }
                        }
                    }
                }
                try {
                    var setName = "dialogorigin";
                    if (!setRequirements._sets.hasOwnProperty(setName) && window.opener) {
                        setRequirements._sets[setName] = 1.1;
                    }
                }
                catch (ex) { }
                return new RequirementMatrix(setRequirements);
            };
            RequirementsMatrixFactory.getClientFullVersionString = function (appContext) {
                var appMinorVersion = appContext.get_appMinorVersion();
                var appMinorVersionString = "";
                var appFullVersion = "";
                var appName = appContext.get_appName();
                var isIOSClient = appName == 1024 ||
                    appName == 4096 ||
                    appName == 8192 ||
                    appName == 65536;
                if (isIOSClient && appContext.get_appVersion() == 1) {
                    if (appName == 4096 && appMinorVersion >= 15) {
                        appFullVersion = "16.00.01";
                    }
                    else {
                        appFullVersion = "16.00";
                    }
                }
                else if (appContext.get_appName() == 64) {
                    appFullVersion = appContext.get_appVersion();
                }
                else {
                    if (appMinorVersion < 10) {
                        appMinorVersionString = "0" + appMinorVersion;
                    }
                    else {
                        appMinorVersionString = "" + appMinorVersion;
                    }
                    appFullVersion = appContext.get_appVersion() + "." + appMinorVersionString;
                }
                return appContext.get_appName() + "-" + appFullVersion;
            };
            RequirementsMatrixFactory.initializeDefaultSetMatrix = function () {
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_RCLIENT_1600] = new ExcelClientDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_RCLIENT_1600] = new WordClientDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_RCLIENT_1600] = new PowerpointClientDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_RCLIENT_1601] = new ExcelClientV1DefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_RCLIENT_1601] = new WordClientV1DefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_RCLIENT_1601] = new PowerpointClientV1DefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_RCLIENT_1600] = new OutlookClientDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_WAC_1600] = new ExcelWebDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_WAC_1600] = new WordWebDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_WAC_1600] = new OutlookWebDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_WAC_1601] = new OutlookWebDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Project_RCLIENT_1600] = new ProjectClientDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Access_WAC_1600] = new AccessWebDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_WAC_1600] = new PowerpointWebDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_IOS_1600] = new ExcelIOSDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.SWAY_WAC_1600] = new SwayWebDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_IOS_1600] = new WordIOSDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_IOS_16001] = new WordIOSV1DefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_IOS_1600] = new PowerpointIOSDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_IOS_1600] = new OutlookIOSDefaultSetRequirement();
            };
            RequirementsMatrixFactory.Excel_RCLIENT_1600 = "1-16.00";
            RequirementsMatrixFactory.Excel_RCLIENT_1601 = "1-16.01";
            RequirementsMatrixFactory.Word_RCLIENT_1600 = "2-16.00";
            RequirementsMatrixFactory.Word_RCLIENT_1601 = "2-16.01";
            RequirementsMatrixFactory.PowerPoint_RCLIENT_1600 = "4-16.00";
            RequirementsMatrixFactory.PowerPoint_RCLIENT_1601 = "4-16.01";
            RequirementsMatrixFactory.Outlook_RCLIENT_1600 = "8-16.00";
            RequirementsMatrixFactory.Excel_WAC_1600 = "16-16.00";
            RequirementsMatrixFactory.Word_WAC_1600 = "32-16.00";
            RequirementsMatrixFactory.Outlook_WAC_1600 = "64-16.00";
            RequirementsMatrixFactory.Outlook_WAC_1601 = "64-16.01";
            RequirementsMatrixFactory.Project_RCLIENT_1600 = "128-16.00";
            RequirementsMatrixFactory.Access_WAC_1600 = "256-16.00";
            RequirementsMatrixFactory.PowerPoint_WAC_1600 = "512-16.00";
            RequirementsMatrixFactory.Excel_IOS_1600 = "1024-16.00";
            RequirementsMatrixFactory.SWAY_WAC_1600 = "2048-16.00";
            RequirementsMatrixFactory.Word_IOS_1600 = "4096-16.00";
            RequirementsMatrixFactory.Word_IOS_16001 = "4096-16.00.01";
            RequirementsMatrixFactory.PowerPoint_IOS_1600 = "8192-16.00";
            RequirementsMatrixFactory.Outlook_IOS_1600 = "65536-16.00";
            RequirementsMatrixFactory.DefaultSetArrayMatrix = {};
            return RequirementsMatrixFactory;
        }());
        Requirement.RequirementsMatrixFactory = RequirementsMatrixFactory;
    })(Requirement = OfficeExt.Requirement || (OfficeExt.Requirement = {}));
})(OfficeExt || (OfficeExt = {}));
OfficeExt.Requirement.RequirementsMatrixFactory.initializeOsfDda();
Microsoft.Office.WebExtension.ApplicationMode = {
    WebEditor: "webEditor",
    WebViewer: "webViewer",
    Client: "client"
};
Microsoft.Office.WebExtension.DocumentMode = {
    ReadOnly: "readOnly",
    ReadWrite: "readWrite"
};
OSF.NamespaceManager = (function OSF_NamespaceManager() {
    var _userOffice;
    var _useShortcut = false;
    return {
        enableShortcut: function OSF_NamespaceManager$enableShortcut() {
            if (!_useShortcut) {
                if (window.Office) {
                    _userOffice = window.Office;
                }
                else {
                    OSF.OUtil.setNamespace("Office", window);
                }
                window.Office = Microsoft.Office.WebExtension;
                _useShortcut = true;
            }
        },
        disableShortcut: function OSF_NamespaceManager$disableShortcut() {
            if (_useShortcut) {
                if (_userOffice) {
                    window.Office = _userOffice;
                }
                else {
                    OSF.OUtil.unsetNamespace("Office", window);
                }
                _useShortcut = false;
            }
        }
    };
})();
OSF.NamespaceManager.enableShortcut();
Microsoft.Office.WebExtension.useShortNamespace = function Microsoft_Office_WebExtension_useShortNamespace(useShortcut) {
    if (useShortcut) {
        OSF.NamespaceManager.enableShortcut();
    }
    else {
        OSF.NamespaceManager.disableShortcut();
    }
};
Microsoft.Office.WebExtension.select = function Microsoft_Office_WebExtension_select(str, errorCallback) {
    var promise;
    if (str && typeof str == "string") {
        var index = str.indexOf("#");
        if (index != -1) {
            var op = str.substring(0, index);
            var target = str.substring(index + 1);
            switch (op) {
                case "binding":
                case "bindings":
                    if (target) {
                        promise = new OSF.DDA.BindingPromise(target);
                    }
                    break;
            }
        }
    }
    if (!promise) {
        if (errorCallback) {
            var callbackType = typeof errorCallback;
            if (callbackType == "function") {
                var callArgs = {};
                callArgs[Microsoft.Office.WebExtension.Parameters.Callback] = errorCallback;
                OSF.DDA.issueAsyncResult(callArgs, OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext, OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext));
            }
            else {
                throw OSF.OUtil.formatString(Strings.OfficeOM.L_CallbackNotAFunction, callbackType);
            }
        }
    }
    else {
        promise.onFail = errorCallback;
        return promise;
    }
};
OSF.DDA.Context = function OSF_DDA_Context(officeAppContext, document, license, appOM, getOfficeTheme) {
    OSF.OUtil.defineEnumerableProperties(this, {
        "contentLanguage": {
            value: officeAppContext.get_dataLocale()
        },
        "displayLanguage": {
            value: officeAppContext.get_appUILocale()
        },
        "touchEnabled": {
            value: officeAppContext.get_touchEnabled()
        },
        "commerceAllowed": {
            value: officeAppContext.get_commerceAllowed()
        },
        "host": {
            value: OfficeExt.HostName.Host.getInstance().getHost()
        },
        "platform": {
            value: OfficeExt.HostName.Host.getInstance().getPlatform()
        },
        "isDialog": {
            value: OSF._OfficeAppFactory.getHostInfo().isDialog
        },
        "diagnostics": {
            value: OfficeExt.HostName.Host.getInstance().getDiagnostics(officeAppContext.get_hostFullVersion())
        }
    });
    if (license) {
        OSF.OUtil.defineEnumerableProperty(this, "license", {
            value: license
        });
    }
    if (officeAppContext.ui) {
        OSF.OUtil.defineEnumerableProperty(this, "ui", {
            value: officeAppContext.ui
        });
    }
    if (officeAppContext.auth) {
        OSF.OUtil.defineEnumerableProperty(this, "auth", {
            value: officeAppContext.auth
        });
    }
    if (officeAppContext.webAuth) {
        OSF.OUtil.defineEnumerableProperty(this, "webAuth", {
            value: officeAppContext.webAuth
        });
    }
    if (officeAppContext.application) {
        OSF.OUtil.defineEnumerableProperty(this, "application", {
            value: officeAppContext.application
        });
    }
    if (officeAppContext.extensionLifeCycle) {
        OSF.OUtil.defineEnumerableProperty(this, "extensionLifeCycle", {
            value: officeAppContext.extensionLifeCycle
        });
    }
    if (officeAppContext.messaging) {
        OSF.OUtil.defineEnumerableProperty(this, "messaging", {
            value: officeAppContext.messaging
        });
    }
    if (officeAppContext.ui && officeAppContext.ui.taskPaneAction) {
        OSF.OUtil.defineEnumerableProperty(this, "taskPaneAction", {
            value: officeAppContext.ui.taskPaneAction
        });
    }
    if (officeAppContext.ui && officeAppContext.ui.ribbonGallery) {
        OSF.OUtil.defineEnumerableProperty(this, "ribbonGallery", {
            value: officeAppContext.ui.ribbonGallery
        });
    }
    if (officeAppContext.get_isDialog()) {
        var requirements = OfficeExt.Requirement.RequirementsMatrixFactory.getDefaultDialogRequirementMatrix(officeAppContext);
        OSF.OUtil.defineEnumerableProperty(this, "requirements", {
            value: requirements
        });
    }
    else {
        if (document) {
            OSF.OUtil.defineEnumerableProperty(this, "document", {
                value: document
            });
        }
        if (appOM) {
            var displayName = appOM.displayName || "appOM";
            delete appOM.displayName;
            OSF.OUtil.defineEnumerableProperty(this, displayName, {
                value: appOM
            });
        }
        if (officeAppContext.get_officeTheme()) {
            OSF.OUtil.defineEnumerableProperty(this, "officeTheme", {
                get: function () {
                    return officeAppContext.get_officeTheme();
                }
            });
        }
        else if (getOfficeTheme) {
            OSF.OUtil.defineEnumerableProperty(this, "officeTheme", {
                get: function () {
                    return getOfficeTheme();
                }
            });
        }
        var requirements = OfficeExt.Requirement.RequirementsMatrixFactory.getDefaultRequirementMatrix(officeAppContext);
        OSF.OUtil.defineEnumerableProperty(this, "requirements", {
            value: requirements
        });
    }
};
OSF.DDA.OutlookContext = function OSF_DDA_OutlookContext(appContext, settings, license, appOM, getOfficeTheme) {
    OSF.DDA.OutlookContext.uber.constructor.call(this, appContext, null, license, appOM, getOfficeTheme);
    if (settings) {
        OSF.OUtil.defineEnumerableProperty(this, "roamingSettings", {
            value: settings
        });
    }
};
OSF.OUtil.extend(OSF.DDA.OutlookContext, OSF.DDA.Context);
OSF.DDA.OutlookAppOm = function OSF_DDA_OutlookAppOm(appContext, window, appReady) { };
OSF.DDA.Application = function OSF_DDA_Application(officeAppContext) {
};
OSF.DDA.Document = function OSF_DDA_Document(officeAppContext, settings) {
    var mode;
    switch (officeAppContext.get_clientMode()) {
        case OSF.ClientMode.ReadOnly:
            mode = Microsoft.Office.WebExtension.DocumentMode.ReadOnly;
            break;
        case OSF.ClientMode.ReadWrite:
            mode = Microsoft.Office.WebExtension.DocumentMode.ReadWrite;
            break;
    }
    ;
    if (settings) {
        OSF.OUtil.defineEnumerableProperty(this, "settings", {
            value: settings
        });
    }
    ;
    OSF.OUtil.defineMutableProperties(this, {
        "mode": {
            value: mode
        },
        "url": {
            value: officeAppContext.get_docUrl()
        }
    });
};
OSF.DDA.JsomDocument = function OSF_DDA_JsomDocument(officeAppContext, bindingFacade, settings) {
    OSF.DDA.JsomDocument.uber.constructor.call(this, officeAppContext, settings);
    if (bindingFacade) {
        OSF.OUtil.defineEnumerableProperty(this, "bindings", {
            get: function OSF_DDA_Document$GetBindings() { return bindingFacade; }
        });
    }
    var am = OSF.DDA.AsyncMethodNames;
    OSF.DDA.DispIdHost.addAsyncMethods(this, [
        am.GetSelectedDataAsync,
        am.SetSelectedDataAsync
    ]);
    OSF.DDA.DispIdHost.addEventSupport(this, new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged]));
};
OSF.OUtil.extend(OSF.DDA.JsomDocument, OSF.DDA.Document);
OSF.OUtil.defineEnumerableProperty(Microsoft.Office.WebExtension, "context", {
    get: function Microsoft_Office_WebExtension$GetContext() {
        var context;
        if (OSF && OSF._OfficeAppFactory) {
            context = OSF._OfficeAppFactory.getContext();
        }
        return context;
    }
});
OSF.DDA.License = function OSF_DDA_License(eToken) {
    OSF.OUtil.defineEnumerableProperty(this, "value", {
        value: eToken
    });
};
OSF.DDA.ApiMethodCall = function OSF_DDA_ApiMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName) {
    var requiredCount = requiredParameters.length;
    var getInvalidParameterString = OSF.OUtil.delayExecutionAndCache(function () {
        return OSF.OUtil.formatString(Strings.OfficeOM.L_InvalidParameters, displayName);
    });
    this.verifyArguments = function OSF_DDA_ApiMethodCall$VerifyArguments(params, args) {
        for (var name in params) {
            var param = params[name];
            var arg = args[name];
            if (param["enum"]) {
                switch (typeof arg) {
                    case "string":
                        if (OSF.OUtil.listContainsValue(param["enum"], arg)) {
                            break;
                        }
                    case "undefined":
                        throw OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration;
                    default:
                        throw getInvalidParameterString();
                }
            }
            if (param["types"]) {
                if (!OSF.OUtil.listContainsValue(param["types"], typeof arg)) {
                    throw getInvalidParameterString();
                }
            }
        }
    };
    this.extractRequiredArguments = function OSF_DDA_ApiMethodCall$ExtractRequiredArguments(userArgs, caller, stateInfo) {
        if (userArgs.length < requiredCount) {
            throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_MissingRequiredArguments);
        }
        var requiredArgs = [];
        var index;
        for (index = 0; index < requiredCount; index++) {
            requiredArgs.push(userArgs[index]);
        }
        this.verifyArguments(requiredParameters, requiredArgs);
        var ret = {};
        for (index = 0; index < requiredCount; index++) {
            var param = requiredParameters[index];
            var arg = requiredArgs[index];
            if (param.verify) {
                var isValid = param.verify(arg, caller, stateInfo);
                if (!isValid) {
                    throw getInvalidParameterString();
                }
            }
            ret[param.name] = arg;
        }
        return ret;
    },
        this.fillOptions = function OSF_DDA_ApiMethodCall$FillOptions(options, requiredArgs, caller, stateInfo) {
            options = options || {};
            for (var optionName in supportedOptions) {
                if (!OSF.OUtil.listContainsKey(options, optionName)) {
                    var value = undefined;
                    var option = supportedOptions[optionName];
                    if (option.calculate && requiredArgs) {
                        value = option.calculate(requiredArgs, caller, stateInfo);
                    }
                    if (!value && option.defaultValue !== undefined) {
                        value = option.defaultValue;
                    }
                    options[optionName] = value;
                }
            }
            return options;
        };
    this.constructCallArgs = function OSF_DAA_ApiMethodCall$ConstructCallArgs(required, options, caller, stateInfo) {
        var callArgs = {};
        for (var r in required) {
            callArgs[r] = required[r];
        }
        for (var o in options) {
            callArgs[o] = options[o];
        }
        for (var s in privateStateCallbacks) {
            callArgs[s] = privateStateCallbacks[s](caller, stateInfo);
        }
        if (checkCallArgs) {
            callArgs = checkCallArgs(callArgs, caller, stateInfo);
        }
        return callArgs;
    };
};
OSF.OUtil.setNamespace("AsyncResultEnum", OSF.DDA);
OSF.DDA.AsyncResultEnum.Properties = {
    Context: "Context",
    Value: "Value",
    Status: "Status",
    Error: "Error"
};
Microsoft.Office.WebExtension.AsyncResultStatus = {
    Succeeded: "succeeded",
    Failed: "failed"
};
OSF.DDA.AsyncResultEnum.ErrorCode = {
    Success: 0,
    Failed: 1
};
OSF.DDA.AsyncResultEnum.ErrorProperties = {
    Name: "Name",
    Message: "Message",
    Code: "Code"
};
OSF.DDA.AsyncMethodNames = {};
OSF.DDA.AsyncMethodNames.addNames = function (methodNames) {
    for (var entry in methodNames) {
        var am = {};
        OSF.OUtil.defineEnumerableProperties(am, {
            "id": {
                value: entry
            },
            "displayName": {
                value: methodNames[entry]
            }
        });
        OSF.DDA.AsyncMethodNames[entry] = am;
    }
};
OSF.DDA.AsyncMethodCall = function OSF_DDA_AsyncMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, onSucceeded, onFailed, checkCallArgs, displayName) {
    var requiredCount = requiredParameters.length;
    var apiMethods = new OSF.DDA.ApiMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName);
    function OSF_DAA_AsyncMethodCall$ExtractOptions(userArgs, requiredArgs, caller, stateInfo) {
        if (userArgs.length > requiredCount + 2) {
            throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyArguments);
        }
        var options, parameterCallback;
        for (var i = userArgs.length - 1; i >= requiredCount; i--) {
            var argument = userArgs[i];
            switch (typeof argument) {
                case "object":
                    if (options) {
                        throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalObjects);
                    }
                    else {
                        options = argument;
                    }
                    break;
                case "function":
                    if (parameterCallback) {
                        throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalFunction);
                    }
                    else {
                        parameterCallback = argument;
                    }
                    break;
                default:
                    throw OsfMsAjaxFactory.msAjaxError.argument(Strings.OfficeOM.L_InValidOptionalArgument);
                    break;
            }
        }
        options = apiMethods.fillOptions(options, requiredArgs, caller, stateInfo);
        if (parameterCallback) {
            if (options[Microsoft.Office.WebExtension.Parameters.Callback]) {
                throw Strings.OfficeOM.L_RedundantCallbackSpecification;
            }
            else {
                options[Microsoft.Office.WebExtension.Parameters.Callback] = parameterCallback;
            }
        }
        apiMethods.verifyArguments(supportedOptions, options);
        return options;
    }
    ;
    this.verifyAndExtractCall = function OSF_DAA_AsyncMethodCall$VerifyAndExtractCall(userArgs, caller, stateInfo) {
        var required = apiMethods.extractRequiredArguments(userArgs, caller, stateInfo);
        var options = OSF_DAA_AsyncMethodCall$ExtractOptions(userArgs, required, caller, stateInfo);
        var callArgs = apiMethods.constructCallArgs(required, options, caller, stateInfo);
        return callArgs;
    };
    this.processResponse = function OSF_DAA_AsyncMethodCall$ProcessResponse(status, response, caller, callArgs) {
        var payload;
        if (status == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
            if (onSucceeded) {
                payload = onSucceeded(response, caller, callArgs);
            }
            else {
                payload = response;
            }
        }
        else {
            if (onFailed) {
                payload = onFailed(status, response);
            }
            else {
                payload = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
            }
        }
        return payload;
    };
    this.getCallArgs = function (suppliedArgs) {
        var options, parameterCallback;
        for (var i = suppliedArgs.length - 1; i >= requiredCount; i--) {
            var argument = suppliedArgs[i];
            switch (typeof argument) {
                case "object":
                    options = argument;
                    break;
                case "function":
                    parameterCallback = argument;
                    break;
            }
        }
        options = options || {};
        if (parameterCallback) {
            options[Microsoft.Office.WebExtension.Parameters.Callback] = parameterCallback;
        }
        return options;
    };
};
OSF.DDA.AsyncMethodCallFactory = (function () {
    return {
        manufacture: function (params) {
            var supportedOptions = params.supportedOptions ? OSF.OUtil.createObject(params.supportedOptions) : [];
            var privateStateCallbacks = params.privateStateCallbacks ? OSF.OUtil.createObject(params.privateStateCallbacks) : [];
            return new OSF.DDA.AsyncMethodCall(params.requiredArguments || [], supportedOptions, privateStateCallbacks, params.onSucceeded, params.onFailed, params.checkCallArgs, params.method.displayName);
        }
    };
})();
OSF.DDA.AsyncMethodCalls = {};
OSF.DDA.AsyncMethodCalls.define = function (callDefinition) {
    OSF.DDA.AsyncMethodCalls[callDefinition.method.id] = OSF.DDA.AsyncMethodCallFactory.manufacture(callDefinition);
};
OSF.DDA.Error = function OSF_DDA_Error(name, message, code) {
    OSF.OUtil.defineEnumerableProperties(this, {
        "name": {
            value: name
        },
        "message": {
            value: message
        },
        "code": {
            value: code
        }
    });
};
OSF.DDA.AsyncResult = function OSF_DDA_AsyncResult(initArgs, errorArgs) {
    OSF.OUtil.defineEnumerableProperties(this, {
        "value": {
            value: initArgs[OSF.DDA.AsyncResultEnum.Properties.Value]
        },
        "status": {
            value: errorArgs ? Microsoft.Office.WebExtension.AsyncResultStatus.Failed : Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded
        }
    });
    if (initArgs[OSF.DDA.AsyncResultEnum.Properties.Context]) {
        OSF.OUtil.defineEnumerableProperty(this, "asyncContext", {
            value: initArgs[OSF.DDA.AsyncResultEnum.Properties.Context]
        });
    }
    if (errorArgs) {
        OSF.OUtil.defineEnumerableProperty(this, "error", {
            value: new OSF.DDA.Error(errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code])
        });
    }
};
OSF.DDA.issueAsyncResult = function OSF_DDA$IssueAsyncResult(callArgs, status, payload) {
    var callback = callArgs[Microsoft.Office.WebExtension.Parameters.Callback];
    if (callback) {
        var asyncInitArgs = {};
        asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Context] = callArgs[Microsoft.Office.WebExtension.Parameters.AsyncContext];
        var errorArgs;
        if (status == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
            asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Value] = payload;
        }
        else {
            errorArgs = {};
            payload = payload || OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
            errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code] = status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
            errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name] = payload.name || payload;
            errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message] = payload.message || payload;
        }
        callback(new OSF.DDA.AsyncResult(asyncInitArgs, errorArgs));
    }
};
OSF.DDA.SyncMethodNames = {};
OSF.DDA.SyncMethodNames.addNames = function (methodNames) {
    for (var entry in methodNames) {
        var am = {};
        OSF.OUtil.defineEnumerableProperties(am, {
            "id": {
                value: entry
            },
            "displayName": {
                value: methodNames[entry]
            }
        });
        OSF.DDA.SyncMethodNames[entry] = am;
    }
};
OSF.DDA.SyncMethodCall = function OSF_DDA_SyncMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName) {
    var requiredCount = requiredParameters.length;
    var apiMethods = new OSF.DDA.ApiMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName);
    function OSF_DAA_SyncMethodCall$ExtractOptions(userArgs, requiredArgs, caller, stateInfo) {
        if (userArgs.length > requiredCount + 1) {
            throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyArguments);
        }
        var options, parameterCallback;
        for (var i = userArgs.length - 1; i >= requiredCount; i--) {
            var argument = userArgs[i];
            switch (typeof argument) {
                case "object":
                    if (options) {
                        throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalObjects);
                    }
                    else {
                        options = argument;
                    }
                    break;
                default:
                    throw OsfMsAjaxFactory.msAjaxError.argument(Strings.OfficeOM.L_InValidOptionalArgument);
                    break;
            }
        }
        options = apiMethods.fillOptions(options, requiredArgs, caller, stateInfo);
        apiMethods.verifyArguments(supportedOptions, options);
        return options;
    }
    ;
    this.verifyAndExtractCall = function OSF_DAA_AsyncMethodCall$VerifyAndExtractCall(userArgs, caller, stateInfo) {
        var required = apiMethods.extractRequiredArguments(userArgs, caller, stateInfo);
        var options = OSF_DAA_SyncMethodCall$ExtractOptions(userArgs, required, caller, stateInfo);
        var callArgs = apiMethods.constructCallArgs(required, options, caller, stateInfo);
        return callArgs;
    };
};
OSF.DDA.SyncMethodCallFactory = (function () {
    return {
        manufacture: function (params) {
            var supportedOptions = params.supportedOptions ? OSF.OUtil.createObject(params.supportedOptions) : [];
            return new OSF.DDA.SyncMethodCall(params.requiredArguments || [], supportedOptions, params.privateStateCallbacks, params.checkCallArgs, params.method.displayName);
        }
    };
})();
OSF.DDA.SyncMethodCalls = {};
OSF.DDA.SyncMethodCalls.define = function (callDefinition) {
    OSF.DDA.SyncMethodCalls[callDefinition.method.id] = OSF.DDA.SyncMethodCallFactory.manufacture(callDefinition);
};
OSF.DDA.ListType = (function () {
    var listTypes = {};
    return {
        setListType: function OSF_DDA_ListType$AddListType(t, prop) { listTypes[t] = prop; },
        isListType: function OSF_DDA_ListType$IsListType(t) { return OSF.OUtil.listContainsKey(listTypes, t); },
        getDescriptor: function OSF_DDA_ListType$getDescriptor(t) { return listTypes[t]; }
    };
})();
OSF.DDA.HostParameterMap = function (specialProcessor, mappings) {
    var toHostMap = "toHost";
    var fromHostMap = "fromHost";
    var sourceData = "sourceData";
    var self = "self";
    var dynamicTypes = {};
    dynamicTypes[Microsoft.Office.WebExtension.Parameters.Data] = {
        toHost: function (data) {
            if (data != null && data.rows !== undefined) {
                var tableData = {};
                tableData[OSF.DDA.TableDataProperties.TableRows] = data.rows;
                tableData[OSF.DDA.TableDataProperties.TableHeaders] = data.headers;
                data = tableData;
            }
            return data;
        },
        fromHost: function (args) {
            return args;
        }
    };
    dynamicTypes[Microsoft.Office.WebExtension.Parameters.SampleData] = dynamicTypes[Microsoft.Office.WebExtension.Parameters.Data];
    function mapValues(preimageSet, mapping) {
        var ret = preimageSet ? {} : undefined;
        for (var entry in preimageSet) {
            var preimage = preimageSet[entry];
            var image;
            if (OSF.DDA.ListType.isListType(entry)) {
                image = [];
                for (var subEntry in preimage) {
                    image.push(mapValues(preimage[subEntry], mapping));
                }
            }
            else if (OSF.OUtil.listContainsKey(dynamicTypes, entry)) {
                image = dynamicTypes[entry][mapping](preimage);
            }
            else if (mapping == fromHostMap && specialProcessor.preserveNesting(entry)) {
                image = mapValues(preimage, mapping);
            }
            else {
                var maps = mappings[entry];
                if (maps) {
                    var map = maps[mapping];
                    if (map) {
                        image = map[preimage];
                        if (image === undefined) {
                            image = preimage;
                        }
                    }
                }
                else {
                    image = preimage;
                }
            }
            ret[entry] = image;
        }
        return ret;
    }
    ;
    function generateArguments(imageSet, parameters) {
        var ret;
        for (var param in parameters) {
            var arg;
            if (specialProcessor.isComplexType(param)) {
                arg = generateArguments(imageSet, mappings[param][toHostMap]);
            }
            else {
                arg = imageSet[param];
            }
            if (arg != undefined) {
                if (!ret) {
                    ret = {};
                }
                var index = parameters[param];
                if (index == self) {
                    index = param;
                }
                ret[index] = specialProcessor.pack(param, arg);
            }
        }
        return ret;
    }
    ;
    function extractArguments(source, parameters, extracted) {
        if (!extracted) {
            extracted = {};
        }
        for (var param in parameters) {
            var index = parameters[param];
            var value;
            if (index == self) {
                value = source;
            }
            else if (index == sourceData) {
                extracted[param] = source.toArray();
                continue;
            }
            else {
                value = source[index];
            }
            if (value === null || value === undefined) {
                extracted[param] = undefined;
            }
            else {
                value = specialProcessor.unpack(param, value);
                var map;
                if (specialProcessor.isComplexType(param)) {
                    map = mappings[param][fromHostMap];
                    if (specialProcessor.preserveNesting(param)) {
                        extracted[param] = extractArguments(value, map);
                    }
                    else {
                        extractArguments(value, map, extracted);
                    }
                }
                else {
                    if (OSF.DDA.ListType.isListType(param)) {
                        map = {};
                        var entryDescriptor = OSF.DDA.ListType.getDescriptor(param);
                        map[entryDescriptor] = self;
                        var extractedValues = new Array(value.length);
                        for (var item in value) {
                            extractedValues[item] = extractArguments(value[item], map);
                        }
                        extracted[param] = extractedValues;
                    }
                    else {
                        extracted[param] = value;
                    }
                }
            }
        }
        return extracted;
    }
    ;
    function applyMap(mapName, preimage, mapping) {
        var parameters = mappings[mapName][mapping];
        var image;
        if (mapping == "toHost") {
            var imageSet = mapValues(preimage, mapping);
            image = generateArguments(imageSet, parameters);
        }
        else if (mapping == "fromHost") {
            var argumentSet = extractArguments(preimage, parameters);
            image = mapValues(argumentSet, mapping);
        }
        return image;
    }
    ;
    if (!mappings) {
        mappings = {};
    }
    this.addMapping = function (mapName, description) {
        var toHost, fromHost;
        if (description.map) {
            toHost = description.map;
            fromHost = {};
            for (var preimage in toHost) {
                var image = toHost[preimage];
                if (image == self) {
                    image = preimage;
                }
                fromHost[image] = preimage;
            }
        }
        else {
            toHost = description.toHost;
            fromHost = description.fromHost;
        }
        var pair = mappings[mapName];
        if (pair) {
            var currMap = pair[toHostMap];
            for (var th in currMap)
                toHost[th] = currMap[th];
            currMap = pair[fromHostMap];
            for (var fh in currMap)
                fromHost[fh] = currMap[fh];
        }
        else {
            pair = mappings[mapName] = {};
        }
        pair[toHostMap] = toHost;
        pair[fromHostMap] = fromHost;
    };
    this.toHost = function (mapName, preimage) { return applyMap(mapName, preimage, toHostMap); };
    this.fromHost = function (mapName, image) { return applyMap(mapName, image, fromHostMap); };
    this.self = self;
    this.sourceData = sourceData;
    this.addComplexType = function (ct) { specialProcessor.addComplexType(ct); };
    this.getDynamicType = function (dt) { return specialProcessor.getDynamicType(dt); };
    this.setDynamicType = function (dt, handler) { specialProcessor.setDynamicType(dt, handler); };
    this.dynamicTypes = dynamicTypes;
    this.doMapValues = function (preimageSet, mapping) { return mapValues(preimageSet, mapping); };
};
OSF.DDA.SpecialProcessor = function (complexTypes, dynamicTypes) {
    this.addComplexType = function OSF_DDA_SpecialProcessor$addComplexType(ct) {
        complexTypes.push(ct);
    };
    this.getDynamicType = function OSF_DDA_SpecialProcessor$getDynamicType(dt) {
        return dynamicTypes[dt];
    };
    this.setDynamicType = function OSF_DDA_SpecialProcessor$setDynamicType(dt, handler) {
        dynamicTypes[dt] = handler;
    };
    this.isComplexType = function OSF_DDA_SpecialProcessor$isComplexType(t) {
        return OSF.OUtil.listContainsValue(complexTypes, t);
    };
    this.isDynamicType = function OSF_DDA_SpecialProcessor$isDynamicType(p) {
        return OSF.OUtil.listContainsKey(dynamicTypes, p);
    };
    this.preserveNesting = function OSF_DDA_SpecialProcessor$preserveNesting(p) {
        var pn = [];
        if (OSF.DDA.PropertyDescriptors)
            pn.push(OSF.DDA.PropertyDescriptors.Subset);
        if (OSF.DDA.DataNodeEventProperties) {
            pn = pn.concat([
                OSF.DDA.DataNodeEventProperties.OldNode,
                OSF.DDA.DataNodeEventProperties.NewNode,
                OSF.DDA.DataNodeEventProperties.NextSiblingNode
            ]);
        }
        return OSF.OUtil.listContainsValue(pn, p);
    };
    this.pack = function OSF_DDA_SpecialProcessor$pack(param, arg) {
        var value;
        if (this.isDynamicType(param)) {
            value = dynamicTypes[param].toHost(arg);
        }
        else {
            value = arg;
        }
        return value;
    };
    this.unpack = function OSF_DDA_SpecialProcessor$unpack(param, arg) {
        var value;
        if (this.isDynamicType(param)) {
            value = dynamicTypes[param].fromHost(arg);
        }
        else {
            value = arg;
        }
        return value;
    };
};
OSF.DDA.getDecoratedParameterMap = function (specialProcessor, initialDefs) {
    var parameterMap = new OSF.DDA.HostParameterMap(specialProcessor);
    var self = parameterMap.self;
    function createObject(properties) {
        var obj = null;
        if (properties) {
            obj = {};
            var len = properties.length;
            for (var i = 0; i < len; i++) {
                obj[properties[i].name] = properties[i].value;
            }
        }
        return obj;
    }
    parameterMap.define = function define(definition) {
        var args = {};
        var toHost = createObject(definition.toHost);
        if (definition.invertible) {
            args.map = toHost;
        }
        else if (definition.canonical) {
            args.toHost = args.fromHost = toHost;
        }
        else {
            args.toHost = toHost;
            args.fromHost = createObject(definition.fromHost);
        }
        parameterMap.addMapping(definition.type, args);
        if (definition.isComplexType)
            parameterMap.addComplexType(definition.type);
    };
    for (var id in initialDefs)
        parameterMap.define(initialDefs[id]);
    return parameterMap;
};
OSF.OUtil.setNamespace("DispIdHost", OSF.DDA);
OSF.DDA.DispIdHost.Methods = {
    InvokeMethod: "invokeMethod",
    AddEventHandler: "addEventHandler",
    RemoveEventHandler: "removeEventHandler",
    OpenDialog: "openDialog",
    CloseDialog: "closeDialog",
    MessageParent: "messageParent",
    SendMessage: "sendMessage"
};
OSF.DDA.DispIdHost.Delegates = {
    ExecuteAsync: "executeAsync",
    RegisterEventAsync: "registerEventAsync",
    UnregisterEventAsync: "unregisterEventAsync",
    ParameterMap: "parameterMap",
    OpenDialog: "openDialog",
    CloseDialog: "closeDialog",
    MessageParent: "messageParent",
    SendMessage: "sendMessage"
};
OSF.DDA.DispIdHost.Facade = function OSF_DDA_DispIdHost_Facade(getDelegateMethods, parameterMap) {
    var dispIdMap = {};
    var jsom = OSF.DDA.AsyncMethodNames;
    var did = OSF.DDA.MethodDispId;
    var methodMap = {
        "GoToByIdAsync": did.dispidNavigateToMethod,
        "GetSelectedDataAsync": did.dispidGetSelectedDataMethod,
        "SetSelectedDataAsync": did.dispidSetSelectedDataMethod,
        "GetDocumentCopyChunkAsync": did.dispidGetDocumentCopyChunkMethod,
        "ReleaseDocumentCopyAsync": did.dispidReleaseDocumentCopyMethod,
        "GetDocumentCopyAsync": did.dispidGetDocumentCopyMethod,
        "AddFromSelectionAsync": did.dispidAddBindingFromSelectionMethod,
        "AddFromPromptAsync": did.dispidAddBindingFromPromptMethod,
        "AddFromNamedItemAsync": did.dispidAddBindingFromNamedItemMethod,
        "GetAllAsync": did.dispidGetAllBindingsMethod,
        "GetByIdAsync": did.dispidGetBindingMethod,
        "ReleaseByIdAsync": did.dispidReleaseBindingMethod,
        "GetDataAsync": did.dispidGetBindingDataMethod,
        "SetDataAsync": did.dispidSetBindingDataMethod,
        "AddRowsAsync": did.dispidAddRowsMethod,
        "AddColumnsAsync": did.dispidAddColumnsMethod,
        "DeleteAllDataValuesAsync": did.dispidClearAllRowsMethod,
        "RefreshAsync": did.dispidLoadSettingsMethod,
        "SaveAsync": did.dispidSaveSettingsMethod,
        "GetActiveViewAsync": did.dispidGetActiveViewMethod,
        "GetFilePropertiesAsync": did.dispidGetFilePropertiesMethod,
        "GetOfficeThemeAsync": did.dispidGetOfficeThemeMethod,
        "GetDocumentThemeAsync": did.dispidGetDocumentThemeMethod,
        "ClearFormatsAsync": did.dispidClearFormatsMethod,
        "SetTableOptionsAsync": did.dispidSetTableOptionsMethod,
        "SetFormatsAsync": did.dispidSetFormatsMethod,
        "GetAccessTokenAsync": did.dispidGetAccessTokenMethod,
        "GetAuthContextAsync": did.dispidGetAuthContextMethod,
        "ExecuteRichApiRequestAsync": did.dispidExecuteRichApiRequestMethod,
        "AppCommandInvocationCompletedAsync": did.dispidAppCommandInvocationCompletedMethod,
        "CloseContainerAsync": did.dispidCloseContainerMethod,
        "OpenBrowserWindow": did.dispidOpenBrowserWindow,
        "CreateDocumentAsync": did.dispidCreateDocumentMethod,
        "InsertFormAsync": did.dispidInsertFormMethod,
        "ExecuteFeature": did.dispidExecuteFeature,
        "QueryFeature": did.dispidQueryFeature,
        "AddDataPartAsync": did.dispidAddDataPartMethod,
        "GetDataPartByIdAsync": did.dispidGetDataPartByIdMethod,
        "GetDataPartsByNameSpaceAsync": did.dispidGetDataPartsByNamespaceMethod,
        "GetPartXmlAsync": did.dispidGetDataPartXmlMethod,
        "GetPartNodesAsync": did.dispidGetDataPartNodesMethod,
        "DeleteDataPartAsync": did.dispidDeleteDataPartMethod,
        "GetNodeValueAsync": did.dispidGetDataNodeValueMethod,
        "GetNodeXmlAsync": did.dispidGetDataNodeXmlMethod,
        "GetRelativeNodesAsync": did.dispidGetDataNodesMethod,
        "SetNodeValueAsync": did.dispidSetDataNodeValueMethod,
        "SetNodeXmlAsync": did.dispidSetDataNodeXmlMethod,
        "AddDataPartNamespaceAsync": did.dispidAddDataNamespaceMethod,
        "GetDataPartNamespaceAsync": did.dispidGetDataUriByPrefixMethod,
        "GetDataPartPrefixAsync": did.dispidGetDataPrefixByUriMethod,
        "GetNodeTextAsync": did.dispidGetDataNodeTextMethod,
        "SetNodeTextAsync": did.dispidSetDataNodeTextMethod,
        "GetSelectedTask": did.dispidGetSelectedTaskMethod,
        "GetTask": did.dispidGetTaskMethod,
        "GetWSSUrl": did.dispidGetWSSUrlMethod,
        "GetTaskField": did.dispidGetTaskFieldMethod,
        "GetSelectedResource": did.dispidGetSelectedResourceMethod,
        "GetResourceField": did.dispidGetResourceFieldMethod,
        "GetProjectField": did.dispidGetProjectFieldMethod,
        "GetSelectedView": did.dispidGetSelectedViewMethod,
        "GetTaskByIndex": did.dispidGetTaskByIndexMethod,
        "GetResourceByIndex": did.dispidGetResourceByIndexMethod,
        "SetTaskField": did.dispidSetTaskFieldMethod,
        "SetResourceField": did.dispidSetResourceFieldMethod,
        "GetMaxTaskIndex": did.dispidGetMaxTaskIndexMethod,
        "GetMaxResourceIndex": did.dispidGetMaxResourceIndexMethod,
        "CreateTask": did.dispidCreateTaskMethod
    };
    for (var method in methodMap) {
        if (jsom[method]) {
            dispIdMap[jsom[method].id] = methodMap[method];
        }
    }
    jsom = OSF.DDA.SyncMethodNames;
    did = OSF.DDA.MethodDispId;
    var syncMethodMap = {
        "MessageParent": did.dispidMessageParentMethod,
        "SendMessage": did.dispidSendMessageMethod
    };
    for (var method in syncMethodMap) {
        if (jsom[method]) {
            dispIdMap[jsom[method].id] = syncMethodMap[method];
        }
    }
    jsom = Microsoft.Office.WebExtension.EventType;
    did = OSF.DDA.EventDispId;
    var eventMap = {
        "SettingsChanged": did.dispidSettingsChangedEvent,
        "DocumentSelectionChanged": did.dispidDocumentSelectionChangedEvent,
        "BindingSelectionChanged": did.dispidBindingSelectionChangedEvent,
        "BindingDataChanged": did.dispidBindingDataChangedEvent,
        "ActiveViewChanged": did.dispidActiveViewChangedEvent,
        "OfficeThemeChanged": did.dispidOfficeThemeChangedEvent,
        "DocumentThemeChanged": did.dispidDocumentThemeChangedEvent,
        "AppCommandInvoked": did.dispidAppCommandInvokedEvent,
        "DialogMessageReceived": did.dispidDialogMessageReceivedEvent,
        "DialogParentMessageReceived": did.dispidDialogParentMessageReceivedEvent,
        "ObjectDeleted": did.dispidObjectDeletedEvent,
        "ObjectSelectionChanged": did.dispidObjectSelectionChangedEvent,
        "ObjectDataChanged": did.dispidObjectDataChangedEvent,
        "ContentControlAdded": did.dispidContentControlAddedEvent,
        "RichApiMessage": did.dispidRichApiMessageEvent,
        "ItemChanged": did.dispidOlkItemSelectedChangedEvent,
        "RecipientsChanged": did.dispidOlkRecipientsChangedEvent,
        "AppointmentTimeChanged": did.dispidOlkAppointmentTimeChangedEvent,
        "RecurrenceChanged": did.dispidOlkRecurrenceChangedEvent,
        "AttachmentsChanged": did.dispidOlkAttachmentsChangedEvent,
        "EnhancedLocationsChanged": did.dispidOlkEnhancedLocationsChangedEvent,
        "InfobarClicked": did.dispidOlkInfobarClickedEvent,
        "TaskSelectionChanged": did.dispidTaskSelectionChangedEvent,
        "ResourceSelectionChanged": did.dispidResourceSelectionChangedEvent,
        "ViewSelectionChanged": did.dispidViewSelectionChangedEvent,
        "DataNodeInserted": did.dispidDataNodeAddedEvent,
        "DataNodeReplaced": did.dispidDataNodeReplacedEvent,
        "DataNodeDeleted": did.dispidDataNodeDeletedEvent
    };
    for (var event in eventMap) {
        if (jsom[event]) {
            dispIdMap[jsom[event]] = eventMap[event];
        }
    }
    function IsObjectEvent(dispId) {
        return (dispId == OSF.DDA.EventDispId.dispidObjectDeletedEvent ||
            dispId == OSF.DDA.EventDispId.dispidObjectSelectionChangedEvent ||
            dispId == OSF.DDA.EventDispId.dispidObjectDataChangedEvent ||
            dispId == OSF.DDA.EventDispId.dispidContentControlAddedEvent);
    }
    function onException(ex, asyncMethodCall, suppliedArgs, callArgs) {
        if (typeof ex == "number") {
            if (!callArgs) {
                callArgs = asyncMethodCall.getCallArgs(suppliedArgs);
            }
            OSF.DDA.issueAsyncResult(callArgs, ex, OSF.DDA.ErrorCodeManager.getErrorArgs(ex));
        }
        else {
            throw ex;
        }
    }
    ;
    this[OSF.DDA.DispIdHost.Methods.InvokeMethod] = function OSF_DDA_DispIdHost_Facade$InvokeMethod(method, suppliedArguments, caller, privateState) {
        var callArgs;
        try {
            var methodName = method.id;
            var asyncMethodCall = OSF.DDA.AsyncMethodCalls[methodName];
            callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, privateState);
            var dispId = dispIdMap[methodName];
            var delegate = getDelegateMethods(methodName);
            var richApiInExcelMethodSubstitution = null;
            if (window.Excel && window.Office.context.requirements.isSetSupported("RedirectV1Api")) {
                window.Excel._RedirectV1APIs = true;
            }
            if (window.Excel && window.Excel._RedirectV1APIs && (richApiInExcelMethodSubstitution = window.Excel._V1APIMap[methodName])) {
                var preprocessedCallArgs = OSF.OUtil.shallowCopy(callArgs);
                delete preprocessedCallArgs[Microsoft.Office.WebExtension.Parameters.AsyncContext];
                if (richApiInExcelMethodSubstitution.preprocess) {
                    preprocessedCallArgs = richApiInExcelMethodSubstitution.preprocess(preprocessedCallArgs);
                }
                var ctx = new window.Excel.RequestContext();
                var result = richApiInExcelMethodSubstitution.call(ctx, preprocessedCallArgs);
                ctx.sync()
                    .then(function () {
                    var response = result.value;
                    var status = response.status;
                    delete response["status"];
                    delete response["@odata.type"];
                    if (richApiInExcelMethodSubstitution.postprocess) {
                        response = richApiInExcelMethodSubstitution.postprocess(response, preprocessedCallArgs);
                    }
                    if (status != 0) {
                        response = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
                    }
                    OSF.DDA.issueAsyncResult(callArgs, status, response);
                })["catch"](function (error) {
                    OSF.DDA.issueAsyncResult(callArgs, OSF.DDA.ErrorCodeManager.errorCodes.ooeFailure, null);
                });
            }
            else {
                var hostCallArgs;
                if (parameterMap.toHost) {
                    hostCallArgs = parameterMap.toHost(dispId, callArgs);
                }
                else {
                    hostCallArgs = callArgs;
                }
                var startTime = (new Date()).getTime();
                delegate[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]({
                    "dispId": dispId,
                    "hostCallArgs": hostCallArgs,
                    "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { },
                    "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { },
                    "onComplete": function (status, hostResponseArgs) {
                        var responseArgs;
                        if (status == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                            if (parameterMap.fromHost) {
                                responseArgs = parameterMap.fromHost(dispId, hostResponseArgs);
                            }
                            else {
                                responseArgs = hostResponseArgs;
                            }
                        }
                        else {
                            responseArgs = hostResponseArgs;
                        }
                        var payload = asyncMethodCall.processResponse(status, responseArgs, caller, callArgs);
                        OSF.DDA.issueAsyncResult(callArgs, status, payload);
                        if (OSF.AppTelemetry && !(OSF.ConstantNames && OSF.ConstantNames.IsCustomFunctionsRuntime)) {
                            OSF.AppTelemetry.onMethodDone(dispId, hostCallArgs, Math.abs((new Date()).getTime() - startTime), status);
                        }
                    }
                });
            }
        }
        catch (ex) {
            onException(ex, asyncMethodCall, suppliedArguments, callArgs);
        }
    };
    this[OSF.DDA.DispIdHost.Methods.AddEventHandler] = function OSF_DDA_DispIdHost_Facade$AddEventHandler(suppliedArguments, eventDispatch, caller, isPopupWindow) {
        var callArgs;
        var eventType, handler;
        var isObjectEvent = false;
        function onEnsureRegistration(status) {
            if (status == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                var added = !isObjectEvent ? eventDispatch.addEventHandler(eventType, handler) :
                    eventDispatch.addObjectEventHandler(eventType, callArgs[Microsoft.Office.WebExtension.Parameters.Id], handler);
                if (!added) {
                    status = OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerAdditionFailed;
                }
            }
            var error;
            if (status != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                error = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
            }
            OSF.DDA.issueAsyncResult(callArgs, status, error);
        }
        try {
            var asyncMethodCall = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.AddHandlerAsync.id];
            callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
            eventType = callArgs[Microsoft.Office.WebExtension.Parameters.EventType];
            handler = callArgs[Microsoft.Office.WebExtension.Parameters.Handler];
            if (isPopupWindow) {
                onEnsureRegistration(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
                return;
            }
            var dispId_1 = dispIdMap[eventType];
            isObjectEvent = IsObjectEvent(dispId_1);
            var targetId_1 = (isObjectEvent ? callArgs[Microsoft.Office.WebExtension.Parameters.Id] : (caller.id || ""));
            var count = isObjectEvent ? eventDispatch.getObjectEventHandlerCount(eventType, targetId_1) : eventDispatch.getEventHandlerCount(eventType);
            if (count == 0) {
                var invoker = getDelegateMethods(eventType)[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync];
                invoker({
                    "eventType": eventType,
                    "dispId": dispId_1,
                    "targetId": targetId_1,
                    "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
                    "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
                    "onComplete": onEnsureRegistration,
                    "onEvent": function handleEvent(hostArgs) {
                        var args = parameterMap.fromHost(dispId_1, hostArgs);
                        if (!isObjectEvent)
                            eventDispatch.fireEvent(OSF.DDA.OMFactory.manufactureEventArgs(eventType, caller, args));
                        else
                            eventDispatch.fireObjectEvent(targetId_1, OSF.DDA.OMFactory.manufactureEventArgs(eventType, targetId_1, args));
                    }
                });
            }
            else {
                onEnsureRegistration(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
            }
        }
        catch (ex) {
            onException(ex, asyncMethodCall, suppliedArguments, callArgs);
        }
    };
    this[OSF.DDA.DispIdHost.Methods.RemoveEventHandler] = function OSF_DDA_DispIdHost_Facade$RemoveEventHandler(suppliedArguments, eventDispatch, caller) {
        var callArgs;
        var eventType, handler;
        var isObjectEvent = false;
        function onEnsureRegistration(status) {
            var error;
            if (status != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                error = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
            }
            OSF.DDA.issueAsyncResult(callArgs, status, error);
        }
        try {
            var asyncMethodCall = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.id];
            callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
            eventType = callArgs[Microsoft.Office.WebExtension.Parameters.EventType];
            handler = callArgs[Microsoft.Office.WebExtension.Parameters.Handler];
            var dispId = dispIdMap[eventType];
            isObjectEvent = IsObjectEvent(dispId);
            var targetId = (isObjectEvent ? callArgs[Microsoft.Office.WebExtension.Parameters.Id] : (caller.id || ""));
            var status, removeSuccess;
            if (handler === null) {
                removeSuccess = isObjectEvent ? eventDispatch.clearObjectEventHandlers(eventType, targetId) : eventDispatch.clearEventHandlers(eventType);
                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess;
            }
            else {
                removeSuccess = isObjectEvent ? eventDispatch.removeObjectEventHandler(eventType, targetId, handler) : eventDispatch.removeEventHandler(eventType, handler);
                status = removeSuccess ? OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess : OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist;
            }
            var count = isObjectEvent ? eventDispatch.getObjectEventHandlerCount(eventType, targetId) : eventDispatch.getEventHandlerCount(eventType);
            if (removeSuccess && count == 0) {
                var invoker = getDelegateMethods(eventType)[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync];
                invoker({
                    "eventType": eventType,
                    "dispId": dispId,
                    "targetId": targetId,
                    "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
                    "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
                    "onComplete": onEnsureRegistration
                });
            }
            else {
                onEnsureRegistration(status);
            }
        }
        catch (ex) {
            onException(ex, asyncMethodCall, suppliedArguments, callArgs);
        }
    };
    this[OSF.DDA.DispIdHost.Methods.OpenDialog] = function OSF_DDA_DispIdHost_Facade$OpenDialog(suppliedArguments, eventDispatch, caller) {
        var callArgs;
        var targetId;
        var dialogMessageEvent = Microsoft.Office.WebExtension.EventType.DialogMessageReceived;
        var dialogOtherEvent = Microsoft.Office.WebExtension.EventType.DialogEventReceived;
        function onEnsureRegistration(status) {
            var payload;
            if (status != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                payload = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
            }
            else {
                var onSucceedArgs = {};
                onSucceedArgs[Microsoft.Office.WebExtension.Parameters.Id] = targetId;
                onSucceedArgs[Microsoft.Office.WebExtension.Parameters.Data] = eventDispatch;
                var payload = asyncMethodCall.processResponse(status, onSucceedArgs, caller, callArgs);
                OSF.DialogShownStatus.hasDialogShown = true;
                eventDispatch.clearEventHandlers(dialogMessageEvent);
                eventDispatch.clearEventHandlers(dialogOtherEvent);
            }
            OSF.DDA.issueAsyncResult(callArgs, status, payload);
        }
        try {
            if (dialogMessageEvent == undefined || dialogOtherEvent == undefined) {
                onEnsureRegistration(OSF.DDA.ErrorCodeManager.ooeOperationNotSupported);
            }
            if (OSF.DDA.AsyncMethodNames.DisplayDialogAsync == null) {
                onEnsureRegistration(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
                return;
            }
            var asyncMethodCall = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.DisplayDialogAsync.id];
            callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
            var dispId = dispIdMap[dialogMessageEvent];
            var delegateMethods = getDelegateMethods(dialogMessageEvent);
            var invoker = delegateMethods[OSF.DDA.DispIdHost.Delegates.OpenDialog] != undefined ?
                delegateMethods[OSF.DDA.DispIdHost.Delegates.OpenDialog] :
                delegateMethods[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync];
            targetId = JSON.stringify(callArgs);
            if (!OSF.DialogShownStatus.hasDialogShown) {
                eventDispatch.clearQueuedEvent(dialogMessageEvent);
                eventDispatch.clearQueuedEvent(dialogOtherEvent);
                eventDispatch.clearQueuedEvent(Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived);
            }
            invoker({
                "eventType": dialogMessageEvent,
                "dispId": dispId,
                "targetId": targetId,
                "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
                "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
                "onComplete": onEnsureRegistration,
                "onEvent": function handleEvent(hostArgs) {
                    var args = parameterMap.fromHost(dispId, hostArgs);
                    var event = OSF.DDA.OMFactory.manufactureEventArgs(dialogMessageEvent, caller, args);
                    if (event.type == dialogOtherEvent) {
                        var payload = OSF.DDA.ErrorCodeManager.getErrorArgs(event.error);
                        var errorArgs = {};
                        errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code] = status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
                        errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name] = payload.name || payload;
                        errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message] = payload.message || payload;
                        event.error = new OSF.DDA.Error(errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code]);
                    }
                    eventDispatch.fireOrQueueEvent(event);
                    if (args[OSF.DDA.PropertyDescriptors.MessageType] == OSF.DialogMessageType.DialogClosed) {
                        eventDispatch.clearEventHandlers(dialogMessageEvent);
                        eventDispatch.clearEventHandlers(dialogOtherEvent);
                        eventDispatch.clearEventHandlers(Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived);
                        OSF.DialogShownStatus.hasDialogShown = false;
                    }
                }
            });
        }
        catch (ex) {
            onException(ex, asyncMethodCall, suppliedArguments, callArgs);
        }
    };
    this[OSF.DDA.DispIdHost.Methods.CloseDialog] = function OSF_DDA_DispIdHost_Facade$CloseDialog(suppliedArguments, targetId, eventDispatch, caller) {
        var callArgs;
        var dialogMessageEvent, dialogOtherEvent;
        var closeStatus = OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess;
        function closeCallback(status) {
            closeStatus = status;
            OSF.DialogShownStatus.hasDialogShown = false;
        }
        try {
            var asyncMethodCall = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.CloseAsync.id];
            callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
            dialogMessageEvent = Microsoft.Office.WebExtension.EventType.DialogMessageReceived;
            dialogOtherEvent = Microsoft.Office.WebExtension.EventType.DialogEventReceived;
            eventDispatch.clearEventHandlers(dialogMessageEvent);
            eventDispatch.clearEventHandlers(dialogOtherEvent);
            var dispId = dispIdMap[dialogMessageEvent];
            var delegateMethods = getDelegateMethods(dialogMessageEvent);
            var invoker = delegateMethods[OSF.DDA.DispIdHost.Delegates.CloseDialog] != undefined ?
                delegateMethods[OSF.DDA.DispIdHost.Delegates.CloseDialog] :
                delegateMethods[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync];
            invoker({
                "eventType": dialogMessageEvent,
                "dispId": dispId,
                "targetId": targetId,
                "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
                "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
                "onComplete": closeCallback
            });
        }
        catch (ex) {
            onException(ex, asyncMethodCall, suppliedArguments, callArgs);
        }
        if (closeStatus != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
            throw OSF.OUtil.formatString(Strings.OfficeOM.L_FunctionCallFailed, OSF.DDA.AsyncMethodNames.CloseAsync.displayName, closeStatus);
        }
    };
    this[OSF.DDA.DispIdHost.Methods.MessageParent] = function OSF_DDA_DispIdHost_Facade$MessageParent(suppliedArguments, caller) {
        var stateInfo = {};
        var syncMethodCall = OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.MessageParent.id];
        var callArgs = syncMethodCall.verifyAndExtractCall(suppliedArguments, caller, stateInfo);
        var delegate = getDelegateMethods(OSF.DDA.SyncMethodNames.MessageParent.id);
        var invoker = delegate[OSF.DDA.DispIdHost.Delegates.MessageParent];
        var dispId = dispIdMap[OSF.DDA.SyncMethodNames.MessageParent.id];
        return invoker({
            "dispId": dispId,
            "hostCallArgs": callArgs,
            "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
            "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); }
        });
    };
    this[OSF.DDA.DispIdHost.Methods.SendMessage] = function OSF_DDA_DispIdHost_Facade$SendMessage(suppliedArguments, eventDispatch, caller) {
        var stateInfo = {};
        var syncMethodCall = OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.SendMessage.id];
        var callArgs = syncMethodCall.verifyAndExtractCall(suppliedArguments, caller, stateInfo);
        var delegate = getDelegateMethods(OSF.DDA.SyncMethodNames.SendMessage.id);
        var invoker = delegate[OSF.DDA.DispIdHost.Delegates.SendMessage];
        var dispId = dispIdMap[OSF.DDA.SyncMethodNames.SendMessage.id];
        return invoker({
            "dispId": dispId,
            "hostCallArgs": callArgs,
            "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
            "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); }
        });
    };
};
OSF.DDA.DispIdHost.addAsyncMethods = function OSF_DDA_DispIdHost$AddAsyncMethods(target, asyncMethodNames, privateState) {
    for (var entry in asyncMethodNames) {
        var method = asyncMethodNames[entry];
        var name = method.displayName;
        if (!target[name]) {
            OSF.OUtil.defineEnumerableProperty(target, name, {
                value: (function (asyncMethod) {
                    return function () {
                        var invokeMethod = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.InvokeMethod];
                        invokeMethod(asyncMethod, arguments, target, privateState);
                    };
                })(method)
            });
        }
    }
};
OSF.DDA.DispIdHost.addEventSupport = function OSF_DDA_DispIdHost$AddEventSupport(target, eventDispatch, isPopupWindow) {
    var add = OSF.DDA.AsyncMethodNames.AddHandlerAsync.displayName;
    var remove = OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.displayName;
    if (!target[add]) {
        OSF.OUtil.defineEnumerableProperty(target, add, {
            value: function () {
                var addEventHandler = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.AddEventHandler];
                addEventHandler(arguments, eventDispatch, target, isPopupWindow);
            }
        });
    }
    if (!target[remove]) {
        OSF.OUtil.defineEnumerableProperty(target, remove, {
            value: function () {
                var removeEventHandler = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.RemoveEventHandler];
                removeEventHandler(arguments, eventDispatch, target);
            }
        });
    }
};
OSF.ShowWindowDialogParameterKeys = {
    Url: "url",
    Width: "width",
    Height: "height",
    DisplayInIframe: "displayInIframe",
    HideTitle: "hideTitle",
    UseDeviceIndependentPixels: "useDeviceIndependentPixels",
    PromptBeforeOpen: "promptBeforeOpen",
    EnforceAppDomain: "enforceAppDomain",
    UrlNoHostInfo: "urlNoHostInfo"
};
OSF.HostThemeButtonStyleKeys = {
    ButtonBorderColor: "buttonBorderColor",
    ButtonBackgroundColor: "buttonBackgroundColor"
};
OSF.OmexPageParameterKeys = {
    AppName: "client",
    AppVersion: "cv",
    AppUILocale: "ui",
    AppDomain: "appDomain",
    StoreLocator: "rs",
    AssetId: "assetid",
    NotificationType: "notificationType",
    AppCorrelationId: "corr",
    AuthType: "authType",
    AppId: "appid",
    Scopes: "scopes"
};
OSF.AuthType = {
    Anonymous: 0,
    MSA: 1,
    OrgId: 2,
    ADAL: 3
};
OSF.OmexMessageKeys = {
    MessageType: "messageType",
    MessageValue: "messageValue"
};
OSF.OmexRemoveAddinMessageKeys = {
    RemoveAddinResultCode: "resultCode",
    RemoveAddinResultValue: "resultValue"
};
OSF.OmexRemoveAddinResultCode = {
    Success: 0,
    ClientError: 400,
    ServerError: 500,
    UnknownError: 600
};
(function (OfficeExt) {
    var WACUtils;
    (function (WACUtils) {
        var _trustedDomain = "^https:\/\/[a-z0-9-]+\.(officeapps\.live|officeapps-df\.live|partner\.officewebapps)\.com\/";
        function parseAppContextFromWindowName(skipSessionStorage, windowName) {
            return OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.AppContext);
        }
        WACUtils.parseAppContextFromWindowName = parseAppContextFromWindowName;
        function serializeObjectToString(obj) {
            if (typeof (JSON) !== "undefined") {
                try {
                    return JSON.stringify(obj);
                }
                catch (ex) {
                }
            }
            return "";
        }
        WACUtils.serializeObjectToString = serializeObjectToString;
        function isHostTrusted() {
            return (new RegExp(_trustedDomain)).test(OSF.getClientEndPoint()._targetUrl.toLowerCase());
        }
        WACUtils.isHostTrusted = isHostTrusted;
        function addHostInfoAsQueryParam(url, hostInfoValue) {
            if (!url) {
                return null;
            }
            url = url.trim() || '';
            var questionMark = "?";
            var hostInfo = "_host_Info=";
            var ampHostInfo = "&_host_Info=";
            var fragmentSeparator = "#";
            var urlParts = url.split(fragmentSeparator);
            var urlWithoutFragment = urlParts.shift();
            var fragment = urlParts.join(fragmentSeparator);
            var querySplits = urlWithoutFragment.split(questionMark);
            var urlWithoutFragmentWithHostInfo;
            if (querySplits.length > 1) {
                urlWithoutFragmentWithHostInfo = urlWithoutFragment + ampHostInfo + hostInfoValue;
            }
            else if (querySplits.length > 0) {
                urlWithoutFragmentWithHostInfo = urlWithoutFragment + questionMark + hostInfo + hostInfoValue;
            }
            if (fragment) {
                return [urlWithoutFragmentWithHostInfo, fragmentSeparator, fragment].join('');
            }
            else {
                return urlWithoutFragmentWithHostInfo;
            }
        }
        WACUtils.addHostInfoAsQueryParam = addHostInfoAsQueryParam;
        function getDomainForUrl(url) {
            if (!url) {
                return null;
            }
            var url_parser = document.createElement('a');
            url_parser.href = url;
            return url_parser.protocol + "//" + url_parser.host;
        }
        WACUtils.getDomainForUrl = getDomainForUrl;
        function shouldUseLocalStorageToPassMessage() {
            try {
                var osList = [
                    "Windows NT 6.1",
                    "Windows NT 6.2",
                    "Windows NT 6.3",
                    "Windows NT 10.0"
                ];
                var userAgent = window.navigator.userAgent;
                for (var i = 0, len = osList.length; i < len; i++) {
                    if (userAgent.indexOf(osList[i]) > -1) {
                        return isInternetExplorer();
                    }
                }
                return false;
            }
            catch (e) {
                logExceptionToBrowserConsole("Error happens in shouldUseLocalStorageToPassMessage.", e);
                return false;
            }
        }
        WACUtils.shouldUseLocalStorageToPassMessage = shouldUseLocalStorageToPassMessage;
        function isInternetExplorer() {
            try {
                var userAgent = window.navigator.userAgent;
                return userAgent.indexOf("MSIE ") > -1 || userAgent.indexOf("Trident/") > -1 || userAgent.indexOf("Edge/") > -1;
            }
            catch (e) {
                logExceptionToBrowserConsole("Error happens in isInternetExplorer.", e);
                return false;
            }
        }
        WACUtils.isInternetExplorer = isInternetExplorer;
        function removeMatchesFromLocalStorage(regexPatterns) {
            var _localStorage = OSF.OUtil.getLocalStorage();
            var keys = _localStorage.getKeysWithPrefix("");
            for (var i = 0, len = keys.length; i < len; i++) {
                var key = keys[i];
                for (var j = 0, lenRegex = regexPatterns.length; j < lenRegex; j++) {
                    if (regexPatterns[j].test(key)) {
                        _localStorage.removeItem(key);
                        break;
                    }
                }
            }
        }
        WACUtils.removeMatchesFromLocalStorage = removeMatchesFromLocalStorage;
        function logExceptionToBrowserConsole(message, exception) {
            OsfMsAjaxFactory.msAjaxDebug.trace(message + " Exception details: " + serializeObjectToString(exception));
        }
        WACUtils.logExceptionToBrowserConsole = logExceptionToBrowserConsole;
        function isTeamsWebView() {
            var ua = navigator.userAgent;
            return /Teams\/((?:(\d+)\.)?(?:(\d+)\.)?(?:(\d+)\.\d+)).* Electron\/((?:(\d+)\.)?(?:(\d+)\.)?(?:(\d+)\.\d+))/.test(ua);
        }
        WACUtils.isTeamsWebView = isTeamsWebView;
        function getHostSecure(url) {
            if (typeof url === "undefined" || !url) {
                return undefined;
            }
            var host = undefined;
            var httpsProtocol = "https:";
            try {
                var urlObj = new URL(url);
                if (urlObj) {
                    host = urlObj.host;
                }
                if (!urlObj.protocol) {
                    throw "fallback";
                }
                else if (urlObj.protocol !== httpsProtocol) {
                    return undefined;
                }
            }
            catch (ex) {
                try {
                    var parser = document.createElement("a");
                    parser.href = url;
                    if (parser.protocol !== httpsProtocol) {
                        return undefined;
                    }
                    var match = url.match(new RegExp("^https://[^/?#]+", "i"));
                    var naiveMatch = (match && match.length == 1) ? match[0].toLowerCase() : "";
                    var parsedUrlWithoutPort = (parser.protocol + "//" + parser.hostname).toLowerCase();
                    var parsedUrlWithPort = (parser.protocol + "//" + parser.host).toLowerCase();
                    if (parsedUrlWithPort === naiveMatch || parsedUrlWithoutPort == naiveMatch) {
                        host = parser.port == "443" ? parser.hostname : parser.host;
                    }
                }
                catch (ex) {
                    return undefined;
                }
            }
            return host ? host.toLowerCase() : undefined;
        }
        WACUtils.getHostSecure = getHostSecure;
        var CacheConstants = (function () {
            function CacheConstants() {
            }
            CacheConstants.GatedCacheKeyPrefix = "__OSF_GATED_OMEX.";
            CacheConstants.AnonymousCacheKeyPrefix = "__OSF_ANONYMOUS_OMEX.";
            CacheConstants.UngatedCacheKeyPrefix = "__OSF_OMEX.";
            CacheConstants.ActivatedCacheKeyPrefix = "__OSF_RUNTIME_.Activated.";
            CacheConstants.AppinstallAuthenticated = "appinstall_authenticated.";
            CacheConstants.Entitlement = "entitle.";
            CacheConstants.AppState = "appState.";
            CacheConstants.AppDetails = "appDetails.";
            CacheConstants.AppInstallInfo = "appInstallInfo.";
            CacheConstants.AuthenticatedAppInstallInfoCacheKey = CacheConstants.GatedCacheKeyPrefix + CacheConstants.AppinstallAuthenticated + "{0}.{1}.{2}.{3}";
            CacheConstants.EntitlementsKey = CacheConstants.Entitlement + "{0}.{1}";
            CacheConstants.AppStateCacheKey = "{0}" + CacheConstants.AppState + "{1}.{2}";
            CacheConstants.AppDetailKey = "{0}" + CacheConstants.AppDetails + "{1}";
            CacheConstants.AppInstallInfoKey = "{0}" + CacheConstants.AppInstallInfo + "{1}.{2}";
            CacheConstants.ActivatedCacheKey = CacheConstants.ActivatedCacheKeyPrefix + "{0}.{1}.{2}";
            return CacheConstants;
        }());
        WACUtils.CacheConstants = CacheConstants;
    })(WACUtils = OfficeExt.WACUtils || (OfficeExt.WACUtils = {}));
})(OfficeExt || (OfficeExt = {}));
OSF.OUtil.setNamespace("Microsoft", window);
OSF.OUtil.setNamespace("Office", Microsoft);
OSF.OUtil.setNamespace("Common", Microsoft.Office);
Microsoft.Office.Common.InvokeType = {
    "async": 0,
    "sync": 1,
    "asyncRegisterEvent": 2,
    "asyncUnregisterEvent": 3,
    "syncRegisterEvent": 4,
    "syncUnregisterEvent": 5
};
OSF.SerializerVersion = {
    MsAjax: 0,
    Browser: 1
};
(function (OfficeExt) {
    function appSpecificCheckOriginFunction(allowed_domains, eventObj, origin, checkOriginFunction) {
        return false;
    }
    ;
    OfficeExt.appSpecificCheckOrigin = appSpecificCheckOriginFunction;
})(OfficeExt || (OfficeExt = {}));
Microsoft.Office.Common.InvokeType = { "async": 0,
    "sync": 1,
    "asyncRegisterEvent": 2,
    "asyncUnregisterEvent": 3,
    "syncRegisterEvent": 4,
    "syncUnregisterEvent": 5
};
Microsoft.Office.Common.InvokeResultCode = {
    "noError": 0,
    "errorInRequest": -1,
    "errorHandlingRequest": -2,
    "errorInResponse": -3,
    "errorHandlingResponse": -4,
    "errorHandlingRequestAccessDenied": -5,
    "errorHandlingMethodCallTimedout": -6
};
Microsoft.Office.Common.MessageType = { "request": 0,
    "response": 1
};
Microsoft.Office.Common.ActionType = { "invoke": 0,
    "registerEvent": 1,
    "unregisterEvent": 2 };
Microsoft.Office.Common.ResponseType = { "forCalling": 0,
    "forEventing": 1
};
Microsoft.Office.Common.HostTrustStatus = {
    "unknown": 0,
    "untrusted": 1,
    "nothttps": 2,
    "trusted": 3
};
Microsoft.Office.Common.MethodObject = function Microsoft_Office_Common_MethodObject(method, invokeType, blockingOthers) {
    this._method = method;
    this._invokeType = invokeType;
    this._blockingOthers = blockingOthers;
};
Microsoft.Office.Common.MethodObject.prototype = {
    getMethod: function Microsoft_Office_Common_MethodObject$getMethod() {
        return this._method;
    },
    getInvokeType: function Microsoft_Office_Common_MethodObject$getInvokeType() {
        return this._invokeType;
    },
    getBlockingFlag: function Microsoft_Office_Common_MethodObject$getBlockingFlag() {
        return this._blockingOthers;
    }
};
Microsoft.Office.Common.EventMethodObject = function Microsoft_Office_Common_EventMethodObject(registerMethodObject, unregisterMethodObject) {
    this._registerMethodObject = registerMethodObject;
    this._unregisterMethodObject = unregisterMethodObject;
};
Microsoft.Office.Common.EventMethodObject.prototype = {
    getRegisterMethodObject: function Microsoft_Office_Common_EventMethodObject$getRegisterMethodObject() {
        return this._registerMethodObject;
    },
    getUnregisterMethodObject: function Microsoft_Office_Common_EventMethodObject$getUnregisterMethodObject() {
        return this._unregisterMethodObject;
    }
};
Microsoft.Office.Common.ServiceEndPoint = function Microsoft_Office_Common_ServiceEndPoint(serviceEndPointId) {
    var e = Function._validateParams(arguments, [
        { name: "serviceEndPointId", type: String, mayBeNull: false }
    ]);
    if (e)
        throw e;
    this._methodObjectList = {};
    this._eventHandlerProxyList = {};
    this._Id = serviceEndPointId;
    this._conversations = {};
    this._policyManager = null;
    this._appDomains = {};
    this._onHandleRequestError = null;
};
Microsoft.Office.Common.ServiceEndPoint.prototype = {
    registerMethod: function Microsoft_Office_Common_ServiceEndPoint$registerMethod(methodName, method, invokeType, blockingOthers) {
        var e = Function._validateParams(arguments, [{ name: "methodName", type: String, mayBeNull: false },
            { name: "method", type: Function, mayBeNull: false },
            { name: "invokeType", type: Number, mayBeNull: false },
            { name: "blockingOthers", type: Boolean, mayBeNull: false }
        ]);
        if (e)
            throw e;
        if (invokeType !== Microsoft.Office.Common.InvokeType.async
            && invokeType !== Microsoft.Office.Common.InvokeType.sync) {
            throw OsfMsAjaxFactory.msAjaxError.argument("invokeType");
        }
        var methodObject = new Microsoft.Office.Common.MethodObject(method, invokeType, blockingOthers);
        this._methodObjectList[methodName] = methodObject;
    },
    unregisterMethod: function Microsoft_Office_Common_ServiceEndPoint$unregisterMethod(methodName) {
        var e = Function._validateParams(arguments, [
            { name: "methodName", type: String, mayBeNull: false }
        ]);
        if (e)
            throw e;
        delete this._methodObjectList[methodName];
    },
    registerEvent: function Microsoft_Office_Common_ServiceEndPoint$registerEvent(eventName, registerMethod, unregisterMethod) {
        var e = Function._validateParams(arguments, [{ name: "eventName", type: String, mayBeNull: false },
            { name: "registerMethod", type: Function, mayBeNull: false },
            { name: "unregisterMethod", type: Function, mayBeNull: false }
        ]);
        if (e)
            throw e;
        var methodObject = new Microsoft.Office.Common.EventMethodObject(new Microsoft.Office.Common.MethodObject(registerMethod, Microsoft.Office.Common.InvokeType.syncRegisterEvent, false), new Microsoft.Office.Common.MethodObject(unregisterMethod, Microsoft.Office.Common.InvokeType.syncUnregisterEvent, false));
        this._methodObjectList[eventName] = methodObject;
    },
    registerEventEx: function Microsoft_Office_Common_ServiceEndPoint$registerEventEx(eventName, registerMethod, registerMethodInvokeType, unregisterMethod, unregisterMethodInvokeType) {
        var e = Function._validateParams(arguments, [{ name: "eventName", type: String, mayBeNull: false },
            { name: "registerMethod", type: Function, mayBeNull: false },
            { name: "registerMethodInvokeType", type: Number, mayBeNull: false },
            { name: "unregisterMethod", type: Function, mayBeNull: false },
            { name: "unregisterMethodInvokeType", type: Number, mayBeNull: false }
        ]);
        if (e)
            throw e;
        var methodObject = new Microsoft.Office.Common.EventMethodObject(new Microsoft.Office.Common.MethodObject(registerMethod, registerMethodInvokeType, false), new Microsoft.Office.Common.MethodObject(unregisterMethod, unregisterMethodInvokeType, false));
        this._methodObjectList[eventName] = methodObject;
    },
    unregisterEvent: function (eventName) {
        var e = Function._validateParams(arguments, [
            { name: "eventName", type: String, mayBeNull: false }
        ]);
        if (e)
            throw e;
        this.unregisterMethod(eventName);
    },
    registerConversation: function Microsoft_Office_Common_ServiceEndPoint$registerConversation(conversationId, conversationUrl, appDomains, serializerVersion) {
        var e = Function._validateParams(arguments, [
            { name: "conversationId", type: String, mayBeNull: false },
            { name: "conversationUrl", type: String, mayBeNull: false, optional: true },
            { name: "appDomains", type: Object, mayBeNull: true, optional: true },
            { name: "serializerVersion", type: Number, mayBeNull: true, optional: true }
        ]);
        if (e)
            throw e;
        ;
        if (appDomains) {
            if (!(appDomains instanceof Array)) {
                throw OsfMsAjaxFactory.msAjaxError.argument("appDomains");
            }
            this._appDomains[conversationId] = appDomains;
        }
        this._conversations[conversationId] = { url: conversationUrl, serializerVersion: serializerVersion };
    },
    unregisterConversation: function Microsoft_Office_Common_ServiceEndPoint$unregisterConversation(conversationId) {
        var e = Function._validateParams(arguments, [
            { name: "conversationId", type: String, mayBeNull: false }
        ]);
        if (e)
            throw e;
        delete this._conversations[conversationId];
    },
    setPolicyManager: function Microsoft_Office_Common_ServiceEndPoint$setPolicyManager(policyManager) {
        var e = Function._validateParams(arguments, [
            { name: "policyManager", type: Object, mayBeNull: false }
        ]);
        if (e)
            throw e;
        if (!policyManager.checkPermission) {
            throw OsfMsAjaxFactory.msAjaxError.argument("policyManager");
        }
        this._policyManager = policyManager;
    },
    getPolicyManager: function Microsoft_Office_Common_ServiceEndPoint$getPolicyManager() {
        return this._policyManager;
    },
    dispose: function Microsoft_Office_Common_ServiceEndPoint$dispose() {
        this._methodObjectList = null;
        this._eventHandlerProxyList = null;
        this._Id = null;
        this._conversations = null;
        this._policyManager = null;
        this._appDomains = null;
        this._onHandleRequestError = null;
    }
};
Microsoft.Office.Common.ClientEndPoint = function Microsoft_Office_Common_ClientEndPoint(conversationId, targetWindow, targetUrl, serializerVersion) {
    var e = Function._validateParams(arguments, [
        { name: "conversationId", type: String, mayBeNull: false },
        { name: "targetWindow", mayBeNull: false },
        { name: "targetUrl", type: String, mayBeNull: false },
        { name: "serializerVersion", type: Number, mayBeNull: true, optional: true }
    ]);
    if (e)
        throw e;
    try {
        if (!targetWindow.postMessage) {
            throw OsfMsAjaxFactory.msAjaxError.argument("targetWindow");
        }
    }
    catch (ex) {
        if (!Object.prototype.hasOwnProperty.call(targetWindow, "postMessage")) {
            throw OsfMsAjaxFactory.msAjaxError.argument("targetWindow");
        }
    }
    this._conversationId = conversationId;
    this._targetWindow = targetWindow;
    this._targetUrl = targetUrl;
    this._callingIndex = 0;
    this._callbackList = {};
    this._eventHandlerList = {};
    if (serializerVersion != null) {
        this._serializerVersion = serializerVersion;
    }
    else {
        this._serializerVersion = OSF.SerializerVersion.Browser;
    }
    this._checkReceiverOriginAndRun = null;
    this._hostTrustCheckStatus = Microsoft.Office.Common.HostTrustStatus.unknown;
    this._checkStatusLogged = false;
};
Microsoft.Office.Common.ClientEndPoint.prototype = {
    invoke: function Microsoft_Office_Common_ClientEndPoint$invoke(targetMethodName, callback, param) {
        var e = Function._validateParams(arguments, [{ name: "targetMethodName", type: String, mayBeNull: false },
            { name: "callback", type: Function, mayBeNull: true },
            { name: "param", mayBeNull: true }
        ]);
        if (e)
            throw e;
        var me = this;
        var funcToRun = function () {
            var correlationId = me._callingIndex++;
            var now = new Date();
            var callbackEntry = { "callback": callback, "createdOn": now.getTime() };
            if (param && typeof param === "object" && typeof param.__timeout__ === "number") {
                callbackEntry.timeout = param.__timeout__;
                delete param.__timeout__;
            }
            me._callbackList[correlationId] = callbackEntry;
            try {
                if (me._hostTrustCheckStatus !== Microsoft.Office.Common.HostTrustStatus.trusted) {
                    if (targetMethodName !== "ContextActivationManager_getAppContextAsync") {
                        throw "Access Denied";
                    }
                }
                var callRequest = new Microsoft.Office.Common.Request(targetMethodName, Microsoft.Office.Common.ActionType.invoke, me._conversationId, correlationId, param);
                var msg = Microsoft.Office.Common.MessagePackager.envelope(callRequest, me._serializerVersion);
                me._targetWindow.postMessage(msg, me._targetUrl);
                Microsoft.Office.Common.XdmCommunicationManager._startMethodTimeoutTimer();
            }
            catch (ex) {
                try {
                    if (callback !== null)
                        callback(Microsoft.Office.Common.InvokeResultCode.errorInRequest, ex);
                }
                finally {
                    delete me._callbackList[correlationId];
                }
            }
        };
        if (me._checkReceiverOriginAndRun) {
            me._checkReceiverOriginAndRun(funcToRun);
        }
        else {
            me._hostTrustCheckStatus = Microsoft.Office.Common.HostTrustStatus.trusted;
            funcToRun();
        }
    },
    registerForEvent: function Microsoft_Office_Common_ClientEndPoint$registerForEvent(targetEventName, eventHandler, callback, data) {
        var e = Function._validateParams(arguments, [{ name: "targetEventName", type: String, mayBeNull: false },
            { name: "eventHandler", type: Function, mayBeNull: false },
            { name: "callback", type: Function, mayBeNull: true },
            { name: "data", mayBeNull: true, optional: true }
        ]);
        if (e)
            throw e;
        var correlationId = this._callingIndex++;
        var now = new Date();
        this._callbackList[correlationId] = { "callback": callback, "createdOn": now.getTime() };
        try {
            var callRequest = new Microsoft.Office.Common.Request(targetEventName, Microsoft.Office.Common.ActionType.registerEvent, this._conversationId, correlationId, data);
            var msg = Microsoft.Office.Common.MessagePackager.envelope(callRequest, this._serializerVersion);
            this._targetWindow.postMessage(msg, this._targetUrl);
            Microsoft.Office.Common.XdmCommunicationManager._startMethodTimeoutTimer();
            this._eventHandlerList[targetEventName] = eventHandler;
        }
        catch (ex) {
            try {
                if (callback !== null) {
                    callback(Microsoft.Office.Common.InvokeResultCode.errorInRequest, ex);
                }
            }
            finally {
                delete this._callbackList[correlationId];
            }
        }
    },
    unregisterForEvent: function Microsoft_Office_Common_ClientEndPoint$unregisterForEvent(targetEventName, callback, data) {
        var e = Function._validateParams(arguments, [{ name: "targetEventName", type: String, mayBeNull: false },
            { name: "callback", type: Function, mayBeNull: true },
            { name: "data", mayBeNull: true, optional: true }
        ]);
        if (e)
            throw e;
        var correlationId = this._callingIndex++;
        var now = new Date();
        this._callbackList[correlationId] = { "callback": callback, "createdOn": now.getTime() };
        try {
            var callRequest = new Microsoft.Office.Common.Request(targetEventName, Microsoft.Office.Common.ActionType.unregisterEvent, this._conversationId, correlationId, data);
            var msg = Microsoft.Office.Common.MessagePackager.envelope(callRequest, this._serializerVersion);
            this._targetWindow.postMessage(msg, this._targetUrl);
            Microsoft.Office.Common.XdmCommunicationManager._startMethodTimeoutTimer();
        }
        catch (ex) {
            try {
                if (callback !== null) {
                    callback(Microsoft.Office.Common.InvokeResultCode.errorInRequest, ex);
                }
            }
            finally {
                delete this._callbackList[correlationId];
            }
        }
        finally {
            delete this._eventHandlerList[targetEventName];
        }
    }
};
Microsoft.Office.Common.XdmCommunicationManager = (function () {
    var _invokerQueue = [];
    var _lastMessageProcessTime = null;
    var _messageProcessingTimer = null;
    var _processInterval = 10;
    var _blockingFlag = false;
    var _methodTimeoutTimer = null;
    var _methodTimeoutProcessInterval = 2000;
    var _methodTimeoutDefault = 65000;
    var _methodTimeout = _methodTimeoutDefault;
    var _serviceEndPoints = {};
    var _clientEndPoints = {};
    var _initialized = false;
    function _lookupServiceEndPoint(conversationId) {
        for (var id in _serviceEndPoints) {
            if (_serviceEndPoints[id]._conversations[conversationId]) {
                return _serviceEndPoints[id];
            }
        }
        OsfMsAjaxFactory.msAjaxDebug.trace("Unknown conversation Id.");
        throw OsfMsAjaxFactory.msAjaxError.argument("conversationId");
    }
    ;
    function _lookupClientEndPoint(conversationId) {
        var clientEndPoint = _clientEndPoints[conversationId];
        if (!clientEndPoint) {
            OsfMsAjaxFactory.msAjaxDebug.trace("Unknown conversation Id.");
        }
        return clientEndPoint;
    }
    ;
    function _lookupMethodObject(serviceEndPoint, messageObject) {
        var methodOrEventMethodObject = serviceEndPoint._methodObjectList[messageObject._actionName];
        if (!methodOrEventMethodObject) {
            OsfMsAjaxFactory.msAjaxDebug.trace("The specified method is not registered on service endpoint:" + messageObject._actionName);
            throw OsfMsAjaxFactory.msAjaxError.argument("messageObject");
        }
        var methodObject = null;
        if (messageObject._actionType === Microsoft.Office.Common.ActionType.invoke) {
            methodObject = methodOrEventMethodObject;
        }
        else if (messageObject._actionType === Microsoft.Office.Common.ActionType.registerEvent) {
            methodObject = methodOrEventMethodObject.getRegisterMethodObject();
        }
        else {
            methodObject = methodOrEventMethodObject.getUnregisterMethodObject();
        }
        return methodObject;
    }
    ;
    function _enqueInvoker(invoker) {
        _invokerQueue.push(invoker);
    }
    ;
    function _dequeInvoker() {
        if (_messageProcessingTimer !== null) {
            if (!_blockingFlag) {
                if (_invokerQueue.length > 0) {
                    var invoker = _invokerQueue.shift();
                    _executeCommand(invoker);
                }
                else {
                    clearInterval(_messageProcessingTimer);
                    _messageProcessingTimer = null;
                }
            }
        }
        else {
            OsfMsAjaxFactory.msAjaxDebug.trace("channel is not ready.");
        }
    }
    ;
    function _executeCommand(invoker) {
        _blockingFlag = invoker.getInvokeBlockingFlag();
        invoker.invoke();
        _lastMessageProcessTime = (new Date()).getTime();
    }
    ;
    function _checkMethodTimeout() {
        if (_methodTimeoutTimer) {
            var clientEndPoint;
            var methodCallsNotTimedout = 0;
            var now = new Date();
            var timeoutValue;
            for (var conversationId in _clientEndPoints) {
                clientEndPoint = _clientEndPoints[conversationId];
                for (var correlationId in clientEndPoint._callbackList) {
                    var callbackEntry = clientEndPoint._callbackList[correlationId];
                    timeoutValue = callbackEntry.timeout ? callbackEntry.timeout : _methodTimeout;
                    if (timeoutValue >= 0 && Math.abs(now.getTime() - callbackEntry.createdOn) >= timeoutValue) {
                        try {
                            if (callbackEntry.callback) {
                                callbackEntry.callback(Microsoft.Office.Common.InvokeResultCode.errorHandlingMethodCallTimedout, null);
                            }
                        }
                        finally {
                            delete clientEndPoint._callbackList[correlationId];
                        }
                    }
                    else {
                        methodCallsNotTimedout++;
                    }
                    ;
                }
            }
            if (methodCallsNotTimedout === 0) {
                clearInterval(_methodTimeoutTimer);
                _methodTimeoutTimer = null;
            }
        }
        else {
            OsfMsAjaxFactory.msAjaxDebug.trace("channel is not ready.");
        }
    }
    ;
    function _postCallbackHandler() {
        _blockingFlag = false;
    }
    ;
    function _registerListener(listener) {
        if (window.addEventListener) {
            window.addEventListener("message", listener, false);
        }
        else if ((navigator.userAgent.indexOf("MSIE") > -1) && window.attachEvent) {
            window.attachEvent("onmessage", listener);
        }
        else {
            OsfMsAjaxFactory.msAjaxDebug.trace("Browser doesn't support the required API.");
            throw OsfMsAjaxFactory.msAjaxError.argument("Browser");
        }
    }
    ;
    function _checkOrigin(url, origin) {
        var res = false;
        if (!url || !origin || url === "null" || origin === "null" || !url.length || !origin.length) {
            return res;
        }
        if (OSF.OUtil.checkFlight(OSF.FlightNames.AddinEnforceHttps)) {
            res = _urlCompareUsingUrlStrings(url, origin);
        }
        else {
            var url_parser, org_parser;
            url_parser = document.createElement('a');
            org_parser = document.createElement('a');
            url_parser.href = url;
            org_parser.href = origin;
            res = _urlCompare(url_parser, org_parser);
        }
        return res;
    }
    function _checkOriginWithAppDomains(allowed_domains, origin) {
        var res = false;
        if (!origin || origin === "null" || !origin.length || !(allowed_domains) || !(allowed_domains instanceof Array) || !allowed_domains.length) {
            return res;
        }
        var org_parser = document.createElement('a');
        var app_domain_parser = document.createElement('a');
        org_parser.href = origin;
        for (var i = 0; i < allowed_domains.length && !res; i++) {
            if (allowed_domains[i].indexOf("://") !== -1) {
                if (OSF.OUtil.checkFlight(OSF.FlightNames.AddinEnforceHttps)) {
                    res = _urlCompareUsingUrlStrings(origin, allowed_domains[i]);
                }
                else {
                    app_domain_parser.href = allowed_domains[i];
                    res = _urlCompare(org_parser, app_domain_parser);
                }
            }
        }
        return res;
    }
    function _isHostNameValidWacDomain(hostName) {
        if (!hostName || hostName === "null") {
            return false;
        }
        var regexHostNameStringArray = new Array("^office-int\\.com$", "^officeapps\\.live-int\\.com$", "^.*\\.dod\\.online\\.office365\\.us$", "^.*\\.gov\\.online\\.office365\\.us$", "^.*\\.officeapps\\.live\\.com$", "^.*\\.officeapps\\.live-int\\.com$", "^.*\\.officeapps-df\\.live\\.com$", "^.*\\.online\\.office\\.de$", "^.*\\.partner\\.officewebapps\\.cn$", "^.*\\.office\\.net$", "^" + document.domain.replace(new RegExp("\\.", "g"), "\\.") + "$");
        var regexHostName = new RegExp(regexHostNameStringArray.join("|"));
        return regexHostName.test(hostName);
    }
    function _isTargetSubdomainOfSourceLocation(sourceLocation, messageOrigin) {
        if (!sourceLocation || !messageOrigin || sourceLocation === "null" || messageOrigin === "null") {
            return false;
        }
        var sourceLocationParser;
        var messageOriginParser;
        if (OSF.OUtil.checkFlight(OSF.FlightNames.AddinEnforceHttps)) {
            sourceLocationParser = OSF.OUtil.parseUrl(sourceLocation, true);
            messageOriginParser = OSF.OUtil.parseUrl(messageOrigin, true);
        }
        else {
            sourceLocationParser = document.createElement('a');
            sourceLocationParser.href = sourceLocation;
            messageOriginParser = document.createElement('a');
            messageOriginParser.href = messageOrigin;
        }
        var isSameProtocol = sourceLocationParser.protocol === messageOriginParser.protocol;
        var isSamePort = sourceLocationParser.port === messageOriginParser.port;
        var originHostName = messageOriginParser.hostname;
        var sourceLocationHostName = sourceLocationParser.hostname;
        var isSameDomain = originHostName === sourceLocationHostName;
        var isSubDomain = false;
        if (!isSameDomain && originHostName.length > sourceLocationHostName.length + 1) {
            isSubDomain = originHostName.slice(-(sourceLocationHostName.length + 1)) === '.' + sourceLocationHostName;
        }
        var isSameDomainOrSubdomain = isSameDomain || isSubDomain;
        return isSamePort && isSameProtocol && isSameDomainOrSubdomain;
    }
    function _urlCompareUsingUrlStrings(url_str1, url_str2) {
        var parsedUrl1 = OSF.OUtil.parseUrl(url_str1, true);
        var parsedUrl2 = OSF.OUtil.parseUrl(url_str2, true);
        return _urlCompare(parsedUrl1, parsedUrl2);
    }
    function _urlCompare(url_parser1, url_parser2) {
        return ((url_parser1.hostname == url_parser2.hostname) &&
            (url_parser1.protocol == url_parser2.protocol) &&
            _hasSamePort(url_parser1, url_parser2));
    }
    function _hasSamePort(url_parser1, url_parser2) {
        var httpPort = "80";
        var httpsPort = "443";
        return ((url_parser1.port == url_parser2.port) ||
            (url_parser1.port == "" && url_parser1.protocol == "http:" && url_parser2.port == httpPort) ||
            (url_parser1.port == "" && url_parser1.protocol == "https:" && url_parser2.port == httpsPort) ||
            (url_parser2.port == "" && url_parser2.protocol == "http:" && url_parser1.port == httpPort) ||
            (url_parser2.port == "" && url_parser2.protocol == "https:" && url_parser1.port == httpsPort));
    }
    function _receive(e) {
        if (!OSF) {
            return;
        }
        if (e.data != '') {
            var messageObject;
            var serializerVersion = OSF.SerializerVersion.Browser;
            var serializedMessage = e.data;
            try {
                messageObject = Microsoft.Office.Common.MessagePackager.unenvelope(serializedMessage, OSF.SerializerVersion.Browser);
                serializerVersion = messageObject._serializerVersion != null ? messageObject._serializerVersion : serializerVersion;
            }
            catch (ex) {
                return;
            }
            if (messageObject._messageType === Microsoft.Office.Common.MessageType.request) {
                var requesterUrl = (e.origin == null || e.origin === "null") ? messageObject._origin : e.origin;
                try {
                    var serviceEndPoint = _lookupServiceEndPoint(messageObject._conversationId);
                    ;
                    var conversation = serviceEndPoint._conversations[messageObject._conversationId];
                    serializerVersion = conversation.serializerVersion != null ? conversation.serializerVersion : serializerVersion;
                    ;
                    var allowedDomains = [conversation.url].concat(serviceEndPoint._appDomains[messageObject._conversationId]);
                    if (!_checkOriginWithAppDomains(allowedDomains, e.origin)) {
                        if (!OfficeExt.appSpecificCheckOrigin(allowedDomains, e, messageObject._origin, _checkOriginWithAppDomains)) {
                            var isOriginSubdomain = _isTargetSubdomainOfSourceLocation(conversation.url, e.origin);
                            if (!isOriginSubdomain) {
                                throw "Failed origin check";
                            }
                        }
                    }
                    var dataWithOrigin = (messageObject._data != null) ? messageObject._data : {};
                    dataWithOrigin.SecurityOrigin = e.origin;
                    var policyManager = serviceEndPoint.getPolicyManager();
                    if (policyManager && !policyManager.checkPermission(messageObject._conversationId, messageObject._actionName, dataWithOrigin)) {
                        throw "Access Denied";
                    }
                    var methodObject = _lookupMethodObject(serviceEndPoint, messageObject);
                    var invokeCompleteCallback = new Microsoft.Office.Common.InvokeCompleteCallback(e.source, requesterUrl, messageObject._actionName, messageObject._conversationId, messageObject._correlationId, _postCallbackHandler, serializerVersion);
                    var invoker = new Microsoft.Office.Common.Invoker(methodObject, dataWithOrigin, invokeCompleteCallback, serviceEndPoint._eventHandlerProxyList, messageObject._conversationId, messageObject._actionName, serializerVersion);
                    var shouldEnque = true;
                    if (_messageProcessingTimer == null) {
                        if ((_lastMessageProcessTime == null || (new Date()).getTime() - _lastMessageProcessTime > _processInterval) && !_blockingFlag) {
                            _executeCommand(invoker);
                            shouldEnque = false;
                        }
                        else {
                            _messageProcessingTimer = setInterval(_dequeInvoker, _processInterval);
                        }
                    }
                    if (shouldEnque) {
                        _enqueInvoker(invoker);
                    }
                }
                catch (ex) {
                    if (serviceEndPoint && serviceEndPoint._onHandleRequestError) {
                        serviceEndPoint._onHandleRequestError(messageObject, ex);
                    }
                    var errorCode = Microsoft.Office.Common.InvokeResultCode.errorHandlingRequest;
                    if (ex == "Access Denied") {
                        errorCode = Microsoft.Office.Common.InvokeResultCode.errorHandlingRequestAccessDenied;
                    }
                    var callResponse = new Microsoft.Office.Common.Response(messageObject._actionName, messageObject._conversationId, messageObject._correlationId, errorCode, Microsoft.Office.Common.ResponseType.forCalling, ex);
                    var envelopedResult = Microsoft.Office.Common.MessagePackager.envelope(callResponse, serializerVersion);
                    var canPostMessage = false;
                    try {
                        canPostMessage = !!(e.source && e.source.postMessage);
                    }
                    catch (ex) {
                    }
                    var isOriginValid = false;
                    if (window.location.href && e.origin && e.origin !== "null" && _isTargetSubdomainOfSourceLocation(window.location.href, e.origin)) {
                        isOriginValid = true;
                    }
                    else {
                        if (e.origin && e.origin !== "null") {
                            if (OSF.OUtil.checkFlight(OSF.FlightNames.AddinEnforceHttps)) {
                                var hostname = OSF.OUtil.parseUrl(e.origin, true).hostname;
                                isOriginValid = _isHostNameValidWacDomain(hostname);
                            }
                            else {
                                var parser = document.createElement("a");
                                parser.href = e.origin;
                                isOriginValid = _isHostNameValidWacDomain(parser.hostname);
                            }
                        }
                    }
                    if (canPostMessage && isOriginValid) {
                        e.source.postMessage(envelopedResult, requesterUrl);
                    }
                }
            }
            else if (messageObject._messageType === Microsoft.Office.Common.MessageType.response) {
                var clientEndPoint = _lookupClientEndPoint(messageObject._conversationId);
                if (!clientEndPoint) {
                    return;
                }
                clientEndPoint._serializerVersion = serializerVersion;
                ;
                if (!_checkOrigin(clientEndPoint._targetUrl, e.origin)) {
                    throw "Failed orgin check";
                }
                if (messageObject._responseType === Microsoft.Office.Common.ResponseType.forCalling) {
                    var callbackEntry = clientEndPoint._callbackList[messageObject._correlationId];
                    if (callbackEntry) {
                        try {
                            if (callbackEntry.callback)
                                callbackEntry.callback(messageObject._errorCode, messageObject._data);
                        }
                        finally {
                            delete clientEndPoint._callbackList[messageObject._correlationId];
                        }
                    }
                }
                else {
                    var eventhandler = clientEndPoint._eventHandlerList[messageObject._actionName];
                    if (eventhandler !== undefined && eventhandler !== null) {
                        eventhandler(messageObject._data);
                    }
                }
            }
            else {
                return;
            }
        }
    }
    ;
    function _initialize() {
        if (!_initialized) {
            _registerListener(_receive);
            _initialized = true;
        }
    }
    ;
    return {
        connect: function Microsoft_Office_Common_XdmCommunicationManager$connect(conversationId, targetWindow, targetUrl, serializerVersion) {
            var clientEndPoint = _clientEndPoints[conversationId];
            if (!clientEndPoint) {
                _initialize();
                clientEndPoint = new Microsoft.Office.Common.ClientEndPoint(conversationId, targetWindow, targetUrl, serializerVersion);
                _clientEndPoints[conversationId] = clientEndPoint;
            }
            return clientEndPoint;
        },
        getClientEndPoint: function Microsoft_Office_Common_XdmCommunicationManager$getClientEndPoint(conversationId) {
            var e = Function._validateParams(arguments, [
                { name: "conversationId", type: String, mayBeNull: false }
            ]);
            if (e)
                throw e;
            return _clientEndPoints[conversationId];
        },
        createServiceEndPoint: function Microsoft_Office_Common_XdmCommunicationManager$createServiceEndPoint(serviceEndPointId) {
            _initialize();
            var serviceEndPoint = new Microsoft.Office.Common.ServiceEndPoint(serviceEndPointId);
            _serviceEndPoints[serviceEndPointId] = serviceEndPoint;
            return serviceEndPoint;
        },
        getServiceEndPoint: function Microsoft_Office_Common_XdmCommunicationManager$getServiceEndPoint(serviceEndPointId) {
            var e = Function._validateParams(arguments, [
                { name: "serviceEndPointId", type: String, mayBeNull: false }
            ]);
            if (e)
                throw e;
            return _serviceEndPoints[serviceEndPointId];
        },
        deleteClientEndPoint: function Microsoft_Office_Common_XdmCommunicationManager$deleteClientEndPoint(conversationId) {
            var e = Function._validateParams(arguments, [
                { name: "conversationId", type: String, mayBeNull: false }
            ]);
            if (e)
                throw e;
            delete _clientEndPoints[conversationId];
        },
        deleteServiceEndPoint: function Microsoft_Office_Common_XdmCommunicationManager$deleteServiceEndPoint(serviceEndPointId) {
            var e = Function._validateParams(arguments, [
                { name: "serviceEndPointId", type: String, mayBeNull: false }
            ]);
            if (e)
                throw e;
            delete _serviceEndPoints[serviceEndPointId];
        },
        urlCompare: function Microsoft_Office_Common_XdmCommunicationManager$_urlCompare(url_parser1, url_parser2) {
            return _urlCompare(url_parser1, url_parser2);
        },
        checkUrlWithAppDomains: function Microsoft_Office_Common_XdmCommunicationManager$_checkUrlWithAppDomains(appDomains, origin) {
            return _checkOriginWithAppDomains(appDomains, origin);
        },
        isTargetSubdomainOfSourceLocation: function Microsoft_Office_Common_XdmCommunicationManager$_isTargetSubdomainOfSourceLocation(sourceLocation, messageOrigin) {
            return _isTargetSubdomainOfSourceLocation(sourceLocation, messageOrigin);
        },
        _setMethodTimeout: function Microsoft_Office_Common_XdmCommunicationManager$_setMethodTimeout(methodTimeout) {
            var e = Function._validateParams(arguments, [
                { name: "methodTimeout", type: Number, mayBeNull: false }
            ]);
            if (e)
                throw e;
            _methodTimeout = (methodTimeout <= 0) ? _methodTimeoutDefault : methodTimeout;
        },
        _startMethodTimeoutTimer: function Microsoft_Office_Common_XdmCommunicationManager$_startMethodTimeoutTimer() {
            if (!_methodTimeoutTimer) {
                _methodTimeoutTimer = setInterval(_checkMethodTimeout, _methodTimeoutProcessInterval);
            }
        },
        isHostNameValidWacDomain: function Microsoft_Office_Common_XdmCommunicationManager$_isHostNameValidWacDomain(hostName) {
            return _isHostNameValidWacDomain(hostName);
        },
        _hasSamePort: function Microsoft_Office_Common_XdmCommunicationManager$_hasSamePort(url_parser1, url_parser2) {
            return _hasSamePort(url_parser1, url_parser2);
        }
    };
})();
Microsoft.Office.Common.Message = function Microsoft_Office_Common_Message(messageType, actionName, conversationId, correlationId, data) {
    var e = Function._validateParams(arguments, [{ name: "messageType", type: Number, mayBeNull: false },
        { name: "actionName", type: String, mayBeNull: false },
        { name: "conversationId", type: String, mayBeNull: false },
        { name: "correlationId", mayBeNull: false },
        { name: "data", mayBeNull: true, optional: true }
    ]);
    if (e)
        throw e;
    this._messageType = messageType;
    this._actionName = actionName;
    this._conversationId = conversationId;
    this._correlationId = correlationId;
    this._origin = window.location.origin;
    if (typeof data == "undefined") {
        this._data = null;
    }
    else {
        this._data = data;
    }
};
Microsoft.Office.Common.Message.prototype = {
    getActionName: function Microsoft_Office_Common_Message$getActionName() {
        return this._actionName;
    },
    getConversationId: function Microsoft_Office_Common_Message$getConversationId() {
        return this._conversationId;
    },
    getCorrelationId: function Microsoft_Office_Common_Message$getCorrelationId() {
        return this._correlationId;
    },
    getOrigin: function Microsoft_Office_Common_Message$getOrigin() {
        return this._origin;
    },
    getData: function Microsoft_Office_Common_Message$getData() {
        return this._data;
    },
    getMessageType: function Microsoft_Office_Common_Message$getMessageType() {
        return this._messageType;
    }
};
Microsoft.Office.Common.Request = function Microsoft_Office_Common_Request(actionName, actionType, conversationId, correlationId, data) {
    Microsoft.Office.Common.Request.uber.constructor.call(this, Microsoft.Office.Common.MessageType.request, actionName, conversationId, correlationId, data);
    this._actionType = actionType;
};
OSF.OUtil.extend(Microsoft.Office.Common.Request, Microsoft.Office.Common.Message);
Microsoft.Office.Common.Request.prototype.getActionType = function Microsoft_Office_Common_Request$getActionType() {
    return this._actionType;
};
Microsoft.Office.Common.Response = function Microsoft_Office_Common_Response(actionName, conversationId, correlationId, errorCode, responseType, data) {
    Microsoft.Office.Common.Response.uber.constructor.call(this, Microsoft.Office.Common.MessageType.response, actionName, conversationId, correlationId, data);
    this._errorCode = errorCode;
    this._responseType = responseType;
};
OSF.OUtil.extend(Microsoft.Office.Common.Response, Microsoft.Office.Common.Message);
Microsoft.Office.Common.Response.prototype.getErrorCode = function Microsoft_Office_Common_Response$getErrorCode() {
    return this._errorCode;
};
Microsoft.Office.Common.Response.prototype.getResponseType = function Microsoft_Office_Common_Response$getResponseType() {
    return this._responseType;
};
Microsoft.Office.Common.MessagePackager = {
    envelope: function Microsoft_Office_Common_MessagePackager$envelope(messageObject, serializerVersion) {
        if (typeof (messageObject) === "object") {
            messageObject._serializerVersion = OSF.SerializerVersion.Browser;
        }
        return JSON.stringify(messageObject);
    },
    unenvelope: function Microsoft_Office_Common_MessagePackager$unenvelope(messageObject, serializerVersion) {
        return JSON.parse(messageObject);
    }
};
Microsoft.Office.Common.ResponseSender = function Microsoft_Office_Common_ResponseSender(requesterWindow, requesterUrl, actionName, conversationId, correlationId, responseType, serializerVersion) {
    var e = Function._validateParams(arguments, [{ name: "requesterWindow", mayBeNull: false },
        { name: "requesterUrl", type: String, mayBeNull: false },
        { name: "actionName", type: String, mayBeNull: false },
        { name: "conversationId", type: String, mayBeNull: false },
        { name: "correlationId", mayBeNull: false },
        { name: "responsetype", type: Number, maybeNull: false },
        { name: "serializerVersion", type: Number, maybeNull: true, optional: true }
    ]);
    if (e)
        throw e;
    this._requesterWindow = requesterWindow;
    this._requesterUrl = requesterUrl;
    this._actionName = actionName;
    this._conversationId = conversationId;
    this._correlationId = correlationId;
    this._invokeResultCode = Microsoft.Office.Common.InvokeResultCode.noError;
    this._responseType = responseType;
    var me = this;
    this._send = function (result) {
        var validateTargetMatchesRequester = function (targetOrigin, requesterUrl) {
            var parsedTargetOrigin = OSF.OUtil.parseUrl(targetOrigin);
            var parsedRequesterUrl = OSF.OUtil.parseUrl(requesterUrl);
            if (!parsedTargetOrigin || !parsedTargetOrigin.hostname || !parsedRequesterUrl || !parsedRequesterUrl.hostname) {
                var errorMessage = "Failed to execute 'postMessage' on 'DOMWindow': The target origin provided or the recipient window's origin are undefined.";
                console.log(errorMessage);
                ;
                return false;
            }
            else if (!Microsoft.Office.Common.XdmCommunicationManager.urlCompare(parsedTargetOrigin, parsedRequesterUrl)) {
                var targetOriginToConsoleLog = parsedTargetOrigin ? (parsedTargetOrigin.protocol + "//" + parsedTargetOrigin.hostname + (parsedTargetOrigin.port ? (":" + parsedTargetOrigin.port) : "")) : 'undefined';
                var requesterUrlToConsoleLog = parsedRequesterUrl ? (parsedRequesterUrl.protocol + "//" + parsedRequesterUrl.hostname + (parsedRequesterUrl.port ? (":" + parsedRequesterUrl.port) : "")) : 'undefined';
                var errorMessage = "Failed to execute 'postMessage' on 'DOMWindow': The target origin provided ('" + targetOriginToConsoleLog + "') does not match the recipient window's origin ('" + requesterUrlToConsoleLog + "').";
                console.log(errorMessage);
                ;
                return false;
            }
            return true;
        };
        try {
            if (me._actionName === "dialogMessageReceived" && result.targetOrigin !== "*") {
                if (result.targetOrigin) {
                    if (!validateTargetMatchesRequester(result.targetOrigin, me._requesterUrl)) {
                        return;
                    }
                }
            }
            else if (me._actionName === "dialogParentMessageReceived" && result.targetOrigin && result.targetOrigin !== "*") {
                if (!validateTargetMatchesRequester(result.targetOrigin, me._requesterUrl)) {
                    return;
                }
            }
            var response = new Microsoft.Office.Common.Response(me._actionName, me._conversationId, me._correlationId, me._invokeResultCode, me._responseType, result);
            var envelopedResult = Microsoft.Office.Common.MessagePackager.envelope(response, serializerVersion);
            me._requesterWindow.postMessage(envelopedResult, me._requesterUrl);
            ;
        }
        catch (ex) {
            OsfMsAjaxFactory.msAjaxDebug.trace("ResponseSender._send error:" + ex.message);
        }
    };
};
Microsoft.Office.Common.ResponseSender.prototype = {
    getRequesterWindow: function Microsoft_Office_Common_ResponseSender$getRequesterWindow() {
        return this._requesterWindow;
    },
    getRequesterUrl: function Microsoft_Office_Common_ResponseSender$getRequesterUrl() {
        return this._requesterUrl;
    },
    getActionName: function Microsoft_Office_Common_ResponseSender$getActionName() {
        return this._actionName;
    },
    getConversationId: function Microsoft_Office_Common_ResponseSender$getConversationId() {
        return this._conversationId;
    },
    getCorrelationId: function Microsoft_Office_Common_ResponseSender$getCorrelationId() {
        return this._correlationId;
    },
    getSend: function Microsoft_Office_Common_ResponseSender$getSend() {
        return this._send;
    },
    setResultCode: function Microsoft_Office_Common_ResponseSender$setResultCode(resultCode) {
        this._invokeResultCode = resultCode;
    }
};
Microsoft.Office.Common.InvokeCompleteCallback = function Microsoft_Office_Common_InvokeCompleteCallback(requesterWindow, requesterUrl, actionName, conversationId, correlationId, postCallbackHandler, serializerVersion) {
    Microsoft.Office.Common.InvokeCompleteCallback.uber.constructor.call(this, requesterWindow, requesterUrl, actionName, conversationId, correlationId, Microsoft.Office.Common.ResponseType.forCalling, serializerVersion);
    this._postCallbackHandler = postCallbackHandler;
    var me = this;
    this._send = function (result, responseCode) {
        if (responseCode != undefined) {
            me._invokeResultCode = responseCode;
        }
        try {
            var response = new Microsoft.Office.Common.Response(me._actionName, me._conversationId, me._correlationId, me._invokeResultCode, me._responseType, result);
            var envelopedResult = Microsoft.Office.Common.MessagePackager.envelope(response, serializerVersion);
            me._requesterWindow.postMessage(envelopedResult, me._requesterUrl);
            me._postCallbackHandler();
        }
        catch (ex) {
            OsfMsAjaxFactory.msAjaxDebug.trace("InvokeCompleteCallback._send error:" + ex.message);
        }
    };
};
OSF.OUtil.extend(Microsoft.Office.Common.InvokeCompleteCallback, Microsoft.Office.Common.ResponseSender);
Microsoft.Office.Common.Invoker = function Microsoft_Office_Common_Invoker(methodObject, paramValue, invokeCompleteCallback, eventHandlerProxyList, conversationId, eventName, serializerVersion) {
    var e = Function._validateParams(arguments, [
        { name: "methodObject", mayBeNull: false },
        { name: "paramValue", mayBeNull: true },
        { name: "invokeCompleteCallback", mayBeNull: false },
        { name: "eventHandlerProxyList", mayBeNull: true },
        { name: "conversationId", type: String, mayBeNull: false },
        { name: "eventName", type: String, mayBeNull: false },
        { name: "serializerVersion", type: Number, mayBeNull: true, optional: true }
    ]);
    if (e)
        throw e;
    this._methodObject = methodObject;
    this._param = paramValue;
    this._invokeCompleteCallback = invokeCompleteCallback;
    this._eventHandlerProxyList = eventHandlerProxyList;
    this._conversationId = conversationId;
    this._eventName = eventName;
    this._serializerVersion = serializerVersion;
};
Microsoft.Office.Common.Invoker.prototype = {
    invoke: function Microsoft_Office_Common_Invoker$invoke() {
        try {
            var result;
            switch (this._methodObject.getInvokeType()) {
                case Microsoft.Office.Common.InvokeType.async:
                    this._methodObject.getMethod()(this._param, this._invokeCompleteCallback.getSend());
                    break;
                case Microsoft.Office.Common.InvokeType.sync:
                    result = this._methodObject.getMethod()(this._param);
                    this._invokeCompleteCallback.getSend()(result);
                    break;
                case Microsoft.Office.Common.InvokeType.syncRegisterEvent:
                    var eventHandlerProxy = this._createEventHandlerProxyObject(this._invokeCompleteCallback);
                    result = this._methodObject.getMethod()(eventHandlerProxy.getSend(), this._param);
                    this._eventHandlerProxyList[this._conversationId + this._eventName] = eventHandlerProxy.getSend();
                    this._invokeCompleteCallback.getSend()(result);
                    break;
                case Microsoft.Office.Common.InvokeType.syncUnregisterEvent:
                    var eventHandler = this._eventHandlerProxyList[this._conversationId + this._eventName];
                    result = this._methodObject.getMethod()(eventHandler, this._param);
                    delete this._eventHandlerProxyList[this._conversationId + this._eventName];
                    this._invokeCompleteCallback.getSend()(result);
                    break;
                case Microsoft.Office.Common.InvokeType.asyncRegisterEvent:
                    var eventHandlerProxyAsync = this._createEventHandlerProxyObject(this._invokeCompleteCallback);
                    this._methodObject.getMethod()(eventHandlerProxyAsync.getSend(), this._invokeCompleteCallback.getSend(), this._param);
                    this._eventHandlerProxyList[this._callerId + this._eventName] = eventHandlerProxyAsync.getSend();
                    break;
                case Microsoft.Office.Common.InvokeType.asyncUnregisterEvent:
                    var eventHandlerAsync = this._eventHandlerProxyList[this._callerId + this._eventName];
                    this._methodObject.getMethod()(eventHandlerAsync, this._invokeCompleteCallback.getSend(), this._param);
                    delete this._eventHandlerProxyList[this._callerId + this._eventName];
                    break;
                default:
                    break;
            }
        }
        catch (ex) {
            this._invokeCompleteCallback.setResultCode(Microsoft.Office.Common.InvokeResultCode.errorInResponse);
            this._invokeCompleteCallback.getSend()(ex);
        }
    },
    getInvokeBlockingFlag: function Microsoft_Office_Common_Invoker$getInvokeBlockingFlag() {
        return this._methodObject.getBlockingFlag();
    },
    _createEventHandlerProxyObject: function Microsoft_Office_Common_Invoker$_createEventHandlerProxyObject(invokeCompleteObject) {
        return new Microsoft.Office.Common.ResponseSender(invokeCompleteObject.getRequesterWindow(), invokeCompleteObject.getRequesterUrl(), invokeCompleteObject.getActionName(), invokeCompleteObject.getConversationId(), invokeCompleteObject.getCorrelationId(), Microsoft.Office.Common.ResponseType.forEventing, this._serializerVersion);
    }
};
OSF.OUtil.setNamespace("WAC", OSF.DDA);
OSF.DDA.WAC.UniqueArguments = {
    Data: "Data",
    Properties: "Properties",
    BindingRequest: "DdaBindingsMethod",
    BindingResponse: "Bindings",
    SingleBindingResponse: "singleBindingResponse",
    GetData: "DdaGetBindingData",
    AddRowsColumns: "DdaAddRowsColumns",
    SetData: "DdaSetBindingData",
    ClearFormats: "DdaClearBindingFormats",
    SetFormats: "DdaSetBindingFormats",
    SettingsRequest: "DdaSettingsMethod",
    BindingEventSource: "ddaBinding",
    ArrayData: "ArrayData"
};
OSF.OUtil.setNamespace("Delegate", OSF.DDA.WAC);
OSF.DDA.WAC.Delegate.SpecialProcessor = function OSF_DDA_WAC_Delegate_SpecialProcessor() {
    var complexTypes = [
        OSF.DDA.WAC.UniqueArguments.SingleBindingResponse,
        OSF.DDA.WAC.UniqueArguments.BindingRequest,
        OSF.DDA.WAC.UniqueArguments.BindingResponse,
        OSF.DDA.WAC.UniqueArguments.GetData,
        OSF.DDA.WAC.UniqueArguments.AddRowsColumns,
        OSF.DDA.WAC.UniqueArguments.SetData,
        OSF.DDA.WAC.UniqueArguments.ClearFormats,
        OSF.DDA.WAC.UniqueArguments.SetFormats,
        OSF.DDA.WAC.UniqueArguments.SettingsRequest,
        OSF.DDA.WAC.UniqueArguments.BindingEventSource
    ];
    var dynamicTypes = {};
    OSF.DDA.WAC.Delegate.SpecialProcessor.uber.constructor.call(this, complexTypes, dynamicTypes);
};
OSF.OUtil.extend(OSF.DDA.WAC.Delegate.SpecialProcessor, OSF.DDA.SpecialProcessor);
OSF.DDA.WAC.Delegate.ParameterMap = OSF.DDA.getDecoratedParameterMap(new OSF.DDA.WAC.Delegate.SpecialProcessor(), []);
OSF.OUtil.setNamespace("WAC", OSF.DDA);
OSF.OUtil.setNamespace("Delegate", OSF.DDA.WAC);
OSF.DDA.WAC.getDelegateMethods = function OSF_DDA_WAC_getDelegateMethods() {
    var delegateMethods = {};
    delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync] = OSF.DDA.WAC.Delegate.executeAsync;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync] = OSF.DDA.WAC.Delegate.registerEventAsync;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync] = OSF.DDA.WAC.Delegate.unregisterEventAsync;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.OpenDialog] = OSF.DDA.WAC.Delegate.openDialog;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.ValidateTaskpaneDomain] = OSF.DDA.WAC.Delegate.validateTaskpaneDomain;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.MessageParent] = OSF.DDA.WAC.Delegate.messageParent;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.SendMessage] = OSF.DDA.WAC.Delegate.sendMessage;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.CloseDialog] = OSF.DDA.WAC.Delegate.closeDialog;
    return delegateMethods;
};
OSF.DDA.WAC.Delegate.version = 1;
OSF.DDA.WAC.Delegate.executeAsync = function OSF_DDA_WAC_Delegate$executeAsync(args) {
    if (!args.hostCallArgs) {
        args.hostCallArgs = {};
    }
    args.hostCallArgs["DdaMethod"] = {
        "ControlId": OSF._OfficeAppFactory.getId(),
        "Version": OSF.DDA.WAC.Delegate.version,
        "DispatchId": args.dispId
    };
    args.hostCallArgs["__timeout__"] = -1;
    if (args.onCalling) {
        args.onCalling();
    }
    if (!OSF.getClientEndPoint()) {
        return;
    }
    OSF.getClientEndPoint().invoke("executeMethod", function OSF_DDA_WAC_Delegate$OMFacade$OnResponse(xdmStatus, payload) {
        if (args.onReceiving) {
            args.onReceiving();
        }
        var error;
        if (xdmStatus == Microsoft.Office.Common.InvokeResultCode.noError) {
            OSF.DDA.WAC.Delegate.version = payload["Version"];
            error = payload["Error"];
        }
        else {
            switch (xdmStatus) {
                case Microsoft.Office.Common.InvokeResultCode.errorHandlingRequestAccessDenied:
                    error = OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;
                    break;
                default:
                    error = OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
                    break;
            }
        }
        if (args.onComplete) {
            args.onComplete(error, payload);
        }
    }, args.hostCallArgs);
};
OSF.DDA.WAC.Delegate._getOnAfterRegisterEvent = function OSF_DDA_WAC_Delegate$GetOnAfterRegisterEvent(register, args) {
    var startTime = (new Date()).getTime();
    return function OSF_DDA_WAC_Delegate$OnAfterRegisterEvent(xdmStatus, payload) {
        if (args.onReceiving) {
            args.onReceiving();
        }
        var status;
        if (xdmStatus != Microsoft.Office.Common.InvokeResultCode.noError) {
            switch (xdmStatus) {
                case Microsoft.Office.Common.InvokeResultCode.errorHandlingRequestAccessDenied:
                    status = OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;
                    break;
                default:
                    status = OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
                    break;
            }
        }
        else {
            if (payload) {
                if (payload["Error"]) {
                    status = payload["Error"];
                }
                else {
                    status = OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess;
                }
            }
            else {
                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
            }
        }
        if (args.onComplete) {
            args.onComplete(status);
        }
        if (OSF.AppTelemetry) {
            OSF.AppTelemetry.onRegisterDone(register, args.dispId, Math.abs((new Date()).getTime() - startTime), status);
        }
    };
};
OSF.DDA.WAC.Delegate.registerEventAsync = function OSF_DDA_WAC_Delegate$RegisterEventAsync(args) {
    if (args.onCalling) {
        args.onCalling();
    }
    if (!OSF.getClientEndPoint()) {
        return;
    }
    OSF.getClientEndPoint().registerForEvent(OSF.DDA.getXdmEventName(args.targetId, args.eventType), function OSF_DDA_WACOMFacade$OnEvent(payload) {
        if (args.onEvent) {
            args.onEvent(payload);
        }
        if (OSF.AppTelemetry) {
            OSF.AppTelemetry.onEventDone(args.dispId);
        }
    }, OSF.DDA.WAC.Delegate._getOnAfterRegisterEvent(true, args), {
        "controlId": OSF._OfficeAppFactory.getId(),
        "eventDispId": args.dispId,
        "targetId": args.targetId
    });
};
OSF.DDA.WAC.Delegate.unregisterEventAsync = function OSF_DDA_WAC_Delegate$UnregisterEventAsync(args) {
    if (args.onCalling) {
        args.onCalling();
    }
    if (!OSF.getClientEndPoint()) {
        return;
    }
    OSF.getClientEndPoint().unregisterForEvent(OSF.DDA.getXdmEventName(args.targetId, args.eventType), OSF.DDA.WAC.Delegate._getOnAfterRegisterEvent(false, args), {
        "controlId": OSF._OfficeAppFactory.getId(),
        "eventDispId": args.dispId,
        "targetId": args.targetId
    });
};
OSF.OUtil.setNamespace("WebApp", OSF);
OSF.OUtil.setNamespace("Messaging", OSF);
OSF.OUtil.setNamespace("ExtensionLifeCycle", OSF);
OSF.OUtil.setNamespace("TaskPaneAction", OSF);
OSF.OUtil.setNamespace("RibbonGallery", OSF);
OSF.WebApp.AddHostInfoAndXdmInfo = function OSF_WebApp$AddHostInfoAndXdmInfo(url) {
    if (OSF._OfficeAppFactory.getWindowLocationSearch && OSF._OfficeAppFactory.getWindowLocationHash) {
        return url + OSF._OfficeAppFactory.getWindowLocationSearch() + OSF._OfficeAppFactory.getWindowLocationHash();
    }
    else {
        return url;
    }
};
OSF.WebApp._UpdateLinksForHostAndXdmInfo = function OSF_WebApp$_UpdateLinksForHostAndXdmInfo() {
    var links = document.querySelectorAll("a[data-officejs-navigate]");
    for (var i = 0; i < links.length; i++) {
        if (OSF.WebApp._isGoodUrl(links[i].href)) {
            links[i].href = OSF.WebApp.AddHostInfoAndXdmInfo(links[i].href);
        }
    }
    var forms = document.querySelectorAll("form[data-officejs-navigate]");
    for (var i = 0; i < forms.length; i++) {
        var form = forms[i];
        if (OSF.WebApp._isGoodUrl(form.action)) {
            form.action = OSF.WebApp.AddHostInfoAndXdmInfo(form.action);
        }
    }
};
OSF.WebApp._isGoodUrl = function OSF_WebApp$_isGoodUrl(url) {
    if (typeof url == 'undefined')
        return false;
    url = url.trim();
    var colonIndex = url.indexOf(':');
    var protocol = colonIndex > 0 ? url.substr(0, colonIndex) : null;
    var goodUrl = protocol !== null ? protocol.toLowerCase() === "http" || protocol.toLowerCase() === "https" : true;
    goodUrl = goodUrl && url != "#" && url != "/" && url != "" && url != OSF._OfficeAppFactory.getWebAppState().webAppUrl;
    return goodUrl;
};
OSF.InitializationHelper = function OSF_InitializationHelper(hostInfo, webAppState, context, settings, hostFacade) {
    this._hostInfo = hostInfo;
    this._webAppState = webAppState;
    this._context = context;
    this._settings = settings;
    this._hostFacade = hostFacade;
    this._appContext = {};
    this._tabbableElements = "a[href]:not([tabindex='-1'])," + "area[href]:not([tabindex='-1'])," + "button:not([disabled]):not([tabindex='-1'])," + "input:not([disabled]):not([tabindex='-1'])," + "select:not([disabled]):not([tabindex='-1'])," + "textarea:not([disabled]):not([tabindex='-1'])," + "*[tabindex]:not([tabindex='-1'])," + "*[contenteditable]:not([disabled]):not([tabindex='-1'])";
    this._initializeSettings = function OSF_InitializationHelper$initializeSettings(appContext, refreshSupported) {
        var settings;
        var serializedSettings = appContext.get_settings();
        var deserializedSettings = OSF.DDA.SettingsManager.deserializeSettings(serializedSettings);
        if (refreshSupported) {
            settings = new OSF.DDA.RefreshableSettings(deserializedSettings);
        }
        else {
            settings = new OSF.DDA.Settings(deserializedSettings);
        }
        return settings;
    };
    var windowOpen = function OSF_InitializationHelper$windowOpen(windowObj) {
        var proxy = window.open;
        windowObj.open = function (strUrl, strWindowName, strWindowFeatures) {
            var windowObject = null;
            try {
                windowObject = proxy(strUrl, strWindowName, strWindowFeatures);
            }
            catch (ex) {
                if (OSF.AppTelemetry) {
                    OSF.AppTelemetry.logAppCommonMessage("Exception happens at windowOpen." + ex);
                }
            }
            if (!windowObject) {
                var params = {
                    "strUrl": strUrl,
                    "strWindowName": strWindowName,
                    "strWindowFeatures": strWindowFeatures
                };
                if (OSF._OfficeAppFactory.getClientEndPoint()) {
                    OSF._OfficeAppFactory.getClientEndPoint().invoke("ContextActivationManager_openWindowInHost", null, params);
                }
            }
            return windowObject;
        };
    };
    windowOpen(window);
};
OSF.InitializationHelper.prototype.saveAndSetDialogInfo = function OSF_InitializationHelper$saveAndSetDialogInfo(hostInfoValue) {
    var getAppIdFromWindowLocation = function OSF_InitializationHelper$getAppIdFromWindowLocation() {
        var xdmInfoValue = OSF.OUtil.parseXdmInfo(true);
        if (xdmInfoValue) {
            var items = xdmInfoValue.split("|");
            return items[1];
        }
        return null;
    };
    var osfSessionStorage = OSF.OUtil.getSessionStorage();
    if (osfSessionStorage) {
        if (!hostInfoValue) {
            hostInfoValue = OSF.OUtil.parseHostInfoFromWindowName(true, OSF._OfficeAppFactory.getWindowName());
        }
        if (hostInfoValue && hostInfoValue.indexOf("isDialog") > -1) {
            var appId = getAppIdFromWindowLocation();
            if (appId != null) {
                osfSessionStorage.setItem(appId + "IsDialog", "true");
            }
            this._hostInfo.isDialog = true;
            return;
        }
        this._hostInfo.isDialog = osfSessionStorage.getItem(OSF.OUtil.getXdmFieldValue(OSF.XdmFieldName.AppId, false) + "IsDialog") != null ? true : false;
    }
};
OSF.InitializationHelper.prototype.getAppContext = function OSF_InitializationHelper$getAppContext(wnd, gotAppContext) {
    if (OSF.AppTelemetry) {
        OSF.AppTelemetry.logAppCommonMessage("OsfControl activation lifecycle: getAppContext got called.");
    }
    var me = this;
    var getInvocationCallbackWebApp = function OSF_InitializationHelper_getAppContextAsync$getInvocationCallbackWebApp(errorCode, appContext) {
        var settings;
        if (appContext._appName === OSF.AppName.ExcelWebApp) {
            var serializedSettings = appContext._settings;
            settings = {};
            for (var index in serializedSettings) {
                var setting = serializedSettings[index];
                settings[setting[0]] = setting[1];
            }
        }
        else {
            settings = appContext._settings;
        }
        if (appContext._appName === OSF.AppName.OutlookWebApp && !!appContext._requirementMatrix && appContext._requirementMatrix.indexOf("react") == -1) {
            OSF.AgaveHostAction.SendTelemetryEvent = undefined;
        }
        if (!me._hostInfo.isDialog || window.opener == null) {
            var pageUrl = window.location.origin;
            me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.UpdateTargetUrl, pageUrl]);
        }
        if (errorCode === 0 && appContext._id != undefined && appContext._appName != undefined && appContext._appVersion != undefined && appContext._appUILocale != undefined && appContext._dataLocale != undefined &&
            appContext._docUrl != undefined && appContext._clientMode != undefined && appContext._settings != undefined && appContext._reason != undefined) {
            me._appContext = appContext;
            var appInstanceId = (appContext._appInstanceId ? appContext._appInstanceId : appContext._id);
            var touchEnabled = false;
            var commerceAllowed = true;
            var minorVersion = 0;
            if (appContext._appMinorVersion != undefined) {
                minorVersion = appContext._appMinorVersion;
            }
            var requirementMatrix = undefined;
            if (appContext._requirementMatrix != undefined) {
                requirementMatrix = appContext._requirementMatrix;
            }
            appContext.eToken = appContext.eToken ? appContext.eToken : "";
            var returnedContext = new OSF.OfficeAppContext(appContext._id, appContext._appName, appContext._appVersion, appContext._appUILocale, appContext._dataLocale, appContext._docUrl, appContext._clientMode, settings, appContext._reason, appContext._osfControlType, appContext._eToken, appContext._correlationId, appInstanceId, touchEnabled, commerceAllowed, minorVersion, requirementMatrix, appContext._hostCustomMessage, appContext._hostFullVersion, appContext._clientWindowHeight, appContext._clientWindowWidth, appContext._addinName, appContext._appDomains, appContext._dialogRequirementMatrix, appContext._featureGates, undefined, appContext._initialDisplayMode);
            returnedContext._wacHostEnvironment = appContext._wacHostEnvironment || "0";
            returnedContext._isFromWacAutomation = !!appContext._isFromWacAutomation;
            if (OSF.AppTelemetry) {
                OSF.AppTelemetry.initialize(returnedContext);
            }
            gotAppContext(returnedContext);
        }
        else {
            var errorMsg = "Function ContextActivationManager_getAppContextAsync call failed. ErrorCode is " + errorCode + ", exception: " + appContext;
            if (OSF.AppTelemetry) {
                OSF.AppTelemetry.logAppException(errorMsg);
            }
            throw errorMsg;
        }
    };
    try {
        if (this._hostInfo.isDialog && window.opener != null) {
            var appContext = OfficeExt.WACUtils.parseAppContextFromWindowName(false, OSF._OfficeAppFactory.getWindowName());
            getInvocationCallbackWebApp(0, appContext);
        }
        else {
            this._webAppState.clientEndPoint.invoke("ContextActivationManager_getAppContextAsync", getInvocationCallbackWebApp, this._webAppState.id);
        }
    }
    catch (ex) {
        if (OSF.AppTelemetry) {
            OSF.AppTelemetry.logAppException("Exception thrown when trying to invoke getAppContextAsync. Exception:[" + ex + "]");
        }
        throw ex;
    }
};
OSF.InitializationHelper.prototype.isHostOriginTrusted = function OSF_InitializationHelper$isHostOriginTrusted(hostOrigin) {
};
OSF.InitializationHelper.prototype.checkReceiverOriginAndRun = function OSF_InitializationHelper$checkReceiverOriginAndRun(funcToRun) {
    var me = this;
    var parsedHostname = OSF.OUtil.parseUrl(me._webAppState.clientEndPoint._targetUrl, false);
    var isHttps = parsedHostname.protocol == "https:";
    var notHttpsErrorMessage = "NotHttps";
    if (me._webAppState.clientEndPoint._hostTrustCheckStatus === Microsoft.Office.Common.HostTrustStatus.unknown) {
        if (!isHttps)
            me._webAppState.clientEndPoint._hostTrustCheckStatus = Microsoft.Office.Common.HostTrustStatus.nothttps;
        if (me._webAppState.clientEndPoint._hostTrustCheckStatus != Microsoft.Office.Common.HostTrustStatus.nothttps) {
            var isOriginValid = Microsoft.Office.Common.XdmCommunicationManager.isHostNameValidWacDomain(parsedHostname.hostname);
            if (me.isHostOriginTrusted) {
                isOriginValid = isOriginValid || me.isHostOriginTrusted(parsedHostname.hostname);
            }
            if (isOriginValid)
                me._webAppState.clientEndPoint._hostTrustCheckStatus = Microsoft.Office.Common.HostTrustStatus.trusted;
        }
    }
    if (!me._webAppState.clientEndPoint._checkStatusLogged && me._hostInfo != null && me._hostInfo !== undefined) {
        OSF.AppTelemetry.onCheckWACHost(me._webAppState.clientEndPoint._hostTrustCheckStatus, me._webAppState.id, me._hostInfo.hostType, me._hostInfo.hostPlatform, me._webAppState.clientEndPoint._targetUrl);
        me._webAppState.clientEndPoint._checkStatusLogged = true;
    }
    if (me._webAppState.clientEndPoint._hostTrustCheckStatus != Microsoft.Office.Common.HostTrustStatus.trusted) {
        var loadAgaveErrorUX = function () {
            var officejsCDNHost = OSF._OfficeAppFactory.getLoadScriptHelper().getOfficeJsBasePath().match(/^https?:\/\/[^:/?#]*(?::([0-9]+))?/);
            if (officejsCDNHost && officejsCDNHost[0]) {
                var agaveErrorUXPath = OSF._OfficeAppFactory.getLoadScriptHelper().getOfficeJsBasePath() + 'AgaveErrorUX/index.html#';
                var hashObj = {
                    error: "NotTrustedWAC",
                    locale: OSF.getSupportedLocale(me._hostInfo.hostLocale, OSF.ConstantNames.DefaultLocale),
                    hostname: parsedHostname.hostname,
                    noHttps: !isHttps,
                    validate: false
                };
                var hostUserTrustIframe = document.createElement("iframe");
                hostUserTrustIframe.style.visibility = "hidden";
                hostUserTrustIframe.style.height = "0";
                hostUserTrustIframe.style.width = "0";
                function hostUserTrustCallback(event) {
                    if ((event.source == hostUserTrustIframe.contentWindow) &&
                        (event.origin == officejsCDNHost[0])) {
                        try {
                            var receivedObj = JSON.parse(event.data);
                            var e = Function._validateParams(receivedObj, [{ name: "hostUserTrusted", type: Boolean, mayBeNull: false }
                            ]);
                            if (receivedObj.hostUserTrusted === true) {
                                me._webAppState.clientEndPoint._hostTrustCheckStatus = Microsoft.Office.Common.HostTrustStatus.trusted;
                                OSF.OUtil.removeEventListener(window, "message", hostUserTrustCallback);
                                document.body.removeChild(hostUserTrustIframe);
                            }
                            else {
                                hashObj.validate = false;
                                window.location.replace(agaveErrorUXPath + encodeURIComponent(JSON.stringify(hashObj)));
                            }
                            funcToRun();
                        }
                        catch (e) {
                            document.body.innerHTML = Strings.OfficeOM.L_NotTrustedWAC;
                        }
                    }
                }
                ;
                OSF.OUtil.addEventListener(window, "message", hostUserTrustCallback);
                hashObj.validate = true;
                hostUserTrustIframe.setAttribute('src', agaveErrorUXPath + encodeURIComponent(JSON.stringify(hashObj)));
                hostUserTrustIframe.onload = function () {
                    var postingObj = {
                        hostname: parsedHostname.hostname,
                        noHttps: !isHttps
                    };
                    hostUserTrustIframe.contentWindow.postMessage(JSON.stringify(postingObj), officejsCDNHost[0]);
                };
                document.body.appendChild(hostUserTrustIframe);
            }
            else {
                document.body.innerHTML = Strings.OfficeOM.L_NotTrustedWAC;
            }
            if (OSF.OUtil.checkFlight(OSF.FlightNames.AddinEnforceHttps)) {
                if (!isHttps)
                    throw new Error(notHttpsErrorMessage);
            }
        };
        if (document.body) {
            loadAgaveErrorUX();
        }
        else {
            var checkDone = false;
            document.addEventListener('DOMContentLoaded', function () {
                if (!checkDone) {
                    checkDone = true;
                    loadAgaveErrorUX();
                }
            });
        }
    }
    else {
        funcToRun();
    }
};
OSF.InitializationHelper.prototype.setAgaveHostCommunication = function OSF_InitializationHelper$setAgaveHostCommunication() {
    try {
        var me = this;
        var xdmInfoValue = OSF.OUtil.parseXdmInfoWithGivenFragment(false, OSF._OfficeAppFactory.getWindowLocationHash());
        if (!xdmInfoValue && OSF._OfficeAppFactory.getWindowName) {
            xdmInfoValue = OSF.OUtil.parseXdmInfoFromWindowName(false, OSF._OfficeAppFactory.getWindowName());
        }
        if (xdmInfoValue) {
            var xdmItems = OSF.OUtil.getInfoItems(xdmInfoValue);
            if (xdmItems != undefined && xdmItems.length >= 3) {
                me._webAppState.conversationID = xdmItems[0];
                me._webAppState.id = xdmItems[1];
                me._webAppState.webAppUrl = xdmItems[2].indexOf(":") >= 0 ? xdmItems[2] : decodeURIComponent(xdmItems[2]);
            }
        }
        me._webAppState.wnd = window.opener != null ? window.opener : window.parent;
        var serializerVersion = OSF.OUtil.parseSerializerVersionWithGivenFragment(false, OSF._OfficeAppFactory.getWindowLocationHash());
        if (isNaN(serializerVersion) && OSF._OfficeAppFactory.getWindowName) {
            serializerVersion = OSF.OUtil.parseSerializerVersionFromWindowName(false, OSF._OfficeAppFactory.getWindowName());
        }
        me._webAppState.serializerVersion = serializerVersion;
        if (this._hostInfo.isDialog && window.opener != null) {
            return;
        }
        me._webAppState.clientEndPoint = Microsoft.Office.Common.XdmCommunicationManager.connect(me._webAppState.conversationID, me._webAppState.wnd, me._webAppState.webAppUrl, me._webAppState.serializerVersion);
        me._webAppState.serviceEndPoint = Microsoft.Office.Common.XdmCommunicationManager.createServiceEndPoint(me._webAppState.id);
        me._webAppState.clientEndPoint._checkReceiverOriginAndRun = function (funcToRun) {
            me.checkReceiverOriginAndRun(funcToRun);
        };
        var notificationConversationId = me._webAppState.conversationID + OSF.SharedConstants.NotificationConversationIdSuffix;
        me._webAppState.serviceEndPoint.registerConversation(notificationConversationId, me._webAppState.webAppUrl);
        var notifyAgave = function OSF__OfficeAppFactory_initialize$notifyAgave(params) {
            var actionId;
            if (typeof params == "string") {
                actionId = params;
            }
            else {
                actionId = params[0];
            }
            switch (actionId) {
                case OSF.AgaveHostAction.Select:
                    me._webAppState.focused = true;
                    break;
                case OSF.AgaveHostAction.UnSelect:
                    me._webAppState.focused = false;
                    break;
                case OSF.AgaveHostAction.TabIn:
                case OSF.AgaveHostAction.CtrlF6In:
                    window.focus();
                    var list = document.querySelectorAll(me._tabbableElements);
                    var focused = OSF.OUtil.focusToFirstTabbable(list, false);
                    if (!focused) {
                        window.blur();
                        me._webAppState.focused = false;
                        me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.ExitNoFocusable]);
                    }
                    break;
                case OSF.AgaveHostAction.TabInShift:
                    window.focus();
                    var list = document.querySelectorAll(me._tabbableElements);
                    var focused = OSF.OUtil.focusToFirstTabbable(list, true);
                    if (!focused) {
                        window.blur();
                        me._webAppState.focused = false;
                        me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.ExitNoFocusableShift]);
                    }
                    break;
                case OSF.AgaveHostAction.SendMessage:
                    if (window.Office.context.messaging.onMessage) {
                        var message = params[1];
                        window.Office.context.messaging.onMessage(message);
                    }
                    break;
                case OSF.AgaveHostAction.TaskPaneHeaderButtonClicked:
                    if (window.Office.context.ui.taskPaneAction.onHeaderButtonClick) {
                        window.Office.context.ui.taskPaneAction.onHeaderButtonClick();
                    }
                    break;
                default:
                    OsfMsAjaxFactory.msAjaxDebug.trace("actionId " + actionId + " notifyAgave is wrong.");
                    break;
            }
        };
        me._webAppState.serviceEndPoint.registerMethod("Office_notifyAgave", notifyAgave, Microsoft.Office.Common.InvokeType.async, false);
        me.addOrRemoveEventListenersForWindow(true);
    }
    catch (ex) {
        if (OSF.AppTelemetry) {
            OSF.AppTelemetry.logAppException("Exception thrown in setAgaveHostCommunication. Exception:[" + ex + "]");
        }
        throw ex;
    }
};
OSF.InitializationHelper.prototype.addOrRemoveEventListenersForWindow = function OSF_InitializationHelper$addOrRemoveEventListenersForWindow(isAdd) {
    var me = this;
    var onWindowFocus = function () {
        if (!me._webAppState.focused) {
            me._webAppState.focused = true;
        }
        me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.Select]);
    };
    var onWindowBlur = function () {
        if (!OSF) {
            return;
        }
        if (me._webAppState.focused) {
            me._webAppState.focused = false;
        }
        me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.UnSelect]);
    };
    var onWindowKeydown = function (e) {
        e.preventDefault = e.preventDefault || function () {
            e.returnValue = false;
        };
        if (e.keyCode == 117 && (e.ctrlKey || e.metaKey)) {
            e.preventDefault();
            var actionId = OSF.AgaveHostAction.CtrlF6Exit;
            if (e.shiftKey) {
                actionId = OSF.AgaveHostAction.CtrlF6ExitShift;
            }
            me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, actionId]);
        }
        else if (e.keyCode == 9) {
            var isPowerPointModernSlideShow = me._appContext._appSettings &&
                (me._appContext._appSettings['PowerPointModernSlideShowEnabled'] || false) &&
                me._appContext._appName == OSF.AppName.PowerpointWebApp &&
                me._appContext._clientMode == OSF.ClientMode.ReadOnly &&
                me._appContext._osfControlType == OSF.OsfControlType.DocumentLevel;
            if (!isPowerPointModernSlideShow) {
                e.preventDefault();
            }
            var allTabbableElements = document.querySelectorAll(me._tabbableElements);
            var focused = OSF.OUtil.focusToNextTabbable(allTabbableElements, e.target || e.srcElement, e.shiftKey);
            if (focused) {
                if (isPowerPointModernSlideShow) {
                    e.preventDefault();
                }
            }
            else {
                if (me._hostInfo.isDialog) {
                    OSF.OUtil.focusToFirstTabbable(allTabbableElements, e.shiftKey);
                }
                else {
                    if (e.shiftKey) {
                        if (!isPowerPointModernSlideShow) {
                            me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.TabExitShift]);
                        }
                    }
                    else {
                        if (OSF.OUtil.checkFlight(OSF.FlightNames.SetFocusToTaskpaneIsEnabled) && e.target && e.target.tagName.toUpperCase() === "BODY") {
                            OSF.OUtil.focusToFirstTabbable(allTabbableElements, e.shiftKey);
                        }
                        else {
                            if (!isPowerPointModernSlideShow) {
                                me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.TabExit]);
                            }
                        }
                    }
                }
            }
        }
        else if (e.keyCode == 27) {
            e.preventDefault();
            me.dismissDialogNotification && me.dismissDialogNotification();
            me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.EscExit]);
        }
        else if (e.keyCode == 113) {
            e.preventDefault();
            me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.F2Exit]);
        }
        else if ((e.ctrlKey || e.metaKey || e.shiftKey || e.altKey) && e.keyCode >= 1 && e.keyCode <= 255) {
            var params = {
                "keyCode": e.keyCode,
                "shiftKey": e.shiftKey,
                "altKey": e.altKey,
                "ctrlKey": e.ctrlKey,
                "metaKey": e.metaKey
            };
            me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.KeyboardShortcuts, params]);
        }
    };
    var onWindowKeypress = function (e) {
        if (e.keyCode == 117 && e.ctrlKey) {
            if (e.preventDefault) {
                e.preventDefault();
            }
            else {
                e.returnValue = false;
            }
        }
    };
    if (isAdd) {
        OSF.OUtil.addEventListener(window, "focus", onWindowFocus);
        OSF.OUtil.addEventListener(window, "blur", onWindowBlur);
        OSF.OUtil.addEventListener(window, "keydown", onWindowKeydown);
        OSF.OUtil.addEventListener(window, "keypress", onWindowKeypress);
    }
    else {
        OSF.OUtil.removeEventListener(window, "focus", onWindowFocus);
        OSF.OUtil.removeEventListener(window, "blur", onWindowBlur);
        OSF.OUtil.removeEventListener(window, "keydown", onWindowKeydown);
        OSF.OUtil.removeEventListener(window, "keypress", onWindowKeypress);
    }
};
OSF.InitializationHelper.prototype.initWebDialog = function OSF_InitializationHelper$initWebDialog(appContext) {
    if (appContext.get_isDialog()) {
        if (OSF.DDA.UI.ChildUI) {
            var isPopupWindow = (window.opener != null);
            appContext.ui = new OSF.DDA.UI.ChildUI(isPopupWindow);
            if (isPopupWindow) {
                this.registerMessageReceivedEventForWindowDialog && this.registerMessageReceivedEventForWindowDialog();
            }
        }
    }
    else {
        if (OSF.DDA.UI.ParentUI) {
            appContext.ui = new OSF.DDA.UI.ParentUI();
            if (OfficeExt.Container) {
                OSF.DDA.DispIdHost.addAsyncMethods(appContext.ui, [OSF.DDA.AsyncMethodNames.CloseContainerAsync]);
            }
        }
    }
};
OSF.InitializationHelper.prototype.initWebAuth = function OSF_InitializationHelper$initWebAuth(appContext) {
    if (OSF.DDA.Auth) {
        appContext.auth = new OSF.DDA.Auth();
        OSF.DDA.DispIdHost.addAsyncMethods(appContext.auth, [OSF.DDA.AsyncMethodNames.GetAccessTokenAsync]);
    }
};
OSF.InitializationHelper.prototype.initWebAuthImplicit = function OSF_InitializationHelper$initWebAuthImplicit(appContext) {
    if (OSF.DDA.WebAuth) {
        appContext.webAuth = new OSF.DDA.WebAuth();
        OSF.DDA.DispIdHost.addAsyncMethods(appContext.webAuth, [OSF.DDA.AsyncMethodNames.GetAuthContextAsync]);
    }
};
OSF.getClientEndPoint = function OSF$getClientEndPoint() {
    var initializationHelper = OSF._OfficeAppFactory.getInitializationHelper();
    return initializationHelper._webAppState.clientEndPoint;
};
OSF.InitializationHelper.prototype.prepareRightAfterWebExtensionInitialize = function OSF_InitializationHelper$prepareRightAfterWebExtensionInitialize() {
    if (this._hostInfo.isDialog) {
        window.focus();
        var list = document.querySelectorAll(this._tabbableElements);
        var focused = OSF.OUtil.focusToFirstTabbable(list, false);
        if (!focused) {
            window.blur();
            this._webAppState.focused = false;
            if (this._webAppState.clientEndPoint) {
                this._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [this._webAppState.id, OSF.AgaveHostAction.ExitNoFocusable]);
            }
        }
    }
};
OSF.Messaging.sendMessage = function OSF_Messaging$sendMessage(message) {
    OSF.getClientEndPoint().invoke("ContextActivationManager_notifyHost", null, [OSF._OfficeAppFactory.getWebAppState().id, OSF.AgaveHostAction.SendMessage, message]);
};
OSF.ExtensionLifeCycle.launchExtensionComponent = function OSF_ExtensionLifeCycle$launchExtensionComponent(params) {
    OSF.getClientEndPoint().invoke("ContextActivationManager_notifyHost", null, [OSF._OfficeAppFactory.getWebAppState().id, OSF.AgaveHostAction.LaunchExtensionComponent, params]);
};
OSF.ExtensionLifeCycle.stopExtensionComponent = function OSF_ExtensionLifeCycle$stopExtensionComponent(params) {
    OSF.getClientEndPoint().invoke("ContextActivationManager_notifyHost", null, [OSF._OfficeAppFactory.getWebAppState().id, OSF.AgaveHostAction.StopExtensionComponent, params]);
};
OSF.ExtensionLifeCycle.restartExtensionComponent = function OSF_ExtensionLifeCycle$restartExtensionComponent(params) {
    OSF.getClientEndPoint().invoke("ContextActivationManager_notifyHost", null, [OSF._OfficeAppFactory.getWebAppState().id, OSF.AgaveHostAction.RestartExtensionComponent, params]);
};
OSF.TaskPaneAction.enableHeaderButton = function OSF_TaskPaneAction$enableHeaderButton(params) {
    OSF.getClientEndPoint().invoke("ContextActivationManager_notifyHost", null, [OSF._OfficeAppFactory.getWebAppState().id, OSF.AgaveHostAction.EnableTaskPaneHeaderButton, params]);
};
OSF.TaskPaneAction.disableHeaderButton = function OSF_TaskPaneAction$disableHeaderButton() {
    OSF.getClientEndPoint().invoke("ContextActivationManager_notifyHost", null, [OSF._OfficeAppFactory.getWebAppState().id, OSF.AgaveHostAction.DisableTaskPaneHeaderButton]);
};
OSF.RibbonGallery.refreshRibbon = function OSF_RibbonGallery$refreshRibbon(params) {
    OSF.getClientEndPoint().invoke("ContextActivationManager_notifyHost", null, [OSF._OfficeAppFactory.getWebAppState().id, OSF.AgaveHostAction.RefreshRibbonGallery, params]);
};
(function () {
    var checkScriptOverride = function OSF$checkScriptOverride() {
        var postScriptOverrideCheckAction = function OSF$postScriptOverrideCheckAction(customizedScriptPath) {
            if (customizedScriptPath) {
                OSF.OUtil.loadScript(customizedScriptPath, function () {
                    OsfMsAjaxFactory.msAjaxDebug.trace("loaded customized script:" + customizedScriptPath);
                });
            }
        };
        var conversationID, webAppUrl, items;
        var clientEndPoint = null;
        var xdmInfoValue = OSF.OUtil.parseXdmInfo();
        if (xdmInfoValue) {
            items = OSF.OUtil.getInfoItems(xdmInfoValue);
            if (items && items.length >= 3) {
                conversationID = items[0];
                webAppUrl = items[2];
                var serializerVersion = OSF.OUtil.parseSerializerVersionWithGivenFragment(false, OSF._OfficeAppFactory.getWindowLocationHash());
                if (isNaN(serializerVersion) && OSF._OfficeAppFactory.getWindowName) {
                    serializerVersion = OSF.OUtil.parseSerializerVersionFromWindowName(false, OSF._OfficeAppFactory.getWindowName());
                }
                clientEndPoint = Microsoft.Office.Common.XdmCommunicationManager.connect(conversationID, window.parent, webAppUrl, serializerVersion);
            }
        }
        var customizedScriptPath = null;
        if (!clientEndPoint) {
            try {
                if (window.external && typeof window.external.getCustomizedScriptPath !== 'undefined') {
                    customizedScriptPath = window.external.getCustomizedScriptPath();
                }
            }
            catch (ex) {
                OsfMsAjaxFactory.msAjaxDebug.trace("no script override through window.external.");
            }
            postScriptOverrideCheckAction(customizedScriptPath);
        }
    };
    var requiresMsAjax = true;
    if (requiresMsAjax && !OsfMsAjaxFactory.isMsAjaxLoaded()) {
        if (!(OSF._OfficeAppFactory && OSF._OfficeAppFactory && OSF._OfficeAppFactory.getLoadScriptHelper && OSF._OfficeAppFactory.getLoadScriptHelper().isScriptLoading(OSF.ConstantNames.MicrosoftAjaxId))) {
            OsfMsAjaxFactory.loadMsAjaxFull(function OSF$loadMSAjaxCallback() {
                if (OsfMsAjaxFactory.isMsAjaxLoaded()) {
                    checkScriptOverride();
                }
                else {
                    throw 'Not able to load MicrosoftAjax.js.';
                }
            });
        }
        else {
            OSF._OfficeAppFactory.getLoadScriptHelper().waitForScripts([OSF.ConstantNames.MicrosoftAjaxId], checkScriptOverride);
        }
    }
    else {
        checkScriptOverride();
    }
})();
var OSFLog;
(function (OSFLog) {
    var BaseUsageData = (function () {
        function BaseUsageData(table) {
            this._table = table;
            this._fields = {};
        }
        Object.defineProperty(BaseUsageData.prototype, "Fields", {
            get: function () {
                return this._fields;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(BaseUsageData.prototype, "Table", {
            get: function () {
                return this._table;
            },
            enumerable: true,
            configurable: true
        });
        BaseUsageData.prototype.SerializeFields = function () {
        };
        BaseUsageData.prototype.SetSerializedField = function (key, value) {
            if (typeof (value) !== "undefined" && value !== null) {
                this._serializedFields[key] = value.toString();
            }
        };
        BaseUsageData.prototype.SerializeRow = function () {
            this._serializedFields = {};
            this.SetSerializedField("Table", this._table);
            this.SerializeFields();
            return JSON.stringify(this._serializedFields);
        };
        return BaseUsageData;
    }());
    OSFLog.BaseUsageData = BaseUsageData;
    var AppActivatedUsageData = (function (_super) {
        __extends(AppActivatedUsageData, _super);
        function AppActivatedUsageData() {
            return _super.call(this, "AppActivated") || this;
        }
        Object.defineProperty(AppActivatedUsageData.prototype, "CorrelationId", {
            get: function () { return this.Fields["CorrelationId"]; },
            set: function (value) { this.Fields["CorrelationId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "SessionId", {
            get: function () { return this.Fields["SessionId"]; },
            set: function (value) { this.Fields["SessionId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "AppId", {
            get: function () { return this.Fields["AppId"]; },
            set: function (value) { this.Fields["AppId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "AppInstanceId", {
            get: function () { return this.Fields["AppInstanceId"]; },
            set: function (value) { this.Fields["AppInstanceId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "AppURL", {
            get: function () { return this.Fields["AppURL"]; },
            set: function (value) { this.Fields["AppURL"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "AssetId", {
            get: function () { return this.Fields["AssetId"]; },
            set: function (value) { this.Fields["AssetId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "Browser", {
            get: function () { return this.Fields["Browser"]; },
            set: function (value) { this.Fields["Browser"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "UserId", {
            get: function () { return this.Fields["UserId"]; },
            set: function (value) { this.Fields["UserId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "Host", {
            get: function () { return this.Fields["Host"]; },
            set: function (value) { this.Fields["Host"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "HostVersion", {
            get: function () { return this.Fields["HostVersion"]; },
            set: function (value) { this.Fields["HostVersion"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "ClientId", {
            get: function () { return this.Fields["ClientId"]; },
            set: function (value) { this.Fields["ClientId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "AppSizeWidth", {
            get: function () { return this.Fields["AppSizeWidth"]; },
            set: function (value) { this.Fields["AppSizeWidth"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "AppSizeHeight", {
            get: function () { return this.Fields["AppSizeHeight"]; },
            set: function (value) { this.Fields["AppSizeHeight"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "Message", {
            get: function () { return this.Fields["Message"]; },
            set: function (value) { this.Fields["Message"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "DocUrl", {
            get: function () { return this.Fields["DocUrl"]; },
            set: function (value) { this.Fields["DocUrl"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "OfficeJSVersion", {
            get: function () { return this.Fields["OfficeJSVersion"]; },
            set: function (value) { this.Fields["OfficeJSVersion"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "HostJSVersion", {
            get: function () { return this.Fields["HostJSVersion"]; },
            set: function (value) { this.Fields["HostJSVersion"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "WacHostEnvironment", {
            get: function () { return this.Fields["WacHostEnvironment"]; },
            set: function (value) { this.Fields["WacHostEnvironment"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "IsFromWacAutomation", {
            get: function () { return this.Fields["IsFromWacAutomation"]; },
            set: function (value) { this.Fields["IsFromWacAutomation"] = value; },
            enumerable: true,
            configurable: true
        });
        AppActivatedUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("CorrelationId", this.CorrelationId);
            this.SetSerializedField("SessionId", this.SessionId);
            this.SetSerializedField("AppId", this.AppId);
            this.SetSerializedField("AppInstanceId", this.AppInstanceId);
            this.SetSerializedField("AppURL", this.AppURL);
            this.SetSerializedField("AssetId", this.AssetId);
            this.SetSerializedField("Browser", this.Browser);
            this.SetSerializedField("UserId", this.UserId);
            this.SetSerializedField("Host", this.Host);
            this.SetSerializedField("HostVersion", this.HostVersion);
            this.SetSerializedField("ClientId", this.ClientId);
            this.SetSerializedField("AppSizeWidth", this.AppSizeWidth);
            this.SetSerializedField("AppSizeHeight", this.AppSizeHeight);
            this.SetSerializedField("Message", this.Message);
            this.SetSerializedField("DocUrl", this.DocUrl);
            this.SetSerializedField("OfficeJSVersion", this.OfficeJSVersion);
            this.SetSerializedField("HostJSVersion", this.HostJSVersion);
            this.SetSerializedField("WacHostEnvironment", this.WacHostEnvironment);
            this.SetSerializedField("IsFromWacAutomation", this.IsFromWacAutomation);
        };
        return AppActivatedUsageData;
    }(BaseUsageData));
    OSFLog.AppActivatedUsageData = AppActivatedUsageData;
    var ScriptLoadUsageData = (function (_super) {
        __extends(ScriptLoadUsageData, _super);
        function ScriptLoadUsageData() {
            return _super.call(this, "ScriptLoad") || this;
        }
        Object.defineProperty(ScriptLoadUsageData.prototype, "CorrelationId", {
            get: function () { return this.Fields["CorrelationId"]; },
            set: function (value) { this.Fields["CorrelationId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ScriptLoadUsageData.prototype, "SessionId", {
            get: function () { return this.Fields["SessionId"]; },
            set: function (value) { this.Fields["SessionId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ScriptLoadUsageData.prototype, "ScriptId", {
            get: function () { return this.Fields["ScriptId"]; },
            set: function (value) { this.Fields["ScriptId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ScriptLoadUsageData.prototype, "StartTime", {
            get: function () { return this.Fields["StartTime"]; },
            set: function (value) { this.Fields["StartTime"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ScriptLoadUsageData.prototype, "ResponseTime", {
            get: function () { return this.Fields["ResponseTime"]; },
            set: function (value) { this.Fields["ResponseTime"] = value; },
            enumerable: true,
            configurable: true
        });
        ScriptLoadUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("CorrelationId", this.CorrelationId);
            this.SetSerializedField("SessionId", this.SessionId);
            this.SetSerializedField("ScriptId", this.ScriptId);
            this.SetSerializedField("StartTime", this.StartTime);
            this.SetSerializedField("ResponseTime", this.ResponseTime);
        };
        return ScriptLoadUsageData;
    }(BaseUsageData));
    OSFLog.ScriptLoadUsageData = ScriptLoadUsageData;
    var AppClosedUsageData = (function (_super) {
        __extends(AppClosedUsageData, _super);
        function AppClosedUsageData() {
            return _super.call(this, "AppClosed") || this;
        }
        Object.defineProperty(AppClosedUsageData.prototype, "CorrelationId", {
            get: function () { return this.Fields["CorrelationId"]; },
            set: function (value) { this.Fields["CorrelationId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppClosedUsageData.prototype, "SessionId", {
            get: function () { return this.Fields["SessionId"]; },
            set: function (value) { this.Fields["SessionId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppClosedUsageData.prototype, "FocusTime", {
            get: function () { return this.Fields["FocusTime"]; },
            set: function (value) { this.Fields["FocusTime"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppClosedUsageData.prototype, "AppSizeFinalWidth", {
            get: function () { return this.Fields["AppSizeFinalWidth"]; },
            set: function (value) { this.Fields["AppSizeFinalWidth"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppClosedUsageData.prototype, "AppSizeFinalHeight", {
            get: function () { return this.Fields["AppSizeFinalHeight"]; },
            set: function (value) { this.Fields["AppSizeFinalHeight"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppClosedUsageData.prototype, "OpenTime", {
            get: function () { return this.Fields["OpenTime"]; },
            set: function (value) { this.Fields["OpenTime"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppClosedUsageData.prototype, "CloseMethod", {
            get: function () { return this.Fields["CloseMethod"]; },
            set: function (value) { this.Fields["CloseMethod"] = value; },
            enumerable: true,
            configurable: true
        });
        AppClosedUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("CorrelationId", this.CorrelationId);
            this.SetSerializedField("SessionId", this.SessionId);
            this.SetSerializedField("FocusTime", this.FocusTime);
            this.SetSerializedField("AppSizeFinalWidth", this.AppSizeFinalWidth);
            this.SetSerializedField("AppSizeFinalHeight", this.AppSizeFinalHeight);
            this.SetSerializedField("OpenTime", this.OpenTime);
            this.SetSerializedField("CloseMethod", this.CloseMethod);
        };
        return AppClosedUsageData;
    }(BaseUsageData));
    OSFLog.AppClosedUsageData = AppClosedUsageData;
    var APIUsageUsageData = (function (_super) {
        __extends(APIUsageUsageData, _super);
        function APIUsageUsageData() {
            return _super.call(this, "APIUsage") || this;
        }
        Object.defineProperty(APIUsageUsageData.prototype, "CorrelationId", {
            get: function () { return this.Fields["CorrelationId"]; },
            set: function (value) { this.Fields["CorrelationId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(APIUsageUsageData.prototype, "SessionId", {
            get: function () { return this.Fields["SessionId"]; },
            set: function (value) { this.Fields["SessionId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(APIUsageUsageData.prototype, "APIType", {
            get: function () { return this.Fields["APIType"]; },
            set: function (value) { this.Fields["APIType"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(APIUsageUsageData.prototype, "APIID", {
            get: function () { return this.Fields["APIID"]; },
            set: function (value) { this.Fields["APIID"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(APIUsageUsageData.prototype, "Parameters", {
            get: function () { return this.Fields["Parameters"]; },
            set: function (value) { this.Fields["Parameters"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(APIUsageUsageData.prototype, "ResponseTime", {
            get: function () { return this.Fields["ResponseTime"]; },
            set: function (value) { this.Fields["ResponseTime"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(APIUsageUsageData.prototype, "ErrorType", {
            get: function () { return this.Fields["ErrorType"]; },
            set: function (value) { this.Fields["ErrorType"] = value; },
            enumerable: true,
            configurable: true
        });
        APIUsageUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("CorrelationId", this.CorrelationId);
            this.SetSerializedField("SessionId", this.SessionId);
            this.SetSerializedField("APIType", this.APIType);
            this.SetSerializedField("APIID", this.APIID);
            this.SetSerializedField("Parameters", this.Parameters);
            this.SetSerializedField("ResponseTime", this.ResponseTime);
            this.SetSerializedField("ErrorType", this.ErrorType);
        };
        return APIUsageUsageData;
    }(BaseUsageData));
    OSFLog.APIUsageUsageData = APIUsageUsageData;
    var AppInitializationUsageData = (function (_super) {
        __extends(AppInitializationUsageData, _super);
        function AppInitializationUsageData() {
            return _super.call(this, "AppInitialization") || this;
        }
        Object.defineProperty(AppInitializationUsageData.prototype, "CorrelationId", {
            get: function () { return this.Fields["CorrelationId"]; },
            set: function (value) { this.Fields["CorrelationId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppInitializationUsageData.prototype, "SessionId", {
            get: function () { return this.Fields["SessionId"]; },
            set: function (value) { this.Fields["SessionId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppInitializationUsageData.prototype, "SuccessCode", {
            get: function () { return this.Fields["SuccessCode"]; },
            set: function (value) { this.Fields["SuccessCode"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppInitializationUsageData.prototype, "Message", {
            get: function () { return this.Fields["Message"]; },
            set: function (value) { this.Fields["Message"] = value; },
            enumerable: true,
            configurable: true
        });
        AppInitializationUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("CorrelationId", this.CorrelationId);
            this.SetSerializedField("SessionId", this.SessionId);
            this.SetSerializedField("SuccessCode", this.SuccessCode);
            this.SetSerializedField("Message", this.Message);
        };
        return AppInitializationUsageData;
    }(BaseUsageData));
    OSFLog.AppInitializationUsageData = AppInitializationUsageData;
    var CheckWACHostUsageData = (function (_super) {
        __extends(CheckWACHostUsageData, _super);
        function CheckWACHostUsageData() {
            return _super.call(this, "CheckWACHost") || this;
        }
        Object.defineProperty(CheckWACHostUsageData.prototype, "isWacKnownHost", {
            get: function () { return this.Fields["isWacKnownHost"]; },
            set: function (value) { this.Fields["isWacKnownHost"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CheckWACHostUsageData.prototype, "instanceId", {
            get: function () { return this.Fields["instanceId"]; },
            set: function (value) { this.Fields["instanceId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CheckWACHostUsageData.prototype, "hostType", {
            get: function () { return this.Fields["hostType"]; },
            set: function (value) { this.Fields["hostType"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CheckWACHostUsageData.prototype, "hostPlatform", {
            get: function () { return this.Fields["hostPlatform"]; },
            set: function (value) { this.Fields["hostPlatform"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CheckWACHostUsageData.prototype, "wacDomain", {
            get: function () { return this.Fields["wacDomain"]; },
            set: function (value) { this.Fields["wacDomain"] = value; },
            enumerable: true,
            configurable: true
        });
        CheckWACHostUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("isWacKnownHost", this.isWacKnownHost);
            this.SetSerializedField("instanceId", this.instanceId);
            this.SetSerializedField("hostType", this.hostType);
            this.SetSerializedField("hostPlatform", this.hostPlatform);
            this.SetSerializedField("wacDomain", this.wacDomain);
        };
        return CheckWACHostUsageData;
    }(BaseUsageData));
    OSFLog.CheckWACHostUsageData = CheckWACHostUsageData;
})(OSFLog || (OSFLog = {}));
var Logger;
(function (Logger) {
    "use strict";
    var TraceLevel;
    (function (TraceLevel) {
        TraceLevel[TraceLevel["info"] = 0] = "info";
        TraceLevel[TraceLevel["warning"] = 1] = "warning";
        TraceLevel[TraceLevel["error"] = 2] = "error";
    })(TraceLevel = Logger.TraceLevel || (Logger.TraceLevel = {}));
    var SendFlag;
    (function (SendFlag) {
        SendFlag[SendFlag["none"] = 0] = "none";
        SendFlag[SendFlag["flush"] = 1] = "flush";
    })(SendFlag = Logger.SendFlag || (Logger.SendFlag = {}));
    function allowUploadingData() {
    }
    Logger.allowUploadingData = allowUploadingData;
    function sendLog(traceLevel, message, flag) {
    }
    Logger.sendLog = sendLog;
    function creatULSEndpoint() {
        try {
            return new ULSEndpointProxy();
        }
        catch (e) {
            return null;
        }
    }
    var ULSEndpointProxy = (function () {
        function ULSEndpointProxy() {
        }
        ULSEndpointProxy.prototype.writeLog = function (log) {
        };
        ULSEndpointProxy.prototype.loadProxyFrame = function () {
        };
        return ULSEndpointProxy;
    }());
    if (!OSF.Logger) {
        OSF.Logger = Logger;
    }
    Logger.ulsEndpoint = creatULSEndpoint();
})(Logger || (Logger = {}));
var OSFAriaLogger;
(function (OSFAriaLogger) {
    var TelemetryEventAppActivated = { name: "AppActivated", enabled: true, critical: true, points: [
            { name: "Browser", type: "string" },
            { name: "Message", type: "string" },
            { name: "Host", type: "string" },
            { name: "AppSizeWidth", type: "int64" },
            { name: "AppSizeHeight", type: "int64" },
            { name: "IsFromWacAutomation", type: "string" },
        ] };
    var TelemetryEventScriptLoad = { name: "ScriptLoad", enabled: true, critical: false, points: [
            { name: "ScriptId", type: "string" },
            { name: "StartTime", type: "double" },
            { name: "ResponseTime", type: "double" },
        ] };
    var enableAPIUsage = shouldAPIUsageBeEnabled();
    var TelemetryEventApiUsage = { name: "APIUsage", enabled: enableAPIUsage, critical: false, points: [
            { name: "APIType", type: "string" },
            { name: "APIID", type: "int64" },
            { name: "Parameters", type: "string" },
            { name: "ResponseTime", type: "int64" },
            { name: "ErrorType", type: "int64" },
        ] };
    var TelemetryEventAppInitialization = { name: "AppInitialization", enabled: true, critical: false, points: [
            { name: "SuccessCode", type: "int64" },
            { name: "Message", type: "string" },
        ] };
    var TelemetryEventAppClosed = { name: "AppClosed", enabled: true, critical: false, points: [
            { name: "FocusTime", type: "int64" },
            { name: "AppSizeFinalWidth", type: "int64" },
            { name: "AppSizeFinalHeight", type: "int64" },
            { name: "OpenTime", type: "int64" },
        ] };
    var TelemetryEventCheckWACHost = { name: "CheckWACHost", enabled: true, critical: false, points: [
            { name: "isWacKnownHost", type: "int64" },
            { name: "solutionId", type: "string" },
            { name: "hostType", type: "string" },
            { name: "hostPlatform", type: "string" },
            { name: "correlationId", type: "string" },
        ] };
    var TelemetryEvents = [
        TelemetryEventAppActivated,
        TelemetryEventScriptLoad,
        TelemetryEventApiUsage,
        TelemetryEventAppInitialization,
        TelemetryEventAppClosed,
        TelemetryEventCheckWACHost,
    ];
    function createDataField(value, point) {
        var key = point.rename === undefined ? point.name : point.rename;
        var type = point.type;
        var field = undefined;
        switch (type) {
            case "string":
                field = oteljs.makeStringDataField(key, value);
                break;
            case "double":
                if (typeof value === "string") {
                    value = parseFloat(value);
                }
                field = oteljs.makeDoubleDataField(key, value);
                break;
            case "int64":
                if (typeof value === "string") {
                    value = parseInt(value);
                }
                field = oteljs.makeInt64DataField(key, value);
                break;
            case "boolean":
                if (typeof value === "string") {
                    value = value === "true";
                }
                field = oteljs.makeBooleanDataField(key, value);
                break;
        }
        return field;
    }
    function getEventDefinition(eventName) {
        for (var _i = 0, TelemetryEvents_1 = TelemetryEvents; _i < TelemetryEvents_1.length; _i++) {
            var event_1 = TelemetryEvents_1[_i];
            if (event_1.name === eventName) {
                return event_1;
            }
        }
        return undefined;
    }
    function eventEnabled(eventName) {
        var eventDefinition = getEventDefinition(eventName);
        if (eventDefinition === undefined) {
            return false;
        }
        return eventDefinition.enabled;
    }
    function shouldAPIUsageBeEnabled() {
        if (!OSF._OfficeAppFactory || !OSF._OfficeAppFactory.getHostInfo) {
            return false;
        }
        var hostInfo = OSF._OfficeAppFactory.getHostInfo();
        if (!hostInfo) {
            return false;
        }
        switch (hostInfo["hostType"]) {
            case "outlook":
                switch (hostInfo["hostPlatform"]) {
                    case "mac":
                    case "web":
                        return true;
                    default:
                        return false;
                }
            default:
                return false;
        }
    }
    function generateTelemetryEvent(eventName, telemetryData) {
        var eventDefinition = getEventDefinition(eventName);
        if (eventDefinition === undefined) {
            return undefined;
        }
        var dataFields = [];
        for (var _i = 0, _a = eventDefinition.points; _i < _a.length; _i++) {
            var point = _a[_i];
            var key = point.name;
            var value = telemetryData[key];
            if (value === undefined) {
                continue;
            }
            var field = createDataField(value, point);
            if (field !== undefined) {
                dataFields.push(field);
            }
        }
        var flags = { dataCategories: oteljs.DataCategories.ProductServiceUsage };
        if (eventDefinition.critical) {
            flags.samplingPolicy = oteljs.SamplingPolicy.CriticalBusinessImpact;
        }
        flags.diagnosticLevel = oteljs.DiagnosticLevel.NecessaryServiceDataEvent;
        var eventNameFull = "Office.Extensibility.OfficeJs." + eventName + "X";
        var event = { eventName: eventNameFull, dataFields: dataFields, eventFlags: flags };
        return event;
    }
    function sendOtelTelemetryEvent(eventName, telemetryData) {
        if (eventEnabled(eventName)) {
            if (typeof OTel !== "undefined") {
                OTel.OTelLogger.onTelemetryLoaded(function () {
                    var event = generateTelemetryEvent(eventName, telemetryData);
                    if (event === undefined) {
                        return;
                    }
                    Microsoft.Office.WebExtension.sendTelemetryEvent(event);
                });
            }
        }
    }
    var AriaLogger = (function () {
        function AriaLogger() {
        }
        AriaLogger.prototype.getAriaCDNLocation = function () {
            return (OSF._OfficeAppFactory.getLoadScriptHelper().getOfficeJsBasePath() + "ariatelemetry/aria-web-telemetry.js");
        };
        AriaLogger.getInstance = function () {
            if (AriaLogger.AriaLoggerObj === undefined) {
                AriaLogger.AriaLoggerObj = new AriaLogger();
            }
            return AriaLogger.AriaLoggerObj;
        };
        AriaLogger.prototype.isIUsageData = function (arg) {
            return arg["Fields"] !== undefined;
        };
        AriaLogger.prototype.shouldSendDirectToAria = function (flavor, version) {
            var BASE10 = 10;
            var MAX_VERSION_WIN32 = [16, 0, 11601];
            var MAX_VERSION_MAC = [16, 28];
            var max_version;
            if (!flavor) {
                return false;
            }
            else if (flavor.toLowerCase() === "win32") {
                max_version = MAX_VERSION_WIN32;
            }
            else if (flavor.toLowerCase() === "mac") {
                max_version = MAX_VERSION_MAC;
            }
            else {
                return true;
            }
            if (!version) {
                return false;
            }
            var versionTokens = version.split('.');
            for (var i = 0; i < max_version.length && i < versionTokens.length; i++) {
                var versionToken = parseInt(versionTokens[i], BASE10);
                if (isNaN(versionToken)) {
                    return false;
                }
                if (versionToken < max_version[i]) {
                    return true;
                }
                if (versionToken > max_version[i]) {
                    return false;
                }
            }
            return false;
        };
        AriaLogger.prototype.isDirectToAriaEnabled = function () {
            if (this.EnableDirectToAria === undefined || this.EnableDirectToAria === null) {
                var flavor = void 0;
                var version = void 0;
                if (OSF._OfficeAppFactory && OSF._OfficeAppFactory.getHostInfo) {
                    flavor = OSF._OfficeAppFactory.getHostInfo()["hostPlatform"];
                }
                if (window.external && typeof window.external.GetContext !== "undefined" && typeof window.external.GetContext().GetHostFullVersion !== "undefined") {
                    version = window.external.GetContext().GetHostFullVersion();
                }
                this.EnableDirectToAria = this.shouldSendDirectToAria(flavor, version);
            }
            return this.EnableDirectToAria;
        };
        AriaLogger.prototype.sendTelemetry = function (tableName, telemetryData) {
            var startAfterMs = 1000;
            var sendAriaEnabled = AriaLogger.EnableSendingTelemetryWithLegacyAria && this.isDirectToAriaEnabled();
            if (sendAriaEnabled) {
                OSF.OUtil.loadScript(this.getAriaCDNLocation(), function () {
                    try {
                        if (!this.ALogger) {
                            var OfficeExtensibilityTenantID = "db334b301e7b474db5e0f02f07c51a47-a1b5bc36-1bbe-482f-a64a-c2d9cb606706-7439";
                            this.ALogger = AWTLogManager.initialize(OfficeExtensibilityTenantID);
                        }
                        var eventProperties = new AWTEventProperties();
                        eventProperties.setName("Office.Extensibility.OfficeJS." + tableName);
                        for (var key in telemetryData) {
                            if (key.toLowerCase() !== "table") {
                                eventProperties.setProperty(key, telemetryData[key]);
                            }
                        }
                        var today = new Date();
                        eventProperties.setProperty("Date", today.toISOString());
                        this.ALogger.logEvent(eventProperties);
                    }
                    catch (e) {
                    }
                }, startAfterMs);
            }
            if (AriaLogger.EnableSendingTelemetryWithOTel) {
                sendOtelTelemetryEvent(tableName, telemetryData);
            }
        };
        AriaLogger.prototype.logData = function (data) {
            if (this.isIUsageData(data)) {
                this.sendTelemetry(data["Table"], data["Fields"]);
            }
            else {
                this.sendTelemetry(data["Table"], data);
            }
        };
        AriaLogger.EnableSendingTelemetryWithOTel = true;
        AriaLogger.EnableSendingTelemetryWithLegacyAria = false;
        return AriaLogger;
    }());
    OSFAriaLogger.AriaLogger = AriaLogger;
})(OSFAriaLogger || (OSFAriaLogger = {}));
var OSFAppTelemetry;
(function (OSFAppTelemetry) {
    "use strict";
    var appInfo;
    var sessionId = OSF.OUtil.Guid.generateNewGuid();
    var osfControlAppCorrelationId = "";
    var omexDomainRegex = new RegExp("^https?://store\\.office(ppe|-int)?\\.com/", "i");
    var privateAddinId = "PRIVATE";
    OSFAppTelemetry.enableTelemetry = true;
    ;
    var AppInfo = (function () {
        function AppInfo() {
        }
        return AppInfo;
    }());
    OSFAppTelemetry.AppInfo = AppInfo;
    var Event = (function () {
        function Event(name, handler) {
            this.name = name;
            this.handler = handler;
        }
        return Event;
    }());
    var AppStorage = (function () {
        function AppStorage() {
            this.clientIDKey = "Office API client";
            this.logIdSetKey = "Office App Log Id Set";
        }
        AppStorage.prototype.getClientId = function () {
            var clientId = this.getValue(this.clientIDKey);
            if (!clientId || clientId.length <= 0 || clientId.length > 40) {
                clientId = OSF.OUtil.Guid.generateNewGuid();
                this.setValue(this.clientIDKey, clientId);
            }
            return clientId;
        };
        AppStorage.prototype.saveLog = function (logId, log) {
            var logIdSet = this.getValue(this.logIdSetKey);
            logIdSet = ((logIdSet && logIdSet.length > 0) ? (logIdSet + ";") : "") + logId;
            this.setValue(this.logIdSetKey, logIdSet);
            this.setValue(logId, log);
        };
        AppStorage.prototype.enumerateLog = function (callback, clean) {
            var logIdSet = this.getValue(this.logIdSetKey);
            if (logIdSet) {
                var ids = logIdSet.split(";");
                for (var id in ids) {
                    var logId = ids[id];
                    var log = this.getValue(logId);
                    if (log) {
                        if (callback) {
                            callback(logId, log);
                        }
                        if (clean) {
                            this.remove(logId);
                        }
                    }
                }
                if (clean) {
                    this.remove(this.logIdSetKey);
                }
            }
        };
        AppStorage.prototype.getValue = function (key) {
            var osfLocalStorage = OSF.OUtil.getLocalStorage();
            var value = "";
            if (osfLocalStorage) {
                value = osfLocalStorage.getItem(key);
            }
            return value;
        };
        AppStorage.prototype.setValue = function (key, value) {
            var osfLocalStorage = OSF.OUtil.getLocalStorage();
            if (osfLocalStorage) {
                osfLocalStorage.setItem(key, value);
            }
        };
        AppStorage.prototype.remove = function (key) {
            var osfLocalStorage = OSF.OUtil.getLocalStorage();
            if (osfLocalStorage) {
                try {
                    osfLocalStorage.removeItem(key);
                }
                catch (ex) {
                }
            }
        };
        return AppStorage;
    }());
    var AppLogger = (function () {
        function AppLogger() {
        }
        AppLogger.prototype.LogData = function (data) {
            if (!OSFAppTelemetry.enableTelemetry) {
                return;
            }
            try {
                OSFAriaLogger.AriaLogger.getInstance().logData(data);
            }
            catch (e) {
            }
        };
        AppLogger.prototype.LogRawData = function (log) {
            if (!OSFAppTelemetry.enableTelemetry) {
                return;
            }
            try {
                OSFAriaLogger.AriaLogger.getInstance().logData(JSON.parse(log));
            }
            catch (e) {
            }
        };
        return AppLogger;
    }());
    function trimStringToLowerCase(input) {
        if (input) {
            input = input.replace(/[{}]/g, "").toLowerCase();
        }
        return (input || "");
    }
    function initialize(context) {
        if (!OSFAppTelemetry.enableTelemetry) {
            return;
        }
        if (appInfo) {
            return;
        }
        appInfo = new AppInfo();
        if (context.get_hostFullVersion()) {
            appInfo.hostVersion = context.get_hostFullVersion();
        }
        else {
            appInfo.hostVersion = context.get_appVersion();
        }
        appInfo.appId = canSendAddinId() ? context.get_id() : privateAddinId;
        appInfo.browser = window.navigator.userAgent;
        appInfo.correlationId = trimStringToLowerCase(context.get_correlationId());
        appInfo.clientId = (new AppStorage()).getClientId();
        appInfo.appInstanceId = context.get_appInstanceId();
        if (appInfo.appInstanceId) {
            appInfo.appInstanceId = trimStringToLowerCase(appInfo.appInstanceId);
            appInfo.appInstanceId = getCompliantAppInstanceId(context.get_id(), appInfo.appInstanceId);
        }
        appInfo.message = context.get_hostCustomMessage();
        appInfo.officeJSVersion = OSF.ConstantNames.FileVersion;
        appInfo.hostJSVersion = "16.0.14516.10000";
        if (context._wacHostEnvironment) {
            appInfo.wacHostEnvironment = context._wacHostEnvironment;
        }
        if (context._isFromWacAutomation !== undefined && context._isFromWacAutomation !== null) {
            appInfo.isFromWacAutomation = context._isFromWacAutomation.toString().toLowerCase();
        }
        var docUrl = context.get_docUrl();
        appInfo.docUrl = omexDomainRegex.test(docUrl) ? docUrl : "";
        var url = location.href;
        if (url) {
            url = url.split("?")[0].split("#")[0];
        }
        appInfo.appURL = "";
        (function getUserIdAndAssetIdFromToken(token, appInfo) {
            var xmlContent;
            var parser;
            var xmlDoc;
            appInfo.assetId = "";
            appInfo.userId = "";
            try {
                xmlContent = decodeURIComponent(token);
                parser = new DOMParser();
                xmlDoc = parser.parseFromString(xmlContent, "text/xml");
                var cidNode = xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("cid");
                var oidNode = xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("oid");
                if (cidNode && cidNode.nodeValue) {
                    appInfo.userId = cidNode.nodeValue;
                }
                else if (oidNode && oidNode.nodeValue) {
                    appInfo.userId = oidNode.nodeValue;
                }
                appInfo.assetId = xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("aid").nodeValue;
            }
            catch (e) {
            }
            finally {
                xmlContent = null;
                xmlDoc = null;
                parser = null;
            }
        })(context.get_eToken(), appInfo);
        appInfo.sessionId = sessionId;
        if (typeof OTel !== "undefined") {
            OTel.OTelLogger.initialize(appInfo);
        }
        (function handleLifecycle() {
            var startTime = new Date();
            var lastFocus = null;
            var focusTime = 0;
            var finished = false;
            var adjustFocusTime = function () {
                if (document.hasFocus()) {
                    if (lastFocus == null) {
                        lastFocus = new Date();
                    }
                }
                else if (lastFocus) {
                    focusTime += Math.abs((new Date()).getTime() - lastFocus.getTime());
                    lastFocus = null;
                }
            };
            var eventList = [];
            eventList.push(new Event("focus", adjustFocusTime));
            eventList.push(new Event("blur", adjustFocusTime));
            eventList.push(new Event("focusout", adjustFocusTime));
            eventList.push(new Event("focusin", adjustFocusTime));
            var exitFunction = function () {
                for (var i = 0; i < eventList.length; i++) {
                    OSF.OUtil.removeEventListener(window, eventList[i].name, eventList[i].handler);
                }
                eventList.length = 0;
                if (!finished) {
                    if (document.hasFocus() && lastFocus) {
                        focusTime += Math.abs((new Date()).getTime() - lastFocus.getTime());
                        lastFocus = null;
                    }
                    OSFAppTelemetry.onAppClosed(Math.abs((new Date()).getTime() - startTime.getTime()), focusTime);
                    finished = true;
                }
            };
            eventList.push(new Event("beforeunload", exitFunction));
            eventList.push(new Event("unload", exitFunction));
            for (var i = 0; i < eventList.length; i++) {
                OSF.OUtil.addEventListener(window, eventList[i].name, eventList[i].handler);
            }
            adjustFocusTime();
        })();
        OSFAppTelemetry.onAppActivated();
    }
    OSFAppTelemetry.initialize = initialize;
    function onAppActivated() {
        if (!appInfo) {
            return;
        }
        (new AppStorage()).enumerateLog(function (id, log) { return (new AppLogger()).LogRawData(log); }, true);
        var data = new OSFLog.AppActivatedUsageData();
        data.SessionId = sessionId;
        data.AppId = appInfo.appId;
        data.AssetId = appInfo.assetId;
        data.AppURL = "";
        data.UserId = "";
        data.ClientId = appInfo.clientId;
        data.Browser = appInfo.browser;
        data.HostVersion = appInfo.hostVersion;
        data.CorrelationId = trimStringToLowerCase(appInfo.correlationId);
        data.AppSizeWidth = window.innerWidth;
        data.AppSizeHeight = window.innerHeight;
        data.AppInstanceId = appInfo.appInstanceId;
        data.Message = appInfo.message;
        data.DocUrl = appInfo.docUrl;
        data.OfficeJSVersion = appInfo.officeJSVersion;
        data.HostJSVersion = appInfo.hostJSVersion;
        if (appInfo.wacHostEnvironment) {
            data.WacHostEnvironment = appInfo.wacHostEnvironment;
        }
        if (appInfo.isFromWacAutomation !== undefined && appInfo.isFromWacAutomation !== null) {
            data.IsFromWacAutomation = appInfo.isFromWacAutomation;
        }
        (new AppLogger()).LogData(data);
    }
    OSFAppTelemetry.onAppActivated = onAppActivated;
    function onScriptDone(scriptId, msStartTime, msResponseTime, appCorrelationId) {
        var data = new OSFLog.ScriptLoadUsageData();
        data.CorrelationId = trimStringToLowerCase(appCorrelationId);
        data.SessionId = sessionId;
        data.ScriptId = scriptId;
        data.StartTime = msStartTime;
        data.ResponseTime = msResponseTime;
        (new AppLogger()).LogData(data);
    }
    OSFAppTelemetry.onScriptDone = onScriptDone;
    function onCallDone(apiType, id, parameters, msResponseTime, errorType) {
        if (!appInfo) {
            return;
        }
        if (!isAllowedHost() || !isAPIUsageEnabledDispId(id, apiType)) {
            return;
        }
        var data = new OSFLog.APIUsageUsageData();
        data.CorrelationId = trimStringToLowerCase(osfControlAppCorrelationId);
        data.SessionId = sessionId;
        data.APIType = apiType;
        data.APIID = id;
        data.Parameters = parameters;
        data.ResponseTime = msResponseTime;
        data.ErrorType = errorType;
        (new AppLogger()).LogData(data);
    }
    OSFAppTelemetry.onCallDone = onCallDone;
    ;
    function onMethodDone(id, args, msResponseTime, errorType) {
        var parameters = null;
        if (args) {
            if (typeof args == "number") {
                parameters = String(args);
            }
            else if (typeof args === "object") {
                for (var index in args) {
                    if (parameters !== null) {
                        parameters += ",";
                    }
                    else {
                        parameters = "";
                    }
                    if (typeof args[index] == "number") {
                        parameters += String(args[index]);
                    }
                }
            }
            else {
                parameters = "";
            }
        }
        OSF.AppTelemetry.onCallDone("method", id, parameters, msResponseTime, errorType);
    }
    OSFAppTelemetry.onMethodDone = onMethodDone;
    function onPropertyDone(propertyName, msResponseTime) {
        OSF.AppTelemetry.onCallDone("property", -1, propertyName, msResponseTime);
    }
    OSFAppTelemetry.onPropertyDone = onPropertyDone;
    function onCheckWACHost(isWacKnownHost, instanceId, hostType, hostPlatform, wacDomain) {
        var data = new OSFLog.CheckWACHostUsageData();
        data.isWacKnownHost = isWacKnownHost;
        data.instanceId = instanceId;
        data.hostType = hostType;
        data.hostPlatform = hostPlatform;
        data.wacDomain = "";
        (new AppLogger()).LogData(data);
    }
    OSFAppTelemetry.onCheckWACHost = onCheckWACHost;
    function onEventDone(id, errorType) {
        OSF.AppTelemetry.onCallDone("event", id, null, 0, errorType);
    }
    OSFAppTelemetry.onEventDone = onEventDone;
    function onRegisterDone(register, id, msResponseTime, errorType) {
        OSF.AppTelemetry.onCallDone(register ? "registerevent" : "unregisterevent", id, null, msResponseTime, errorType);
    }
    OSFAppTelemetry.onRegisterDone = onRegisterDone;
    function onAppClosed(openTime, focusTime) {
        if (!appInfo) {
            return;
        }
        var data = new OSFLog.AppClosedUsageData();
        data.CorrelationId = trimStringToLowerCase(osfControlAppCorrelationId);
        data.SessionId = sessionId;
        data.FocusTime = focusTime;
        data.OpenTime = openTime;
        data.AppSizeFinalWidth = window.innerWidth;
        data.AppSizeFinalHeight = window.innerHeight;
        (new AppStorage()).saveLog(sessionId, data.SerializeRow());
    }
    OSFAppTelemetry.onAppClosed = onAppClosed;
    function setOsfControlAppCorrelationId(correlationId) {
        osfControlAppCorrelationId = trimStringToLowerCase(correlationId);
    }
    OSFAppTelemetry.setOsfControlAppCorrelationId = setOsfControlAppCorrelationId;
    function doAppInitializationLogging(isException, message) {
        var data = new OSFLog.AppInitializationUsageData();
        data.CorrelationId = trimStringToLowerCase(osfControlAppCorrelationId);
        data.SessionId = sessionId;
        data.SuccessCode = isException ? 1 : 0;
        data.Message = message;
        (new AppLogger()).LogData(data);
    }
    OSFAppTelemetry.doAppInitializationLogging = doAppInitializationLogging;
    function logAppCommonMessage(message) {
        doAppInitializationLogging(false, message);
    }
    OSFAppTelemetry.logAppCommonMessage = logAppCommonMessage;
    function logAppException(errorMessage) {
        doAppInitializationLogging(true, errorMessage);
    }
    OSFAppTelemetry.logAppException = logAppException;
    function isAllowedHost() {
        if (!OSF._OfficeAppFactory || !OSF._OfficeAppFactory.getHostInfo) {
            return false;
        }
        var hostInfo = OSF._OfficeAppFactory.getHostInfo();
        if (!hostInfo) {
            return false;
        }
        switch (hostInfo["hostType"]) {
            case "outlook":
                switch (hostInfo["hostPlatform"]) {
                    case "mac":
                    case "web":
                        return true;
                    default:
                        return false;
                }
            default:
                return false;
        }
    }
    function isAPIUsageEnabledDispId(dispId, apiType) {
        if (apiType === "method") {
            switch (dispId) {
                case 3:
                case 4:
                case 38:
                case 37:
                case 10:
                case 12:
                    return true;
                default:
                    return false;
            }
        }
        return false;
    }
    function canSendAddinId() {
        var isPublic = (OSF._OfficeAppFactory.getHostInfo().flags & OSF.HostInfoFlags.PublicAddin) != 0;
        if (isPublic) {
            return isPublic;
        }
        if (!appInfo) {
            return false;
        }
        var hostPlatform = OSF._OfficeAppFactory.getHostInfo().hostPlatform;
        var hostVersion = appInfo.hostVersion;
        return _isComplianceExceptedHost(hostPlatform, hostVersion);
    }
    OSFAppTelemetry.canSendAddinId = canSendAddinId;
    function getCompliantAppInstanceId(addinId, appInstanceId) {
        if (!canSendAddinId() && appInstanceId === addinId) {
            return privateAddinId;
        }
        return appInstanceId;
    }
    OSFAppTelemetry.getCompliantAppInstanceId = getCompliantAppInstanceId;
    function _isComplianceExceptedHost(hostPlatform, hostVersion) {
        var excepted = false;
        var versionExtractor = /^(\d+)\.(\d+)\.(\d+)\.(\d+)$/;
        var result = versionExtractor.exec(hostVersion);
        if (result) {
            var major = parseInt(result[1]);
            var minor = parseInt(result[2]);
            var build = parseInt(result[3]);
            if (hostPlatform == "win32") {
                if (major < 16 || major == 16 && build < 14225) {
                    excepted = true;
                }
            }
            else if (hostPlatform == "mac") {
                if (major < 16 || (major == 16 && (minor < 52 || minor == 52 && build < 808))) {
                    excepted = true;
                }
            }
        }
        return excepted;
    }
    OSFAppTelemetry._isComplianceExceptedHost = _isComplianceExceptedHost;
    OSF.AppTelemetry = OSFAppTelemetry;
})(OSFAppTelemetry || (OSFAppTelemetry = {}));
var OSFPerfUtil;
(function (OSFPerfUtil) {
    function prepareDataFieldsForOtel(resource, name) {
        name = name + "_Resource";
        if (oteljs !== undefined) {
            return [
                oteljs.makeStringDataField(name + "_name", resource.name),
                oteljs.makeDoubleDataField(name + "_responseEnd", resource.responseEnd),
                oteljs.makeDoubleDataField(name + "_responseStart", resource.responseStart),
                oteljs.makeDoubleDataField(name + "_startTime", resource.startTime),
                oteljs.makeDoubleDataField(name + "_transferSize", resource.transferSize)
            ];
        }
    }
    function sendPerformanceTelemetry() {
        if (typeof OTel !== "undefined" && OSF.AppTelemetry.enableTelemetry && typeof OSFPerformance !== "undefined" && typeof (performance) != "undefined" && performance.getEntriesByType) {
            var hostPerfResource_1;
            var officePerfResource_1;
            var resources = performance.getEntriesByType("resource");
            resources.forEach(function (resource) {
                if (OSF.OUtil.stringEndsWith(resource.name, OSFPerformance.hostSpecificFileName)) {
                    hostPerfResource_1 = resource;
                }
                else if (OSF.OUtil.stringEndsWith(resource.name, OSF.ConstantNames.OfficeDebugJS) ||
                    OSF.OUtil.stringEndsWith(resource.name, OSF.ConstantNames.OfficeJS)) {
                    officePerfResource_1 = resource;
                }
            });
            OTel.OTelLogger.onTelemetryLoaded(function () {
                var dataFields = prepareDataFieldsForOtel(hostPerfResource_1, "HostJs");
                dataFields = dataFields.concat(prepareDataFieldsForOtel(officePerfResource_1, "OfficeJs"));
                dataFields = dataFields.concat([
                    oteljs.makeDoubleDataField("officeExecuteStartDate", OSFPerformance.officeExecuteStartDate),
                    oteljs.makeDoubleDataField("officeExecuteStart", OSFPerformance.officeExecuteStart),
                    oteljs.makeDoubleDataField("officeExecuteEnd", OSFPerformance.officeExecuteEnd),
                    oteljs.makeDoubleDataField("hostInitializationStart", OSFPerformance.hostInitializationStart),
                    oteljs.makeDoubleDataField("hostInitializationEnd", OSFPerformance.hostInitializationEnd),
                    oteljs.makeDoubleDataField("totalJSHeapSize", OSFPerformance.totalJSHeapSize),
                    oteljs.makeDoubleDataField("usedJSHeapSize", OSFPerformance.usedJSHeapSize),
                    oteljs.makeDoubleDataField("jsHeapSizeLimit", OSFPerformance.jsHeapSizeLimit),
                    oteljs.makeDoubleDataField("getAppContextStart", OSFPerformance.getAppContextStart),
                    oteljs.makeDoubleDataField("getAppContextEnd", OSFPerformance.getAppContextEnd),
                    oteljs.makeDoubleDataField("createOMEnd", OSFPerformance.createOMEnd),
                    oteljs.makeDoubleDataField("officeOnReady", OSFPerformance.officeOnReady),
                    oteljs.makeBooleanDataField("isSharedRuntime", (OSF._OfficeAppFactory.getHostInfo().flags & OSF.HostInfoFlags.SharedApp) !== 0)
                ]);
                Microsoft.Office.WebExtension.sendTelemetryEvent({
                    eventName: "Office.Extensibility.OfficeJs.JSPerformanceTelemetryV06",
                    dataFields: dataFields,
                    eventFlags: {
                        dataCategories: oteljs.DataCategories.ProductServiceUsage,
                        diagnosticLevel: oteljs.DiagnosticLevel.NecessaryServiceDataEvent
                    }
                });
            });
        }
    }
    OSFPerfUtil.sendPerformanceTelemetry = sendPerformanceTelemetry;
})(OSFPerfUtil || (OSFPerfUtil = {}));
Microsoft.Office.WebExtension.EventType = {};
OSF.EventDispatch = function OSF_EventDispatch(eventTypes) {
    this._eventHandlers = {};
    this._objectEventHandlers = {};
    this._queuedEventsArgs = {};
    if (eventTypes != null) {
        for (var i = 0; i < eventTypes.length; i++) {
            var eventType = eventTypes[i];
            var isObjectEvent = (eventType == "objectDeleted" || eventType == "objectSelectionChanged" || eventType == "objectDataChanged" || eventType == "contentControlAdded");
            if (!isObjectEvent)
                this._eventHandlers[eventType] = [];
            else
                this._objectEventHandlers[eventType] = {};
            this._queuedEventsArgs[eventType] = [];
        }
    }
};
OSF.EventDispatch.prototype = {
    getSupportedEvents: function OSF_EventDispatch$getSupportedEvents() {
        var events = [];
        for (var eventName in this._eventHandlers)
            events.push(eventName);
        for (var eventName in this._objectEventHandlers)
            events.push(eventName);
        return events;
    },
    supportsEvent: function OSF_EventDispatch$supportsEvent(event) {
        for (var eventName in this._eventHandlers) {
            if (event == eventName)
                return true;
        }
        for (var eventName in this._objectEventHandlers) {
            if (event == eventName)
                return true;
        }
        return false;
    },
    hasEventHandler: function OSF_EventDispatch$hasEventHandler(eventType, handler) {
        var handlers = this._eventHandlers[eventType];
        if (handlers && handlers.length > 0) {
            for (var i = 0; i < handlers.length; i++) {
                if (handlers[i] === handler)
                    return true;
            }
        }
        return false;
    },
    hasObjectEventHandler: function OSF_EventDispatch$hasObjectEventHandler(eventType, objectId, handler) {
        var handlers = this._objectEventHandlers[eventType];
        if (handlers != null) {
            var _handlers = handlers[objectId];
            for (var i = 0; _handlers != null && i < _handlers.length; i++) {
                if (_handlers[i] === handler)
                    return true;
            }
        }
        return false;
    },
    addEventHandler: function OSF_EventDispatch$addEventHandler(eventType, handler) {
        if (typeof handler != "function") {
            return false;
        }
        var handlers = this._eventHandlers[eventType];
        if (handlers && !this.hasEventHandler(eventType, handler)) {
            handlers.push(handler);
            return true;
        }
        else {
            return false;
        }
    },
    addObjectEventHandler: function OSF_EventDispatch$addObjectEventHandler(eventType, objectId, handler) {
        if (typeof handler != "function") {
            return false;
        }
        var handlers = this._objectEventHandlers[eventType];
        if (handlers && !this.hasObjectEventHandler(eventType, objectId, handler)) {
            if (handlers[objectId] == null)
                handlers[objectId] = [];
            handlers[objectId].push(handler);
            return true;
        }
        return false;
    },
    addEventHandlerAndFireQueuedEvent: function OSF_EventDispatch$addEventHandlerAndFireQueuedEvent(eventType, handler) {
        var handlers = this._eventHandlers[eventType];
        var isFirstHandler = handlers.length == 0;
        var succeed = this.addEventHandler(eventType, handler);
        if (isFirstHandler && succeed) {
            this.fireQueuedEvent(eventType);
        }
        return succeed;
    },
    removeEventHandler: function OSF_EventDispatch$removeEventHandler(eventType, handler) {
        var handlers = this._eventHandlers[eventType];
        if (handlers && handlers.length > 0) {
            for (var index = 0; index < handlers.length; index++) {
                if (handlers[index] === handler) {
                    handlers.splice(index, 1);
                    return true;
                }
            }
        }
        return false;
    },
    removeObjectEventHandler: function OSF_EventDispatch$removeObjectEventHandler(eventType, objectId, handler) {
        var handlers = this._objectEventHandlers[eventType];
        if (handlers != null) {
            var _handlers = handlers[objectId];
            for (var i = 0; _handlers != null && i < _handlers.length; i++) {
                if (_handlers[i] === handler) {
                    _handlers.splice(i, 1);
                    return true;
                }
            }
        }
        return false;
    },
    clearEventHandlers: function OSF_EventDispatch$clearEventHandlers(eventType) {
        if (typeof this._eventHandlers[eventType] != "undefined" && this._eventHandlers[eventType].length > 0) {
            this._eventHandlers[eventType] = [];
            return true;
        }
        return false;
    },
    clearObjectEventHandlers: function OSF_EventDispatch$clearObjectEventHandlers(eventType, objectId) {
        if (this._objectEventHandlers[eventType] != null && this._objectEventHandlers[eventType][objectId] != null) {
            this._objectEventHandlers[eventType][objectId] = [];
            return true;
        }
        return false;
    },
    getEventHandlerCount: function OSF_EventDispatch$getEventHandlerCount(eventType) {
        return this._eventHandlers[eventType] != undefined ? this._eventHandlers[eventType].length : -1;
    },
    getObjectEventHandlerCount: function OSF_EventDispatch$getObjectEventHandlerCount(eventType, objectId) {
        if (this._objectEventHandlers[eventType] == null || this._objectEventHandlers[eventType][objectId] == null)
            return 0;
        return this._objectEventHandlers[eventType][objectId].length;
    },
    fireEvent: function OSF_EventDispatch$fireEvent(eventArgs) {
        if (eventArgs.type == undefined)
            return false;
        var eventType = eventArgs.type;
        if (eventType && this._eventHandlers[eventType]) {
            var eventHandlers = this._eventHandlers[eventType];
            for (var i = 0; i < eventHandlers.length; i++) {
                eventHandlers[i](eventArgs);
            }
            return true;
        }
        else {
            return false;
        }
    },
    fireObjectEvent: function OSF_EventDispatch$fireObjectEvent(objectId, eventArgs) {
        if (eventArgs.type == undefined)
            return false;
        var eventType = eventArgs.type;
        if (eventType && this._objectEventHandlers[eventType]) {
            var eventHandlers = this._objectEventHandlers[eventType];
            var _handlers = eventHandlers[objectId];
            if (_handlers != null) {
                for (var i = 0; i < _handlers.length; i++)
                    _handlers[i](eventArgs);
                return true;
            }
        }
        return false;
    },
    fireOrQueueEvent: function OSF_EventDispatch$fireOrQueueEvent(eventArgs) {
        var eventType = eventArgs.type;
        if (eventType && this._eventHandlers[eventType]) {
            var eventHandlers = this._eventHandlers[eventType];
            var queuedEvents = this._queuedEventsArgs[eventType];
            if (eventHandlers.length == 0) {
                queuedEvents.push(eventArgs);
            }
            else {
                this.fireEvent(eventArgs);
            }
            return true;
        }
        else {
            return false;
        }
    },
    fireQueuedEvent: function OSF_EventDispatch$queueEvent(eventType) {
        if (eventType && this._eventHandlers[eventType]) {
            var eventHandlers = this._eventHandlers[eventType];
            var queuedEvents = this._queuedEventsArgs[eventType];
            if (eventHandlers.length > 0) {
                var eventHandler = eventHandlers[0];
                while (queuedEvents.length > 0) {
                    var eventArgs = queuedEvents.shift();
                    eventHandler(eventArgs);
                }
                return true;
            }
        }
        return false;
    },
    clearQueuedEvent: function OSF_EventDispatch$clearQueuedEvent(eventType) {
        if (eventType && this._eventHandlers[eventType]) {
            var queuedEvents = this._queuedEventsArgs[eventType];
            if (queuedEvents) {
                this._queuedEventsArgs[eventType] = [];
            }
        }
    }
};
OSF.DDA.OMFactory = OSF.DDA.OMFactory || {};
OSF.DDA.OMFactory.manufactureEventArgs = function OSF_DDA_OMFactory$manufactureEventArgs(eventType, target, eventProperties) {
    var args;
    switch (eventType) {
        case Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged:
            args = new OSF.DDA.DocumentSelectionChangedEventArgs(target);
            break;
        case Microsoft.Office.WebExtension.EventType.BindingSelectionChanged:
            args = new OSF.DDA.BindingSelectionChangedEventArgs(this.manufactureBinding(eventProperties, target.document), eventProperties[OSF.DDA.PropertyDescriptors.Subset]);
            break;
        case Microsoft.Office.WebExtension.EventType.BindingDataChanged:
            args = new OSF.DDA.BindingDataChangedEventArgs(this.manufactureBinding(eventProperties, target.document));
            break;
        case Microsoft.Office.WebExtension.EventType.SettingsChanged:
            args = new OSF.DDA.SettingsChangedEventArgs(target);
            break;
        case Microsoft.Office.WebExtension.EventType.ActiveViewChanged:
            args = new OSF.DDA.ActiveViewChangedEventArgs(eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.OfficeThemeChanged:
            args = new OSF.DDA.Theming.OfficeThemeChangedEventArgs(eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.DocumentThemeChanged:
            args = new OSF.DDA.Theming.DocumentThemeChangedEventArgs(eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.AppCommandInvoked:
            args = OSF.DDA.AppCommand.AppCommandInvokedEventArgs.create(eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.ObjectDeleted:
        case Microsoft.Office.WebExtension.EventType.ObjectSelectionChanged:
        case Microsoft.Office.WebExtension.EventType.ObjectDataChanged:
        case Microsoft.Office.WebExtension.EventType.ContentControlAdded:
            args = new OSF.DDA.ObjectEventArgs(eventType, eventProperties[Microsoft.Office.WebExtension.Parameters.Id]);
            break;
        case Microsoft.Office.WebExtension.EventType.RichApiMessage:
            args = new OSF.DDA.RichApiMessageEventArgs(eventType, eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.DataNodeInserted:
            args = new OSF.DDA.NodeInsertedEventArgs(this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.NewNode]), eventProperties[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
            break;
        case Microsoft.Office.WebExtension.EventType.DataNodeReplaced:
            args = new OSF.DDA.NodeReplacedEventArgs(this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.OldNode]), this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.NewNode]), eventProperties[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
            break;
        case Microsoft.Office.WebExtension.EventType.DataNodeDeleted:
            args = new OSF.DDA.NodeDeletedEventArgs(this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.OldNode]), this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.NextSiblingNode]), eventProperties[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
            break;
        case Microsoft.Office.WebExtension.EventType.TaskSelectionChanged:
            args = new OSF.DDA.TaskSelectionChangedEventArgs(target);
            break;
        case Microsoft.Office.WebExtension.EventType.ResourceSelectionChanged:
            args = new OSF.DDA.ResourceSelectionChangedEventArgs(target);
            break;
        case Microsoft.Office.WebExtension.EventType.ViewSelectionChanged:
            args = new OSF.DDA.ViewSelectionChangedEventArgs(target);
            break;
        case Microsoft.Office.WebExtension.EventType.DialogMessageReceived:
            args = new OSF.DDA.DialogEventArgs(eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived:
            args = new OSF.DDA.DialogParentEventArgs(eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.ItemChanged:
            if (OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook") {
                args = new OSF.DDA.OlkItemSelectedChangedEventArgs(eventProperties);
                target.initialize(args["initialData"]);
                if (OSF._OfficeAppFactory.getHostInfo()["hostPlatform"] == "win32" || OSF._OfficeAppFactory.getHostInfo()["hostPlatform"] == "mac") {
                    target.setCurrentItemNumber(args["itemNumber"].itemNumber);
                }
            }
            else {
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
            }
            break;
        case Microsoft.Office.WebExtension.EventType.RecipientsChanged:
            if (OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook") {
                args = new OSF.DDA.OlkRecipientsChangedEventArgs(eventProperties);
            }
            else {
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
            }
            break;
        case Microsoft.Office.WebExtension.EventType.AppointmentTimeChanged:
            if (OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook") {
                args = new OSF.DDA.OlkAppointmentTimeChangedEventArgs(eventProperties);
            }
            else {
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
            }
            break;
        case Microsoft.Office.WebExtension.EventType.RecurrenceChanged:
            if (OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook") {
                args = new OSF.DDA.OlkRecurrenceChangedEventArgs(eventProperties);
            }
            else {
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
            }
            break;
        case Microsoft.Office.WebExtension.EventType.AttachmentsChanged:
            if (OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook") {
                args = new OSF.DDA.OlkAttachmentsChangedEventArgs(eventProperties);
            }
            else {
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
            }
            break;
        case Microsoft.Office.WebExtension.EventType.EnhancedLocationsChanged:
            if (OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook") {
                args = new OSF.DDA.OlkEnhancedLocationsChangedEventArgs(eventProperties);
            }
            else {
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
            }
            break;
        case Microsoft.Office.WebExtension.EventType.InfobarClicked:
            if (OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook") {
                args = new OSF.DDA.OlkInfobarClickedEventArgs(eventProperties);
            }
            else {
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
            }
            break;
        default:
            throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
    }
    return args;
};
OSF.DDA.AsyncMethodNames.addNames({
    AddHandlerAsync: "addHandlerAsync",
    RemoveHandlerAsync: "removeHandlerAsync"
});
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.AddHandlerAsync,
    requiredArguments: [{
            "name": Microsoft.Office.WebExtension.Parameters.EventType,
            "enum": Microsoft.Office.WebExtension.EventType,
            "verify": function (eventType, caller, eventDispatch) { return eventDispatch.supportsEvent(eventType); }
        },
        {
            "name": Microsoft.Office.WebExtension.Parameters.Handler,
            "types": ["function"]
        }
    ],
    supportedOptions: [],
    privateStateCallbacks: []
});
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.RemoveHandlerAsync,
    requiredArguments: [
        {
            "name": Microsoft.Office.WebExtension.Parameters.EventType,
            "enum": Microsoft.Office.WebExtension.EventType,
            "verify": function (eventType, caller, eventDispatch) { return eventDispatch.supportsEvent(eventType); }
        }
    ],
    supportedOptions: [
        {
            name: Microsoft.Office.WebExtension.Parameters.Handler,
            value: {
                "types": ["function", "object"],
                "defaultValue": null
            }
        }
    ],
    privateStateCallbacks: []
});
(function (OfficeExt) {
    var AppCommand;
    (function (AppCommand) {
        var AppCommandManager = (function () {
            function AppCommandManager() {
                var _this = this;
                this._pseudoDocument = null;
                this._eventDispatch = null;
                this._processAppCommandInvocation = function (args) {
                    var verifyResult = _this._verifyManifestCallback(args.callbackName);
                    if (verifyResult.errorCode != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                        _this._invokeAppCommandCompletedMethod(args.appCommandId, verifyResult.errorCode, "");
                        return;
                    }
                    var eventObj = _this._constructEventObjectForCallback(args);
                    if (eventObj) {
                        window.setTimeout(function () { verifyResult.callback(eventObj); }, 0);
                    }
                    else {
                        _this._invokeAppCommandCompletedMethod(args.appCommandId, OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError, "");
                    }
                };
            }
            AppCommandManager.initializeOsfDda = function () {
                OSF.DDA.AsyncMethodNames.addNames({
                    AppCommandInvocationCompletedAsync: "appCommandInvocationCompletedAsync"
                });
                OSF.DDA.AsyncMethodCalls.define({
                    method: OSF.DDA.AsyncMethodNames.AppCommandInvocationCompletedAsync,
                    requiredArguments: [{
                            "name": Microsoft.Office.WebExtension.Parameters.Id,
                            "types": ["string"]
                        },
                        {
                            "name": Microsoft.Office.WebExtension.Parameters.Status,
                            "types": ["number"]
                        },
                        {
                            "name": Microsoft.Office.WebExtension.Parameters.AppCommandInvocationCompletedData,
                            "types": ["string"]
                        }
                    ]
                });
                OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
                    AppCommandInvokedEvent: "AppCommandInvokedEvent"
                });
                OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
                    AppCommandInvoked: "appCommandInvoked"
                });
                OSF.OUtil.setNamespace("AppCommand", OSF.DDA);
                OSF.DDA.AppCommand.AppCommandInvokedEventArgs = OfficeExt.AppCommand.AppCommandInvokedEventArgs;
            };
            AppCommandManager.prototype.initializeAndChangeOnce = function (callback) {
                AppCommand.registerDdaFacade();
                this._pseudoDocument = {};
                OSF.DDA.DispIdHost.addAsyncMethods(this._pseudoDocument, [
                    OSF.DDA.AsyncMethodNames.AppCommandInvocationCompletedAsync,
                ]);
                this._eventDispatch = new OSF.EventDispatch([
                    Microsoft.Office.WebExtension.EventType.AppCommandInvoked,
                ]);
                var onRegisterCompleted = function (result) {
                    if (callback) {
                        if (result.status == "succeeded") {
                            callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
                        }
                        else {
                            callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
                        }
                    }
                };
                OSF.DDA.DispIdHost.addEventSupport(this._pseudoDocument, this._eventDispatch);
                this._pseudoDocument.addHandlerAsync(Microsoft.Office.WebExtension.EventType.AppCommandInvoked, this._processAppCommandInvocation, onRegisterCompleted);
            };
            AppCommandManager.prototype._verifyManifestCallback = function (callbackName) {
                var defaultResult = { callback: null, errorCode: OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidCallback };
                callbackName = callbackName.trim();
                try {
                    var callbackFunc = this._getCallbackFunc(callbackName);
                    if (typeof callbackFunc != "function") {
                        return defaultResult;
                    }
                }
                catch (e) {
                    return defaultResult;
                }
                return { callback: callbackFunc, errorCode: OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess };
            };
            AppCommandManager.prototype._getCallbackFuncFromWindow = function (callbackName) {
                var callList = callbackName.split(".");
                var parentObject = window;
                for (var i = 0; i < callList.length - 1; i++) {
                    if (parentObject[callList[i]] && (typeof parentObject[callList[i]] == "object" || typeof parentObject[callList[i]] == "function")) {
                        parentObject = parentObject[callList[i]];
                    }
                    else {
                        return null;
                    }
                }
                var callbackFunc = parentObject[callList[callList.length - 1]];
                return callbackFunc;
            };
            AppCommandManager.prototype._getCallbackFuncFromActionAssociateTable = function (callbackName) {
                var nameUpperCase = callbackName.toUpperCase();
                return Office.actions._association.mappings[nameUpperCase];
            };
            AppCommandManager.prototype._getCallbackFunc = function (callbackName) {
                var callbackFunc = this._getCallbackFuncFromWindow(callbackName);
                if (!callbackFunc) {
                    callbackFunc = this._getCallbackFuncFromActionAssociateTable(callbackName);
                }
                return callbackFunc;
            };
            AppCommandManager.prototype._invokeAppCommandCompletedMethod = function (appCommandId, resultCode, data) {
                this._pseudoDocument.appCommandInvocationCompletedAsync(appCommandId, resultCode, data);
            };
            AppCommandManager.prototype._constructEventObjectForCallback = function (args) {
                var _this = this;
                var eventObj = new AppCommandCallbackEventArgs();
                try {
                    var jsonData = JSON.parse(args.eventObjStr);
                    this._translateEventObjectInternal(jsonData, eventObj);
                    Object.defineProperty(eventObj, 'completed', {
                        value: function (completedContext) {
                            eventObj.completedContext = completedContext;
                            var jsonString = JSON.stringify(eventObj);
                            _this._invokeAppCommandCompletedMethod(args.appCommandId, OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess, jsonString);
                        },
                        enumerable: true
                    });
                }
                catch (e) {
                    eventObj = null;
                }
                return eventObj;
            };
            AppCommandManager.prototype._translateEventObjectInternal = function (input, output) {
                for (var key in input) {
                    if (!input.hasOwnProperty(key))
                        continue;
                    var inputChild = input[key];
                    if (typeof inputChild == "object" && inputChild != null) {
                        OSF.OUtil.defineEnumerableProperty(output, key, {
                            value: {}
                        });
                        this._translateEventObjectInternal(inputChild, output[key]);
                    }
                    else {
                        Object.defineProperty(output, key, {
                            value: inputChild,
                            enumerable: true,
                            writable: true
                        });
                    }
                }
            };
            AppCommandManager.prototype._constructObjectByTemplate = function (template, input) {
                var output = {};
                if (!template || !input)
                    return output;
                for (var key in template) {
                    if (template.hasOwnProperty(key)) {
                        output[key] = null;
                        if (input[key] != null) {
                            var templateChild = template[key];
                            var inputChild = input[key];
                            var inputChildType = typeof inputChild;
                            if (typeof templateChild == "object" && templateChild != null) {
                                output[key] = this._constructObjectByTemplate(templateChild, inputChild);
                            }
                            else if (inputChildType == "number" || inputChildType == "string" || inputChildType == "boolean") {
                                output[key] = inputChild;
                            }
                        }
                    }
                }
                return output;
            };
            AppCommandManager.instance = function () {
                if (AppCommandManager._instance == null) {
                    AppCommandManager._instance = new AppCommandManager();
                }
                return AppCommandManager._instance;
            };
            AppCommandManager._instance = null;
            return AppCommandManager;
        }());
        AppCommand.AppCommandManager = AppCommandManager;
        var AppCommandInvokedEventArgs = (function () {
            function AppCommandInvokedEventArgs(appCommandId, callbackName, eventObjStr) {
                this.type = Microsoft.Office.WebExtension.EventType.AppCommandInvoked;
                this.appCommandId = appCommandId;
                this.callbackName = callbackName;
                this.eventObjStr = eventObjStr;
            }
            AppCommandInvokedEventArgs.create = function (eventProperties) {
                return new AppCommandInvokedEventArgs(eventProperties[AppCommand.AppCommandInvokedEventEnums.AppCommandId], eventProperties[AppCommand.AppCommandInvokedEventEnums.CallbackName], eventProperties[AppCommand.AppCommandInvokedEventEnums.EventObjStr]);
            };
            return AppCommandInvokedEventArgs;
        }());
        AppCommand.AppCommandInvokedEventArgs = AppCommandInvokedEventArgs;
        var AppCommandCallbackEventArgs = (function () {
            function AppCommandCallbackEventArgs() {
            }
            return AppCommandCallbackEventArgs;
        }());
        AppCommand.AppCommandCallbackEventArgs = AppCommandCallbackEventArgs;
        AppCommand.AppCommandInvokedEventEnums = {
            AppCommandId: "appCommandId",
            CallbackName: "callbackName",
            EventObjStr: "eventObjStr"
        };
    })(AppCommand = OfficeExt.AppCommand || (OfficeExt.AppCommand = {}));
})(OfficeExt || (OfficeExt = {}));
OfficeExt.AppCommand.AppCommandManager.initializeOsfDda();
OSF.OUtil.setNamespace("Marshaling", OSF.DDA);
OSF.OUtil.setNamespace("AppCommand", OSF.DDA.Marshaling);
var OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys;
(function (OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys) {
    OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys[OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys["AppCommandId"] = 0] = "AppCommandId";
    OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys[OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys["CallbackName"] = 1] = "CallbackName";
    OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys[OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys["EventObjStr"] = 2] = "EventObjStr";
})(OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys || (OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys = {}));
;
OSF.DDA.Marshaling.AppCommand.AppCommandInvokedEventKeys = OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys;
var OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys;
(function (OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys) {
    OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys[OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys["Id"] = 0] = "Id";
    OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys[OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys["Status"] = 1] = "Status";
    OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys[OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys["Data"] = 2] = "Data";
})(OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys || (OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys = {}));
;
OSF.DDA.Marshaling.AppCommand.AppCommandCompletedMethodParameterKeys = OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys;
(function (OfficeExt) {
    var AppCommand;
    (function (AppCommand) {
        function registerDdaFacade() {
            if (OSF.DDA.WAC) {
                var parameterMap = OSF.DDA.WAC.Delegate.ParameterMap;
                parameterMap.define({
                    type: OSF.DDA.MethodDispId.dispidAppCommandInvocationCompletedMethod,
                    toHost: [
                        { name: Microsoft.Office.WebExtension.Parameters.Id, value: OSF.DDA.Marshaling.AppCommand.AppCommandCompletedMethodParameterKeys.Id },
                        { name: Microsoft.Office.WebExtension.Parameters.Status, value: OSF.DDA.Marshaling.AppCommand.AppCommandCompletedMethodParameterKeys.Status },
                        { name: Microsoft.Office.WebExtension.Parameters.AppCommandInvocationCompletedData, value: OSF.DDA.Marshaling.AppCommand.AppCommandCompletedMethodParameterKeys.Data }
                    ]
                });
                parameterMap.define({
                    type: OSF.DDA.EventDispId.dispidAppCommandInvokedEvent,
                    fromHost: [
                        { name: OSF.DDA.EventDescriptors.AppCommandInvokedEvent, value: parameterMap.self }
                    ]
                });
                parameterMap.addComplexType(OSF.DDA.EventDescriptors.AppCommandInvokedEvent);
                parameterMap.define({
                    type: OSF.DDA.EventDescriptors.AppCommandInvokedEvent,
                    fromHost: [
                        { name: OfficeExt.AppCommand.AppCommandInvokedEventEnums.AppCommandId, value: OSF.DDA.Marshaling.AppCommand.AppCommandInvokedEventKeys.AppCommandId },
                        { name: OfficeExt.AppCommand.AppCommandInvokedEventEnums.CallbackName, value: OSF.DDA.Marshaling.AppCommand.AppCommandInvokedEventKeys.CallbackName },
                        { name: OfficeExt.AppCommand.AppCommandInvokedEventEnums.EventObjStr, value: OSF.DDA.Marshaling.AppCommand.AppCommandInvokedEventKeys.EventObjStr },
                    ]
                });
            }
        }
        AppCommand.registerDdaFacade = registerDdaFacade;
    })(AppCommand = OfficeExt.AppCommand || (OfficeExt.AppCommand = {}));
})(OfficeExt || (OfficeExt = {}));
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
    DialogParentMessageReceivedEvent: "DialogParentMessageReceivedEvent"
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
    DialogParentMessageReceived: "dialogParentMessageReceived",
    DialogParentEventReceived: "dialogParentEventReceived"
});
OSF.DialogParentMessageEventDispatch = new OSF.EventDispatch([
    Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived,
    Microsoft.Office.WebExtension.EventType.DialogParentEventReceived
]);
OSF.DDA.UI.EnableMessageChildDialogAPI = true;
OSF.DialogShownStatus = { hasDialogShown: false, isWindowDialog: false };
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
    DialogMessageReceivedEvent: "DialogMessageReceivedEvent"
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
    DialogMessageReceived: "dialogMessageReceived",
    DialogEventReceived: "dialogEventReceived"
});
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
    MessageType: "messageType",
    MessageContent: "messageContent",
    MessageOrigin: "messageOrigin"
});
OSF.DDA.DialogEventType = {};
OSF.OUtil.augmentList(OSF.DDA.DialogEventType, {
    DialogClosed: "dialogClosed",
    NavigationFailed: "naviationFailed"
});
OSF.DDA.AsyncMethodNames.addNames({
    DisplayDialogAsync: "displayDialogAsync",
    CloseAsync: "close"
});
OSF.DDA.SyncMethodNames.addNames({
    MessageParent: "messageParent",
    MessageChild: "messageChild",
    SendMessage: "sendMessage",
    AddMessageHandler: "addEventHandler"
});
OSF.DDA.UI.ParentUI = function OSF_DDA_ParentUI() {
    var eventDispatch;
    if (Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived != null) {
        eventDispatch = new OSF.EventDispatch([
            Microsoft.Office.WebExtension.EventType.DialogMessageReceived,
            Microsoft.Office.WebExtension.EventType.DialogEventReceived,
            Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived
        ]);
    }
    else {
        eventDispatch = new OSF.EventDispatch([
            Microsoft.Office.WebExtension.EventType.DialogMessageReceived,
            Microsoft.Office.WebExtension.EventType.DialogEventReceived
        ]);
    }
    var openDialogName = OSF.DDA.AsyncMethodNames.DisplayDialogAsync.displayName;
    var target = this;
    if (!target[openDialogName]) {
        OSF.OUtil.defineEnumerableProperty(target, openDialogName, {
            value: function () {
                var openDialog = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.OpenDialog];
                openDialog(arguments, eventDispatch, target);
            }
        });
    }
    OSF.OUtil.finalizeProperties(this);
};
OSF.DDA.UI.ChildUI = function OSF_DDA_ChildUI(isPopupWindow) {
    var messageParentName = OSF.DDA.SyncMethodNames.MessageParent.displayName;
    var target = this;
    if (!target[messageParentName]) {
        OSF.OUtil.defineEnumerableProperty(target, messageParentName, {
            value: function () {
                var messageParent = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.MessageParent];
                return messageParent(arguments, target);
            }
        });
    }
    var addEventHandler = OSF.DDA.SyncMethodNames.AddMessageHandler.displayName;
    if (!target[addEventHandler] && typeof OSF.DialogParentMessageEventDispatch != "undefined") {
        OSF.DDA.DispIdHost.addEventSupport(target, OSF.DialogParentMessageEventDispatch, isPopupWindow);
    }
    OSF.OUtil.finalizeProperties(this);
};
OSF.DialogHandler = function OSF_DialogHandler() { };
OSF.DDA.DialogEventArgs = function OSF_DDA_DialogEventArgs(message) {
    if (message[OSF.DDA.PropertyDescriptors.MessageType] == OSF.DialogMessageType.DialogMessageReceived) {
        OSF.OUtil.defineEnumerableProperties(this, {
            "type": {
                value: Microsoft.Office.WebExtension.EventType.DialogMessageReceived
            },
            "message": {
                value: message[OSF.DDA.PropertyDescriptors.MessageContent]
            },
            "origin": {
                value: message[OSF.DDA.PropertyDescriptors.MessageOrigin]
            }
        });
    }
    else {
        OSF.OUtil.defineEnumerableProperties(this, {
            "type": {
                value: Microsoft.Office.WebExtension.EventType.DialogEventReceived
            },
            "error": {
                value: message[OSF.DDA.PropertyDescriptors.MessageType]
            }
        });
    }
};
OSF.DDA.DialogParentEventArgs = function OSF_DDA_DialogParentEventArgs(message) {
    OSF.OUtil.defineEnumerableProperties(this, {
        "type": {
            value: Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived
        },
        "message": {
            value: message[OSF.DDA.PropertyDescriptors.MessageContent]
        },
        "origin": {
            value: message[OSF.DDA.PropertyDescriptors.MessageOrigin]
        }
    });
};
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.DisplayDialogAsync,
    requiredArguments: [
        {
            "name": Microsoft.Office.WebExtension.Parameters.Url,
            "types": ["string"]
        }
    ],
    supportedOptions: [
        {
            name: Microsoft.Office.WebExtension.Parameters.Width,
            value: {
                "types": ["number"],
                "defaultValue": 99
            }
        },
        {
            name: Microsoft.Office.WebExtension.Parameters.Height,
            value: {
                "types": ["number"],
                "defaultValue": 99
            }
        },
        {
            name: Microsoft.Office.WebExtension.Parameters.RequireHTTPs,
            value: {
                "types": ["boolean"],
                "defaultValue": true
            }
        },
        {
            name: Microsoft.Office.WebExtension.Parameters.DisplayInIframe,
            value: {
                "types": ["boolean"],
                "defaultValue": false
            }
        },
        {
            name: Microsoft.Office.WebExtension.Parameters.HideTitle,
            value: {
                "types": ["boolean"],
                "defaultValue": false
            }
        },
        {
            name: Microsoft.Office.WebExtension.Parameters.UseDeviceIndependentPixels,
            value: {
                "types": ["boolean"],
                "defaultValue": false
            }
        },
        {
            name: Microsoft.Office.WebExtension.Parameters.PromptBeforeOpen,
            value: {
                "types": ["boolean"],
                "defaultValue": true
            }
        },
        {
            name: Microsoft.Office.WebExtension.Parameters.EnforceAppDomain,
            value: {
                "types": ["boolean"],
                "defaultValue": true
            }
        },
        {
            name: Microsoft.Office.WebExtension.Parameters.UrlNoHostInfo,
            value: {
                "types": ["boolean"],
                "defaultValue": false
            }
        }
    ],
    privateStateCallbacks: [],
    onSucceeded: function (args, caller, callArgs) {
        var targetId = args[Microsoft.Office.WebExtension.Parameters.Id];
        var eventDispatch = args[Microsoft.Office.WebExtension.Parameters.Data];
        var dialog = new OSF.DialogHandler();
        var closeDialog = OSF.DDA.AsyncMethodNames.CloseAsync.displayName;
        OSF.OUtil.defineEnumerableProperty(dialog, closeDialog, {
            value: function () {
                var closeDialogfunction = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.CloseDialog];
                closeDialogfunction(arguments, targetId, eventDispatch, dialog);
            }
        });
        var addHandler = OSF.DDA.SyncMethodNames.AddMessageHandler.displayName;
        OSF.OUtil.defineEnumerableProperty(dialog, addHandler, {
            value: function () {
                var syncMethodCall = OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.AddMessageHandler.id];
                var callArgs = syncMethodCall.verifyAndExtractCall(arguments, dialog, eventDispatch);
                var eventType = callArgs[Microsoft.Office.WebExtension.Parameters.EventType];
                var handler = callArgs[Microsoft.Office.WebExtension.Parameters.Handler];
                return eventDispatch.addEventHandlerAndFireQueuedEvent(eventType, handler);
            }
        });
        if (OSF.DDA.UI.EnableSendMessageDialogAPI === true) {
            var sendMessage = OSF.DDA.SyncMethodNames.SendMessage.displayName;
            OSF.OUtil.defineEnumerableProperty(dialog, sendMessage, {
                value: function () {
                    var execute = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.SendMessage];
                    return execute(arguments, eventDispatch, dialog);
                }
            });
        }
        if (OSF.DDA.UI.EnableMessageChildDialogAPI === true) {
            var messageChild = OSF.DDA.SyncMethodNames.MessageChild.displayName;
            OSF.OUtil.defineEnumerableProperty(dialog, messageChild, {
                value: function () {
                    var execute = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.SendMessage];
                    return execute(arguments, eventDispatch, dialog);
                }
            });
        }
        return dialog;
    },
    checkCallArgs: function (callArgs, caller, stateInfo) {
        if (callArgs[Microsoft.Office.WebExtension.Parameters.Width] <= 0) {
            callArgs[Microsoft.Office.WebExtension.Parameters.Width] = 1;
        }
        if (!callArgs[Microsoft.Office.WebExtension.Parameters.UseDeviceIndependentPixels] && callArgs[Microsoft.Office.WebExtension.Parameters.Width] > 100) {
            callArgs[Microsoft.Office.WebExtension.Parameters.Width] = 99;
        }
        if (callArgs[Microsoft.Office.WebExtension.Parameters.Height] <= 0) {
            callArgs[Microsoft.Office.WebExtension.Parameters.Height] = 1;
        }
        if (!callArgs[Microsoft.Office.WebExtension.Parameters.UseDeviceIndependentPixels] && callArgs[Microsoft.Office.WebExtension.Parameters.Height] > 100) {
            callArgs[Microsoft.Office.WebExtension.Parameters.Height] = 99;
        }
        if (!callArgs[Microsoft.Office.WebExtension.Parameters.RequireHTTPs]) {
            callArgs[Microsoft.Office.WebExtension.Parameters.RequireHTTPs] = true;
        }
        return callArgs;
    }
});
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.CloseAsync,
    requiredArguments: [],
    supportedOptions: [],
    privateStateCallbacks: []
});
OSF.DDA.SyncMethodCalls.define({
    method: OSF.DDA.SyncMethodNames.MessageParent,
    requiredArguments: [
        {
            "name": Microsoft.Office.WebExtension.Parameters.MessageToParent,
            "types": ["string", "number", "boolean"]
        }
    ],
    supportedOptions: [
        {
            name: Microsoft.Office.WebExtension.Parameters.TargetOrigin,
            value: {
                "types": ["string"],
                "defaultValue": ""
            }
        }
    ]
});
OSF.DDA.SyncMethodCalls.define({
    method: OSF.DDA.SyncMethodNames.AddMessageHandler,
    requiredArguments: [
        {
            "name": Microsoft.Office.WebExtension.Parameters.EventType,
            "enum": Microsoft.Office.WebExtension.EventType,
            "verify": function (eventType, caller, eventDispatch) { return eventDispatch.supportsEvent(eventType); }
        },
        {
            "name": Microsoft.Office.WebExtension.Parameters.Handler,
            "types": ["function"]
        }
    ],
    supportedOptions: []
});
OSF.DDA.SyncMethodCalls.define({
    method: OSF.DDA.SyncMethodNames.SendMessage,
    requiredArguments: [
        {
            "name": Microsoft.Office.WebExtension.Parameters.MessageContent,
            "types": ["string"]
        }
    ],
    supportedOptions: [
        {
            name: Microsoft.Office.WebExtension.Parameters.TargetOrigin,
            value: {
                "types": ["string"],
                "defaultValue": ""
            }
        }
    ],
    privateStateCallbacks: []
});
OSF.OUtil.setNamespace("Marshaling", OSF.DDA);
OSF.OUtil.setNamespace("Dialog", OSF.DDA.Marshaling);
OSF.DDA.Marshaling.Dialog.DialogMessageReceivedEventKeys = {
    MessageType: "messageType",
    MessageContent: "messageContent",
    MessageOrigin: "messageOrigin",
    TargetOrigin: "targetOrigin"
};
OSF.DDA.Marshaling.Dialog.DialogParentMessageReceivedEventKeys = {
    MessageType: "messageType",
    MessageContent: "messageContent",
    MessageOrigin: "messageOrigin",
    TargetOrigin: "targetOrigin"
};
OSF.DDA.Marshaling.MessageParentKeys = {
    MessageToParent: "messageToParent",
    TargetOrigin: "targetOrigin"
};
OSF.DDA.Marshaling.DialogNotificationShownEventType = {
    DialogNotificationShown: "dialogNotificationShown"
};
OSF.DDA.Marshaling.SendMessageKeys = {
    MessageContent: "messageContent",
    TargetOrigin: "targetOrigin"
};
(function (OfficeExt) {
    var WacCommonUICssManager;
    (function (WacCommonUICssManager) {
        var hostType = {
            Excel: "excel",
            Word: "word",
            PowerPoint: "powerpoint",
            Outlook: "outlook",
            Visio: "visio"
        };
        function getDialogCssManager(applicationHostType) {
            switch (applicationHostType) {
                case hostType.Excel:
                case hostType.Word:
                case hostType.PowerPoint:
                case hostType.Outlook:
                case hostType.Visio:
                    return new DefaultDialogCSSManager();
                default:
                    return new DefaultDialogCSSManager();
            }
            return null;
        }
        WacCommonUICssManager.getDialogCssManager = getDialogCssManager;
        var DefaultDialogCSSManager = (function () {
            function DefaultDialogCSSManager() {
                this.overlayElementCSS = [
                    "position: absolute",
                    "top: 0",
                    "left: 0",
                    "width: 100%",
                    "height: 100%",
                    "background-color: rgba(198, 198, 198, 0.5)",
                    "z-index: 99998"
                ];
                this.dialogNotificationPanelCSS = [
                    "width: 100%",
                    "height: 190px",
                    "position: absolute",
                    "z-index: 99999",
                    "background-color: rgba(255, 255, 255, 1)",
                    "left: 0px",
                    "top: 50%",
                    "margin-top: -95px"
                ];
                this.newWindowNotificationTextPanelCSS = [
                    "margin: 20px 14px",
                    "font-family: Segoe UI,Arial,Verdana,sans-serif",
                    "font-size: 14px",
                    "height: 100px",
                    "line-height: 100px"
                ];
                this.newWindowNotificationTextSpanCSS = [
                    "display: inline-block",
                    "line-height: normal",
                    "vertical-align: middle"
                ];
                this.crossZoneNotificationTextPanelCSS = [
                    "margin: 20px 14px",
                    "font-family: Segoe UI,Arial,Verdana,sans-serif",
                    "font-size: 14px",
                    "height: 100px",
                ];
                this.dialogNotificationButtonPanelCSS = "margin:0px 9px";
                this.buttonStyleCSS = [
                    "text-align: center",
                    "width: 70px",
                    "height: 25px",
                    "font-size: 14px",
                    "font-family: Segoe UI,Arial,Verdana,sans-serif",
                    "margin: 0px 5px",
                    "border-width: 1px",
                    "border-style: solid"
                ];
            }
            DefaultDialogCSSManager.prototype.getOverlayElementCSS = function () {
                return this.overlayElementCSS.join(";");
            };
            DefaultDialogCSSManager.prototype.getDialogNotificationPanelCSS = function () {
                return this.dialogNotificationPanelCSS.join(";");
            };
            DefaultDialogCSSManager.prototype.getNewWindowNotificationTextPanelCSS = function () {
                return this.newWindowNotificationTextPanelCSS.join(";");
            };
            DefaultDialogCSSManager.prototype.getNewWindowNotificationTextSpanCSS = function () {
                return this.newWindowNotificationTextSpanCSS.join(";");
            };
            DefaultDialogCSSManager.prototype.getCrossZoneNotificationTextPanelCSS = function () {
                return this.crossZoneNotificationTextPanelCSS.join(";");
            };
            DefaultDialogCSSManager.prototype.getDialogNotificationButtonPanelCSS = function () {
                return this.dialogNotificationButtonPanelCSS;
            };
            DefaultDialogCSSManager.prototype.getDialogButtonCSS = function () {
                return this.buttonStyleCSS.join(";");
            };
            return DefaultDialogCSSManager;
        }());
        WacCommonUICssManager.DefaultDialogCSSManager = DefaultDialogCSSManager;
    })(WacCommonUICssManager = OfficeExt.WacCommonUICssManager || (OfficeExt.WacCommonUICssManager = {}));
})(OfficeExt || (OfficeExt = {}));
(function (OfficeExt) {
    var AddinNativeAction;
    (function (AddinNativeAction) {
        var Dialog;
        (function (Dialog) {
            var windowInstance = null;
            var handler = null;
            var overlayElement = null;
            var dialogNotificationPanel = null;
            var closeDialogKey = "osfDialogInternal:action=closeDialog";
            var showDialogCallback = null;
            var hasCrossZoneNotification = false;
            var checkWindowDialogCloseInterval = -1;
            var messageParentKey = "messageParentKey";
            var hostThemeButtonStyle = null;
            var commonButtonBorderColor = "#ababab";
            var commonButtonBackgroundColor = "#ffffff";
            var commonEventInButtonBackgroundColor = "#ccc";
            var newWindowNotificationId = "newWindowNotificaiton";
            var crossZoneNotificationId = "crossZoneNotification";
            var configureBrowserLinkId = "configureBrowserLink";
            var dialogNotificationTextPanelId = "dialogNotificationTextPanel";
            var shouldUseLocalStorageToPassMessage = OfficeExt.WACUtils.shouldUseLocalStorageToPassMessage();
            var registerDialogNotificationShownArgs = {
                "dispId": OSF.DDA.EventDispId.dispidDialogNotificationShownInAddinEvent,
                "eventType": OSF.DDA.Marshaling.DialogNotificationShownEventType.DialogNotificationShown,
                "onComplete": null,
                "onCalling": null
            };
            function setHostThemeButtonStyle(args) {
                var hostThemeButtonStyleArgs = args.input;
                if (hostThemeButtonStyleArgs != null) {
                    hostThemeButtonStyle = {
                        HostButtonBorderColor: hostThemeButtonStyleArgs[OSF.HostThemeButtonStyleKeys.ButtonBorderColor],
                        HostButtonBackgroundColor: hostThemeButtonStyleArgs[OSF.HostThemeButtonStyleKeys.ButtonBackgroundColor]
                    };
                }
                args.completed();
            }
            Dialog.setHostThemeButtonStyle = setHostThemeButtonStyle;
            function removeEventListenersForDialog(args) {
                OSF._OfficeAppFactory.getInitializationHelper().addOrRemoveEventListenersForWindow(false);
                args.completed();
            }
            Dialog.removeEventListenersForDialog = removeEventListenersForDialog;
            function handleNewWindowDialog(dialogInfo) {
                try {
                    if (!checkAppDomain(dialogInfo)) {
                        showDialogCallback(OSF.DDA.ErrorCodeManager.errorCodes.ooeAppDomains);
                        return;
                    }
                    if (!dialogInfo[OSF.ShowWindowDialogParameterKeys.PromptBeforeOpen]) {
                        showDialog(dialogInfo);
                        return;
                    }
                    hasCrossZoneNotification = false;
                    var ignoreButtonKeyDownClick = false;
                    var hostInfoObj = OSF._OfficeAppFactory.getInitializationHelper()._hostInfo;
                    var dialogCssManager = OfficeExt.WacCommonUICssManager.getDialogCssManager(hostInfoObj.hostType);
                    var notificationText = OSF.OUtil.formatString(Strings.OfficeOM.L_ShowWindowDialogNotification, OSF._OfficeAppFactory.getInitializationHelper()._appContext._addinName);
                    overlayElement = createOverlayElement(dialogCssManager);
                    var docBodyChildren = removeAndStoreAllChildrenFromNode(document.body);
                    document.body.appendChild(overlayElement);
                    dialogNotificationPanel = createNotificationPanel(dialogCssManager, notificationText);
                    dialogNotificationPanel.id = newWindowNotificationId;
                    var dialogNotificationButtonPanel = createButtonPanel(dialogCssManager);
                    var allowButton = createButtonControl(dialogCssManager, Strings.OfficeOM.L_ShowWindowDialogNotificationAllow);
                    var ignoreButton = createButtonControl(dialogCssManager, Strings.OfficeOM.L_ShowWindowDialogNotificationIgnore);
                    dialogNotificationButtonPanel.appendChild(allowButton);
                    dialogNotificationButtonPanel.appendChild(ignoreButton);
                    dialogNotificationPanel.appendChild(dialogNotificationButtonPanel);
                    document.body.insertBefore(dialogNotificationPanel, document.body.firstChild);
                    allowButton.onclick = function (event) {
                        showDialog(dialogInfo);
                        if (!hasCrossZoneNotification) {
                            dismissDialogNotification();
                        }
                        restoreChildrenToNode(document.body, docBodyChildren);
                        docBodyChildren = [];
                        event.preventDefault();
                        event.stopPropagation();
                    };
                    function removeAndStoreAllChildrenFromNode(node) {
                        var children = [];
                        try {
                            while (node.firstChild && node.firstChild != null) {
                                children.push(node.firstChild);
                                node.removeChild(node.firstChild);
                            }
                        }
                        catch (e) { }
                        return children;
                    }
                    function restoreChildrenToNode(node, children) {
                        try {
                            for (var i = 0; i < children.length; i++) {
                                node.appendChild(children[i]);
                            }
                        }
                        catch (e) { }
                    }
                    function ignoreButtonClickEventHandler(event) {
                        function unregisterDialogNotificationShownEventCallback(status) {
                            removeDialogNotificationElement();
                            setFocusOnFirstElement(status);
                            showDialogCallback(OSF.DDA.ErrorCodeManager.errorCodes.ooeEndUserIgnore);
                        }
                        registerDialogNotificationShownArgs.onCalling = unregisterDialogNotificationShownEventCallback;
                        OSF.DDA.WAC.Delegate.unregisterEventAsync(registerDialogNotificationShownArgs);
                        restoreChildrenToNode(document.body, docBodyChildren);
                        docBodyChildren = [];
                        event.preventDefault();
                        event.stopPropagation();
                    }
                    ignoreButton.onclick = ignoreButtonClickEventHandler;
                    allowButton.addEventListener("keydown", function (event) {
                        if (event.shiftKey && event.keyCode == 9) {
                            handleButtonControlEventOut(allowButton);
                            handleButtonControlEventIn(ignoreButton);
                            ignoreButton.focus();
                            event.preventDefault();
                            event.stopPropagation();
                        }
                    }, false);
                    ignoreButton.addEventListener("keydown", function (event) {
                        if (!event.shiftKey && event.keyCode == 9) {
                            handleButtonControlEventOut(ignoreButton);
                            handleButtonControlEventIn(allowButton);
                            allowButton.focus();
                            event.preventDefault();
                            event.stopPropagation();
                        }
                        else if (event.keyCode == 13) {
                            ignoreButtonKeyDownClick = true;
                            event.preventDefault();
                            event.stopPropagation();
                        }
                    }, false);
                    ignoreButton.addEventListener("keyup", function (event) {
                        if (event.keyCode == 13 && ignoreButtonKeyDownClick) {
                            ignoreButtonKeyDownClick = false;
                            ignoreButtonClickEventHandler(event);
                        }
                    }, false);
                    window.focus();
                    function registerDialogNotificationShownEventCallback(status) {
                        allowButton.focus();
                    }
                    registerDialogNotificationShownArgs.onCalling = registerDialogNotificationShownEventCallback;
                    OSF.DDA.WAC.Delegate.registerEventAsync(registerDialogNotificationShownArgs);
                }
                catch (e) {
                    if (OSF.AppTelemetry) {
                        OSF.AppTelemetry.logAppException("Exception happens at new window dialog." + e);
                    }
                    showDialogCallback(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
                }
            }
            Dialog.handleNewWindowDialog = handleNewWindowDialog;
            function closeDialog(callback) {
                try {
                    if (windowInstance != null) {
                        var appDomains = OSF._OfficeAppFactory.getInitializationHelper()._appContext._appDomains;
                        if (appDomains) {
                            for (var i = 0; i < appDomains.length && appDomains[i].indexOf("://") !== -1; i++) {
                                windowInstance.postMessage(closeDialogKey, appDomains[i]);
                            }
                        }
                        if (windowInstance != null && !windowInstance.closed) {
                            windowInstance.close();
                        }
                        if (shouldUseLocalStorageToPassMessage) {
                            window.removeEventListener("storage", storageChangedHandler);
                        }
                        else {
                            window.removeEventListener("message", receiveMessage);
                        }
                        window.clearInterval(checkWindowDialogCloseInterval);
                        windowInstance = null;
                        callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
                    }
                    else {
                        callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
                    }
                }
                catch (e) {
                    if (OSF.AppTelemetry) {
                        OSF.AppTelemetry.logAppException("Exception happens at close window dialog." + e);
                    }
                    callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
                }
            }
            Dialog.closeDialog = closeDialog;
            function messageParent(params) {
                var message = params.hostCallArgs[Microsoft.Office.WebExtension.Parameters.MessageToParent];
                var targetOrigin = params.hostCallArgs[Microsoft.Office.WebExtension.Parameters.TargetOrigin] || null;
                if (shouldUseLocalStorageToPassMessage) {
                    try {
                        var messageKey = OSF._OfficeAppFactory.getInitializationHelper()._webAppState.id + messageParentKey;
                        window.localStorage.setItem(messageKey, message);
                    }
                    catch (e) {
                        if (OSF.AppTelemetry) {
                            OSF.AppTelemetry.logAppException("Error happened during messageParent method:" + e);
                        }
                    }
                }
                else {
                    postDialogMessage(window.opener, { message: message, targetOrigin: targetOrigin });
                }
            }
            Dialog.messageParent = messageParent;
            function sendMessage(params) {
                if (windowInstance != null) {
                    var message = params.hostCallArgs;
                    var targetOrigin = message[Microsoft.Office.WebExtension.Parameters.TargetOrigin] || null;
                    delete message[Microsoft.Office.WebExtension.Parameters.TargetOrigin];
                    if (typeof message != "string") {
                        message = JSON.stringify(message);
                    }
                    postDialogMessage(windowInstance, { message: message, targetOrigin: targetOrigin });
                }
            }
            Dialog.sendMessage = sendMessage;
            function postDialogMessage(targetWindow, args) {
                var message = args.message;
                var targetOrigin = args.targetOrigin;
                if (targetOrigin) {
                    targetWindow.postMessage(message, targetOrigin);
                }
                else {
                    var appDomains = OSF._OfficeAppFactory.getInitializationHelper()._appContext._appDomains;
                    var currentOrigin = window.location.origin;
                    if (!currentOrigin) {
                        currentOrigin = window.location.protocol + "//"
                            + window.location.hostname
                            + (window.location.port ? ':' + window.location.port : '');
                    }
                    if (appDomains) {
                        for (var i = 0; i < appDomains.length && appDomains[i].indexOf("://") !== -1; i++) {
                            targetWindow.postMessage(message, appDomains[i]);
                        }
                    }
                    if (!Microsoft.Office.Common.XdmCommunicationManager.checkUrlWithAppDomains(appDomains, currentOrigin)) {
                        targetWindow.postMessage(message, currentOrigin);
                    }
                }
            }
            Dialog.postDialogMessage = postDialogMessage;
            function registerMessageReceivedEvent() {
                function receiveCloseDialogMessage(event) {
                    var origin = event.origin;
                    if (event.source == window.opener &&
                        (window.office_disable_receive_dialog_message_prompt === true ||
                            OfficeExt.AddinNativeAction.Dialog.validateTaskpaneDomain(origin, true))) {
                        if (typeof event.data === "string" && event.data.indexOf(closeDialogKey) > -1) {
                            window.close();
                        }
                        else {
                            var rawMessage = event.data, type = typeof rawMessage;
                            if (rawMessage && (type == "object" || type == "string")) {
                                var parsedMessage = (type == "string") ? JSON.parse(rawMessage) : rawMessage;
                                var message = {};
                                message[OSF.DDA.PropertyDescriptors.MessageContent] = parsedMessage.messageContent;
                                message[OSF.DDA.PropertyDescriptors.MessageOrigin] = event.origin;
                                var eventArgs = OSF.DDA.OMFactory.manufactureEventArgs(Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived, null, message);
                                OSF.DialogParentMessageEventDispatch.fireEvent(eventArgs);
                            }
                        }
                    }
                }
                window.addEventListener("message", receiveCloseDialogMessage);
            }
            Dialog.registerMessageReceivedEvent = registerMessageReceivedEvent;
            function setHandlerAndShowDialogCallback(onEventHandler, callback) {
                handler = onEventHandler;
                showDialogCallback = callback;
            }
            Dialog.setHandlerAndShowDialogCallback = setHandlerAndShowDialogCallback;
            function escDismissDialogNotification() {
                try {
                    if (dialogNotificationPanel && (dialogNotificationPanel.id == newWindowNotificationId) && showDialogCallback) {
                        showDialogCallback(OSF.DDA.ErrorCodeManager.errorCodes.ooeEndUserIgnore);
                    }
                }
                catch (e) {
                    if (OSF.AppTelemetry) {
                        OSF.AppTelemetry.logAppException("Error happened during executing displayDialogAsync callback." + e);
                    }
                }
                dismissDialogNotification();
            }
            Dialog.escDismissDialogNotification = escDismissDialogNotification;
            function showCrossZoneNotification(windowUrl, hostType) {
                var okButtonKeyDownClick = false;
                var dialogCssManager = OfficeExt.WacCommonUICssManager.getDialogCssManager(hostType);
                overlayElement = createOverlayElement(dialogCssManager);
                document.body.insertBefore(overlayElement, document.body.firstChild);
                dialogNotificationPanel = createNotificationPanelForCrossZoneIssue(dialogCssManager, windowUrl);
                dialogNotificationPanel.id = crossZoneNotificationId;
                var dialogNotificationButtonPanel = createButtonPanel(dialogCssManager);
                var okButton = createButtonControl(dialogCssManager, Strings.OfficeOM.L_DialogOK ? Strings.OfficeOM.L_DialogOK : "OK");
                dialogNotificationButtonPanel.appendChild(okButton);
                dialogNotificationPanel.appendChild(dialogNotificationButtonPanel);
                document.body.insertBefore(dialogNotificationPanel, document.body.firstChild);
                hasCrossZoneNotification = true;
                okButton.onclick = function () {
                    dismissDialogNotification();
                };
                okButton.addEventListener("keydown", function (event) {
                    if (event.keyCode == 9) {
                        document.getElementById(configureBrowserLinkId).focus();
                        event.preventDefault();
                        event.stopPropagation();
                    }
                    else if (event.keyCode == 13) {
                        okButtonKeyDownClick = true;
                        event.preventDefault();
                        event.stopPropagation();
                    }
                }, false);
                okButton.addEventListener("keyup", function (event) {
                    if (event.keyCode == 13 && okButtonKeyDownClick) {
                        okButtonKeyDownClick = false;
                        dismissDialogNotification();
                        event.preventDefault();
                        event.stopPropagation();
                    }
                }, false);
                document.getElementById(configureBrowserLinkId).addEventListener("keydown", function (event) {
                    if (event.keyCode == 9) {
                        okButton.focus();
                        event.preventDefault();
                        event.stopPropagation();
                    }
                }, false);
                window.focus();
                okButton.focus();
            }
            Dialog.showCrossZoneNotification = showCrossZoneNotification;
            function getWithExpiry(storage, key) {
                var itemStr = storage.getItem(key);
                if (!itemStr) {
                    return 'undefined';
                }
                var item = JSON.parse(itemStr);
                if (!item.expiry || !item.value) {
                    return 'undefined';
                }
                var now = new Date();
                var isExpired = now.getTime() > item.expiry;
                if (isExpired) {
                    storage.removeItem(key);
                    return 'undefined';
                }
                return item.value;
            }
            Dialog.getWithExpiry = getWithExpiry;
            function setWithExpiry(storage, key, value) {
                var ttlInMs = 86400000;
                var now = new Date();
                var item = {
                    value: value,
                    expiry: now.getTime() + ttlInMs
                };
                storage.setItem(key, JSON.stringify(item));
            }
            Dialog.setWithExpiry = setWithExpiry;
            function getLocalStorage() {
                if (OfficeExt.SafeStorage) {
                    return new OfficeExt.SafeStorage(window.localStorage);
                }
                else {
                    return window.localStorage;
                }
            }
            Dialog.getLocalStorage = getLocalStorage;
            function getSessionStorage() {
                if (OfficeExt.SafeStorage) {
                    return new OfficeExt.SafeStorage(window.sessionStorage);
                }
                else {
                    return window.sessionStorage;
                }
            }
            Dialog.getSessionStorage = getSessionStorage;
            function getUrlProtocolHostnamePort(parsedUrl) {
                var port = parsedUrl.port ? (':' + parsedUrl.port) : '';
                return parsedUrl.protocol + "//" + parsedUrl.hostname + port;
            }
            Dialog.getUrlProtocolHostnamePort = getUrlProtocolHostnamePort;
            function validateTaskpaneDomain(taskpaneUrl, allowSubDomains) {
                function urlCompare(url_parser1, url_parser2) {
                    return ((url_parser1.hostname == url_parser2.hostname) &&
                        (url_parser1.protocol == url_parser2.protocol) &&
                        hasSamePort(url_parser1, url_parser2));
                }
                function hasSamePort(url_parser1, url_parser2) {
                    var httpPort = "80";
                    var httpsPort = "443";
                    return ((url_parser1.port == url_parser2.port) ||
                        (url_parser1.port == "" && url_parser1.protocol == "http:" && url_parser2.port == httpPort) ||
                        (url_parser1.port == "" && url_parser1.protocol == "https:" && url_parser2.port == httpsPort) ||
                        (url_parser2.port == "" && url_parser2.protocol == "http:" && url_parser1.port == httpPort) ||
                        (url_parser2.port == "" && url_parser2.protocol == "https:" && url_parser1.port == httpsPort));
                }
                try {
                    if (!taskpaneUrl) {
                        return false;
                    }
                    if (!allowSubDomains) {
                        allowSubDomains = true;
                    }
                    var dialogUrl = window.location.origin;
                    if (!dialogUrl) {
                        dialogUrl = window.location.protocol + "//"
                            + window.location.hostname
                            + (window.location.port ? ':' + window.location.port : '');
                    }
                    var parsedDialogUrl = OSF.OUtil.parseUrl(dialogUrl, true);
                    var parsedTaskpaneUrl = OSF.OUtil.parseUrl(taskpaneUrl, true);
                    if (!parsedDialogUrl.protocol || !parsedDialogUrl.hostname ||
                        !parsedTaskpaneUrl.protocol || !parsedTaskpaneUrl.hostname || parsedTaskpaneUrl.port === undefined) {
                        return false;
                    }
                    var isSameDomain = urlCompare(parsedDialogUrl, parsedTaskpaneUrl);
                    var isSubdomain = false;
                    if (allowSubDomains) {
                        isSubdomain = Microsoft.Office.Common.XdmCommunicationManager.isTargetSubdomainOfSourceLocation(taskpaneUrl, dialogUrl);
                    }
                    if (isSameDomain || isSubdomain) {
                        return true;
                    }
                    var localStorage = this.getLocalStorage();
                    var sessionStorage = this.getSessionStorage();
                    var taskpaneUrlPrefix = parsedTaskpaneUrl.protocol + "//" + parsedTaskpaneUrl.hostname + (parsedTaskpaneUrl.port ? (":" + parsedTaskpaneUrl.port) : "");
                    var storageKeyUrl = '_trusts_' + taskpaneUrlPrefix;
                    var urlTrustValueInLocalStorage = this.getWithExpiry(localStorage, storageKeyUrl);
                    var urlTrustValueInSessionStorage = this.getWithExpiry(sessionStorage, storageKeyUrl);
                    if (urlTrustValueInLocalStorage === 'true') {
                        return true;
                    }
                    else if (urlTrustValueInSessionStorage === 'false') {
                        return false;
                    }
                    var shortTaskpaneUrl = this.getUrlProtocolHostnamePort(parsedTaskpaneUrl);
                    var shortDialogUrl = this.getUrlProtocolHostnamePort(parsedDialogUrl);
                    var promptText = 'You are about to send and receive potentially sensitive information from "' + shortDialogUrl + '". Only click OK if you trust the following website recieving the sensitive information: "' + shortTaskpaneUrl + '".';
                    if (Strings.OfficeOM["L_ConfirmDialogApiTrustsParent"]) {
                        promptText = Strings.OfficeOM["L_ConfirmDialogApiTrustsParent"].replace('{0}', shortDialogUrl).replace('{1}', shortTaskpaneUrl);
                    }
                    if (window.confirm(promptText)) {
                        this.setWithExpiry(localStorage, storageKeyUrl, 'true');
                        this.setWithExpiry(sessionStorage, storageKeyUrl, 'true');
                        return true;
                    }
                    else {
                        this.setWithExpiry(localStorage, storageKeyUrl, 'false');
                        this.setWithExpiry(sessionStorage, storageKeyUrl, 'false');
                        return false;
                    }
                }
                catch (e) {
                    if (OSF.AppTelemetry) {
                        OSF.AppTelemetry.logAppException("Error happened during validateTaskpaneDomain method:" + e);
                    }
                    return false;
                }
            }
            Dialog.validateTaskpaneDomain = validateTaskpaneDomain;
            function validateDialogDomain(dialogUrl, taskpaneUrl, allowSubdomains) {
                if (allowSubdomains === void 0) { allowSubdomains = true; }
                if (!dialogUrl || !taskpaneUrl) {
                    return false;
                }
                var httpsIdentifyString = "https:";
                var parsedDialogUrl = OSF.OUtil.parseUrl(dialogUrl);
                var parsedTaskpaneUrl = OSF.OUtil.parseUrl(taskpaneUrl);
                var appDomains = OSF._OfficeAppFactory.getInitializationHelper()._appContext._appDomains;
                var isHttps = parsedDialogUrl.protocol === httpsIdentifyString;
                if (!isHttps) {
                    return false;
                }
                var isSameDomain = parsedDialogUrl.protocol === parsedTaskpaneUrl.protocol
                    && parsedDialogUrl.hostname === parsedTaskpaneUrl.hostname
                    && parsedDialogUrl.port === parsedTaskpaneUrl.port;
                var isInAppDomains = Microsoft.Office.Common.XdmCommunicationManager.checkUrlWithAppDomains(appDomains, dialogUrl);
                var isTrustedDomain = isSameDomain || isInAppDomains;
                if (!isTrustedDomain && allowSubdomains) {
                    isTrustedDomain = Microsoft.Office.Common.XdmCommunicationManager.isTargetSubdomainOfSourceLocation(taskpaneUrl, dialogUrl);
                }
                return isTrustedDomain;
            }
            function receiveMessage(event) {
                if (event.source == windowInstance) {
                    try {
                        var dialogOrigin = event.origin;
                        var taskpaneUrl = OSF._OfficeAppFactory.getInitializationHelper()._appContext._addInSourceUrl;
                        var isTrustedDomain = validateDialogDomain(dialogOrigin, taskpaneUrl, true);
                        if (!isTrustedDomain) {
                            throw new Error("Received a message from a dialog with an untrusted domain.");
                        }
                        var dialogMessageReceivedArgs = {};
                        dialogMessageReceivedArgs[OSF.DDA.Marshaling.Dialog.DialogMessageReceivedEventKeys.MessageType] = OSF.DialogMessageType.DialogMessageReceived;
                        dialogMessageReceivedArgs[OSF.DDA.Marshaling.Dialog.DialogMessageReceivedEventKeys.MessageContent] = event.data;
                        dialogMessageReceivedArgs[OSF.DDA.Marshaling.Dialog.DialogMessageReceivedEventKeys.MessageOrigin] = dialogOrigin;
                        handler(dialogMessageReceivedArgs);
                    }
                    catch (e) {
                        if (OSF.AppTelemetry) {
                            OSF.AppTelemetry.logAppException("Error happened during receive message handler." + e);
                        }
                    }
                }
            }
            function storageChangedHandler(event) {
                var messageKey = OSF._OfficeAppFactory.getInitializationHelper()._webAppState.id + messageParentKey;
                if (event.key == messageKey) {
                    try {
                        var dialogMessageReceivedArgs = {};
                        dialogMessageReceivedArgs[OSF.DDA.Marshaling.Dialog.DialogMessageReceivedEventKeys.MessageType] = OSF.DialogMessageType.DialogMessageReceived;
                        dialogMessageReceivedArgs[OSF.DDA.Marshaling.Dialog.DialogMessageReceivedEventKeys.MessageContent] = event.newValue;
                        dialogMessageReceivedArgs[OSF.DDA.Marshaling.Dialog.DialogMessageReceivedEventKeys.MessageOrigin] = event.origin;
                        handler(dialogMessageReceivedArgs);
                    }
                    catch (e) {
                        if (OSF.AppTelemetry) {
                            OSF.AppTelemetry.logAppException("Error happened during storage changed handler." + e);
                        }
                    }
                }
            }
            function checkAppDomain(dialogInfo) {
                var appDomains = OSF._OfficeAppFactory.getInitializationHelper()._appContext._appDomains;
                var url = dialogInfo[OSF.ShowWindowDialogParameterKeys.Url];
                var fInDomain = Microsoft.Office.Common.XdmCommunicationManager.checkUrlWithAppDomains(appDomains, url);
                if (!fInDomain) {
                    return OSF._OfficeAppFactory.getInitializationHelper()._appContext._addInSourceUrl
                        && Microsoft.Office.Common.XdmCommunicationManager.isTargetSubdomainOfSourceLocation(OSF._OfficeAppFactory.getInitializationHelper()._appContext._addInSourceUrl, url);
                }
                return fInDomain;
            }
            function showDialog(dialogInfo) {
                var hostInfoObj = OSF._OfficeAppFactory.getInitializationHelper()._hostInfo;
                var hostInfoVals;
                if (OSF.OUtil.checkFlight(OSF.FlightTreatmentNames.IsPrivateAddin)) {
                    hostInfoVals = [
                        hostInfoObj.hostType,
                        hostInfoObj.hostPlatform,
                        hostInfoObj.hostSpecificFileVersion,
                        hostInfoObj.hostLocale,
                        hostInfoObj.osfControlAppCorrelationId,
                        "isDialog",
                        hostInfoObj.disableLogging ? "disableLogging" : "",
                        (hostInfoObj.flags & OSF.HostInfoFlags.PublicAddin)
                    ];
                }
                else {
                    hostInfoVals = [
                        hostInfoObj.hostType,
                        hostInfoObj.hostPlatform,
                        hostInfoObj.hostSpecificFileVersion,
                        hostInfoObj.hostLocale,
                        hostInfoObj.osfControlAppCorrelationId,
                        "isDialog",
                        hostInfoObj.disableLogging ? "disableLogging" : "",
                    ];
                }
                var hostInfo = hostInfoVals.join("$");
                var appContext = OSF._OfficeAppFactory.getInitializationHelper()._appContext;
                appContext._taskpaneUrl = window.location.origin;
                if (!appContext._taskpaneUrl) {
                    appContext._taskpaneUrl = window.location.protocol + "//"
                        + window.location.hostname
                        + (window.location.port ? ':' + window.location.port : '');
                }
                var windowUrl = dialogInfo[OSF.ShowWindowDialogParameterKeys.Url];
                if (!dialogInfo[OSF.ShowWindowDialogParameterKeys.UrlNoHostInfo]) {
                    windowUrl = OfficeExt.WACUtils.addHostInfoAsQueryParam(windowUrl, hostInfo);
                }
                var windowName = JSON.parse(window.name);
                windowName[OSF.WindowNameItemKeys.HostInfo] = hostInfo;
                windowName[OSF.WindowNameItemKeys.AppContext] = appContext;
                var width = dialogInfo[OSF.ShowWindowDialogParameterKeys.Width] * screen.width / 100;
                var height = dialogInfo[OSF.ShowWindowDialogParameterKeys.Height] * screen.height / 100;
                var left = appContext._clientWindowWidth / 2 - width / 2;
                var top = appContext._clientWindowHeight / 2 - height / 2;
                var windowSpecs = "width=" + width + ", height=" + height + ", left=" + left + ", top=" + top + ",channelmode=no,directories=no,fullscreen=no,location=no,menubar=no,resizable=yes,scrollbars=yes,status=no,titlebar=yes,toolbar=no";
                windowInstance = window.open(windowUrl, OfficeExt.WACUtils.serializeObjectToString(windowName), windowSpecs);
                if (windowInstance == null) {
                    OSF.AppTelemetry.logAppCommonMessage("Encountered cross zone issue in displayDialogAsync api.");
                    removeDialogNotificationElement();
                    showCrossZoneNotification(windowUrl, hostInfoObj.hostType);
                    showDialogCallback(OSF.DDA.ErrorCodeManager.errorCodes.ooeCrossZone);
                    return;
                }
                if (shouldUseLocalStorageToPassMessage) {
                    window.addEventListener("storage", storageChangedHandler);
                }
                else {
                    window.addEventListener("message", receiveMessage);
                }
                function checkWindowClose() {
                    try {
                        if (windowInstance == null || windowInstance.closed) {
                            window.clearInterval(checkWindowDialogCloseInterval);
                            if (shouldUseLocalStorageToPassMessage) {
                                window.removeEventListener("storage", storageChangedHandler);
                            }
                            else {
                                window.removeEventListener("message", receiveMessage);
                            }
                            var dialogClosedArgs = {};
                            dialogClosedArgs[OSF.DDA.Marshaling.Dialog.DialogMessageReceivedEventKeys.MessageType] = OSF.DialogMessageType.DialogClosed;
                            handler(dialogClosedArgs);
                        }
                    }
                    catch (e) {
                        if (OSF.AppTelemetry) {
                            OSF.AppTelemetry.logAppException("Error happened during check or handle window close." + e);
                        }
                    }
                }
                checkWindowDialogCloseInterval = window.setInterval(checkWindowClose, 1000);
                if (showDialogCallback != null) {
                    showDialogCallback(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
                }
                else {
                    if (OSF.AppTelemetry) {
                        OSF.AppTelemetry.logAppException("showDialogCallback can not be null.");
                    }
                }
            }
            function createButtonControl(dialogCssManager, buttonValue) {
                var buttonControl = document.createElement("input");
                buttonControl.setAttribute("type", "button");
                buttonControl.style.cssText = dialogCssManager.getDialogButtonCSS();
                buttonControl.style.borderColor = commonButtonBorderColor;
                buttonControl.style.backgroundColor = commonButtonBackgroundColor;
                buttonControl.setAttribute("value", buttonValue);
                var buttonControlEventInHandler = function () {
                    handleButtonControlEventIn(buttonControl);
                };
                var buttonControlEventOutHandler = function () {
                    handleButtonControlEventOut(buttonControl);
                };
                buttonControl.addEventListener("mouseover", buttonControlEventInHandler);
                buttonControl.addEventListener("focus", buttonControlEventInHandler);
                buttonControl.addEventListener("mouseout", buttonControlEventOutHandler);
                buttonControl.addEventListener("focusout", buttonControlEventOutHandler);
                return buttonControl;
            }
            function handleButtonControlEventIn(buttonControl) {
                if (hostThemeButtonStyle != null) {
                    buttonControl.style.borderColor = hostThemeButtonStyle.HostButtonBorderColor;
                    buttonControl.style.backgroundColor = hostThemeButtonStyle.HostButtonBackgroundColor;
                }
                else if (OSF.CommonUI && OSF.CommonUI.HostButtonBorderColor && OSF.CommonUI.HostButtonBackgroundColor) {
                    buttonControl.style.borderColor = OSF.CommonUI.HostButtonBorderColor;
                    buttonControl.style.backgroundColor = OSF.CommonUI.HostButtonBackgroundColor;
                }
                else {
                    buttonControl.style.backgroundColor = commonEventInButtonBackgroundColor;
                }
            }
            function handleButtonControlEventOut(buttonControl) {
                buttonControl.style.borderColor = commonButtonBorderColor;
                buttonControl.style.backgroundColor = commonButtonBackgroundColor;
            }
            function dismissDialogNotification() {
                function unregisterDialogNotificationShownEventCallback(status) {
                    removeDialogNotificationElement();
                    setFocusOnFirstElement(status);
                }
                registerDialogNotificationShownArgs.onCalling = unregisterDialogNotificationShownEventCallback;
                OSF.DDA.WAC.Delegate.unregisterEventAsync(registerDialogNotificationShownArgs);
            }
            function removeDialogNotificationElement() {
                if (dialogNotificationPanel != null) {
                    document.body.removeChild(dialogNotificationPanel);
                    dialogNotificationPanel = null;
                }
                if (overlayElement != null) {
                    document.body.removeChild(overlayElement);
                    overlayElement = null;
                }
            }
            function createOverlayElement(dialogCssManager) {
                var overlayElement = document.createElement("div");
                overlayElement.style.cssText = dialogCssManager.getOverlayElementCSS();
                return overlayElement;
            }
            function createNotificationPanel(dialogCssManager, notificationString) {
                var dialogNotificationPanel = document.createElement("div");
                dialogNotificationPanel.style.cssText = dialogCssManager.getDialogNotificationPanelCSS();
                setAttributeForDialogNotificationPanel(dialogNotificationPanel);
                var dialogNotificationTextPanel = document.createElement("div");
                dialogNotificationTextPanel.style.cssText = dialogCssManager.getNewWindowNotificationTextPanelCSS();
                dialogNotificationTextPanel.id = dialogNotificationTextPanelId;
                if (document.documentElement.getAttribute("dir") == "rtl") {
                    dialogNotificationTextPanel.style.paddingRight = "30px";
                }
                else {
                    dialogNotificationTextPanel.style.paddingLeft = "30px";
                }
                var dialogNotificationTextSpan = document.createElement("span");
                dialogNotificationTextSpan.style.cssText = dialogCssManager.getNewWindowNotificationTextSpanCSS();
                dialogNotificationTextSpan.innerText = notificationString;
                dialogNotificationTextPanel.appendChild(dialogNotificationTextSpan);
                dialogNotificationPanel.appendChild(dialogNotificationTextPanel);
                return dialogNotificationPanel;
            }
            function createButtonPanel(dialogCssManager) {
                var dialogNotificationButtonPanel = document.createElement("div");
                dialogNotificationButtonPanel.style.cssText = dialogCssManager.getDialogNotificationButtonPanelCSS();
                if (document.documentElement.getAttribute("dir") == "rtl") {
                    dialogNotificationButtonPanel.style.cssFloat = "left";
                }
                else {
                    dialogNotificationButtonPanel.style.cssFloat = "right";
                }
                return dialogNotificationButtonPanel;
            }
            function setFocusOnFirstElement(status) {
                if (status != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                    var list = document.querySelectorAll(OSF._OfficeAppFactory.getInitializationHelper()._tabbableElements);
                    OSF.OUtil.focusToFirstTabbable(list, false);
                }
            }
            function createNotificationPanelForCrossZoneIssue(dialogCssManager, windowUrl) {
                var dialogNotificationPanel = document.createElement("div");
                dialogNotificationPanel.style.cssText = dialogCssManager.getDialogNotificationPanelCSS();
                setAttributeForDialogNotificationPanel(dialogNotificationPanel);
                var dialogNotificationTextPanel = document.createElement("div");
                dialogNotificationTextPanel.style.cssText = dialogCssManager.getCrossZoneNotificationTextPanelCSS();
                dialogNotificationTextPanel.id = dialogNotificationTextPanelId;
                var configureBrowserLink = document.createElement("a");
                configureBrowserLink.id = configureBrowserLinkId;
                configureBrowserLink.href = "#";
                configureBrowserLink.innerText = Strings.OfficeOM.L_NewWindowCrossZoneConfigureBrowserLink;
                configureBrowserLink.setAttribute("onclick", "window.open('https://support.microsoft.com/en-us/help/17479/windows-internet-explorer-11-change-security-privacy-settings', '_blank', 'fullscreen=1')");
                var dialogNotificationTextSpan = document.createElement("span");
                if (Strings.OfficeOM.L_NewWindowCrossZone) {
                    dialogNotificationTextSpan.innerHTML = OSF.OUtil.formatString(Strings.OfficeOM.L_NewWindowCrossZone, configureBrowserLink.outerHTML, OfficeExt.WACUtils.getDomainForUrl(windowUrl));
                }
                dialogNotificationTextPanel.appendChild(dialogNotificationTextSpan);
                dialogNotificationPanel.appendChild(dialogNotificationTextPanel);
                return dialogNotificationPanel;
            }
            function setAttributeForDialogNotificationPanel(dialogNotificationDiv) {
                dialogNotificationDiv.setAttribute("role", "dialog");
                dialogNotificationDiv.setAttribute("aria-describedby", dialogNotificationTextPanelId);
            }
        })(Dialog = AddinNativeAction.Dialog || (AddinNativeAction.Dialog = {}));
    })(AddinNativeAction = OfficeExt.AddinNativeAction || (OfficeExt.AddinNativeAction = {}));
})(OfficeExt || (OfficeExt = {}));
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidDialogMessageReceivedEvent,
    fromHost: [
        { name: OSF.DDA.EventDescriptors.DialogMessageReceivedEvent, value: OSF.DDA.WAC.Delegate.ParameterMap.self }
    ]
});
OSF.DDA.WAC.Delegate.ParameterMap.addComplexType(OSF.DDA.EventDescriptors.DialogMessageReceivedEvent);
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDescriptors.DialogMessageReceivedEvent,
    fromHost: [
        { name: OSF.DDA.PropertyDescriptors.MessageType, value: OSF.DDA.Marshaling.Dialog.DialogMessageReceivedEventKeys.MessageType },
        { name: OSF.DDA.PropertyDescriptors.MessageContent, value: OSF.DDA.Marshaling.Dialog.DialogMessageReceivedEventKeys.MessageContent },
        { name: OSF.DDA.PropertyDescriptors.MessageOrigin, value: OSF.DDA.Marshaling.Dialog.DialogMessageReceivedEventKeys.MessageOrigin }
    ]
});
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidDialogParentMessageReceivedEvent,
    fromHost: [
        { name: OSF.DDA.EventDescriptors.DialogParentMessageReceivedEvent, value: OSF.DDA.WAC.Delegate.ParameterMap.self }
    ]
});
OSF.DDA.WAC.Delegate.ParameterMap.addComplexType(OSF.DDA.EventDescriptors.DialogParentMessageReceivedEvent);
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDescriptors.DialogParentMessageReceivedEvent,
    fromHost: [
        { name: OSF.DDA.PropertyDescriptors.MessageType, value: OSF.DDA.Marshaling.Dialog.DialogParentMessageReceivedEventKeys.MessageType },
        { name: OSF.DDA.PropertyDescriptors.MessageContent, value: OSF.DDA.Marshaling.Dialog.DialogParentMessageReceivedEventKeys.MessageContent },
        { name: OSF.DDA.PropertyDescriptors.MessageOrigin, value: OSF.DDA.Marshaling.Dialog.DialogParentMessageReceivedEventKeys.MessageOrigin }
    ]
});
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidMessageParentMethod,
    toHost: [
        { name: Microsoft.Office.WebExtension.Parameters.MessageToParent, value: OSF.DDA.Marshaling.MessageParentKeys.MessageToParent },
        { name: Microsoft.Office.WebExtension.Parameters.TargetOrigin, value: OSF.DDA.Marshaling.MessageParentKeys.TargetOrigin }
    ]
});
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidSendMessageMethod,
    toHost: [
        { name: Microsoft.Office.WebExtension.Parameters.MessageContent, value: OSF.DDA.Marshaling.SendMessageKeys.MessageContent },
        { name: Microsoft.Office.WebExtension.Parameters.TargetOrigin, value: OSF.DDA.Marshaling.MessageParentKeys.TargetOrigin }
    ]
});
OSF.DDA.WAC.Delegate.openDialog = function OSF_DDA_WAC_Delegate$OpenDialog(args) {
    var httpsIdentifyString = "https://";
    var httpIdentifyString = "http://";
    var dialogInfo = JSON.parse(args.targetId);
    var callback = OSF.DDA.WAC.Delegate._getOnAfterRegisterEvent(true, args);
    function showDialogCallback(status) {
        var payload = { "Error": status };
        try {
            callback(Microsoft.Office.Common.InvokeResultCode.noError, payload);
        }
        catch (e) {
            if (OSF.AppTelemetry) {
                OSF.AppTelemetry.logAppException("Exception happens at showDialogCallback." + e);
            }
        }
    }
    if (OSF.DialogShownStatus.hasDialogShown) {
        showDialogCallback(OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogAlreadyOpened);
        return;
    }
    var dialogUrl = dialogInfo[OSF.ShowWindowDialogParameterKeys.Url].toLowerCase();
    var taskpaneUrl = (window.location.href).toLowerCase();
    if (OSF.AppTelemetry) {
        var isSameDomain = false;
        var parentIsSubdomain = false;
        var childIsSubdomain = false;
        var isAppDomain = false;
        var dialogUrlPortionAllowedToLog = "";
        var taskpaneUrlPortionAllowedToLog = "";
        if (OSF.OUtil) {
            var parsedDialogUrl = OSF.OUtil.parseUrl(dialogUrl);
            var parsedTaskpaneUrl = OSF.OUtil.parseUrl(taskpaneUrl);
            isSameDomain = parsedDialogUrl.protocol === parsedTaskpaneUrl.protocol
                && parsedDialogUrl.hostname === parsedTaskpaneUrl.hostname
                && parsedDialogUrl.port === parsedTaskpaneUrl.port;
            dialogUrlPortionAllowedToLog = OSF.OUtil.getHostnamePortionForLogging(parsedDialogUrl.hostname);
            if (isSameDomain) {
                taskpaneUrlPortionAllowedToLog = dialogUrlPortionAllowedToLog;
            }
            else {
                taskpaneUrlPortionAllowedToLog = OSF.OUtil.getHostnamePortionForLogging(parsedTaskpaneUrl.hostname);
                parentIsSubdomain = Microsoft.Office.Common.XdmCommunicationManager.isTargetSubdomainOfSourceLocation(dialogUrl, taskpaneUrl);
                childIsSubdomain = Microsoft.Office.Common.XdmCommunicationManager.isTargetSubdomainOfSourceLocation(taskpaneUrl, dialogUrl);
            }
            var appDomains = OSF._OfficeAppFactory.getInitializationHelper()._appContext._appDomains;
            isAppDomain = Microsoft.Office.Common.XdmCommunicationManager.checkUrlWithAppDomains(appDomains, dialogUrl);
        }
        var logJsonAsString = "openDialog isInline: " + dialogInfo[OSF.ShowWindowDialogParameterKeys.DisplayInIframe].toString() + ", " + "taskpaneHostname: " + taskpaneUrlPortionAllowedToLog + ", " + "dialogHostName: " + dialogUrlPortionAllowedToLog + ", " + "isSameDomain: " + isSameDomain.toString() + ", " + "parentIsSubdomain: " + parentIsSubdomain.toString() + ", " + "childIsSubdomain: " + childIsSubdomain.toString() + ", " + "isAppDomain: " + isAppDomain.toString();
        OSF.AppTelemetry.logAppCommonMessage(logJsonAsString);
    }
    if (dialogUrl == null || !(dialogUrl.substr(0, httpsIdentifyString.length) === httpsIdentifyString)) {
        if (dialogUrl.substr(0, httpIdentifyString.length) === httpIdentifyString) {
            showDialogCallback(OSF.DDA.ErrorCodeManager.errorCodes.ooeRequireHTTPS);
        }
        else {
            showDialogCallback(OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidScheme);
        }
        return;
    }
    if (!dialogInfo[OSF.ShowWindowDialogParameterKeys.DisplayInIframe]) {
        OSF.DialogShownStatus.isWindowDialog = true;
        OfficeExt.AddinNativeAction.Dialog.setHandlerAndShowDialogCallback(function OSF_DDA_WACDelegate$RegisterEventAsync_OnEvent(payload) {
            if (args.onEvent) {
                args.onEvent(payload);
            }
            if (OSF.AppTelemetry) {
                OSF.AppTelemetry.onEventDone(args.dispId);
            }
        }, showDialogCallback);
        OfficeExt.AddinNativeAction.Dialog.handleNewWindowDialog(dialogInfo);
    }
    else {
        OSF.DialogShownStatus.isWindowDialog = false;
        OSF.DDA.WAC.Delegate.registerEventAsync(args);
    }
};
OSF.DDA.WAC.Delegate.validateTaskpaneDomain = function OSF_DDA_WAC_Delegate$ValidateTaskpaneDomain(taskpaneUrl, allowSubDomains) {
    return OfficeExt.AddinNativeAction.Dialog.validateTaskpaneDomain(taskpaneUrl, allowSubDomains);
};
OSF.DDA.WAC.Delegate.messageParent = function OSF_DDA_WAC_Delegate$MessageParent(args) {
    var targetOrigin = args.hostCallArgs["targetOrigin"];
    var hasTargetOrigin = !!targetOrigin;
    var allowSubDomains = true;
    if (hasTargetOrigin && targetOrigin != "*") {
        var localStorage = OfficeExt.AddinNativeAction.Dialog.getLocalStorage();
        var sessionStorage = OfficeExt.AddinNativeAction.Dialog.getSessionStorage();
        var storageKey = '_trusts_';
        var parsedTargetOriginUrl = OSF.OUtil.parseUrl(targetOrigin, true);
        var targetOriginUrlPrefix = parsedTargetOriginUrl.protocol + "//" + parsedTargetOriginUrl.hostname + (parsedTargetOriginUrl.port ? (":" + parsedTargetOriginUrl.port) : "");
        storageKey = storageKey + targetOriginUrlPrefix;
        OfficeExt.AddinNativeAction.Dialog.setWithExpiry(localStorage, storageKey, 'true');
        OfficeExt.AddinNativeAction.Dialog.setWithExpiry(sessionStorage, storageKey, 'true');
    }
    if (window.opener != null) {
        if (!hasTargetOrigin) {
            var taskpaneUrl = OSF._OfficeAppFactory.getInitializationHelper()._appContext._taskpaneUrl;
            if (!taskpaneUrl) {
                taskpaneUrl = window.location.origin;
            }
            else if (!OfficeExt.AddinNativeAction.Dialog.validateTaskpaneDomain(taskpaneUrl, allowSubDomains)) {
                var errorMessage = "messageParent called but the taskpane domain is untrusted: " + taskpaneUrl;
                if (OSF.AppTelemetry) {
                    OSF.AppTelemetry.logAppException(errorMessage);
                }
                throw new Error(errorMessage);
            }
            args.hostCallArgs["targetOrigin"] = taskpaneUrl;
        }
        OfficeExt.AddinNativeAction.Dialog.messageParent(args);
    }
    else {
        OSF.DDA.WAC.Delegate.executeAsync(args);
    }
};
OSF.DDA.WAC.Delegate.sendMessage = function OSF_DDA_WAC_Delegate$SendMessage(args) {
    if (OSF.DialogShownStatus.hasDialogShown) {
        if (OSF.DialogShownStatus.isWindowDialog) {
            OfficeExt.AddinNativeAction.Dialog.sendMessage(args);
        }
        else {
            OSF.DDA.WAC.Delegate.executeAsync(args);
        }
    }
};
OSF.DDA.WAC.Delegate.closeDialog = function OSF_DDA_WAC_Delegate$CloseDialog(args) {
    var callback = OSF.DDA.WAC.Delegate._getOnAfterRegisterEvent(false, args);
    function closeDialogCallback(status) {
        var payload = { "Error": status };
        try {
            callback(Microsoft.Office.Common.InvokeResultCode.noError, payload);
        }
        catch (e) {
            if (OSF.AppTelemetry) {
                OSF.AppTelemetry.logAppException("Exception happens at closeDialogCallback." + e);
            }
        }
    }
    if (!OSF.DialogShownStatus.hasDialogShown) {
        closeDialogCallback(OSF.DDA.ErrorCodeManager.errorCodes.ooeWebDialogClosed);
    }
    else {
        if (OSF.DialogShownStatus.isWindowDialog) {
            if (args.onCalling) {
                args.onCalling();
            }
            OfficeExt.AddinNativeAction.Dialog.closeDialog(closeDialogCallback);
        }
        else {
            OSF.DDA.WAC.Delegate.unregisterEventAsync(args);
        }
    }
};
OSF.InitializationHelper.prototype.dismissDialogNotification = function OSF_InitializationHelper$dismissDialogNotification() {
    OfficeExt.AddinNativeAction.Dialog.escDismissDialogNotification();
};
OSF.InitializationHelper.prototype.registerMessageReceivedEventForWindowDialog = function OSF_InitializationHelper$registerMessageReceivedEventForWindowDialog() {
    OfficeExt.AddinNativeAction.Dialog.registerMessageReceivedEvent();
};
OSF.DDA.AsyncMethodNames.addNames({
    CloseContainerAsync: "closeContainer"
});
(function (OfficeExt) {
    var Container = (function () {
        function Container(parameters) {
        }
        return Container;
    }());
    OfficeExt.Container = Container;
})(OfficeExt || (OfficeExt = {}));
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.CloseContainerAsync,
    requiredArguments: [],
    supportedOptions: [],
    privateStateCallbacks: []
});
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidCloseContainerMethod,
    fromHost: [],
    toHost: []
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { ItemChanged: "olkItemSelectedChanged" });
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, { OlkItemSelectedData: "OlkItemSelectedData" });
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { RecipientsChanged: "olkRecipientsChanged" });
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, { OlkRecipientsData: "OlkRecipientsData" });
OSF.DDA.OlkRecipientsChangedEventArgs = function OSF_DDA_OlkRecipientsChangedEventArgs(eventData) {
    var changedRecipientFields = eventData[OSF.DDA.EventDescriptors.OlkRecipientsData][0];
    if (changedRecipientFields === "") {
        changedRecipientFields = null;
    }
    OSF.OUtil.defineEnumerableProperties(this, {
        "type": {
            value: Microsoft.Office.WebExtension.EventType.RecipientsChanged
        },
        "changedRecipientFields": {
            value: JSON.parse(changedRecipientFields)
        }
    });
};
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { AppointmentTimeChanged: "olkAppointmentTimeChanged" });
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, { OlkAppointmentTimeChangedData: "OlkAppointmentTimeChangedData" });
OSF.DDA.OlkAppointmentTimeChangedEventArgs = function OSF_DDA_OlkAppointmentTimeChangedEventArgs(eventData) {
    var appointmentTimeString = eventData[OSF.DDA.EventDescriptors.OlkAppointmentTimeChangedData][0];
    var start;
    var end;
    try {
        var appointmentTime = JSON.parse(appointmentTimeString);
        start = new Date(appointmentTime.start).toISOString();
        end = new Date(appointmentTime.end).toISOString();
    }
    catch (e) {
        start = null;
        end = null;
    }
    OSF.OUtil.defineEnumerableProperties(this, {
        "type": {
            value: Microsoft.Office.WebExtension.EventType.AppointmentTimeChanged
        },
        "start": {
            value: start
        },
        "end": {
            value: end
        }
    });
};
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { RecurrenceChanged: "olkRecurrenceChanged" });
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, { OlkRecurrenceData: "OlkRecurrenceData" });
OSF.DDA.OlkRecurrenceChangedEventArgs = function OSF_DDA_OlkRecurrenceChangedEventArgs(eventData) {
    var recurrenceObject = null;
    try {
        var dataObject = JSON.parse(eventData[OSF.DDA.EventDescriptors.OlkRecurrenceChangedData][0]);
        if (dataObject.recurrence != null) {
            recurrenceObject = JSON.parse(dataObject.recurrence);
            recurrenceObject = Microsoft.Office.WebExtension.OutlookBase.SeriesTimeJsonConverter(recurrenceObject);
        }
    }
    catch (e) {
        recurrenceObject = null;
    }
    OSF.OUtil.defineEnumerableProperties(this, {
        "type": {
            value: Microsoft.Office.WebExtension.EventType.RecurrenceChanged
        },
        "recurrence": {
            value: recurrenceObject
        }
    });
};
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { OfficeThemeChanged: "officeThemeChanged" });
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, { OfficeThemeData: "OfficeThemeData" });
OSF.OUtil.setNamespace("Theming", OSF.DDA);
OSF.DDA.Theming.OfficeThemeChangedEventArgs = function OSF_DDA_Theming_OfficeThemeChangedEventArgs(officeTheme) {
    var themeData = JSON.parse(officeTheme.OfficeThemeData[0]);
    var themeDataHex = {};
    for (var color in themeData) {
        themeDataHex[color] = OSF.OUtil.convertIntToCssHexColor(themeData[color]);
    }
    OSF.OUtil.defineEnumerableProperties(this, {
        "type": {
            value: Microsoft.Office.WebExtension.EventType.OfficeThemeChanged
        },
        "officeTheme": {
            value: themeDataHex
        }
    });
};
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { AttachmentsChanged: "olkAttachmentsChanged" });
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, { OlkAttachmentsChangedData: "OlkAttachmentsChangedData" });
OSF.DDA.OlkAttachmentsChangedEventArgs = function OSF_DDA_OlkAttachmentsChangedEventArgs(eventData) {
    var attachmentStatus;
    var attachmentDetails;
    try {
        var attachmentChangedObject = JSON.parse(eventData[OSF.DDA.EventDescriptors.OlkAttachmentsChangedData][0]);
        attachmentStatus = attachmentChangedObject.attachmentStatus;
        attachmentDetails = Microsoft.Office.WebExtension.OutlookBase.CreateAttachmentDetails(attachmentChangedObject.attachmentDetails);
    }
    catch (e) {
        attachmentStatus = null;
        attachmentDetails = null;
    }
    OSF.OUtil.defineEnumerableProperties(this, {
        "type": {
            value: Microsoft.Office.WebExtension.EventType.AttachmentsChanged
        },
        "attachmentStatus": {
            value: attachmentStatus
        },
        "attachmentDetails": {
            value: attachmentDetails
        }
    });
};
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { EnhancedLocationsChanged: "olkEnhancedLocationsChanged" });
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, { OlkEnhancedLocationsChangedData: "OlkEnhancedLocationsChangedData" });
OSF.DDA.OlkEnhancedLocationsChangedEventArgs = function OSF_DDA_OlkEnhancedLocationsChangedEventArgs(eventData) {
    var enhancedLocations;
    try {
        var enhancedLocationsChangedObject = JSON.parse(eventData[OSF.DDA.EventDescriptors.OlkEnhancedLocationsChangedData][0]);
        enhancedLocations = enhancedLocationsChangedObject.enhancedLocations;
    }
    catch (e) {
        enhancedLocations = null;
    }
    OSF.OUtil.defineEnumerableProperties(this, {
        "type": {
            value: Microsoft.Office.WebExtension.EventType.EnhancedLocationsChanged
        },
        "enhancedLocations": {
            value: enhancedLocations
        }
    });
};
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { InfobarClicked: "olkInfobarClicked" });
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, { OlkInfobarClickedData: "OlkInfobarClickedData" });
OSF.DDA.OlkInfobarClickedEventArgs = function OSF_DDA_OlkInfobarClickedEventArgs(eventData) {
    var infobarDetails;
    try {
        infobarDetails = eventData[OSF.DDA.EventDescriptors.OlkInfobarClickedData][0];
    }
    catch (e) {
        infobarDetails = null;
    }
    OSF.OUtil.defineEnumerableProperties(this, {
        "type": {
            value: Microsoft.Office.WebExtension.EventType.InfobarClicked
        },
        "infobarDetails": {
            value: infobarDetails
        }
    });
};
OSF.DDA.OlkItemSelectedChangedEventArgs = function OSF_DDA_OlkItemSelectedChangedEventArgs(eventData) {
    var initialDataSource = eventData[OSF.DDA.EventDescriptors.OlkItemSelectedData];
    if (initialDataSource === "") {
        initialDataSource = null;
    }
    OSF.OUtil.defineEnumerableProperties(this, {
        "type": {
            value: Microsoft.Office.WebExtension.EventType.ItemChanged
        },
        "initialData": {
            value: initialDataSource
        }
    });
};
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkItemSelectedChangedEvent,
    fromHost: [{
            name: OSF.DDA.EventDescriptors.OlkItemSelectedData,
            value: OSF.DDA.WAC.Delegate.ParameterMap.self
        }],
    isComplexType: true
});
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkRecipientsChangedEvent,
    fromHost: [{
            name: OSF.DDA.EventDescriptors.OlkRecipientsData,
            value: OSF.DDA.WAC.Delegate.ParameterMap.self
        }],
    isComplexType: true
});
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkAppointmentTimeChangedEvent,
    fromHost: [{
            name: OSF.DDA.EventDescriptors.OlkAppointmentTimeChangedData,
            value: OSF.DDA.WAC.Delegate.ParameterMap.self
        }],
    isComplexType: true
});
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkRecurrenceChangedEvent,
    fromHost: [{
            name: OSF.DDA.EventDescriptors.OlkRecurrenceChangedData,
            value: OSF.DDA.WAC.Delegate.ParameterMap.self
        }],
    isComplexType: true
});
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOfficeThemeChangedEvent,
    fromHost: [{
            name: OSF.DDA.EventDescriptors.OfficeThemeData,
            value: OSF.DDA.WAC.Delegate.ParameterMap.self
        }],
    isComplexType: true
});
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkAttachmentsChangedEvent,
    fromHost: [{
            name: OSF.DDA.EventDescriptors.OlkAttachmentsChangedData,
            value: OSF.DDA.WAC.Delegate.ParameterMap.self
        }],
    isComplexType: true
});
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkEnhancedLocationsChangedEvent,
    fromHost: [{
            name: OSF.DDA.EventDescriptors.OlkEnhancedLocationsChangedData,
            value: OSF.DDA.WAC.Delegate.ParameterMap.self
        }],
    isComplexType: true
});
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkInfobarClickedEvent,
    fromHost: [{
            name: OSF.DDA.EventDescriptors.OlkInfobarClickedData,
            value: OSF.DDA.WAC.Delegate.ParameterMap.self
        }],
    isComplexType: true
});
Microsoft.Office.WebExtension.AccountTypeFilter = {
    NoFilter: "noFilter",
    AAD: "aad",
    MSA: "msa"
};
OSF.DDA.AsyncMethodNames.addNames({ GetAccessTokenAsync: "getAccessTokenAsync" });
OSF.DDA.Auth = function OSF_DDA_Auth() {
};
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.GetAccessTokenAsync,
    requiredArguments: [],
    supportedOptions: [
        {
            name: Microsoft.Office.WebExtension.Parameters.ForceConsent,
            value: {
                "types": ["boolean"],
                "defaultValue": false
            }
        },
        {
            name: Microsoft.Office.WebExtension.Parameters.ForceAddAccount,
            value: {
                "types": ["boolean"],
                "defaultValue": false
            }
        },
        {
            name: Microsoft.Office.WebExtension.Parameters.AuthChallenge,
            value: {
                "types": ["string"],
                "defaultValue": ""
            }
        },
        {
            name: Microsoft.Office.WebExtension.Parameters.AllowConsentPrompt,
            value: {
                "types": ["boolean"],
                "defaultValue": false
            }
        },
        {
            name: Microsoft.Office.WebExtension.Parameters.ForMSGraphAccess,
            value: {
                "types": ["boolean"],
                "defaultValue": false
            }
        },
        {
            name: Microsoft.Office.WebExtension.Parameters.AllowSignInPrompt,
            value: {
                "types": ["boolean"],
                "defaultValue": false
            }
        },
        {
            name: Microsoft.Office.WebExtension.Parameters.EnableNewHosts,
            value: {
                "types": ["number"],
                "defaultValue": 0
            }
        },
        {
            name: Microsoft.Office.WebExtension.Parameters.AccountTypeFilter,
            value: {
                "enum": Microsoft.Office.WebExtension.AccountTypeFilter,
                "defaultValue": Microsoft.Office.WebExtension.AccountTypeFilter.NoFilter
            }
        }
    ],
    checkCallArgs: function (callArgs, caller, stateInfo) {
        var _a;
        var appContext = OSF._OfficeAppFactory.getInitializationHelper()._appContext;
        if (appContext && appContext._wopiHostOriginForSingleSignOn) {
            var addinTrustId = OSF.OUtil.Guid.generateNewGuid();
            window.parent.parent.postMessage("{\"MessageId\":\"AddinTrustedOrigin\",\"AddinTrustId\":\"" + addinTrustId + "\"}", appContext._wopiHostOriginForSingleSignOn);
            callArgs[Microsoft.Office.WebExtension.Parameters.AddinTrustId] = addinTrustId;
        }
        if (window.Office.context.requirements.isSetSupported("JsonPayloadSSO")) {
            var jsonParameterMap = (_a = {},
                _a[Microsoft.Office.WebExtension.Parameters.ForceConsent] = false,
                _a[Microsoft.Office.WebExtension.Parameters.ForceAddAccount] = false,
                _a[Microsoft.Office.WebExtension.Parameters.AuthChallenge] = true,
                _a[Microsoft.Office.WebExtension.Parameters.AllowConsentPrompt] = true,
                _a[Microsoft.Office.WebExtension.Parameters.ForMSGraphAccess] = true,
                _a[Microsoft.Office.WebExtension.Parameters.AllowSignInPrompt] = true,
                _a[Microsoft.Office.WebExtension.Parameters.EnableNewHosts] = true,
                _a[Microsoft.Office.WebExtension.Parameters.AccountTypeFilter] = true,
                _a);
            var jsonPayload = {};
            for (var _i = 0, _b = Object.keys(jsonParameterMap); _i < _b.length; _i++) {
                var key = _b[_i];
                if (jsonParameterMap[key]) {
                    jsonPayload[key] = callArgs[key];
                }
                delete callArgs[key];
            }
            callArgs[Microsoft.Office.WebExtension.Parameters.JsonPayload] = JSON.stringify(jsonPayload);
        }
        return callArgs;
    },
    onSucceeded: function (dataDescriptor, caller, callArgs) {
        var data = dataDescriptor[Microsoft.Office.WebExtension.Parameters.Data];
        return data;
    }
});
OSF.OUtil.setNamespace("Marshaling", OSF.DDA);
OSF.OUtil.setNamespace("SingleSignOn", OSF.DDA.Marshaling);
OSF.DDA.Marshaling.SingleSignOn.GetAccessTokenKeys = {
    ForceConsent: "forceConsent",
    ForceAddAccount: "forceAddAccount",
    AuthChallenge: "authChallenge",
    AllowConsentPrompt: "allowConsentPrompt",
    ForMSGraphAccess: "forMSGraphAccess",
    AllowSignInPrompt: "allowSignInPrompt",
    EnableNewHosts: "enableNewHosts",
    AccountTypeFilter: "accountTypeFilter",
    AddinTrustId: "addinTrustId"
};
OSF.DDA.Marshaling.SingleSignOn.AccessTokenResultKeys = {
    AccessToken: "accessToken"
};
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidGetAccessTokenMethod,
    toHost: [
        { name: Microsoft.Office.WebExtension.Parameters.ForceConsent, value: OSF.DDA.Marshaling.SingleSignOn.GetAccessTokenKeys.ForceConsent },
        { name: Microsoft.Office.WebExtension.Parameters.ForceAddAccount, value: OSF.DDA.Marshaling.SingleSignOn.GetAccessTokenKeys.ForceAddAccount },
        { name: Microsoft.Office.WebExtension.Parameters.AuthChallenge, value: OSF.DDA.Marshaling.SingleSignOn.GetAccessTokenKeys.AuthChallenge },
        { name: Microsoft.Office.WebExtension.Parameters.AllowConsentPrompt, value: OSF.DDA.Marshaling.SingleSignOn.GetAccessTokenKeys.AllowConsentPrompt },
        { name: Microsoft.Office.WebExtension.Parameters.ForMSGraphAccess, value: OSF.DDA.Marshaling.SingleSignOn.GetAccessTokenKeys.ForMSGraphAccess },
        { name: Microsoft.Office.WebExtension.Parameters.AllowSignInPrompt, value: OSF.DDA.Marshaling.SingleSignOn.GetAccessTokenKeys.AllowSignInPrompt },
        { name: Microsoft.Office.WebExtension.Parameters.EnableNewHosts, value: OSF.DDA.Marshaling.SingleSignOn.GetAccessTokenKeys.EnableNewHosts },
        { name: Microsoft.Office.WebExtension.Parameters.AccountTypeFilter, value: OSF.DDA.Marshaling.SingleSignOn.GetAccessTokenKeys.AccountTypeFilter },
        { name: Microsoft.Office.WebExtension.Parameters.AddinTrustId, value: OSF.DDA.Marshaling.SingleSignOn.GetAccessTokenKeys.AddinTrustId }
    ],
    fromHost: [
        { name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.Marshaling.SingleSignOn.AccessTokenResultKeys.AccessToken }
    ]
});
window.OfficeRuntime = window.OfficeRuntime || {};
window.OfficeRuntime.auth = {
    getAccessToken: function (params) {
        var internalPromise = window.Promise ? window.Promise : window.Office.Promise;
        return new internalPromise(function (resolve, reject) {
            try {
                window.Office.context.auth.getAccessTokenAsync(params || {}, function (result) {
                    if (result.status === "succeeded") {
                        resolve(result.value);
                    }
                    else {
                        reject(result.error);
                    }
                });
            }
            catch (error) {
                reject(error);
            }
        });
    }
};
OSF.InitializationHelper.prototype.prepareRightBeforeWebExtensionInitialize = function OSF_InitializationHelper$prepareRightBeforeWebExtensionInitialize(appContext) {
    OSF.WebApp._UpdateLinksForHostAndXdmInfo();
};
OSF.InitializationHelper.prototype.prepareRightAfterWebExtensionInitialize = function () {
    var appCommandHandler = OfficeExt.AppCommand.AppCommandManager.instance();
    appCommandHandler.initializeAndChangeOnce();
};
OSF.InitializationHelper.prototype.getInitializationReason = function OSF_InitializationHelper$getInitializationReason(appContext) {
    return appContext.get_reason();
};
var executeAsyncBase = OSF.DDA.WAC.Delegate.executeAsync;
OSF.DDA.WAC.Delegate.executeAsync = function OSF_DDA_WAC_Delegate$executeAsyncOverride(args) {
    var onCallingBase = args.onCalling;
    args.onCalling = function OSF_DDA_WAC_Delegate$executeAsync$onCalling() {
        args.hostCallArgs = OSF.DDA.OutlookAppOm.addAdditionalArgs(args.dispId, args.hostCallArgs);
        onCallingBase && onCallingBase();
    };
    executeAsyncBase(args);
};
;

OSF.InitializationHelper.prototype.isHostOriginTrusted = function OSF_InitializationHelper$isHostOriginTrusted(hostOrigin) {
    if (!hostOrigin || hostOrigin === "null") {
        return false;
    }

    var regexHostNameStringArray = [
        "^outlook\\.office\\.com$",
        "^outlook-sdf\\.office\\.com$",
        "^outlook\\.office\\.com$",
        "^outlook-sdf\\.office\\.com$",
        "^outlook\\.live\\.com$",
        "^outlook-sdf\\.live\\.com$",
        "^consumer\\.live-int\\.com$",
        "^outlook-tdf\\.live\\.com$",
        "^sdfpilot\\.live\\.com$",
        "^outlook\\.office\\.de$",
        "^outlook\\.office365\\.us$",
        "^outlook\\.office365\\.com$",
        "^partner\\.outlook\\.cn$",
        "^exchangelabs\\.live-int\\.com$"
    ];

    var regexHostName = new RegExp(regexHostNameStringArray.join("|"));
    return regexHostName.test(hostOrigin);
};  

OSF.InitializationHelper.prototype.prepareApiSurface = function OSF_InitializationHelper$prepareApiSurface(appContext)
{
    var license = new OSF.DDA.License(appContext.get_eToken());
    if ((appContext.get_appName() == OSF.AppName.OutlookWebApp)) {
        OSF.WebApp._UpdateLinksForHostAndXdmInfo();
        this.initWebDialog(appContext);
        this.initWebAuth(appContext);
        OSF._OfficeAppFactory.setContext(new OSF.DDA.OutlookContext(appContext, this._settings, license, appContext.appOM));
        OSF._OfficeAppFactory.setHostFacade(new OSF.DDA.DispIdHost.Facade(OSF.DDA.WAC.getDelegateMethods,OSF.DDA.WAC.Delegate.ParameterMap));
	}
    else {
        OfficeJsClient_OutlookWin32.prepareApiSurface(appContext); 
        OSF._OfficeAppFactory.setContext(new OSF.DDA.OutlookContext(appContext, this._settings, license, appContext.appOM, OSF.DDA.OfficeTheme ? OSF.DDA.OfficeTheme.getOfficeTheme : null, appContext.ui));
        OSF._OfficeAppFactory.setHostFacade(new OSF.DDA.DispIdHost.Facade(OSF.DDA.DispIdHost.getClientDelegateMethods,OSF.DDA.SafeArray.Delegate.ParameterMap));
    }
}
OSF.DDA.SettingsManager = {
    SerializedSettings: "serializedSettings",
    DateJSONPrefix: "Date(",
    DataJSONSuffix: ")",
    serializeSettings: function OSF_DDA_SettingsManager$serializeSettings(settingsCollection) {
        var ret = {};
        for (var key in settingsCollection) {
            var value = settingsCollection[key];
            try  {
                if (JSON) {
                    value = JSON.stringify(value, function dateReplacer(k, v) {
                        return OSF.OUtil.isDate(this[k]) ? OSF.DDA.SettingsManager.DateJSONPrefix + this[k].getTime() + OSF.DDA.SettingsManager.DataJSONSuffix : v;
                    });
                } else {
                    value = Sys.Serialization.JavaScriptSerializer.serialize(value);
                }
                ret[key] = value;
            } catch (ex) {
            }
        }
        return ret;
    },
    deserializeSettings: function OSF_DDA_SettingsManager$deserializeSettings(serializedSettings) {
        var ret = {};
        serializedSettings = serializedSettings || {};
        for (var key in serializedSettings) {
            var value = serializedSettings[key];
            try  {
                if (JSON) {
                    value = JSON.parse(value, function dateReviver(k, v) {
                        var d;
                        if (typeof v === 'string' && v && v.length > 6 && v.slice(0, 5) === OSF.DDA.SettingsManager.DateJSONPrefix && v.slice(-1) === OSF.DDA.SettingsManager.DataJSONSuffix) {
                            d = new Date(parseInt(v.slice(5, -1)));
                            if (d) {
                                return d;
                            }
                        }
                        return v;
                    });
                } else {
                    value = Sys.Serialization.JavaScriptSerializer.deserialize(value, true);
                }
                ret[key] = value;
            } catch (ex) {
            }
        }
        return ret;
    }
};
OSF.InitializationHelper.prototype.loadAppSpecificScriptAndCreateOM = function OSF_InitializationHelper$loadAppSpecificScriptAndCreateOM(appContext, appReady, basePath) {

var Outlook = typeof Outlook === "object" ? Outlook : {}; Outlook["OutlookAppOm"] =
/******/ (function(modules) { // webpackBootstrap
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
/******/ 	__webpack_require__.p = "/";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 2);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, exports) {

module.exports = OSF;

/***/ }),
/* 1 */
/***/ (function(module, exports) {

module.exports = Microsoft;

/***/ }),
/* 2 */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);

// CONCATENATED MODULE: ./src/utils/isNullOrUndefined.ts
function isNullOrUndefined(value) {
  return value === null || value === undefined;
}
// CONCATENATED MODULE: ./src/types/ExtensibilityStrings.ts

var OfficeStringJS = "office_strings.js";
var OfficeStringDebugJS = "office_strings.debug.js";
var ExtensibilityStringJS = "outlook_strings.js";
var tempWindow = window;
var ExtensibilityStrings;
function getString(string) {
  return ExtensibilityStrings[string];
}
var ExtensibilityStrings_url = "";
var baseUrl = "";
var scriptElement = null;
var stringLoadedCallback;
var stringsAreLoaded = false;

function createScriptElement(url) {
  var scriptElement = document.createElement("script");
  scriptElement.type = "text/javascript";
  scriptElement.src = url;
  return scriptElement;
}

function loadLocalizedScript(initializeAppCallback) {
  stringLoadedCallback = initializeAppCallback;
  var officeIndex;
  var scripts = document.getElementsByTagName("script");

  for (var i = 0; i < scripts.length; i++) {
    var tag = scripts.item(i);

    if (tag && tag.src) {
      var filename = tag.src || "";
      filename = filename.toLowerCase();
      officeIndex = filename.indexOf(OfficeStringJS);

      if (filename && officeIndex > 0) {
        ExtensibilityStrings_url = filename.replace(OfficeStringJS, ExtensibilityStringJS);
        baseUrl = saveBaseUrl(baseUrl, officeIndex, filename);
        break;
      }

      officeIndex = filename.indexOf(OfficeStringDebugJS);

      if (filename && officeIndex > 0) {
        ExtensibilityStrings_url = filename.replace(OfficeStringDebugJS, ExtensibilityStringJS);
        baseUrl = saveBaseUrl(baseUrl, officeIndex, filename);
        break;
      }
    }
  }

  if (ExtensibilityStrings_url) {
    var head = document.getElementsByTagName("head")[0];
    scriptElement = createScriptElement(ExtensibilityStrings_url);
    scriptElement.onload = scriptElementCallback;
    scriptElement.onreadystatechange = scriptElementCallback;
    window.setTimeout(failureCallback, 2000);
    head.appendChild(scriptElement);
  }
}

function scriptElementCallback() {
  stringsAreLoaded = true;

  if (!isNullOrUndefined(stringLoadedCallback) && (isNullOrUndefined(scriptElement.readyState) || !isNullOrUndefined(scriptElement.readyState) && (scriptElement.readyState === "loaded" || scriptElement.readyState === "complete"))) {
    scriptElement.onload = null;
    scriptElement.onreadystatechange = null;

    if (typeof tempWindow._u !== "undefined") {
      ExtensibilityStrings = tempWindow._u.ExtensibilityStrings;
    }

    stringLoadedCallback();
  }
}

function failureCallback() {
  if (!stringsAreLoaded) {
    var head = document.getElementsByTagName("head")[0];
    var fallbackUrl = baseUrl + "en-us/" + ExtensibilityStringJS;
    scriptElement.onload = null;
    scriptElement.onreadystatechange = null;
    scriptElement = createScriptElement(fallbackUrl);
    scriptElement.onload = scriptElementCallback;
    scriptElement.onreadystatechange = scriptElementCallback;
    head.appendChild(scriptElement);
  }
}

function saveBaseUrl(baseUrl, officeIndex, filename) {
  var languageUrl = filename.substring(0, officeIndex);
  var lastIndexOfSlash = languageUrl.lastIndexOf("/", languageUrl.length - 2);

  if (lastIndexOfSlash === -1) {
    lastIndexOfSlash = languageUrl.lastIndexOf("\\", languageUrl.length - 2);
  }

  if (lastIndexOfSlash !== -1 && languageUrl.length > lastIndexOfSlash + 1) {
    baseUrl = languageUrl.substring(0, lastIndexOfSlash + 1);
  }

  return baseUrl;
}
// CONCATENATED MODULE: ./src/utils/ApiTelemetryConstants.ts
var ApiTelemetryCode = function () {
  function ApiTelemetryCode() {}

  ApiTelemetryCode.success = 0;
  ApiTelemetryCode.noResponseDictionary = -900;
  ApiTelemetryCode.noErrorCodeForStandardInvokeMethod = -901;
  ApiTelemetryCode.genericProxyError = -902;
  ApiTelemetryCode.genericLegacyApiError = -903;
  ApiTelemetryCode.genericUnknownError = -904;
  return ApiTelemetryCode;
}();


// CONCATENATED MODULE: ./src/utils/getErrorForTelemetry.ts


var getErrorForTelemetry_getErrorForTelemetry = function getErrorForTelemetry(resultCode, responseDictionary) {
  if (responseDictionary) {
    if ("error" in responseDictionary) {
      if (!responseDictionary["error"]) return ApiTelemetryCode.success;
      if ("errorCode" in responseDictionary) return responseDictionary["errorCode"];else return ApiTelemetryCode.noErrorCodeForStandardInvokeMethod;
    }

    if ("wasProxySuccessful" in responseDictionary) return responseDictionary["wasProxySuccessful"] ? ApiTelemetryCode.success : ApiTelemetryCode.genericProxyError;
    if ("wasSuccessful" in responseDictionary) return responseDictionary["wasSuccessful"] ? ApiTelemetryCode.success : ApiTelemetryCode.genericLegacyApiError;
  }

  if (!isNullOrUndefined(resultCode)) return resultCode;
  return ApiTelemetryCode.genericUnknownError;
};
// CONCATENATED MODULE: ./src/utils/isOwaOnly.ts
var isOwaOnly = function isOwaOnly(dispid) {
  switch (dispid) {
    case 402:
    case 401:
    case 400:
    case 403:
      return true;

    default:
      return false;
  }
};
// CONCATENATED MODULE: ./src/utils/InvokeResultCode.ts
var InvokeResultCode;

(function (InvokeResultCode) {
  InvokeResultCode[InvokeResultCode["noError"] = 0] = "noError";
  InvokeResultCode[InvokeResultCode["errorInRequest"] = -1] = "errorInRequest";
  InvokeResultCode[InvokeResultCode["errorHandlingRequest"] = -2] = "errorHandlingRequest";
  InvokeResultCode[InvokeResultCode["errorInResponse"] = -3] = "errorInResponse";
  InvokeResultCode[InvokeResultCode["errorHandlingResponse"] = -4] = "errorHandlingResponse";
  InvokeResultCode[InvokeResultCode["errorHandlingRequestAccessDenied"] = -5] = "errorHandlingRequestAccessDenied";
  InvokeResultCode[InvokeResultCode["errorHandlingMethodCallTimedout"] = -6] = "errorHandlingMethodCallTimedout";
})(InvokeResultCode || (InvokeResultCode = {}));
// CONCATENATED MODULE: ./src/utils/getErrorArgs.ts


var getErrorArgs_OSF = __webpack_require__(0);

var isInitialized = false;
function getErrorArgs(detailedErrorCode) {
  if (!isInitialized) {
    initialize();
  }

  return getErrorArgs_OSF.DDA.ErrorCodeManager.getErrorArgs(detailedErrorCode);
}
var totalRecipientsLimit = 500;
var sessionDataLengthLimit = 50000;
function initialize() {
  addErrorMessage(9000, "AttachmentSizeExceeded", getString("l_AttachmentExceededSize_Text"));
  addErrorMessage(9001, "NumberOfAttachmentsExceeded", getString("l_ExceededMaxNumberOfAttachments_Text"));
  addErrorMessage(9002, "InternalFormatError", getString("l_InternalFormatError_Text"));
  addErrorMessage(9003, "InvalidAttachmentId", getString("l_InvalidAttachmentId_Text"));
  addErrorMessage(9004, "InvalidAttachmentPath", getString("l_InvalidAttachmentPath_Text"));
  addErrorMessage(9005, "CannotAddAttachmentBeforeUpgrade", getString("l_CannotAddAttachmentBeforeUpgrade_Text"));
  addErrorMessage(9006, "AttachmentDeletedBeforeUploadCompletes", getString("l_AttachmentDeletedBeforeUploadCompletes_Text"));
  addErrorMessage(9007, "AttachmentUploadGeneralFailure", getString("l_AttachmentUploadGeneralFailure_Text"));
  addErrorMessage(9008, "AttachmentToDeleteDoesNotExist", getString("l_DeleteAttachmentDoesNotExist_Text"));
  addErrorMessage(9009, "AttachmentDeleteGeneralFailure", getString("l_AttachmentDeleteGeneralFailure_Text"));
  addErrorMessage(9010, "InvalidEndTime", getString("l_InvalidEndTime_Text"));
  addErrorMessage(9011, "HtmlSanitizationFailure", getString("l_HtmlSanitizationFailure_Text"));
  addErrorMessage(9012, "NumberOfRecipientsExceeded", getString("l_NumberOfRecipientsExceeded_Text").replace("{0}", totalRecipientsLimit));
  addErrorMessage(9013, "NoValidRecipientsProvided", getString("l_NoValidRecipientsProvided_Text"));
  addErrorMessage(9014, "CursorPositionChanged", getString("l_CursorPositionChanged_Text"));
  addErrorMessage(9016, "InvalidSelection", getString("l_InvalidSelection_Text"));
  addErrorMessage(9017, "AccessRestricted", "");
  addErrorMessage(9018, "GenericTokenError", "");
  addErrorMessage(9019, "GenericSettingsError", "");
  addErrorMessage(9020, "GenericResponseError", "");
  addErrorMessage(9021, "SaveError", getString("l_SaveError_Text"));
  addErrorMessage(9022, "MessageInDifferentStoreError", getString("l_MessageInDifferentStoreError_Text"));
  addErrorMessage(9023, "DuplicateNotificationKey", getString("l_DuplicateNotificationKey_Text"));
  addErrorMessage(9024, "NotificationKeyNotFound", getString("l_NotificationKeyNotFound_Text"));
  addErrorMessage(9025, "NumberOfNotificationsExceeded", getString("l_NumberOfNotificationsExceeded_Text"));
  addErrorMessage(9026, "PersistedNotificationArrayReadError", getString("l_PersistedNotificationArrayReadError_Text"));
  addErrorMessage(9027, "PersistedNotificationArraySaveError", getString("l_PersistedNotificationArraySaveError_Text"));
  addErrorMessage(9028, "CannotPersistPropertyInUnsavedDraftError", getString("l_CannotPersistPropertyInUnsavedDraftError_Text"));
  addErrorMessage(9029, "CanOnlyGetTokenForSavedItem", getString("l_CallSaveAsyncBeforeToken_Text"));
  addErrorMessage(9030, "APICallFailedDueToItemChange", getString("l_APICallFailedDueToItemChange_Text"));
  addErrorMessage(9031, "InvalidParameterValueError", getString("l_InvalidParameterValueError_Text"));
  addErrorMessage(9032, "ApiCallNotSupportedByExtensionPoint", getString("l_API_Not_Supported_By_ExtensionPoint_Error_Text"));
  addErrorMessage(9033, "SetRecurrenceOnInstanceError", getString("l_Recurrence_Error_Instance_SetAsync_Text"));
  addErrorMessage(9034, "InvalidRecurrenceError", getString("l_Recurrence_Error_Properties_Invalid_Text"));
  addErrorMessage(9035, "RecurrenceZeroOccurrences", getString("l_RecurrenceErrorZeroOccurrences_Text"));
  addErrorMessage(9036, "RecurrenceMaxOccurrences", getString("l_RecurrenceErrorMaxOccurrences_Text"));
  addErrorMessage(9037, "RecurrenceInvalidTimeZone", getString("l_RecurrenceInvalidTimeZone_Text"));
  addErrorMessage(9038, "InsufficientItemPermissionsError", getString("l_Insufficient_Item_Permissions_Text"));
  addErrorMessage(9039, "RecurrenceUnsupportedAlternateCalendar", getString("l_RecurrenceUnsupportedAlternateCalendar_Text"));
  addErrorMessage(9040, "HTTPRequestFailure", getString("l_Olk_Http_Error_Text"));
  addErrorMessage(9041, "NetworkError", getString("l_Internet_Not_Connected_Error_Text"));
  addErrorMessage(9042, "InternalServerError", getString("l_Internal_Server_Error_Text"));
  addErrorMessage(9043, "AttachmentTypeNotSupported", getString("l_AttachmentNotSupported_Text"));
  addErrorMessage(9044, "InvalidCategory", getString("l_Invalid_Category_Error_Text"));
  addErrorMessage(9045, "DuplicateCategory", getString("l_Duplicate_Category_Error_Text"));
  addErrorMessage(9046, "ItemNotSaved", getString("l_Item_Not_Saved_Error_Text"));
  addErrorMessage(9047, "MissingExtendedPermissionsForAPIError", getString("l_Missing_Extended_Permissions_For_API"));
  addErrorMessage(9048, "TokenAccessDenied", getString("l_TokenAccessDeniedWithoutItemContext_Text"));
  addErrorMessage(9049, "ItemNotFound", getString("l_ItemNotFound_Text"));
  addErrorMessage(9050, "KeyNotFound", getString("l_KeyNotFound_Text"));
  addErrorMessage(9051, "SessionObjectMaxLengthExceeded", getString("l_SessionDataObjectMaxLengthExceeded_Text").replace("{0}", sessionDataLengthLimit));
  addErrorMessage(9052, "AttachmentResourceNotFound", getString("l_Attachment_Resource_Not_Found"));
  addErrorMessage(9053, "AttachmentResourceUnAuthorizedAccess", getString("l_Attachment_Resource_UnAuthorizedAccess"));
  addErrorMessage(9054, "AttachmentDownloadFailed", getString("l_Attachment_Download_Failed_Generic_Error"));
  addErrorMessage(9055, "APINotSupportedForSharedFolders", getString("l_API_Not_Supported_For_Shared_Folders_Error"));
  isInitialized = true;
}
function addErrorMessage(code, error, message) {
  getErrorArgs_OSF.DDA.ErrorCodeManager.addErrorMessage(code, {
    name: error,
    message: message
  });
}
// CONCATENATED MODULE: ./src/utils/AdditionalGlobalParameters.ts
var additionalOutlookGlobalParameters;
var getAdditionalGlobalParametersSingleton = function getAdditionalGlobalParametersSingleton() {
  return additionalOutlookGlobalParameters;
};
var recreateAdditionalGlobalParametersSingleton = function recreateAdditionalGlobalParametersSingleton(parameterBlobSupported) {
  additionalOutlookGlobalParameters = new AdditionalGlobalParameters();
  additionalOutlookGlobalParameters.parameterBlobSupported = true;
  return additionalOutlookGlobalParameters;
};

var AdditionalGlobalParameters = function () {
  function AdditionalGlobalParameters() {
    this._parameterBlobSupported = true;
    this._itemNumber = 0;
    additionalOutlookGlobalParameters = this;
  }

  Object.defineProperty(AdditionalGlobalParameters.prototype, "parameterBlobSupported", {
    set: function set(supported) {
      this._parameterBlobSupported = supported;
    },
    enumerable: true,
    configurable: true
  });

  AdditionalGlobalParameters.prototype.setActionsDefinition = function (actionsDefinitionIn) {
    this._actionsDefinition = actionsDefinitionIn;
  };

  AdditionalGlobalParameters.prototype.setCurrentItemNumber = function (itemNumberIn) {
    if (itemNumberIn > 0) {
      this._itemNumber = itemNumberIn;
    }
  };

  Object.defineProperty(AdditionalGlobalParameters.prototype, "itemNumber", {
    get: function get() {
      return this._itemNumber;
    },
    enumerable: true,
    configurable: true
  });
  Object.defineProperty(AdditionalGlobalParameters.prototype, "actionsDefinition", {
    get: function get() {
      return this._actionsDefinition;
    },
    enumerable: true,
    configurable: true
  });

  AdditionalGlobalParameters.prototype.updateOutlookExecuteParameters = function (executeParameters, additionalApiParameters) {
    var outParameters = executeParameters;

    if (this._parameterBlobSupported) {
      if (this._itemNumber > 0) {
        additionalApiParameters.itemNumber = this._itemNumber.toString();
      }

      if (this._actionsDefinition != null) {
        additionalApiParameters.actions = this.actionsDefinition;
      }

      if (Object.keys(additionalApiParameters).length === 0) {
        return outParameters;
      }

      if (outParameters == null) {
        outParameters = [];
      }

      outParameters.push(JSON.stringify(additionalApiParameters));
    }

    return outParameters;
  };

  return AdditionalGlobalParameters;
}();


// CONCATENATED MODULE: ./src/utils/callOutlookNativeDispatcher.ts


var callOutlookNativeDispatcher_OSF = __webpack_require__(0);

var callOutlookNativeDispatcher = function callOutlookNativeDispatcher(dispid, data, responseCallback) {
  var executeParameters = callOutlookNativeDispatcher_convertToOutlookNativeParameters(dispid, data);
  callOutlookNativeDispatcher_OSF.ClientHostController.execute(dispid, executeParameters, function (nativeData, resultCode) {
    var responseData = nativeData.toArray();
    var deserializedData = callOutlookNativeDispatcher_deserializeResponseData(responseData);

    if (responseCallback != null) {
      responseCallback(resultCode, deserializedData);
    }
  });
};
var callOutlookNativeDispatcher_deserializeResponseData = function deserializeResponseData(responseData) {
  if (responseData.length == 0) {
    return null;
  }

  var itemNumberFromOutlookResponse = getItemNumberFromOutlookResponse(responseData);
  var isValidItemNumberFromOutlookResponse = itemNumberFromOutlookResponse > 0;
  var itemNumberInternal = 0;

  if (getAdditionalGlobalParametersSingleton()) {
    itemNumberInternal = getAdditionalGlobalParametersSingleton().itemNumber;
  }

  var isValidItemNumberInternal = itemNumberInternal > 0;
  var itemChanged = isValidItemNumberFromOutlookResponse && isValidItemNumberInternal && itemNumberFromOutlookResponse > itemNumberInternal;
  return createDeserializedData(responseData, itemChanged);
};
var callOutlookNativeDispatcher_convertToOutlookNativeParameters = function convertToOutlookNativeParameters(dispid, data) {
  var executeParameters = null;
  var optionalParameters = {};

  switch (dispid) {
    case 12:
      optionalParameters.isRest = data.isRest;
      break;

    case 4:
      {
        var jsonProperty = JSON.stringify(data.customProperties);
        executeParameters = [jsonProperty];
        break;
      }

    case 5:
      executeParameters = new Array(data.body);
      break;

    case 8:
    case 9:
    case 179:
    case 180:
      executeParameters = new Array(data.itemId);
      break;

    case 7:
    case 177:
      executeParameters = new Array(convertRecipientArrayParameterForOutlookForDisplayApi(data.requiredAttendees), convertRecipientArrayParameterForOutlookForDisplayApi(data.optionalAttendees), data.start, data.end, data.location, convertRecipientArrayParameterForOutlookForDisplayApi(data.resources), data.subject, data.body);
      break;

    case 44:
    case 178:
      executeParameters = [convertRecipientArrayParameterForOutlookForDisplayApi(data.toRecipients), convertRecipientArrayParameterForOutlookForDisplayApi(data.ccRecipients), convertRecipientArrayParameterForOutlookForDisplayApi(data.bccRecipients), data.subject, data.htmlBody, data.attachments];
      break;

    case 43:
      executeParameters = [data.ewsIdOrEmail];
      break;

    case 45:
      executeParameters = [data.module, data.queryString];
      break;

    case 40:
      executeParameters = [data.extensionId, data.consentState];
      break;

    case 11:
    case 10:
    case 184:
    case 183:
      executeParameters = [data.htmlBody];
      break;

    case 31:
    case 30:
    case 182:
    case 181:
      executeParameters = [data.htmlBody, data.attachments];
      break;

    case 23:
    case 13:
    case 38:
    case 29:
      executeParameters = [data.data, data.coercionType];
      break;

    case 37:
    case 28:
      executeParameters = [data.coercionType];
      break;

    case 17:
      executeParameters = [data.subject];
      break;

    case 15:
      executeParameters = [data.recipientField];
      break;

    case 22:
    case 21:
      executeParameters = [data.recipientField, convertComposeEmailDictionaryParameterForSetApi(data.recipientArray)];
      break;

    case 19:
      executeParameters = [data.itemId, data.name];
      break;

    case 16:
      executeParameters = [data.uri, data.name, data.isInline];
      break;

    case 148:
      executeParameters = [data.base64String, data.name, data.isInline];
      break;

    case 20:
      executeParameters = [data.attachmentIndex];
      break;

    case 25:
      executeParameters = [data.TimeProperty, data.time];
      break;

    case 24:
      executeParameters = [data.TimeProperty];
      break;

    case 27:
      executeParameters = [data.location];
      break;

    case 33:
    case 35:
      executeParameters = [data.key, data.type, data.persistent, data.message, data.icon];
      getAdditionalGlobalParametersSingleton().setActionsDefinition(data.actions);
      break;

    case 36:
      executeParameters = [data.key];
      break;

    default:
      optionalParameters = data || {};
      break;
  }

  if (dispid !== 1) {
    executeParameters = getAdditionalGlobalParametersSingleton().updateOutlookExecuteParameters(executeParameters, optionalParameters);
  }

  return executeParameters;
};

var convertRecipientArrayParameterForOutlookForDisplayApi = function convertRecipientArrayParameterForOutlookForDisplayApi(recipients) {
  return recipients != null ? recipients.join(";") : "";
};

var convertComposeEmailDictionaryParameterForSetApi = function convertComposeEmailDictionaryParameterForSetApi(recipients) {
  var results = [];

  if (recipients == null) {
    return results;
  }

  for (var i = 0; i < recipients.length; i++) {
    var newRecipient = [recipients[i].address, recipients[i].name];
    results.push(newRecipient);
  }

  return results;
};

var getItemNumberFromOutlookResponse = function getItemNumberFromOutlookResponse(responseData) {
  var itemNumber = 0;

  if (responseData.length > 2) {
    var extraParameters = JSON.parse(responseData[2]);

    if (!!extraParameters && typeof extraParameters === "object") {
      itemNumber = extraParameters.itemNumber;
    }
  }

  return itemNumber;
};
var createDeserializedData = function createDeserializedData(responseData, itemChanged) {
  var deserializedData = null;
  var returnValues = JSON.parse(responseData[0]);

  if (typeof returnValues === "number") {
    deserializedData = createDeserializedDataWithInt(responseData, itemChanged);
  } else if (!!returnValues && typeof returnValues === "object") {
    deserializedData = createDeserializedDataWithDictionary(responseData, itemChanged);
  } else {
    throw new Error("Return data type from host must be Object or Number");
  }

  return deserializedData;
};

var createDeserializedDataWithDictionary = function createDeserializedDataWithDictionary(responseData, itemChanged) {
  var deserializedData = JSON.parse(responseData[0]);

  if (itemChanged) {
    deserializedData.error = true;
    deserializedData.errorCode = 9030;
  } else if (responseData.length > 1 && responseData[1] !== 0) {
    deserializedData.error = true;
    deserializedData.errorCode = responseData[1];

    if (responseData.length > 2) {
      var diagnosticsData = JSON.parse(responseData[2]);
      deserializedData.diagnostics = diagnosticsData["Diagnostics"];
    }

    if (responseData.length >= 5) {
      deserializedData.errorMessage = responseData[3];
      deserializedData.errorName = responseData[4];
    }
  } else {
    deserializedData.error = false;
  }

  return deserializedData;
};

var createDeserializedDataWithInt = function createDeserializedDataWithInt(responseData, itemChanged) {
  var deserializedData = {};
  deserializedData.error = true;
  deserializedData.errorCode = responseData[0];
  return deserializedData;
};
// CONCATENATED MODULE: ./src/utils/isOutlookJs.ts
var outlookJs;
outlookJs = false;
var isOutlookJs = function isOutlookJs() {
  return outlookJs;
};
// CONCATENATED MODULE: ./src/api/standardInvokeHostMethod.ts








var standardInvokeHostMethod_OSF = __webpack_require__(0);

function standardInvokeHostMethod(dispid, userContext, callback, data, format, customResponse) {
  standardInvokeHostMethod_invokeHostMethod(dispid, data, function (resultCode, response) {
    if (callback) {
      var asyncResult = void 0;
      var wasSuccessful = true;

      if (typeof response === "object" && response !== null) {
        if (response.wasSuccessful !== undefined) {
          wasSuccessful = response.wasSuccessful;
        }

        if (response.error !== undefined || response.errorCode !== undefined || response.data !== undefined) {
          if (!response.error) {
            var formattedData = format ? format(response.data) : response.data;
            asyncResult = createAsyncResult(formattedData, standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.ErrorCode.Success, 0, userContext);
          } else {
            var errorCode = response.errorCode;
            asyncResult = createAsyncResult(undefined, standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, errorCode, userContext);
          }
        }

        if (customResponse) {
          asyncResult = customResponse(response, userContext, resultCode);
        }

        if (!asyncResult && resultCode !== InvokeResultCode.noError) {
          asyncResult = createAsyncResult(undefined, standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 9002, userContext);
        }

        if (!asyncResult && resultCode === InvokeResultCode.noError && wasSuccessful === false) {
          asyncResult = createAsyncResult(undefined, standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, standardInvokeHostMethod_OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupported, userContext);
        }

        if (!asyncResult && !customResponse && !response.errorCode && wasSuccessful === true) {
          var formattedData = format ? format(response.data) : response.data;
          asyncResult = createAsyncResult(formattedData, standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.ErrorCode.Success, 0, userContext);
        }

        callback(asyncResult);
      }
    }
  });
}
function createAsyncResult(value, errorCode, detailedErrorCode, userContext, errorMessage, errorName) {
  var initArgs = {};
  var errorArgs;
  initArgs[standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.Properties.Value] = value;
  initArgs[standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.Properties.Context] = userContext;

  if (standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.ErrorCode.Success !== errorCode) {
    errorArgs = {};
    var errorProperties = void 0;
    errorProperties = getErrorArgs(detailedErrorCode);
    errorArgs[standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.ErrorProperties.Name] = !errorName ? errorProperties.name : errorName;
    errorArgs[standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.ErrorProperties.Message] = !errorMessage ? errorProperties.message : errorMessage;
    errorArgs[standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.ErrorProperties.Code] = detailedErrorCode;
  }

  return new standardInvokeHostMethod_OSF.DDA.AsyncResult(initArgs, errorArgs);
}
var standardInvokeHostMethod_invokeHostMethod = function invokeHostMethod(dispid, data, responseCallback) {
  if (isOutlookJs()) {
    standardInvokeHostMethod_invokeHostMethodOutlookJs(dispid, data, responseCallback);
  } else {
    standardInvokeHostMethod_invokeHostMethodInternal(dispid, data, responseCallback);
  }
};

var standardInvokeHostMethod_invokeHostMethodInternal = function invokeHostMethodInternal(dispid, data, responseCallback) {
  if (standardInvokeHostMethod_OSF.AppName.OutlookWebApp !== getAppName() && isOwaOnly(dispid)) {
    responseCallback(InvokeResultCode.errorHandlingRequest, null);
    return;
  }

  var start = performance && performance.now();

  var invokeResponseCallback = function invokeResponseCallback(resultCode, resultData) {
    standardInvokeHostMethod_logTelemetry(resultCode, resultData, dispid, start);

    if (responseCallback) {
      responseCallback(resultCode, resultData);
    }
  };

  if (standardInvokeHostMethod_OSF.AppName.OutlookWebApp === getAppName()) {
    var args = {
      ApiParams: data,
      MethodData: {
        ControlId: standardInvokeHostMethod_OSF._OfficeAppFactory.getId(),
        DispatchId: dispid
      }
    };

    if (dispid === 1) {
      standardInvokeHostMethod_OSF._OfficeAppFactory.getClientEndPoint().invoke("GetInitialData", invokeResponseCallback, args);
    } else {
      standardInvokeHostMethod_OSF._OfficeAppFactory.getClientEndPoint().invoke("ExecuteMethod", invokeResponseCallback, args);
    }
  } else {
    callOutlookNativeDispatcher(dispid, data, invokeResponseCallback);
  }
};

var standardInvokeHostMethod_invokeHostMethodOutlookJs = function invokeHostMethodOutlookJs(dispid, data, responseCallback) {
  if (standardInvokeHostMethod_OSF.AppName.OutlookWebApp !== getAppName() && isOwaOnly(dispid)) {
    responseCallback(InvokeResultCode.errorHandlingRequest, null);
    return;
  }

  var dataTransform = standardInvokeHostMethod_createDataTransform(dispid, data);
  var start = performance && performance.now();

  standardInvokeHostMethod_OSF._OfficeAppFactory.getAsyncMethodExecutor().executeAsync(dispid, dataTransform, function (resultCode, response) {
    standardInvokeHostMethod_logTelemetry(resultCode, response, dispid, start);

    if (responseCallback) {
      var deserializedData = response;

      if (standardInvokeHostMethod_OSF.AppName.OutlookWebApp !== getAppName()) {
        deserializedData = callOutlookNativeDispatcher_deserializeResponseData(response);
      }

      responseCallback(resultCode, deserializedData);
    }
  });
};

var standardInvokeHostMethod_logTelemetry = function logTelemetry(resultCode, response, dispid, start) {
  if (standardInvokeHostMethod_OSF.AppTelemetry) {
    var detailedErrorCode = getErrorForTelemetry_getErrorForTelemetry(resultCode, response);
    var end = performance && performance.now();
    standardInvokeHostMethod_OSF.AppTelemetry.onMethodDone(dispid, null, Math.round(end - start), detailedErrorCode);
  }
};

var standardInvokeHostMethod_createDataTransform = function createDataTransform(dispid, data) {
  return {
    toSafeArrayHost: function toSafeArrayHost() {
      return callOutlookNativeDispatcher_convertToOutlookNativeParameters(dispid, data);
    },
    fromSafeArrayHost: function fromSafeArrayHost(payload) {
      return payload;
    },
    toWebHost: function toWebHost() {
      return data;
    },
    fromWebHost: function fromWebHost(payload) {
      return payload;
    }
  };
};
// CONCATENATED MODULE: ./src/utils/getPermissionLevel.ts


var getPermissionLevel_getPermissionLevel = function getPermissionLevel() {
  var permissionLevel = getInitialDataProp("permissionLevel");

  if (isNullOrUndefined(permissionLevel)) {
    return -1;
  }

  return permissionLevel;
};
// CONCATENATED MODULE: ./src/utils/createError.ts
function createError(message, errorInfo) {
  var err = new Error(message);
  err.message = message || "";

  if (errorInfo) {
    for (var v in errorInfo) {
      err[v] = errorInfo[v];
    }
  }

  return err;
}
function createBetaError(featureName) {
  var displayMessage = "The feature {0}, is only enabled on the beta api endpoint".replace("{0}", featureName);
  var err = createError(displayMessage, {
    name: "Sys.FeatureNotEnabled"
  });
  return err;
}
function createParameterCountError(message) {
  var displayMessage = "Sys.ParameterCountException: " + (message ? message : "Parameter count mismatch.");
  var err = createError(displayMessage, {
    name: "Sys.ParameterCountException"
  });
  return err;
}
function createArgumentError(paramName, message) {
  var displayMessage = "Sys.ArgumentException: " + (message ? message : "Value does not fall within the expected range.");

  if (paramName) {
    displayMessage += "\n" + "Parameter name: {0}".replace("{0}", paramName);
  }

  var err = createError(displayMessage, {
    name: "Sys.ArgumentException",
    paramName: paramName
  });
  return err;
}
function createNullItemError(namespace) {
  var displayMessage = "Invalid operation ({0}) when Office.context.mailbox.item is null.".replace("{0}", namespace);
  var err = createError(displayMessage);
  return err;
}
function createNullArgumentError(paramName, message) {
  var displayMessage = "Sys.ArgumentNullException: " + (message ? message : "Value cannot be null.");

  if (paramName) {
    displayMessage += "\n" + "Parameter name: {0}".replace("{0}", paramName);
  }

  var err = createError(displayMessage, {
    name: "Sys.ArgumentNullException",
    paramName: paramName
  });
  return err;
}
function createArgumentOutOfRange(paramName, actualValue, message) {
  var displayMessage = "Sys.ArgumentOutOfRangeException: " + (message ? message : "Specified argument was out of the range of valid values.");

  if (paramName) {
    displayMessage += "\n" + "Parameter name: {0}".replace("{0}", paramName);
  }

  if (typeof actualValue !== "undefined" && actualValue !== null) {
    displayMessage += "\n" + "Actual value was {0}.".replace("{0}", actualValue);
  }

  var err = createError(displayMessage, {
    name: "Sys.ArgumentOutOfRangeException",
    paramName: paramName,
    actualValue: actualValue
  });
  return err;
}
function createArgumentTypeError(paramName, actualType, expectedType, message) {
  var displayMessage = "Sys.ArgumentTypeException: ";

  if (message) {
    displayMessage += message;
  } else if (actualType && expectedType) {
    displayMessage += "Object of type '{0}' cannot be converted to type '{1}'.".replace("{0}", actualType.getName ? actualType.getName() : actualType).replace("{1}", expectedType.getName ? expectedType.getName() : expectedType);
  } else {
    displayMessage += "Object cannot be converted to the required type.";
  }

  if (paramName) {
    displayMessage += "\n" + "Parameter name: {0}".replace("{0}", paramName);
  }

  var err = createError(displayMessage, {
    name: "Sys.ArgumentTypeException",
    paramName: paramName,
    actualType: actualType,
    expectedType: expectedType
  });
  return err;
}
// CONCATENATED MODULE: ./src/utils/checkPermissionsAndThrow.ts



function checkPermissionsAndThrow(permissions, namespace) {
  if (getPermissionLevel_getPermissionLevel() == -1) {
    throw createNullItemError(namespace);
  }

  if (getPermissionLevel_getPermissionLevel() < permissions) {
    throw createError(getString("l_ElevatedPermissionNeededForMethod_Text").replace("{0}", namespace));
  }
}
// CONCATENATED MODULE: ./src/utils/parseCommonArgs.ts


function parseCommonArgs(args, isCallbackRequired, tryLegacy) {
  var result = {};

  if (tryLegacy) {
    result = tryParseLegacy(args);

    if (result.callback) {
      return result;
    }
  }

  if (args.length === 1) {
    if (typeof args[0] === "function") {
      result.callback = args[0];
    } else if (typeof args[0] === "object") {
      result.options = args[0];
    } else {
      throw createArgumentTypeError();
    }
  } else if (args.length === 2) {
    if (typeof args[0] !== "object") {
      throw createArgumentError("options");
    }

    if (typeof args[1] !== "function") {
      throw createArgumentError("callback");
    }

    result.callback = args[1];
    result.options = args[0];
  } else if (args.length !== 0) {
    throw createParameterCountError(getString("l_ParametersNotAsExpected_Text"));
  }

  if (isCallbackRequired && !result.callback) {
    throw createNullArgumentError("callback");
  }

  if (result.options && result.options.asyncContext) {
    result.asyncContext = result.options.asyncContext;
  }

  return result;
}

function tryParseLegacy(args) {
  var result = {};

  if (args.length === 1 || args.length === 2) {
    if (typeof args[0] !== "function") {
      return result;
    }

    result.callback = args[0];

    if (args.length === 2) {
      result.asyncContext = args[1];
    }

    return result;
  }

  return result;
}
// CONCATENATED MODULE: ./src/validation/recipientConstants.ts
var RecipientFields;

(function (RecipientFields) {
  RecipientFields[RecipientFields["to"] = 0] = "to";
  RecipientFields[RecipientFields["cc"] = 1] = "cc";
  RecipientFields[RecipientFields["bcc"] = 2] = "bcc";
  RecipientFields[RecipientFields["requiredAttendees"] = 0] = "requiredAttendees";
  RecipientFields[RecipientFields["optionalAttendees"] = 1] = "optionalAttendees";
})(RecipientFields || (RecipientFields = {}));

var displayNameLengthLimit = 255;
var recipientsLimit = 100;
var recipientConstants_totalRecipientsLimit = 500;
var maxSmtpLength = 571;
// CONCATENATED MODULE: ./src/validation/displayConstants.ts
var maxLocationLength = 255;
var maxBodyLength = 32 * 1024;
var maxSubjectLength = 255;
var maxRecipients = 100;
var MaxAttachmentNameLength = 255;
var MaxUrlLength = 2048;
var MaxItemIdLength = 200;
var MaxRemoveIdLength = 200;
// CONCATENATED MODULE: ./src/utils/throwOnOutOfRange.ts

function throwOnOutOfRange(value, minValue, maxValue, argumentName) {
  if (value < minValue || value > maxValue) {
    throw createArgumentOutOfRange(String(argumentName));
  }
}
// CONCATENATED MODULE: ./src/utils/OutlookEnums.ts
var MailboxEnums = {};
MailboxEnums.EntityType = {
  MeetingSuggestion: "meetingSuggestion",
  TaskSuggestion: "taskSuggestion",
  Address: "address",
  EmailAddress: "emailAddress",
  Url: "url",
  PhoneNumber: "phoneNumber",
  Contact: "contact",
  FlightReservations: "flightReservations",
  ParcelDeliveries: "parcelDeliveries"
};
MailboxEnums.ItemType = {
  Message: "message",
  Appointment: "appointment"
};
MailboxEnums.ResponseType = {
  None: "none",
  Organizer: "organizer",
  Tentative: "tentative",
  Accepted: "accepted",
  Declined: "declined"
};
MailboxEnums.RecipientType = {
  Other: "other",
  DistributionList: "distributionList",
  User: "user",
  ExternalUser: "externalUser"
};
MailboxEnums.AttachmentType = {
  File: "file",
  Item: "item",
  Cloud: "cloud"
};
MailboxEnums.AttachmentStatus = {
  Added: "added",
  Removed: "removed"
};
MailboxEnums.AttachmentContentFormat = {
  Base64: "base64",
  Url: "url",
  Eml: "eml",
  ICalendar: "iCalendar"
};
MailboxEnums.BodyType = {
  Text: "text",
  Html: "html"
};
MailboxEnums.ItemNotificationMessageType = {
  ProgressIndicator: "progressIndicator",
  InformationalMessage: "informationalMessage",
  ErrorMessage: "errorMessage",
  InsightMessage: "insightMessage"
};
MailboxEnums.Folder = {
  Inbox: "inbox",
  Junk: "junk",
  DeletedItems: "deletedItems"
};
MailboxEnums.ComposeType = {
  Forward: "forward",
  NewMail: "newMail",
  Reply: "reply"
};
var CoercionType = {
  Text: "text",
  Html: "html"
};
MailboxEnums.UserProfileType = {
  Office365: "office365",
  OutlookCom: "outlookCom",
  Enterprise: "enterprise"
};
MailboxEnums.RestVersion = {
  v1_0: "v1.0",
  v2_0: "v2.0",
  Beta: "beta"
};
MailboxEnums.ModuleType = {
  Addins: "addins"
};
MailboxEnums.ActionType = {
  ShowTaskPane: "showTaskPane"
};
MailboxEnums.Days = {
  Mon: "mon",
  Tue: "tue",
  Wed: "wed",
  Thu: "thu",
  Fri: "fri",
  Sat: "sat",
  Sun: "sun",
  Weekday: "weekday",
  WeekendDay: "weekendDay",
  Day: "day"
};
MailboxEnums.WeekNumber = {
  First: "first",
  Second: "second",
  Third: "third",
  Fourth: "fourth",
  Last: "last"
};
MailboxEnums.RecurrenceType = {
  Daily: "daily",
  Weekday: "weekday",
  Weekly: "weekly",
  Monthly: "monthly",
  Yearly: "yearly"
};
MailboxEnums.Month = {
  Jan: "jan",
  Feb: "feb",
  Mar: "mar",
  Apr: "apr",
  May: "may",
  Jun: "jun",
  Jul: "jul",
  Aug: "aug",
  Sep: "sep",
  Oct: "oct",
  Nov: "nov",
  Dec: "dec"
};
MailboxEnums.DelegatePermissions = {
  Read: 0x00000001,
  Write: 0x00000002,
  DeleteOwn: 0x00000004,
  DeleteAll: 0x00000008,
  EditOwn: 0x00000010,
  EditAll: 0x00000020
};
MailboxEnums.TimeZone = {
  AfghanistanStandardTime: "Afghanistan Standard Time",
  AlaskanStandardTime: "Alaskan Standard Time",
  AleutianStandardTime: "Aleutian Standard Time",
  AltaiStandardTime: "Altai Standard Time",
  ArabStandardTime: "Arab Standard Time",
  ArabianStandardTime: "Arabian Standard Time",
  ArabicStandardTime: "Arabic Standard Time",
  ArgentinaStandardTime: "Argentina Standard Time",
  AstrakhanStandardTime: "Astrakhan Standard Time",
  AtlanticStandardTime: "Atlantic Standard Time",
  AUSCentralStandardTime: "AUS Central Standard Time",
  AusCentralWStandardTime: "Aus Central W. Standard Time",
  AUSEasternStandardTime: "AUS Eastern Standard Time",
  AzerbaijanStandardTime: "Azerbaijan Standard Time",
  AzoresStandardTime: "Azores Standard Time",
  BahiaStandardTime: "Bahia Standard Time",
  BangladeshStandardTime: "Bangladesh Standard Time",
  BelarusStandardTime: "Belarus Standard Time",
  BougainvilleStandardTime: "Bougainville Standard Time",
  CanadaCentralStandardTime: "Canada Central Standard Time",
  CapeVerdeStandardTime: "Cape Verde Standard Time",
  CaucasusStandardTime: "Caucasus Standard Time",
  CenAustraliaStandardTime: "Cen. Australia Standard Time",
  CentralAmericaStandardTime: "Central America Standard Time",
  CentralAsiaStandardTime: "Central Asia Standard Time",
  CentralBrazilianStandardTime: "Central Brazilian Standard Time",
  CentralEuropeStandardTime: "Central Europe Standard Time",
  CentralEuropeanStandardTime: "Central European Standard Time",
  CentralPacificStandardTime: "Central Pacific Standard Time",
  CentralStandardTime: "Central Standard Time",
  CentralStandardTime_Mexico: "Central Standard Time (Mexico)",
  ChathamIslandsStandardTime: "Chatham Islands Standard Time",
  ChinaStandardTime: "China Standard Time",
  CubaStandardTime: "Cuba Standard Time",
  DatelineStandardTime: "Dateline Standard Time",
  EAfricaStandardTime: "E. Africa Standard Time",
  EAustraliaStandardTime: "E. Australia Standard Time",
  EEuropeStandardTime: "E. Europe Standard Time",
  ESouthAmericaStandardTime: "E. South America Standard Time",
  EasterIslandStandardTime: "Easter Island Standard Time",
  EasternStandardTime: "Eastern Standard Time",
  EasternStandardTime_Mexico: "Eastern Standard Time (Mexico)",
  EgyptStandardTime: "Egypt Standard Time",
  EkaterinburgStandardTime: "Ekaterinburg Standard Time",
  FijiStandardTime: "Fiji Standard Time",
  FLEStandardTime: "FLE Standard Time",
  GeorgianStandardTime: "Georgian Standard Time",
  GMTStandardTime: "GMT Standard Time",
  GreenlandStandardTime: "Greenland Standard Time",
  GreenwichStandardTime: "Greenwich Standard Time",
  GTBStandardTime: "GTB Standard Time",
  HaitiStandardTime: "Haiti Standard Time",
  HawaiianStandardTime: "Hawaiian Standard Time",
  IndiaStandardTime: "India Standard Time",
  IranStandardTime: "Iran Standard Time",
  IsraelStandardTime: "Israel Standard Time",
  JordanStandardTime: "Jordan Standard Time",
  KaliningradStandardTime: "Kaliningrad Standard Time",
  KamchatkaStandardTime: "Kamchatka Standard Time",
  KoreaStandardTime: "Korea Standard Time",
  LibyaStandardTime: "Libya Standard Time",
  LineIslandsStandardTime: "Line Islands Standard Time",
  LordHoweStandardTime: "Lord Howe Standard Time",
  MagadanStandardTime: "Magadan Standard Time",
  MagallanesStandardTime: "Magallanes Standard Time",
  MarquesasStandardTime: "Marquesas Standard Time",
  MauritiusStandardTime: "Mauritius Standard Time",
  MidAtlanticStandardTime: "Mid-Atlantic Standard Time",
  MiddleEastStandardTime: "Middle East Standard Time",
  MontevideoStandardTime: "Montevideo Standard Time",
  MoroccoStandardTime: "Morocco Standard Time",
  MountainStandardTime: "Mountain Standard Time",
  MountainStandardTime_Mexico: "Mountain Standard Time (Mexico)",
  MyanmarStandardTime: "Myanmar Standard Time",
  NCentralAsiaStandardTime: "N. Central Asia Standard Time",
  NamibiaStandardTime: "Namibia Standard Time",
  NepalStandardTime: "Nepal Standard Time",
  NewZealandStandardTime: "New Zealand Standard Time",
  NewfoundlandStandardTime: "Newfoundland Standard Time",
  NorfolkStandardTime: "Norfolk Standard Time",
  NorthAsiaEastStandardTime: "North Asia East Standard Time",
  NorthAsiaStandardTime: "North Asia Standard Time",
  NorthKoreaStandardTime: "North Korea Standard Time",
  OmskStandardTime: "Omsk Standard Time",
  PacificSAStandardTime: "Pacific SA Standard Time",
  PacificStandardTime: "Pacific Standard Time",
  PacificStandardTime_Mexico: "Pacific Standard Time (Mexico)",
  PakistanStandardTime: "Pakistan Standard Time",
  ParaguayStandardTime: "Paraguay Standard Time",
  RomanceStandardTime: "Romance Standard Time",
  RussiaTimeZone10: "Russia Time Zone 10",
  RussiaTimeZone11: "Russia Time Zone 11",
  RussiaTimeZone3: "Russia Time Zone 3",
  RussianStandardTime: "Russian Standard Time",
  SAEasternStandardTime: "SA Eastern Standard Time",
  SAPacificStandardTime: "SA Pacific Standard Time",
  SAWesternStandardTime: "SA Western Standard Time",
  SaintPierreStandardTime: "Saint Pierre Standard Time",
  SakhalinStandardTime: "Sakhalin Standard Time",
  SamoaStandardTime: "Samoa Standard Time",
  SaratovStandardTime: "Saratov Standard Time",
  SEAsiaStandardTime: "SE Asia Standard Time",
  SingaporeStandardTime: "Singapore Standard Time",
  SouthAfricaStandardTime: "South Africa Standard Time",
  SriLankaStandardTime: "Sri Lanka Standard Time",
  SudanStandardTime: "Sudan Standard Time",
  SyriaStandardTime: "Syria Standard Time",
  TaipeiStandardTime: "Taipei Standard Time",
  TasmaniaStandardTime: "Tasmania Standard Time",
  TocantinsStandardTime: "Tocantins Standard Time",
  TokyoStandardTime: "Tokyo Standard Time",
  TomskStandardTime: "Tomsk Standard Time",
  TongaStandardTime: "Tonga Standard Time",
  TransbaikalStandardTime: "Transbaikal Standard Time",
  TurkeyStandardTime: "Turkey Standard Time",
  TurksAndCaicosStandardTime: "Turks And Caicos Standard Time",
  UlaanbaatarStandardTime: "Ulaanbaatar Standard Time",
  USEasternStandardTime: "US Eastern Standard Time",
  USMountainStandardTime: "US Mountain Standard Time",
  UTC: "UTC",
  UTCPLUS12: "UTC+12",
  UTCPLUS13: "UTC+13",
  UTCMINUS02: "UTC-02",
  UTCMINUS08: "UTC-08",
  UTCMINUS09: "UTC-09",
  UTCMINUS11: "UTC-11",
  VenezuelaStandardTime: "Venezuela Standard Time",
  VladivostokStandardTime: "Vladivostok Standard Time",
  WAustraliaStandardTime: "W. Australia Standard Time",
  WCentralAfricaStandardTime: "W. Central Africa Standard Time",
  WEuropeStandardTime: "W. Europe Standard Time",
  WMongoliaStandardTime: "W. Mongolia Standard Time",
  WestAsiaStandardTime: "West Asia Standard Time",
  WestBankStandardTime: "West Bank Standard Time",
  WestPacificStandardTime: "West Pacific Standard Time",
  YakutskStandardTime: "Yakutsk Standard Time"
};
MailboxEnums.LocationType = {
  Custom: "custom",
  Room: "room"
};
MailboxEnums.AppointmentSensitivityType = {
  Normal: "normal",
  Personal: "personal",
  Private: "private",
  Confidential: "confidential"
};
MailboxEnums.CategoryColor = {
  None: "None",
  Preset0: "Preset0",
  Preset1: "Preset1",
  Preset2: "Preset2",
  Preset3: "Preset3",
  Preset4: "Preset4",
  Preset5: "Preset5",
  Preset6: "Preset6",
  Preset7: "Preset7",
  Preset8: "Preset8",
  Preset9: "Preset9",
  Preset10: "Preset10",
  Preset11: "Preset11",
  Preset12: "Preset12",
  Preset13: "Preset13",
  Preset14: "Preset14",
  Preset15: "Preset15",
  Preset16: "Preset16",
  Preset17: "Preset17",
  Preset18: "Preset18",
  Preset19: "Preset19",
  Preset20: "Preset20",
  Preset21: "Preset21",
  Preset22: "Preset22",
  Preset23: "Preset23",
  Preset24: "Preset24"
};
// CONCATENATED MODULE: ./src/utils/throwOnInvalidRestVersion.ts


function throwOnInvalidRestVersion(restVersion) {
  if (restVersion === null || restVersion === undefined) {
    throw createNullArgumentError(restVersion);
  }

  if (restVersion !== MailboxEnums.RestVersion.v1_0 && restVersion !== MailboxEnums.RestVersion.v2_0 && restVersion !== MailboxEnums.RestVersion.Beta) {
    throw createArgumentError(restVersion);
  }
}
// CONCATENATED MODULE: ./src/utils/convertToRestId.ts


function convertToRestId(itemId, restVersion) {
  if (itemId === null || itemId === undefined) {
    throw createNullArgumentError(itemId);
  }

  throwOnInvalidRestVersion(restVersion);
  return itemId.replace(new RegExp("[/]", "g"), "-").replace(new RegExp("[+]", "g"), "_");
}
// CONCATENATED MODULE: ./src/utils/convertToEwsId.ts


function convertToEwsId(itemId, restVersion) {
  if (itemId === null || itemId === undefined) {
    throw createNullArgumentError(itemId);
  }

  throwOnInvalidRestVersion(restVersion);
  return itemId.replace(new RegExp("[-]", "g"), "/").replace(new RegExp("[_]", "g"), "+");
}
// CONCATENATED MODULE: ./src/validation/validateDisplayForms.ts









function validateRecipientEmails(emailset, name) {
  if (!Array.isArray(emailset)) {
    throw createArgumentTypeError("name");
  }

  throwOnOutOfRange(emailset.length, 0, maxRecipients, "{0}.length".replace("{0}", name));
}
function normalizeRecipientEmails(emailset, name) {
  var originalAttendees = emailset;
  var updatedAttendees = [];

  for (var i = 0; i < originalAttendees.length; i++) {
    if (typeof originalAttendees[i] === "object") {
      throwOnInvalidEmailAddressDetails(originalAttendees[i]);
      updatedAttendees[i] = originalAttendees[i].emailAddress;

      if (typeof updatedAttendees[i] !== "string") {
        throw createArgumentError("{0}[{1}]".replace(name, String(i)));
      }
    } else {
      if (!(typeof originalAttendees[i] === "string")) {
        throw createArgumentError("{0}[{1}]".replace(name, String(i)));
      }

      updatedAttendees[i] = originalAttendees[i];
    }
  }

  return updatedAttendees;
}
function throwOnInvalidEmailAddressDetails(originalAttendee) {
  if (!isNullOrUndefined(originalAttendee.displayName)) {
    if (typeof originalAttendee.displayName === "string" && originalAttendee.displayName.length > displayNameLengthLimit) {
      throw createArgumentOutOfRange("displayName");
    }
  }

  if (!isNullOrUndefined(originalAttendee.emailAddress)) {
    if (typeof originalAttendee.emailAddress === "string" && originalAttendee.emailAddress.length > maxSmtpLength) {
      throw createArgumentOutOfRange("emailAddress");
    }
  }

  if (!isNullOrUndefined(originalAttendee.appointmentResponse)) {
    if (typeof originalAttendee.appointmentResponse !== "string") {
      throw createArgumentOutOfRange("appointmentResponse");
    }
  }

  if (!isNullOrUndefined(originalAttendee.recipientType)) {
    if (typeof originalAttendee.recipientType !== "string") {
      throw createArgumentOutOfRange("recipientType");
    }
  }
}
function validateDisplayFormParameters(itemId) {
  if (typeof itemId === "string") {
    throwOnInvalidItemId(itemId);
  } else {
    throw createArgumentTypeError("itemId");
  }
}

function throwOnInvalidItemId(itemId) {
  if (isNullOrUndefined(itemId) || itemId === "") {
    throw createNullArgumentError("itemId");
  }
}

function getItemIdBasedOnHost(itemId) {
  if (getInitialDataProp("isRestIdSupported")) {
    return convertToRestId(itemId, MailboxEnums.RestVersion.v1_0);
  }

  return convertToEwsId(itemId, MailboxEnums.RestVersion.v1_0);
}
// CONCATENATED MODULE: ./src/methods/displayAppointmentForm.ts
var __spreadArrays = undefined && undefined.__spreadArrays || function () {
  for (var s = 0, i = 0, il = arguments.length; i < il; i++) {
    s += arguments[i].length;
  }

  for (var r = Array(s), k = 0, i = 0; i < il; i++) {
    for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++) {
      r[k] = a[j];
    }
  }

  return r;
};





function displayAppointmentForm(itemId) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  displayAppointmentFormHelper.apply(void 0, __spreadArrays([9, itemId], args));
}
function displayAppointmentFormAsync(itemId) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  displayAppointmentFormHelper.apply(void 0, __spreadArrays([180, itemId], args));
}

function displayAppointmentFormHelper(dispidToInvoke, itemId) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "mailbox.displayAppointmentForm");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    itemId: itemId
  };
  validateParameters(parameters);
  standardInvokeHostMethod(dispidToInvoke, commonParameters.asyncContext, commonParameters.callback, {
    itemId: getItemIdBasedOnHost(parameters.itemId)
  }, undefined);
}

function validateParameters(parameters) {
  validateDisplayFormParameters(parameters.itemId);
}
// CONCATENATED MODULE: ./src/methods/displayMessageForm.ts
var displayMessageForm_spreadArrays = undefined && undefined.__spreadArrays || function () {
  for (var s = 0, i = 0, il = arguments.length; i < il; i++) {
    s += arguments[i].length;
  }

  for (var r = Array(s), k = 0, i = 0; i < il; i++) {
    for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++) {
      r[k] = a[j];
    }
  }

  return r;
};





function displayMessageForm(itemId) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  displayMessageFormHelper.apply(void 0, displayMessageForm_spreadArrays([8, itemId], args));
}
function displayMessageFormAsync(itemId) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  displayMessageFormHelper.apply(void 0, displayMessageForm_spreadArrays([179, itemId], args));
}

function displayMessageFormHelper(dispidToInvoke, itemId) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "mailbox.displayMessageForm");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    itemId: itemId
  };
  displayMessageForm_validateParameters(parameters);
  standardInvokeHostMethod(dispidToInvoke, commonParameters.asyncContext, commonParameters.callback, {
    itemId: getItemIdBasedOnHost(parameters.itemId)
  }, undefined);
}

function displayMessageForm_validateParameters(parameters) {
  validateDisplayFormParameters(parameters.itemId);
}
// CONCATENATED MODULE: ./src/utils/validateOptionalStringParameter.ts


function validateOptionalStringParameter(value, minLength, maxlength, name) {
  if (typeof value === "string") {
    throwOnOutOfRange(value.length, minLength, maxlength, name);
  } else {
    throw createArgumentError(String(name));
  }
}
// CONCATENATED MODULE: ./src/utils/isDateObject.ts
var isDateObject = function isDateObject(objectIn) {
  return objectIn instanceof Date || Object.prototype.toString.call(objectIn) == "[object Date]";
};
// CONCATENATED MODULE: ./src/methods/displayNewAppointmentForm.ts
var displayNewAppointmentForm_spreadArrays = undefined && undefined.__spreadArrays || function () {
  for (var s = 0, i = 0, il = arguments.length; i < il; i++) {
    s += arguments[i].length;
  }

  for (var r = Array(s), k = 0, i = 0; i < il; i++) {
    for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++) {
      r[k] = a[j];
    }
  }

  return r;
};











function displayNewAppointmentForm(parameters) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  return displayNewAppointmentFormHelper.apply(void 0, displayNewAppointmentForm_spreadArrays([7, parameters], args));
}
function displayNewAppointmentFormAsync(parameters) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  return displayNewAppointmentFormHelper.apply(void 0, displayNewAppointmentForm_spreadArrays([177, parameters], args));
}

function displayNewAppointmentFormHelper(dispidToInvoke, parameters) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "mailbox.displayNewAppointmentForm");
  var commonParameters = parseCommonArgs(args, false, false);
  displayNewAppointmentForm_validateParameters(parameters);
  var updatedParameters = normalizeParameters(parameters);
  standardInvokeHostMethod(dispidToInvoke, commonParameters.asyncContext, commonParameters.callback, updatedParameters, undefined);
}

function displayNewAppointmentForm_validateParameters(parameters) {
  if (!isNullOrUndefined(parameters.requiredAttendees)) {
    validateRecipientEmails(parameters.requiredAttendees, "requiredAttendees");
  }

  if (!isNullOrUndefined(parameters.optionalAttendees)) {
    validateRecipientEmails(parameters.optionalAttendees, "optionalAttendees");
  }

  if (!isNullOrUndefined(parameters.location)) {
    validateOptionalStringParameter(parameters.location, 0, maxLocationLength, "location");
  }

  if (!isNullOrUndefined(parameters.body)) {
    validateOptionalStringParameter(parameters.body, 0, maxBodyLength, "body");
  }

  if (!isNullOrUndefined(parameters.subject)) {
    validateOptionalStringParameter(parameters.subject, 0, maxSubjectLength, "subject");
  }

  if (!isNullOrUndefined(parameters.start)) {
    if (!isDateObject(parameters.start)) {
      throw createArgumentError("start");
    }

    if (!isNullOrUndefined(parameters.end)) {
      if (!isDateObject(parameters.end)) {
        throw createArgumentError("end");
      }

      if (parameters.end && parameters.start && parameters.end < parameters.start) {
        throw createArgumentError("end", getString("l_InvalidEventDates_Text"));
      }
    }
  }
}

function normalizeParameters(parameters) {
  var normalizedRequiredAttendees = null;
  var normalizedOptionalAttendees = null;

  if (!isNullOrUndefined(parameters.requiredAttendees)) {
    normalizedRequiredAttendees = normalizeRecipientEmails(parameters.requiredAttendees, "requiredAttendees");
  }

  if (!isNullOrUndefined(parameters.optionalAttendees)) {
    normalizedOptionalAttendees = normalizeRecipientEmails(parameters.optionalAttendees, "optionalAttendees");
  }

  if (!isNullOrUndefined(parameters.start)) {
    var startDate = parameters.start;
    parameters.start = startDate.getTime();
  }

  if (!isNullOrUndefined(parameters.end)) {
    var endDate = parameters.end;
    parameters.end = endDate.getTime();
  }

  var updatedParameters = JSON.parse(JSON.stringify(parameters));

  if (normalizedRequiredAttendees || normalizedOptionalAttendees) {
    if (!isNullOrUndefined(parameters.requiredAttendees)) {
      updatedParameters.requiredAttendees = normalizedRequiredAttendees;
    }

    if (!isNullOrUndefined(parameters.optionalAttendees)) {
      updatedParameters.optionalAttendees = normalizedOptionalAttendees;
    }
  }

  return updatedParameters;
}
// CONCATENATED MODULE: ./src/methods/displayNewMessageForm.ts
var displayNewMessageForm_spreadArrays = undefined && undefined.__spreadArrays || function () {
  for (var s = 0, i = 0, il = arguments.length; i < il; i++) {
    s += arguments[i].length;
  }

  for (var r = Array(s), k = 0, i = 0; i < il; i++) {
    for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++) {
      r[k] = a[j];
    }
  }

  return r;
};











function displayNewMessageForm(parameters) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  return displayNewMessageFormHelper.apply(void 0, displayNewMessageForm_spreadArrays([44, parameters], args));
}
function displayNewMessageFormAsync(parameters) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  return displayNewMessageFormHelper.apply(void 0, displayNewMessageForm_spreadArrays([178, parameters], args));
}

function displayNewMessageFormHelper(dispidToInvoke, parameters) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "mailbox.displayNewMessageForm");
  var commonParameters = parseCommonArgs(args, false, false);
  displayNewMessageForm_validateParameters(parameters);
  var updatedParameters = normailzeParameters(parameters);
  standardInvokeHostMethod(dispidToInvoke, commonParameters.asyncContext, commonParameters.callback, updatedParameters === null || updatedParameters === undefined ? parameters : updatedParameters, undefined);
}

function displayNewMessageForm_validateParameters(parameters) {
  if (parameters !== null && parameters !== null) {
    if (!isNullOrUndefined(parameters.toRecipients)) {
      validateRecipientEmails(parameters.toRecipients, "toRecipients");
    }

    if (!isNullOrUndefined(parameters.ccRecipients)) {
      validateRecipientEmails(parameters.ccRecipients, "ccRecipients");
    }

    if (!isNullOrUndefined(parameters.bccRecipients)) {
      validateRecipientEmails(parameters.bccRecipients, "bccRecipients");
    }

    if (!isNullOrUndefined(parameters.htmlBody)) {
      validateOptionalStringParameter(parameters.htmlBody, 0, maxBodyLength, "htmlBody");
    }

    if (!isNullOrUndefined(parameters.subject)) {
      validateOptionalStringParameter(parameters.subject, 0, maxSubjectLength, "subject");
    }
  }
}

function normailzeParameters(parameters) {
  var updatedParameters = JSON.parse(JSON.stringify(parameters));

  if (!isNullOrUndefined(parameters)) {
    if (parameters.toRecipients) {
      updatedParameters.toRecipients = normalizeRecipientEmails(parameters.toRecipients, "toRecipients");
    }

    if (parameters.ccRecipients) {
      updatedParameters.ccRecipients = normalizeRecipientEmails(parameters.ccRecipients, "ccRecipients");
    }

    if (parameters.bccRecipients) {
      updatedParameters.bccRecipients = normalizeRecipientEmails(parameters.bccRecipients, "bccRecipients");
    }

    var attachments = getAttachments(parameters);

    if (parameters.attachments) {
      updatedParameters.attachments = createAttachmentsDataForHost(attachments);
    }
  }

  return updatedParameters;
}
function getAttachments(data) {
  var attachments = [];

  if (data.attachments) {
    attachments = data.attachments;
    throwOnInvalidAttachmentsArray(attachments);
  }

  return attachments;
}
function throwOnInvalidAttachmentsArray(attachments) {
  if (!isNullOrUndefined(attachments) && !Array.isArray(attachments)) {
    throw createArgumentError("attachments");
  }
}
function createAttachmentsDataForHost(attachments) {
  var attachmentsData = [];

  for (var i = 0; i < attachments.length; i++) {
    if (typeof attachments[i] === "object") {
      var attachment = attachments[i];
      throwOnInvalidAttachment(attachment);
      attachmentsData.push(createAttachmentData(attachment));
    } else {
      throw createArgumentError("attachments");
    }
  }

  return attachmentsData;
}
function throwOnInvalidAttachment(attachment) {
  if (typeof attachment !== "object") {
    throw createArgumentError("attachments");
  }

  if (!attachment.type || !attachment.name) {
    throw createArgumentError("attachments");
  }

  if (!attachment.url && !attachment.itemId) {
    throw createArgumentError("attachments");
  }
}
function createAttachmentData(attachment) {
  var attachmentData = null;

  if (attachment.type === MailboxEnums.AttachmentType.File) {
    var url = attachment.url;
    var name_1 = attachment.name;
    var isInline = !!attachment.isInline;
    throwOnInvalidAttachmentUrlOrName(url, name_1);
    attachmentData = [MailboxEnums.AttachmentType.File, name_1, url, isInline];
  } else if (attachment.type === MailboxEnums.AttachmentType.Item) {
    var itemId = getItemIdBasedOnHost(attachment.itemId);
    var name_2 = attachment.name;
    throwOnInvalidAttachmentItemIdOrName(itemId, name_2);
    attachmentData = [MailboxEnums.AttachmentType.Item, name_2, itemId];
  } else {
    throw createArgumentError("attachments");
  }

  return attachmentData;
}
function throwOnInvalidAttachmentUrlOrName(url, name) {
  if (!(typeof url === "string") && !(typeof name === "string")) {
    throw createArgumentError("attachments");
  }

  if (url.length > MaxUrlLength) {
    throw createArgumentOutOfRange("attachments", url.length, getString("l_AttachmentUrlTooLong_Text"));
  }

  throwOnInvalidAttachmentName(name);
}
function throwOnInvalidAttachmentName(name) {
  if (name.length > MaxAttachmentNameLength) {
    throw createArgumentOutOfRange("attachments", name.length, getString("l_AttachmentNameTooLong_Text"));
  }
}
function throwOnInvalidAttachmentItemIdOrName(itemId, name) {
  if (!(typeof itemId === "string") || !(typeof name === "string")) {
    throw createArgumentError("attachments");
  }

  if (itemId.length > MaxItemIdLength) {
    throw createArgumentOutOfRange("attachments", itemId.length, getString("l_AttachmentItemIdTooLong_Text"));
  }

  throwOnInvalidAttachmentName(name);
}
// CONCATENATED MODULE: ./src/utils/handleTokenResponse.ts





var handleTokenResponse_OSF = __webpack_require__(0);

function handleTokenResponse(response, context, resultCode) {
  var asyncResult = undefined;

  if (getAppName() === handleTokenResponse_OSF.AppName.Outlook && response.error !== undefined && response.errorCode !== undefined && !!response.error && response.errorCode === 9030) {
    asyncResult = createAsyncResult(undefined, handleTokenResponse_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, response.errorCode, context, response.errorMessage, response.errorName);
  } else if (!!resultCode && resultCode !== InvokeResultCode.noError) {
    asyncResult = createAsyncResult(undefined, handleTokenResponse_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 9017, context, getString("l_InternalProtocolError_Text").replace("{0}", resultCode));

    if (!!asyncResult) {
      asyncResult.diagnostics = {
        InvokeCodeResult: resultCode
      };
    }
  } else {
    if (!!response.wasSuccessful) {
      asyncResult = createAsyncResult(response.token, handleTokenResponse_OSF.DDA.AsyncResultEnum.ErrorCode.Success, 0, context);
    } else {
      asyncResult = createAsyncResult(undefined, handleTokenResponse_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, response.errorCode, context, response.errorMessage, response.errorName);
    }

    if (response.diagnostics) {
      asyncResult.diagnostics = response.diagnostics;
    }
  }

  return asyncResult;
}
// CONCATENATED MODULE: ./src/methods/getCallbackToken.ts








function getCallbackToken() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "mailbox.getCallbackTokenAsync");
  var commonParameters = parseCommonArgs(args, true, true);
  var isRest = false;

  if (commonParameters.options && !!commonParameters.options.isRest) {
    isRest = true;
  }

  if (getIsNoItemContextWebExt()) {
    if (!isRest || getPermissionLevel_getPermissionLevel() < 3) {
      throw createError(getString("l_TokenAccessDeniedWithoutItemContext_Text"));
    }
  }

  standardInvokeHostMethod(12, commonParameters.asyncContext, commonParameters.callback, {
    isRest: isRest
  }, undefined, handleTokenResponse);
}
// CONCATENATED MODULE: ./src/methods/getUserIdentityToken.ts




function getUserIdentityToken() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "mailbox.getUserIdentityToken");
  var commonParameters = parseCommonArgs(args, true, true);
  standardInvokeHostMethod(2, commonParameters.asyncContext, commonParameters.callback, undefined, undefined, handleTokenResponse);
}
// CONCATENATED MODULE: ./src/methods/makeEwsRequest.ts







var makeEwsRequest_OSF = __webpack_require__(0);

var maxEwsRequestSize = 1000000;
function makeEwsRequest(body) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(3, "mailbox.makeEwsRequest");
  var commonParameters = parseCommonArgs(args, true, true);

  if (body === null || body === undefined) {
    throw createNullArgumentError("data");
  }

  if (typeof body !== "string") {
    throw createArgumentTypeError("data", typeof body, "string");
  }

  if (body.length > maxEwsRequestSize) {
    throw createArgumentError("data", getString("l_EwsRequestOversized_Text"));
  }

  standardInvokeHostMethod(5, commonParameters.asyncContext, commonParameters.callback, {
    body: body
  }, undefined, handleCustomResponse);
}

function handleCustomResponse(data, context, responseCode) {
  if (!!responseCode && responseCode !== InvokeResultCode.noError) {
    return createAsyncResult(undefined, makeEwsRequest_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 9017, context, getString("l_InternalProtocolError_Text").replace("{0}", responseCode));
  } else if (data.wasProxySuccessful === false) {
    return createAsyncResult(undefined, makeEwsRequest_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 9020, context, data.errorMessage);
  } else {
    return createAsyncResult(data.body, makeEwsRequest_OSF.DDA.AsyncResultEnum.ErrorCode.Success, 0, context);
  }
}
// CONCATENATED MODULE: ./src/utils/objectDefine.ts
var objectDefine = function objectDefine(o, props) {
  var keys = Object.keys(props);
  var values = keys.map(function (prop) {
    return {
      value: props[prop],
      writable: false
    };
  });
  var properties = {};
  keys.forEach(function (key, index) {
    properties[key] = values[index];
  });
  return Object.defineProperties(o, properties);
};
// CONCATENATED MODULE: ./src/api/getDiagnostics.ts



var getDiagnostics_OSF = __webpack_require__(0);

var getDiagnostics_getHostName = function getHostName() {
  var appName = getAppName();

  switch (appName) {
    case getDiagnostics_OSF.AppName.Outlook:
      return "Outlook";

    case getDiagnostics_OSF.AppName.OutlookWebApp:
      return "OutlookWebApp";

    case getDiagnostics_OSF.AppName.OutlookIOS:
      return "OutlookIOS";

    case getDiagnostics_OSF.AppName.OutlookAndroid:
      return "OutlookAndroid";

    default:
      return undefined;
  }
};
function getDiagnosticsSurface() {
  return objectDefine({}, {
    hostName: getDiagnostics_getHostName(),
    hostVersion: getInitialDataProp("hostVersion"),
    OWAView: getInitialDataProp("owaView")
  });
}
// CONCATENATED MODULE: ./src/api/getUserProfile.ts


function getUserProfileSurface() {
  return objectDefine({}, {
    accountType: getInitialDataProp("userProfileType"),
    displayName: getInitialDataProp("userDisplayName"),
    emailAddress: getInitialDataProp("userEmailAddress"),
    timeZone: getInitialDataProp("userTimeZone")
  });
}
// CONCATENATED MODULE: ./src/validation/categoryConstants.ts

var CategoryColor = MailboxEnums.CategoryColor;
var categoriesCharacterLimit = 255;
var colorPresets = [CategoryColor.None, CategoryColor.Preset0, CategoryColor.Preset1, CategoryColor.Preset2, CategoryColor.Preset3, CategoryColor.Preset4, CategoryColor.Preset5, CategoryColor.Preset6, CategoryColor.Preset7, CategoryColor.Preset8, CategoryColor.Preset9, CategoryColor.Preset10, CategoryColor.Preset11, CategoryColor.Preset12, CategoryColor.Preset13, CategoryColor.Preset14, CategoryColor.Preset15, CategoryColor.Preset16, CategoryColor.Preset17, CategoryColor.Preset18, CategoryColor.Preset19, CategoryColor.Preset20, CategoryColor.Preset21, CategoryColor.Preset22, CategoryColor.Preset23, CategoryColor.Preset24];
// CONCATENATED MODULE: ./src/validation/validateCategoryDetailsArray.ts


function validateCategoryDetailsArray(categoryDetails) {
  if (!categoryDetails) {
    throw createArgumentError("categoryDetails");
  }

  if (!Array.isArray(categoryDetails)) {
    throw createArgumentTypeError("categoryDetails", typeof categoryDetails, typeof []);
  }

  if (categoryDetails.length === 0) {
    throw createArgumentError("categoryDetails");
  }

  categoryDetails.forEach(validateCategoryDetails);
}

function validateCategoryDetails(categoryDetails) {
  if (!categoryDetails) {
    throw createArgumentError("categoryDetails");
  }

  if (!categoryDetails.color || !categoryDetails.displayName) {
    throw createArgumentError("categoryDetails");
  }

  if (typeof categoryDetails.color !== "string") {
    throw createArgumentTypeError("categoryDetails.color", typeof categoryDetails.color, "string");
  }

  if (typeof categoryDetails.displayName !== "string") {
    throw createArgumentTypeError("categoryDetails.displayName", typeof categoryDetails.displayName, "string");
  }

  if (categoryDetails.displayName.length > categoriesCharacterLimit) {
    throw createArgumentOutOfRange("categoryDetails.displayName", categoryDetails.displayName.length);
  }

  if (colorPresets.indexOf(categoryDetails.color) === -1) {
    throw createArgumentError("categoryDetails.color");
  }
}
// CONCATENATED MODULE: ./src/methods/addMasterCategories.ts




function addMasterCategories(categoryDetails) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(3, "masterCategories.addAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    categoryDetails: categoryDetails
  };
  validateCategoryDetailsArray(categoryDetails);
  standardInvokeHostMethod(161, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/getMasterCategories.ts



function getMasterCategories() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(3, "masterCategories.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(160, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/validation/validateCategoryArray.ts


function validateCategoryArray(categories) {
  if (!categories) {
    throw createArgumentError("categories");
  }

  if (!Array.isArray(categories)) {
    throw createArgumentTypeError("categories", typeof categories, typeof Array);
  }

  if (categories.length === 0) {
    throw createArgumentError("categories");
  }

  categories.forEach(validateCategory);
}

function validateCategory(category) {
  if (!category) {
    throw createArgumentError("categories");
  }

  if (typeof category !== "string") {
    throw createArgumentTypeError("categories", typeof category, "string");
  }

  if (category.length > categoriesCharacterLimit) {
    throw createArgumentOutOfRange("categories", category.length);
  }
}
// CONCATENATED MODULE: ./src/methods/removeMasterCategories.ts




function removeMasterCategories(categories) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(3, "masterCategories.removeAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    categories: categories
  };
  validateCategoryArray(categories);
  standardInvokeHostMethod(162, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/api/getMasterCategoriesSurface.ts




function getMasterCategoriesSurface() {
  return objectDefine({}, {
    addAsync: addMasterCategories,
    getAsync: getMasterCategories,
    removeAsync: removeMasterCategories
  });
}
// CONCATENATED MODULE: ./src/methods/closeApp.ts

function closeApp() {
  standardInvokeHostMethod(42, undefined, undefined, undefined, undefined);
}
// CONCATENATED MODULE: ./src/utils/getHostItemType.ts

var getHostItemType_getHostItemType = function getHostItemType() {
  return getInitialDataProp("itemType");
};
// CONCATENATED MODULE: ./src/utils/HostItemType.ts
var HostItemType;

(function (HostItemType) {
  HostItemType[HostItemType["Message"] = 1] = "Message";
  HostItemType[HostItemType["Appointment"] = 2] = "Appointment";
  HostItemType[HostItemType["MeetingRequest"] = 3] = "MeetingRequest";
  HostItemType[HostItemType["MessageCompose"] = 4] = "MessageCompose";
  HostItemType[HostItemType["AppointmentCompose"] = 5] = "AppointmentCompose";
  HostItemType[HostItemType["ItemLess"] = 6] = "ItemLess";
})(HostItemType || (HostItemType = {}));
// CONCATENATED MODULE: ./src/methods/getInitializationContext.ts



function getInitializationContext() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.getInitializationContext");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(99, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/validation/customPropertiesConstants.ts
var DatePrefix = "Date(";
var DatePostfix = ")";
var MaxCustomPropertiesLength = 2500;
var CustomPropertyType;

(function (CustomPropertyType) {
  CustomPropertyType[CustomPropertyType["NonTransmittable"] = 0] = "NonTransmittable";
})(CustomPropertyType || (CustomPropertyType = {}));
// CONCATENATED MODULE: ./src/methods/saveCustomProperties.ts





function saveCustomProperties(customProperties) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.saveCustomProperties");
  var commonParameters = parseCommonArgs(args, false, true);
  saveCustomProperties_validateParameters(customProperties);
  standardInvokeHostMethod(4, commonParameters.asyncContext, commonParameters.callback, {
    customProperties: customProperties
  }, undefined);
}

function saveCustomProperties_validateParameters(customProperties) {
  if (JSON.stringify(customProperties).length > MaxCustomPropertiesLength) {
    throw createArgumentOutOfRange("customProperties");
  }
}
// CONCATENATED MODULE: ./src/api/CustomProperties.ts
var CustomProperties_spreadArrays = undefined && undefined.__spreadArrays || function () {
  for (var s = 0, i = 0, il = arguments.length; i < il; i++) {
    s += arguments[i].length;
  }

  for (var r = Array(s), k = 0, i = 0; i < il; i++) {
    for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++) {
      r[k] = a[j];
    }
  }

  return r;
};







var CustomProperties_CustomProperties = function () {
  function CustomProperties(deserializedData) {
    if (isNullOrUndefined(deserializedData)) {
      createNullArgumentError("data");
    }

    if (Array.isArray(deserializedData)) {
      var customropertiesArray = deserializedData;

      if (customropertiesArray.length > CustomPropertyType.NonTransmittable) {
        deserializedData = customropertiesArray[CustomPropertyType.NonTransmittable];
      } else {
        throw createArgumentError("data");
      }
    } else {
      this.rawData = deserializedData;
    }
  }

  CustomProperties.prototype.get = function (key) {
    var value = this.rawData[key];

    if (typeof value === "string") {
      var valueString = value;

      if (valueString.length > DatePrefix.length + DatePostfix.length && valueString.startsWith(DatePrefix) && valueString.endsWith(DatePostfix)) {
        var ticksString = valueString.substring(DatePrefix.length, valueString.length - 1);
        var ticks = parseInt(ticksString);

        if (!isNaN(ticks)) {
          var dateTimeValue = new Date(ticks);

          if (!isNullOrUndefined(dateTimeValue)) {
            value = dateTimeValue;
          }
        }
      }
    }

    return value;
  };

  CustomProperties.prototype.set = function (key, value) {
    if (isDateObject(value)) {
      value = DatePrefix + value.getTime() + DatePostfix;
    }

    this.rawData[key] = value;
  };

  CustomProperties.prototype.remove = function (key) {
    delete this.rawData[key];
  };

  CustomProperties.prototype.saveAsync = function () {
    var args = [];

    for (var _i = 0; _i < arguments.length; _i++) {
      args[_i] = arguments[_i];
    }

    saveCustomProperties.apply(void 0, CustomProperties_spreadArrays([this.rawData], args));
  };

  CustomProperties.prototype.getAll = function () {
    var _this = this;

    var dictionary = {};
    var keys = Object.keys(this.rawData);
    keys.forEach(function (key) {
      dictionary[key] = _this.get(key);
    });
    return dictionary;
  };

  return CustomProperties;
}();


// CONCATENATED MODULE: ./src/methods/loadCustomProperties.ts






var loadCustomProperties_OSF = __webpack_require__(0);

function loadCustomProperties() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  var commonParameters = parseCommonArgs(args, true, true);
  standardInvokeHostMethod(3, commonParameters.asyncContext, commonParameters.callback, undefined, undefined, loadCustomProperties_handleCustomResponse);
}

function loadCustomProperties_handleCustomResponse(data, context, responseCode) {
  if (typeof responseCode !== "undefined" && responseCode !== InvokeResultCode.noError) {
    return createAsyncResult(undefined, loadCustomProperties_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 9017, context, getString("l_InternalProtocolError_Text").replace("{0}", responseCode));
  } else if (data.wasSuccessful) {
    var props = JSON.parse(data.customProperties);
    var value = new CustomProperties_CustomProperties(props);
    return createAsyncResult(value, loadCustomProperties_OSF.DDA.AsyncResultEnum.ErrorCode.Success, 0, context);
  } else {
    return createAsyncResult(undefined, loadCustomProperties_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 9020, context, data.errorMessage);
  }
}
// CONCATENATED MODULE: ./src/utils/bodyUtils.ts



var bodyUtils_OSF = __webpack_require__(0);

var HostCoercionType;

(function (HostCoercionType) {
  HostCoercionType[HostCoercionType["Text"] = 0] = "Text";
  HostCoercionType[HostCoercionType["Html"] = 3] = "Html";
})(HostCoercionType || (HostCoercionType = {}));

function addCoercionTypeParameter(parameters, args) {
  if (!!args.options && typeof args.options.coercionType === "string") {
    parameters.coercionType = getCoercionTypeFromString(args.options.coercionType);
  } else {
    parameters.coercionType = HostCoercionType.Text;
  }
}
function getCoercionTypeFromString(coercionType) {
  if (coercionType === CoercionType.Html) {
    return HostCoercionType.Html;
  } else if (coercionType === CoercionType.Text) {
    return HostCoercionType.Text;
  } else {
    return undefined;
  }
}
function invokeCallbackWithCoercionTypeError(args) {
  args.callback && args.callback(createAsyncResult(undefined, bodyUtils_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 1000, args.asyncContext));
}
// CONCATENATED MODULE: ./src/methods/getBody.ts





function getBody(coercionType) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "body.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  var parameters = {
    coercionType: getCoercionTypeFromString(coercionType)
  };

  if (parameters.coercionType === undefined) {
    throw createArgumentError("coercionType");
  }

  standardInvokeHostMethod(37, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/getBodyType.ts



function getBodyType() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "body.getTypeAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(14, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/validation/validateBodyParameters.ts

var maxDataLengthForBodyApi = 1000000;
var maxAppendOnSendLength = 5000;
var maxDataLengthForSignatureBodyApi = 120000;
function validateAppendOnSendBodyParamters(parameters) {
  if (typeof parameters.appendTxt !== "string") {
    throw createArgumentTypeError("data", typeof parameters.appendTxt, "string");
  }

  if (parameters.appendTxt.length > maxAppendOnSendLength) {
    throw createArgumentOutOfRange("data", parameters.appendTxt.length);
  }
}
function validateBodyParameters(parameters) {
  if (typeof parameters.data !== "string") {
    throw createArgumentTypeError("data", typeof parameters.data, "string");
  }

  if (parameters.data.length > maxDataLengthForBodyApi) {
    throw createArgumentOutOfRange("data", parameters.data.length);
  }
}
function validateSignatureBodyParameters(parameters) {
  if (typeof parameters.data !== "string") {
    throw createArgumentTypeError("data", typeof parameters.data, "string");
  }

  if (parameters.data.length > maxDataLengthForSignatureBodyApi) {
    throw createArgumentOutOfRange("data", parameters.data.length);
  }
}
// CONCATENATED MODULE: ./src/methods/setBody.ts





function setBody(data) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "body.setAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    data: data
  };
  validateBodyParameters(parameters);
  addCoercionTypeParameter(parameters, commonParameters);

  if (parameters.coercionType === undefined) {
    invokeCallbackWithCoercionTypeError(commonParameters);
    return;
  }

  standardInvokeHostMethod(38, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/bodyPrepend.ts





function bodyPrepend(data) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "body.prependAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    data: data
  };
  validateBodyParameters(parameters);
  addCoercionTypeParameter(parameters, commonParameters);

  if (parameters.coercionType === undefined) {
    invokeCallbackWithCoercionTypeError(commonParameters);
    return;
  }

  standardInvokeHostMethod(23, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/appendOnSend.ts






function appendOnSend(data) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "body.appendOnSendAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    appendTxt: data
  };

  if (isNullOrUndefined(parameters.appendTxt)) {
    parameters.appendTxt = "";
  } else {
    validateAppendOnSendBodyParamters(parameters);
  }

  addCoercionTypeParameter(parameters, commonParameters);

  if (parameters.coercionType === undefined) {
    invokeCallbackWithCoercionTypeError(commonParameters);
    return;
  }

  standardInvokeHostMethod(100, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/setSelectedData.ts





function setSelectedData(dispid) {
  return function (data) {
    var args = [];

    for (var _i = 1; _i < arguments.length; _i++) {
      args[_i - 1] = arguments[_i];
    }

    checkPermissionsAndThrow(2, "body.setSelectedDataAsync");
    var commonParameters = parseCommonArgs(args, false, false);
    var parameters = {
      data: data
    };
    validateBodyParameters(parameters);
    addCoercionTypeParameter(parameters, commonParameters);

    if (parameters.coercionType === undefined) {
      invokeCallbackWithCoercionTypeError(commonParameters);
      return;
    }

    standardInvokeHostMethod(dispid, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
  };
}
// CONCATENATED MODULE: ./src/utils/RuntimeFlighting.ts

var beta = 2;
var production = 1;
var currentLevel;
currentLevel = production;
var getCurrentLevel = function getCurrentLevel() {
  return currentLevel;
};
var Features = {
  featureSampleProduction: production,
  featureSampleBeta: beta,
  calendarItems: production,
  signature: production,
  replyCallback: beta,
  sessionData: beta
};
function isFeatureEnabled(feature) {
  return feature <= getCurrentLevel();
}
function checkFeatureEnabledAndThrow(feature, featureName) {
  if (!isFeatureEnabled(feature)) {
    throw createBetaError(featureName);
  }
}
// CONCATENATED MODULE: ./src/methods/setSignature.ts







function setSignature(data) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "item.body.setSignatureAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    data: data
  };
  checkFeatureEnabledAndThrow(Features.signature, "setSignatureAsync");

  if (isNullOrUndefined(parameters.data)) {
    parameters.data = "";
  } else {
    validateSignatureBodyParameters(parameters);
  }

  addCoercionTypeParameter(parameters, commonParameters);

  if (parameters.coercionType === undefined) {
    invokeCallbackWithCoercionTypeError(commonParameters);
    return;
  }

  standardInvokeHostMethod(173, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/api/getBodySurface.ts








function getBodySurface(isCompose) {
  var body = objectDefine({}, {
    getAsync: getBody
  });

  if (isCompose) {
    objectDefine(body, {
      appendOnSendAsync: appendOnSend,
      getTypeAsync: getBodyType,
      prependAsync: bodyPrepend,
      setAsync: setBody,
      setSelectedDataAsync: setSelectedData(13),
      setSignatureAsync: setSignature
    });
  }

  return body;
}
// CONCATENATED MODULE: ./src/methods/getAllInternetHeaders.ts



function getAllInternetHeaders() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.getAllInternetHeadersAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(168, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/types/ItemNotificationMessageType.ts
var ItemNotificationMessageType;

(function (ItemNotificationMessageType) {
  ItemNotificationMessageType[ItemNotificationMessageType["informationalMessage"] = 0] = "informationalMessage";
  ItemNotificationMessageType[ItemNotificationMessageType["progressIndicator"] = 1] = "progressIndicator";
  ItemNotificationMessageType[ItemNotificationMessageType["errorMessage"] = 2] = "errorMessage";
  ItemNotificationMessageType[ItemNotificationMessageType["insightMessage"] = 3] = "insightMessage";
})(ItemNotificationMessageType || (ItemNotificationMessageType = {}));
// CONCATENATED MODULE: ./src/utils/validateString.ts


function validateStringParam(paramName, paramValue) {
  if (isNullOrUndefined(paramValue) || paramValue === "") {
    throw createNullArgumentError(paramName);
  }

  if (!(typeof paramValue === "string")) {
    throw createArgumentTypeError(paramName, typeof paramValue, "string");
  }
}
function validateStringParamWithEmptyAllowed(paramName, paramValue) {
  if (isNullOrUndefined(paramValue)) {
    throw createNullArgumentError(paramName);
  }

  if (!(typeof paramValue === "string")) {
    throw createArgumentTypeError(paramName, typeof paramValue, "string");
  }
}
// CONCATENATED MODULE: ./src/validation/notificationMessagesConstants.ts
var MaximumKeyLength = 32;
var MaximumIconLength = 32;
var MaximumMessageLength = 150;
var MaximumActionTextLength = 30;
var NotificationsKeyParameterName = "key";
var NotificationsTypeParameterName = "type";
var NotificationsIconParameterName = "icon";
var NotificationsMessageParameterName = "message";
var NotificationsPersistentParameterName = "persistent";
var NotificationsActionsDefinitionParameterName = "actions";
var NotificationsActionTypeParameterName = "actionType";
var NotificationsActionTextParameterName = "actionText";
var NotificationsActionCommandIdParameterName = "commandId";
var NotificationsActionShowTaskPaneActionId = "showTaskPane";
// CONCATENATED MODULE: ./src/validation/validateNotificationMessages.ts






function validateKey(key) {
  validateStringParam(NotificationsKeyParameterName, key);

  if (key.length > MaximumKeyLength) {
    throw createArgumentOutOfRange(NotificationsKeyParameterName, key.length);
  }
}
function validateData(data) {
  validateStringParam(NotificationsTypeParameterName, data.type);

  if (data.type === MailboxEnums.ItemNotificationMessageType.InformationalMessage) {
    validateStringParam(NotificationsIconParameterName, data.icon);

    if (data.icon.length > MaximumIconLength) {
      throw createArgumentOutOfRange(NotificationsIconParameterName, data.icon.length);
    }

    if (isNullOrUndefined(data.persistent)) {
      throw createNullArgumentError(NotificationsPersistentParameterName);
    }

    if (typeof data.persistent !== "boolean") {
      throw createArgumentTypeError(NotificationsPersistentParameterName, typeof data.persistent, "boolean");
    }

    if (!isNullOrUndefined(data.actions)) {
      throw createArgumentError(NotificationsActionsDefinitionParameterName, getString("l_ActionsDefinitionWrongNotificationMessageError_Text"));
    }
  } else if (data.type === MailboxEnums.ItemNotificationMessageType.InsightMessage) {
    validateInsightMessageParameters(data);
  } else {
    if (!isNullOrUndefined(data.icon)) {
      throw createArgumentError(NotificationsIconParameterName);
    }

    if (!isNullOrUndefined(data.persistent)) {
      throw createArgumentError(NotificationsPersistentParameterName);
    }

    if (!isNullOrUndefined(data.actions)) {
      throw createArgumentError(NotificationsActionsDefinitionParameterName, getString("l_ActionsDefinitionWrongNotificationMessageError_Text"));
    }
  }

  validateStringParam(NotificationsMessageParameterName, data.message);

  if (data.message.length > MaximumMessageLength) {
    throw createArgumentOutOfRange(NotificationsMessageParameterName, data.message.length);
  }
}

function validateInsightMessageParameters(data) {
  validateStringParam(NotificationsIconParameterName, data.icon);

  if (data.icon.length > MaximumIconLength) {
    throw createArgumentOutOfRange(NotificationsIconParameterName, data.icon.length);
  }

  if (!isNullOrUndefined(data.persistent)) {
    throw createArgumentError(NotificationsPersistentParameterName);
  }

  if (isNullOrUndefined(data.actions)) {
    throw createNullArgumentError(NotificationsActionsDefinitionParameterName);
  } else {
    validateActionsDefinitionBlob(data.actions);
  }
}

function validateActionsDefinitionBlob(actionsDefinitionBlob) {
  var actionsDefinition = extractActionsDefinition(actionsDefinitionBlob);

  if (isNullOrUndefined(actionsDefinition)) {
    return;
  }

  validateActionsDefinitionActionsType(actionsDefinition);
  validateActionsDefinitionActionsText(actionsDefinition);
}

function extractActionsDefinition(actionsDefinitionBlob) {
  var actionsDefinition = null;

  if (Array.isArray(actionsDefinitionBlob)) {
    if (actionsDefinitionBlob.length === 1) {
      actionsDefinition = actionsDefinitionBlob[0];
    } else if (actionsDefinitionBlob.length > 1) {
      throw createArgumentError(NotificationsActionsDefinitionParameterName, getString("l_ActionsDefinitionMultipleActionsError_Text"));
    }
  } else {
    throw createArgumentError(NotificationsActionsDefinitionParameterName);
  }

  return actionsDefinition;
}

function validateActionsDefinitionActionsType(actionsDefinition) {
  if (isNullOrUndefined(actionsDefinition.actionType)) {
    throw createNullArgumentError(NotificationsActionTypeParameterName);
  }

  if (NotificationsActionShowTaskPaneActionId !== actionsDefinition.actionType) {
    throw createArgumentError(NotificationsActionTypeParameterName, getString("l_InvalidActionType_Text"));
  } else {
    if (isNullOrUndefined(actionsDefinition.commandId) || typeof actionsDefinition.commandId !== "string" || actionsDefinition.commandId === "") {
      throw createArgumentError(NotificationsActionCommandIdParameterName, getString("l_InvalidCommandIdError_Text"));
    }
  }
}

function validateActionsDefinitionActionsText(actionsDefinition) {
  if (isNullOrUndefined(actionsDefinition.actionText) || actionsDefinition.actionText === "" || typeof actionsDefinition.actionText !== "string") {
    throw createNullArgumentError(NotificationsActionTextParameterName);
  }

  if (actionsDefinition.actionText.length > MaximumActionTextLength) {
    throw createArgumentOutOfRange(NotificationsActionTextParameterName, actionsDefinition.actionText.length);
  }
}
// CONCATENATED MODULE: ./src/methods/addNotificationMessage.ts







function addNotificationMessage(key, data) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "notificationMessages.addAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  validateKey(key);
  validateData(data);
  var type = ItemNotificationMessageType[data.type];

  if (isNullOrUndefined(type)) {
    throw createArgumentError("type");
  }

  var message = data.message;
  var icon = data.icon;
  var persistent = data.persistent;
  var actions = data.actions;
  var parameters = {
    key: key,
    message: message,
    type: type,
    icon: icon,
    persistent: persistent,
    actions: actions
  };
  standardInvokeHostMethod(33, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/getAllNotificationMessages.ts



function getAllNotificationMessages() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "notificationMessages.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(34, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/removeNotificationMessage.ts




function removeNotificationMessage(key) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "notificationMessages.removeAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  validateKey(key);
  var parameters = {
    key: key
  };
  standardInvokeHostMethod(36, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/replaceNotificationMessage.ts







function replaceNotificationMessage(key, data) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "notificationMessages.replaceAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  validateKey(key);
  validateData(data);
  var type = ItemNotificationMessageType[data.type];

  if (isNullOrUndefined(type)) {
    throw createArgumentError("type");
  }

  var message = data.message;
  var icon = data.icon;
  var persistent = data.persistent;
  var actions = data.actions;
  var parameters = {
    key: key,
    message: message,
    type: type,
    icon: icon,
    persistent: persistent,
    actions: actions
  };
  standardInvokeHostMethod(35, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/api/getNotificationMessagesSurface.ts





function getNotificationMessageSurface() {
  return objectDefine({}, {
    addAsync: addNotificationMessage,
    getAllAsync: getAllNotificationMessages,
    removeAsync: removeNotificationMessage,
    replaceAsync: replaceNotificationMessage
  });
}
// CONCATENATED MODULE: ./src/validation/validateDisplayReplyForm.ts





function validateStringParameters(formData) {
  if (!isNullOrUndefined(formData)) {
    throwOnOutOfRange(formData.length, 0, maxBodyLength, "htmlBody");
  }
}
function validateAndGetHtmlBody(data) {
  var htmlBody = "";

  if (data.htmlBody) {
    throwOnInvalidHtmlBody(data.htmlBody);
    htmlBody = data.htmlBody;
  }

  return htmlBody;
}
function throwOnInvalidHtmlBody(htmlBody) {
  if (!(typeof htmlBody === "string")) {
    throw createArgumentTypeError("htmlBody", typeof htmlBody, "string");
  }

  if (isNullOrUndefined(htmlBody)) {
    throw createNullArgumentError("htmlBody");
  }

  throwOnOutOfRange(htmlBody.length, 0, maxBodyLength, "htmlBody");
}
function validateAndGetAttachments(data) {
  var attachments = [];

  if (data.attachments) {
    attachments = data.attachments;
    throwOnInvalidAttachmentsArray(attachments);
  }

  return attachments;
}
// CONCATENATED MODULE: ./src/utils/getOptionsAndCallback.ts

function getOptionsAndCallback(data) {
  var args = [];

  if (!isNullOrUndefined(data.options)) {
    args[0] = data.options;
  }

  if (!isNullOrUndefined(data.callback)) {
    args[args.length] = data.callback;
  }

  return args;
}
// CONCATENATED MODULE: ./src/methods/displayReplyForm.ts
var displayReplyForm_spreadArrays = undefined && undefined.__spreadArrays || function () {
  for (var s = 0, i = 0, il = arguments.length; i < il; i++) {
    s += arguments[i].length;
  }

  for (var r = Array(s), k = 0, i = 0; i < il; i++) {
    for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++) {
      r[k] = a[j];
    }
  }

  return r;
};










function displayReplyForm(formData) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  return displayReplyFormHelper.apply(void 0, displayReplyForm_spreadArrays([false, false, formData], args));
}
function displayReplyAllForm(formData) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  return displayReplyFormHelper.apply(void 0, displayReplyForm_spreadArrays([true, false, formData], args));
}
function displayReplyFormAsync(formData) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  return displayReplyFormHelper.apply(void 0, displayReplyForm_spreadArrays([false, true, formData], args));
}
function displayReplyAllFormAsync(formData) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  return displayReplyFormHelper.apply(void 0, displayReplyForm_spreadArrays([true, true, formData], args));
}

function displayReplyFormHelper(isReplyAll, isAsync, formData) {
  var args = [];

  for (var _i = 3; _i < arguments.length; _i++) {
    args[_i - 3] = arguments[_i];
  }

  var dispidToInvoke;
  checkPermissionsAndThrow(1, "mailbox.displayReplyForm");
  var commonParameters = parseCommonArgs(getOptionsAndCallback(formData), false, false);

  if (isFeatureEnabled(Features.replyCallback)) {
    if (isNullOrUndefined(commonParameters) || commonParameters.options === undefined && commonParameters.callback === undefined) {
      commonParameters = parseCommonArgs(args, false, false);
    }
  }

  var parameters = {
    formData: formData
  };
  var updatedHtmlBody = null;
  var updateAttachments = null;

  if (typeof parameters.formData === "string") {
    if (isReplyAll) {
      if (isAsync) {
        dispidToInvoke = 184;
      } else {
        dispidToInvoke = 11;
      }
    } else {
      if (isAsync) {
        dispidToInvoke = 183;
      } else {
        dispidToInvoke = 10;
      }
    }

    validateStringParameters(parameters.formData);
    standardInvokeHostMethod(dispidToInvoke, commonParameters.asyncContext, commonParameters.callback, {
      htmlBody: parameters.formData
    }, undefined);
  } else if (typeof parameters.formData === "object") {
    updatedHtmlBody = validateAndGetHtmlBody(parameters.formData);
    updateAttachments = createAttachmentsDataForHost(validateAndGetAttachments(parameters.formData));

    if (isReplyAll) {
      if (isAsync) {
        dispidToInvoke = 182;
      } else {
        dispidToInvoke = 31;
      }
    } else {
      if (isAsync) {
        dispidToInvoke = 181;
      } else {
        dispidToInvoke = 30;
      }
    }

    standardInvokeHostMethod(dispidToInvoke, commonParameters.asyncContext, commonParameters.callback, {
      htmlBody: updatedHtmlBody,
      attachments: updateAttachments
    }, undefined);
  } else {
    throw createArgumentError();
  }
}
// CONCATENATED MODULE: ./src/methods/addCategories.ts




function addCategories(categories) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "categories.addAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    categories: categories
  };
  validateCategoryArray(categories);
  standardInvokeHostMethod(158, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/getCategories.ts



function getCategories() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "categories.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(157, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/removeCategories.ts




function removeCategories(categories) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "categories.removeAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    categories: categories
  };
  validateCategoryArray(categories);
  standardInvokeHostMethod(159, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/api/getCategoriesSurface.ts




function getCategoriesSurface() {
  return objectDefine({}, {
    addAsync: addCategories,
    getAsync: getCategories,
    removeAsync: removeCategories
  });
}
// CONCATENATED MODULE: ./src/methods/getAttachmentContent.ts




function getAttachmentContent(id) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.getAttachmentContentAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  var parameters = {
    id: id
  };
  getAttachmentContent_validateParameters(parameters);
  standardInvokeHostMethod(150, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function getAttachmentContent_validateParameters(parameters) {
  validateStringParam("attachmentId", parameters.id);
}
// CONCATENATED MODULE: ./src/methods/moveToFolder.ts





var Folder = MailboxEnums.Folder;
function moveToFolder(destinationFolder) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(3, "item.move");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    destinationFolder: destinationFolder
  };
  moveToFolder_validateParameters(destinationFolder);
  standardInvokeHostMethod(101, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function moveToFolder_validateParameters(destinationFolder) {
  if (destinationFolder !== Folder.Inbox && destinationFolder !== Folder.Junk && destinationFolder !== Folder.DeletedItems) {
    throw createArgumentError("destinationFolder");
  }
}
// CONCATENATED MODULE: ./src/utils/createEmailAddressDetails.ts

var ResponseType = MailboxEnums.ResponseType;
var RecipientType = MailboxEnums.RecipientType;
var responseTypeMap = [ResponseType.None, ResponseType.Organizer, ResponseType.Tentative, ResponseType.Accepted, ResponseType.Declined];
var recipientTypeMap = [RecipientType.Other, RecipientType.DistributionList, RecipientType.User, RecipientType.ExternalUser];
var createEmailAddressDetails = function createEmailAddressDetails(input) {
  var response = input.appointmentResponse;
  var type = input.recipientType;
  var emailAddressDetails = {
    emailAddress: input.address,
    displayName: input.name
  };

  if (typeof input.appointmentResponse === "number") {
    emailAddressDetails.appointmentResponse = response < responseTypeMap.length ? responseTypeMap[response] : ResponseType.None;
  }

  if (typeof input.recipientType === "number") {
    emailAddressDetails.recipientType = type < recipientTypeMap.length ? recipientTypeMap[type] : RecipientType.Other;
  }

  return emailAddressDetails;
};
function createEmailAddressDetailsForEntity(data) {
  return createEmailAddressDetails({
    name: data.Name || "",
    address: data.UserId || ""
  });
}
// CONCATENATED MODULE: ./src/methods/getDelayDelivery.ts



function getDelayDelivery() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "delayDeliveryTime.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(166, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/setDelayDelivery.ts







function setDelayDelivery(dateTime) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "delayDeliveryTime.setAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  validateParamerters(dateTime);
  standardInvokeHostMethod(167, commonParameters.asyncContext, commonParameters.callback, {
    time: dateTime.getTime()
  }, undefined);
}

function validateParamerters(dateTime) {
  if (isNullOrUndefined(dateTime)) {
    throw createNullArgumentError("dateTime", "You cannot conduct to a null dateTime");
  }

  if (!isDateObject(dateTime)) {
    throw createArgumentTypeError("dateTime", typeof dateTime, typeof Date);
  }

  if (isNaN(dateTime.getTime())) {
    throw createArgumentError("dateTime");
  }

  throwOnOutOfRange(dateTime.getTime(), -8640000000000000, 8640000000000000, "dateTime");
}
// CONCATENATED MODULE: ./src/api/getDelayDeliverySurface.ts



function getDelayDeliverySurface(isCompose) {
  var delayDelivery = objectDefine({}, {
    getAsync: getDelayDelivery
  });

  if (isCompose) {
    objectDefine(delayDelivery, {
      setAsync: setDelayDelivery
    });
  }

  return delayDelivery;
}
// CONCATENATED MODULE: ./src/utils/removeDuplicates.ts
function removeDuplicates(array, comparator) {
  for (var matchIndex1 = array.length - 1; matchIndex1 >= 0; matchIndex1--) {
    var removeMatch = false;

    for (var matchIndex2 = matchIndex1 - 1; matchIndex2 >= 0; matchIndex2--) {
      if (comparator(array[matchIndex1], array[matchIndex2])) {
        removeMatch = true;
        break;
      }
    }

    if (removeMatch) {
      array.splice(matchIndex1, 1);
    }
  }

  return array;
}
var stringComparator = function stringComparator(a, b) {
  return a === b;
};
var meetingComparator = function meetingComparator(a, b) {
  if (a === b) {
    return true;
  } else if (!a || !b) {
    return false;
  } else {
    return a.meetingString === b.meetingString;
  }
};
var taskComparator = function taskComparator(a, b) {
  if (a === b) {
    return true;
  } else if (!a || !b) {
    return false;
  } else {
    return a.taskString === b.taskString;
  }
};
var contactComparator = function contactComparator(a, b) {
  if (a === b) {
    return true;
  } else if (!a || !b) {
    return false;
  } else {
    return a.contactString === b.contactString;
  }
};
// CONCATENATED MODULE: ./src/utils/isLegacyEntityExtraction.ts

function isLegacyEntityExtraction() {
  return !!getInitialDataProp("entities") && getInitialDataProp("entities").IsLegacyExtraction !== undefined && getInitialDataProp("entities").IsLegacyExtraction;
}
// CONCATENATED MODULE: ./src/utils/resolveDate.ts


var totalBits = 18;
var typeBits = 3;
var preciseDateTypeBits = 3;
var yearBits = 7;
var monthBits = 4;
var dayBits = 5;
var modifierBits = 2;
var unitBits = 3;
var offsetBits = 6;
var tagBits = 4;
var preciseDateType = 0;
var relativeDateType = 1;
var oneDayInMilliseconds = 86400000;
var baseDate = new Date("0001-01-01T00:00:00Z");
function resolveDate(input, sentTime) {
  if (!sentTime) {
    return input;
  }

  var date = null;

  try {
    var sentDate = new Date(sentTime.getFullYear(), sentTime.getMonth(), sentTime.getDate(), 0, 0, 0, 0);
    var extractedDate = decode(input);

    if (!extractedDate) {
      return input;
    } else {
      var preciseDate = extractedDate;

      if (preciseDate.day && preciseDate.month && preciseDate.year !== undefined) {
        date = resolvePreciseDate(sentDate, extractedDate);
      } else {
        var relativeDate = extractedDate;

        if (relativeDate.modifier !== undefined && relativeDate.offset !== undefined && relativeDate.tag !== undefined && relativeDate.unit !== undefined) {
          date = resolveRelativeDate(sentDate, extractedDate);
        } else {
          date = sentDate;
        }
      }
    }

    if (isNaN(date.getTime())) {
      return sentTime;
    }

    date.setMilliseconds(date.getMilliseconds() + (isLegacyEntityExtraction() ? getTimeOfDayInMillisecondsUTC(input) : getTimeOfDayInMilliseconds(input)));
    return date;
  } catch (e) {
    return sentTime;
  }
}

function decode(input) {
  var dateValueMask = (1 << totalBits - typeBits) - 1;
  var time = 0;

  if (input == null) {
    return undefined;
  }

  if (isLegacyEntityExtraction()) {
    time = getTimeOfDayInMillisecondsUTC(input);
  } else {
    time = getTimeOfDayInMilliseconds(input);
  }

  var inDateAtMidnight = input.getTime() - time;
  var value = (inDateAtMidnight - baseDate.getTime()) / oneDayInMilliseconds;

  if (value < 0) {
    return undefined;
  } else if (value >= 1 << totalBits) {
    return undefined;
  } else {
    var type = value >> totalBits - typeBits;
    value = value & dateValueMask;

    switch (type) {
      case preciseDateType:
        return decodePreciseDate(value);

      case relativeDateType:
        return decodeRelativeDate(value);

      default:
        return undefined;
    }
  }
}

function decodePreciseDate(value) {
  var cSubTypeMask = (1 << preciseDateTypeBits) - 1;
  var cMonthMask = (1 << monthBits) - 1;
  var cDayMask = (1 << dayBits) - 1;
  var cYearMask = (1 << yearBits) - 1;
  var year = 0;
  var month = 0;
  var day = 0;
  var subType = value >> totalBits - typeBits - preciseDateTypeBits & cSubTypeMask;

  if ((subType & 4) == 4) {
    year = value >> totalBits - typeBits - preciseDateTypeBits - yearBits & cYearMask;

    if ((subType & 2) == 2) {
      if ((subType & 1) == 1) {
        return undefined;
      }

      month = value >> totalBits - typeBits - preciseDateTypeBits - yearBits - monthBits & cMonthMask;
    }
  } else {
    if ((subType & 2) == 2) {
      month = value >> totalBits - typeBits - preciseDateTypeBits - monthBits & cMonthMask;
    }

    if ((subType & 1) == 1) {
      day = value >> totalBits - typeBits - preciseDateTypeBits - monthBits - dayBits & cDayMask;
    }
  }

  return createPreciseDate(day, month, year);
}

function resolvePreciseDate(sentDate, precise) {
  var year = precise.year;
  var month = precise.month == 0 ? sentDate.getMonth() : precise.month - 1;
  var day = precise.day;

  if (day == 0) {
    return sentDate;
  }

  var candidate;

  if (isNullOrUndefined(year)) {
    candidate = new Date(sentDate.getFullYear(), month, day);

    if (candidate.getTime() < sentDate.getTime()) {
      candidate = new Date(sentDate.getFullYear() + 1, month, day);
    }
  } else {
    candidate = new Date(year < 50 ? 2000 + year : 1900 + year, month, day);
  }

  if (candidate.getMonth() != month) {
    return sentDate;
  }

  return candidate;
}

function resolveRelativeDate(sentDate, relative) {
  var date;

  switch (relative.unit) {
    case 0:
      date = new Date(sentDate.getFullYear(), sentDate.getMonth(), sentDate.getDate());
      date.setDate(date.getDate() + relative.offset);
      return date;

    case 5:
      return findBestDateForWeekDate(sentDate, relative.offset, relative.tag);

    case 2:
      {
        var days = 1;

        switch (relative.modifier) {
          case 1:
            break;

          case 2:
            days = 16;
            break;

          default:
            if (relative.offset == 0) {
              days = sentDate.getDate();
            }

            break;
        }

        date = new Date(sentDate.getFullYear(), sentDate.getMonth(), days);
        date.setMonth(date.getMonth() + relative.offset);

        if (date.getTime() < sentDate.getTime()) {
          date.setDate(date.getDate() + sentDate.getDate() - 1);
        }

        return date;
      }

    case 1:
      date = new Date(sentDate.getFullYear(), sentDate.getMonth(), sentDate.getDate());
      date.setDate(sentDate.getDate() + 7 * relative.offset);

      if (relative.modifier == 1 || relative.modifier == 0) {
        date.setDate(date.getDate() + 1 - date.getDay());

        if (date.getTime() < sentDate.getTime()) {
          return sentDate;
        }

        return date;
      } else if (relative.modifier == 2) {
        date.setDate(date.getDate() + 5 - date.getDay());
        return date;
      }

      break;

    case 4:
      return findBestDateForWeekOfMonthDate(sentDate, relative);

    case 3:
      if (relative.offset > 0) {
        return new Date(sentDate.getFullYear() + relative.offset, 0, 1);
      }

      break;

    default:
      break;
  }

  return sentDate;
}

function findBestDateForWeekDate(sentDate, offset, tag) {
  if (offset > -5 && offset < 5) {
    var dayOfWeek = (tag + 6) % 7 + 1;
    var days = 7 * offset + (dayOfWeek - sentDate.getDay());
    sentDate.setDate(sentDate.getDate() + days);
    return sentDate;
  } else {
    var days = (tag - sentDate.getDay()) % 7;

    if (days < 0) {
      days += 7;
    }

    sentDate.setDate(sentDate.getDate() + days);
    return sentDate;
  }
}

function findBestDateForWeekOfMonthDate(sentDate, relative) {
  var date;
  var firstDay;
  var newDate;
  date = sentDate;

  if (relative.tag <= 0 || relative.tag > 12 || relative.offset <= 0 || relative.offset > 5) {
    return sentDate;
  }

  var monthOffset = (12 + relative.tag - date.getMonth() - 1) % 12;
  firstDay = new Date(date.getFullYear(), date.getMonth() + monthOffset, 1);

  if (relative.modifier == 1) {
    if (relative.offset == 1 && firstDay.getDay() != 6 && firstDay.getDay() != 0) {
      return firstDay;
    } else {
      newDate = new Date(firstDay.getFullYear(), firstDay.getMonth(), firstDay.getDate());
      newDate.setDate(newDate.getDate() + (7 + (1 - firstDay.getDay())) % 7);

      if (firstDay.getDay() != 6 && firstDay.getDay() != 0 && firstDay.getDay() != 1) {
        newDate.setDate(newDate.getDate() - 7);
      }

      newDate.setDate(newDate.getDate() + 7 * (relative.offset - 1));

      if (newDate.getMonth() + 1 != relative.tag) {
        return sentDate;
      }

      return newDate;
    }
  } else {
    newDate = new Date(firstDay.getFullYear(), firstDay.getMonth(), daysInMonth(firstDay.getMonth(), firstDay.getFullYear()));
    var offset = 1 - newDate.getDay();

    if (offset > 0) {
      offset = offset - 7;
    }

    newDate.setDate(newDate.getDate() + offset);
    newDate.setDate(newDate.getDate() + 7 * (1 - relative.offset));

    if (newDate.getMonth() + 1 != relative.tag) {
      if (firstDay.getDay() != 6 && firstDay.getDay() != 0) {
        return firstDay;
      } else {
        return sentDate;
      }
    } else {
      return newDate;
    }
  }
}

function decodeRelativeDate(value) {
  var tagMask = (1 << tagBits) - 1;
  var offsetMask = (1 << offsetBits) - 1;
  var unitMask = (1 << unitBits) - 1;
  var modifierMask = (1 << modifierBits) - 1;
  var tag = value & tagMask;
  value >>= tagBits;
  var offset = fromComplement(value & offsetMask, offsetBits);
  value >>= offsetBits;
  var unit = value & unitMask;
  value >>= unitBits;
  var modifier = value & modifierMask;

  try {
    return createRelativeDate(modifier, offset, unit, tag);
  } catch (_a) {
    return undefined;
  }
}

function fromComplement(value, n) {
  var signed = 1 << n - 1;
  var mask = (1 << n) - 1;

  if ((value & signed) == signed) {
    return -((value ^ mask) + 1);
  } else {
    return value;
  }
}

function daysInMonth(month, year) {
  return 32 - new Date(year, month, 32).getDate();
}

function getTimeOfDayInMilliseconds(inputTime) {
  var timeOfDay = 0;
  timeOfDay += inputTime.getHours() * 3600;
  timeOfDay += inputTime.getMinutes() * 60;
  timeOfDay += inputTime.getSeconds();
  timeOfDay *= 1000;
  timeOfDay += inputTime.getMilliseconds();
  return timeOfDay;
}

function getTimeOfDayInMillisecondsUTC(inputTime) {
  var timeOfDay = 0;
  timeOfDay += inputTime.getUTCHours() * 3600;
  timeOfDay += inputTime.getUTCMinutes() * 60;
  timeOfDay += inputTime.getUTCSeconds();
  timeOfDay *= 1000;
  timeOfDay += inputTime.getUTCMilliseconds();
  return timeOfDay;
}

function createPreciseDate(day, month, year) {
  return {
    day: day,
    month: month,
    year: year % 100
  };
}

function createRelativeDate(modifier, offset, unit, tag) {
  return {
    modifier: modifier,
    offset: offset,
    unit: unit,
    tag: tag
  };
}
// CONCATENATED MODULE: ./src/utils/findOffset.ts



function findOffset(value) {
  var ranges = getInitialDataProp("timeZoneOffsets");

  for (var r = 0; r < ranges.length; r++) {
    var range = ranges[r];
    var start = parseInt(range.start);
    var end = parseInt(range.end);

    if (value.getTime() - start >= 0 && value.getTime() - end < 0) {
      return parseInt(range.offset);
    }
  }

  throw createArgumentError("input", getString("l_InvalidDate_Text"));
}
// CONCATENATED MODULE: ./src/utils/convertToUtcClientTime.ts





function convertToUtcClientTime(input) {
  var retValue = localClientTimeToDate(input);

  if (!isNullOrUndefined(getInitialDataProp("timeZoneOffsets"))) {
    var offset = findOffset(retValue);
    retValue.setUTCMinutes(retValue.getUTCMinutes() - offset);
    offset = !input["timezoneOffset"] ? retValue.getTimezoneOffset() * -1 : input["timezoneOffset"];
    retValue.setUTCMinutes(retValue.getUTCMinutes() + offset);
  }

  return retValue;
}
function localClientTimeToDate(input) {
  var retValue = new Date(input["year"], input["month"], input["date"], input["hours"], input["minutes"], input["seconds"], input["milliseconds"] === null ? 0 : input["milliseconds"]);

  if (isNaN(retValue.getTime())) {
    throw createArgumentError("input", getString("l_InvalidDate_Text"));
  }

  return retValue;
}
// CONCATENATED MODULE: ./src/utils/dateToDictionary.ts
function dateToDictionary(date) {
  return {
    month: date.getMonth(),
    date: date.getDate(),
    year: date.getFullYear(),
    hours: date.getHours(),
    minutes: date.getMinutes(),
    seconds: date.getSeconds(),
    milliseconds: date.getMilliseconds()
  };
}
// CONCATENATED MODULE: ./src/utils/createEntities.ts









var EntityKeys;

(function (EntityKeys) {
  EntityKeys["meetingSuggestion"] = "MeetingSuggestions";
  EntityKeys["taskSuggestion"] = "TaskSuggestions";
  EntityKeys["address"] = "Addresses";
  EntityKeys["emailAddress"] = "EmailAddresses";
  EntityKeys["url"] = "Urls";
  EntityKeys["phoneNumber"] = "PhoneNumbers";
  EntityKeys["contact"] = "Contacts";
  EntityKeys["flightReservations"] = "FlightReservations";
  EntityKeys["parcelDeliveries"] = "ParcelDeliveries";
})(EntityKeys || (EntityKeys = {}));

function createEntities(data) {
  if (isNullOrUndefined(data)) {
    return {
      addresses: [],
      emailAddresses: [],
      urls: [],
      taskSuggestions: [],
      meetingSuggestions: [],
      phoneNumbers: [],
      contacts: [],
      flightReservations: [],
      parcelDelivery: []
    };
  } else {
    return {
      addresses: createEntities_createAddresses(data[EntityKeys.address]),
      emailAddresses: createEntities_createEmailAddresses(data[EntityKeys.emailAddress]),
      urls: createUrls(data[EntityKeys.url]),
      taskSuggestions: createEntities_createTaskSuggestions(data[EntityKeys.taskSuggestion]),
      meetingSuggestions: createEntities_createMeetingSuggestions(data[EntityKeys.meetingSuggestion]),
      phoneNumbers: createPhoneNumbers(data[EntityKeys.phoneNumber]),
      contacts: createEntities_createContacts(data[EntityKeys.contact]),
      flightReservations: createEntities_createReadItemArray(data[EntityKeys.flightReservations]),
      parcelDelivery: createEntities_createReadItemArray(data[EntityKeys.parcelDeliveries])
    };
  }
}
function createFilteredEntities(data, name) {
  checkPermissionsAndThrow(1, "item.getFilteredEntitiesByName");
  var results = Object.keys(data).map(function (entities) {
    var results = data[entities][name];
    if (results) return {
      entityType: entities,
      name: name,
      entities: data[entities][name]
    };else return undefined;
  }).filter(function (results) {
    return results !== undefined;
  });

  if (results.length === 0) {
    return null;
  }

  var matchedRule = results[0];

  switch (matchedRule.entityType) {
    case EntityKeys.meetingSuggestion:
      return createEntities_createMeetingSuggestions(matchedRule.entities);

    case EntityKeys.address:
      return createEntities_createAddresses(matchedRule.entities);

    case EntityKeys.contact:
      return createEntities_createContacts(matchedRule.entities);

    case EntityKeys.emailAddress:
      return createEntities_createEmailAddresses(matchedRule.entities);

    case EntityKeys.phoneNumber:
      return createPhoneNumbers(matchedRule.entities);

    case EntityKeys.taskSuggestion:
      return createEntities_createTaskSuggestions(matchedRule.entities);

    case EntityKeys.url:
      return createUrls(matchedRule.entities);

    default:
      return createEntities_createReadItemArray(matchedRule.entities);
  }
}
var createEntities_createAddresses = function createAddresses(data) {
  var addresses = data || [];
  return removeDuplicates(addresses, stringComparator);
};
var createEntities_createEmailAddresses = function createEmailAddresses(data) {
  if (getPermissionLevel_getPermissionLevel() === 0) {
    return [];
  }

  return data || [];
};
var createUrls = function createUrls(data) {
  return data || [];
};
var createEntities_createTaskSuggestions = function createTaskSuggestions(data) {
  if (getPermissionLevel_getPermissionLevel() === 0) {
    return [];
  }

  var tasks = data || [];
  tasks = tasks.map(function (task) {
    return {
      assignees: (task.Assignees || []).map(createEmailAddressDetailsForEntity),
      taskString: task.TaskString
    };
  });
  return removeDuplicates(tasks, taskComparator);
};
var createEntities_createMeetingSuggestions = function createMeetingSuggestions(data) {
  if (getPermissionLevel_getPermissionLevel() === 0) {
    return [];
  }

  var meetings = data || [];
  meetings = meetings.map(function (meeting) {
    var start = meeting.StartTime !== "" ? getDate(meeting.StartTime) : undefined;
    var end = meeting.EndTime !== "" ? getDate(meeting.EndTime) : undefined;
    return {
      meetingString: meeting.MeetingString,
      attendees: (meeting.Attendees || []).map(createEmailAddressDetailsForEntity),
      location: meeting.Location,
      subject: meeting.Subject,
      start: meeting.StartTime !== undefined ? start : undefined,
      end: meeting.EndTime !== undefined ? end : undefined
    };
  });
  return removeDuplicates(meetings, meetingComparator);
};

function getDate(date) {
  var result = resolveDate(new Date(date), new Date(getInitialDataProp("dateTimeSent")));

  if (result.getTime() !== new Date(date).getTime()) {
    return convertToUtcClientTime(dateToDictionary(result));
  }

  return new Date(date);
}

var createPhoneNumbers = function createPhoneNumbers(data) {
  var phoneNumbers = data || [];
  return phoneNumbers.map(function (number) {
    return {
      phoneString: number.PhoneString,
      originalPhoneString: number.OriginalPhoneString,
      type: number.Type
    };
  });
};
var createEntities_createContacts = function createContacts(data) {
  if (getPermissionLevel_getPermissionLevel() === 0) {
    return [];
  }

  var contacts = data || [];
  contacts = contacts.map(function (contact) {
    return {
      personName: contact.PersonName,
      businessName: contact.BusinessName,
      phoneNumbers: createPhoneNumbers(contact.PhoneNumbers || []),
      emailAddresses: contact.EmailAddresses || [],
      urls: contact.Urls || [],
      addresses: contact.Addresses || [],
      contactString: contact.ContactString
    };
  });
  return removeDuplicates(contacts, contactComparator);
};
var createEntities_createReadItemArray = function createReadItemArray(data) {
  if (getPermissionLevel_getPermissionLevel() === 0) {
    return [];
  }

  return data || [];
};
// CONCATENATED MODULE: ./src/api/Entities.ts



var entityPermissions = {
  meetingSuggestion: 1,
  taskSuggestion: 1,
  address: 0,
  emailAddress: 1,
  url: 0,
  phoneNumber: 0,
  contact: 1,
  flightReservations: 1,
  parcelDeliveries: 1
};
var entityKeys = {
  meetingSuggestion: "meetingSuggestions",
  taskSuggestion: "taskSuggestions",
  address: "addresses",
  emailAddress: "emailAddresses",
  url: "urls",
  phoneNumber: "phoneNumbers",
  contact: "contacts",
  flightReservations: "flightReservations",
  parcelDeliveries: "parcelDeliveries"
};
var Entities_getEntities = function getEntities() {
  return createEntities(getInitialDataProp("entities"));
};
var Entities_getEntitiesByType = function getEntitiesByType(entityType) {
  var entities = createEntities(getInitialDataProp("entities"));
  checkPermissionsAndThrow(entityPermissions[entityType] !== undefined ? entityPermissions[entityType] : 1, entityType);
  var entityProperty = entityKeys[entityType];

  if (entityProperty === undefined) {
    return null;
  }

  return entities[entityProperty];
};
var Entities_getFilteredEntitiesByName = function getFilteredEntitiesByName(name) {
  return createFilteredEntities(getInitialDataProp("filteredEntities"), name);
};
var Entities_getRegExMatches = function getRegExMatches() {
  return getInitialDataProp("regExMatches");
};
var Entities_getRegExMatchesByName = function getRegExMatchesByName(name) {
  var regExMatches = getInitialDataProp("regExMatches") || {};
  return regExMatches[name];
};
var Entities_getSelectedEntities = function getSelectedEntities() {
  return createEntities(getInitialDataProp("selectedEntities"));
};
var Entities_getSelectedRegExMatches = function getSelectedRegExMatches() {
  return getInitialDataProp("selectedRegExMatches");
};
// CONCATENATED MODULE: ./src/utils/CustomJsonAttachmentsResponse.ts


function CustomJsonAttachmentsResponse(arrayOfAttachmentJsonData) {
  var customJsonResponse = [];

  if (getPermissionLevel_getPermissionLevel() === 0) {
    return [];
  }

  if (!!arrayOfAttachmentJsonData) {
    for (var i = 0; i < arrayOfAttachmentJsonData.length; i++) {
      if (!!arrayOfAttachmentJsonData[i]) {
        var newAttachment = convertAttachmentType(arrayOfAttachmentJsonData[i]);
        customJsonResponse.push(newAttachment);
      }
    }
  }

  return customJsonResponse;
}
function convertAttachmentType(attachmentDetails) {
  if (attachmentDetails.attachmentType !== null || attachmentDetails.attachmentType !== undefined) {
    switch (attachmentDetails.attachmentType) {
      case 0:
        {
          attachmentDetails.attachmentType = MailboxEnums.AttachmentType.File;
          break;
        }

      case 1:
        {
          attachmentDetails.attachmentType = MailboxEnums.AttachmentType.Item;
          break;
        }

      case 2:
        {
          attachmentDetails.attachmentType = MailboxEnums.AttachmentType.Cloud;
          break;
        }
    }
  }

  return attachmentDetails;
}
// CONCATENATED MODULE: ./src/methods/deepClone.ts
function deepClone(original) {
  return JSON.parse(JSON.stringify(original));
}
// CONCATENATED MODULE: ./src/validation/seriesTimeConstants.ts
var StartYearKey = "startYear";
var StartMonthKey = "startMonth";
var StartDayKey = "startDay";
var EndYearKey = "endYear";
var EndMonthKey = "endMonth";
var EndDayKey = "endDay";
var NoEndDateKey = "noEndDate";
var StartTimeMinKey = "startTimeMin";
var DurationMinKey = "durationMin";
// CONCATENATED MODULE: ./src/validation/recurrenceConstants.ts
var StartDateKey = "startDate";
var EndDateKey = "endDate";
var StartTimeKey = "startTime";
var EndTimeKey = "endTime";
var RecurrenceTypeKey = "recurrenceType";
var SeriesTimeKey = "seriesTime";
var SeriesTimeJsonKey = "seriesTimeJson";
var RecurrenceTimeZoneKey = "recurrenceTimeZone";
var RecurrenceTimeZoneName = "name";
var RecurrencePropertiesKey = "recurrenceProperties";
var IntervalKey = "interval";
var DaysKey = "days";
var DayOfMonthKey = "dayOfMonth";
var DayOfWeekKey = "dayOfWeek";
var WeekNumberKey = "weekNumber";
var MonthKey = "month";
var FirstDayOfWeekKey = "firstDayOfWeek";
// CONCATENATED MODULE: ./src/utils/seriesTimeUtils.ts



function prependZeroToString(number) {
  if (number < 0) {
    number = 1;
  }

  if (number < 10) {
    return "0" + number.toString();
  }

  return number.toString();
}
function throwOnInvalidDate(year, month, day) {
  if (!isValidDate(year, month, day)) {
    throw createArgumentError(SeriesTimeKey, getString("l_InvalidDate_Text"));
  }
}
function isValidDate(year, month, day) {
  if (year < 1601 || month < 1 || month > 12 || day < 1 || day > 31) {
    return false;
  }

  return true;
}
function throwOnInvalidDateString(dateString) {
  var regEx = new RegExp("^\\d{4}-(?:[0]\\d|1[0-2])-(?:[0-2]\\d|3[01])$");

  if (!regEx.test(dateString)) {
    throw createArgumentError(SeriesTimeKey, getString("l_InvalidDate_Text"));
  }
}
// CONCATENATED MODULE: ./src/api/SeriesTime.ts






var SeriesTime_SeriesTime = function () {
  function SeriesTime() {
    this.startYear = 0;
    this.startMonth = 0;
    this.startDay = 0;
    this.endYear = 0;
    this.endMonth = 0;
    this.endDay = 0;
    this.startTimeMinutes = 0;
    this.durationMinutes = 0;
  }

  SeriesTime.prototype.getDuration = function () {
    return this.durationMinutes;
  };

  SeriesTime.prototype.getEndTime = function () {
    var endTimeMinutes = this.startTimeMinutes + this.durationMinutes;
    var minutes = endTimeMinutes % 60;
    var hours = Math.floor(endTimeMinutes / 60) % 24;
    return "T" + prependZeroToString(hours) + ":" + prependZeroToString(minutes) + ":00.000";
  };

  SeriesTime.prototype.getEndDate = function () {
    if (this.endYear === 0 && this.endMonth === 0 && this.endDay === 0) {
      return null;
    }

    return this.endYear.toString() + "-" + prependZeroToString(this.endMonth) + "-" + prependZeroToString(this.endDay);
  };

  SeriesTime.prototype.getStartDate = function () {
    return this.startYear.toString() + "-" + prependZeroToString(this.startMonth) + "-" + prependZeroToString(this.startDay);
  };

  SeriesTime.prototype.getStartTime = function () {
    var minutes = this.startTimeMinutes % 60;
    var hours = Math.floor(this.startTimeMinutes / 60);
    return "T" + prependZeroToString(hours) + ":" + prependZeroToString(minutes) + ":00.000";
  };

  SeriesTime.prototype.setDuration = function (minutes) {
    if (minutes >= 0) {
      this.durationMinutes = minutes;
    } else {
      throw createArgumentError(undefined, getString("l_InvalidTime_Text"));
    }
  };

  SeriesTime.prototype.setEndDate = function (yearOrDateString, month, day) {
    if (yearOrDateString !== null && !isNullOrUndefined(month) && day !== null) {
      this.setDateHelper(false, yearOrDateString, month, day);
    } else if (yearOrDateString !== null) {
      this.setDateHelper(false, yearOrDateString);
    } else if (yearOrDateString == null) {
      this.endYear = 0;
      this.endMonth = 0;
      this.endDay = 0;
    }
  };

  SeriesTime.prototype.setStartDate = function (yearOrDateString, month, day) {
    if (yearOrDateString !== null && !isNullOrUndefined(month) && day !== null) {
      this.setDateHelper(true, yearOrDateString, month, day);
    } else if (yearOrDateString !== null) {
      this.setDateHelper(true, yearOrDateString);
    }
  };

  SeriesTime.prototype.setStartTime = function (hoursOrTimeString, minutes) {
    if (!isNullOrUndefined(hoursOrTimeString) && !isNullOrUndefined(minutes)) {
      var totalMinutes = hoursOrTimeString * 60 + minutes;

      if (totalMinutes >= 0) {
        this.startTimeMinutes = totalMinutes;
      } else {
        throw createArgumentError(undefined, getString("l_InvalidTime_Text"));
      }
    } else if (!isNullOrUndefined(hoursOrTimeString)) {
      var timeString = hoursOrTimeString;
      var newDateString = "2017-01-15" + timeString + "Z";
      var regEx = new RegExp("^T[0-2]\\d:[0-5]\\d:[0-5]\\d\\.\\d{3}$");

      if (!regEx.test(timeString)) {
        throw createArgumentError(undefined, getString("l_InvalidTime_Text"));
      }

      var dateObject = new Date(newDateString);

      if (!isNullOrUndefined(dateObject) && !isNaN(dateObject.getUTCHours()) && !isNaN(dateObject.getUTCMinutes())) {
        this.startTimeMinutes = dateObject.getUTCHours() * 60 + dateObject.getUTCMinutes();
      } else {
        throw createArgumentError(undefined, getString("l_InvalidTime_Text"));
      }
    }
  };

  SeriesTime.prototype.isValid = function () {
    if (!isValidDate(this.startYear, this.startMonth, this.startDay)) {
      return false;
    }

    if (this.endDay !== 0 && this.endMonth !== 0 && this.endYear !== 0) {
      if (!isValidDate(this.endYear, this.endMonth, this.endDay)) {
        return false;
      }
    }

    if (this.startTimeMinutes < 0 || this.durationMinutes <= 0) {
      return false;
    }

    return true;
  };

  SeriesTime.prototype.exportToSeriesTimeJson = function () {
    var result = {};
    result[StartYearKey] = this.startYear;
    result[StartMonthKey] = this.startMonth;
    result[StartDayKey] = this.startDay;

    if (this.endYear === 0 && this.endMonth === 0 && this.endDay === 0) {
      result[NoEndDateKey] = true;
    } else {
      result[EndYearKey] = this.endYear;
      result[EndMonthKey] = this.endMonth;
      result[EndDayKey] = this.endDay;
    }

    result[StartTimeMinKey] = this.startTimeMinutes;

    if (this.durationMinutes > 0) {
      result[DurationMinKey] = this.durationMinutes;
    }

    return result;
  };

  SeriesTime.prototype.importFromSeriesTimeJsonObject = function (jsonObject) {
    this.startYear = jsonObject[StartYearKey];
    this.startMonth = jsonObject[StartMonthKey];
    this.startDay = jsonObject[StartDayKey];

    if (jsonObject[NoEndDateKey] != null && typeof jsonObject[NoEndDateKey] === "boolean") {
      this.endYear = 0;
      this.endMonth = 0;
      this.endDay = 0;
    } else {
      this.endYear = jsonObject[EndYearKey];
      this.endMonth = jsonObject[EndMonthKey];
      this.endDay = jsonObject[EndDayKey];
    }

    this.startTimeMinutes = jsonObject[StartTimeMinKey];
    this.durationMinutes = jsonObject[DurationMinKey];
  };

  SeriesTime.prototype.setDateHelper = function (isStart, yearOrDateString, month, day) {
    var yearCalculated = 0;
    var monthCalculated = 0;
    var dayCalculated = 0;

    if (yearOrDateString !== null && !isNullOrUndefined(month) && day !== null) {
      throwOnInvalidDate(yearOrDateString, month + 1, day);
      yearCalculated = yearOrDateString;
      monthCalculated = month + 1;
      dayCalculated = day;
    } else if (yearOrDateString !== null) {
      var dateString = yearOrDateString;
      throwOnInvalidDateString(dateString);
      var dateObject = new Date(dateString);

      if (dateObject !== null && !isNaN(dateObject.getUTCFullYear()) && !isNaN(dateObject.getUTCMonth()) && !isNaN(dateObject.getUTCDate())) {
        throwOnInvalidDate(dateObject.getUTCFullYear(), dateObject.getUTCMonth() + 1, dateObject.getUTCDate());
        yearCalculated = dateObject.getUTCFullYear();
        monthCalculated = dateObject.getUTCMonth() + 1;
        dayCalculated = dateObject.getUTCDate();
      }
    }

    if (yearCalculated !== 0 && monthCalculated !== 0 && dayCalculated !== 0) {
      if (isStart) {
        this.startYear = yearCalculated;
        this.startMonth = monthCalculated;
        this.startDay = dayCalculated;
      } else {
        this.endYear = yearCalculated;
        this.endMonth = monthCalculated;
        this.endDay = dayCalculated;
      }
    }
  };

  SeriesTime.prototype.isEndAfterStart = function () {
    if (this.endYear === 0 && this.endMonth === 0 && this.endDay === 0) {
      return true;
    }

    var startDateTime = new Date();
    startDateTime.setFullYear(this.startYear);
    startDateTime.setMonth(this.startMonth - 1);
    startDateTime.setDate(this.startDay);
    var endDateTime = new Date();
    endDateTime.setFullYear(this.endYear);
    endDateTime.setMonth(this.endMonth - 1);
    endDateTime.setDate(this.endDay);
    return endDateTime >= startDateTime;
  };

  return SeriesTime;
}();


// CONCATENATED MODULE: ./src/utils/recurrenceUtils.ts



function copyRecurrenceObjectConvertSeriesTimeJson(recurrenceOriginal) {
  if (isNullOrUndefined(recurrenceOriginal) || isNullOrUndefined(recurrenceOriginal.seriesTimeJson)) {
    return recurrenceOriginal;
  }

  var recurrenceCopy = {
    recurrenceType: "",
    recurrenceProperties: null,
    recurrenceTimeZone: null
  };
  var newSeriesTime = new SeriesTime_SeriesTime();

  if (!isNullOrUndefined(recurrenceOriginal.recurrenceProperties)) {
    recurrenceCopy.recurrenceProperties = deepClone(recurrenceOriginal.recurrenceProperties);
  }

  recurrenceCopy.recurrenceType = recurrenceOriginal.recurrenceType;

  if (!isNullOrUndefined(recurrenceOriginal.recurrenceTimeZone)) {
    recurrenceCopy.recurrenceTimeZone = deepClone(recurrenceOriginal.recurrenceTimeZone);
  }

  newSeriesTime.importFromSeriesTimeJsonObject(recurrenceOriginal.seriesTimeJson);
  recurrenceCopy.seriesTime = newSeriesTime;
  return recurrenceCopy;
}
// CONCATENATED MODULE: ./src/api/getMessageRead.ts
















function getMessageRead() {
  var sender = getInitialDataProp("sender");
  var from = getInitialDataProp("from");
  var dateTimeCreated = getInitialDataProp("dateTimeCreated");
  var dateTimeModified = getInitialDataProp("dateTimeModified");
  var end = getInitialDataProp("end");
  var start = getInitialDataProp("start");
  var messageRead = objectDefine({}, {
    attachments: CustomJsonAttachmentsResponse(getInitialDataProp("attachments")),
    bcc: (getInitialDataProp("bcc") || []).map(createEmailAddressDetails),
    body: getBodySurface(false),
    categories: getCategoriesSurface(),
    cc: (getInitialDataProp("cc") || []).map(createEmailAddressDetails),
    conversationId: getInitialDataProp("conversationId"),
    dateTimeCreated: dateTimeCreated ? new Date(dateTimeCreated) : undefined,
    dateTimeModified: dateTimeModified ? new Date(dateTimeModified) : undefined,
    end: end ? new Date(end) : undefined,
    from: from ? createEmailAddressDetails(from) : undefined,
    getAllInternetHeadersAsync: getAllInternetHeaders,
    internetMessageId: getInitialDataProp("internetMessageId"),
    itemClass: getInitialDataProp("itemClass"),
    itemId: getInitialDataProp("id"),
    itemType: "message",
    location: getInitialDataProp("location"),
    move: moveToFolder,
    normalizedSubject: getInitialDataProp("normalizedSubject"),
    notificationMessages: getNotificationMessageSurface(),
    recurrence: copyRecurrenceObjectConvertSeriesTimeJson(getInitialDataProp("recurrence")),
    seriesId: getInitialDataProp("seriesId"),
    sender: sender ? createEmailAddressDetails(sender) : undefined,
    start: start ? new Date(start) : undefined,
    subject: getInitialDataProp("subject"),
    to: (getInitialDataProp("to") || []).map(createEmailAddressDetails),
    displayReplyForm: displayReplyForm,
    displayReplyFormAsync: displayReplyFormAsync,
    displayReplyAllForm: displayReplyAllForm,
    displayReplyAllFormAsync: displayReplyAllFormAsync,
    getAttachmentContentAsync: getAttachmentContent,
    getEntities: Entities_getEntities,
    getEntitiesByType: Entities_getEntitiesByType,
    getFilteredEntitiesByName: Entities_getFilteredEntitiesByName,
    getInitializationContextAsync: getInitializationContext,
    getRegExMatches: Entities_getRegExMatches,
    getRegExMatchesByName: Entities_getRegExMatchesByName,
    getSelectedEntities: Entities_getSelectedEntities,
    getSelectedRegExMatches: Entities_getSelectedRegExMatches,
    loadCustomPropertiesAsync: loadCustomProperties,
    delayDeliveryTime: getDelayDeliverySurface(false),
    isAllDayEvent: getInitialDataProp("isAllDayEvent"),
    sensitivity: getInitialDataProp("sensitivity")
  });
  return messageRead;
}
// CONCATENATED MODULE: ./src/validation/validateAttachments.ts




function validateAddFileAttachmentApis(attachmentName) {
  if (isNullOrUndefined(attachmentName) || attachmentName === "" || !(typeof attachmentName === "string")) {
    throw createArgumentError("attachmentName");
  }

  throwOnOutOfRange(attachmentName.length, 0, MaxAttachmentNameLength, "attachmentName");
}
// CONCATENATED MODULE: ./src/validation/attachmentsConstants.ts
var AddItemAttachmentClientEndPointTimeoutInMilliseconds = 600000;
var MaxBase64AttachmentSize = 27892122;
// CONCATENATED MODULE: ./src/methods/addFileAttachment.ts








function addFileAttachment(uri, attachmentName) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "item.addFileAttachmentAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var isInline = false;

  if (!!commonParameters.options) {
    isInline = !!commonParameters.options.isInline;
  }

  var name = attachmentName;
  var parameters = {
    uri: uri,
    name: name,
    isInline: isInline,
    __timeout__: AddItemAttachmentClientEndPointTimeoutInMilliseconds
  };
  addFileAttachment_validateParameters(parameters);
  standardInvokeHostMethod(16, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function addFileAttachment_validateParameters(parameters) {
  validateStringParam("uri", parameters.uri);
  throwOnOutOfRange(parameters.uri.length, 0, MaxUrlLength, "uri");
  validateAddFileAttachmentApis(parameters.name);
}
// CONCATENATED MODULE: ./src/methods/addBase64FileAttachment.ts







function addBase64FileAttachment(base64String, name) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "item.addBase64FileAttachmentAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var isInline = false;

  if (!!commonParameters.options) {
    isInline = !!commonParameters.options.isInline;
  }

  var parameters = {
    base64String: base64String,
    name: name,
    isInline: isInline,
    __timeout__: AddItemAttachmentClientEndPointTimeoutInMilliseconds
  };
  addBase64FileAttachment_validateParameters(parameters);
  standardInvokeHostMethod(148, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function addBase64FileAttachment_validateParameters(parameters) {
  validateStringParam("base64Encoded", parameters.base64String);
  throwOnOutOfRange(parameters.base64String.length, 0, MaxBase64AttachmentSize, "base64File");
  validateAddFileAttachmentApis(parameters.name);
}
// CONCATENATED MODULE: ./src/methods/addItemAttachment.ts








function addItemAttachment(itemId, name) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "item.addItemAttachmentAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    itemId: itemId,
    name: name
  };
  addItemAttachment_validateParameters(parameters);
  standardInvokeHostMethod(19, commonParameters.asyncContext, commonParameters.callback, {
    itemId: getItemIdBasedOnHost(parameters.itemId),
    name: parameters.name,
    __timeout__: AddItemAttachmentClientEndPointTimeoutInMilliseconds
  }, undefined);
}

function addItemAttachment_validateParameters(parameters) {
  validateStringParam("itemId", parameters.itemId);
  validateStringParam("attachmentName", parameters.name);
  throwOnOutOfRange(parameters.itemId.length, 0, MaxItemIdLength, "itemId");
  throwOnOutOfRange(parameters.name.length, 0, MaxAttachmentNameLength, "attachmentName");
}
// CONCATENATED MODULE: ./src/methods/close.ts

function close_close() {
  standardInvokeHostMethod(41, undefined, undefined, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/getAttachments.ts




function getAttachments_getAttachments() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.getAttachmentsAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(149, commonParameters.asyncContext, commonParameters.callback, undefined, CustomJsonAttachmentsResponse);
}
// CONCATENATED MODULE: ./src/methods/getSelectedData.ts





function getSelectedData(coercionType) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "item.getSelectedDataAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  var parameters = {
    coercionType: getCoercionTypeFromString(coercionType)
  };

  if (parameters.coercionType === undefined) {
    throw createArgumentError("coercionType");
  }

  standardInvokeHostMethod(28, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/removeAttachment.ts






function removeAttachment(attachmentId) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "item.removeAttachmentAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    attachmentIndex: attachmentId
  };
  removeAttachment_validateParameters(parameters);
  standardInvokeHostMethod(20, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function removeAttachment_validateParameters(parameters) {
  validateStringParam("attachmentId", parameters.attachmentIndex);
  throwOnOutOfRange(parameters.attachmentIndex.length, 0, MaxRemoveIdLength, "attachmentId");
}
// CONCATENATED MODULE: ./src/methods/save.ts



function save() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "item.saveAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  standardInvokeHostMethod(32, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/validation/validateRecipientParameters.ts




function validateRecipientParameters(parameters) {
  if (Array.isArray(parameters.recipientArray)) {
    if (parameters.recipientArray.length > recipientsLimit) {
      throw createArgumentOutOfRange("recipients", parameters.recipientArray.length);
    }

    var validatedRecipients = parameters.recipientArray.map(function (recipient) {
      if (isNullOrUndefined(recipient)) {
        throw createArgumentError("recipients");
      }

      if (typeof recipient === "string") {
        throwOnInvalidDisplayNameOrEmail(recipient, recipient);
        return createEmailAddressForHost(recipient, recipient);
      } else if (typeof recipient === "object") {
        throwOnInvalidDisplayNameOrEmail(recipient.displayName, recipient.emailAddress);
        return createEmailAddressForHost(recipient.displayName, recipient.emailAddress);
      } else {
        throw createArgumentError("recipients");
      }
    });
    parameters.recipientArray = validatedRecipients;
  } else {
    throw createArgumentError("recipients");
  }
}

function throwOnInvalidDisplayNameOrEmail(displayName, email) {
  if (!displayName && !email) {
    throw createArgumentError("recipients");
  } else if (typeof displayName === "string" && displayName.length > displayNameLengthLimit) {
    throw createArgumentOutOfRange("recipients", displayName.length, getString("l_DisplayNameTooLong_Text"));
  } else if (typeof email === "string" && email.length > maxSmtpLength) {
    throw createArgumentOutOfRange("recipients", email.length, getString("l_EmailAddressTooLong_Text"));
  } else if (typeof displayName !== "string" && typeof email !== "string") {
    throw createArgumentError("recipients");
  }
}

function createEmailAddressForHost(displayName, email) {
  return {
    address: email,
    name: displayName
  };
}
// CONCATENATED MODULE: ./src/methods/addRecipients.ts





function addRecipients(namespace) {
  return function (recipientArray) {
    var args = [];

    for (var _i = 1; _i < arguments.length; _i++) {
      args[_i - 1] = arguments[_i];
    }

    checkPermissionsAndThrow(2, namespace + ".addAsync");
    var commonParameters = parseCommonArgs(args, false, false);
    var parameters = {
      recipientField: RecipientFields[namespace],
      recipientArray: recipientArray
    };
    validateRecipientParameters(parameters);
    standardInvokeHostMethod(22, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
  };
}
// CONCATENATED MODULE: ./src/methods/getRecipients.ts





function getRecipients(namespace) {
  return function () {
    var args = [];

    for (var _i = 0; _i < arguments.length; _i++) {
      args[_i] = arguments[_i];
    }

    checkPermissionsAndThrow(1, namespace + ".getAsync");
    var commonParameters = parseCommonArgs(args, true, false);
    standardInvokeHostMethod(15, commonParameters.asyncContext, commonParameters.callback, {
      recipientField: RecipientFields[namespace]
    }, getRecipients_format);
  };
}

function getRecipients_format(rawInput) {
  if (rawInput === null || rawInput === undefined) {
    return [];
  }

  return rawInput.map(function (input) {
    return createEmailAddressDetails(input);
  });
}
// CONCATENATED MODULE: ./src/methods/setRecipients.ts





function setRecipients(namespace) {
  return function (recipientArray) {
    var args = [];

    for (var _i = 1; _i < arguments.length; _i++) {
      args[_i - 1] = arguments[_i];
    }

    checkPermissionsAndThrow(2, namespace + ".setAsync");
    var commonParameters = parseCommonArgs(args, false, false);
    var parameters = {
      recipientField: RecipientFields[namespace],
      recipientArray: recipientArray
    };
    validateRecipientParameters(parameters);
    standardInvokeHostMethod(21, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
  };
}
// CONCATENATED MODULE: ./src/api/getRecipientsSurface.ts




function getRecipientsSurface(namespace) {
  return objectDefine({}, {
    addAsync: addRecipients(namespace),
    getAsync: getRecipients(namespace),
    setAsync: setRecipients(namespace)
  });
}
// CONCATENATED MODULE: ./src/methods/getFrom.ts





function getFrom(namespace) {
  return function () {
    var args = [];

    for (var _i = 0; _i < arguments.length; _i++) {
      args[_i] = arguments[_i];
    }

    checkPermissionsAndThrow(1, namespace + ".getAsync");
    var commonParameters = parseCommonArgs(args, true, false);
    standardInvokeHostMethod(107, commonParameters.asyncContext, commonParameters.callback, undefined, getFrom_format);
  };
}

function getFrom_format(rawInput) {
  return isNullOrUndefined(rawInput) ? null : createEmailAddressDetails(rawInput);
}
// CONCATENATED MODULE: ./src/api/getFromSurface.ts


function getFromSurface(namespace) {
  return objectDefine({}, {
    getAsync: getFrom(namespace)
  });
}
// CONCATENATED MODULE: ./src/validation/validateInternetHeaders.ts



function validateInternetHeaderArray(internetHeaderArray) {
  if (isNullOrUndefined(internetHeaderArray)) {
    throw createArgumentError("internetHeaders");
  }

  if (!Array.isArray(internetHeaderArray)) {
    throw createArgumentTypeError("internetHeaders", typeof internetHeaderArray, "Array");
  }

  if (internetHeaderArray.length === 0) {
    throw createArgumentError("internetHeaders");
  }

  for (var _i = 0, internetHeaderArray_1 = internetHeaderArray; _i < internetHeaderArray_1.length; _i++) {
    var internetHeader = internetHeaderArray_1[_i];
    validateStringParam("internetHeaders", internetHeader);
  }
}
// CONCATENATED MODULE: ./src/methods/removeInternetHeaders.ts




function removeInternetHeaders(internetHeaderKeys) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "internetHeaders.removeAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    internetHeaderKeys: internetHeaderKeys
  };
  removeInternetHeaders_validateParameters(parameters);
  standardInvokeHostMethod(153, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function removeInternetHeaders_validateParameters(parameters) {
  validateInternetHeaderArray(parameters.internetHeaderKeys);
}
// CONCATENATED MODULE: ./src/methods/getInternetHeaders.ts




function getInternetHeaders(internetHeaderKeys) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "internetHeaders.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  var parameters = {
    internetHeaderKeys: internetHeaderKeys
  };
  getInternetHeaders_validateParameters(parameters);
  standardInvokeHostMethod(151, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function getInternetHeaders_validateParameters(parameters) {
  validateInternetHeaderArray(parameters.internetHeaderKeys);
}
// CONCATENATED MODULE: ./src/methods/setInternetHeaders.ts







var InternetHeadersLimit = 998;
function setInternetHeaders(internetHeaderNameValuePairs) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "internetHeaders.setAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    internetHeaderNameValuePairs: internetHeaderNameValuePairs
  };
  setInternetHeaders_validateParameters(parameters);
  standardInvokeHostMethod(152, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function setInternetHeaders_validateParameters(parameters) {
  if (isNullOrUndefined(parameters.internetHeaderNameValuePairs)) {
    throw createNullArgumentError("internetHeaders");
  }

  var keys = Object.keys(parameters.internetHeaderNameValuePairs);

  if (keys.length === 0) {
    throw createArgumentError("internetHeaders");
  }

  for (var _i = 0, keys_1 = keys; _i < keys_1.length; _i++) {
    var key = keys_1[_i];
    var value = parameters.internetHeaderNameValuePairs[key];
    validateStringParam("internetHeaders", key);

    if (!(typeof value === "string")) {
      throw createArgumentTypeError("internetHeaders", typeof value, "string");
    }

    throwOnOutOfRange(key.length + value.length, 0, InternetHeadersLimit, key);
  }
}
// CONCATENATED MODULE: ./src/api/getInternetHeadersSurface.ts




function getInternetHeadersSurface(isCompose) {
  var internetHeaders = objectDefine({}, {
    getAsync: getInternetHeaders
  });

  if (isCompose) {
    objectDefine(internetHeaders, {
      removeAsync: removeInternetHeaders,
      setAsync: setInternetHeaders
    });
  }

  return internetHeaders;
}
// CONCATENATED MODULE: ./src/methods/getSubject.ts



function getSubject() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "subject.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(18, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/setSubject.ts





var MaximumSubjectLength = 255;
function setSubject(subject) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "subject.setAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    subject: subject
  };
  setSubject_validateParameters(parameters);
  standardInvokeHostMethod(17, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function setSubject_validateParameters(parameters) {
  if (!(typeof parameters.subject === "string")) {
    throw createArgumentTypeError("subject", typeof parameters.subject, "string");
  }

  throwOnOutOfRange(parameters.subject.length, 0, MaximumSubjectLength, "subject");
}
// CONCATENATED MODULE: ./src/api/getSubjectSurface.ts



function getSubjectSurface() {
  return objectDefine({}, {
    getAsync: getSubject,
    setAsync: setSubject
  });
}
// CONCATENATED MODULE: ./src/methods/getItemId.ts



function getItemId() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.getItemIdAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(164, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/getComposeType.ts




function getComposeType() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.getComposeTypeAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  checkFeatureEnabledAndThrow(Features.signature, "getComposeTypeAsync");
  standardInvokeHostMethod(174, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/isClientSignatureEnabled.ts




function isClientSignatureEnabled() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "isClientSignatureEnabledAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  checkFeatureEnabledAndThrow(Features.signature, "isClientSignatureEnabledAsync");
  standardInvokeHostMethod(175, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/disableClientSignature.ts




function disableClientSignature() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "disableClientSignatureAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  checkFeatureEnabledAndThrow(Features.signature, "disableClientSignatureAsync");
  standardInvokeHostMethod(176, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/getSessionData.ts





function getSessionData(name) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "sessionData.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  checkFeatureEnabledAndThrow(Features.sessionData, "sessionData.getAsync");
  var parameters = {
    name: name
  };
  getSessionData_validateParameters(parameters);
  standardInvokeHostMethod(186, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function getSessionData_validateParameters(parameters) {
  validateStringParam("name", parameters.name);
}
// CONCATENATED MODULE: ./src/methods/setSessionData.ts





function setSessionData(name, value) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "sessionData.setAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    name: name,
    value: value
  };
  checkFeatureEnabledAndThrow(Features.sessionData, "sessionData.setAsync");
  setSessionData_validateParameters(parameters);
  standardInvokeHostMethod(185, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function setSessionData_validateParameters(parameters) {
  validateStringParam("name", parameters.name);
  validateStringParamWithEmptyAllowed("value", parameters.value);
}
// CONCATENATED MODULE: ./src/methods/getAllSessionData.ts




function getAllSessionData() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "sessionData.getAllAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  checkFeatureEnabledAndThrow(Features.sessionData, "sessionData.getAllAsync");
  standardInvokeHostMethod(187, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/clearSessionData.ts




function clearSessionData() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "sessionData.clearAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  checkFeatureEnabledAndThrow(Features.sessionData, "sessionData.clearAsync");
  standardInvokeHostMethod(188, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/removeSessionData.ts





function removeSessionData(name) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "sessionData.removeAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    name: name
  };
  checkFeatureEnabledAndThrow(Features.sessionData, "sessionData.removeAsync");
  removeSessionData_validateParameters(parameters);
  standardInvokeHostMethod(189, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function removeSessionData_validateParameters(parameters) {
  validateStringParam("name", parameters.name);
}
// CONCATENATED MODULE: ./src/api/getSessionDataSurface.ts






function getSessionDataSurface() {
  return objectDefine({}, {
    getAsync: getSessionData,
    setAsync: setSessionData,
    getAllAsync: getAllSessionData,
    clearAsync: clearSessionData,
    removeAsync: removeSessionData
  });
}
// CONCATENATED MODULE: ./src/api/getMessageCompose.ts



























function getMessageCompose() {
  var messageCompose = objectDefine({}, {
    bcc: getRecipientsSurface("bcc"),
    body: getBodySurface(true),
    categories: getCategoriesSurface(),
    cc: getRecipientsSurface("cc"),
    conversationId: getInitialDataProp("conversationId"),
    from: getFromSurface("from"),
    internetHeaders: getInternetHeadersSurface(true),
    itemType: "message",
    notificationMessages: getNotificationMessageSurface(),
    seriesId: getInitialDataProp("seriesId"),
    subject: getSubjectSurface(),
    to: getRecipientsSurface("to"),
    addFileAttachmentAsync: addFileAttachment,
    addFileAttachmentFromBase64Async: addBase64FileAttachment,
    addItemAttachmentAsync: addItemAttachment,
    close: close_close,
    getAttachmentsAsync: getAttachments_getAttachments,
    getAttachmentContentAsync: getAttachmentContent,
    getInitializationContextAsync: getInitializationContext,
    getItemIdAsync: getItemId,
    getSelectedDataAsync: getSelectedData,
    loadCustomPropertiesAsync: loadCustomProperties,
    removeAttachmentAsync: removeAttachment,
    saveAsync: save,
    setSelectedDataAsync: setSelectedData(29),
    delayDeliveryTime: getDelayDeliverySurface(true),
    getComposeTypeAsync: getComposeType,
    isClientSignatureEnabledAsync: isClientSignatureEnabled,
    disableClientSignatureAsync: disableClientSignature,
    sessionData: getSessionDataSurface()
  });
  return messageCompose;
}
// CONCATENATED MODULE: ./src/validation/validateEnhancedLocation.ts




function validateLocationIdentifiers(locationIdentifiers) {
  if (isNullOrUndefined(locationIdentifiers)) {
    throw createNullArgumentError("locationIdentifier");
  }

  if (!Array.isArray(locationIdentifiers)) {
    throw createArgumentTypeError("locationIdentifier", typeof locationIdentifiers, "Array");
  }

  if (locationIdentifiers.length === 0) {
    throw createArgumentError("locationIdentifier");
  }

  for (var _i = 0, locationIdentifiers_1 = locationIdentifiers; _i < locationIdentifiers_1.length; _i++) {
    var locationIdentifier = locationIdentifiers_1[_i];
    validateLocationIdentifier(locationIdentifier);
  }
}

function validateLocationIdentifier(locationIdentifier) {
  if (isNullOrUndefined(locationIdentifier) || isNullOrUndefined(locationIdentifier.id) || isNullOrUndefined(locationIdentifier.type)) {
    throw createNullArgumentError("locationIdentifier");
  }

  if (locationIdentifier.type === MailboxEnums.LocationType.Room || locationIdentifier.type === MailboxEnums.LocationType.Custom) {
    validateIdParameter(locationIdentifier.id, locationIdentifier.type);
  } else {
    throw createArgumentError("type");
  }
}

function validateIdParameter(id, type) {
  if (id === "") {
    throw createArgumentError("id");
  }

  if (type === MailboxEnums.LocationType.Room) {
    if (id.length > maxSmtpLength) {
      throw createArgumentError("id");
    }
  }
}
// CONCATENATED MODULE: ./src/methods/addEnhancedLocations.ts




function addEnhancedLocations(enhancedLocations) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "enhancedLocations.addAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    enhancedLocations: enhancedLocations
  };
  addEnhancedLocations_validateParameters(parameters);
  standardInvokeHostMethod(155, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function addEnhancedLocations_validateParameters(parameters) {
  validateLocationIdentifiers(parameters.enhancedLocations);
}
// CONCATENATED MODULE: ./src/methods/getEnhancedLocations.ts



function getEnhancedLocations() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "enhancedLocations.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(154, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/removeEnhancedLocations.ts




function removeEnhancedLocations(enhancedLocations) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "enhancedLocations.removeAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    enhancedLocations: enhancedLocations
  };
  removeEnhancedLocations_validateParameters(parameters);
  standardInvokeHostMethod(156, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function removeEnhancedLocations_validateParameters(parameters) {
  validateLocationIdentifiers(parameters.enhancedLocations);
}
// CONCATENATED MODULE: ./src/api/getEnhancedLocationSurface.ts




function getEnhancedLocationsSurface(isCompose) {
  var enhancedLocations = objectDefine({}, {
    getAsync: getEnhancedLocations
  });

  if (isCompose) {
    objectDefine(enhancedLocations, {
      addAsync: addEnhancedLocations,
      removeAsync: removeEnhancedLocations
    });
  }

  return enhancedLocations;
}
// CONCATENATED MODULE: ./src/api/getAppointmentRead.ts














function getAppointmentRead() {
  var organizer = getInitialDataProp("organizer");
  var dateTimeCreated = getInitialDataProp("dateTimeCreated");
  var dateTimeModified = getInitialDataProp("dateTimeModified");
  var end = getInitialDataProp("end");
  var start = getInitialDataProp("start");
  var appointmentRead = objectDefine({}, {
    attachments: CustomJsonAttachmentsResponse(getInitialDataProp("attachments")),
    body: getBodySurface(false),
    categories: getCategoriesSurface(),
    dateTimeCreated: dateTimeCreated ? new Date(dateTimeCreated) : undefined,
    dateTimeModified: dateTimeModified ? new Date(dateTimeModified) : undefined,
    end: end ? new Date(end) : undefined,
    enhancedLocation: getEnhancedLocationsSurface(false),
    itemClass: getInitialDataProp("itemClass"),
    itemId: getInitialDataProp("id"),
    itemType: "appointment",
    location: getInitialDataProp("location"),
    normalizedSubject: getInitialDataProp("normalizedSubject"),
    notificationMessages: getNotificationMessageSurface(),
    optionalAttendees: (getInitialDataProp("cc") || []).map(createEmailAddressDetails),
    organizer: organizer ? createEmailAddressDetails(organizer) : undefined,
    recurrence: copyRecurrenceObjectConvertSeriesTimeJson(getInitialDataProp("recurrence")),
    requiredAttendees: (getInitialDataProp("to") || []).map(createEmailAddressDetails),
    start: start ? new Date(start) : undefined,
    seriesId: getInitialDataProp("seriesId"),
    subject: getInitialDataProp("subject"),
    displayReplyForm: displayReplyForm,
    displayReplyFormAsync: displayReplyFormAsync,
    displayReplyAllForm: displayReplyAllForm,
    displayReplyAllFormAsync: displayReplyAllFormAsync,
    getAttachmentContentAsync: getAttachmentContent,
    getEntities: Entities_getEntities,
    getEntitiesByType: Entities_getEntitiesByType,
    getFilteredEntitiesByName: Entities_getFilteredEntitiesByName,
    getInitializationContextAsync: getInitializationContext,
    getRegExMatches: Entities_getRegExMatches,
    getRegExMatchesByName: Entities_getRegExMatchesByName,
    getSelectedEntities: Entities_getSelectedEntities,
    getSelectedRegExMatches: Entities_getSelectedRegExMatches,
    loadCustomPropertiesAsync: loadCustomProperties,
    isAllDayEvent: getInitialDataProp("isAllDayEvent"),
    sensitivity: getInitialDataProp("sensitivity")
  });
  return appointmentRead;
}
// CONCATENATED MODULE: ./src/validation/timeConstants.ts
var TimeType;

(function (TimeType) {
  TimeType[TimeType["start"] = 1] = "start";
  TimeType[TimeType["end"] = 2] = "end";
})(TimeType || (TimeType = {}));
// CONCATENATED MODULE: ./src/methods/getTime.ts




function getTime(namespace) {
  return function () {
    var args = [];

    for (var _i = 0; _i < arguments.length; _i++) {
      args[_i] = arguments[_i];
    }

    checkPermissionsAndThrow(1, namespace + ".getAsync");
    var commonParameters = parseCommonArgs(args, true, false);
    standardInvokeHostMethod(24, commonParameters.asyncContext, commonParameters.callback, {
      TimeProperty: TimeType[namespace]
    }, getTime_format);
  };
}

function getTime_format(rawInput) {
  var ticks = rawInput;
  return new Date(ticks);
}
// CONCATENATED MODULE: ./src/methods/setTime.ts






var maxTime = 8640000000000000;
var minTime = -8640000000000000;
function setTime(namespace) {
  return function (date) {
    var args = [];

    for (var _i = 1; _i < arguments.length; _i++) {
      args[_i - 1] = arguments[_i];
    }

    checkPermissionsAndThrow(2, namespace + ".setAsync");
    var commonParameters = parseCommonArgs(args, false, false);
    var parameters = {
      date: date
    };
    setTime_validateParameters(parameters);
    standardInvokeHostMethod(25, commonParameters.asyncContext, commonParameters.callback, {
      TimeProperty: TimeType[namespace],
      time: parameters.date.getTime()
    }, undefined);
  };
}

function setTime_validateParameters(parameters) {
  if (!isDateObject(parameters.date)) {
    throw createArgumentTypeError("dateTime", typeof parameters.date, typeof Date);
  }

  if (isNaN(parameters.date.getTime())) {
    throw createArgumentError("dateTime");
  }

  if (parameters.date.getTime() < minTime || parameters.date.getTime() > maxTime) {
    throw createArgumentOutOfRange("dateTime");
  }
}
// CONCATENATED MODULE: ./src/api/getTimeSurface.ts



function getTimeSurface(namespace) {
  return objectDefine({}, {
    getAsync: getTime(namespace),
    setAsync: setTime(namespace)
  });
}
// CONCATENATED MODULE: ./src/methods/getLocation.ts



function getLocation() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "location.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(26, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/setLocation.ts






var MaximumLocationLength = 255;
function setLocation(location) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "location.setAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    location: location
  };
  setLocation_validateParameters(parameters);
  standardInvokeHostMethod(27, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function setLocation_validateParameters(parameters) {
  if (!isNullOrUndefined(parameters.location)) {
    if (!(typeof parameters.location === "string")) {
      throw createArgumentTypeError("location", typeof parameters.location, "string");
    }

    throwOnOutOfRange(parameters.location.length, 0, MaximumLocationLength, "location");
  } else {
    throw createNullArgumentError("location");
  }
}
// CONCATENATED MODULE: ./src/api/getLocationSurface.ts



function getLocationSurface() {
  return objectDefine({}, {
    getAsync: getLocation,
    setAsync: setLocation
  });
}
// CONCATENATED MODULE: ./src/methods/getRecurrence.ts




function getRecurrence() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "recurrenceProperties.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(103, commonParameters.asyncContext, commonParameters.callback, undefined, seriesTimeJsonConverter);
}
function seriesTimeJsonConverter(rawInput) {
  if (rawInput !== null) {
    if (rawInput.seriesTimeJson !== null) {
      var seriesTime = new SeriesTime_SeriesTime();
      seriesTime.importFromSeriesTimeJsonObject(rawInput.seriesTimeJson);
      delete rawInput.seriesTimeJson;
      rawInput.seriesTime = seriesTime;
    }
  }

  return rawInput;
}
// CONCATENATED MODULE: ./src/validation/validateRecurrenceObject.ts






function validateRecurrenceObject(recurrenceObject) {
  if (isNullOrUndefined(recurrenceObject)) {
    return;
  }

  recurrenceObject = recurrenceObject;

  if (isNullOrUndefined(recurrenceObject.recurrenceType)) {
    throw createNullArgumentError(RecurrenceTypeKey);
  }

  if (isNullOrUndefined(recurrenceObject.seriesTime)) {
    throw createNullArgumentError(SeriesTimeKey);
  }

  if (!(recurrenceObject.seriesTime instanceof SeriesTime_SeriesTime) || !recurrenceObject.seriesTime.isValid()) {
    throw createArgumentError(SeriesTimeKey);
  }

  if (!recurrenceObject.seriesTime.isEndAfterStart()) {
    throw createArgumentError(SeriesTimeKey, getString("l_InvalidEventDates_Text"));
  }

  throwOnInvalidRecurrenceType(recurrenceObject.recurrenceType);

  if (recurrenceObject.recurrenceType !== MailboxEnums.RecurrenceType.Weekday) {
    if (isNullOrUndefined(recurrenceObject.recurrenceProperties)) {
      throw createNullArgumentError(RecurrenceTypeKey);
    }
  }

  if (!isNullOrUndefined(recurrenceObject.recurrenceTimeZone)) {
    if (isNullOrUndefined(recurrenceObject.recurrenceTimeZone.name)) {
      throw createNullArgumentError(RecurrenceTimeZoneName);
    }

    if (typeof recurrenceObject.recurrenceTimeZone.name !== "string") {
      throw createArgumentTypeError(RecurrenceTimeZoneName, typeof recurrenceObject.recurrenceTimeZone.name, "string");
    }
  }

  if (recurrenceObject.recurrenceType === MailboxEnums.RecurrenceType.Daily) {
    throwOnInvalidDailyRecurrence(recurrenceObject.recurrenceProperties);
  } else if (recurrenceObject.recurrenceType === MailboxEnums.RecurrenceType.Weekly) {
    throwOnInvalidWeeklyRecurrence(recurrenceObject.recurrenceProperties);
  } else if (recurrenceObject.recurrenceType === MailboxEnums.RecurrenceType.Monthly) {
    throwOnInvalidMonthlyRecurrence(recurrenceObject.recurrenceProperties);
  } else if (recurrenceObject.recurrenceType === MailboxEnums.RecurrenceType.Yearly) {
    throwOnInvalidYearlyRecurrence(recurrenceObject.recurrenceProperties);
  }
}

function throwOnInvalidRecurrenceType(recurrenceType) {
  if (recurrenceType !== MailboxEnums.RecurrenceType.Daily && recurrenceType !== MailboxEnums.RecurrenceType.Weekly && recurrenceType !== MailboxEnums.RecurrenceType.Weekday && recurrenceType !== MailboxEnums.RecurrenceType.Yearly && recurrenceType !== MailboxEnums.RecurrenceType.Monthly) {
    throw createArgumentError(RecurrenceTypeKey);
  }
}

function throwOnInvalidRecurrenceInterval(recurrenceProperties) {
  if (isNullOrUndefined(recurrenceProperties.interval)) {
    throw createNullArgumentError(IntervalKey);
  }

  if (typeof recurrenceProperties.interval !== "number") {
    throw createArgumentTypeError(IntervalKey, typeof recurrenceProperties.interval, "number");
  }

  if (recurrenceProperties.interval <= 0) {
    throw createArgumentError(IntervalKey);
  }
}

function throwOnInvalidDailyRecurrence(recurrenceProperties) {
  throwOnInvalidRecurrenceInterval(recurrenceProperties);
}

function throwOnInvalidWeeklyRecurrence(recurrenceProperties) {
  throwOnInvalidRecurrenceInterval(recurrenceProperties);

  if (isNullOrUndefined(recurrenceProperties.days)) {
    throw createArgumentTypeError(DaysKey);
  }

  if (!Array.isArray(recurrenceProperties.days)) {
    throw createArgumentTypeError(DaysKey);
  }

  throwOnInvalidDaysArray(recurrenceProperties.days);

  if (!isNullOrUndefined(recurrenceProperties.firstDayOfWeek)) {
    if (typeof recurrenceProperties.firstDayOfWeek !== "string") {
      throw createArgumentTypeError(FirstDayOfWeekKey);
    }

    if (!verifyDays(recurrenceProperties.firstDayOfWeek, false)) {
      throw createArgumentError(FirstDayOfWeekKey);
    }
  }
}

function throwOnInvalidDaysArray(daysArray) {
  for (var i = 0; i < daysArray.length; i++) {
    if (!verifyDays(daysArray[i], false)) {
      throw createArgumentError(DaysKey);
    }
  }
}

function verifyDays(dayEnum, checkGroupedDays) {
  var fRegularDay = dayEnum === MailboxEnums.Days.Mon || dayEnum === MailboxEnums.Days.Tue || dayEnum === MailboxEnums.Days.Wed || dayEnum === MailboxEnums.Days.Thu || dayEnum === MailboxEnums.Days.Fri || dayEnum === MailboxEnums.Days.Sat || dayEnum === MailboxEnums.Days.Sun;

  if (checkGroupedDays) {
    var fGroupedDay = dayEnum === MailboxEnums.Days.WeekendDay || dayEnum === MailboxEnums.Days.Weekday || dayEnum === MailboxEnums.Days.Day;
    return fGroupedDay || fRegularDay;
  } else {
    return fRegularDay;
  }
}

function throwOnInvalidMonthlyRecurrence(recurrenceProperties) {
  if (isNullOrUndefined(recurrenceProperties.interval)) {
    throw createNullArgumentError(IntervalKey);
  }

  if (typeof recurrenceProperties.interval !== "number") {
    throw createArgumentTypeError(IntervalKey, typeof recurrenceProperties.interval, "number");
  }

  if (!isNullOrUndefined(recurrenceProperties.dayOfMonth)) {
    if (typeof recurrenceProperties.dayOfMonth !== "number") {
      throw createArgumentTypeError(DayOfMonthKey, typeof recurrenceProperties.dayOfMonth, "number");
    }

    throwOnInvalidDayOfMonth(recurrenceProperties.dayOfMonth);
  } else if (!isNullOrUndefined(recurrenceProperties.dayOfWeek) && !isNullOrUndefined(recurrenceProperties.weekNumber)) {
    if (typeof recurrenceProperties.dayOfWeek !== "string") {
      throw createArgumentTypeError(DayOfWeekKey, typeof recurrenceProperties.dayOfWeek, "string");
    }

    if (!verifyDays(recurrenceProperties.dayOfWeek, true)) {
      throw createArgumentError(DayOfWeekKey);
    }

    if (typeof recurrenceProperties.weekNumber !== "string") {
      throw createArgumentTypeError(WeekNumberKey, typeof recurrenceProperties.weekNumber, "string");
    }

    throwOnInvalidWeekNumber(recurrenceProperties.weekNumber);
  } else {
    throw createArgumentError(undefined, getString("l_Recurrence_Error_Properties_Invalid_Text"));
  }
}

function throwOnInvalidWeekNumber(weekNumber) {
  if (weekNumber !== MailboxEnums.WeekNumber.First && weekNumber !== MailboxEnums.WeekNumber.Second && weekNumber !== MailboxEnums.WeekNumber.Third && weekNumber !== MailboxEnums.WeekNumber.Fourth && weekNumber !== MailboxEnums.WeekNumber.Last) {
    throw createArgumentError(WeekNumberKey);
  }
}

function throwOnInvalidDayOfMonth(iDayOfMonth) {
  if (iDayOfMonth < 1 || iDayOfMonth > 31) {
    throw createArgumentError(DayOfMonthKey);
  }
}

function throwOnInvalidYearlyRecurrence(recurrenceProperties) {
  if (isNullOrUndefined(recurrenceProperties.interval)) {
    throw createNullArgumentError(IntervalKey);
  }

  if (typeof recurrenceProperties.interval !== "number") {
    throw createArgumentTypeError(IntervalKey, typeof recurrenceProperties.interval, "number");
  }

  if (isNullOrUndefined(recurrenceProperties.month)) {
    throw createNullArgumentError(MonthKey);
  }

  if (typeof recurrenceProperties.month !== "string") {
    throw createArgumentTypeError(MonthKey, typeof recurrenceProperties.month, "string");
  }

  throwOnInvalidMonth(recurrenceProperties.month);

  if (!isNullOrUndefined(recurrenceProperties.dayOfMonth)) {
    if (typeof recurrenceProperties.dayOfMonth !== "number") {
      throw createArgumentTypeError(DayOfMonthKey, typeof recurrenceProperties.dayOfMonth, "number");
    }

    throwOnInvalidDayOfMonth(recurrenceProperties.dayOfMonth);
  } else if (!isNullOrUndefined(recurrenceProperties.weekNumber) && !isNullOrUndefined(recurrenceProperties.dayOfWeek)) {
    if (typeof recurrenceProperties.dayOfWeek !== "string") {
      throw createArgumentTypeError(DayOfWeekKey, typeof recurrenceProperties.dayOfWeek, "string");
    }

    if (!verifyDays(recurrenceProperties.dayOfWeek, true)) {
      throw createArgumentError(DayOfWeekKey);
    }

    if (typeof recurrenceProperties.weekNumber !== "string") {
      throw createArgumentTypeError(WeekNumberKey, typeof recurrenceProperties.weekNumber, "string");
    }

    throwOnInvalidWeekNumber(recurrenceProperties.weekNumber);
  } else {
    throw createArgumentError(undefined, getString("l_Recurrence_Error_Properties_Invalid_Text"));
  }
}

function throwOnInvalidMonth(month) {
  if (month !== MailboxEnums.Month.Jan && month !== MailboxEnums.Month.Feb && month !== MailboxEnums.Month.Mar && month !== MailboxEnums.Month.Apr && month !== MailboxEnums.Month.May && month !== MailboxEnums.Month.Jun && month !== MailboxEnums.Month.Jul && month !== MailboxEnums.Month.Aug && month !== MailboxEnums.Month.Sep && month !== MailboxEnums.Month.Oct && month !== MailboxEnums.Month.Nov && month !== MailboxEnums.Month.Dec) {
    throw createArgumentError(MonthKey);
  }
}
// CONCATENATED MODULE: ./src/methods/setRecurrence.ts









function setRecurrence(recurrencePattern) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "recurrenceProperties.setAsync");
  var seriesId = getAppointmentCompose().seriesId;

  if (!isNullOrUndefined(seriesId) && seriesId.length > 0) {
    throw createArgumentError(undefined, getString("l_Recurrence_Error_Instance_SetAsync_Text"));
  }

  validateRecurrenceObject(recurrencePattern);
  var commonParameters = parseCommonArgs(args, false, false);
  var recurrenceData = convertSeriesTime(recurrencePattern);
  var parameters = {
    recurrenceData: recurrenceData
  };
  standardInvokeHostMethod(104, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function convertSeriesTime(recurrencePattern) {
  if (recurrencePattern !== null && recurrencePattern.seriesTime !== null) {
    if (recurrencePattern.seriesTime instanceof SeriesTime_SeriesTime) {
      var recurrencePatternCopy = {
        recurrenceProperties: recurrencePattern.recurrenceProperties,
        recurrenceTimeZone: recurrencePattern.recurrenceTimeZone,
        recurrenceType: recurrencePattern.recurrenceType,
        seriesTimeJson: recurrencePattern.seriesTime.exportToSeriesTimeJson()
      };
      return recurrencePatternCopy;
    }
  }

  return recurrencePattern;
}
// CONCATENATED MODULE: ./src/api/getRecurrenceSurface.ts



function getRecurrenceSurface(isCompose) {
  var recurrence = objectDefine({}, {
    getAsync: getRecurrence
  });

  if (isCompose) {
    objectDefine(recurrence, {
      setAsync: setRecurrence
    });
  }

  return recurrence;
}
// CONCATENATED MODULE: ./src/methods/getAllDayEvent.ts




function getAllDayEvent() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "isAllDayEvent.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  checkFeatureEnabledAndThrow(Features.calendarItems, "isAllDayEvent.getAsync");
  standardInvokeHostMethod(169, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/setAllDayEvent.ts






function setAllDayEvent(isAllDayEvent) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "isAllDayEvent.setAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  var parameters = {
    isAllDayEvent: isAllDayEvent
  };
  checkFeatureEnabledAndThrow(Features.calendarItems, "isAllDayEvent.setAsync");
  setAllDayEvent_validateParameters(parameters);
  standardInvokeHostMethod(170, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function setAllDayEvent_validateParameters(parameters) {
  if (isNullOrUndefined(parameters.isAllDayEvent)) {
    throw createNullArgumentError("isAllDayEvent");
  }

  if (typeof parameters.isAllDayEvent !== "boolean") {
    throw createArgumentTypeError("isAllDayEvent", typeof parameters.isAllDayEvent, "boolean");
  }
}
// CONCATENATED MODULE: ./src/api/getAllDayEventSurface.ts



function getAllDayEventSurface() {
  return objectDefine({}, {
    getAsync: getAllDayEvent,
    setAsync: setAllDayEvent
  });
}
// CONCATENATED MODULE: ./src/methods/setSensitivity.ts







function setSensitivity(sensitivity) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "sensitivity.setAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  var parameters = {
    sensitivity: sensitivity
  };
  checkFeatureEnabledAndThrow(Features.calendarItems, "sensitivity.setAsync");
  setSensitivity_validateParameters(parameters);
  standardInvokeHostMethod(172, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function setSensitivity_validateParameters(parameters) {
  validateStringParam("sensitivity", parameters.sensitivity);
  throwOnInvalidSensitivityType(parameters.sensitivity);
}

function throwOnInvalidSensitivityType(sensitivity) {
  if (sensitivity !== MailboxEnums.AppointmentSensitivityType.Normal && sensitivity !== MailboxEnums.AppointmentSensitivityType.Personal && sensitivity !== MailboxEnums.AppointmentSensitivityType.Private && sensitivity !== MailboxEnums.AppointmentSensitivityType.Confidential) {
    throw createArgumentError("sensitivity");
  }
}
// CONCATENATED MODULE: ./src/methods/getSensitivity.ts




function getSensitivity() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "sensitivity.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  checkFeatureEnabledAndThrow(Features.calendarItems, "sensitivity.getAsync");
  standardInvokeHostMethod(171, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/api/getSensitivitySurface.ts



function getSensitivitySurface() {
  return objectDefine({}, {
    getAsync: getSensitivity,
    setAsync: setSensitivity
  });
}
// CONCATENATED MODULE: ./src/api/getAppointmentCompose.ts






























function getAppointmentCompose() {
  var appointmentCompose = objectDefine({}, {
    body: getBodySurface(true),
    categories: getCategoriesSurface(),
    end: getTimeSurface("end"),
    enhancedLocation: getEnhancedLocationsSurface(true),
    itemType: "appointment",
    location: getLocationSurface(),
    notificationMessages: getNotificationMessageSurface(),
    optionalAttendees: getRecipientsSurface("optionalAttendees"),
    organizer: getFromSurface("organizer"),
    recurrence: getRecurrenceSurface(true),
    requiredAttendees: getRecipientsSurface("requiredAttendees"),
    seriesId: getInitialDataProp("seriesId"),
    start: getTimeSurface("start"),
    subject: getSubjectSurface(),
    addFileAttachmentAsync: addFileAttachment,
    addFileAttachmentFromBase64Async: addBase64FileAttachment,
    addItemAttachmentAsync: addItemAttachment,
    close: close_close,
    getAttachmentsAsync: getAttachments_getAttachments,
    getAttachmentContentAsync: getAttachmentContent,
    getInitializationContextAsync: getInitializationContext,
    getItemIdAsync: getItemId,
    getSelectedDataAsync: getSelectedData,
    loadCustomPropertiesAsync: loadCustomProperties,
    removeAttachmentAsync: removeAttachment,
    saveAsync: save,
    setSelectedDataAsync: setSelectedData(29),
    isAllDayEvent: getAllDayEventSurface(),
    sensitivity: getSensitivitySurface(),
    isClientSignatureEnabledAsync: isClientSignatureEnabled,
    disableClientSignatureAsync: disableClientSignature,
    sessionData: getSessionDataSurface()
  });
  return appointmentCompose;
}
// CONCATENATED MODULE: ./src/utils/addEventSupport.ts
var addEventSupport_OSF = __webpack_require__(0);

var addEventSupport_Microsoft = __webpack_require__(1);

var addEventSupport = function addEventSupport(target) {
  addEventSupport_OSF.DDA.DispIdHost.addEventSupport(target, new addEventSupport_OSF.EventDispatch([addEventSupport_Microsoft.Office.WebExtension.EventType.RecipientsChanged, addEventSupport_Microsoft.Office.WebExtension.EventType.AppointmentTimeChanged, addEventSupport_Microsoft.Office.WebExtension.EventType.AttachmentsChanged, addEventSupport_Microsoft.Office.WebExtension.EventType.EnhancedLocationsChanged, addEventSupport_Microsoft.Office.WebExtension.EventType.InfobarClicked, addEventSupport_Microsoft.Office.WebExtension.EventType.RecurrenceChanged]));
};
// CONCATENATED MODULE: ./src/methods/registerConsent.ts



var ConsentStateType;

(function (ConsentStateType) {
  ConsentStateType[ConsentStateType["NotResponded"] = 0] = "NotResponded";
  ConsentStateType[ConsentStateType["NotConsented"] = 1] = "NotConsented";
  ConsentStateType[ConsentStateType["Consented"] = 2] = "Consented";
})(ConsentStateType || (ConsentStateType = {}));

function registerConsent(consentState) {
  var parameters = {
    consentState: consentState,
    extensionId: getInitialDataProp("extensionId")
  };
  registerConsent_validateParameters(consentState);
  standardInvokeHostMethod(40, undefined, undefined, parameters, undefined);
}

function registerConsent_validateParameters(consentState) {
  if (consentState !== ConsentStateType.Consented && consentState !== ConsentStateType.NotConsented && consentState !== ConsentStateType.NotResponded) {
    throw createArgumentOutOfRange("consentState");
  }
}
// CONCATENATED MODULE: ./src/methods/navigateToModule.ts





function navigateToModule(moduleName) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    module: moduleName
  };
  navigateToModule_validateParameters(moduleName);

  if (moduleName === MailboxEnums.ModuleType.Addins) {
    if (!!commonParameters.options && !!commonParameters.options.queryString) {
      parameters.queryString = commonParameters.options.queryString;
    } else {
      parameters.queryString = "";
    }
  }

  standardInvokeHostMethod(45, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function navigateToModule_validateParameters(moduleName) {
  if (isNullOrUndefined(moduleName)) {
    throw createNullArgumentError("module");
  }

  if (moduleName === "") {
    throw createArgumentError("module");
  }

  if (moduleName !== MailboxEnums.ModuleType.Addins) {
    throw createArgumentError("module");
  }
}
// CONCATENATED MODULE: ./src/methods/recordDataPoint.ts



function recordDataPoint(data) {
  if (isNullOrUndefined(data)) {
    throw createNullArgumentError("data");
  }

  standardInvokeHostMethod(402, undefined, undefined, data, undefined);
}
// CONCATENATED MODULE: ./src/methods/recordTrace.ts



function recordTrace(data) {
  if (isNullOrUndefined(data)) {
    throw createNullArgumentError("data");
  }

  standardInvokeHostMethod(401, undefined, undefined, data, undefined);
}
// CONCATENATED MODULE: ./src/methods/trackCtq.ts



function trackCtq(data) {
  if (isNullOrUndefined(data)) {
    throw createNullArgumentError("data");
  }

  standardInvokeHostMethod(400, undefined, undefined, data, undefined);
}
// CONCATENATED MODULE: ./src/methods/windowOpenOverrideHandler.ts

function windowOpenOverrideHandler(url, target, features, replace) {
  standardInvokeHostMethod(403, undefined, undefined, {
    launchUrl: url
  }, undefined);
  return window;
}
// CONCATENATED MODULE: ./src/methods/logTelemetry.ts



function logTelemetry_logTelemetry(data) {
  if (isNullOrUndefined(data)) {
    throw createNullArgumentError("data");
  }

  standardInvokeHostMethod(163, undefined, undefined, {
    telemetryData: data
  }, undefined);
}
// CONCATENATED MODULE: ./src/methods/logCustomerContentTelemetry.ts



function logCustomerContentTelemetry(data) {
  if (isNullOrUndefined(data)) {
    throw createNullArgumentError("data");
  }

  standardInvokeHostMethod(193, undefined, undefined, {
    telemetryData: data
  }, undefined);
}
// CONCATENATED MODULE: ./src/utils/convertToLocalClientTime.ts
var __assign = undefined && undefined.__assign || function () {
  __assign = Object.assign || function (t) {
    for (var s, i = 1, n = arguments.length; i < n; i++) {
      s = arguments[i];

      for (var p in s) {
        if (Object.prototype.hasOwnProperty.call(s, p)) t[p] = s[p];
      }
    }

    return t;
  };

  return __assign.apply(this, arguments);
};







function convertToLocalClientTime(timeValue) {
  if (!isDateObject(timeValue)) {
    throw createArgumentError("timeValue");
  }

  var date = new Date(timeValue.getTime());
  var offset = date.getTimezoneOffset() * -1;

  if (!isNullOrUndefined(getInitialDataProp("timeZoneOffsets"))) {
    date.setUTCMinutes(date.getUTCMinutes() - offset);
    offset = findOffset(date);
    date.setUTCMinutes(date.getUTCMinutes() + offset);
  }

  var retValue = __assign({
    timezoneOffset: offset
  }, dateToDictionary(date));

  return retValue;
}
// CONCATENATED MODULE: ./src/methods/displayPersonaCardAsync.ts




function displayPersonaCardAsync(ewsIdOrEmail) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    ewsIdOrEmail: ewsIdOrEmail
  };
  displayPersonaCardAsync_validateParameters(parameters);
  standardInvokeHostMethod(43, commonParameters.asyncContext, commonParameters.callback, {
    ewsIdOrEmail: ewsIdOrEmail.trim()
  }, undefined);
}

function displayPersonaCardAsync_validateParameters(parameters) {
  if (!isNullOrUndefined(parameters.ewsIdOrEmail)) {
    displayPersonaCardAsync_throwOnInvalidItemId(parameters.ewsIdOrEmail);

    if (parameters.ewsIdOrEmail === "") {
      throw createArgumentError("ewsIdOrEmail", "ewsIdOrEmail cannot be empty.");
    }
  } else {
    throw createNullArgumentError("ewsIdOrEmail");
  }
}

function displayPersonaCardAsync_throwOnInvalidItemId(ewsIdOrEmail) {
  if (!(typeof ewsIdOrEmail === "string")) {
    throw createArgumentError("ewsIdOrEmail");
  }
}
// CONCATENATED MODULE: ./src/methods/getSharedProperties.ts



function getSharedProperties() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.getSharedPropertiesAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(108, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/utils/addSharedPropertiesSupport.ts





var addSharedPropertiesSupport_addSharedPropertiesSupport = function addSharedPropertiesSupport(target) {
  if (target && getInitialDataProp("isFromSharedFolder") && getHostItemType_getHostItemType() !== HostItemType.ItemLess) {
    objectDefine(target, {
      getSharedPropertiesAsync: getSharedProperties
    });
  }
};
// CONCATENATED MODULE: ./src/api/prepareApiSurface.ts




































var prepareApiSurface_OSF = __webpack_require__(0);

var prepareApiSurface_createMailboxSurface = function createMailboxSurface(target) {
  objectDefine(target, {
    ewsUrl: getInitialDataProp("ewsUrl"),
    restUrl: getInitialDataProp("restUrl"),
    displayAppointmentForm: displayAppointmentForm,
    displayAppointmentFormAsync: displayAppointmentFormAsync,
    displayMessageForm: displayMessageForm,
    displayMessageFormAsync: displayMessageFormAsync,
    displayPersonaCardAsync: displayPersonaCardAsync,
    getCallbackTokenAsync: getCallbackToken,
    getUserIdentityTokenAsync: getUserIdentityToken,
    logTelemetry: logTelemetry_logTelemetry,
    logCustomerContentTelemetry: logCustomerContentTelemetry,
    makeEwsRequestAsync: makeEwsRequest,
    masterCategories: getMasterCategoriesSurface(),
    navigateToModuleAsync: navigateToModule,
    diagnostics: getDiagnosticsSurface(),
    userProfile: getUserProfileSurface(),
    convertToEwsId: convertToEwsId,
    convertToLocalClientTime: convertToLocalClientTime,
    convertToRestId: convertToRestId,
    convertToUtcClientTime: convertToUtcClientTime,
    RegisterConsentAsync: registerConsent,
    GetIsRead: function GetIsRead() {
      return getInitialDataProp("isRead");
    },
    GetEndPointUrl: function GetEndPointUrl() {
      return getInitialDataProp("endNodeUrl");
    },
    GetConsentMetaData: function GetConsentMetaData() {
      return getInitialDataProp("consentMetadata");
    },
    GetMarketplaceContentMarket: function GetMarketplaceContentMarket() {
      return getInitialDataProp("marketplaceContentMarket");
    },
    GetMarketplaceAssetId: function GetMarketplaceAssetId() {
      return getInitialDataProp("marketplaceAssetId");
    },
    GetExtensionId: function GetExtensionId() {
      return getInitialDataProp("extensionId");
    },
    CloseApp: closeApp,
    recordDataPoint: recordDataPoint,
    recordTrace: recordTrace,
    trackCtq: trackCtq
  });

  if (getHostItemType_getHostItemType() !== HostItemType.MessageCompose && getHostItemType_getHostItemType() !== HostItemType.AppointmentCompose) {
    objectDefine(target, {
      displayNewAppointmentForm: displayNewAppointmentForm,
      displayNewMessageForm: displayNewMessageForm,
      displayNewAppointmentFormAsync: displayNewAppointmentFormAsync,
      displayNewMessageFormAsync: displayNewMessageFormAsync
    });
  }

  if (getAppName() === prepareApiSurface_OSF.AppName.OutlookWebApp && getInitialDataProp("openWindowOpen")) {
    window.open = windowOpenOverrideHandler;
  }

  return target;
};
var prepareApiSurface_getItem = function getItem() {
  var item = undefined;

  switch (getHostItemType_getHostItemType()) {
    case HostItemType.Message:
      item = getMessageRead();
      break;

    case HostItemType.MessageCompose:
      item = getMessageCompose();
      break;

    case HostItemType.Appointment:
      item = getAppointmentRead();
      break;

    case HostItemType.AppointmentCompose:
      item = getAppointmentCompose();
      break;

    case HostItemType.MeetingRequest:
      item = getMessageRead();
      break;

    default:
      return undefined;
  }

  if (isOutlookJs()) {
    prepareApiSurface_OSF.OutlookInitializationHelper.addEventDispatchToTarget(item, prepareApiSurface_OSF.OutlookInitializationHelper.getMailboxItemEventDispatch());
  } else {
    addEventSupport(item);
  }

  addSharedPropertiesSupport_addSharedPropertiesSupport(item);
  return item;
};
// CONCATENATED MODULE: ./src/utils/isOutlook16.ts

var isOutlook16_isOutlook16OrGreater = function isOutlook16OrGreater(hostVersion) {
  var endIndex = 0;
  var majorVersionNumber = 0;

  if (!isNullOrUndefined(hostVersion)) {
    endIndex = hostVersion.indexOf(".");
    majorVersionNumber = parseInt(hostVersion.substring(0, endIndex));
  }

  return majorVersionNumber >= 16;
};
// CONCATENATED MODULE: ./src/utils/isApiVersionSupported.ts
var isApiVersionSupported = function isApiVersionSupported(requirementSet, officeAppContext) {
  var apiSupported = false;

  try {
    var requirementDict = JSON.parse(officeAppContext.get_requirementMatrix());
    var hostApiVersion = requirementDict["Mailbox"];
    var hostApiVersionParts = hostApiVersion.split(".");
    var requirementSetParts = requirementSet.split(".");

    if (parseInt(hostApiVersionParts[0]) > parseInt(requirementSetParts[0]) || parseInt(hostApiVersionParts[0]) === parseInt(requirementSetParts[0]) && parseInt(hostApiVersionParts[1]) >= parseInt(requirementSetParts[1])) {
      apiSupported = true;
    }
  } catch (_a) {}

  return apiSupported;
};
// CONCATENATED MODULE: ./src/api/OutlookAppOm.ts








var OutlookAppOm_OSF = __webpack_require__(0);

var appInstance;
var whenStringsFinish;
var getInitialDataProp = function getInitialDataProp(key) {
  return appInstance && appInstance.getInitialDataProp(key);
};
var getIsNoItemContextWebExt = function getIsNoItemContextWebExt() {
  return !appInstance || !appInstance.item;
};
var getAppName = function getAppName() {
  return appInstance && appInstance.getAppName();
};

var OutlookAppOm_OutlookAppOm = function () {
  function OutlookAppOm(appContext, targetWindow, appReadyCallback) {
    var _this = this;

    this.displayName = "mailbox";

    this.stringLoadedCallback = function () {
      if (!!_this.appReadyCallback) {
        if (!_this.officeAppContext.get_isDialog()) {
          standardInvokeHostMethod_invokeHostMethod(1, undefined, _this.onInitialDataResponse);
        } else {
          setTimeout(function () {
            return _this.appReadyCallback();
          }, 0);
        }
      }
    };

    this.initialize = function (data) {
      if (data === null || data === undefined) {
        recreateAdditionalGlobalParametersSingleton(true);
        _this.initialData = null;
        _this.item = null;
      } else {
        _this.initialData = data;
        _this.initialData.permissionLevel = calculatePermissionLevel();
        _this.item = prepareApiSurface_getItem();
        var supportsAdditionalParameters = false;
        supportsAdditionalParameters = getAppName() !== OutlookAppOm_OSF.AppName.Outlook || isOutlook16_isOutlook16OrGreater(getInitialDataProp("hostVersion")) || isApiVersionSupported("1.5", _this.officeAppContext);
        recreateAdditionalGlobalParametersSingleton(supportsAdditionalParameters);

        if (typeof data.itemNumber !== "undefined") {
          getAdditionalGlobalParametersSingleton().setCurrentItemNumber(data.itemNumber);
        }
      }
    };

    this.onInitialDataResponse = function (resultCode, data) {
      if (!!resultCode && resultCode !== InvokeResultCode.noError) {
        return;
      }

      _this.initialize(data);

      prepareApiSurface_createMailboxSurface(_this);
      setTimeout(function () {
        return _this.appReadyCallback();
      }, 0);
    };

    this.officeAppContext = appContext;
    this.targetWindow = window;
    this.appReadyCallback = appReadyCallback;
    appInstance = this;
    loadLocalizedScript(this.stringLoadedCallback);
  }

  OutlookAppOm.prototype.getAppName = function () {
    var retVal = -1;
    retVal = this.officeAppContext.get_appName();
    return retVal;
  };

  OutlookAppOm.prototype.getInitialDataProp = function (key) {
    return this.initialData && this.initialData[key];
  };

  OutlookAppOm.prototype.setCurrentItemNumber = function (newItemNumber) {
    getAdditionalGlobalParametersSingleton().setCurrentItemNumber(newItemNumber);
  };

  OutlookAppOm.addAdditionalArgs = function (dispid, hostCallingArgs) {
    return hostCallingArgs;
  };

  OutlookAppOm.shouldRunInitialDataResponse = function () {
    return true;
  };

  return OutlookAppOm;
}();



var calculatePermissionLevel = function calculatePermissionLevel() {
  var HostReadItem = 1;
  var HostReadWriteMailbox = 2;
  var HostReadWriteItem = 3;
  var permissionLevelFromHost = getInitialDataProp("permissionLevel");

  if (permissionLevelFromHost === undefined) {
    return 0;
  }

  switch (permissionLevelFromHost) {
    case HostReadItem:
      return 1;

    case HostReadWriteItem:
      return 2;

    case HostReadWriteMailbox:
      return 3;

    default:
      return 0;
  }
};
// CONCATENATED MODULE: ./src/methods/saveSettingsRequest.ts





var saveSettingsRequest_OSF = __webpack_require__(0);

var settingsMaxNumberOfCharacters = 32 * 1024;
function saveSettingsRequest(data) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  var commonParameters = parseCommonArgs(args, false, false);
  var serializedSettings = saveSettingsRequest_OSF.DDA.SettingsManager.serializeSettings(data);

  if (JSON.stringify(serializedSettings).length > settingsMaxNumberOfCharacters) {
    var asyncResult_1 = createAsyncResult(undefined, saveSettingsRequest_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 9019, commonParameters.asyncContext, "");

    if (!!commonParameters.callback) {
      setTimeout(function () {
        if (!!commonParameters.callback) commonParameters.callback(asyncResult_1);
      }, 0);
    }

    return;
  }

  if (saveSettingsRequest_OSF.AppName.OutlookWebApp === getAppName()) {
    saveSettingsForOwa(commonParameters, serializedSettings);
  } else {
    saveSettingsForOutlookDesktop(commonParameters, serializedSettings);
  }
}

function saveSettingsForOwa(commonParameters, serializedSettings) {
  standardInvokeHostMethod(404, commonParameters.asyncContext, commonParameters.callback, [serializedSettings], undefined);
}

function saveSettingsForOutlookDesktop(commonParameters, serializedSettings) {
  var detailedErrorCode = -1;
  var storedException = null;

  try {
    var jsonSettings = JSON.stringify(serializedSettings);
    var settingsObjectToSave = {};
    settingsObjectToSave.SettingsKey = jsonSettings;
    saveSettingsRequest_OSF.DDA.ClientSettingsManager.write(settingsObjectToSave);
  } catch (ex) {
    storedException = ex;
  }

  var asyncResult = undefined;

  if (storedException != null) {
    detailedErrorCode = 9019;
    asyncResult = createAsyncResult(undefined, saveSettingsRequest_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, detailedErrorCode, commonParameters.asyncContext, storedException.Message);
  } else {
    detailedErrorCode = 0;
    asyncResult = createAsyncResult(undefined, saveSettingsRequest_OSF.DDA.AsyncResultEnum.ErrorCode.Success, detailedErrorCode, commonParameters.asyncContext);
  }

  if (!!commonParameters.callback) {
    commonParameters.callback(asyncResult);
  }
}
// CONCATENATED MODULE: ./src/api/Settings.ts
var Settings_spreadArrays = undefined && undefined.__spreadArrays || function () {
  for (var s = 0, i = 0, il = arguments.length; i < il; i++) {
    s += arguments[i].length;
  }

  for (var r = Array(s), k = 0, i = 0; i < il; i++) {
    for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++) {
      r[k] = a[j];
    }
  }

  return r;
};




var Settings_OSF = __webpack_require__(0);

var Settings_Settings = function () {
  function Settings(deserializedData) {
    this.rawData = deserializedData;
    this.settingsData = null;
  }

  Settings.prototype.getSettingsData = function () {
    if (this.settingsData == null) {
      this.settingsData = this.convertFromRawSettings(this.rawData);
      this.rawData = null;
    }

    return this.settingsData;
  };

  Settings.prototype.get = function (key) {
    return this.getSettingsData()[key];
  };

  Settings.prototype.set = function (key, value) {
    this.getSettingsData()[key] = value;
  };

  Settings.prototype.remove = function (key) {
    delete this.getSettingsData()[key];
  };

  Settings.prototype.saveAsync = function () {
    var args = [];

    for (var _i = 0; _i < arguments.length; _i++) {
      args[_i] = arguments[_i];
    }

    saveSettingsRequest.apply(void 0, Settings_spreadArrays([this.getSettingsData()], args));
  };

  Settings.prototype.convertFromRawSettings = function (rawSettings) {
    if (rawSettings == null) {
      return {};
    }

    if (getAppName() !== Settings_OSF.AppName.OutlookWebApp) {
      var outlookSettings = rawSettings.SettingsKey;

      if (!!outlookSettings) {
        return Settings_OSF.DDA.SettingsManager.deserializeSettings(outlookSettings);
      }
    }

    return rawSettings;
  };

  return Settings;
}();


// CONCATENATED MODULE: ./src/api/Intellisense.ts



var Intellisense = {
  toItemRead: function toItemRead(item) {
    var hostItemtype = getHostItemType_getHostItemType();

    if (hostItemtype === HostItemType.Message || hostItemtype === HostItemType.Appointment || hostItemtype === HostItemType.MeetingRequest) {
      return item;
    }

    throw createArgumentTypeError();
  },
  toItemCompose: function toItemCompose(item) {
    var hostItemtype = getHostItemType_getHostItemType();

    if (hostItemtype === HostItemType.MessageCompose || hostItemtype === HostItemType.AppointmentCompose) {
      return item;
    }

    throw createArgumentTypeError();
  },
  toMessage: function toMessage(item) {
    return Intellisense.toMessageRead(item);
  },
  toMessageRead: function toMessageRead(item) {
    if (getHostItemType_getHostItemType() === HostItemType.Message || getHostItemType_getHostItemType() === HostItemType.MeetingRequest) {
      return item;
    }

    throw createArgumentTypeError();
  },
  toMessageCompose: function toMessageCompose(item) {
    if (getHostItemType_getHostItemType() === HostItemType.MessageCompose) {
      return item;
    }

    throw createArgumentTypeError();
  },
  toMeetingRequest: function toMeetingRequest(item) {
    if (getHostItemType_getHostItemType() === HostItemType.MeetingRequest) {
      return item;
    }

    throw createArgumentTypeError();
  },
  toAppointment: function toAppointment(item) {
    if (getHostItemType_getHostItemType() === HostItemType.Appointment) {
      return item;
    }

    throw createArgumentTypeError();
  },
  toAppointmentRead: function toAppointmentRead(item) {
    if (getHostItemType_getHostItemType() === HostItemType.Appointment) {
      return item;
    }

    throw createArgumentTypeError();
  },
  toAppointmentCompose: function toAppointmentCompose(item) {
    if (getHostItemType_getHostItemType() === HostItemType.AppointmentCompose) {
      return item;
    }

    throw createArgumentTypeError();
  }
};
// CONCATENATED MODULE: ./src/api/OutlookBase.ts


var OutlookBase = {
  SeriesTimeJsonConverter: function SeriesTimeJsonConverter(rawInput) {
    if (rawInput !== null && typeof rawInput === "object") {
      if (rawInput.seriesTimeJson !== null) {
        var seriesTime = new SeriesTime_SeriesTime();
        seriesTime.importFromSeriesTimeJsonObject(rawInput.seriesTimeJson);
        delete rawInput["seriesTimeJson"];
        rawInput.seriesTime = seriesTime;
      }
    }

    return rawInput;
  },
  CreateAttachmentDetails: function CreateAttachmentDetails(data) {
    convertAttachmentType(data);
    return data;
  }
};
// CONCATENATED MODULE: ./src/index.tsx






OSF = typeof OSF === "object" ? OSF : {};
OSF.DDA = OSF.DDA || {};
OSF.DDA.Settings = Settings_Settings;
OSF = typeof OSF === "object" ? OSF : {};
OSF.DDA = OSF.DDA || {};
OSF.DDA.OutlookAppOm = OutlookAppOm_OutlookAppOm;
Office = typeof Office === "object" ? Office : {};
Office.cast = Office.cast || {};
Office.cast.item = Intellisense;
Microsoft.Office.WebExtension.MailboxEnums = MailboxEnums;
Microsoft.Office.WebExtension.CoercionType = CoercionType;
Microsoft.Office.WebExtension.SeriesTime = SeriesTime_SeriesTime;
Microsoft.Office.WebExtension.OutlookBase = OutlookBase;
/* harmony default export */ var src = __webpack_exports__["default"] = (OutlookAppOm_OutlookAppOm);
var hWindow = window;
hWindow.$h = typeof $h === "object" ? $h : {};
hWindow.$h.Message = $h.Message || {};
hWindow.$h.Appointment = $h.Appointment || {};

hWindow.$h.Message.isInstanceOfType = function (item) {
  return item && item.itemType === "message";
};

hWindow.$h.Appointment.isInstanceOfType = function (item) {
  return item && item.itemType === "appointment";
};

/***/ })
/******/ ])["default"];
//# sourceMappingURL=outlook.web.js.map    
    OSF.DDA.ErrorCodeManager.initializeErrorMessages(Strings.OfficeOM);
    if (appContext.get_appName() == OSF.AppName.Outlook && OSF.DDA.RichApi && OSF.DDA.AsyncMethodNames.ExecuteRichApiRequestAsync)
    {
        OSF.DDA.DispIdHost.addAsyncMethods(OSF.DDA.RichApi, [OSF.DDA.AsyncMethodNames.ExecuteRichApiRequestAsync]);
        OSF.DDA.RichApi.richApiMessageManager = new OfficeExt.RichApiMessageManager();
    }
    if (appContext.get_appName() == OSF.AppName.OutlookWebApp || appContext.get_appName() == OSF.AppName.OutlookIOS || appContext.get_appName() == OSF.AppName.OutlookAndroid)
    {
        this._settings = this._initializeSettings(appContext, false);
    }
    else if (OSF._OfficeAppFactory.getHostInfo()["hostPlatform"] == "mac" && appContext.get_appName() == OSF.AppName.Outlook)
    {
        this._settings = this.initializeMacSettings(appContext, false);
    }
    else
    {
        this._settings = this._initializeSettings(false);
    }
    appContext.appOM = new OSF.DDA.OutlookAppOm(appContext, this._webAppState.wnd, appReady);

    if (appContext.get_appName() == OSF.AppName.Outlook || appContext.get_appName() == OSF.AppName.OutlookWebApp || appContext.get_appName() == OSF.AppName.OutlookIOS || appContext.get_appName() == OSF.AppName.OutlookAndroid)
    {
        // Add OSF's eventing mechanism.
        OSF.DDA.DispIdHost.addEventSupport(
            appContext.appOM,
            new OSF.EventDispatch(
                [
                    Microsoft.Office.WebExtension.EventType.ItemChanged,
                    Microsoft.Office.WebExtension.EventType.OfficeThemeChanged
                ]
            )
        );
    }
};if (typeof OSFPerformance !== "undefined") {
    OSFPerformance.hostInitializationEnd = OSFPerformance.now();
}
