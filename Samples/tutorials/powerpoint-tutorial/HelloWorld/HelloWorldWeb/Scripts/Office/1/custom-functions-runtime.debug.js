/* Office JavaScript API library - Custom Functions */

/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

/*
	This file incorporates the "whatwg-fetch" implementation, version 2.0.3, licensed under MIT with the following licensing notice:
	(See github.com/github/fetch/blob/master/LICENSE)

		Copyright (c) 2014-2016 GitHub, Inc.

		Permission is hereby granted, free of charge, to any person obtaining
		a copy of this software and associated documentation files (the
		"Software"), to deal in the Software without restriction, including
		without limitation the rights to use, copy, modify, merge, publish,
		distribute, sublicense, and/or sell copies of the Software, and to
		permit persons to whom the Software is furnished to do so, subject to
		the following conditions:

		The above copyright notice and this permission notice shall be
		included in all copies or substantial portions of the Software.

		THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
		EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
		MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
		NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
		LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
		OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
		WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

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
var OSF = OSF || {};
OSF.ConstantNames = {
    FileVersion: "0.0.0.0",
    OfficeJS: "custom-functions-runtime.js",
    OfficeDebugJS: "custom-functions-runtime.debug.js",
    HostFileScriptSuffix: "core",
    IsCustomFunctionsRuntime: true
};
var OSF = OSF || {};
OSF.HostSpecificFileVersionDefault = "16.00";
OSF.HostSpecificFileVersionMap = {
    "access": {
        "web": "16.00"
    },
    "agavito": {
        "winrt": "16.00"
    },
    "excel": {
        "ios": "16.00",
        "mac": "16.00",
        "web": "16.00",
        "win32": "16.01",
        "winrt": "16.00"
    },
    "onenote": {
        "android": "16.00",
        "web": "16.00",
        "win32": "16.00",
        "winrt": "16.00"
    },
    "outlook": {
        "ios": "16.00",
        "mac": "16.00",
        "web": "16.01",
        "win32": "16.02"
    },
    "powerpoint": {
        "ios": "16.00",
        "mac": "16.00",
        "web": "16.00",
        "win32": "16.01",
        "winrt": "16.00"
    },
    "project": {
        "win32": "16.00"
    },
    "sway": {
        "web": "16.00"
    },
    "word": {
        "ios": "16.00",
        "mac": "16.00",
        "web": "16.00",
        "win32": "16.01",
        "winrt": "16.00"
    },
    "visio": {
        "web": "16.00",
        "win32": "16.00"
    }
};
OSF.SupportedLocales = {
    "ar-sa": true,
    "bg-bg": true,
    "bn-in": true,
    "ca-es": true,
    "cs-cz": true,
    "da-dk": true,
    "de-de": true,
    "el-gr": true,
    "en-us": true,
    "es-es": true,
    "et-ee": true,
    "eu-es": true,
    "fa-ir": true,
    "fi-fi": true,
    "fr-fr": true,
    "gl-es": true,
    "he-il": true,
    "hi-in": true,
    "hr-hr": true,
    "hu-hu": true,
    "id-id": true,
    "it-it": true,
    "ja-jp": true,
    "kk-kz": true,
    "ko-kr": true,
    "lo-la": true,
    "lt-lt": true,
    "lv-lv": true,
    "ms-my": true,
    "nb-no": true,
    "nl-nl": true,
    "nn-no": true,
    "pl-pl": true,
    "pt-br": true,
    "pt-pt": true,
    "ro-ro": true,
    "ru-ru": true,
    "sk-sk": true,
    "sl-si": true,
    "sr-cyrl-cs": true,
    "sr-cyrl-rs": true,
    "sr-latn-cs": true,
    "sr-latn-rs": true,
    "sv-se": true,
    "th-th": true,
    "tr-tr": true,
    "uk-ua": true,
    "ur-pk": true,
    "vi-vn": true,
    "zh-cn": true,
    "zh-tw": true
};
OSF.AssociatedLocales = {
    ar: "ar-sa",
    bg: "bg-bg",
    bn: "bn-in",
    ca: "ca-es",
    cs: "cs-cz",
    da: "da-dk",
    de: "de-de",
    el: "el-gr",
    en: "en-us",
    es: "es-es",
    et: "et-ee",
    eu: "eu-es",
    fa: "fa-ir",
    fi: "fi-fi",
    fr: "fr-fr",
    gl: "gl-es",
    he: "he-il",
    hi: "hi-in",
    hr: "hr-hr",
    hu: "hu-hu",
    id: "id-id",
    it: "it-it",
    ja: "ja-jp",
    kk: "kk-kz",
    ko: "ko-kr",
    lo: "lo-la",
    lt: "lt-lt",
    lv: "lv-lv",
    ms: "ms-my",
    nb: "nb-no",
    nl: "nl-nl",
    nn: "nn-no",
    pl: "pl-pl",
    pt: "pt-br",
    ro: "ro-ro",
    ru: "ru-ru",
    sk: "sk-sk",
    sl: "sl-si",
    sr: "sr-cyrl-cs",
    sv: "sv-se",
    th: "th-th",
    tr: "tr-tr",
    uk: "uk-ua",
    ur: "ur-pk",
    vi: "vi-vn",
    zh: "zh-cn"
};
OSF.getSupportedLocale = function OSF$getSupportedLocale(locale, defaultLocale) {
    if (defaultLocale === void 0) { defaultLocale = "en-us"; }
    if (!locale) {
        return defaultLocale;
    }
    var supportedLocale;
    locale = locale.toLowerCase();
    if (locale in OSF.SupportedLocales) {
        supportedLocale = locale;
    }
    else {
        var localeParts = locale.split('-', 1);
        if (localeParts && localeParts.length > 0) {
            supportedLocale = OSF.AssociatedLocales[localeParts[0]];
        }
    }
    if (!supportedLocale) {
        supportedLocale = defaultLocale;
    }
    return supportedLocale;
};
var ScriptLoading;
(function (ScriptLoading) {
    var ScriptInfo = (function () {
        function ScriptInfo(url, isReady, hasStarted, timer, pendingCallback) {
            this.url = url;
            this.isReady = isReady;
            this.hasStarted = hasStarted;
            this.timer = timer;
            this.hasError = false;
            this.pendingCallbacks = [];
            this.pendingCallbacks.push(pendingCallback);
        }
        return ScriptInfo;
    }());
    var ScriptTelemetry = (function () {
        function ScriptTelemetry(scriptId, startTime, msResponseTime) {
            this.scriptId = scriptId;
            this.startTime = startTime;
            this.msResponseTime = msResponseTime;
        }
        return ScriptTelemetry;
    }());
    var LoadScriptHelper = (function () {
        function LoadScriptHelper(constantNames) {
            if (constantNames === void 0) { constantNames = {
                OfficeJS: "office.js",
                OfficeDebugJS: "office.debug.js"
            }; }
            this.constantNames = constantNames;
            this.defaultScriptLoadingTimeout = 10000;
            this.loadedScriptByIds = {};
            this.scriptTelemetryBuffer = [];
            this.osfControlAppCorrelationId = "";
            this.basePath = null;
        }
        LoadScriptHelper.prototype.isScriptLoading = function (id) {
            return !!(this.loadedScriptByIds[id] && this.loadedScriptByIds[id].hasStarted);
        };
        LoadScriptHelper.prototype.getOfficeJsBasePath = function () {
            if (this.basePath) {
                return this.basePath;
            }
            else {
                var getScriptBase = function (scriptSrc, scriptNameToCheck) {
                    var scriptBase, indexOfJS, scriptSrcLowerCase;
                    scriptSrcLowerCase = scriptSrc.toLowerCase();
                    indexOfJS = scriptSrcLowerCase.indexOf(scriptNameToCheck);
                    if (indexOfJS >= 0 && indexOfJS === (scriptSrc.length - scriptNameToCheck.length) && (indexOfJS === 0 || scriptSrc.charAt(indexOfJS - 1) === '/' || scriptSrc.charAt(indexOfJS - 1) === '\\')) {
                        scriptBase = scriptSrc.substring(0, indexOfJS);
                    }
                    else if (indexOfJS >= 0
                        && indexOfJS < (scriptSrc.length - scriptNameToCheck.length)
                        && scriptSrc.charAt(indexOfJS + scriptNameToCheck.length) === '?'
                        && (indexOfJS === 0 || scriptSrc.charAt(indexOfJS - 1) === '/' || scriptSrc.charAt(indexOfJS - 1) === '\\')) {
                        scriptBase = scriptSrc.substring(0, indexOfJS);
                    }
                    return scriptBase;
                };
                var scripts = document.getElementsByTagName("script");
                var scriptsCount = scripts.length;
                var officeScripts = [this.constantNames.OfficeJS, this.constantNames.OfficeDebugJS];
                var officeScriptsCount = officeScripts.length;
                var i, j;
                for (i = 0; !this.basePath && i < scriptsCount; i++) {
                    if (scripts[i].src) {
                        for (j = 0; !this.basePath && j < officeScriptsCount; j++) {
                            this.basePath = getScriptBase(scripts[i].src, officeScripts[j]);
                        }
                    }
                }
                return this.basePath;
            }
        };
        LoadScriptHelper.prototype.loadScript = function (url, scriptId, callback, highPriority, timeoutInMs) {
            this.loadScriptInternal(url, scriptId, callback, highPriority, timeoutInMs);
        };
        LoadScriptHelper.prototype.loadScriptParallel = function (url, scriptId, timeoutInMs) {
            this.loadScriptInternal(url, scriptId, null, false, timeoutInMs);
        };
        LoadScriptHelper.prototype.waitForFunction = function (scriptLoadTest, callback, numberOfTries, delay) {
            var attemptsRemaining = numberOfTries;
            var timerId;
            var validateFunction = function () {
                attemptsRemaining--;
                if (scriptLoadTest()) {
                    callback(true);
                    return;
                }
                else if (attemptsRemaining > 0) {
                    timerId = window.setTimeout(validateFunction, delay);
                    attemptsRemaining--;
                }
                else {
                    window.clearTimeout(timerId);
                    callback(false);
                }
            };
            validateFunction();
        };
        LoadScriptHelper.prototype.waitForScripts = function (ids, callback) {
            var _this = this;
            if (this.invokeCallbackIfScriptsReady(ids, callback) == false) {
                for (var i = 0; i < ids.length; i++) {
                    var id = ids[i];
                    var loadedScriptEntry = this.loadedScriptByIds[id];
                    if (loadedScriptEntry) {
                        loadedScriptEntry.pendingCallbacks.push(function () {
                            _this.invokeCallbackIfScriptsReady(ids, callback);
                        });
                    }
                }
            }
        };
        LoadScriptHelper.prototype.logScriptLoading = function (scriptId, startTime, msResponseTime) {
            startTime = Math.floor(startTime);
            if (OSF.AppTelemetry && OSF.AppTelemetry.onScriptDone) {
                if (OSF.AppTelemetry.onScriptDone.length == 3) {
                    OSF.AppTelemetry.onScriptDone(scriptId, startTime, msResponseTime);
                }
                else {
                    OSF.AppTelemetry.onScriptDone(scriptId, startTime, msResponseTime, this.osfControlAppCorrelationId);
                }
            }
            else {
                var scriptTelemetry = new ScriptTelemetry(scriptId, startTime, msResponseTime);
                this.scriptTelemetryBuffer.push(scriptTelemetry);
            }
        };
        LoadScriptHelper.prototype.setAppCorrelationId = function (appCorrelationId) {
            this.osfControlAppCorrelationId = appCorrelationId;
        };
        LoadScriptHelper.prototype.invokeCallbackIfScriptsReady = function (ids, callback) {
            var hasError = false;
            for (var i = 0; i < ids.length; i++) {
                var id = ids[i];
                var loadedScriptEntry = this.loadedScriptByIds[id];
                if (!loadedScriptEntry) {
                    loadedScriptEntry = new ScriptInfo("", false, false, null, null);
                    this.loadedScriptByIds[id] = loadedScriptEntry;
                }
                if (loadedScriptEntry.isReady == false) {
                    return false;
                }
                else if (loadedScriptEntry.hasError) {
                    hasError = true;
                }
            }
            callback(!hasError);
            return true;
        };
        LoadScriptHelper.prototype.getScriptEntryByUrl = function (url) {
            for (var key in this.loadedScriptByIds) {
                var scriptEntry = this.loadedScriptByIds[key];
                if (this.loadedScriptByIds.hasOwnProperty(key) && scriptEntry.url === url) {
                    return scriptEntry;
                }
            }
            return null;
        };
        LoadScriptHelper.prototype.loadScriptInternal = function (url, scriptId, callback, highPriority, timeoutInMs) {
            if (url) {
                var self = this;
                var doc = window.document;
                var loadedScriptEntry = (scriptId && this.loadedScriptByIds[scriptId]) ? this.loadedScriptByIds[scriptId] : this.getScriptEntryByUrl(url);
                if (!loadedScriptEntry || loadedScriptEntry.hasError || loadedScriptEntry.url.toLowerCase() != url.toLowerCase()) {
                    var script = doc.createElement("script");
                    script.type = "text/javascript";
                    if (scriptId) {
                        script.id = scriptId;
                    }
                    if (!loadedScriptEntry) {
                        loadedScriptEntry = new ScriptInfo(url, false, false, null, null);
                        this.loadedScriptByIds[(scriptId ? scriptId : url)] = loadedScriptEntry;
                    }
                    else {
                        loadedScriptEntry.url = url;
                        loadedScriptEntry.hasError = false;
                        loadedScriptEntry.isReady = false;
                    }
                    if (callback) {
                        if (highPriority) {
                            loadedScriptEntry.pendingCallbacks.unshift(callback);
                        }
                        else {
                            loadedScriptEntry.pendingCallbacks.push(callback);
                        }
                    }
                    var timeFromPageInit = -1;
                    if (window.performance && window.performance.now) {
                        timeFromPageInit = window.performance.now();
                    }
                    var startTime = (new Date()).getTime();
                    var logTelemetry = function (succeeded) {
                        if (scriptId) {
                            var totalTime = (new Date()).getTime() - startTime;
                            if (!succeeded) {
                                totalTime = -totalTime;
                            }
                            self.logScriptLoading(scriptId, timeFromPageInit, totalTime);
                        }
                        self.flushTelemetryBuffer();
                    };
                    var onLoadCallback = function OSF_OUtil_loadScript$onLoadCallback() {
                        if (!OSF._OfficeAppFactory.getLoggingAllowed() && (typeof OSF.AppTelemetry !== 'undefined')) {
                            OSF.AppTelemetry.enableTelemetry = false;
                        }
                        logTelemetry(true);
                        loadedScriptEntry.isReady = true;
                        if (loadedScriptEntry.timer != null) {
                            clearTimeout(loadedScriptEntry.timer);
                            delete loadedScriptEntry.timer;
                        }
                        var pendingCallbackCount = loadedScriptEntry.pendingCallbacks.length;
                        for (var i = 0; i < pendingCallbackCount; i++) {
                            var currentCallback = loadedScriptEntry.pendingCallbacks.shift();
                            if (currentCallback) {
                                var result = currentCallback(true);
                                if (result === false) {
                                    break;
                                }
                            }
                        }
                    };
                    var onLoadError = function () {
                        logTelemetry(false);
                        loadedScriptEntry.hasError = true;
                        loadedScriptEntry.isReady = true;
                        if (loadedScriptEntry.timer != null) {
                            clearTimeout(loadedScriptEntry.timer);
                            delete loadedScriptEntry.timer;
                        }
                        var pendingCallbackCount = loadedScriptEntry.pendingCallbacks.length;
                        for (var i = 0; i < pendingCallbackCount; i++) {
                            var currentCallback = loadedScriptEntry.pendingCallbacks.shift();
                            if (currentCallback) {
                                var result = currentCallback(false);
                                if (result === false) {
                                    break;
                                }
                            }
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
                    timeoutInMs = timeoutInMs || this.defaultScriptLoadingTimeout;
                    loadedScriptEntry.timer = setTimeout(onLoadError, timeoutInMs);
                    loadedScriptEntry.hasStarted = true;
                    script.setAttribute("crossOrigin", "anonymous");
                    script.src = url;
                    doc.getElementsByTagName("head")[0].appendChild(script);
                }
                else if (loadedScriptEntry.isReady) {
                    callback(true);
                }
                else {
                    if (highPriority) {
                        loadedScriptEntry.pendingCallbacks.unshift(callback);
                    }
                    else {
                        loadedScriptEntry.pendingCallbacks.push(callback);
                    }
                }
            }
        };
        LoadScriptHelper.prototype.flushTelemetryBuffer = function () {
            if (OSF.AppTelemetry && OSF.AppTelemetry.onScriptDone) {
                for (var i = 0; i < this.scriptTelemetryBuffer.length; i++) {
                    var scriptTelemetry = this.scriptTelemetryBuffer[i];
                    if (OSF.AppTelemetry.onScriptDone.length == 3) {
                        OSF.AppTelemetry.onScriptDone(scriptTelemetry.scriptId, scriptTelemetry.startTime, scriptTelemetry.msResponseTime);
                    }
                    else {
                        OSF.AppTelemetry.onScriptDone(scriptTelemetry.scriptId, scriptTelemetry.startTime, scriptTelemetry.msResponseTime, this.osfControlAppCorrelationId);
                    }
                }
                this.scriptTelemetryBuffer = [];
            }
        };
        return LoadScriptHelper;
    }());
    ScriptLoading.LoadScriptHelper = LoadScriptHelper;
})(ScriptLoading || (ScriptLoading = {}));
var OfficeExt;
(function (OfficeExt) {
    var HostName;
    (function (HostName) {
        var Host = (function () {
            function Host() {
                this.getDiagnostics = function _getDiagnostics(version) {
                    var diagnostics = {
                        host: this.getHost(),
                        version: (version || this.getDefaultVersion()),
                        platform: this.getPlatform()
                    };
                    return diagnostics;
                };
                this.platformRemappings = {
                    web: Microsoft.Office.WebExtension.PlatformType.OfficeOnline,
                    winrt: Microsoft.Office.WebExtension.PlatformType.Universal,
                    win32: Microsoft.Office.WebExtension.PlatformType.PC,
                    mac: Microsoft.Office.WebExtension.PlatformType.Mac,
                    ios: Microsoft.Office.WebExtension.PlatformType.iOS,
                    android: Microsoft.Office.WebExtension.PlatformType.Android
                };
                this.camelCaseMappings = {
                    powerpoint: Microsoft.Office.WebExtension.HostType.PowerPoint,
                    onenote: Microsoft.Office.WebExtension.HostType.OneNote
                };
                this.hostInfo = OSF._OfficeAppFactory.getHostInfo();
                this.getHost = this.getHost.bind(this);
                this.getPlatform = this.getPlatform.bind(this);
                this.getDiagnostics = this.getDiagnostics.bind(this);
            }
            Host.prototype.capitalizeFirstLetter = function (input) {
                if (input) {
                    return (input[0].toUpperCase() + input.slice(1).toLowerCase());
                }
                return input;
            };
            Host.getInstance = function () {
                if (Host.hostObj === undefined) {
                    Host.hostObj = new Host();
                }
                return Host.hostObj;
            };
            Host.prototype.getPlatform = function (appNumber) {
                if (this.hostInfo.hostPlatform) {
                    var hostPlatform = this.hostInfo.hostPlatform.toLowerCase();
                    if (this.platformRemappings[hostPlatform]) {
                        return this.platformRemappings[hostPlatform];
                    }
                }
                return null;
            };
            Host.prototype.getHost = function (appNumber) {
                if (this.hostInfo.hostType) {
                    var hostType = this.hostInfo.hostType.toLowerCase();
                    if (this.camelCaseMappings[hostType]) {
                        return this.camelCaseMappings[hostType];
                    }
                    hostType = this.capitalizeFirstLetter(this.hostInfo.hostType);
                    if (Microsoft.Office.WebExtension.HostType[hostType]) {
                        return Microsoft.Office.WebExtension.HostType[hostType];
                    }
                }
                return null;
            };
            Host.prototype.getDefaultVersion = function () {
                if (this.getHost()) {
                    return "16.0.0000.0000";
                }
                return null;
            };
            return Host;
        }());
        HostName.Host = Host;
    })(HostName = OfficeExt.HostName || (OfficeExt.HostName = {}));
})(OfficeExt || (OfficeExt = {}));
var Office;
(function (Office) {
    var _Internal;
    (function (_Internal) {
        var PromiseImpl;
        (function (PromiseImpl) {
            function Init() {
                return (function () {
                    "use strict";
                    function lib$es6$promise$utils$$objectOrFunction(x) {
                        return typeof x === 'function' || (typeof x === 'object' && x !== null);
                    }
                    function lib$es6$promise$utils$$isFunction(x) {
                        return typeof x === 'function';
                    }
                    function lib$es6$promise$utils$$isMaybeThenable(x) {
                        return typeof x === 'object' && x !== null;
                    }
                    var lib$es6$promise$utils$$_isArray;
                    if (!Array.isArray) {
                        lib$es6$promise$utils$$_isArray = function (x) {
                            return Object.prototype.toString.call(x) === '[object Array]';
                        };
                    }
                    else {
                        lib$es6$promise$utils$$_isArray = Array.isArray;
                    }
                    var lib$es6$promise$utils$$isArray = lib$es6$promise$utils$$_isArray;
                    var lib$es6$promise$asap$$len = 0;
                    var lib$es6$promise$asap$$toString = {}.toString;
                    var lib$es6$promise$asap$$vertxNext;
                    var lib$es6$promise$asap$$customSchedulerFn;
                    var lib$es6$promise$asap$$asap = function asap(callback, arg) {
                        lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len] = callback;
                        lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len + 1] = arg;
                        lib$es6$promise$asap$$len += 2;
                        if (lib$es6$promise$asap$$len === 2) {
                            if (lib$es6$promise$asap$$customSchedulerFn) {
                                lib$es6$promise$asap$$customSchedulerFn(lib$es6$promise$asap$$flush);
                            }
                            else {
                                lib$es6$promise$asap$$scheduleFlush();
                            }
                        }
                    };
                    function lib$es6$promise$asap$$setScheduler(scheduleFn) {
                        lib$es6$promise$asap$$customSchedulerFn = scheduleFn;
                    }
                    function lib$es6$promise$asap$$setAsap(asapFn) {
                        lib$es6$promise$asap$$asap = asapFn;
                    }
                    var lib$es6$promise$asap$$browserWindow = (typeof window !== 'undefined') ? window : undefined;
                    var lib$es6$promise$asap$$browserGlobal = lib$es6$promise$asap$$browserWindow || {};
                    var lib$es6$promise$asap$$BrowserMutationObserver = lib$es6$promise$asap$$browserGlobal.MutationObserver || lib$es6$promise$asap$$browserGlobal.WebKitMutationObserver;
                    var lib$es6$promise$asap$$isNode = typeof process !== 'undefined' && {}.toString.call(process) === '[object process]';
                    var lib$es6$promise$asap$$isWorker = typeof Uint8ClampedArray !== 'undefined' &&
                        typeof importScripts !== 'undefined' &&
                        typeof MessageChannel !== 'undefined';
                    function lib$es6$promise$asap$$useNextTick() {
                        var nextTick = process.nextTick;
                        var version = process.versions.node.match(/^(?:(\d+)\.)?(?:(\d+)\.)?(\*|\d+)$/);
                        if (Array.isArray(version) && version[1] === '0' && version[2] === '10') {
                            nextTick = setImmediate;
                        }
                        return function () {
                            nextTick(lib$es6$promise$asap$$flush);
                        };
                    }
                    function lib$es6$promise$asap$$useVertxTimer() {
                        return function () {
                            lib$es6$promise$asap$$vertxNext(lib$es6$promise$asap$$flush);
                        };
                    }
                    function lib$es6$promise$asap$$useMutationObserver() {
                        var iterations = 0;
                        var observer = new lib$es6$promise$asap$$BrowserMutationObserver(lib$es6$promise$asap$$flush);
                        var node = document.createTextNode('');
                        observer.observe(node, { characterData: true });
                        return function () {
                            node.data = (iterations = ++iterations % 2);
                        };
                    }
                    function lib$es6$promise$asap$$useMessageChannel() {
                        var channel = new MessageChannel();
                        channel.port1.onmessage = lib$es6$promise$asap$$flush;
                        return function () {
                            channel.port2.postMessage(0);
                        };
                    }
                    function lib$es6$promise$asap$$useSetTimeout() {
                        return function () {
                            setTimeout(lib$es6$promise$asap$$flush, 1);
                        };
                    }
                    var lib$es6$promise$asap$$queue = new Array(1000);
                    function lib$es6$promise$asap$$flush() {
                        for (var i = 0; i < lib$es6$promise$asap$$len; i += 2) {
                            var callback = lib$es6$promise$asap$$queue[i];
                            var arg = lib$es6$promise$asap$$queue[i + 1];
                            callback(arg);
                            lib$es6$promise$asap$$queue[i] = undefined;
                            lib$es6$promise$asap$$queue[i + 1] = undefined;
                        }
                        lib$es6$promise$asap$$len = 0;
                    }
                    var lib$es6$promise$asap$$scheduleFlush;
                    if (lib$es6$promise$asap$$isNode) {
                        lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useNextTick();
                    }
                    else if (lib$es6$promise$asap$$isWorker) {
                        lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useMessageChannel();
                    }
                    else {
                        lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useSetTimeout();
                    }
                    function lib$es6$promise$$internal$$noop() { }
                    var lib$es6$promise$$internal$$PENDING = void 0;
                    var lib$es6$promise$$internal$$FULFILLED = 1;
                    var lib$es6$promise$$internal$$REJECTED = 2;
                    var lib$es6$promise$$internal$$GET_THEN_ERROR = new lib$es6$promise$$internal$$ErrorObject();
                    function lib$es6$promise$$internal$$selfFullfillment() {
                        return new TypeError("You cannot resolve a promise with itself");
                    }
                    function lib$es6$promise$$internal$$cannotReturnOwn() {
                        return new TypeError('A promises callback cannot return that same promise.');
                    }
                    function lib$es6$promise$$internal$$getThen(promise) {
                        try {
                            return promise.then;
                        }
                        catch (error) {
                            lib$es6$promise$$internal$$GET_THEN_ERROR.error = error;
                            return lib$es6$promise$$internal$$GET_THEN_ERROR;
                        }
                    }
                    function lib$es6$promise$$internal$$tryThen(then, value, fulfillmentHandler, rejectionHandler) {
                        try {
                            then.call(value, fulfillmentHandler, rejectionHandler);
                        }
                        catch (e) {
                            return e;
                        }
                    }
                    function lib$es6$promise$$internal$$handleForeignThenable(promise, thenable, then) {
                        lib$es6$promise$asap$$asap(function (promise) {
                            var sealed = false;
                            var error = lib$es6$promise$$internal$$tryThen(then, thenable, function (value) {
                                if (sealed) {
                                    return;
                                }
                                sealed = true;
                                if (thenable !== value) {
                                    lib$es6$promise$$internal$$resolve(promise, value);
                                }
                                else {
                                    lib$es6$promise$$internal$$fulfill(promise, value);
                                }
                            }, function (reason) {
                                if (sealed) {
                                    return;
                                }
                                sealed = true;
                                lib$es6$promise$$internal$$reject(promise, reason);
                            }, 'Settle: ' + (promise._label || ' unknown promise'));
                            if (!sealed && error) {
                                sealed = true;
                                lib$es6$promise$$internal$$reject(promise, error);
                            }
                        }, promise);
                    }
                    function lib$es6$promise$$internal$$handleOwnThenable(promise, thenable) {
                        if (thenable._state === lib$es6$promise$$internal$$FULFILLED) {
                            lib$es6$promise$$internal$$fulfill(promise, thenable._result);
                        }
                        else if (thenable._state === lib$es6$promise$$internal$$REJECTED) {
                            lib$es6$promise$$internal$$reject(promise, thenable._result);
                        }
                        else {
                            lib$es6$promise$$internal$$subscribe(thenable, undefined, function (value) {
                                lib$es6$promise$$internal$$resolve(promise, value);
                            }, function (reason) {
                                lib$es6$promise$$internal$$reject(promise, reason);
                            });
                        }
                    }
                    function lib$es6$promise$$internal$$handleMaybeThenable(promise, maybeThenable) {
                        if (maybeThenable.constructor === promise.constructor) {
                            lib$es6$promise$$internal$$handleOwnThenable(promise, maybeThenable);
                        }
                        else {
                            var then = lib$es6$promise$$internal$$getThen(maybeThenable);
                            if (then === lib$es6$promise$$internal$$GET_THEN_ERROR) {
                                lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$GET_THEN_ERROR.error);
                            }
                            else if (then === undefined) {
                                lib$es6$promise$$internal$$fulfill(promise, maybeThenable);
                            }
                            else if (lib$es6$promise$utils$$isFunction(then)) {
                                lib$es6$promise$$internal$$handleForeignThenable(promise, maybeThenable, then);
                            }
                            else {
                                lib$es6$promise$$internal$$fulfill(promise, maybeThenable);
                            }
                        }
                    }
                    function lib$es6$promise$$internal$$resolve(promise, value) {
                        if (promise === value) {
                            lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$selfFullfillment());
                        }
                        else if (lib$es6$promise$utils$$objectOrFunction(value)) {
                            lib$es6$promise$$internal$$handleMaybeThenable(promise, value);
                        }
                        else {
                            lib$es6$promise$$internal$$fulfill(promise, value);
                        }
                    }
                    function lib$es6$promise$$internal$$publishRejection(promise) {
                        if (promise._onerror) {
                            promise._onerror(promise._result);
                        }
                        lib$es6$promise$$internal$$publish(promise);
                    }
                    function lib$es6$promise$$internal$$fulfill(promise, value) {
                        if (promise._state !== lib$es6$promise$$internal$$PENDING) {
                            return;
                        }
                        promise._result = value;
                        promise._state = lib$es6$promise$$internal$$FULFILLED;
                        if (promise._subscribers.length !== 0) {
                            lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, promise);
                        }
                    }
                    function lib$es6$promise$$internal$$reject(promise, reason) {
                        if (promise._state !== lib$es6$promise$$internal$$PENDING) {
                            return;
                        }
                        promise._state = lib$es6$promise$$internal$$REJECTED;
                        promise._result = reason;
                        lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publishRejection, promise);
                    }
                    function lib$es6$promise$$internal$$subscribe(parent, child, onFulfillment, onRejection) {
                        var subscribers = parent._subscribers;
                        var length = subscribers.length;
                        parent._onerror = null;
                        subscribers[length] = child;
                        subscribers[length + lib$es6$promise$$internal$$FULFILLED] = onFulfillment;
                        subscribers[length + lib$es6$promise$$internal$$REJECTED] = onRejection;
                        if (length === 0 && parent._state) {
                            lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, parent);
                        }
                    }
                    function lib$es6$promise$$internal$$publish(promise) {
                        var subscribers = promise._subscribers;
                        var settled = promise._state;
                        if (subscribers.length === 0) {
                            return;
                        }
                        var child, callback, detail = promise._result;
                        for (var i = 0; i < subscribers.length; i += 3) {
                            child = subscribers[i];
                            callback = subscribers[i + settled];
                            if (child) {
                                lib$es6$promise$$internal$$invokeCallback(settled, child, callback, detail);
                            }
                            else {
                                callback(detail);
                            }
                        }
                        promise._subscribers.length = 0;
                    }
                    function lib$es6$promise$$internal$$ErrorObject() {
                        this.error = null;
                    }
                    var lib$es6$promise$$internal$$TRY_CATCH_ERROR = new lib$es6$promise$$internal$$ErrorObject();
                    function lib$es6$promise$$internal$$tryCatch(callback, detail) {
                        try {
                            return callback(detail);
                        }
                        catch (e) {
                            lib$es6$promise$$internal$$TRY_CATCH_ERROR.error = e;
                            return lib$es6$promise$$internal$$TRY_CATCH_ERROR;
                        }
                    }
                    function lib$es6$promise$$internal$$invokeCallback(settled, promise, callback, detail) {
                        var hasCallback = lib$es6$promise$utils$$isFunction(callback), value, error, succeeded, failed;
                        if (hasCallback) {
                            value = lib$es6$promise$$internal$$tryCatch(callback, detail);
                            if (value === lib$es6$promise$$internal$$TRY_CATCH_ERROR) {
                                failed = true;
                                error = value.error;
                                value = null;
                            }
                            else {
                                succeeded = true;
                            }
                            if (promise === value) {
                                lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$cannotReturnOwn());
                                return;
                            }
                        }
                        else {
                            value = detail;
                            succeeded = true;
                        }
                        if (promise._state !== lib$es6$promise$$internal$$PENDING) {
                        }
                        else if (hasCallback && succeeded) {
                            lib$es6$promise$$internal$$resolve(promise, value);
                        }
                        else if (failed) {
                            lib$es6$promise$$internal$$reject(promise, error);
                        }
                        else if (settled === lib$es6$promise$$internal$$FULFILLED) {
                            lib$es6$promise$$internal$$fulfill(promise, value);
                        }
                        else if (settled === lib$es6$promise$$internal$$REJECTED) {
                            lib$es6$promise$$internal$$reject(promise, value);
                        }
                    }
                    function lib$es6$promise$$internal$$initializePromise(promise, resolver) {
                        try {
                            resolver(function resolvePromise(value) {
                                lib$es6$promise$$internal$$resolve(promise, value);
                            }, function rejectPromise(reason) {
                                lib$es6$promise$$internal$$reject(promise, reason);
                            });
                        }
                        catch (e) {
                            lib$es6$promise$$internal$$reject(promise, e);
                        }
                    }
                    function lib$es6$promise$enumerator$$Enumerator(Constructor, input) {
                        var enumerator = this;
                        enumerator._instanceConstructor = Constructor;
                        enumerator.promise = new Constructor(lib$es6$promise$$internal$$noop);
                        if (enumerator._validateInput(input)) {
                            enumerator._input = input;
                            enumerator.length = input.length;
                            enumerator._remaining = input.length;
                            enumerator._init();
                            if (enumerator.length === 0) {
                                lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
                            }
                            else {
                                enumerator.length = enumerator.length || 0;
                                enumerator._enumerate();
                                if (enumerator._remaining === 0) {
                                    lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
                                }
                            }
                        }
                        else {
                            lib$es6$promise$$internal$$reject(enumerator.promise, enumerator._validationError());
                        }
                    }
                    lib$es6$promise$enumerator$$Enumerator.prototype._validateInput = function (input) {
                        return lib$es6$promise$utils$$isArray(input);
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._validationError = function () {
                        return new Error('Array Methods must be provided an Array');
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._init = function () {
                        this._result = new Array(this.length);
                    };
                    var lib$es6$promise$enumerator$$default = lib$es6$promise$enumerator$$Enumerator;
                    lib$es6$promise$enumerator$$Enumerator.prototype._enumerate = function () {
                        var enumerator = this;
                        var length = enumerator.length;
                        var promise = enumerator.promise;
                        var input = enumerator._input;
                        for (var i = 0; promise._state === lib$es6$promise$$internal$$PENDING && i < length; i++) {
                            enumerator._eachEntry(input[i], i);
                        }
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._eachEntry = function (entry, i) {
                        var enumerator = this;
                        var c = enumerator._instanceConstructor;
                        if (lib$es6$promise$utils$$isMaybeThenable(entry)) {
                            if (entry.constructor === c && entry._state !== lib$es6$promise$$internal$$PENDING) {
                                entry._onerror = null;
                                enumerator._settledAt(entry._state, i, entry._result);
                            }
                            else {
                                enumerator._willSettleAt(c.resolve(entry), i);
                            }
                        }
                        else {
                            enumerator._remaining--;
                            enumerator._result[i] = entry;
                        }
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._settledAt = function (state, i, value) {
                        var enumerator = this;
                        var promise = enumerator.promise;
                        if (promise._state === lib$es6$promise$$internal$$PENDING) {
                            enumerator._remaining--;
                            if (state === lib$es6$promise$$internal$$REJECTED) {
                                lib$es6$promise$$internal$$reject(promise, value);
                            }
                            else {
                                enumerator._result[i] = value;
                            }
                        }
                        if (enumerator._remaining === 0) {
                            lib$es6$promise$$internal$$fulfill(promise, enumerator._result);
                        }
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._willSettleAt = function (promise, i) {
                        var enumerator = this;
                        lib$es6$promise$$internal$$subscribe(promise, undefined, function (value) {
                            enumerator._settledAt(lib$es6$promise$$internal$$FULFILLED, i, value);
                        }, function (reason) {
                            enumerator._settledAt(lib$es6$promise$$internal$$REJECTED, i, reason);
                        });
                    };
                    function lib$es6$promise$promise$all$$all(entries) {
                        return new lib$es6$promise$enumerator$$default(this, entries).promise;
                    }
                    var lib$es6$promise$promise$all$$default = lib$es6$promise$promise$all$$all;
                    function lib$es6$promise$promise$race$$race(entries) {
                        var Constructor = this;
                        var promise = new Constructor(lib$es6$promise$$internal$$noop);
                        if (!lib$es6$promise$utils$$isArray(entries)) {
                            lib$es6$promise$$internal$$reject(promise, new TypeError('You must pass an array to race.'));
                            return promise;
                        }
                        var length = entries.length;
                        function onFulfillment(value) {
                            lib$es6$promise$$internal$$resolve(promise, value);
                        }
                        function onRejection(reason) {
                            lib$es6$promise$$internal$$reject(promise, reason);
                        }
                        for (var i = 0; promise._state === lib$es6$promise$$internal$$PENDING && i < length; i++) {
                            lib$es6$promise$$internal$$subscribe(Constructor.resolve(entries[i]), undefined, onFulfillment, onRejection);
                        }
                        return promise;
                    }
                    var lib$es6$promise$promise$race$$default = lib$es6$promise$promise$race$$race;
                    function lib$es6$promise$promise$resolve$$resolve(object) {
                        var Constructor = this;
                        if (object && typeof object === 'object' && object.constructor === Constructor) {
                            return object;
                        }
                        var promise = new Constructor(lib$es6$promise$$internal$$noop);
                        lib$es6$promise$$internal$$resolve(promise, object);
                        return promise;
                    }
                    var lib$es6$promise$promise$resolve$$default = lib$es6$promise$promise$resolve$$resolve;
                    function lib$es6$promise$promise$reject$$reject(reason) {
                        var Constructor = this;
                        var promise = new Constructor(lib$es6$promise$$internal$$noop);
                        lib$es6$promise$$internal$$reject(promise, reason);
                        return promise;
                    }
                    var lib$es6$promise$promise$reject$$default = lib$es6$promise$promise$reject$$reject;
                    var lib$es6$promise$promise$$counter = 0;
                    function lib$es6$promise$promise$$needsResolver() {
                        throw new TypeError('You must pass a resolver function as the first argument to the promise constructor');
                    }
                    function lib$es6$promise$promise$$needsNew() {
                        throw new TypeError("Failed to construct 'Promise': Please use the 'new' operator, this object constructor cannot be called as a function.");
                    }
                    var lib$es6$promise$promise$$default = lib$es6$promise$promise$$Promise;
                    function lib$es6$promise$promise$$Promise(resolver) {
                        this._id = lib$es6$promise$promise$$counter++;
                        this._state = undefined;
                        this._result = undefined;
                        this._subscribers = [];
                        if (lib$es6$promise$$internal$$noop !== resolver) {
                            if (!lib$es6$promise$utils$$isFunction(resolver)) {
                                lib$es6$promise$promise$$needsResolver();
                            }
                            if (!(this instanceof lib$es6$promise$promise$$Promise)) {
                                lib$es6$promise$promise$$needsNew();
                            }
                            lib$es6$promise$$internal$$initializePromise(this, resolver);
                        }
                    }
                    lib$es6$promise$promise$$Promise.all = lib$es6$promise$promise$all$$default;
                    lib$es6$promise$promise$$Promise.race = lib$es6$promise$promise$race$$default;
                    lib$es6$promise$promise$$Promise.resolve = lib$es6$promise$promise$resolve$$default;
                    lib$es6$promise$promise$$Promise.reject = lib$es6$promise$promise$reject$$default;
                    lib$es6$promise$promise$$Promise._setScheduler = lib$es6$promise$asap$$setScheduler;
                    lib$es6$promise$promise$$Promise._setAsap = lib$es6$promise$asap$$setAsap;
                    lib$es6$promise$promise$$Promise._asap = lib$es6$promise$asap$$asap;
                    lib$es6$promise$promise$$Promise.prototype = {
                        constructor: lib$es6$promise$promise$$Promise,
                        then: function (onFulfillment, onRejection) {
                            var parent = this;
                            var state = parent._state;
                            if (state === lib$es6$promise$$internal$$FULFILLED && !onFulfillment || state === lib$es6$promise$$internal$$REJECTED && !onRejection) {
                                return this;
                            }
                            var child = new this.constructor(lib$es6$promise$$internal$$noop);
                            var result = parent._result;
                            if (state) {
                                var callback = arguments[state - 1];
                                lib$es6$promise$asap$$asap(function () {
                                    lib$es6$promise$$internal$$invokeCallback(state, child, callback, result);
                                });
                            }
                            else {
                                lib$es6$promise$$internal$$subscribe(parent, child, onFulfillment, onRejection);
                            }
                            return child;
                        },
                        'catch': function (onRejection) {
                            return this.then(null, onRejection);
                        }
                    };
                    return lib$es6$promise$promise$$default;
                }).call(this);
            }
            PromiseImpl.Init = Init;
        })(PromiseImpl = _Internal.PromiseImpl || (_Internal.PromiseImpl = {}));
    })(_Internal = Office._Internal || (Office._Internal = {}));
    (function (_Internal) {
        function isEdgeLessThan14() {
            var userAgent = window.navigator.userAgent;
            var versionIdx = userAgent.indexOf("Edge/");
            if (versionIdx >= 0) {
                userAgent = userAgent.substring(versionIdx + 5, userAgent.length);
                if (userAgent < "14.14393")
                    return true;
                else
                    return false;
            }
            return false;
        }
        function determinePromise() {
            if (typeof (window) === "undefined" && typeof (Promise) === "function") {
                return Promise;
            }
            if (typeof (window) !== "undefined" && window.Promise) {
                if (isEdgeLessThan14()) {
                    return _Internal.PromiseImpl.Init();
                }
                else {
                    return window.Promise;
                }
            }
            else {
                return _Internal.PromiseImpl.Init();
            }
        }
        _Internal.OfficePromise = determinePromise();
    })(_Internal = Office._Internal || (Office._Internal = {}));
    var OfficePromise = _Internal.OfficePromise;
    Office.Promise = OfficePromise;
})(Office || (Office = {}));
var OTel;
(function (OTel) {
    var CDN_PATH_OTELJS_AGAVE = 'telemetry/oteljs_agave.js';
    var OTelLogger = (function () {
        function OTelLogger() {
        }
        OTelLogger.loaded = function () {
            return !(OTelLogger.logger === undefined);
        };
        OTelLogger.getOtelSinkCDNLocation = function () {
            return (OSF._OfficeAppFactory.getLoadScriptHelper().getOfficeJsBasePath() + CDN_PATH_OTELJS_AGAVE);
        };
        OTelLogger.getMapName = function (map, name) {
            if (name !== undefined && map.hasOwnProperty(name)) {
                return map[name];
            }
            return name;
        };
        OTelLogger.getHost = function () {
            var host = OSF._OfficeAppFactory.getHostInfo()["hostType"];
            var map = {
                "excel": "Excel",
                "onenote": "OneNote",
                "outlook": "Outlook",
                "powerpoint": "PowerPoint",
                "project": "Project",
                "visio": "Visio",
                "word": "Word"
            };
            var mappedName = OTelLogger.getMapName(map, host);
            return mappedName;
        };
        OTelLogger.getFlavor = function () {
            var flavor = OSF._OfficeAppFactory.getHostInfo()["hostPlatform"];
            var map = {
                "android": "Android",
                "ios": "iOS",
                "mac": "Mac",
                "universal": "Universal",
                "web": "Web",
                "win32": "Win32"
            };
            var mappedName = OTelLogger.getMapName(map, flavor);
            return mappedName;
        };
        OTelLogger.ensureValue = function (value, alternative) {
            if (!value) {
                return alternative;
            }
            return value;
        };
        OTelLogger.create = function (info) {
            var contract = {
                id: info.appId,
                assetId: info.assetId,
                officeJsVersion: info.officeJSVersion,
                hostJsVersion: info.hostJSVersion,
                browserToken: info.clientId,
                instanceId: info.appInstanceId,
                sessionId: info.sessionId
            };
            var fields = oteljs.Contracts.Office.System.SDX.getFields("SDX", contract);
            var host = OTelLogger.getHost();
            var flavor = OTelLogger.getFlavor();
            var version = (flavor === "Web" && info.hostVersion.slice(0, 2) === "0.") ? "16.0.0.0" : info.hostVersion;
            var context = {
                'App.Name': host,
                'App.Platform': flavor,
                'App.Version': version,
                'Session.Id': OTelLogger.ensureValue(info.correlationId, "00000000-0000-0000-0000-000000000000")
            };
            var sink = oteljs_agave.AgaveSink.createInstance(context);
            var namespace = "Office.Extensibility.OfficeJs";
            var ariaTenantToken = 'db334b301e7b474db5e0f02f07c51a47-a1b5bc36-1bbe-482f-a64a-c2d9cb606706-7439';
            var nexusTenantToken = 1755;
            var logger = new oteljs.TelemetryLogger(undefined, fields);
            logger.addSink(sink);
            logger.setTenantToken(namespace, ariaTenantToken, nexusTenantToken);
            return logger;
        };
        OTelLogger.initialize = function (info) {
            if (!OTelLogger.Enabled) {
                OTelLogger.promises = [];
                return;
            }
            var timeoutAfterOneSecond = 1000;
            var afterOnReady = function () {
                if ((typeof oteljs === "undefined") || (typeof oteljs_agave === "undefined")) {
                    return;
                }
                if (!OTelLogger.loaded()) {
                    OTelLogger.logger = OTelLogger.create(info);
                }
                if (OTelLogger.loaded()) {
                    OTelLogger.promises.forEach(function (resolve) {
                        resolve();
                    });
                }
            };
            var afterLoadOtelSink = function () {
                Microsoft.Office.WebExtension.onReadyInternal().then(function () { return afterOnReady(); });
            };
            OSF.OUtil.loadScript(OTelLogger.getOtelSinkCDNLocation(), afterLoadOtelSink, timeoutAfterOneSecond);
        };
        OTelLogger.sendTelemetryEvent = function (telemetryEvent) {
            OTelLogger.onTelemetryLoaded(function () {
                try {
                    OTelLogger.logger.sendTelemetryEvent(telemetryEvent);
                }
                catch (e) {
                }
            });
        };
        OTelLogger.onTelemetryLoaded = function (resolve) {
            if (!OTelLogger.Enabled) {
                return;
            }
            if (OTelLogger.loaded()) {
                resolve();
            }
            else {
                OTelLogger.promises.push(resolve);
            }
        };
        OTelLogger.promises = [];
        OTelLogger.Enabled = true;
        return OTelLogger;
    }());
    OTel.OTelLogger = OTelLogger;
})(OTel || (OTel = {}));
(function (OfficeExt) {
    var Association = (function () {
        function Association() {
            this.m_mappings = {};
            this.m_onchangeHandlers = [];
        }
        Association.prototype.associate = function (arg1, arg2) {
            function consoleWarn(message) {
                if (typeof console !== 'undefined' && console.warn) {
                    console.warn(message);
                }
            }
            if (arguments.length == 1 && typeof arguments[0] === 'object' && arguments[0]) {
                var mappings = arguments[0];
                for (var key in mappings) {
                    this.associate(key, mappings[key]);
                }
            }
            else if (arguments.length == 2) {
                var name_1 = arguments[0];
                var func = arguments[1];
                if (typeof name_1 !== 'string') {
                    consoleWarn('[InvalidArg] Function=associate');
                    return;
                }
                if (typeof func !== 'function') {
                    consoleWarn('[InvalidArg] Function=associate');
                    return;
                }
                var nameUpperCase = name_1.toUpperCase();
                if (this.m_mappings[nameUpperCase]) {
                    consoleWarn('[DuplicatedName] Function=' + name_1);
                }
                this.m_mappings[nameUpperCase] = func;
                for (var i = 0; i < this.m_onchangeHandlers.length; i++) {
                    this.m_onchangeHandlers[i]();
                }
            }
            else {
                consoleWarn('[InvalidArg] Function=associate');
            }
        };
        Association.prototype.onchange = function (handler) {
            if (handler) {
                this.m_onchangeHandlers.push(handler);
            }
        };
        Object.defineProperty(Association.prototype, "mappings", {
            get: function () {
                return this.m_mappings;
            },
            enumerable: true,
            configurable: true
        });
        return Association;
    }());
    OfficeExt.Association = Association;
})(OfficeExt || (OfficeExt = {}));
var CustomFunctionMappings = window.CustomFunctionMappings || {};
var CustomFunctions;
(function (CustomFunctions) {
    function delayInitialization() {
        CustomFunctionMappings['__delay__'] = true;
    }
    CustomFunctions.delayInitialization = delayInitialization;
    ;
    CustomFunctions._association = new OfficeExt.Association();
    function associate() {
        CustomFunctions._association.associate.apply(CustomFunctions._association, arguments);
        delete CustomFunctionMappings['__delay__'];
    }
    CustomFunctions.associate = associate;
    ;
})(CustomFunctions || (CustomFunctions = {}));
(function (Office) {
    var actions;
    (function (actions) {
        actions._association = new OfficeExt.Association();
        function associate() {
            actions._association.associate.apply(actions._association, arguments);
        }
        actions.associate = associate;
        ;
    })(actions = Office.actions || (Office.actions = {}));
})(Office || (Office = {}));
var g_isExpEnabled = g_isExpEnabled || false;
var g_isOfflineLibrary = g_isOfflineLibrary || false;
(function () {
    var previousConstantNames = OSF.ConstantNames || {};
    OSF.ConstantNames = {
        FileVersion: "0.0.0.0",
        OfficeJS: "office.js",
        OfficeDebugJS: "office.debug.js",
        DefaultLocale: "en-us",
        LocaleStringLoadingTimeout: 5000,
        MicrosoftAjaxId: "MSAJAX",
        OfficeStringsId: "OFFICESTRINGS",
        OfficeJsId: "OFFICEJS",
        HostFileId: "HOST",
        O15MappingId: "O15Mapping",
        OfficeStringJS: "office_strings.debug.js",
        O15InitHelper: "o15apptofilemappingtable.debug.js",
        SupportedLocales: OSF.SupportedLocales,
        AssociatedLocales: OSF.AssociatedLocales,
        ExperimentScriptSuffix: "experiment"
    };
    for (var key in previousConstantNames) {
        OSF.ConstantNames[key] = previousConstantNames[key];
    }
})();
OSF.InitializationHelper = function OSF_InitializationHelper(hostInfo, webAppState, context, settings, hostFacade) {
    this._hostInfo = hostInfo;
    this._webAppState = webAppState;
    this._context = context;
    this._settings = settings;
    this._hostFacade = hostFacade;
};
OSF.InitializationHelper.prototype.saveAndSetDialogInfo = function OSF_InitializationHelper$saveAndSetDialogInfo(hostInfoValue) {
};
OSF.InitializationHelper.prototype.getAppContext = function OSF_InitializationHelper$getAppContext(wnd, gotAppContext) {
};
OSF.InitializationHelper.prototype.setAgaveHostCommunication = function OSF_InitializationHelper$setAgaveHostCommunication() {
};
OSF.InitializationHelper.prototype.prepareRightBeforeWebExtensionInitialize = function OSF_InitializationHelper$prepareRightBeforeWebExtensionInitialize(appContext) {
};
OSF.InitializationHelper.prototype.loadAppSpecificScriptAndCreateOM = function OSF_InitializationHelper$loadAppSpecificScriptAndCreateOM(appContext, appReady, basePath) {
};
OSF.InitializationHelper.prototype.prepareRightAfterWebExtensionInitialize = function OSF_InitializationHelper$prepareRightAfterWebExtensionInitialize() {
};
OSF.HostInfoFlags = {
    SharedApp: 1,
    CustomFunction: 2,
    ProtectedDocDisable: 4,
    ExperimentJsEnabled: 8,
    PublicAddin: 0x10
};
OSF._OfficeAppFactory = (function OSF__OfficeAppFactory() {
    var _setNamespace = function OSF_OUtil$_setNamespace(name, parent) {
        if (parent && name && !parent[name]) {
            parent[name] = {};
        }
    };
    _setNamespace("Office", window);
    _setNamespace("Microsoft", window);
    _setNamespace("Office", Microsoft);
    _setNamespace("WebExtension", Microsoft.Office);
    if (typeof (window.Office) === 'object') {
        for (var p in window.Office) {
            if (window.Office.hasOwnProperty(p)) {
                Microsoft.Office.WebExtension[p] = window.Office[p];
            }
        }
    }
    window.Office = Microsoft.Office.WebExtension;
    var initialDisplayModeMappings = {
        0: "Unknown",
        1: "Hidden",
        2: "Taskpane",
        3: "Dialog"
    };
    Microsoft.Office.WebExtension.PlatformType = {
        PC: "PC",
        OfficeOnline: "OfficeOnline",
        Mac: "Mac",
        iOS: "iOS",
        Android: "Android",
        Universal: "Universal"
    };
    Microsoft.Office.WebExtension.HostType = {
        Word: "Word",
        Excel: "Excel",
        PowerPoint: "PowerPoint",
        Outlook: "Outlook",
        OneNote: "OneNote",
        Project: "Project",
        Access: "Access",
        Visio: "Visio"
    };
    var _context = {};
    var _settings = {};
    var _hostFacade = {};
    var _WebAppState = { id: null, webAppUrl: null, conversationID: null, clientEndPoint: null, wnd: window.parent, focused: false };
    var _hostInfo = { isO15: true, isRichClient: true, hostType: "", hostPlatform: "", hostSpecificFileVersion: "", hostLocale: "", osfControlAppCorrelationId: "", isDialog: false, disableLogging: false, flags: 0 };
    var _isLoggingAllowed = true;
    var _initializationHelper = {};
    var _appInstanceId = null;
    var _isOfficeJsLoaded = false;
    var _officeOnReadyPendingResolves = [];
    var _isOfficeOnReadyCalled = false;
    var _officeOnReadyHostAndPlatformInfo = { host: null, platform: null, addin: null };
    var _loadScriptHelper = new ScriptLoading.LoadScriptHelper({
        OfficeJS: OSF.ConstantNames.OfficeJS,
        OfficeDebugJS: OSF.ConstantNames.OfficeDebugJS
    });
    if (window.performance && window.performance.now) {
        _loadScriptHelper.logScriptLoading(OSF.ConstantNames.OfficeJsId, -1, window.performance.now());
    }
    var _windowLocationHash = window.location.hash;
    var _windowLocationSearch = window.location.search;
    var _windowName = window.name;
    var setOfficeJsAsLoadedAndDispatchPendingOnReadyCallbacks = function (_a) {
        var host = _a.host, platform = _a.platform, addin = _a.addin;
        _isOfficeJsLoaded = true;
        if (typeof OSFPerformance !== "undefined") {
            OSFPerformance.officeOnReady = OSFPerformance.now();
        }
        _officeOnReadyHostAndPlatformInfo = { host: host, platform: platform, addin: addin };
        while (_officeOnReadyPendingResolves.length > 0) {
            _officeOnReadyPendingResolves.shift()(_officeOnReadyHostAndPlatformInfo);
        }
    };
    Microsoft.Office.WebExtension.FeatureGates = {};
    Microsoft.Office.WebExtension.sendTelemetryEvent = function Microsoft_Office_WebExtension_sendTelemetryEvent(telemetryEvent) {
        OTel.OTelLogger.sendTelemetryEvent(telemetryEvent);
    };
    Microsoft.Office.WebExtension.onReadyInternal = function Microsoft_Office_WebExtension_onReadyInternal(callback) {
        if (_isOfficeJsLoaded) {
            var host_1 = _officeOnReadyHostAndPlatformInfo.host, platform_1 = _officeOnReadyHostAndPlatformInfo.platform, addin_1 = _officeOnReadyHostAndPlatformInfo.addin;
            if (callback) {
                var result = callback({ host: host_1, platform: platform_1, addin: addin_1 });
                if (result && result.then && typeof result.then === "function") {
                    return result.then(function () { return Office.Promise.resolve({ host: host_1, platform: platform_1, addin: addin_1 }); });
                }
            }
            return Office.Promise.resolve({ host: host_1, platform: platform_1, addin: addin_1 });
        }
        if (callback) {
            return new Office.Promise(function (resolve) {
                _officeOnReadyPendingResolves.push(function (receivedHostAndPlatform) {
                    var result = callback(receivedHostAndPlatform);
                    if (result && result.then && typeof result.then === "function") {
                        return result.then(function () { return resolve(receivedHostAndPlatform); });
                    }
                    resolve(receivedHostAndPlatform);
                });
            });
        }
        return new Office.Promise(function (resolve) {
            _officeOnReadyPendingResolves.push(resolve);
        });
    };
    Microsoft.Office.WebExtension.onReady = function Microsoft_Office_WebExtension_onReady(callback) {
        _isOfficeOnReadyCalled = true;
        return Microsoft.Office.WebExtension.onReadyInternal(callback);
    };
    var getQueryStringValue = function OSF__OfficeAppFactory$getQueryStringValue(paramName) {
        var hostInfoValue;
        var searchString = window.location.search;
        if (searchString) {
            var hostInfoParts = searchString.split(paramName + "=");
            if (hostInfoParts.length > 1) {
                var hostInfoValueRestString = hostInfoParts[1];
                var separatorRegex = new RegExp("[&#]", "g");
                var hostInfoValueParts = hostInfoValueRestString.split(separatorRegex);
                if (hostInfoValueParts.length > 0) {
                    hostInfoValue = hostInfoValueParts[0];
                }
            }
        }
        return hostInfoValue;
    };
    var compareVersions = function _compareVersions(version1, version2) {
        var splitVersion1 = version1.split(".");
        var splitVersion2 = version2.split(".");
        var iter;
        for (iter in splitVersion1) {
            if (parseInt(splitVersion1[iter]) < parseInt(splitVersion2[iter])) {
                return false;
            }
            else if (parseInt(splitVersion1[iter]) > parseInt(splitVersion2[iter])) {
                return true;
            }
        }
        return false;
    };
    var shouldLoadOldOutlookMacJs = function _shouldLoadOldOutlookMacJs() {
        try {
            var versionToUseNewJS = "15.30.1128.0";
            var currentHostVersion = window.external.GetContext().GetHostFullVersion();
        }
        catch (ex) {
            return false;
        }
        return !!compareVersions(versionToUseNewJS, currentHostVersion);
    };
    var _retrieveLoggingAllowed = function OSF__OfficeAppFactory$_retrieveLoggingAllowed() {
        _isLoggingAllowed = true;
        try {
            if (_hostInfo.disableLogging) {
                _isLoggingAllowed = false;
                return;
            }
            window.external = window.external || {};
            if (typeof window.external.GetLoggingAllowed === 'undefined') {
                _isLoggingAllowed = true;
            }
            else {
                _isLoggingAllowed = window.external.GetLoggingAllowed();
            }
        }
        catch (Exception) {
        }
    };
    var _retrieveHostInfo = function OSF__OfficeAppFactory$_retrieveHostInfo() {
        var hostInfoParaName = "_host_Info";
        var hostInfoValue = getQueryStringValue(hostInfoParaName);
        if (!hostInfoValue) {
            try {
                var windowNameObj = JSON.parse(_windowName);
                hostInfoValue = windowNameObj ? windowNameObj["hostInfo"] : null;
            }
            catch (Exception) {
            }
        }
        if (!hostInfoValue) {
            try {
                window.external = window.external || {};
                if (typeof agaveHost !== "undefined" && agaveHost.GetHostInfo) {
                    window.external.GetHostInfo = function () {
                        return agaveHost.GetHostInfo();
                    };
                }
                var fallbackHostInfo = window.external.GetHostInfo();
                if (fallbackHostInfo == "isDialog") {
                    _hostInfo.isO15 = true;
                    _hostInfo.isDialog = true;
                }
                else if (fallbackHostInfo.toLowerCase().indexOf("mac") !== -1 && fallbackHostInfo.toLowerCase().indexOf("outlook") !== -1 && shouldLoadOldOutlookMacJs()) {
                    _hostInfo.isO15 = true;
                }
                else {
                    var hostInfoParts = fallbackHostInfo.split(hostInfoParaName + "=");
                    if (hostInfoParts.length > 1) {
                        hostInfoValue = hostInfoParts[1];
                    }
                    else {
                        hostInfoValue = fallbackHostInfo;
                    }
                }
            }
            catch (Exception) {
            }
        }
        var getSessionStorage = function OSF__OfficeAppFactory$_retrieveHostInfo$getSessionStorage() {
            var osfSessionStorage = null;
            try {
                if (window.sessionStorage) {
                    osfSessionStorage = window.sessionStorage;
                }
            }
            catch (ex) {
            }
            return osfSessionStorage;
        };
        var osfSessionStorage = getSessionStorage();
        if (!hostInfoValue && osfSessionStorage && osfSessionStorage.getItem("hostInfoValue")) {
            hostInfoValue = osfSessionStorage.getItem("hostInfoValue");
        }
        if (hostInfoValue) {
            hostInfoValue = decodeURIComponent(hostInfoValue);
            _hostInfo.isO15 = false;
            var items = hostInfoValue.split("$");
            if (typeof items[2] == "undefined") {
                items = hostInfoValue.split("|");
            }
            _hostInfo.hostType = (typeof items[0] == "undefined") ? "" : items[0].toLowerCase();
            _hostInfo.hostPlatform = (typeof items[1] == "undefined") ? "" : items[1].toLowerCase();
            ;
            _hostInfo.hostSpecificFileVersion = (typeof items[2] == "undefined") ? "" : items[2].toLowerCase();
            _hostInfo.hostLocale = (typeof items[3] == "undefined") ? "" : items[3].toLowerCase();
            _hostInfo.osfControlAppCorrelationId = (typeof items[4] == "undefined") ? "" : items[4];
            if (_hostInfo.osfControlAppCorrelationId == "telemetry") {
                _hostInfo.osfControlAppCorrelationId = "";
            }
            _hostInfo.isDialog = (((typeof items[5]) != "undefined") && items[5] == "isDialog") ? true : false;
            _hostInfo.disableLogging = (((typeof items[6]) != "undefined") && items[6] == "disableLogging") ? true : false;
            _hostInfo.flags = (((typeof items[7]) === "string") && items[7].length > 0) ? parseInt(items[7]) : 0;
            if (g_isOfflineLibrary) {
                g_isExpEnabled = false;
            }
            else {
                g_isExpEnabled = g_isExpEnabled || !!(_hostInfo.flags & OSF.HostInfoFlags.ExperimentJsEnabled);
            }
            var hostSpecificFileVersionValue = parseFloat(_hostInfo.hostSpecificFileVersion);
            var fallbackVersion = OSF.HostSpecificFileVersionDefault;
            if (OSF.HostSpecificFileVersionMap[_hostInfo.hostType] && OSF.HostSpecificFileVersionMap[_hostInfo.hostType][_hostInfo.hostPlatform]) {
                fallbackVersion = OSF.HostSpecificFileVersionMap[_hostInfo.hostType][_hostInfo.hostPlatform];
            }
            if (hostSpecificFileVersionValue > parseFloat(fallbackVersion)) {
                _hostInfo.hostSpecificFileVersion = fallbackVersion;
            }
            if (osfSessionStorage) {
                try {
                    osfSessionStorage.setItem("hostInfoValue", hostInfoValue);
                }
                catch (e) {
                }
            }
        }
        else {
            _hostInfo.isO15 = true;
            _hostInfo.hostLocale = getQueryStringValue("locale");
        }
    };
    var getAppContextAsync = function OSF__OfficeAppFactory$getAppContextAsync(wnd, gotAppContext) {
        if (OSF.AppTelemetry && OSF.AppTelemetry.logAppCommonMessage) {
            OSF.AppTelemetry.logAppCommonMessage("getAppContextAsync starts");
        }
        _initializationHelper.getAppContext(wnd, gotAppContext);
    };
    var initialize = function OSF__OfficeAppFactory$initialize() {
        _retrieveHostInfo();
        _retrieveLoggingAllowed();
        if (_hostInfo.hostPlatform == "web" && _hostInfo.isDialog && window == window.top && window.opener == null) {
            window.open('', '_self', '');
            window.close();
        }
        if ((_hostInfo.flags & (OSF.HostInfoFlags.SharedApp | OSF.HostInfoFlags.CustomFunction)) !== 0) {
            if (typeof (window.Promise) === 'undefined') {
                window.Promise = window.Office.Promise;
            }
        }
        _loadScriptHelper.setAppCorrelationId(_hostInfo.osfControlAppCorrelationId);
        var basePath = _loadScriptHelper.getOfficeJsBasePath();
        var requiresMsAjax = false;
        if (!basePath)
            throw "Office Web Extension script library file name should be " + OSF.ConstantNames.OfficeJS + " or " + OSF.ConstantNames.OfficeDebugJS + ".";
        var isMicrosftAjaxLoaded = function OSF$isMicrosftAjaxLoaded() {
            if ((typeof (Sys) !== 'undefined' && typeof (Type) !== 'undefined' &&
                Sys.StringBuilder && typeof (Sys.StringBuilder) === "function" &&
                Type.registerNamespace && typeof (Type.registerNamespace) === "function" &&
                Type.registerClass && typeof (Type.registerClass) === "function") ||
                (typeof (OfficeExt) !== "undefined" && OfficeExt.MsAjaxError)) {
                return true;
            }
            else {
                return false;
            }
        };
        var officeStrings = null;
        var loadLocaleStrings = function OSF__OfficeAppFactory_initialize$loadLocaleStrings(appLocale) {
            var fallbackLocaleTried = false;
            var loadLocaleStringCallback = function OSF__OfficeAppFactory_initialize$loadLocaleStringCallback() {
                if (typeof Strings == 'undefined' || typeof Strings.OfficeOM == 'undefined') {
                    if (!fallbackLocaleTried) {
                        fallbackLocaleTried = true;
                        var fallbackLocaleStringFile = basePath + OSF.ConstantNames.DefaultLocale + "/" + OSF.ConstantNames.OfficeStringJS;
                        _loadScriptHelper.loadScript(fallbackLocaleStringFile, OSF.ConstantNames.OfficeStringsId, loadLocaleStringCallback, true, OSF.ConstantNames.LocaleStringLoadingTimeout);
                        return false;
                    }
                    else {
                        throw "Neither the locale, " + appLocale.toLowerCase() + ", provided by the host app nor the fallback locale " + OSF.ConstantNames.DefaultLocale + " are supported.";
                    }
                }
                else {
                    fallbackLocaleTried = false;
                    officeStrings = Strings.OfficeOM;
                }
            };
            if (!isMicrosftAjaxLoaded()) {
                window.Type = Function;
                Type.registerNamespace = function (ns) {
                    window[ns] = window[ns] || {};
                };
                Type.prototype.registerClass = function (cls) {
                    cls = {};
                };
            }
            var localeStringFile = basePath + OSF.getSupportedLocale(appLocale, OSF.ConstantNames.DefaultLocale) + "/" + OSF.ConstantNames.OfficeStringJS;
            _loadScriptHelper.loadScript(localeStringFile, OSF.ConstantNames.OfficeStringsId, loadLocaleStringCallback, true, OSF.ConstantNames.LocaleStringLoadingTimeout);
        };
        var onAppCodeAndMSAjaxReady = function OSF__OfficeAppFactory_initialize$onAppCodeAndMSAjaxReady(loadSuccess) {
            if (loadSuccess) {
                _initializationHelper = new OSF.InitializationHelper(_hostInfo, _WebAppState, _context, _settings, _hostFacade);
                if (_hostInfo.hostPlatform == "web" && _initializationHelper.saveAndSetDialogInfo) {
                    _initializationHelper.saveAndSetDialogInfo(getQueryStringValue("_host_Info"));
                }
                _initializationHelper.setAgaveHostCommunication();
                if (typeof OSFPerformance !== "undefined") {
                    OSFPerformance.getAppContextStart = OSFPerformance.now();
                }
                getAppContextAsync(_WebAppState.wnd, function (appContext) {
                    if (typeof OSFPerformance !== "undefined") {
                        OSFPerformance.getAppContextEnd = OSFPerformance.now();
                    }
                    if (OSF.AppTelemetry && OSF.AppTelemetry.logAppCommonMessage) {
                        OSF.AppTelemetry.logAppCommonMessage("getAppContextAsync callback start");
                    }
                    _appInstanceId = appContext._appInstanceId;
                    if (appContext.get_featureGates) {
                        var featureGates = appContext.get_featureGates();
                        if (featureGates) {
                            Microsoft.Office.WebExtension.FeatureGates = featureGates;
                        }
                    }
                    var updateVersionInfo = function updateVersionInfo() {
                        var hostVersionItems = _hostInfo.hostSpecificFileVersion.split(".");
                        if (appContext.get_appMinorVersion) {
                            var isIOS = _hostInfo.hostPlatform == "ios";
                            if (!isIOS) {
                                if (isNaN(appContext.get_appMinorVersion())) {
                                    appContext._appMinorVersion = parseInt(hostVersionItems[1]);
                                }
                                else if (hostVersionItems.length > 1 && !isNaN(Number(hostVersionItems[1]))) {
                                    appContext._appMinorVersion = parseInt(hostVersionItems[1]);
                                }
                            }
                        }
                        if (_hostInfo.isDialog) {
                            appContext._isDialog = _hostInfo.isDialog;
                        }
                    };
                    updateVersionInfo();
                    var appReady = function appReady() {
                        _initializationHelper.prepareApiSurface && _initializationHelper.prepareApiSurface(appContext);
                        _loadScriptHelper.waitForFunction(function () { return (Microsoft.Office.WebExtension.initialize != undefined || _isOfficeOnReadyCalled); }, function (initializedDeclaredOrOfficeOnReadyCalled) {
                            if (initializedDeclaredOrOfficeOnReadyCalled) {
                                if (_initializationHelper.prepareApiSurface) {
                                    if (Microsoft.Office.WebExtension.initialize) {
                                        Microsoft.Office.WebExtension.initialize(_initializationHelper.getInitializationReason(appContext));
                                    }
                                }
                                else {
                                    if (!Microsoft.Office.WebExtension.initialize) {
                                        Microsoft.Office.WebExtension.initialize = function () { };
                                    }
                                    _initializationHelper.prepareRightBeforeWebExtensionInitialize(appContext);
                                }
                                _initializationHelper.prepareRightAfterWebExtensionInitialize && _initializationHelper.prepareRightAfterWebExtensionInitialize();
                                var appNumber = appContext.get_appName();
                                var addinInfo = null;
                                if ((_hostInfo.flags & OSF.HostInfoFlags.SharedApp) !== 0) {
                                    addinInfo = {
                                        visibilityMode: initialDisplayModeMappings[(appContext.get_initialDisplayMode && typeof appContext.get_initialDisplayMode === 'function') ? appContext.get_initialDisplayMode() : 0]
                                    };
                                }
                                setOfficeJsAsLoadedAndDispatchPendingOnReadyCallbacks({
                                    host: OfficeExt.HostName.Host.getInstance().getHost(appNumber),
                                    platform: OfficeExt.HostName.Host.getInstance().getPlatform(appNumber),
                                    addin: addinInfo
                                });
                            }
                            else {
                                throw new Error("Office.js has not fully loaded. Your app must call \"Office.onReady()\" as part of it's loading sequence (or set the \"Office.initialize\" function). If your app has this functionality, try reloading this page.");
                            }
                        }, 400, 50);
                    };
                    if (!_loadScriptHelper.isScriptLoading(OSF.ConstantNames.OfficeStringsId)) {
                        loadLocaleStrings(appContext.get_appUILocale());
                    }
                    _loadScriptHelper.waitForScripts([OSF.ConstantNames.OfficeStringsId], function () {
                        if (officeStrings && !Strings.OfficeOM) {
                            Strings.OfficeOM = officeStrings;
                        }
                        _initializationHelper.loadAppSpecificScriptAndCreateOM(appContext, appReady, basePath);
                        if (typeof OSFPerformance !== "undefined") {
                            OSFPerformance.createOMEnd = OSFPerformance.now();
                        }
                    });
                });
                if (_hostInfo.isO15) {
                    var wacXdmInfoIsMissing = (OSF.OUtil.parseXdmInfo() == null);
                    if (wacXdmInfoIsMissing) {
                        var isPlainBrowser = true;
                        if (window.external && typeof window.external.GetContext !== 'undefined') {
                            try {
                                window.external.GetContext();
                                isPlainBrowser = false;
                            }
                            catch (e) {
                            }
                        }
                        if (typeof OsfOptOut === "undefined" && isPlainBrowser && window.top !== window.self) {
                            if (window.console && window.console.log) {
                                window.console.log("The add-in is not hosted in plain browser top window.");
                            }
                            window.location.href = "about:blank";
                        }
                        if (isPlainBrowser) {
                            setOfficeJsAsLoadedAndDispatchPendingOnReadyCallbacks({
                                host: null,
                                platform: null,
                                addin: null
                            });
                        }
                    }
                }
            }
            else {
                var errorMsg = "MicrosoftAjax.js is not loaded successfully.";
                if (OSF.AppTelemetry && OSF.AppTelemetry.logAppException) {
                    OSF.AppTelemetry.logAppException(errorMsg);
                }
                throw errorMsg;
            }
        };
        var onAppCodeReady = function OSF__OfficeAppFactory_initialize$onAppCodeReady() {
            if (OSF.AppTelemetry && OSF.AppTelemetry.setOsfControlAppCorrelationId) {
                OSF.AppTelemetry.setOsfControlAppCorrelationId(_hostInfo.osfControlAppCorrelationId);
            }
            if (_loadScriptHelper.isScriptLoading(OSF.ConstantNames.MicrosoftAjaxId)) {
                _loadScriptHelper.waitForScripts([OSF.ConstantNames.MicrosoftAjaxId], onAppCodeAndMSAjaxReady);
            }
            else {
                _loadScriptHelper.waitForFunction(isMicrosftAjaxLoaded, onAppCodeAndMSAjaxReady, 500, 100);
            }
        };
        if (_hostInfo.isO15) {
            _loadScriptHelper.loadScript(basePath + OSF.ConstantNames.O15InitHelper, OSF.ConstantNames.O15MappingId, onAppCodeReady);
        }
        else {
            var hostSpecificFileName;
            if (g_isExpEnabled) {
                hostSpecificFileName = ([
                    _hostInfo.hostType,
                    _hostInfo.hostPlatform,
                    OSF.ConstantNames.ExperimentScriptSuffix || null,
                    OSF.ConstantNames.HostFileScriptSuffix || null,
                ]
                    .filter(function (part) { return part != null; })
                    .join("-"))
                    +
                        ".debug.js";
            }
            else {
                hostSpecificFileName = ([
                    _hostInfo.hostType,
                    _hostInfo.hostPlatform,
                    _hostInfo.hostSpecificFileVersion,
                    OSF.ConstantNames.HostFileScriptSuffix || null,
                ]
                    .filter(function (part) { return part != null; })
                    .join("-"))
                    +
                        ".debug.js";
            }
            _loadScriptHelper.loadScript(basePath + hostSpecificFileName.toLowerCase(), OSF.ConstantNames.HostFileId, onAppCodeReady);
            if (typeof OSFPerformance !== "undefined") {
                OSFPerformance.hostSpecificFileName = hostSpecificFileName;
            }
        }
        if (_hostInfo.hostLocale) {
            loadLocaleStrings(_hostInfo.hostLocale);
        }
        if (requiresMsAjax && !isMicrosftAjaxLoaded()) {
            var msAjaxCDNPath = (window.location.protocol.toLowerCase() === 'https:' ? 'https:' : 'http:') + '//ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js';
            _loadScriptHelper.loadScriptParallel(msAjaxCDNPath, OSF.ConstantNames.MicrosoftAjaxId);
        }
        window.confirm = function OSF__OfficeAppFactory_initialize$confirm(message) {
            throw new Error('Function window.confirm is not supported.');
        };
        window.alert = function OSF__OfficeAppFactory_initialize$alert(message) {
            throw new Error('Function window.alert is not supported.');
        };
        window.prompt = function OSF__OfficeAppFactory_initialize$prompt(message, defaultvalue) {
            throw new Error('Function window.prompt is not supported.');
        };
        var isOutlookAndroid = _hostInfo.hostType == "outlook" && _hostInfo.hostPlatform == "android";
        if (!isOutlookAndroid) {
            window.history.replaceState = null;
            window.history.pushState = null;
        }
    };
    initialize();
    if (window.addEventListener) {
        window.addEventListener('DOMContentLoaded', function (event) {
            Microsoft.Office.WebExtension.onReadyInternal(function () {
                if (typeof OSFPerfUtil !== 'undefined') {
                    OSFPerfUtil.sendPerformanceTelemetry();
                }
            });
        });
    }
    return {
        getId: function OSF__OfficeAppFactory$getId() { return _WebAppState.id; },
        getClientEndPoint: function OSF__OfficeAppFactory$getClientEndPoint() { return _WebAppState.clientEndPoint; },
        getContext: function OSF__OfficeAppFactory$getContext() { return _context; },
        setContext: function OSF__OfficeAppFactory$setContext(context) { _context = context; },
        getHostInfo: function OSF_OfficeAppFactory$getHostInfo() { return _hostInfo; },
        getLoggingAllowed: function OSF_OfficeAppFactory$getLoggingAllowed() { return _isLoggingAllowed; },
        getHostFacade: function OSF__OfficeAppFactory$getHostFacade() { return _hostFacade; },
        setHostFacade: function setHostFacade(hostFacade) { _hostFacade = hostFacade; },
        getInitializationHelper: function OSF__OfficeAppFactory$getInitializationHelper() { return _initializationHelper; },
        getCachedSessionSettingsKey: function OSF__OfficeAppFactory$getCachedSessionSettingsKey() {
            return (_WebAppState.conversationID != null ? _WebAppState.conversationID : _appInstanceId) + "CachedSessionSettings";
        },
        getWebAppState: function OSF__OfficeAppFactory$getWebAppState() { return _WebAppState; },
        getWindowLocationHash: function OSF__OfficeAppFactory$getHash() { return _windowLocationHash; },
        getWindowLocationSearch: function OSF__OfficeAppFactory$getSearch() { return _windowLocationSearch; },
        getLoadScriptHelper: function OSF__OfficeAppFactory$getLoadScriptHelper() { return _loadScriptHelper; },
        getWindowName: function OSF__OfficeAppFactory$getWindowName() { return _windowName; }
    };
})();



!function(modules) {
    var installedModules = {};
    function __webpack_require__(moduleId) {
        if (installedModules[moduleId]) return installedModules[moduleId].exports;
        var module = installedModules[moduleId] = {
            i: moduleId,
            l: !1,
            exports: {}
        };
        return modules[moduleId].call(module.exports, module, module.exports, __webpack_require__), 
        module.l = !0, module.exports;
    }
    __webpack_require__.m = modules, __webpack_require__.c = installedModules, __webpack_require__.d = function(exports, name, getter) {
        __webpack_require__.o(exports, name) || Object.defineProperty(exports, name, {
            enumerable: !0,
            get: getter
        });
    }, __webpack_require__.r = function(exports) {
        "undefined" != typeof Symbol && Symbol.toStringTag && Object.defineProperty(exports, Symbol.toStringTag, {
            value: "Module"
        }), Object.defineProperty(exports, "__esModule", {
            value: !0
        });
    }, __webpack_require__.t = function(value, mode) {
        if (1 & mode && (value = __webpack_require__(value)), 8 & mode) return value;
        if (4 & mode && "object" == typeof value && value && value.__esModule) return value;
        var ns = Object.create(null);
        if (__webpack_require__.r(ns), Object.defineProperty(ns, "default", {
            enumerable: !0,
            value: value
        }), 2 & mode && "string" != typeof value) for (var key in value) __webpack_require__.d(ns, key, function(key) {
            return value[key];
        }.bind(null, key));
        return ns;
    }, __webpack_require__.n = function(module) {
        var getter = module && module.__esModule ? function() {
            return module.default;
        } : function() {
            return module;
        };
        return __webpack_require__.d(getter, "a", getter), getter;
    }, __webpack_require__.o = function(object, property) {
        return Object.prototype.hasOwnProperty.call(object, property);
    }, __webpack_require__.p = "", __webpack_require__(__webpack_require__.s = 3);
}([ function(module, exports, __webpack_require__) {
    "use strict";
    var extendStatics, __extends = this && this.__extends || (extendStatics = function(d, b) {
        return (extendStatics = Object.setPrototypeOf || {
            __proto__: []
        } instanceof Array && function(d, b) {
            d.__proto__ = b;
        } || function(d, b) {
            for (var p in b) b.hasOwnProperty(p) && (d[p] = b[p]);
        })(d, b);
    }, function(d, b) {
        function __() {
            this.constructor = d;
        }
        extendStatics(d, b), d.prototype = null === b ? Object.create(b) : (__.prototype = b.prototype, 
        new __);
    });
    Object.defineProperty(exports, "__esModule", {
        value: !0
    }), exports.TestUtility = exports.CoreUtility = exports.RichApiMessageUtility = exports.CoreConstants = exports.CoreResourceStrings = exports.CoreErrorCodes = exports.Error = exports._Internal = exports.HostBridge = exports.HttpUtility = exports.SessionBase = void 0;
    var SessionBase = function() {
        function SessionBase() {}
        return SessionBase.prototype._resolveRequestUrlAndHeaderInfo = function() {
            return CoreUtility._createPromiseFromResult(null);
        }, SessionBase.prototype._createRequestExecutorOrNull = function() {
            return null;
        }, Object.defineProperty(SessionBase.prototype, "eventRegistration", {
            get: function() {
                return null;
            },
            enumerable: !1,
            configurable: !0
        }), SessionBase;
    }();
    exports.SessionBase = SessionBase;
    var HttpUtility = function() {
        function HttpUtility() {}
        return HttpUtility.setCustomSendRequestFunc = function(func) {
            HttpUtility.s_customSendRequestFunc = func;
        }, HttpUtility.xhrSendRequestFunc = function(request) {
            return CoreUtility.createPromise((function(resolve, reject) {
                var xhr = new XMLHttpRequest;
                if (xhr.open(request.method, request.url), xhr.onload = function() {
                    var resp = {
                        statusCode: xhr.status,
                        headers: CoreUtility._parseHttpResponseHeaders(xhr.getAllResponseHeaders()),
                        body: xhr.responseText
                    };
                    resolve(resp);
                }, xhr.onerror = function() {
                    reject(new _Internal.RuntimeError({
                        code: CoreErrorCodes.connectionFailure,
                        httpStatusCode: xhr.status,
                        message: CoreUtility._getResourceString(CoreResourceStrings.connectionFailureWithStatus, xhr.statusText)
                    }));
                }, request.headers) for (var key in request.headers) xhr.setRequestHeader(key, request.headers[key]);
                xhr.send(CoreUtility._getRequestBodyText(request));
            }));
        }, HttpUtility.fetchSendRequestFunc = function(request) {
            var requestBodyText = CoreUtility._getRequestBodyText(request);
            return "" === requestBodyText && (requestBodyText = void 0), fetch(request.url, {
                method: request.method,
                headers: request.headers,
                body: requestBodyText
            }).then((function(resp) {
                return resp.text().then((function(body) {
                    var statusCode = resp.status, headers = {};
                    return resp.headers.forEach((function(value, name) {
                        headers[name] = value;
                    })), {
                        statusCode: statusCode,
                        headers: headers,
                        body: body
                    };
                }));
            }));
        }, HttpUtility.sendRequest = function(request) {
            HttpUtility.validateAndNormalizeRequest(request);
            var func = HttpUtility.s_customSendRequestFunc;
            return func || (func = "undefined" != typeof fetch ? HttpUtility.fetchSendRequestFunc : HttpUtility.xhrSendRequestFunc), 
            func(request);
        }, HttpUtility.setCustomSendLocalDocumentRequestFunc = function(func) {
            HttpUtility.s_customSendLocalDocumentRequestFunc = func;
        }, HttpUtility.sendLocalDocumentRequest = function(request) {
            return HttpUtility.validateAndNormalizeRequest(request), (HttpUtility.s_customSendLocalDocumentRequestFunc || HttpUtility.officeJsSendLocalDocumentRequestFunc)(request);
        }, HttpUtility.officeJsSendLocalDocumentRequestFunc = function(request) {
            request = CoreUtility._validateLocalDocumentRequest(request);
            var requestSafeArray = CoreUtility._buildRequestMessageSafeArray(request);
            return CoreUtility.createPromise((function(resolve, reject) {
                OSF.DDA.RichApi.executeRichApiRequestAsync(requestSafeArray, (function(asyncResult) {
                    var response;
                    response = "succeeded" == asyncResult.status ? {
                        statusCode: RichApiMessageUtility.getResponseStatusCode(asyncResult),
                        headers: RichApiMessageUtility.getResponseHeaders(asyncResult),
                        body: RichApiMessageUtility.getResponseBody(asyncResult)
                    } : RichApiMessageUtility.buildHttpResponseFromOfficeJsError(asyncResult.error.code, asyncResult.error.message), 
                    CoreUtility.log("Response:"), CoreUtility.log(JSON.stringify(response)), resolve(response);
                }));
            }));
        }, HttpUtility.validateAndNormalizeRequest = function(request) {
            if (CoreUtility.isNullOrUndefined(request)) throw _Internal.RuntimeError._createInvalidArgError({
                argumentName: "request"
            });
            CoreUtility.isNullOrEmptyString(request.method) && (request.method = "GET"), request.method = request.method.toUpperCase();
        }, HttpUtility.logRequest = function(request) {
            if (CoreUtility._logEnabled) {
                if (CoreUtility.log("---HTTP Request---"), CoreUtility.log(request.method + " " + request.url), 
                request.headers) for (var key in request.headers) CoreUtility.log(key + ": " + request.headers[key]);
                HttpUtility._logBodyEnabled && CoreUtility.log(CoreUtility._getRequestBodyText(request));
            }
        }, HttpUtility.logResponse = function(response) {
            if (CoreUtility._logEnabled) {
                if (CoreUtility.log("---HTTP Response---"), CoreUtility.log("" + response.statusCode), 
                response.headers) for (var key in response.headers) CoreUtility.log(key + ": " + response.headers[key]);
                HttpUtility._logBodyEnabled && CoreUtility.log(response.body);
            }
        }, HttpUtility._logBodyEnabled = !1, HttpUtility;
    }();
    exports.HttpUtility = HttpUtility;
    var _Internal, HostBridge = function() {
        function HostBridge(m_bridge) {
            var _this = this;
            this.m_bridge = m_bridge, this.m_promiseResolver = {}, this.m_handlers = [], this.m_bridge.onMessageFromHost = function(messageText) {
                var message = JSON.parse(messageText);
                if (3 == message.type) {
                    var genericMessageBody = message.message;
                    if (genericMessageBody && genericMessageBody.entries) for (var i = 0; i < genericMessageBody.entries.length; i++) {
                        var entryObjectOrArray = genericMessageBody.entries[i];
                        if (Array.isArray(entryObjectOrArray)) {
                            var entry = {
                                messageCategory: entryObjectOrArray[0],
                                messageType: entryObjectOrArray[1],
                                targetId: entryObjectOrArray[2],
                                message: entryObjectOrArray[3],
                                id: entryObjectOrArray[4]
                            };
                            genericMessageBody.entries[i] = entry;
                        }
                    }
                }
                _this.dispatchMessage(message);
            };
        }
        return HostBridge.init = function(bridge) {
            if ("object" == typeof bridge && bridge) {
                var instance = new HostBridge(bridge);
                HostBridge.s_instance = instance, HttpUtility.setCustomSendLocalDocumentRequestFunc((function(request) {
                    request = CoreUtility._validateLocalDocumentRequest(request);
                    var requestFlags = 0;
                    CoreUtility.isReadonlyRestRequest(request.method) || (requestFlags = 1);
                    var index = request.url.indexOf("?");
                    if (index >= 0) {
                        var query = request.url.substr(index + 1), flagsAndCustomData = CoreUtility._parseRequestFlagsAndCustomDataFromQueryStringIfAny(query);
                        flagsAndCustomData.flags >= 0 && (requestFlags = flagsAndCustomData.flags);
                    }
                    var bridgeMessage = {
                        id: HostBridge.nextId(),
                        type: 1,
                        flags: requestFlags,
                        message: request
                    };
                    return instance.sendMessageToHostAndExpectResponse(bridgeMessage).then((function(bridgeResponse) {
                        return bridgeResponse.message;
                    }));
                }));
                for (var i = 0; i < HostBridge.s_onInitedHandlers.length; i++) HostBridge.s_onInitedHandlers[i](instance);
            }
        }, Object.defineProperty(HostBridge, "instance", {
            get: function() {
                return HostBridge.s_instance;
            },
            enumerable: !1,
            configurable: !0
        }), HostBridge.prototype.sendMessageToHost = function(message) {
            this.m_bridge.sendMessageToHost(JSON.stringify(message));
        }, HostBridge.prototype.sendMessageToHostAndExpectResponse = function(message) {
            var _this = this, ret = CoreUtility.createPromise((function(resolve, reject) {
                _this.m_promiseResolver[message.id] = resolve;
            }));
            return this.m_bridge.sendMessageToHost(JSON.stringify(message)), ret;
        }, HostBridge.prototype.addHostMessageHandler = function(handler) {
            this.m_handlers.push(handler);
        }, HostBridge.prototype.removeHostMessageHandler = function(handler) {
            var index = this.m_handlers.indexOf(handler);
            index >= 0 && this.m_handlers.splice(index, 1);
        }, HostBridge.onInited = function(handler) {
            HostBridge.s_onInitedHandlers.push(handler), HostBridge.s_instance && handler(HostBridge.s_instance);
        }, HostBridge.prototype.dispatchMessage = function(message) {
            if ("number" == typeof message.id) {
                var resolve = this.m_promiseResolver[message.id];
                if (resolve) return resolve(message), void delete this.m_promiseResolver[message.id];
            }
            for (var i = 0; i < this.m_handlers.length; i++) this.m_handlers[i](message);
        }, HostBridge.nextId = function() {
            return HostBridge.s_nextId++;
        }, HostBridge.s_onInitedHandlers = [], HostBridge.s_nextId = 1, HostBridge;
    }();
    exports.HostBridge = HostBridge, "object" == typeof _richApiNativeBridge && _richApiNativeBridge && HostBridge.init(_richApiNativeBridge), 
    function(_Internal) {
        var RuntimeError = function(_super) {
            function RuntimeError(error) {
                var _this = _super.call(this, "string" == typeof error ? error : error.message) || this;
                if (Object.setPrototypeOf(_this, RuntimeError.prototype), _this.name = "RichApi.Error", 
                "string" == typeof error ? _this.message = error : (_this.code = error.code, _this.message = error.message, 
                _this.traceMessages = error.traceMessages || [], _this.innerError = error.innerError || null, 
                _this.debugInfo = _this._createDebugInfo(error.debugInfo || {}), _this.httpStatusCode = error.httpStatusCode, 
                _this.data = error.data), CoreUtility.isNullOrUndefined(_this.httpStatusCode) || 200 === _this.httpStatusCode) {
                    var mapping = {};
                    mapping[CoreErrorCodes.accessDenied] = 401, mapping[CoreErrorCodes.connectionFailure] = 500, 
                    mapping[CoreErrorCodes.generalException] = 500, mapping[CoreErrorCodes.invalidArgument] = 400, 
                    mapping[CoreErrorCodes.invalidObjectPath] = 400, mapping[CoreErrorCodes.invalidOrTimedOutSession] = 408, 
                    mapping[CoreErrorCodes.invalidRequestContext] = 400, mapping[CoreErrorCodes.timeout] = 408, 
                    mapping[CoreErrorCodes.valueNotLoaded] = 400, _this.httpStatusCode = mapping[_this.code];
                }
                return CoreUtility.isNullOrUndefined(_this.httpStatusCode) && (_this.httpStatusCode = 500), 
                _this;
            }
            return __extends(RuntimeError, _super), RuntimeError.prototype.toString = function() {
                return this.code + ": " + this.message;
            }, RuntimeError.prototype._createDebugInfo = function(partialDebugInfo) {
                var debugInfo = {
                    code: this.code,
                    message: this.message,
                    toString: function() {
                        return JSON.stringify(this);
                    }
                };
                for (var key in partialDebugInfo) debugInfo[key] = partialDebugInfo[key];
                return this.innerError && (this.innerError instanceof _Internal.RuntimeError ? debugInfo.innerError = this.innerError.debugInfo : debugInfo.innerError = this.innerError), 
                debugInfo;
            }, RuntimeError._createInvalidArgError = function(error) {
                return new _Internal.RuntimeError({
                    code: CoreErrorCodes.invalidArgument,
                    httpStatusCode: 400,
                    message: CoreUtility.isNullOrEmptyString(error.argumentName) ? CoreUtility._getResourceString(CoreResourceStrings.invalidArgumentGeneric) : CoreUtility._getResourceString(CoreResourceStrings.invalidArgument, error.argumentName),
                    debugInfo: error.errorLocation ? {
                        errorLocation: error.errorLocation
                    } : {},
                    innerError: error.innerError
                });
            }, RuntimeError;
        }(Error);
        _Internal.RuntimeError = RuntimeError;
    }(_Internal = exports._Internal || (exports._Internal = {})), exports.Error = _Internal.RuntimeError;
    var CoreErrorCodes = function() {
        function CoreErrorCodes() {}
        return CoreErrorCodes.apiNotFound = "ApiNotFound", CoreErrorCodes.accessDenied = "AccessDenied", 
        CoreErrorCodes.generalException = "GeneralException", CoreErrorCodes.activityLimitReached = "ActivityLimitReached", 
        CoreErrorCodes.invalidArgument = "InvalidArgument", CoreErrorCodes.connectionFailure = "ConnectionFailure", 
        CoreErrorCodes.timeout = "Timeout", CoreErrorCodes.invalidOrTimedOutSession = "InvalidOrTimedOutSession", 
        CoreErrorCodes.invalidObjectPath = "InvalidObjectPath", CoreErrorCodes.invalidRequestContext = "InvalidRequestContext", 
        CoreErrorCodes.valueNotLoaded = "ValueNotLoaded", CoreErrorCodes.requestPayloadSizeLimitExceeded = "RequestPayloadSizeLimitExceeded", 
        CoreErrorCodes;
    }();
    exports.CoreErrorCodes = CoreErrorCodes;
    var CoreResourceStrings = function() {
        function CoreResourceStrings() {}
        return CoreResourceStrings.apiNotFoundDetails = "ApiNotFoundDetails", CoreResourceStrings.connectionFailureWithStatus = "ConnectionFailureWithStatus", 
        CoreResourceStrings.connectionFailureWithDetails = "ConnectionFailureWithDetails", 
        CoreResourceStrings.invalidArgument = "InvalidArgument", CoreResourceStrings.invalidArgumentGeneric = "InvalidArgumentGeneric", 
        CoreResourceStrings.timeout = "Timeout", CoreResourceStrings.invalidOrTimedOutSessionMessage = "InvalidOrTimedOutSessionMessage", 
        CoreResourceStrings.invalidObjectPath = "InvalidObjectPath", CoreResourceStrings.invalidRequestContext = "InvalidRequestContext", 
        CoreResourceStrings.valueNotLoaded = "ValueNotLoaded", CoreResourceStrings;
    }();
    exports.CoreResourceStrings = CoreResourceStrings;
    var CoreConstants = function() {
        function CoreConstants() {}
        return CoreConstants.flags = "flags", CoreConstants.sourceLibHeader = "SdkVersion", 
        CoreConstants.processQuery = "ProcessQuery", CoreConstants.localDocument = "http://document.localhost/", 
        CoreConstants.localDocumentApiPrefix = "http://document.localhost/_api/", CoreConstants.customData = "customdata", 
        CoreConstants;
    }();
    exports.CoreConstants = CoreConstants;
    var RichApiMessageUtility = function() {
        function RichApiMessageUtility() {}
        return RichApiMessageUtility.buildMessageArrayForIRequestExecutor = function(customData, requestFlags, requestMessage, sourceLibHeaderValue) {
            var requestMessageText = JSON.stringify(requestMessage.Body);
            CoreUtility.log("Request:"), CoreUtility.log(requestMessageText);
            var headers = {};
            return CoreUtility._copyHeaders(requestMessage.Headers, headers), headers[CoreConstants.sourceLibHeader] = sourceLibHeaderValue, 
            RichApiMessageUtility.buildRequestMessageSafeArray(customData, requestFlags, "POST", CoreConstants.processQuery, headers, requestMessageText);
        }, RichApiMessageUtility.buildResponseOnSuccess = function(responseBody, responseHeaders) {
            var response = {
                HttpStatusCode: 200,
                ErrorCode: "",
                ErrorMessage: "",
                Headers: null,
                Body: null
            };
            return response.Body = JSON.parse(responseBody), response.Headers = responseHeaders, 
            response;
        }, RichApiMessageUtility.buildResponseOnError = function(errorCode, message) {
            var response = {
                HttpStatusCode: 500,
                ErrorCode: "",
                ErrorMessage: "",
                Headers: null,
                Body: null
            };
            return response.ErrorCode = CoreErrorCodes.generalException, response.ErrorMessage = message, 
            errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability ? (response.ErrorCode = CoreErrorCodes.accessDenied, 
            response.HttpStatusCode = 401) : errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached ? (response.ErrorCode = CoreErrorCodes.activityLimitReached, 
            response.HttpStatusCode = 429) : errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidOrTimedOutSession ? (response.ErrorCode = CoreErrorCodes.invalidOrTimedOutSession, 
            response.HttpStatusCode = 408, response.ErrorMessage = CoreUtility._getResourceString(CoreResourceStrings.invalidOrTimedOutSessionMessage)) : errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeRequestPayloadSizeLimitExceeded && (response.ErrorCode = CoreErrorCodes.requestPayloadSizeLimitExceeded, 
            response.HttpStatusCode = 400), response;
        }, RichApiMessageUtility.buildHttpResponseFromOfficeJsError = function(errorCode, message) {
            var statusCode = 500, errorBody = {
                error: {}
            };
            return errorBody.error.code = CoreErrorCodes.generalException, errorBody.error.message = message, 
            errorCode === RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability ? (statusCode = 403, 
            errorBody.error.code = CoreErrorCodes.accessDenied) : errorCode === RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached && (statusCode = 429, 
            errorBody.error.code = CoreErrorCodes.activityLimitReached), {
                statusCode: statusCode,
                headers: {},
                body: JSON.stringify(errorBody)
            };
        }, RichApiMessageUtility.buildRequestMessageSafeArray = function(customData, requestFlags, method, path, headers, body) {
            var headerArray = [];
            if (headers) for (var headerName in headers) headerArray.push(headerName), headerArray.push(headers[headerName]);
            return [ customData, method, path, headerArray, body, 0, requestFlags, "", "", "" ];
        }, RichApiMessageUtility.getResponseBody = function(result) {
            return RichApiMessageUtility.getResponseBodyFromSafeArray(result.value.data);
        }, RichApiMessageUtility.getResponseHeaders = function(result) {
            return RichApiMessageUtility.getResponseHeadersFromSafeArray(result.value.data);
        }, RichApiMessageUtility.getResponseBodyFromSafeArray = function(data) {
            var ret = data[2];
            return "string" == typeof ret ? ret : ret.join("");
        }, RichApiMessageUtility.getResponseHeadersFromSafeArray = function(data) {
            var arrayHeader = data[1];
            if (!arrayHeader) return null;
            for (var headers = {}, i = 0; i < arrayHeader.length - 1; i += 2) headers[arrayHeader[i]] = arrayHeader[i + 1];
            return headers;
        }, RichApiMessageUtility.getResponseStatusCode = function(result) {
            return RichApiMessageUtility.getResponseStatusCodeFromSafeArray(result.value.data);
        }, RichApiMessageUtility.getResponseStatusCodeFromSafeArray = function(data) {
            return data[0];
        }, RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidOrTimedOutSession = 5012, RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached = 5102, 
        RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability = 7e3, RichApiMessageUtility.OfficeJsErrorCode_ooeRequestPayloadSizeLimitExceeded = 5103, 
        RichApiMessageUtility;
    }();
    exports.RichApiMessageUtility = RichApiMessageUtility, function(_Internal) {
        _Internal.getPromiseType = function() {
            if ("undefined" != typeof Promise) return Promise;
            if ("undefined" != typeof Office && Office.Promise) return Office.Promise;
            if ("undefined" != typeof OfficeExtension && OfficeExtension.Promise) return OfficeExtension.Promise;
            throw new _Internal.Error("No Promise implementation found");
        };
    }(_Internal = exports._Internal || (exports._Internal = {}));
    var CoreUtility = function() {
        function CoreUtility() {}
        return CoreUtility.log = function(message) {
            CoreUtility._logEnabled && "undefined" != typeof console && console.log && console.log(message);
        }, CoreUtility.checkArgumentNull = function(value, name) {
            if (CoreUtility.isNullOrUndefined(value)) throw _Internal.RuntimeError._createInvalidArgError({
                argumentName: name
            });
        }, CoreUtility.isNullOrUndefined = function(value) {
            return null === value || void 0 === value;
        }, CoreUtility.isUndefined = function(value) {
            return void 0 === value;
        }, CoreUtility.isNullOrEmptyString = function(value) {
            return null === value || (void 0 === value || 0 == value.length);
        }, CoreUtility.isPlainJsonObject = function(value) {
            if (CoreUtility.isNullOrUndefined(value)) return !1;
            if ("object" != typeof value) return !1;
            if ("[object Object]" !== Object.prototype.toString.apply(value)) return !1;
            if (value.constructor && !Object.prototype.hasOwnProperty.call(value, "constructor") && !Object.prototype.hasOwnProperty.call(value.constructor.prototype, "hasOwnProperty")) return !1;
            for (var key in value) if (!Object.prototype.hasOwnProperty.call(value, key)) return !1;
            return !0;
        }, CoreUtility.trim = function(str) {
            return str.replace(new RegExp("^\\s+|\\s+$", "g"), "");
        }, CoreUtility.caseInsensitiveCompareString = function(str1, str2) {
            return CoreUtility.isNullOrUndefined(str1) ? CoreUtility.isNullOrUndefined(str2) : !CoreUtility.isNullOrUndefined(str2) && str1.toUpperCase() == str2.toUpperCase();
        }, CoreUtility.isReadonlyRestRequest = function(method) {
            return CoreUtility.caseInsensitiveCompareString(method, "GET");
        }, CoreUtility._getResourceString = function(resourceId, arg) {
            var ret;
            if ("undefined" != typeof window && window.Strings && window.Strings.OfficeOM) {
                var stringName = "L_" + resourceId, stringValue = window.Strings.OfficeOM[stringName];
                stringValue && (ret = stringValue);
            }
            if (ret || (ret = CoreUtility.s_resourceStringValues[resourceId]), ret || (ret = resourceId), 
            !CoreUtility.isNullOrUndefined(arg)) if (Array.isArray(arg)) {
                var arrArg = arg;
                ret = CoreUtility._formatString(ret, arrArg);
            } else ret = ret.replace("{0}", arg);
            return ret;
        }, CoreUtility._formatString = function(format, arrArg) {
            return format.replace(/\{\d\}/g, (function(v) {
                var position = parseInt(v.substr(1, v.length - 2));
                if (position < arrArg.length) return arrArg[position];
                throw _Internal.RuntimeError._createInvalidArgError({
                    argumentName: "format"
                });
            }));
        }, Object.defineProperty(CoreUtility, "Promise", {
            get: function() {
                return _Internal.getPromiseType();
            },
            enumerable: !1,
            configurable: !0
        }), CoreUtility.createPromise = function(executor) {
            return new CoreUtility.Promise(executor);
        }, CoreUtility._createPromiseFromResult = function(value) {
            return CoreUtility.createPromise((function(resolve, reject) {
                resolve(value);
            }));
        }, CoreUtility._createPromiseFromException = function(reason) {
            return CoreUtility.createPromise((function(resolve, reject) {
                reject(reason);
            }));
        }, CoreUtility._createTimeoutPromise = function(timeout) {
            return CoreUtility.createPromise((function(resolve, reject) {
                setTimeout((function() {
                    resolve(null);
                }), timeout);
            }));
        }, CoreUtility._createInvalidArgError = function(error) {
            return _Internal.RuntimeError._createInvalidArgError(error);
        }, CoreUtility._isLocalDocumentUrl = function(url) {
            return CoreUtility._getLocalDocumentUrlPrefixLength(url) > 0;
        }, CoreUtility._getLocalDocumentUrlPrefixLength = function(url) {
            for (var localDocumentPrefixes = [ "http://document.localhost", "https://document.localhost", "//document.localhost" ], urlLower = url.toLowerCase().trim(), i = 0; i < localDocumentPrefixes.length; i++) {
                if (urlLower === localDocumentPrefixes[i]) return localDocumentPrefixes[i].length;
                if (urlLower.substr(0, localDocumentPrefixes[i].length + 1) === localDocumentPrefixes[i] + "/") return localDocumentPrefixes[i].length + 1;
            }
            return 0;
        }, CoreUtility._validateLocalDocumentRequest = function(request) {
            var index = CoreUtility._getLocalDocumentUrlPrefixLength(request.url);
            if (index <= 0) throw _Internal.RuntimeError._createInvalidArgError({
                argumentName: "request"
            });
            var path = request.url.substr(index), pathLower = path.toLowerCase();
            return "_api" === pathLower ? path = "" : "_api/" === pathLower.substr(0, "_api/".length) && (path = path.substr("_api/".length)), 
            {
                method: request.method,
                url: path,
                headers: request.headers,
                body: request.body
            };
        }, CoreUtility._parseRequestFlagsAndCustomDataFromQueryStringIfAny = function(queryString) {
            for (var ret = {
                flags: -1,
                customData: ""
            }, parts = queryString.split("&"), i = 0; i < parts.length; i++) {
                var keyvalue = parts[i].split("=");
                if (keyvalue[0].toLowerCase() === CoreConstants.flags) {
                    var flags = parseInt(keyvalue[1]);
                    flags &= 4095, ret.flags = flags;
                } else keyvalue[0].toLowerCase() === CoreConstants.customData && (ret.customData = decodeURIComponent(keyvalue[1]));
            }
            return ret;
        }, CoreUtility._getRequestBodyText = function(request) {
            var body = "";
            return "string" == typeof request.body ? body = request.body : request.body && "object" == typeof request.body && (body = JSON.stringify(request.body)), 
            body;
        }, CoreUtility._parseResponseBody = function(response) {
            if ("string" == typeof response.body) {
                var bodyText = CoreUtility.trim(response.body);
                return JSON.parse(bodyText);
            }
            return response.body;
        }, CoreUtility._buildRequestMessageSafeArray = function(request) {
            var requestFlags = 0;
            CoreUtility.isReadonlyRestRequest(request.method) || (requestFlags = 1);
            var customData = "";
            if (request.url.substr(0, CoreConstants.processQuery.length).toLowerCase() === CoreConstants.processQuery.toLowerCase()) {
                var index = request.url.indexOf("?");
                if (index > 0) {
                    var queryString = request.url.substr(index + 1), flagsAndCustomData = CoreUtility._parseRequestFlagsAndCustomDataFromQueryStringIfAny(queryString);
                    flagsAndCustomData.flags >= 0 && (requestFlags = flagsAndCustomData.flags), customData = flagsAndCustomData.customData;
                }
            }
            return RichApiMessageUtility.buildRequestMessageSafeArray(customData, requestFlags, request.method, request.url, request.headers, CoreUtility._getRequestBodyText(request));
        }, CoreUtility._parseHttpResponseHeaders = function(allResponseHeaders) {
            var responseHeaders = {};
            if (!CoreUtility.isNullOrEmptyString(allResponseHeaders)) for (var regex = new RegExp("\r?\n"), entries = allResponseHeaders.split(regex), i = 0; i < entries.length; i++) {
                var entry = entries[i];
                if (null != entry) {
                    var index = entry.indexOf(":");
                    if (index > 0) {
                        var key = entry.substr(0, index), value = entry.substr(index + 1);
                        key = CoreUtility.trim(key), value = CoreUtility.trim(value), responseHeaders[key.toUpperCase()] = value;
                    }
                }
            }
            return responseHeaders;
        }, CoreUtility._parseErrorResponse = function(responseInfo) {
            var errorObj = null;
            if (CoreUtility.isPlainJsonObject(responseInfo.body)) errorObj = responseInfo.body; else if (!CoreUtility.isNullOrEmptyString(responseInfo.body)) {
                var errorResponseBody = CoreUtility.trim(responseInfo.body);
                try {
                    errorObj = JSON.parse(errorResponseBody);
                } catch (e) {
                    CoreUtility.log("Error when parse " + errorResponseBody);
                }
            }
            var statusCode = responseInfo.statusCode.toString();
            if (CoreUtility.isNullOrUndefined(errorObj) || "object" != typeof errorObj || !errorObj.error) return CoreUtility._createDefaultErrorResponse(statusCode);
            var error = errorObj.error, innerError = error.innerError;
            return innerError && innerError.code ? CoreUtility._createErrorResponse(innerError.code, statusCode, innerError.message) : error.code ? CoreUtility._createErrorResponse(error.code, statusCode, error.message) : CoreUtility._createDefaultErrorResponse(statusCode);
        }, CoreUtility._createDefaultErrorResponse = function(statusCode) {
            return {
                errorCode: CoreErrorCodes.connectionFailure,
                errorMessage: CoreUtility._getResourceString(CoreResourceStrings.connectionFailureWithStatus, statusCode)
            };
        }, CoreUtility._createErrorResponse = function(code, statusCode, message) {
            return {
                errorCode: code,
                errorMessage: CoreUtility._getResourceString(CoreResourceStrings.connectionFailureWithDetails, [ statusCode, code, message ])
            };
        }, CoreUtility._copyHeaders = function(src, dest) {
            if (src && dest) for (var key in src) dest[key] = src[key];
        }, CoreUtility.addResourceStringValues = function(values) {
            for (var key in values) CoreUtility.s_resourceStringValues[key] = values[key];
        }, CoreUtility._logEnabled = !1, CoreUtility.s_resourceStringValues = {
            ApiNotFoundDetails: "The method or property {0} is part of the {1} requirement set, which is not available in your version of {2}.",
            ConnectionFailureWithStatus: "The request failed with status code of {0}.",
            ConnectionFailureWithDetails: "The request failed with status code of {0}, error code {1} and the following error message: {2}",
            InvalidArgument: "The argument '{0}' doesn't work for this situation, is missing, or isn't in the right format.",
            InvalidObjectPath: 'The object path \'{0}\' isn\'t working for what you\'re trying to do. If you\'re using the object across multiple "context.sync" calls and outside the sequential execution of a ".run" batch, please use the "context.trackedObjects.add()" and "context.trackedObjects.remove()" methods to manage the object\'s lifetime.',
            InvalidRequestContext: "Cannot use the object across different request contexts.",
            Timeout: "The operation has timed out.",
            ValueNotLoaded: 'The value of the result object has not been loaded yet. Before reading the value property, call "context.sync()" on the associated request context.'
        }, CoreUtility;
    }();
    exports.CoreUtility = CoreUtility;
    var TestUtility = function() {
        function TestUtility() {}
        return TestUtility.setMock = function(value) {
            TestUtility.s_isMock = value;
        }, TestUtility.isMock = function() {
            return TestUtility.s_isMock;
        }, TestUtility;
    }();
    exports.TestUtility = TestUtility;
}, function(module, exports, __webpack_require__) {
    "use strict";
    var extendStatics, __extends = this && this.__extends || (extendStatics = function(d, b) {
        return (extendStatics = Object.setPrototypeOf || {
            __proto__: []
        } instanceof Array && function(d, b) {
            d.__proto__ = b;
        } || function(d, b) {
            for (var p in b) b.hasOwnProperty(p) && (d[p] = b[p]);
        })(d, b);
    }, function(d, b) {
        function __() {
            this.constructor = d;
        }
        extendStatics(d, b), d.prototype = null === b ? Object.create(b) : (__.prototype = b.prototype, 
        new __);
    }), __createBinding = this && this.__createBinding || (Object.create ? function(o, m, k, k2) {
        void 0 === k2 && (k2 = k), Object.defineProperty(o, k2, {
            enumerable: !0,
            get: function() {
                return m[k];
            }
        });
    } : function(o, m, k, k2) {
        void 0 === k2 && (k2 = k), o[k2] = m[k];
    }), __exportStar = this && this.__exportStar || function(m, exports) {
        for (var p in m) "default" === p || exports.hasOwnProperty(p) || __createBinding(exports, m, p);
    };
    Object.defineProperty(exports, "__esModule", {
        value: !0
    }), exports.CommonResourceStrings = exports.CommonUtility = exports.CommonConstants = exports.HttpRequestExecutor = exports.ClientResult = exports.ClientRequestBase = exports.ClientRequestContextBase = exports.ObjectPath = exports.Action = exports.ClientObjectBase = exports.CommonActionFactory = exports.config = exports._internalConfig = void 0;
    var Core = __webpack_require__(0);
    __exportStar(__webpack_require__(0), exports), exports._internalConfig = {
        showDisposeInfoInDebugInfo: !1,
        showInternalApiInDebugInfo: !1,
        enableEarlyDispose: !0,
        alwaysPolyfillClientObjectUpdateMethod: !1,
        alwaysPolyfillClientObjectRetrieveMethod: !1,
        enableConcurrentFlag: !0,
        enableUndoableFlag: !0,
        appendTypeNameToObjectPathInfo: !1
    }, exports.config = {
        extendedErrorLogging: !1
    };
    var CommonActionFactory = function() {
        function CommonActionFactory() {}
        return CommonActionFactory.createSetPropertyAction = function(context, parent, propertyName, value, flags) {
            CommonUtility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 4,
                Name: propertyName,
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            }, args = [ value ], referencedArgumentObjectPaths = CommonUtility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
            CommonUtility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
            var action = new Action(actionInfo, 0, flags);
            return action.referencedObjectPath = parent._objectPath, action.referencedArgumentObjectPaths = referencedArgumentObjectPaths, 
            parent._addAction(action);
        }, CommonActionFactory.createQueryAction = function(context, parent, queryOption, resultHandler) {
            CommonUtility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 2,
                Name: "",
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                QueryInfo: queryOption
            }, action = new Action(actionInfo, 1, 4);
            return action.referencedObjectPath = parent._objectPath, parent._addAction(action, resultHandler);
        }, CommonActionFactory.createQueryAsJsonAction = function(context, parent, queryOption, resultHandler) {
            CommonUtility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 7,
                Name: "",
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                QueryInfo: queryOption
            }, action = new Action(actionInfo, 1, 4);
            return action.referencedObjectPath = parent._objectPath, parent._addAction(action, resultHandler);
        }, CommonActionFactory.createUpdateAction = function(context, parent, objectState) {
            CommonUtility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 9,
                Name: "",
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                ObjectState: objectState
            }, action = new Action(actionInfo, 0, 0);
            return action.referencedObjectPath = parent._objectPath, parent._addAction(action);
        }, CommonActionFactory;
    }();
    exports.CommonActionFactory = CommonActionFactory;
    var ClientObjectBase = function() {
        function ClientObjectBase(contextBase, objectPath) {
            this.m_contextBase = contextBase, this.m_objectPath = objectPath;
        }
        return Object.defineProperty(ClientObjectBase.prototype, "_objectPath", {
            get: function() {
                return this.m_objectPath;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ClientObjectBase.prototype, "_context", {
            get: function() {
                return this.m_contextBase;
            },
            enumerable: !1,
            configurable: !0
        }), ClientObjectBase.prototype._addAction = function(action, resultHandler) {
            var _this = this;
            return void 0 === resultHandler && (resultHandler = null), Core.CoreUtility.createPromise((function(resolve, reject) {
                _this._context._addServiceApiAction(action, resultHandler, resolve, reject);
            }));
        }, ClientObjectBase.prototype._retrieve = function(option, resultHandler) {
            var shouldPolyfill = exports._internalConfig.alwaysPolyfillClientObjectRetrieveMethod;
            shouldPolyfill || (shouldPolyfill = !CommonUtility.isSetSupported("RichApiRuntime", "1.1"));
            var queryOption = ClientRequestContextBase._parseQueryOption(option);
            return shouldPolyfill ? CommonActionFactory.createQueryAction(this._context, this, queryOption, resultHandler) : CommonActionFactory.createQueryAsJsonAction(this._context, this, queryOption, resultHandler);
        }, ClientObjectBase.prototype._recursivelyUpdate = function(properties) {
            var shouldPolyfill = exports._internalConfig.alwaysPolyfillClientObjectUpdateMethod;
            shouldPolyfill || (shouldPolyfill = !CommonUtility.isSetSupported("RichApiRuntime", "1.2"));
            try {
                var scalarPropNames = this[CommonConstants.scalarPropertyNames];
                scalarPropNames || (scalarPropNames = []);
                var scalarPropUpdatable = this[CommonConstants.scalarPropertyUpdateable];
                if (!scalarPropUpdatable) {
                    scalarPropUpdatable = [];
                    for (var i = 0; i < scalarPropNames.length; i++) scalarPropUpdatable.push(!1);
                }
                var navigationPropNames = this[CommonConstants.navigationPropertyNames];
                navigationPropNames || (navigationPropNames = []);
                var scalarProps = {}, navigationProps = {}, scalarPropCount = 0;
                for (var propName in properties) {
                    var index = scalarPropNames.indexOf(propName);
                    if (index >= 0) {
                        if (!scalarPropUpdatable[index]) throw new Core._Internal.RuntimeError({
                            code: Core.CoreErrorCodes.invalidArgument,
                            httpStatusCode: 400,
                            message: Core.CoreUtility._getResourceString(CommonResourceStrings.attemptingToSetReadOnlyProperty, propName),
                            debugInfo: {
                                errorLocation: propName
                            }
                        });
                        scalarProps[propName] = properties[propName], ++scalarPropCount;
                    } else {
                        if (!(navigationPropNames.indexOf(propName) >= 0)) throw new Core._Internal.RuntimeError({
                            code: Core.CoreErrorCodes.invalidArgument,
                            httpStatusCode: 400,
                            message: Core.CoreUtility._getResourceString(CommonResourceStrings.propertyDoesNotExist, propName),
                            debugInfo: {
                                errorLocation: propName
                            }
                        });
                        navigationProps[propName] = properties[propName];
                    }
                }
                if (scalarPropCount > 0) if (shouldPolyfill) for (i = 0; i < scalarPropNames.length; i++) {
                    var propValue = scalarProps[propName = scalarPropNames[i]];
                    CommonUtility.isUndefined(propValue) || CommonActionFactory.createSetPropertyAction(this._context, this, propName, propValue);
                } else CommonActionFactory.createUpdateAction(this._context, this, scalarProps);
                for (var propName in navigationProps) {
                    var navigationPropProxy = this[propName], navigationPropValue = navigationProps[propName];
                    navigationPropProxy._recursivelyUpdate(navigationPropValue);
                }
            } catch (innerError) {
                throw new Core._Internal.RuntimeError({
                    code: Core.CoreErrorCodes.invalidArgument,
                    httpStatusCode: 400,
                    message: Core.CoreUtility._getResourceString(Core.CoreResourceStrings.invalidArgument, "properties"),
                    debugInfo: {
                        errorLocation: this._className + ".update"
                    },
                    innerError: innerError
                });
            }
        }, ClientObjectBase;
    }();
    exports.ClientObjectBase = ClientObjectBase;
    var Action = function() {
        function Action(actionInfo, operationType, flags) {
            this.m_actionInfo = actionInfo, this.m_operationType = operationType, this.m_flags = flags;
        }
        return Object.defineProperty(Action.prototype, "actionInfo", {
            get: function() {
                return this.m_actionInfo;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(Action.prototype, "operationType", {
            get: function() {
                return this.m_operationType;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(Action.prototype, "flags", {
            get: function() {
                return this.m_flags;
            },
            enumerable: !1,
            configurable: !0
        }), Action;
    }();
    exports.Action = Action;
    var ObjectPath = function() {
        function ObjectPath(objectPathInfo, parentObjectPath, isCollection, isInvalidAfterRequest, operationType, flags) {
            this.m_objectPathInfo = objectPathInfo, this.m_parentObjectPath = parentObjectPath, 
            this.m_isCollection = isCollection, this.m_isInvalidAfterRequest = isInvalidAfterRequest, 
            this.m_isValid = !0, this.m_operationType = operationType, this.m_flags = flags;
        }
        return Object.defineProperty(ObjectPath.prototype, "id", {
            get: function() {
                var argumentInfo = this.m_objectPathInfo.ArgumentInfo;
                if (argumentInfo) {
                    var argument = argumentInfo.Arguments;
                    if (argument) return argument[0];
                }
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ObjectPath.prototype, "parent", {
            get: function() {
                var parent = this.m_parentObjectPath;
                if (parent) return parent;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ObjectPath.prototype, "parentId", {
            get: function() {
                return this.parent ? this.parent.id : void 0;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ObjectPath.prototype, "objectPathInfo", {
            get: function() {
                return this.m_objectPathInfo;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ObjectPath.prototype, "operationType", {
            get: function() {
                return this.m_operationType;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ObjectPath.prototype, "flags", {
            get: function() {
                return this.m_flags;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ObjectPath.prototype, "isCollection", {
            get: function() {
                return this.m_isCollection;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ObjectPath.prototype, "isInvalidAfterRequest", {
            get: function() {
                return this.m_isInvalidAfterRequest;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ObjectPath.prototype, "parentObjectPath", {
            get: function() {
                return this.m_parentObjectPath;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ObjectPath.prototype, "argumentObjectPaths", {
            get: function() {
                return this.m_argumentObjectPaths;
            },
            set: function(value) {
                this.m_argumentObjectPaths = value;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ObjectPath.prototype, "isValid", {
            get: function() {
                return this.m_isValid;
            },
            set: function(value) {
                this.m_isValid = value, !value && 6 === this.m_objectPathInfo.ObjectPathType && this.m_savedObjectPathInfo && (ObjectPath.copyObjectPathInfo(this.m_savedObjectPathInfo.pathInfo, this.m_objectPathInfo), 
                this.m_parentObjectPath = this.m_savedObjectPathInfo.parent, this.m_isValid = !0, 
                this.m_savedObjectPathInfo = null);
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ObjectPath.prototype, "originalObjectPathInfo", {
            get: function() {
                return this.m_originalObjectPathInfo;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ObjectPath.prototype, "getByIdMethodName", {
            get: function() {
                return this.m_getByIdMethodName;
            },
            set: function(value) {
                this.m_getByIdMethodName = value;
            },
            enumerable: !1,
            configurable: !0
        }), ObjectPath.prototype._updateAsNullObject = function() {
            this.resetForUpdateUsingObjectData(), this.m_objectPathInfo.ObjectPathType = 7, 
            this.m_objectPathInfo.Name = "", this.m_parentObjectPath = null;
        }, ObjectPath.prototype.saveOriginalObjectPathInfo = function() {
            exports.config.extendedErrorLogging && !this.m_originalObjectPathInfo && (this.m_originalObjectPathInfo = {}, 
            ObjectPath.copyObjectPathInfo(this.m_objectPathInfo, this.m_originalObjectPathInfo));
        }, ObjectPath.prototype.updateUsingObjectData = function(value, clientObject) {
            var referenceId = value[CommonConstants.referenceId];
            if (!Core.CoreUtility.isNullOrEmptyString(referenceId)) {
                if (!this.m_savedObjectPathInfo && !this.isInvalidAfterRequest && ObjectPath.isRestorableObjectPath(this.m_objectPathInfo.ObjectPathType)) {
                    var pathInfo = {};
                    ObjectPath.copyObjectPathInfo(this.m_objectPathInfo, pathInfo), this.m_savedObjectPathInfo = {
                        pathInfo: pathInfo,
                        parent: this.m_parentObjectPath
                    };
                }
                return this.saveOriginalObjectPathInfo(), this.resetForUpdateUsingObjectData(), 
                this.m_objectPathInfo.ObjectPathType = 6, this.m_objectPathInfo.Name = referenceId, 
                delete this.m_objectPathInfo.ParentObjectPathId, void (this.m_parentObjectPath = null);
            }
            if (clientObject) {
                var collectionPropertyPath = clientObject[CommonConstants.collectionPropertyPath];
                if (!Core.CoreUtility.isNullOrEmptyString(collectionPropertyPath) && clientObject.context) {
                    var id = CommonUtility.tryGetObjectIdFromLoadOrRetrieveResult(value);
                    if (!Core.CoreUtility.isNullOrUndefined(id)) {
                        for (var propNames = collectionPropertyPath.split("."), parent_1 = clientObject.context[propNames[0]], i = 1; i < propNames.length; i++) parent_1 = parent_1[propNames[i]];
                        return this.saveOriginalObjectPathInfo(), this.resetForUpdateUsingObjectData(), 
                        this.m_parentObjectPath = parent_1._objectPath, this.m_objectPathInfo.ParentObjectPathId = this.m_parentObjectPath.objectPathInfo.Id, 
                        this.m_objectPathInfo.ObjectPathType = 5, this.m_objectPathInfo.Name = "", void (this.m_objectPathInfo.ArgumentInfo.Arguments = [ id ]);
                    }
                }
            }
            var parentIsCollection = this.parentObjectPath && this.parentObjectPath.isCollection, getByIdMethodName = this.getByIdMethodName;
            if (parentIsCollection || !Core.CoreUtility.isNullOrEmptyString(getByIdMethodName)) {
                id = CommonUtility.tryGetObjectIdFromLoadOrRetrieveResult(value);
                if (!Core.CoreUtility.isNullOrUndefined(id)) return this.saveOriginalObjectPathInfo(), 
                this.resetForUpdateUsingObjectData(), Core.CoreUtility.isNullOrEmptyString(getByIdMethodName) ? (this.m_objectPathInfo.ObjectPathType = 5, 
                this.m_objectPathInfo.Name = "") : (this.m_objectPathInfo.ObjectPathType = 3, this.m_objectPathInfo.Name = getByIdMethodName), 
                void (this.m_objectPathInfo.ArgumentInfo.Arguments = [ id ]);
            }
        }, ObjectPath.prototype.resetForUpdateUsingObjectData = function() {
            this.m_isInvalidAfterRequest = !1, this.m_isValid = !0, this.m_operationType = 1, 
            this.m_flags = 4, this.m_objectPathInfo.ArgumentInfo = {}, this.m_argumentObjectPaths = null, 
            this.m_getByIdMethodName = null;
        }, ObjectPath.isRestorableObjectPath = function(objectPathType) {
            return 1 === objectPathType || 5 === objectPathType || 3 === objectPathType || 4 === objectPathType;
        }, ObjectPath.copyObjectPathInfo = function(src, dest) {
            dest.Id = src.Id, dest.ArgumentInfo = src.ArgumentInfo, dest.Name = src.Name, dest.ObjectPathType = src.ObjectPathType, 
            dest.ParentObjectPathId = src.ParentObjectPathId;
        }, ObjectPath;
    }();
    exports.ObjectPath = ObjectPath;
    var ClientRequestContextBase = function() {
        function ClientRequestContextBase() {
            this.m_nextId = 0;
        }
        return ClientRequestContextBase.prototype._nextId = function() {
            return ++this.m_nextId;
        }, ClientRequestContextBase.prototype._addServiceApiAction = function(action, resultHandler, resolve, reject) {
            this.m_serviceApiQueue || (this.m_serviceApiQueue = new ServiceApiQueue(this)), 
            this.m_serviceApiQueue.add(action, resultHandler, resolve, reject);
        }, ClientRequestContextBase._parseQueryOption = function(option) {
            var queryOption = {};
            if ("string" == typeof option) {
                var select = option;
                queryOption.Select = CommonUtility._parseSelectExpand(select);
            } else if (Array.isArray(option)) queryOption.Select = option; else if ("object" == typeof option) {
                var loadOption = option;
                if (ClientRequestContextBase.isLoadOption(loadOption)) {
                    if ("string" == typeof loadOption.select) queryOption.Select = CommonUtility._parseSelectExpand(loadOption.select); else if (Array.isArray(loadOption.select)) queryOption.Select = loadOption.select; else if (!CommonUtility.isNullOrUndefined(loadOption.select)) throw Core._Internal.RuntimeError._createInvalidArgError({
                        argumentName: "option.select"
                    });
                    if ("string" == typeof loadOption.expand) queryOption.Expand = CommonUtility._parseSelectExpand(loadOption.expand); else if (Array.isArray(loadOption.expand)) queryOption.Expand = loadOption.expand; else if (!CommonUtility.isNullOrUndefined(loadOption.expand)) throw Core._Internal.RuntimeError._createInvalidArgError({
                        argumentName: "option.expand"
                    });
                    if ("number" == typeof loadOption.top) queryOption.Top = loadOption.top; else if (!CommonUtility.isNullOrUndefined(loadOption.top)) throw Core._Internal.RuntimeError._createInvalidArgError({
                        argumentName: "option.top"
                    });
                    if ("number" == typeof loadOption.skip) queryOption.Skip = loadOption.skip; else if (!CommonUtility.isNullOrUndefined(loadOption.skip)) throw Core._Internal.RuntimeError._createInvalidArgError({
                        argumentName: "option.skip"
                    });
                } else queryOption = ClientRequestContextBase.parseStrictLoadOption(option);
            } else if (!CommonUtility.isNullOrUndefined(option)) throw Core._Internal.RuntimeError._createInvalidArgError({
                argumentName: "option"
            });
            return queryOption;
        }, ClientRequestContextBase.isLoadOption = function(loadOption) {
            if (!CommonUtility.isUndefined(loadOption.select) && ("string" == typeof loadOption.select || Array.isArray(loadOption.select))) return !0;
            if (!CommonUtility.isUndefined(loadOption.expand) && ("string" == typeof loadOption.expand || Array.isArray(loadOption.expand))) return !0;
            if (!CommonUtility.isUndefined(loadOption.top) && "number" == typeof loadOption.top) return !0;
            if (!CommonUtility.isUndefined(loadOption.skip) && "number" == typeof loadOption.skip) return !0;
            for (var i in loadOption) return !1;
            return !0;
        }, ClientRequestContextBase.parseStrictLoadOption = function(option) {
            var ret = {
                Select: []
            };
            return ClientRequestContextBase.parseStrictLoadOptionHelper(ret, "", "option", option), 
            ret;
        }, ClientRequestContextBase.combineQueryPath = function(pathPrefix, key, separator) {
            return 0 === pathPrefix.length ? key : pathPrefix + separator + key;
        }, ClientRequestContextBase.parseStrictLoadOptionHelper = function(queryInfo, pathPrefix, argPrefix, option) {
            for (var key in option) {
                var value = option[key];
                if ("$all" === key) {
                    if ("boolean" != typeof value) throw Core._Internal.RuntimeError._createInvalidArgError({
                        argumentName: ClientRequestContextBase.combineQueryPath(argPrefix, key, ".")
                    });
                    value && queryInfo.Select.push(ClientRequestContextBase.combineQueryPath(pathPrefix, "*", "/"));
                } else if ("$top" === key) {
                    if ("number" != typeof value || pathPrefix.length > 0) throw Core._Internal.RuntimeError._createInvalidArgError({
                        argumentName: ClientRequestContextBase.combineQueryPath(argPrefix, key, ".")
                    });
                    queryInfo.Top = value;
                } else if ("$skip" === key) {
                    if ("number" != typeof value || pathPrefix.length > 0) throw Core._Internal.RuntimeError._createInvalidArgError({
                        argumentName: ClientRequestContextBase.combineQueryPath(argPrefix, key, ".")
                    });
                    queryInfo.Skip = value;
                } else if ("boolean" == typeof value) value && queryInfo.Select.push(ClientRequestContextBase.combineQueryPath(pathPrefix, key, "/")); else {
                    if ("object" != typeof value) throw Core._Internal.RuntimeError._createInvalidArgError({
                        argumentName: ClientRequestContextBase.combineQueryPath(argPrefix, key, ".")
                    });
                    ClientRequestContextBase.parseStrictLoadOptionHelper(queryInfo, ClientRequestContextBase.combineQueryPath(pathPrefix, key, "/"), ClientRequestContextBase.combineQueryPath(argPrefix, key, "."), value);
                }
            }
        }, ClientRequestContextBase;
    }();
    exports.ClientRequestContextBase = ClientRequestContextBase;
    var InstantiateActionUpdateObjectPathHandler = function() {
        function InstantiateActionUpdateObjectPathHandler(m_objectPath) {
            this.m_objectPath = m_objectPath;
        }
        return InstantiateActionUpdateObjectPathHandler.prototype._handleResult = function(value) {
            Core.CoreUtility.isNullOrUndefined(value) ? this.m_objectPath._updateAsNullObject() : this.m_objectPath.updateUsingObjectData(value, null);
        }, InstantiateActionUpdateObjectPathHandler;
    }(), ClientRequestBase = function() {
        function ClientRequestBase(context) {
            this.m_contextBase = context, this.m_actions = [], this.m_actionResultHandler = {}, 
            this.m_referencedObjectPaths = {}, this.m_instantiatedObjectPaths = {}, this.m_preSyncPromises = [];
        }
        return ClientRequestBase.prototype.addAction = function(action) {
            this.m_actions.push(action), 1 == action.actionInfo.ActionType && (this.m_instantiatedObjectPaths[action.actionInfo.ObjectPathId] = action);
        }, Object.defineProperty(ClientRequestBase.prototype, "hasActions", {
            get: function() {
                return this.m_actions.length > 0;
            },
            enumerable: !1,
            configurable: !0
        }), ClientRequestBase.prototype._getLastAction = function() {
            return this.m_actions[this.m_actions.length - 1];
        }, ClientRequestBase.prototype.ensureInstantiateObjectPath = function(objectPath) {
            if (objectPath) {
                if (this.m_instantiatedObjectPaths[objectPath.objectPathInfo.Id]) return;
                if (this.ensureInstantiateObjectPath(objectPath.parentObjectPath), this.ensureInstantiateObjectPaths(objectPath.argumentObjectPaths), 
                !this.m_instantiatedObjectPaths[objectPath.objectPathInfo.Id]) {
                    var actionInfo = {
                        Id: this.m_contextBase._nextId(),
                        ActionType: 1,
                        Name: "",
                        ObjectPathId: objectPath.objectPathInfo.Id
                    }, instantiateAction = new Action(actionInfo, 1, 4);
                    instantiateAction.referencedObjectPath = objectPath, this.addReferencedObjectPath(objectPath), 
                    this.addAction(instantiateAction);
                    var resultHandler = new InstantiateActionUpdateObjectPathHandler(objectPath);
                    this.addActionResultHandler(instantiateAction, resultHandler);
                }
            }
        }, ClientRequestBase.prototype.ensureInstantiateObjectPaths = function(objectPaths) {
            if (objectPaths) for (var i = 0; i < objectPaths.length; i++) this.ensureInstantiateObjectPath(objectPaths[i]);
        }, ClientRequestBase.prototype.addReferencedObjectPath = function(objectPath) {
            if (objectPath && !this.m_referencedObjectPaths[objectPath.objectPathInfo.Id]) {
                if (!objectPath.isValid) throw new Core._Internal.RuntimeError({
                    code: Core.CoreErrorCodes.invalidObjectPath,
                    httpStatusCode: 400,
                    message: Core.CoreUtility._getResourceString(Core.CoreResourceStrings.invalidObjectPath, CommonUtility.getObjectPathExpression(objectPath)),
                    debugInfo: {
                        errorLocation: CommonUtility.getObjectPathExpression(objectPath)
                    }
                });
                for (;objectPath; ) this.m_referencedObjectPaths[objectPath.objectPathInfo.Id] = objectPath, 
                3 == objectPath.objectPathInfo.ObjectPathType && this.addReferencedObjectPaths(objectPath.argumentObjectPaths), 
                objectPath = objectPath.parentObjectPath;
            }
        }, ClientRequestBase.prototype.addReferencedObjectPaths = function(objectPaths) {
            if (objectPaths) for (var i = 0; i < objectPaths.length; i++) this.addReferencedObjectPath(objectPaths[i]);
        }, ClientRequestBase.prototype.addActionResultHandler = function(action, resultHandler) {
            this.m_actionResultHandler[action.actionInfo.Id] = resultHandler;
        }, ClientRequestBase.prototype.aggregrateRequestFlags = function(requestFlags, operationType, flags) {
            return 0 === operationType && (requestFlags |= 1, 0 == (2 & flags) && (requestFlags &= -17), 
            0 == (8 & flags) && (requestFlags &= -257), requestFlags &= -5), 1 & flags && (requestFlags |= 2), 
            0 == (4 & flags) && (requestFlags &= -5), requestFlags;
        }, ClientRequestBase.prototype.finallyNormalizeFlags = function(requestFlags) {
            return 0 == (1 & requestFlags) && (requestFlags &= -17, requestFlags &= -257), exports._internalConfig.enableConcurrentFlag || (requestFlags &= -5), 
            exports._internalConfig.enableUndoableFlag || (requestFlags &= -17), CommonUtility.isSetSupported("RichApiRuntimeFlag", "1.1") || (requestFlags &= -5, 
            requestFlags &= -17), CommonUtility.isSetSupported("RichApiRuntimeFlag", "1.2") || (requestFlags &= -257), 
            "number" == typeof this.m_flagsForTesting && (requestFlags = this.m_flagsForTesting), 
            requestFlags;
        }, ClientRequestBase.prototype.buildRequestMessageBodyAndRequestFlags = function() {
            exports._internalConfig.enableEarlyDispose && ClientRequestBase._calculateLastUsedObjectPathIds(this.m_actions);
            var requestFlags = 276, objectPaths = {};
            for (var i in this.m_referencedObjectPaths) requestFlags = this.aggregrateRequestFlags(requestFlags, this.m_referencedObjectPaths[i].operationType, this.m_referencedObjectPaths[i].flags), 
            objectPaths[i] = this.m_referencedObjectPaths[i].objectPathInfo;
            for (var actions = [], hasKeepReference = !1, index = 0; index < this.m_actions.length; index++) {
                var action = this.m_actions[index];
                3 === action.actionInfo.ActionType && action.actionInfo.Name === CommonConstants.keepReference && (hasKeepReference = !0), 
                requestFlags = this.aggregrateRequestFlags(requestFlags, action.operationType, action.flags), 
                actions.push(action.actionInfo);
            }
            return requestFlags = this.finallyNormalizeFlags(requestFlags), {
                body: {
                    AutoKeepReference: this.m_contextBase._autoCleanup && hasKeepReference,
                    Actions: actions,
                    ObjectPaths: objectPaths
                },
                flags: requestFlags
            };
        }, ClientRequestBase.prototype.processResponse = function(actionResults) {
            if (actionResults) for (var i = 0; i < actionResults.length; i++) {
                var actionResult = actionResults[i], handler = this.m_actionResultHandler[actionResult.ActionId];
                handler && handler._handleResult(actionResult.Value);
            }
        }, ClientRequestBase.prototype.invalidatePendingInvalidObjectPaths = function() {
            for (var i in this.m_referencedObjectPaths) this.m_referencedObjectPaths[i].isInvalidAfterRequest && (this.m_referencedObjectPaths[i].isValid = !1);
        }, ClientRequestBase.prototype._addPreSyncPromise = function(value) {
            this.m_preSyncPromises.push(value);
        }, Object.defineProperty(ClientRequestBase.prototype, "_preSyncPromises", {
            get: function() {
                return this.m_preSyncPromises;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ClientRequestBase.prototype, "_actions", {
            get: function() {
                return this.m_actions;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ClientRequestBase.prototype, "_objectPaths", {
            get: function() {
                return this.m_referencedObjectPaths;
            },
            enumerable: !1,
            configurable: !0
        }), ClientRequestBase.prototype._removeKeepReferenceAction = function(objectPathId) {
            for (var i = this.m_actions.length - 1; i >= 0; i--) {
                var actionInfo = this.m_actions[i].actionInfo;
                if (actionInfo.ObjectPathId === objectPathId && 3 === actionInfo.ActionType && actionInfo.Name === CommonConstants.keepReference) {
                    this.m_actions.splice(i, 1);
                    break;
                }
            }
        }, ClientRequestBase._updateLastUsedActionIdOfObjectPathId = function(lastUsedActionIdOfObjectPathId, objectPath, actionId) {
            for (;objectPath; ) {
                if (lastUsedActionIdOfObjectPathId[objectPath.objectPathInfo.Id]) return;
                lastUsedActionIdOfObjectPathId[objectPath.objectPathInfo.Id] = actionId;
                var argumentObjectPaths = objectPath.argumentObjectPaths;
                if (argumentObjectPaths) for (var argumentObjectPathsLength = argumentObjectPaths.length, i = 0; i < argumentObjectPathsLength; i++) ClientRequestBase._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, argumentObjectPaths[i], actionId);
                objectPath = objectPath.parentObjectPath;
            }
        }, ClientRequestBase._calculateLastUsedObjectPathIds = function(actions) {
            for (var lastUsedActionIdOfObjectPathId = {}, actionsLength = actions.length, index = actionsLength - 1; index >= 0; --index) {
                var actionId = (action = actions[index]).actionInfo.Id;
                action.referencedObjectPath && ClientRequestBase._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, action.referencedObjectPath, actionId);
                var referencedObjectPaths = action.referencedArgumentObjectPaths;
                if (referencedObjectPaths) for (var referencedObjectPathsLength = referencedObjectPaths.length, refIndex = 0; refIndex < referencedObjectPathsLength; refIndex++) ClientRequestBase._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, referencedObjectPaths[refIndex], actionId);
            }
            var lastUsedObjectPathIdsOfAction = {};
            for (var key in lastUsedActionIdOfObjectPathId) {
                var objectPathIds = lastUsedObjectPathIdsOfAction[actionId = lastUsedActionIdOfObjectPathId[key]];
                objectPathIds || (objectPathIds = [], lastUsedObjectPathIdsOfAction[actionId] = objectPathIds), 
                objectPathIds.push(parseInt(key));
            }
            for (index = 0; index < actionsLength; index++) {
                var action, lastUsedObjectPathIds = lastUsedObjectPathIdsOfAction[(action = actions[index]).actionInfo.Id];
                lastUsedObjectPathIds && lastUsedObjectPathIds.length > 0 ? action.actionInfo.L = lastUsedObjectPathIds : action.actionInfo.L && delete action.actionInfo.L;
            }
        }, ClientRequestBase;
    }();
    exports.ClientRequestBase = ClientRequestBase;
    var ClientResult = function() {
        function ClientResult(m_type) {
            this.m_type = m_type;
        }
        return Object.defineProperty(ClientResult.prototype, "value", {
            get: function() {
                if (!this.m_isLoaded) throw new Core._Internal.RuntimeError({
                    code: Core.CoreErrorCodes.valueNotLoaded,
                    httpStatusCode: 400,
                    message: Core.CoreUtility._getResourceString(Core.CoreResourceStrings.valueNotLoaded),
                    debugInfo: {
                        errorLocation: "clientResult.value"
                    }
                });
                return this.m_value;
            },
            enumerable: !1,
            configurable: !0
        }), ClientResult.prototype._handleResult = function(value) {
            this.m_isLoaded = !0, "object" == typeof value && value && value._IsNull || (1 === this.m_type ? this.m_value = CommonUtility.adjustToDateTime(value) : this.m_value = value);
        }, ClientResult;
    }();
    exports.ClientResult = ClientResult;
    var ServiceApiQueue = function() {
        function ServiceApiQueue(m_context) {
            this.m_context = m_context, this.m_actions = [];
        }
        return ServiceApiQueue.prototype.add = function(action, resultHandler, resolve, reject) {
            var _this = this;
            this.m_actions.push({
                action: action,
                resultHandler: resultHandler,
                resolve: resolve,
                reject: reject
            }), 1 === this.m_actions.length && setTimeout((function() {
                return _this.processActions();
            }), 0);
        }, ServiceApiQueue.prototype.processActions = function() {
            var _this = this;
            if (0 !== this.m_actions.length) {
                var actions = this.m_actions;
                this.m_actions = [];
                for (var request = new ClientRequestBase(this.m_context), i = 0; i < actions.length; i++) {
                    var action = actions[i];
                    request.ensureInstantiateObjectPath(action.action.referencedObjectPath), request.ensureInstantiateObjectPaths(action.action.referencedArgumentObjectPaths), 
                    request.addAction(action.action), request.addReferencedObjectPath(action.action.referencedObjectPath), 
                    request.addReferencedObjectPaths(action.action.referencedArgumentObjectPaths);
                }
                var _a = request.buildRequestMessageBodyAndRequestFlags(), body = _a.body, flags = _a.flags, requestMessage = {
                    Url: Core.CoreConstants.localDocumentApiPrefix,
                    Headers: null,
                    Body: body
                };
                Core.CoreUtility.log("Request:"), Core.CoreUtility.log(JSON.stringify(body)), (new HttpRequestExecutor).executeAsync(this.m_context._customData, flags, requestMessage).then((function(response) {
                    _this.processResponse(request, actions, response);
                })).catch((function(ex) {
                    for (var i = 0; i < actions.length; i++) {
                        actions[i].reject(ex);
                    }
                }));
            }
        }, ServiceApiQueue.prototype.processResponse = function(request, actions, response) {
            var error = this.getErrorFromResponse(response), actionResults = null;
            response.Body.Results ? actionResults = response.Body.Results : response.Body.ProcessedResults && response.Body.ProcessedResults.Results && (actionResults = response.Body.ProcessedResults.Results), 
            actionResults || (actionResults = []), this.processActionResults(request, actions, actionResults, error);
        }, ServiceApiQueue.prototype.getErrorFromResponse = function(response) {
            return Core.CoreUtility.isNullOrEmptyString(response.ErrorCode) ? response.Body && response.Body.Error ? new Core._Internal.RuntimeError({
                code: response.Body.Error.Code,
                httpStatusCode: response.Body.Error.HttpStatusCode,
                message: response.Body.Error.Message
            }) : null : new Core._Internal.RuntimeError({
                code: response.ErrorCode,
                httpStatusCode: response.HttpStatusCode,
                message: response.ErrorMessage
            });
        }, ServiceApiQueue.prototype.processActionResults = function(request, actions, actionResults, err) {
            request.processResponse(actionResults);
            for (var i = 0; i < actions.length; i++) {
                for (var action = actions[i], actionId = action.action.actionInfo.Id, hasResult = !1, j = 0; j < actionResults.length; j++) if (actionId == actionResults[j].ActionId) {
                    var resultValue = actionResults[j].Value;
                    action.resultHandler && (action.resultHandler._handleResult(resultValue), resultValue = action.resultHandler.value), 
                    action.resolve && action.resolve(resultValue), hasResult = !0;
                    break;
                }
                !hasResult && action.reject && (err ? action.reject(err) : action.reject("No response for the action."));
            }
        }, ServiceApiQueue;
    }(), HttpRequestExecutor = function() {
        function HttpRequestExecutor() {}
        return HttpRequestExecutor.prototype.getRequestUrl = function(baseUrl, requestFlags) {
            return "/" != baseUrl.charAt(baseUrl.length - 1) && (baseUrl += "/"), baseUrl = (baseUrl += Core.CoreConstants.processQuery) + "?" + Core.CoreConstants.flags + "=" + requestFlags.toString();
        }, HttpRequestExecutor.prototype.executeAsync = function(customData, requestFlags, requestMessage) {
            var requestInfo = {
                method: "POST",
                url: this.getRequestUrl(requestMessage.Url, requestFlags),
                headers: {},
                body: requestMessage.Body
            };
            if (requestInfo.headers[Core.CoreConstants.sourceLibHeader] = HttpRequestExecutor.SourceLibHeaderValue, 
            requestInfo.headers["CONTENT-TYPE"] = "application/json", requestMessage.Headers) for (var key in requestMessage.Headers) requestInfo.headers[key] = requestMessage.Headers[key];
            return (Core.CoreUtility._isLocalDocumentUrl(requestInfo.url) ? Core.HttpUtility.sendLocalDocumentRequest : Core.HttpUtility.sendRequest)(requestInfo).then((function(responseInfo) {
                var response;
                if (200 === responseInfo.statusCode) response = {
                    HttpStatusCode: responseInfo.statusCode,
                    ErrorCode: null,
                    ErrorMessage: null,
                    Headers: responseInfo.headers,
                    Body: Core.CoreUtility._parseResponseBody(responseInfo)
                }; else {
                    Core.CoreUtility.log("Error Response:" + responseInfo.body);
                    var error = Core.CoreUtility._parseErrorResponse(responseInfo);
                    response = {
                        HttpStatusCode: responseInfo.statusCode,
                        ErrorCode: error.errorCode,
                        ErrorMessage: error.errorMessage,
                        Headers: responseInfo.headers,
                        Body: null
                    };
                }
                return response;
            }));
        }, HttpRequestExecutor.SourceLibHeaderValue = "officejs-rest", HttpRequestExecutor;
    }();
    exports.HttpRequestExecutor = HttpRequestExecutor;
    var CommonConstants = function(_super) {
        function CommonConstants() {
            return null !== _super && _super.apply(this, arguments) || this;
        }
        return __extends(CommonConstants, _super), CommonConstants.collectionPropertyPath = "_collectionPropertyPath", 
        CommonConstants.id = "Id", CommonConstants.idLowerCase = "id", CommonConstants.idPrivate = "_Id", 
        CommonConstants.keepReference = "_KeepReference", CommonConstants.objectPathIdPrivate = "_ObjectPathId", 
        CommonConstants.referenceId = "_ReferenceId", CommonConstants.items = "_Items", 
        CommonConstants.itemsLowerCase = "items", CommonConstants.scalarPropertyNames = "_scalarPropertyNames", 
        CommonConstants.scalarPropertyOriginalNames = "_scalarPropertyOriginalNames", CommonConstants.navigationPropertyNames = "_navigationPropertyNames", 
        CommonConstants.scalarPropertyUpdateable = "_scalarPropertyUpdateable", CommonConstants;
    }(Core.CoreConstants);
    exports.CommonConstants = CommonConstants;
    var CommonUtility = function(_super) {
        function CommonUtility() {
            return null !== _super && _super.apply(this, arguments) || this;
        }
        return __extends(CommonUtility, _super), CommonUtility.validateObjectPath = function(clientObject) {
            for (var objectPath = clientObject._objectPath; objectPath; ) {
                if (!objectPath.isValid) throw new Core._Internal.RuntimeError({
                    code: Core.CoreErrorCodes.invalidObjectPath,
                    httpStatusCode: 400,
                    message: Core.CoreUtility._getResourceString(Core.CoreResourceStrings.invalidObjectPath, CommonUtility.getObjectPathExpression(objectPath)),
                    debugInfo: {
                        errorLocation: CommonUtility.getObjectPathExpression(objectPath)
                    }
                });
                objectPath = objectPath.parentObjectPath;
            }
        }, CommonUtility.validateReferencedObjectPaths = function(objectPaths) {
            if (objectPaths) for (var i = 0; i < objectPaths.length; i++) for (var objectPath = objectPaths[i]; objectPath; ) {
                if (!objectPath.isValid) throw new Core._Internal.RuntimeError({
                    code: Core.CoreErrorCodes.invalidObjectPath,
                    httpStatusCode: 400,
                    message: Core.CoreUtility._getResourceString(Core.CoreResourceStrings.invalidObjectPath, CommonUtility.getObjectPathExpression(objectPath))
                });
                objectPath = objectPath.parentObjectPath;
            }
        }, CommonUtility._toCamelLowerCase = function(name) {
            if (Core.CoreUtility.isNullOrEmptyString(name)) return name;
            for (var index = 0; index < name.length && name.charCodeAt(index) >= 65 && name.charCodeAt(index) <= 90; ) index++;
            return index < name.length ? name.substr(0, index).toLowerCase() + name.substr(index) : name.toLowerCase();
        }, CommonUtility.adjustToDateTime = function(value) {
            if (Core.CoreUtility.isNullOrUndefined(value)) return null;
            if ("string" == typeof value) return new Date(value);
            if (Array.isArray(value)) {
                for (var arr = value, i = 0; i < arr.length; i++) arr[i] = CommonUtility.adjustToDateTime(arr[i]);
                return arr;
            }
            throw Core.CoreUtility._createInvalidArgError({
                argumentName: "date"
            });
        }, CommonUtility.tryGetObjectIdFromLoadOrRetrieveResult = function(value) {
            var id = value[CommonConstants.id];
            return Core.CoreUtility.isNullOrUndefined(id) && (id = value[CommonConstants.idLowerCase]), 
            Core.CoreUtility.isNullOrUndefined(id) && (id = value[CommonConstants.idPrivate]), 
            id;
        }, CommonUtility.getObjectPathExpression = function(objectPath) {
            for (var ret = ""; objectPath; ) {
                switch (objectPath.objectPathInfo.ObjectPathType) {
                  case 1:
                    ret = ret;
                    break;

                  case 2:
                    ret = "new()" + (ret.length > 0 ? "." : "") + ret;
                    break;

                  case 3:
                    ret = CommonUtility.normalizeName(objectPath.objectPathInfo.Name) + "()" + (ret.length > 0 ? "." : "") + ret;
                    break;

                  case 4:
                    ret = CommonUtility.normalizeName(objectPath.objectPathInfo.Name) + (ret.length > 0 ? "." : "") + ret;
                    break;

                  case 5:
                    ret = "getItem()" + (ret.length > 0 ? "." : "") + ret;
                    break;

                  case 6:
                    ret = "_reference()" + (ret.length > 0 ? "." : "") + ret;
                }
                objectPath = objectPath.parentObjectPath;
            }
            return ret;
        }, CommonUtility.setMethodArguments = function(context, argumentInfo, args) {
            if (Core.CoreUtility.isNullOrUndefined(args)) return null;
            var referencedObjectPaths = new Array, referencedObjectPathIds = new Array, hasOne = CommonUtility.collectObjectPathInfos(context, args, referencedObjectPaths, referencedObjectPathIds);
            return argumentInfo.Arguments = args, hasOne && (argumentInfo.ReferencedObjectPathIds = referencedObjectPathIds), 
            referencedObjectPaths;
        }, CommonUtility.validateContext = function(context, obj) {
            if (context && obj && obj._context !== context) throw new Core._Internal.RuntimeError({
                code: Core.CoreErrorCodes.invalidRequestContext,
                httpStatusCode: 400,
                message: Core.CoreUtility._getResourceString(Core.CoreResourceStrings.invalidRequestContext)
            });
        }, CommonUtility.isSetSupported = function(apiSetName, apiSetVersion) {
            return !("undefined" != typeof window && window.Office && window.Office.context && window.Office.context.requirements) || window.Office.context.requirements.isSetSupported(apiSetName, apiSetVersion);
        }, CommonUtility.throwIfApiNotSupported = function(apiFullName, apiSetName, apiSetVersion, hostName) {
            if (CommonUtility._doApiNotSupportedCheck && !CommonUtility.isSetSupported(apiSetName, apiSetVersion)) {
                var message = Core.CoreUtility._getResourceString(Core.CoreResourceStrings.apiNotFoundDetails, [ apiFullName, apiSetName + " " + apiSetVersion, hostName ]);
                throw new Core._Internal.RuntimeError({
                    code: Core.CoreErrorCodes.apiNotFound,
                    httpStatusCode: 404,
                    message: message,
                    debugInfo: {
                        errorLocation: apiFullName
                    }
                });
            }
        }, CommonUtility.calculateApiFlags = function(apiFlags, undoableApiSetName, undoableApiSetVersion) {
            return CommonUtility.isSetSupported(undoableApiSetName, undoableApiSetVersion) || (apiFlags &= -3), 
            apiFlags;
        }, CommonUtility._parseSelectExpand = function(select) {
            var args = [];
            if (!Core.CoreUtility.isNullOrEmptyString(select)) for (var propertyNames = select.split(","), i = 0; i < propertyNames.length; i++) {
                var propertyName = propertyNames[i];
                (propertyName = sanitizeForAnyItemsSlash(propertyName.trim())).length > 0 && args.push(propertyName);
            }
            return args;
            function sanitizeForAnyItemsSlash(propertyName) {
                var propertyNameLower = propertyName.toLowerCase();
                if ("items" === propertyNameLower || "items/" === propertyNameLower) return "*";
                return ("items/" === propertyNameLower.substr(0, 6) || "items." === propertyNameLower.substr(0, 6)) && (propertyName = propertyName.substr(6)), 
                propertyName.replace(new RegExp("[/.]items[/.]", "gi"), "/");
            }
        }, CommonUtility.changePropertyNameToCamelLowerCase = function(value) {
            if (Array.isArray(value)) {
                for (var ret = [], i = 0; i < value.length; i++) ret.push(this.changePropertyNameToCamelLowerCase(value[i]));
                return ret;
            }
            if ("object" == typeof value && null !== value) {
                ret = {};
                for (var key in value) {
                    var propValue = value[key];
                    if (key === CommonConstants.items) {
                        (ret = {})[CommonConstants.itemsLowerCase] = this.changePropertyNameToCamelLowerCase(propValue);
                        break;
                    }
                    ret[CommonUtility._toCamelLowerCase(key)] = this.changePropertyNameToCamelLowerCase(propValue);
                }
                return ret;
            }
            return value;
        }, CommonUtility.purifyJson = function(value) {
            if (Array.isArray(value)) {
                for (var ret = [], i = 0; i < value.length; i++) ret.push(this.purifyJson(value[i]));
                return ret;
            }
            if ("object" == typeof value && null !== value) {
                ret = {};
                for (var key in value) if (95 !== key.charCodeAt(0)) {
                    var propValue = value[key];
                    "object" == typeof propValue && null !== propValue && Array.isArray(propValue.items) && (propValue = propValue.items), 
                    ret[key] = this.purifyJson(propValue);
                }
                return ret;
            }
            return value;
        }, CommonUtility.collectObjectPathInfos = function(context, args, referencedObjectPaths, referencedObjectPathIds) {
            for (var hasOne = !1, i = 0; i < args.length; i++) if (args[i] instanceof ClientObjectBase) {
                var clientObject = args[i];
                CommonUtility.validateContext(context, clientObject), args[i] = clientObject._objectPath.objectPathInfo.Id, 
                referencedObjectPathIds.push(clientObject._objectPath.objectPathInfo.Id), referencedObjectPaths.push(clientObject._objectPath), 
                hasOne = !0;
            } else if (Array.isArray(args[i])) {
                var childArrayObjectPathIds = new Array;
                CommonUtility.collectObjectPathInfos(context, args[i], referencedObjectPaths, childArrayObjectPathIds) ? (referencedObjectPathIds.push(childArrayObjectPathIds), 
                hasOne = !0) : referencedObjectPathIds.push(0);
            } else Core.CoreUtility.isPlainJsonObject(args[i]) ? (referencedObjectPathIds.push(0), 
            CommonUtility.replaceClientObjectPropertiesWithObjectPathIds(args[i], referencedObjectPaths)) : referencedObjectPathIds.push(0);
            return hasOne;
        }, CommonUtility.replaceClientObjectPropertiesWithObjectPathIds = function(value, referencedObjectPaths) {
            var _a, _b;
            for (var key in value) {
                var propValue = value[key];
                if (propValue instanceof ClientObjectBase) referencedObjectPaths.push(propValue._objectPath), 
                value[key] = ((_a = {})[CommonConstants.objectPathIdPrivate] = propValue._objectPath.objectPathInfo.Id, 
                _a); else if (Array.isArray(propValue)) for (var i = 0; i < propValue.length; i++) if (propValue[i] instanceof ClientObjectBase) {
                    var elem = propValue[i];
                    referencedObjectPaths.push(elem._objectPath), propValue[i] = ((_b = {})[CommonConstants.objectPathIdPrivate] = elem._objectPath.objectPathInfo.Id, 
                    _b);
                } else Core.CoreUtility.isPlainJsonObject(propValue[i]) && CommonUtility.replaceClientObjectPropertiesWithObjectPathIds(propValue[i], referencedObjectPaths); else Core.CoreUtility.isPlainJsonObject(propValue) && CommonUtility.replaceClientObjectPropertiesWithObjectPathIds(propValue, referencedObjectPaths);
            }
        }, CommonUtility.normalizeName = function(name) {
            return name.substr(0, 1).toLowerCase() + name.substr(1);
        }, CommonUtility._doApiNotSupportedCheck = !1, CommonUtility;
    }(Core.CoreUtility);
    exports.CommonUtility = CommonUtility;
    var CommonResourceStrings = function(_super) {
        function CommonResourceStrings() {
            return null !== _super && _super.apply(this, arguments) || this;
        }
        return __extends(CommonResourceStrings, _super), CommonResourceStrings.propertyDoesNotExist = "PropertyDoesNotExist", 
        CommonResourceStrings.attemptingToSetReadOnlyProperty = "AttemptingToSetReadOnlyProperty", 
        CommonResourceStrings;
    }(Core.CoreResourceStrings);
    exports.CommonResourceStrings = CommonResourceStrings;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    }), exports.LibraryBuilder = exports.BatchApiHelper = exports._internalConfig = exports.Utility = exports.SessionBase = exports.ResourceStrings = exports.HttpUtility = exports.GenericEventHandlers = exports.EventHandlers = exports.ErrorCodes = exports.Error = exports.CoreUtility = exports.Constants = exports.config = exports.CommonUtility = exports.ClientObject = exports.ClientResult = exports.ClientRequestContext = void 0;
    var core_1 = __webpack_require__(0);
    Object.defineProperty(exports, "CoreUtility", {
        enumerable: !0,
        get: function() {
            return core_1.CoreUtility;
        }
    }), Object.defineProperty(exports, "Error", {
        enumerable: !0,
        get: function() {
            return core_1.Error;
        }
    }), Object.defineProperty(exports, "HttpUtility", {
        enumerable: !0,
        get: function() {
            return core_1.HttpUtility;
        }
    }), Object.defineProperty(exports, "SessionBase", {
        enumerable: !0,
        get: function() {
            return core_1.SessionBase;
        }
    });
    var common_1 = __webpack_require__(1);
    Object.defineProperty(exports, "CommonUtility", {
        enumerable: !0,
        get: function() {
            return common_1.CommonUtility;
        }
    }), Object.defineProperty(exports, "ClientResult", {
        enumerable: !0,
        get: function() {
            return common_1.ClientResult;
        }
    });
    var batch_runtime_1 = __webpack_require__(4);
    Object.defineProperty(exports, "ClientRequestContext", {
        enumerable: !0,
        get: function() {
            return batch_runtime_1.ClientRequestContext;
        }
    }), Object.defineProperty(exports, "ClientObject", {
        enumerable: !0,
        get: function() {
            return batch_runtime_1.ClientObject;
        }
    }), Object.defineProperty(exports, "config", {
        enumerable: !0,
        get: function() {
            return batch_runtime_1.config;
        }
    }), Object.defineProperty(exports, "Constants", {
        enumerable: !0,
        get: function() {
            return batch_runtime_1.Constants;
        }
    }), Object.defineProperty(exports, "ErrorCodes", {
        enumerable: !0,
        get: function() {
            return batch_runtime_1.ErrorCodes;
        }
    }), Object.defineProperty(exports, "EventHandlers", {
        enumerable: !0,
        get: function() {
            return batch_runtime_1.EventHandlers;
        }
    }), Object.defineProperty(exports, "GenericEventHandlers", {
        enumerable: !0,
        get: function() {
            return batch_runtime_1.GenericEventHandlers;
        }
    }), Object.defineProperty(exports, "ResourceStrings", {
        enumerable: !0,
        get: function() {
            return batch_runtime_1.ResourceStrings;
        }
    }), Object.defineProperty(exports, "Utility", {
        enumerable: !0,
        get: function() {
            return batch_runtime_1.Utility;
        }
    }), Object.defineProperty(exports, "_internalConfig", {
        enumerable: !0,
        get: function() {
            return batch_runtime_1._internalConfig;
        }
    });
    var BatchApiHelper = function() {
        function BatchApiHelper() {}
        return BatchApiHelper.invokeMethod = function(obj, methodName, operationType, args, flags, resultProcessType) {
            var action = batch_runtime_1.ActionFactory.createMethodAction(obj.context, obj, methodName, operationType, args, flags), result = new common_1.ClientResult(resultProcessType);
            return batch_runtime_1.Utility._addActionResultHandler(obj, action, result), result;
        }, BatchApiHelper.invokeEnsureUnchanged = function(obj, objectState) {
            batch_runtime_1.ActionFactory.createEnsureUnchangedAction(obj.context, obj, objectState);
        }, BatchApiHelper.invokeSetProperty = function(obj, propName, propValue, flags) {
            batch_runtime_1.ActionFactory.createSetPropertyAction(obj.context, obj, propName, propValue, flags);
        }, BatchApiHelper.createRootServiceObject = function(type, context) {
            return new type(context, batch_runtime_1.ObjectPathFactory.createGlobalObjectObjectPath(context));
        }, BatchApiHelper.createObjectFromReferenceId = function(type, context, referenceId) {
            return new type(context, batch_runtime_1.ObjectPathFactory.createReferenceIdObjectPath(context, referenceId));
        }, BatchApiHelper.createTopLevelServiceObject = function(type, context, typeName, isCollection, flags) {
            return new type(context, batch_runtime_1.ObjectPathFactory.createNewObjectObjectPath(context, typeName, isCollection, flags));
        }, BatchApiHelper.createPropertyObject = function(type, parent, propertyName, isCollection, flags) {
            var objectPath = batch_runtime_1.ObjectPathFactory.createPropertyObjectPath(parent.context, parent, propertyName, isCollection, !1, flags);
            return new type(parent.context, objectPath);
        }, BatchApiHelper.createIndexerObject = function(type, parent, args) {
            var objectPath = batch_runtime_1.ObjectPathFactory.createIndexerObjectPath(parent.context, parent, args);
            return new type(parent.context, objectPath);
        }, BatchApiHelper.createMethodObject = function(type, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, flags) {
            var objectPath = batch_runtime_1.ObjectPathFactory.createMethodObjectPath(parent.context, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, flags);
            return new type(parent.context, objectPath);
        }, BatchApiHelper.createChildItemObject = function(type, hasIndexerMethod, parent, chileItem, index) {
            var objectPath = batch_runtime_1.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt(hasIndexerMethod, parent.context, parent, chileItem, index);
            return new type(parent.context, objectPath);
        }, BatchApiHelper;
    }();
    exports.BatchApiHelper = BatchApiHelper;
    var LibraryBuilder = function() {
        function LibraryBuilder(options) {
            if (this.m_namespaceMap = {}, this.m_namespace = options.metadata.name, this.m_targetNamespaceObject = options.targetNamespaceObject, 
            this.m_namespaceMap[this.m_namespace] = options.targetNamespaceObject, options.namespaceMap) for (var ns in options.namespaceMap) this.m_namespaceMap[ns] = options.namespaceMap[ns];
            this.m_defaultApiSetName = options.metadata.defaultApiSetName, this.m_hostName = options.metadata.hostName;
            var metadata = options.metadata;
            if (metadata.enumTypes) for (var i = 0; i < metadata.enumTypes.length; i++) this.buildEnumType(metadata.enumTypes[i]);
            if (metadata.apiSets) {
                for (i = 0; i < metadata.apiSets.length; i++) {
                    var elem = metadata.apiSets[i];
                    Array.isArray(elem) && (metadata.apiSets[i] = {
                        version: elem[0],
                        name: elem[1] || this.m_defaultApiSetName
                    });
                }
                this.m_apiSets = metadata.apiSets;
            }
            if (this.m_strings = metadata.strings, metadata.clientObjectTypes) for (i = 0; i < metadata.clientObjectTypes.length; i++) {
                elem = metadata.clientObjectTypes[i];
                Array.isArray(elem) && (this.ensureArraySize(elem, 11), metadata.clientObjectTypes[i] = {
                    name: this.getString(elem[0]),
                    behaviorFlags: elem[1],
                    collectionPropertyPath: this.getString(elem[6]),
                    newObjectServerTypeFullName: this.getString(elem[9]),
                    newObjectApiFlags: elem[10],
                    childItemTypeFullName: this.getString(elem[7]),
                    scalarProperties: elem[2],
                    navigationProperties: elem[3],
                    scalarMethods: elem[4],
                    navigationMethods: elem[5],
                    events: elem[8]
                }), this.buildClientObjectType(metadata.clientObjectTypes[i], options.fullyInitialize);
            }
        }
        return LibraryBuilder.prototype.ensureArraySize = function(value, size) {
            for (var count = size - value.length; count > 0; ) value.push(0), count--;
        }, LibraryBuilder.prototype.getString = function(ordinalOrValue) {
            return "number" == typeof ordinalOrValue ? ordinalOrValue > 0 ? this.m_strings[ordinalOrValue - 1] : null : ordinalOrValue;
        }, LibraryBuilder.prototype.buildEnumType = function(elem) {
            var enumType;
            if (Array.isArray(elem)) {
                (enumType = {
                    name: elem[0],
                    fields: elem[2]
                }).fields || (enumType.fields = {});
                var fieldsWithCamelUpperCaseValue = elem[1];
                if (Array.isArray(fieldsWithCamelUpperCaseValue)) for (var index = 0; index < fieldsWithCamelUpperCaseValue.length; index++) enumType.fields[fieldsWithCamelUpperCaseValue[index]] = this.toSimpleCamelUpperCase(fieldsWithCamelUpperCaseValue[index]);
            } else enumType = elem;
            this.m_targetNamespaceObject[enumType.name] = enumType.fields;
        }, LibraryBuilder.prototype.buildClientObjectType = function(typeInfo, fullyInitialize) {
            var thisBuilder = this, type = function(context, objectPath) {
                batch_runtime_1.ClientObject.apply(this, arguments), thisBuilder.m_targetNamespaceObject[typeInfo.name]._typeInited || (thisBuilder.buildPrototype(thisBuilder.m_targetNamespaceObject[typeInfo.name], typeInfo), 
                thisBuilder.m_targetNamespaceObject[typeInfo.name]._typeInited = !0), batch_runtime_1._internalConfig.appendTypeNameToObjectPathInfo && this._objectPath && this._objectPath.objectPathInfo && this._className && (this._objectPath.objectPathInfo.T = this._className);
            };
            (this.m_targetNamespaceObject[typeInfo.name] = type, this.extendsType(type, batch_runtime_1.ClientObject), 
            this.buildNewObject(type, typeInfo), 0 != (2 & typeInfo.behaviorFlags) && (type.prototype._KeepReference = function() {
                BatchApiHelper.invokeMethod(this, "_KeepReference", 1, [], 0, 0);
            }), 0 != (32 & typeInfo.behaviorFlags)) && this.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_StaticCustomize").call(null, type);
            fullyInitialize && (this.buildPrototype(type, typeInfo), type._typeInited = !0);
        }, LibraryBuilder.prototype.extendsType = function(d, b) {
            function __() {
                this.constructor = d;
            }
            d.prototype = null === b ? Object.create(b) : (__.prototype = b.prototype, new __);
        }, LibraryBuilder.prototype.findObjectUnderPath = function(top, paths, pathStartIndex) {
            for (var obj = top, i = pathStartIndex; i < paths.length; i++) {
                if ("object" != typeof obj) throw new core_1.Error("Cannot find " + paths.join("."));
                obj = obj[paths[i]];
            }
            return obj;
        }, LibraryBuilder.prototype.getFunction = function(fullName) {
            var ret = this.resolveObjectByFullName(fullName);
            if ("function" != typeof ret) throw new core_1.Error("Cannot find function or type: " + fullName);
            return ret;
        }, LibraryBuilder.prototype.resolveObjectByFullName = function(fullName) {
            var parts = fullName.split(".");
            if (1 === parts.length) return this.m_targetNamespaceObject[parts[0]];
            var rootName = parts[0];
            return rootName === this.m_namespace ? this.findObjectUnderPath(this.m_targetNamespaceObject, parts, 1) : this.m_namespaceMap[rootName] ? this.findObjectUnderPath(this.m_namespaceMap[rootName], parts, 1) : this.findObjectUnderPath(this.m_targetNamespaceObject, parts, 0);
        }, LibraryBuilder.prototype.evaluateSimpleExpression = function(expression, thisObj) {
            if (batch_runtime_1.Utility.isNullOrUndefined(expression)) return null;
            var paths = expression.split(".");
            if (3 === paths.length && "OfficeExtension" === paths[0] && "Constants" === paths[1]) return batch_runtime_1.Constants[paths[2]];
            if ("this" === paths[0]) {
                for (var obj = thisObj, i = 1; i < paths.length; i++) obj = "toString()" == paths[i] ? obj.toString() : "()" === paths[i].substr(paths[i].length - 2) ? obj[paths[i].substr(0, paths[i].length - 2)]() : obj[paths[i]];
                return obj;
            }
            throw new core_1.Error("Cannot evaluate: " + expression);
        }, LibraryBuilder.prototype.evaluateEventTargetId = function(targetIdExpression, thisObj) {
            return batch_runtime_1.Utility.isNullOrEmptyString(targetIdExpression) ? "" : this.evaluateSimpleExpression(targetIdExpression, thisObj);
        }, LibraryBuilder.prototype.isAllDigits = function(expression) {
            for (var charZero = "0".charCodeAt(0), charNine = "9".charCodeAt(0), i = 0; i < expression.length; i++) if (expression.charCodeAt(i) < charZero || expression.charCodeAt(i) > charNine) return !1;
            return !0;
        }, LibraryBuilder.prototype.evaluateEventType = function(eventTypeExpression) {
            if (batch_runtime_1.Utility.isNullOrEmptyString(eventTypeExpression)) return 0;
            if (this.isAllDigits(eventTypeExpression)) return parseInt(eventTypeExpression);
            var ret = this.resolveObjectByFullName(eventTypeExpression);
            if ("number" != typeof ret) throw new core_1.Error("Invalid event type: " + eventTypeExpression);
            return ret;
        }, LibraryBuilder.prototype.buildPrototype = function(type, typeInfo) {
            this.buildScalarProperties(type, typeInfo), this.buildNavigationProperties(type, typeInfo), 
            this.buildScalarMethods(type, typeInfo), this.buildNavigationMethods(type, typeInfo), 
            this.buildEvents(type, typeInfo), this.buildHandleResult(type, typeInfo), this.buildHandleIdResult(type, typeInfo), 
            this.buildHandleRetrieveResult(type, typeInfo), this.buildLoad(type, typeInfo), 
            this.buildRetrieve(type, typeInfo), this.buildSetMockData(type, typeInfo), this.buildEnsureUnchanged(type, typeInfo), 
            this.buildUpdate(type, typeInfo), this.buildSet(type, typeInfo), this.buildToJSON(type, typeInfo), 
            this.buildItems(type, typeInfo), this.buildTypeMetadataInfo(type, typeInfo), this.buildTrackUntrack(type, typeInfo), 
            this.buildMixin(type, typeInfo);
        }, LibraryBuilder.prototype.toSimpleCamelUpperCase = function(name) {
            return name.substr(0, 1).toUpperCase() + name.substr(1);
        }, LibraryBuilder.prototype.ensureOriginalName = function(member) {
            null === member.originalName && (member.originalName = this.toSimpleCamelUpperCase(member.name));
        }, LibraryBuilder.prototype.getFieldName = function(member) {
            return "m_" + member.name;
        }, LibraryBuilder.prototype.throwIfApiNotSupported = function(typeInfo, member) {
            if (this.m_apiSets && member.apiSetInfoOrdinal > 0) {
                var apiSetInfo = this.m_apiSets[member.apiSetInfoOrdinal - 1];
                apiSetInfo && batch_runtime_1.Utility.throwIfApiNotSupported(typeInfo.name + "." + member.name, apiSetInfo.name, apiSetInfo.version, this.m_hostName);
            }
        }, LibraryBuilder.prototype.buildScalarProperties = function(type, typeInfo) {
            if (Array.isArray(typeInfo.scalarProperties)) for (var i = 0; i < typeInfo.scalarProperties.length; i++) {
                var elem = typeInfo.scalarProperties[i];
                Array.isArray(elem) && (this.ensureArraySize(elem, 6), typeInfo.scalarProperties[i] = {
                    name: this.getString(elem[0]),
                    behaviorFlags: elem[1],
                    apiSetInfoOrdinal: elem[2],
                    originalName: this.getString(elem[3]),
                    setMethodApiFlags: elem[4],
                    undoableApiSetInfoOrdinal: elem[5]
                }), this.buildScalarProperty(type, typeInfo, typeInfo.scalarProperties[i]);
            }
        }, LibraryBuilder.prototype.calculateApiFlags = function(apiFlags, undoableApiSetInfoOrdinal) {
            if (undoableApiSetInfoOrdinal > 0) {
                var undoableApiSetInfo = this.m_apiSets[undoableApiSetInfoOrdinal - 1];
                undoableApiSetInfo && (apiFlags = common_1.CommonUtility.calculateApiFlags(apiFlags, undoableApiSetInfo.name, undoableApiSetInfo.version));
            }
            return apiFlags;
        }, LibraryBuilder.prototype.buildScalarProperty = function(type, typeInfo, propInfo) {
            this.ensureOriginalName(propInfo);
            var thisBuilder = this, fieldName = this.getFieldName(propInfo), descriptor = {
                get: function() {
                    return batch_runtime_1.Utility.throwIfNotLoaded(propInfo.name, this[fieldName], typeInfo.name, this._isNull), 
                    thisBuilder.throwIfApiNotSupported(typeInfo, propInfo), this[fieldName];
                },
                enumerable: !0,
                configurable: !0
            };
            0 == (2 & propInfo.behaviorFlags) && (descriptor.set = function(value) {
                if (4 & propInfo.behaviorFlags && thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_" + propInfo.originalName + "_Set").call(this, this, value).handled) return;
                this[fieldName] = value;
                var apiFlags = thisBuilder.calculateApiFlags(propInfo.setMethodApiFlags, propInfo.undoableApiSetInfoOrdinal);
                BatchApiHelper.invokeSetProperty(this, propInfo.originalName, value, apiFlags);
            }), Object.defineProperty(type.prototype, propInfo.name, descriptor);
        }, LibraryBuilder.prototype.buildNavigationProperties = function(type, typeInfo) {
            if (Array.isArray(typeInfo.navigationProperties)) for (var i = 0; i < typeInfo.navigationProperties.length; i++) {
                var elem = typeInfo.navigationProperties[i];
                Array.isArray(elem) && (this.ensureArraySize(elem, 8), typeInfo.navigationProperties[i] = {
                    name: this.getString(elem[0]),
                    behaviorFlags: elem[2],
                    apiSetInfoOrdinal: elem[3],
                    originalName: this.getString(elem[4]),
                    getMethodApiFlags: elem[5],
                    setMethodApiFlags: elem[6],
                    propertyTypeFullName: this.getString(elem[1]),
                    undoableApiSetInfoOrdinal: elem[7]
                }), this.buildNavigationProperty(type, typeInfo, typeInfo.navigationProperties[i]);
            }
        }, LibraryBuilder.prototype.buildNavigationProperty = function(type, typeInfo, propInfo) {
            this.ensureOriginalName(propInfo);
            var thisBuilder = this, fieldName = this.getFieldName(propInfo), descriptor = {
                get: function() {
                    (this[thisBuilder.getFieldName(propInfo)] || (thisBuilder.throwIfApiNotSupported(typeInfo, propInfo), 
                    this[fieldName] = BatchApiHelper.createPropertyObject(thisBuilder.getFunction(propInfo.propertyTypeFullName), this, propInfo.originalName, 0 != (16 & propInfo.behaviorFlags), propInfo.getMethodApiFlags)), 
                    64 & propInfo.behaviorFlags) && thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_" + propInfo.originalName + "_Get").call(this, this, this[fieldName]);
                    return this[fieldName];
                },
                enumerable: !0,
                configurable: !0
            };
            0 == (2 & propInfo.behaviorFlags) && (descriptor.set = function(value) {
                if (4 & propInfo.behaviorFlags && thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_" + propInfo.originalName + "_Set").call(this, this, value).handled) return;
                this[fieldName] = value;
                var apiFlags = thisBuilder.calculateApiFlags(propInfo.setMethodApiFlags, propInfo.undoableApiSetInfoOrdinal);
                BatchApiHelper.invokeSetProperty(this, propInfo.originalName, value, apiFlags);
            }), Object.defineProperty(type.prototype, propInfo.name, descriptor);
        }, LibraryBuilder.prototype.buildScalarMethods = function(type, typeInfo) {
            if (Array.isArray(typeInfo.scalarMethods)) for (var i = 0; i < typeInfo.scalarMethods.length; i++) {
                var elem = typeInfo.scalarMethods[i];
                Array.isArray(elem) && (this.ensureArraySize(elem, 7), typeInfo.scalarMethods[i] = {
                    name: this.getString(elem[0]),
                    behaviorFlags: elem[2],
                    apiSetInfoOrdinal: elem[3],
                    originalName: this.getString(elem[5]),
                    apiFlags: elem[4],
                    parameterCount: elem[1],
                    undoableApiSetInfoOrdinal: elem[6]
                }), this.buildScalarMethod(type, typeInfo, typeInfo.scalarMethods[i]);
            }
        }, LibraryBuilder.prototype.buildScalarMethod = function(type, typeInfo, methodInfo) {
            this.ensureOriginalName(methodInfo);
            var thisBuilder = this;
            type.prototype[methodInfo.name] = function() {
                var args = [];
                if (64 & methodInfo.behaviorFlags && methodInfo.parameterCount > 0) {
                    for (var i = 0; i < methodInfo.parameterCount - 1; i++) args.push(arguments[i]);
                    var rest = [];
                    for (i = methodInfo.parameterCount - 1; i < arguments.length; i++) rest.push(arguments[i]);
                    args.push(rest);
                } else for (i = 0; i < arguments.length; i++) args.push(arguments[i]);
                if (1 & methodInfo.behaviorFlags) {
                    var customizationFunc = thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_" + methodInfo.originalName), applyArgs = [ this ];
                    for (i = 0; i < args.length; i++) applyArgs.push(args[i]);
                    var _a = customizationFunc.apply(this, applyArgs), handled = _a.handled, result = _a.result;
                    if (handled) return result;
                }
                thisBuilder.throwIfApiNotSupported(typeInfo, methodInfo);
                var resultProcessType = 0;
                32 & methodInfo.behaviorFlags && (resultProcessType = 1);
                var operationType = 0;
                2 & methodInfo.behaviorFlags && (operationType = 1);
                var apiFlags = thisBuilder.calculateApiFlags(methodInfo.apiFlags, methodInfo.undoableApiSetInfoOrdinal);
                return BatchApiHelper.invokeMethod(this, methodInfo.originalName, operationType, args, apiFlags, resultProcessType);
            };
        }, LibraryBuilder.prototype.buildNavigationMethods = function(type, typeInfo) {
            if (Array.isArray(typeInfo.navigationMethods)) for (var i = 0; i < typeInfo.navigationMethods.length; i++) {
                var elem = typeInfo.navigationMethods[i];
                Array.isArray(elem) && (this.ensureArraySize(elem, 9), typeInfo.navigationMethods[i] = {
                    name: this.getString(elem[0]),
                    behaviorFlags: elem[3],
                    apiSetInfoOrdinal: elem[4],
                    originalName: this.getString(elem[6]),
                    apiFlags: elem[5],
                    parameterCount: elem[2],
                    returnTypeFullName: this.getString(elem[1]),
                    returnObjectGetByIdMethodName: this.getString(elem[7]),
                    undoableApiSetInfoOrdinal: elem[8]
                }), this.buildNavigationMethod(type, typeInfo, typeInfo.navigationMethods[i]);
            }
        }, LibraryBuilder.prototype.buildNavigationMethod = function(type, typeInfo, methodInfo) {
            this.ensureOriginalName(methodInfo);
            var thisBuilder = this;
            type.prototype[methodInfo.name] = function() {
                var args = [];
                if (64 & methodInfo.behaviorFlags && methodInfo.parameterCount > 0) {
                    for (var i = 0; i < methodInfo.parameterCount - 1; i++) args.push(arguments[i]);
                    var rest = [];
                    for (i = methodInfo.parameterCount - 1; i < arguments.length; i++) rest.push(arguments[i]);
                    args.push(rest);
                } else for (i = 0; i < arguments.length; i++) args.push(arguments[i]);
                if (1 & methodInfo.behaviorFlags) {
                    var customizationFunc = thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_" + methodInfo.originalName), applyArgs = [ this ];
                    for (i = 0; i < args.length; i++) applyArgs.push(args[i]);
                    var _a = customizationFunc.apply(this, applyArgs), handled = _a.handled, result = _a.result;
                    if (handled) return result;
                }
                if (thisBuilder.throwIfApiNotSupported(typeInfo, methodInfo), 0 != (16 & methodInfo.behaviorFlags)) return BatchApiHelper.createIndexerObject(thisBuilder.getFunction(methodInfo.returnTypeFullName), this, args);
                var operationType = 0;
                2 & methodInfo.behaviorFlags && (operationType = 1);
                var apiFlags = thisBuilder.calculateApiFlags(methodInfo.apiFlags, methodInfo.undoableApiSetInfoOrdinal);
                return BatchApiHelper.createMethodObject(thisBuilder.getFunction(methodInfo.returnTypeFullName), this, methodInfo.originalName, operationType, args, 0 != (4 & methodInfo.behaviorFlags), 0 != (8 & methodInfo.behaviorFlags), methodInfo.returnObjectGetByIdMethodName, apiFlags);
            };
        }, LibraryBuilder.prototype.buildHandleResult = function(type, typeInfo) {
            var thisBuilder = this;
            type.prototype._handleResult = function(value) {
                if (batch_runtime_1.ClientObject.prototype._handleResult.call(this, value), !batch_runtime_1.Utility.isNullOrUndefined(value)) {
                    if (batch_runtime_1.Utility.fixObjectPathIfNecessary(this, value), 8 & typeInfo.behaviorFlags) thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_HandleResult").call(this, this, value);
                    if (typeInfo.scalarProperties) for (var i_1 = 0; i_1 < typeInfo.scalarProperties.length; i_1++) batch_runtime_1.Utility.isUndefined(value[typeInfo.scalarProperties[i_1].originalName]) || (0 != (8 & typeInfo.scalarProperties[i_1].behaviorFlags) ? this[thisBuilder.getFieldName(typeInfo.scalarProperties[i_1])] = batch_runtime_1.Utility.adjustToDateTime(value[typeInfo.scalarProperties[i_1].originalName]) : this[thisBuilder.getFieldName(typeInfo.scalarProperties[i_1])] = value[typeInfo.scalarProperties[i_1].originalName]);
                    if (typeInfo.navigationProperties) {
                        for (var propNames = [], i_2 = 0; i_2 < typeInfo.navigationProperties.length; i_2++) propNames.push(typeInfo.navigationProperties[i_2].name), 
                        propNames.push(typeInfo.navigationProperties[i_2].originalName);
                        batch_runtime_1.Utility._handleNavigationPropertyResults(this, value, propNames);
                    }
                    if (0 != (1 & typeInfo.behaviorFlags)) {
                        var hasIndexerMethod = thisBuilder.hasIndexMethod(typeInfo);
                        if (!batch_runtime_1.Utility.isNullOrUndefined(value[batch_runtime_1.Constants.items])) {
                            this.m__items = [];
                            for (var _data = value[batch_runtime_1.Constants.items], childItemType = thisBuilder.getFunction(typeInfo.childItemTypeFullName), i = 0; i < _data.length; i++) {
                                var _item = BatchApiHelper.createChildItemObject(childItemType, hasIndexerMethod, this, _data[i], i);
                                _item._handleResult(_data[i]), this.m__items.push(_item);
                            }
                        }
                    }
                }
            };
        }, LibraryBuilder.prototype.buildHandleRetrieveResult = function(type, typeInfo) {
            var thisBuilder = this;
            type.prototype._handleRetrieveResult = function(value, result) {
                if (batch_runtime_1.ClientObject.prototype._handleRetrieveResult.call(this, value, result), 
                !batch_runtime_1.Utility.isNullOrUndefined(value)) {
                    if (typeInfo.scalarProperties) for (var i = 0; i < typeInfo.scalarProperties.length; i++) 8 & typeInfo.scalarProperties[i].behaviorFlags && (batch_runtime_1.Utility.isNullOrUndefined(value[typeInfo.scalarProperties[i].name]) || (value[typeInfo.scalarProperties[i].name] = batch_runtime_1.Utility.adjustToDateTime(value[typeInfo.scalarProperties[i].name])));
                    if (1 & typeInfo.behaviorFlags) {
                        var hasIndexerMethod_1 = thisBuilder.hasIndexMethod(typeInfo), childItemType_1 = thisBuilder.getFunction(typeInfo.childItemTypeFullName), thisObj_1 = this;
                        batch_runtime_1.Utility.processRetrieveResult(thisObj_1, value, result, (function(childItemData, index) {
                            return BatchApiHelper.createChildItemObject(childItemType_1, hasIndexerMethod_1, thisObj_1, childItemData, index);
                        }));
                    } else batch_runtime_1.Utility.processRetrieveResult(this, value, result);
                }
            };
        }, LibraryBuilder.prototype.buildHandleIdResult = function(type, typeInfo) {
            var thisBuilder = this;
            type.prototype._handleIdResult = function(value) {
                if (batch_runtime_1.ClientObject.prototype._handleIdResult.call(this, value), !batch_runtime_1.Utility.isNullOrUndefined(value)) {
                    if (16 & typeInfo.behaviorFlags) thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_HandleIdResult").call(this, this, value);
                    if (typeInfo.scalarProperties) for (var i = 0; i < typeInfo.scalarProperties.length; i++) {
                        var propName = typeInfo.scalarProperties[i].originalName;
                        "Id" !== propName && "_Id" !== propName && "_ReferenceId" !== propName || batch_runtime_1.Utility.isNullOrUndefined(value[typeInfo.scalarProperties[i].originalName]) || (this[thisBuilder.getFieldName(typeInfo.scalarProperties[i])] = value[typeInfo.scalarProperties[i].originalName]);
                    }
                }
            };
        }, LibraryBuilder.prototype.buildLoad = function(type, typeInfo) {
            type.prototype.load = function(options) {
                return batch_runtime_1.Utility.load(this, options);
            };
        }, LibraryBuilder.prototype.buildRetrieve = function(type, typeInfo) {
            type.prototype.retrieve = function(options) {
                return batch_runtime_1.Utility.retrieve(this, options);
            };
        }, LibraryBuilder.prototype.buildNewObject = function(type, typeInfo) {
            batch_runtime_1.Utility.isNullOrEmptyString(typeInfo.newObjectServerTypeFullName) || (type.newObject = function(context) {
                return BatchApiHelper.createTopLevelServiceObject(type, context, typeInfo.newObjectServerTypeFullName, 0 != (1 & typeInfo.behaviorFlags), typeInfo.newObjectApiFlags);
            });
        }, LibraryBuilder.prototype.buildSetMockData = function(type, typeInfo) {
            var thisBuilder = this;
            if (1 & typeInfo.behaviorFlags) {
                var hasIndexMethod_1 = thisBuilder.hasIndexMethod(typeInfo);
                type.prototype.setMockData = function(data) {
                    var thisObj = this;
                    batch_runtime_1.Utility.setMockData(thisObj, data, (function(childItemData, index) {
                        return BatchApiHelper.createChildItemObject(thisBuilder.getFunction(typeInfo.childItemTypeFullName), hasIndexMethod_1, thisObj, childItemData, index);
                    }), (function(items) {
                        thisObj.m__items = items;
                    }));
                };
            } else type.prototype.setMockData = function(data) {
                batch_runtime_1.Utility.setMockData(this, data);
            };
        }, LibraryBuilder.prototype.buildEnsureUnchanged = function(type, typeInfo) {
            type.prototype.ensureUnchanged = function(data) {
                BatchApiHelper.invokeEnsureUnchanged(this, data);
            };
        }, LibraryBuilder.prototype.buildUpdate = function(type, typeInfo) {
            type.prototype.update = function(properties) {
                this._recursivelyUpdate(properties);
            };
        }, LibraryBuilder.prototype.buildSet = function(type, typeInfo) {
            if (0 == (1 & typeInfo.behaviorFlags)) {
                var notAllowedToBeSetPropertyNames = [], allowedScalarPropertyNames = [];
                if (typeInfo.scalarProperties) for (var i = 0; i < typeInfo.scalarProperties.length; i++) 0 == (2 & typeInfo.scalarProperties[i].behaviorFlags) && 0 != (1 & typeInfo.scalarProperties[i].behaviorFlags) ? allowedScalarPropertyNames.push(typeInfo.scalarProperties[i].name) : notAllowedToBeSetPropertyNames.push(typeInfo.scalarProperties[i].name);
                var allowedNavigationPropertyNames = [];
                if (typeInfo.navigationProperties) for (i = 0; i < typeInfo.navigationProperties.length; i++) 0 != (16 & typeInfo.navigationProperties[i].behaviorFlags) || 0 == (1 & typeInfo.navigationProperties[i].behaviorFlags) || 0 == (32 & typeInfo.navigationProperties[i].behaviorFlags) ? notAllowedToBeSetPropertyNames.push(typeInfo.navigationProperties[i].name) : allowedNavigationPropertyNames.push(typeInfo.navigationProperties[i].name);
                0 === allowedNavigationPropertyNames.length && 0 === allowedScalarPropertyNames.length || (type.prototype.set = function(properties, options) {
                    this._recursivelySet(properties, options, allowedScalarPropertyNames, allowedNavigationPropertyNames, notAllowedToBeSetPropertyNames);
                });
            }
        }, LibraryBuilder.prototype.buildItems = function(type, typeInfo) {
            0 != (1 & typeInfo.behaviorFlags) && Object.defineProperty(type.prototype, "items", {
                get: function() {
                    return batch_runtime_1.Utility.throwIfNotLoaded("items", this.m__items, typeInfo.name, this._isNull), 
                    this.m__items;
                },
                enumerable: !0,
                configurable: !0
            });
        }, LibraryBuilder.prototype.buildToJSON = function(type, typeInfo) {
            var thisBuilder = this;
            0 == (1 & typeInfo.behaviorFlags) ? type.prototype.toJSON = function() {
                var scalarProperties = {};
                if (typeInfo.scalarProperties) for (var i = 0; i < typeInfo.scalarProperties.length; i++) 0 != (1 & typeInfo.scalarProperties[i].behaviorFlags) && (scalarProperties[typeInfo.scalarProperties[i].name] = this[thisBuilder.getFieldName(typeInfo.scalarProperties[i])]);
                var navProperties = {};
                if (typeInfo.navigationProperties) for (i = 0; i < typeInfo.navigationProperties.length; i++) 0 != (1 & typeInfo.navigationProperties[i].behaviorFlags) && (navProperties[typeInfo.navigationProperties[i].name] = this[thisBuilder.getFieldName(typeInfo.navigationProperties[i])]);
                return batch_runtime_1.Utility.toJson(this, scalarProperties, navProperties);
            } : type.prototype.toJSON = function() {
                return batch_runtime_1.Utility.toJson(this, {}, {}, this.m__items);
            };
        }, LibraryBuilder.prototype.buildTypeMetadataInfo = function(type, typeInfo) {
            Object.defineProperty(type.prototype, "_className", {
                get: function() {
                    return typeInfo.name;
                },
                enumerable: !0,
                configurable: !0
            }), Object.defineProperty(type.prototype, "_isCollection", {
                get: function() {
                    return 0 != (1 & typeInfo.behaviorFlags);
                },
                enumerable: !0,
                configurable: !0
            }), batch_runtime_1.Utility.isNullOrEmptyString(typeInfo.collectionPropertyPath) || Object.defineProperty(type.prototype, "_collectionPropertyPath", {
                get: function() {
                    return typeInfo.collectionPropertyPath;
                },
                enumerable: !0,
                configurable: !0
            }), typeInfo.scalarProperties && typeInfo.scalarProperties.length > 0 && (Object.defineProperty(type.prototype, "_scalarPropertyNames", {
                get: function() {
                    return this.m__scalarPropertyNames || (this.m__scalarPropertyNames = typeInfo.scalarProperties.map((function(p) {
                        return p.name;
                    }))), this.m__scalarPropertyNames;
                },
                enumerable: !0,
                configurable: !0
            }), Object.defineProperty(type.prototype, "_scalarPropertyOriginalNames", {
                get: function() {
                    return this.m__scalarPropertyOriginalNames || (this.m__scalarPropertyOriginalNames = typeInfo.scalarProperties.map((function(p) {
                        return p.originalName;
                    }))), this.m__scalarPropertyOriginalNames;
                },
                enumerable: !0,
                configurable: !0
            }), Object.defineProperty(type.prototype, "_scalarPropertyUpdateable", {
                get: function() {
                    return this.m__scalarPropertyUpdateable || (this.m__scalarPropertyUpdateable = typeInfo.scalarProperties.map((function(p) {
                        return 0 == (2 & p.behaviorFlags);
                    }))), this.m__scalarPropertyUpdateable;
                },
                enumerable: !0,
                configurable: !0
            })), typeInfo.navigationProperties && typeInfo.navigationProperties.length > 0 && Object.defineProperty(type.prototype, "_navigationPropertyNames", {
                get: function() {
                    return this.m__navigationPropertyNames || (this.m__navigationPropertyNames = typeInfo.navigationProperties.map((function(p) {
                        return p.name;
                    }))), this.m__navigationPropertyNames;
                },
                enumerable: !0,
                configurable: !0
            });
        }, LibraryBuilder.prototype.buildTrackUntrack = function(type, typeInfo) {
            2 & typeInfo.behaviorFlags && (type.prototype.track = function() {
                return this.context.trackedObjects.add(this), this;
            }, type.prototype.untrack = function() {
                return this.context.trackedObjects.remove(this), this;
            });
        }, LibraryBuilder.prototype.buildMixin = function(type, typeInfo) {
            if (4 & typeInfo.behaviorFlags) {
                var mixinType = this.getFunction(typeInfo.name + "Custom");
                batch_runtime_1.Utility.applyMixin(type, mixinType);
            }
        }, LibraryBuilder.prototype.getOnEventName = function(name) {
            return "_" === name[0] ? "_on" + name.substr(1) : "on" + name;
        }, LibraryBuilder.prototype.buildEvents = function(type, typeInfo) {
            if (typeInfo.events) for (var i = 0; i < typeInfo.events.length; i++) {
                var elem = typeInfo.events[i];
                Array.isArray(elem) && (this.ensureArraySize(elem, 7), typeInfo.events[i] = {
                    name: this.getString(elem[0]),
                    behaviorFlags: elem[1],
                    apiSetInfoOrdinal: elem[2],
                    typeExpression: this.getString(elem[3]),
                    targetIdExpression: this.getString(elem[4]),
                    register: this.getString(elem[5]),
                    unregister: this.getString(elem[6])
                }), this.buildEvent(type, typeInfo, typeInfo.events[i]);
            }
        }, LibraryBuilder.prototype.buildEvent = function(type, typeInfo, evt) {
            1 & evt.behaviorFlags ? this.buildV0Event(type, typeInfo, evt) : this.buildV2Event(type, typeInfo, evt);
        }, LibraryBuilder.prototype.buildV2Event = function(type, typeInfo, evt) {
            var thisBuilder = this, eventName = this.getOnEventName(evt.name), fieldName = this.getFieldName(evt);
            Object.defineProperty(type.prototype, eventName, {
                get: function() {
                    if (!this[fieldName]) {
                        thisBuilder.throwIfApiNotSupported(typeInfo, evt);
                        var thisObj = this, registerFunc = null;
                        "null" !== evt.register && (registerFunc = this[evt.register].bind(this));
                        var unregisterFunc = null;
                        "null" !== evt.unregister && (unregisterFunc = this[evt.unregister].bind(this));
                        var func = null;
                        2 & evt.behaviorFlags && (func = thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_" + evt.name + "_EventArgsTransform"));
                        var eventType = thisBuilder.evaluateEventType(evt.typeExpression);
                        this[fieldName] = new batch_runtime_1.GenericEventHandlers(this.context, this, evt.name, {
                            eventType: eventType,
                            getTargetIdFunc: function() {
                                return thisBuilder.evaluateEventTargetId(evt.targetIdExpression, thisObj);
                            },
                            registerFunc: registerFunc,
                            unregisterFunc: unregisterFunc,
                            eventArgsTransformFunc: function(value) {
                                return func && (value = func.call(thisObj, thisObj, value)), batch_runtime_1.Utility._createPromiseFromResult(value);
                            }
                        });
                    }
                    return this[fieldName];
                },
                enumerable: !0,
                configurable: !0
            });
        }, LibraryBuilder.prototype.buildV0Event = function(type, typeInfo, evt) {
            var thisBuilder = this, eventName = this.getOnEventName(evt.name), fieldName = this.getFieldName(evt);
            Object.defineProperty(type.prototype, eventName, {
                get: function() {
                    if (!this[fieldName]) {
                        thisBuilder.throwIfApiNotSupported(typeInfo, evt);
                        var thisObj = this, registerFunc = null;
                        if (batch_runtime_1.Utility.isNullOrEmptyString(evt.register)) {
                            var eventType_1 = thisBuilder.evaluateEventType(evt.typeExpression);
                            registerFunc = function(handlerCallback) {
                                var targetId = thisBuilder.evaluateEventTargetId(evt.targetIdExpression, thisObj);
                                return thisObj.context.eventRegistration.register(eventType_1, targetId, handlerCallback);
                            };
                        } else if ("null" !== evt.register) {
                            var func_1 = thisBuilder.getFunction(evt.register);
                            registerFunc = function(handlerCallback) {
                                return func_1.call(thisObj, thisObj, handlerCallback);
                            };
                        }
                        var unregisterFunc = null;
                        if (batch_runtime_1.Utility.isNullOrEmptyString(evt.unregister)) {
                            var eventType_2 = thisBuilder.evaluateEventType(evt.typeExpression);
                            unregisterFunc = function(handlerCallback) {
                                var targetId = thisBuilder.evaluateEventTargetId(evt.targetIdExpression, thisObj);
                                return thisObj.context.eventRegistration.unregister(eventType_2, targetId, handlerCallback);
                            };
                        } else if ("null" !== evt.unregister) {
                            var func_2 = thisBuilder.getFunction(evt.unregister);
                            unregisterFunc = function(handlerCallback) {
                                return func_2.call(thisObj, thisObj, handlerCallback);
                            };
                        }
                        var func = null;
                        2 & evt.behaviorFlags && (func = thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_" + evt.name + "_EventArgsTransform"));
                        this[fieldName] = new batch_runtime_1.EventHandlers(this.context, this, evt.name, {
                            registerFunc: registerFunc,
                            unregisterFunc: unregisterFunc,
                            eventArgsTransformFunc: function(value) {
                                return func && (value = func.call(thisObj, thisObj, value)), batch_runtime_1.Utility._createPromiseFromResult(value);
                            }
                        });
                    }
                    return this[fieldName];
                },
                enumerable: !0,
                configurable: !0
            });
        }, LibraryBuilder.prototype.hasIndexMethod = function(typeInfo) {
            var ret = !1;
            if (typeInfo.navigationMethods) for (var i = 0; i < typeInfo.navigationMethods.length; i++) if (0 != (16 & typeInfo.navigationMethods[i].behaviorFlags)) {
                ret = !0;
                break;
            }
            return ret;
        }, LibraryBuilder.CustomizationCodeNamespace = "_CC", LibraryBuilder;
    }();
    exports.LibraryBuilder = LibraryBuilder;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var OfficeExtensionBatch = __webpack_require__(2);
    __webpack_require__(5), __webpack_require__(6), window.OfficeExtensionBatch = OfficeExtensionBatch, 
    "undefined" == typeof CustomFunctionMappings && (window.CustomFunctionMappings = {}), 
    "undefined" == typeof Promise && (window.Promise = Office.Promise), window.OfficeExtension = {
        Promise: Promise,
        Error: OfficeExtensionBatch.Error,
        ErrorCodes: OfficeExtensionBatch.ErrorCodes
    }, __webpack_require__(7).default(!0);
}, function(module, exports, __webpack_require__) {
    "use strict";
    var extendStatics, __extends = this && this.__extends || (extendStatics = function(d, b) {
        return (extendStatics = Object.setPrototypeOf || {
            __proto__: []
        } instanceof Array && function(d, b) {
            d.__proto__ = b;
        } || function(d, b) {
            for (var p in b) b.hasOwnProperty(p) && (d[p] = b[p]);
        })(d, b);
    }, function(d, b) {
        function __() {
            this.constructor = d;
        }
        extendStatics(d, b), d.prototype = null === b ? Object.create(b) : (__.prototype = b.prototype, 
        new __);
    }), __createBinding = this && this.__createBinding || (Object.create ? function(o, m, k, k2) {
        void 0 === k2 && (k2 = k), Object.defineProperty(o, k2, {
            enumerable: !0,
            get: function() {
                return m[k];
            }
        });
    } : function(o, m, k, k2) {
        void 0 === k2 && (k2 = k), o[k2] = m[k];
    }), __exportStar = this && this.__exportStar || function(m, exports) {
        for (var p in m) "default" === p || exports.hasOwnProperty(p) || __createBinding(exports, m, p);
    };
    Object.defineProperty(exports, "__esModule", {
        value: !0
    }), exports.Utility = exports.ResourceStrings = exports.TrackedObjects = exports.ObjectPathFactory = exports.GenericEventHandlers = exports._testTriggerRichApiMessageEvent = exports._testSetRichApiMessageReadyImpl = exports.GenericEventRegistration = exports.EventRegistration = exports._Internal = exports.EventHandlerResult = exports.EventHandlers = exports.ClientRequest = exports.Constants = exports.ClientRequestContext = exports.HostBridgeSession = exports.ClientObject = exports.ActionFactory = exports.ErrorCodes = void 0;
    var Core = __webpack_require__(0), Common = __webpack_require__(1);
    __exportStar(__webpack_require__(1), exports);
    var ErrorCodes = function(_super) {
        function ErrorCodes() {
            return null !== _super && _super.apply(this, arguments) || this;
        }
        return __extends(ErrorCodes, _super), ErrorCodes.propertyNotLoaded = "PropertyNotLoaded", 
        ErrorCodes.runMustReturnPromise = "RunMustReturnPromise", ErrorCodes.cannotRegisterEvent = "CannotRegisterEvent", 
        ErrorCodes.invalidOrTimedOutSession = "InvalidOrTimedOutSession", ErrorCodes.cannotUpdateReadOnlyProperty = "CannotUpdateReadOnlyProperty", 
        ErrorCodes;
    }(Core.CoreErrorCodes);
    exports.ErrorCodes = ErrorCodes;
    var TraceMarkerActionResultHandler = function() {
        function TraceMarkerActionResultHandler(callback) {
            this.m_callback = callback;
        }
        return TraceMarkerActionResultHandler.prototype._handleResult = function(value) {
            this.m_callback && this.m_callback();
        }, TraceMarkerActionResultHandler;
    }(), ActionFactory = function(_super) {
        function ActionFactory() {
            return null !== _super && _super.apply(this, arguments) || this;
        }
        return __extends(ActionFactory, _super), ActionFactory.createMethodAction = function(context, parent, methodName, operationType, args, flags) {
            Utility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 3,
                Name: methodName,
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            }, referencedArgumentObjectPaths = Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
            Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
            var action = new Common.Action(actionInfo, operationType, Utility._fixupApiFlags(flags));
            return action.referencedObjectPath = parent._objectPath, action.referencedArgumentObjectPaths = referencedArgumentObjectPaths, 
            parent._addAction(action), action;
        }, ActionFactory.createRecursiveQueryAction = function(context, parent, query) {
            Utility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 6,
                Name: "",
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                RecursiveQueryInfo: query
            }, action = new Common.Action(actionInfo, 1, 4);
            return action.referencedObjectPath = parent._objectPath, parent._addAction(action), 
            action;
        }, ActionFactory.createEnsureUnchangedAction = function(context, parent, objectState) {
            Utility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 8,
                Name: "",
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                ObjectState: objectState
            }, action = new Common.Action(actionInfo, 1, 4);
            return action.referencedObjectPath = parent._objectPath, parent._addAction(action), 
            action;
        }, ActionFactory.createInstantiateAction = function(context, obj) {
            Utility.validateObjectPath(obj), context._pendingRequest.ensureInstantiateObjectPath(obj._objectPath.parentObjectPath), 
            context._pendingRequest.ensureInstantiateObjectPaths(obj._objectPath.argumentObjectPaths);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 1,
                Name: "",
                ObjectPathId: obj._objectPath.objectPathInfo.Id
            }, action = new Common.Action(actionInfo, 1, 4);
            return action.referencedObjectPath = obj._objectPath, obj._addAction(action, new InstantiateActionResultHandler(obj), !0), 
            action;
        }, ActionFactory.createTraceAction = function(context, message, addTraceMessage) {
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 5,
                Name: "Trace",
                ObjectPathId: 0
            }, ret = new Common.Action(actionInfo, 1, 4);
            return context._pendingRequest.addAction(ret), addTraceMessage && context._pendingRequest.addTrace(actionInfo.Id, message), 
            ret;
        }, ActionFactory.createTraceMarkerForCallback = function(context, callback) {
            var action = ActionFactory.createTraceAction(context, null, !1);
            context._pendingRequest.addActionResultHandler(action, new TraceMarkerActionResultHandler(callback));
        }, ActionFactory;
    }(Common.CommonActionFactory);
    exports.ActionFactory = ActionFactory;
    var ClientObject = function(_super) {
        function ClientObject(context, objectPath) {
            var _this = _super.call(this, context, objectPath) || this;
            return Utility.checkArgumentNull(context, "context"), _this.m_context = context, 
            _this._objectPath && (!context._processingResult && context._pendingRequest && (ActionFactory.createInstantiateAction(context, _this), 
            context._autoCleanup && _this._KeepReference && context.trackedObjects._autoAdd(_this)), 
            Common._internalConfig.appendTypeNameToObjectPathInfo && _this._objectPath.objectPathInfo && _this._className && (_this._objectPath.objectPathInfo.T = _this._className)), 
            _this;
        }
        return __extends(ClientObject, _super), Object.defineProperty(ClientObject.prototype, "context", {
            get: function() {
                return this.m_context;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ClientObject.prototype, "isNull", {
            get: function() {
                return (void 0 !== this.m_isNull || !Core.TestUtility.isMock()) && (Utility.throwIfNotLoaded("isNull", this._isNull, null, this._isNull), 
                this._isNull);
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ClientObject.prototype, "isNullObject", {
            get: function() {
                return (void 0 !== this.m_isNull || !Core.TestUtility.isMock()) && (Utility.throwIfNotLoaded("isNullObject", this._isNull, null, this._isNull), 
                this._isNull);
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ClientObject.prototype, "_isNull", {
            get: function() {
                return this.m_isNull;
            },
            set: function(value) {
                this.m_isNull = value, value && this._objectPath && this._objectPath._updateAsNullObject();
            },
            enumerable: !1,
            configurable: !0
        }), ClientObject.prototype._addAction = function(action, resultHandler, isInstantiationEnsured) {
            return void 0 === resultHandler && (resultHandler = null), isInstantiationEnsured || (this.context._pendingRequest.ensureInstantiateObjectPath(this._objectPath), 
            this.context._pendingRequest.ensureInstantiateObjectPaths(action.referencedArgumentObjectPaths)), 
            this.context._pendingRequest.addAction(action), this.context._pendingRequest.addReferencedObjectPath(this._objectPath), 
            this.context._pendingRequest.addReferencedObjectPaths(action.referencedArgumentObjectPaths), 
            this.context._pendingRequest.addActionResultHandler(action, resultHandler), Core.CoreUtility._createPromiseFromResult(null);
        }, ClientObject.prototype._handleResult = function(value) {
            this._isNull = Utility.isNullOrUndefined(value), this.context.trackedObjects._autoTrackIfNecessaryWhenHandleObjectResultValue(this, value);
        }, ClientObject.prototype._handleIdResult = function(value) {
            this._isNull = Utility.isNullOrUndefined(value), Utility.fixObjectPathIfNecessary(this, value), 
            this.context.trackedObjects._autoTrackIfNecessaryWhenHandleObjectResultValue(this, value);
        }, ClientObject.prototype._handleRetrieveResult = function(value, result) {
            this._handleIdResult(value);
        }, ClientObject.prototype._recursivelySet = function(input, options, scalarWriteablePropertyNames, objectPropertyNames, notAllowedToBeSetPropertyNames) {
            var isClientObject = input instanceof ClientObject, originalInput = input;
            if (isClientObject) {
                if (Object.getPrototypeOf(this) !== Object.getPrototypeOf(input)) throw Core._Internal.RuntimeError._createInvalidArgError({
                    argumentName: "properties",
                    errorLocation: this._className + ".set"
                });
                input = JSON.parse(JSON.stringify(input));
            }
            try {
                for (var prop, i = 0; i < scalarWriteablePropertyNames.length; i++) prop = scalarWriteablePropertyNames[i], 
                input.hasOwnProperty(prop) && void 0 !== input[prop] && (this[prop] = input[prop]);
                for (i = 0; i < objectPropertyNames.length; i++) if (prop = objectPropertyNames[i], 
                input.hasOwnProperty(prop) && void 0 !== input[prop]) {
                    var dataToPassToSet = isClientObject ? originalInput[prop] : input[prop];
                    this[prop].set(dataToPassToSet, options);
                }
                var throwOnReadOnly = !isClientObject;
                options && !Utility.isNullOrUndefined(throwOnReadOnly) && (throwOnReadOnly = options.throwOnReadOnly);
                for (i = 0; i < notAllowedToBeSetPropertyNames.length; i++) if (prop = notAllowedToBeSetPropertyNames[i], 
                input.hasOwnProperty(prop) && void 0 !== input[prop] && throwOnReadOnly) throw new Core._Internal.RuntimeError({
                    code: Core.CoreErrorCodes.invalidArgument,
                    httpStatusCode: 400,
                    message: Core.CoreUtility._getResourceString(ResourceStrings.cannotApplyPropertyThroughSetMethod, prop),
                    debugInfo: {
                        errorLocation: prop
                    }
                });
                for (prop in input) if (scalarWriteablePropertyNames.indexOf(prop) < 0 && objectPropertyNames.indexOf(prop) < 0) {
                    var propertyDescriptor = Object.getOwnPropertyDescriptor(Object.getPrototypeOf(this), prop);
                    if (!propertyDescriptor) throw new Core._Internal.RuntimeError({
                        code: Core.CoreErrorCodes.invalidArgument,
                        httpStatusCode: 400,
                        message: Core.CoreUtility._getResourceString(Common.CommonResourceStrings.propertyDoesNotExist, prop),
                        debugInfo: {
                            errorLocation: prop
                        }
                    });
                    if (throwOnReadOnly && !propertyDescriptor.set) throw new Core._Internal.RuntimeError({
                        code: Core.CoreErrorCodes.invalidArgument,
                        httpStatusCode: 400,
                        message: Core.CoreUtility._getResourceString(Common.CommonResourceStrings.attemptingToSetReadOnlyProperty, prop),
                        debugInfo: {
                            errorLocation: prop
                        }
                    });
                }
            } catch (innerError) {
                throw new Core._Internal.RuntimeError({
                    code: Core.CoreErrorCodes.invalidArgument,
                    httpStatusCode: 400,
                    message: Core.CoreUtility._getResourceString(Core.CoreResourceStrings.invalidArgument, "properties"),
                    debugInfo: {
                        errorLocation: this._className + ".set"
                    },
                    innerError: innerError
                });
            }
        }, ClientObject;
    }(Common.ClientObjectBase);
    exports.ClientObject = ClientObject;
    var HostBridgeRequestExecutor = function() {
        function HostBridgeRequestExecutor(session) {
            this.m_session = session;
        }
        return HostBridgeRequestExecutor.prototype.executeAsync = function(customData, requestFlags, requestMessage) {
            var httpRequestInfo = {
                url: Core.CoreConstants.processQuery,
                method: "POST",
                headers: requestMessage.Headers,
                body: requestMessage.Body
            }, message = {
                id: Core.HostBridge.nextId(),
                type: 1,
                flags: requestFlags,
                message: httpRequestInfo
            };
            return Core.CoreUtility.log(JSON.stringify(message)), this.m_session.sendMessageToHost(message).then((function(nativeBridgeResponse) {
                Core.CoreUtility.log("Received response: " + JSON.stringify(nativeBridgeResponse));
                var response, responseInfo = nativeBridgeResponse.message;
                if (200 === responseInfo.statusCode) response = {
                    HttpStatusCode: responseInfo.statusCode,
                    ErrorCode: null,
                    ErrorMessage: null,
                    Headers: responseInfo.headers,
                    Body: Core.CoreUtility._parseResponseBody(responseInfo)
                }; else {
                    Core.CoreUtility.log("Error Response:" + responseInfo.body);
                    var error = Core.CoreUtility._parseErrorResponse(responseInfo);
                    response = {
                        HttpStatusCode: responseInfo.statusCode,
                        ErrorCode: error.errorCode,
                        ErrorMessage: error.errorMessage,
                        Headers: responseInfo.headers,
                        Body: null
                    };
                }
                return response;
            }));
        }, HostBridgeRequestExecutor;
    }(), HostBridgeSession = function(_super) {
        function HostBridgeSession(m_bridge) {
            var _this = _super.call(this) || this;
            return _this.m_bridge = m_bridge, _this.m_bridge.addHostMessageHandler((function(message) {
                3 === message.type && GenericEventRegistration.getGenericEventRegistration()._handleRichApiMessage(message.message);
            })), _this;
        }
        return __extends(HostBridgeSession, _super), HostBridgeSession.getInstanceIfHostBridgeInited = function() {
            return Core.HostBridge.instance ? ((Core.CoreUtility.isNullOrUndefined(HostBridgeSession.s_instance) || HostBridgeSession.s_instance.m_bridge !== Core.HostBridge.instance) && (HostBridgeSession.s_instance = new HostBridgeSession(Core.HostBridge.instance)), 
            HostBridgeSession.s_instance) : null;
        }, HostBridgeSession.prototype._resolveRequestUrlAndHeaderInfo = function() {
            return Core.CoreUtility._createPromiseFromResult(null);
        }, HostBridgeSession.prototype._createRequestExecutorOrNull = function() {
            return Core.CoreUtility.log("NativeBridgeSession::CreateRequestExecutor"), new HostBridgeRequestExecutor(this);
        }, Object.defineProperty(HostBridgeSession.prototype, "eventRegistration", {
            get: function() {
                return GenericEventRegistration.getGenericEventRegistration();
            },
            enumerable: !1,
            configurable: !0
        }), HostBridgeSession.prototype.sendMessageToHost = function(message) {
            return this.m_bridge.sendMessageToHostAndExpectResponse(message);
        }, HostBridgeSession;
    }(Core.SessionBase);
    exports.HostBridgeSession = HostBridgeSession;
    var ClientRequestContext = function(_super) {
        function ClientRequestContext(url) {
            var _this = _super.call(this) || this;
            if (_this.m_customRequestHeaders = {}, _this.m_batchMode = 0, _this._onRunFinishedNotifiers = [], 
            Core.SessionBase._overrideSession) _this.m_requestUrlAndHeaderInfoResolver = Core.SessionBase._overrideSession; else if ((Utility.isNullOrUndefined(url) || "string" == typeof url && 0 === url.length) && ((url = ClientRequestContext.defaultRequestUrlAndHeaders) || (url = {
                url: Core.CoreConstants.localDocument,
                headers: {}
            })), "string" == typeof url) _this.m_requestUrlAndHeaderInfo = {
                url: url,
                headers: {}
            }; else if (ClientRequestContext.isRequestUrlAndHeaderInfoResolver(url)) _this.m_requestUrlAndHeaderInfoResolver = url; else {
                if (!ClientRequestContext.isRequestUrlAndHeaderInfo(url)) throw Core._Internal.RuntimeError._createInvalidArgError({
                    argumentName: "url"
                });
                var requestInfo = url;
                _this.m_requestUrlAndHeaderInfo = {
                    url: requestInfo.url,
                    headers: {}
                }, Core.CoreUtility._copyHeaders(requestInfo.headers, _this.m_requestUrlAndHeaderInfo.headers);
            }
            return !_this.m_requestUrlAndHeaderInfoResolver && _this.m_requestUrlAndHeaderInfo && Core.CoreUtility._isLocalDocumentUrl(_this.m_requestUrlAndHeaderInfo.url) && HostBridgeSession.getInstanceIfHostBridgeInited() && (_this.m_requestUrlAndHeaderInfo = null, 
            _this.m_requestUrlAndHeaderInfoResolver = HostBridgeSession.getInstanceIfHostBridgeInited()), 
            _this.m_requestUrlAndHeaderInfoResolver instanceof Core.SessionBase && (_this.m_session = _this.m_requestUrlAndHeaderInfoResolver), 
            _this._processingResult = !1, _this._customData = Constants.iterativeExecutor, _this.sync = _this.sync.bind(_this), 
            _this;
        }
        return __extends(ClientRequestContext, _super), Object.defineProperty(ClientRequestContext.prototype, "session", {
            get: function() {
                return this.m_session;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ClientRequestContext.prototype, "eventRegistration", {
            get: function() {
                return this.m_session ? this.m_session.eventRegistration : _Internal.officeJsEventRegistration;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ClientRequestContext.prototype, "_url", {
            get: function() {
                return this.m_requestUrlAndHeaderInfo ? this.m_requestUrlAndHeaderInfo.url : null;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ClientRequestContext.prototype, "_pendingRequest", {
            get: function() {
                return null == this.m_pendingRequest && (this.m_pendingRequest = new ClientRequest(this)), 
                this.m_pendingRequest;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ClientRequestContext.prototype, "debugInfo", {
            get: function() {
                return {
                    pendingStatements: new RequestPrettyPrinter(this._rootObjectPropertyName, this._pendingRequest._objectPaths, this._pendingRequest._actions, Common._internalConfig.showDisposeInfoInDebugInfo).process()
                };
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ClientRequestContext.prototype, "trackedObjects", {
            get: function() {
                return this.m_trackedObjects || (this.m_trackedObjects = new TrackedObjects(this)), 
                this.m_trackedObjects;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ClientRequestContext.prototype, "requestHeaders", {
            get: function() {
                return this.m_customRequestHeaders;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ClientRequestContext.prototype, "batchMode", {
            get: function() {
                return this.m_batchMode;
            },
            enumerable: !1,
            configurable: !0
        }), ClientRequestContext.prototype.ensureInProgressBatchIfBatchMode = function() {
            if (1 === this.m_batchMode && !this.m_explicitBatchInProgress) throw Utility.createRuntimeError(Core.CoreErrorCodes.generalException, Core.CoreUtility._getResourceString(ResourceStrings.notInsideBatch), null);
        }, ClientRequestContext.prototype.load = function(clientObj, option) {
            Utility.validateContext(this, clientObj);
            var queryOption = ClientRequestContext._parseQueryOption(option);
            Common.CommonActionFactory.createQueryAction(this, clientObj, queryOption, clientObj);
        }, ClientRequestContext.prototype.loadRecursive = function(clientObj, options, maxDepth) {
            if (!Utility.isPlainJsonObject(options)) throw Core._Internal.RuntimeError._createInvalidArgError({
                argumentName: "options"
            });
            var quries = {};
            for (var key in options) quries[key] = ClientRequestContext._parseQueryOption(options[key]);
            var action = ActionFactory.createRecursiveQueryAction(this, clientObj, {
                Queries: quries,
                MaxDepth: maxDepth
            });
            this._pendingRequest.addActionResultHandler(action, clientObj);
        }, ClientRequestContext.prototype.trace = function(message) {
            ActionFactory.createTraceAction(this, message, !0);
        }, ClientRequestContext.prototype._processOfficeJsErrorResponse = function(officeJsErrorCode, response) {}, 
        ClientRequestContext.prototype.ensureRequestUrlAndHeaderInfo = function() {
            var _this = this;
            return Utility._createPromiseFromResult(null).then((function() {
                if (!_this.m_requestUrlAndHeaderInfo) return _this.m_requestUrlAndHeaderInfoResolver._resolveRequestUrlAndHeaderInfo().then((function(value) {
                    if (_this.m_requestUrlAndHeaderInfo = value, _this.m_requestUrlAndHeaderInfo || (_this.m_requestUrlAndHeaderInfo = {
                        url: Core.CoreConstants.localDocument,
                        headers: {}
                    }), Utility.isNullOrEmptyString(_this.m_requestUrlAndHeaderInfo.url) && (_this.m_requestUrlAndHeaderInfo.url = Core.CoreConstants.localDocument), 
                    _this.m_requestUrlAndHeaderInfo.headers || (_this.m_requestUrlAndHeaderInfo.headers = {}), 
                    "function" == typeof _this.m_requestUrlAndHeaderInfoResolver._createRequestExecutorOrNull) {
                        var executor = _this.m_requestUrlAndHeaderInfoResolver._createRequestExecutorOrNull();
                        executor && (_this._requestExecutor = executor);
                    }
                }));
            }));
        }, ClientRequestContext.prototype.syncPrivateMain = function() {
            var _this = this;
            return this.ensureRequestUrlAndHeaderInfo().then((function() {
                var req = _this._pendingRequest;
                return _this.m_pendingRequest = null, _this.processPreSyncPromises(req).then((function() {
                    return _this.syncPrivate(req);
                }));
            }));
        }, ClientRequestContext.prototype.syncPrivate = function(req) {
            var _this = this;
            if (Core.TestUtility.isMock()) return Core.CoreUtility._createPromiseFromResult(null);
            if (!req.hasActions) return this.processPendingEventHandlers(req);
            var _a = req.buildRequestMessageBodyAndRequestFlags(), msgBody = _a.body, requestFlags = _a.flags;
            this._requestFlagModifier && (requestFlags |= this._requestFlagModifier), this._requestExecutor || (Core.CoreUtility._isLocalDocumentUrl(this.m_requestUrlAndHeaderInfo.url) ? this._requestExecutor = new OfficeJsRequestExecutor(this) : this._requestExecutor = new Common.HttpRequestExecutor);
            var requestExecutor = this._requestExecutor, headers = {};
            Core.CoreUtility._copyHeaders(this.m_requestUrlAndHeaderInfo.headers, headers), 
            Core.CoreUtility._copyHeaders(this.m_customRequestHeaders, headers), delete this.m_customRequestHeaders[Constants.officeScriptEventId];
            var requestExecutorRequestMessage = {
                Url: this.m_requestUrlAndHeaderInfo.url,
                Headers: headers,
                Body: msgBody
            };
            req.invalidatePendingInvalidObjectPaths();
            var errorFromResponse = null, errorFromProcessEventHandlers = null;
            return this._lastSyncStart = "undefined" == typeof performance ? 0 : performance.now(), 
            this._lastRequestFlags = requestFlags, requestExecutor.executeAsync(this._customData, requestFlags, requestExecutorRequestMessage).then((function(response) {
                return _this._lastSyncEnd = "undefined" == typeof performance ? 0 : performance.now(), 
                errorFromResponse = _this.processRequestExecutorResponseMessage(req, response), 
                _this.processPendingEventHandlers(req).catch((function(ex) {
                    Core.CoreUtility.log("Error in processPendingEventHandlers"), Core.CoreUtility.log(JSON.stringify(ex)), 
                    errorFromProcessEventHandlers = ex;
                }));
            })).then((function() {
                if (errorFromResponse) throw Core.CoreUtility.log("Throw error from response: " + JSON.stringify(errorFromResponse)), 
                errorFromResponse;
                if (errorFromProcessEventHandlers) {
                    Core.CoreUtility.log("Throw error from ProcessEventHandler: " + JSON.stringify(errorFromProcessEventHandlers));
                    var transformedError = null;
                    if (errorFromProcessEventHandlers instanceof Core._Internal.RuntimeError) (transformedError = errorFromProcessEventHandlers).traceMessages = req._responseTraceMessages; else {
                        var message = null;
                        message = "string" == typeof errorFromProcessEventHandlers ? errorFromProcessEventHandlers : errorFromProcessEventHandlers.message, 
                        Utility.isNullOrEmptyString(message) && (message = Core.CoreUtility._getResourceString(ResourceStrings.cannotRegisterEvent)), 
                        transformedError = new Core._Internal.RuntimeError({
                            code: ErrorCodes.cannotRegisterEvent,
                            httpStatusCode: 400,
                            message: message,
                            traceMessages: req._responseTraceMessages
                        });
                    }
                    throw transformedError;
                }
            }));
        }, ClientRequestContext.prototype.processRequestExecutorResponseMessage = function(req, response) {
            response.Body && response.Body.TraceIds && req._setResponseTraceIds(response.Body.TraceIds);
            var traceMessages = req._responseTraceMessages, errorStatementInfo = null;
            if (response.Body) {
                if (response.Body.Error && response.Body.Error.ActionIndex >= 0) {
                    var prettyPrinter = new RequestPrettyPrinter(this._rootObjectPropertyName, req._objectPaths, req._actions, !1, !0), debugInfoStatementInfo = prettyPrinter.processForDebugStatementInfo(response.Body.Error.ActionIndex);
                    errorStatementInfo = {
                        statement: debugInfoStatementInfo.statement,
                        surroundingStatements: debugInfoStatementInfo.surroundingStatements,
                        fullStatements: [ "Please enable config.extendedErrorLogging to see full statements." ]
                    }, Common.config.extendedErrorLogging && (prettyPrinter = new RequestPrettyPrinter(this._rootObjectPropertyName, req._objectPaths, req._actions, !1, !1), 
                    errorStatementInfo.fullStatements = prettyPrinter.process());
                }
                var actionResults = null;
                if (response.Body.Results ? actionResults = response.Body.Results : response.Body.ProcessedResults && response.Body.ProcessedResults.Results && (actionResults = response.Body.ProcessedResults.Results), 
                actionResults) {
                    this._processingResult = !0;
                    try {
                        req.processResponse(actionResults);
                    } finally {
                        this._processingResult = !1;
                    }
                }
            }
            if (!Utility.isNullOrEmptyString(response.ErrorCode)) return new Core._Internal.RuntimeError({
                code: response.ErrorCode,
                httpStatusCode: response.HttpStatusCode,
                message: response.ErrorMessage,
                traceMessages: traceMessages
            });
            if (response.Body && response.Body.Error) {
                var debugInfo = {
                    errorLocation: response.Body.Error.Location
                };
                return errorStatementInfo && (debugInfo.statement = errorStatementInfo.statement, 
                debugInfo.surroundingStatements = errorStatementInfo.surroundingStatements, debugInfo.fullStatements = errorStatementInfo.fullStatements), 
                new Core._Internal.RuntimeError({
                    code: response.Body.Error.Code,
                    httpStatusCode: response.Body.Error.HttpStatusCode,
                    message: response.Body.Error.Message,
                    traceMessages: traceMessages,
                    debugInfo: debugInfo
                });
            }
            return null;
        }, ClientRequestContext.prototype.processPendingEventHandlers = function(req) {
            for (var ret = Utility._createPromiseFromResult(null), i = 0; i < req._pendingProcessEventHandlers.length; i++) {
                var eventHandlers = req._pendingProcessEventHandlers[i];
                ret = ret.then(this.createProcessOneEventHandlersFunc(eventHandlers, req));
            }
            return ret;
        }, ClientRequestContext.prototype.createProcessOneEventHandlersFunc = function(eventHandlers, req) {
            return function() {
                return eventHandlers._processRegistration(req);
            };
        }, ClientRequestContext.prototype.processPreSyncPromises = function(req) {
            for (var ret = Utility._createPromiseFromResult(null), i = 0; i < req._preSyncPromises.length; i++) {
                var p = req._preSyncPromises[i];
                ret = ret.then(this.createProcessOneProSyncFunc(p));
            }
            return ret;
        }, ClientRequestContext.prototype.createProcessOneProSyncFunc = function(p) {
            return function() {
                return p;
            };
        }, ClientRequestContext.prototype.sync = function(passThroughValue) {
            return Core.TestUtility.isMock() ? Core.CoreUtility._createPromiseFromResult(passThroughValue) : this.syncPrivateMain().then((function() {
                return passThroughValue;
            }));
        }, ClientRequestContext.prototype.batch = function(batchBody) {
            var _this = this;
            if (1 !== this.m_batchMode) return Core.CoreUtility._createPromiseFromException(Utility.createRuntimeError(Core.CoreErrorCodes.generalException, null, null));
            if (this.m_explicitBatchInProgress) return Core.CoreUtility._createPromiseFromException(Utility.createRuntimeError(Core.CoreErrorCodes.generalException, Core.CoreUtility._getResourceString(ResourceStrings.pendingBatchInProgress), null));
            if (Utility.isNullOrUndefined(batchBody)) return Utility._createPromiseFromResult(null);
            this.m_explicitBatchInProgress = !0;
            var batchBodyResult, request, batchBodyResultPromise, previousRequest = this.m_pendingRequest;
            this.m_pendingRequest = new ClientRequest(this);
            try {
                batchBodyResult = batchBody(this._rootObject, this);
            } catch (ex) {
                return this.m_explicitBatchInProgress = !1, this.m_pendingRequest = previousRequest, 
                Core.CoreUtility._createPromiseFromException(ex);
            }
            return "object" == typeof batchBodyResult && batchBodyResult && "function" == typeof batchBodyResult.then ? batchBodyResultPromise = Utility._createPromiseFromResult(null).then((function() {
                return batchBodyResult;
            })).then((function(result) {
                return _this.m_explicitBatchInProgress = !1, request = _this.m_pendingRequest, _this.m_pendingRequest = previousRequest, 
                result;
            })).catch((function(ex) {
                return _this.m_explicitBatchInProgress = !1, request = _this.m_pendingRequest, _this.m_pendingRequest = previousRequest, 
                Core.CoreUtility._createPromiseFromException(ex);
            })) : (this.m_explicitBatchInProgress = !1, request = this.m_pendingRequest, this.m_pendingRequest = previousRequest, 
            batchBodyResultPromise = Utility._createPromiseFromResult(batchBodyResult)), batchBodyResultPromise.then((function(result) {
                return _this.ensureRequestUrlAndHeaderInfo().then((function() {
                    return _this.syncPrivate(request);
                })).then((function() {
                    return result;
                }));
            }));
        }, ClientRequestContext._run = function(ctxInitializer, runBody, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
            return void 0 === numCleanupAttempts && (numCleanupAttempts = 3), void 0 === retryDelay && (retryDelay = 5e3), 
            ClientRequestContext._runCommon("run", null, ctxInitializer, 0, runBody, numCleanupAttempts, retryDelay, null, onCleanupSuccess, onCleanupFailure);
        }, ClientRequestContext.isValidRequestInfo = function(value) {
            return "string" == typeof value || ClientRequestContext.isRequestUrlAndHeaderInfo(value) || ClientRequestContext.isRequestUrlAndHeaderInfoResolver(value);
        }, ClientRequestContext.isRequestUrlAndHeaderInfo = function(value) {
            return "object" == typeof value && null !== value && Object.getPrototypeOf(value) === Object.getPrototypeOf({}) && !Utility.isNullOrUndefined(value.url);
        }, ClientRequestContext.isRequestUrlAndHeaderInfoResolver = function(value) {
            return "object" == typeof value && null !== value && "function" == typeof value._resolveRequestUrlAndHeaderInfo;
        }, ClientRequestContext._runBatch = function(functionName, receivedRunArgs, ctxInitializer, onBeforeRun, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
            return void 0 === numCleanupAttempts && (numCleanupAttempts = 3), void 0 === retryDelay && (retryDelay = 5e3), 
            ClientRequestContext._runBatchCommon(0, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onBeforeRun, onCleanupSuccess, onCleanupFailure);
        }, ClientRequestContext._runExplicitBatch = function(functionName, receivedRunArgs, ctxInitializer, onBeforeRun, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
            return void 0 === numCleanupAttempts && (numCleanupAttempts = 3), void 0 === retryDelay && (retryDelay = 5e3), 
            ClientRequestContext._runBatchCommon(1, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onBeforeRun, onCleanupSuccess, onCleanupFailure);
        }, ClientRequestContext._runBatchCommon = function(batchMode, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onBeforeRun, onCleanupSuccess, onCleanupFailure) {
            var ctxRetriever, batch;
            void 0 === numCleanupAttempts && (numCleanupAttempts = 3), void 0 === retryDelay && (retryDelay = 5e3);
            var requestInfo = null, previousObjects = null, argOffset = 0, options = null;
            if (receivedRunArgs.length > 0) if (ClientRequestContext.isValidRequestInfo(receivedRunArgs[0])) requestInfo = receivedRunArgs[0], 
            argOffset = 1; else if (Utility.isPlainJsonObject(receivedRunArgs[0])) {
                if (null != (requestInfo = (options = receivedRunArgs[0]).session) && !ClientRequestContext.isValidRequestInfo(requestInfo)) return ClientRequestContext.createErrorPromise(functionName);
                previousObjects = options.previousObjects, argOffset = 1;
            }
            if (receivedRunArgs.length == argOffset + 1) batch = receivedRunArgs[argOffset + 0]; else {
                if (null != options || receivedRunArgs.length != argOffset + 2) return ClientRequestContext.createErrorPromise(functionName);
                previousObjects = receivedRunArgs[argOffset + 0], batch = receivedRunArgs[argOffset + 1];
            }
            if (null != previousObjects) if (previousObjects instanceof ClientObject) ctxRetriever = function() {
                return previousObjects.context;
            }; else if (previousObjects instanceof ClientRequestContext) ctxRetriever = function() {
                return previousObjects;
            }; else {
                if (!Array.isArray(previousObjects)) return ClientRequestContext.createErrorPromise(functionName);
                var array = previousObjects;
                if (0 == array.length) return ClientRequestContext.createErrorPromise(functionName);
                for (var i = 0; i < array.length; i++) {
                    if (!(array[i] instanceof ClientObject)) return ClientRequestContext.createErrorPromise(functionName);
                    if (array[i].context != array[0].context) return ClientRequestContext.createErrorPromise(functionName, ResourceStrings.invalidRequestContext);
                }
                ctxRetriever = function() {
                    return array[0].context;
                };
            } else ctxRetriever = ctxInitializer;
            var onBeforeRunWithOptions = null;
            return onBeforeRun && (onBeforeRunWithOptions = function(context) {
                return onBeforeRun(options || {}, context);
            }), ClientRequestContext._runCommon(functionName, requestInfo, ctxRetriever, batchMode, batch, numCleanupAttempts, retryDelay, onBeforeRunWithOptions, onCleanupSuccess, onCleanupFailure);
        }, ClientRequestContext.createErrorPromise = function(functionName, code) {
            return void 0 === code && (code = Core.CoreResourceStrings.invalidArgument), Core.CoreUtility._createPromiseFromException(Utility.createRuntimeError(code, Core.CoreUtility._getResourceString(code), functionName));
        }, ClientRequestContext._runCommon = function(functionName, requestInfo, ctxRetriever, batchMode, runBody, numCleanupAttempts, retryDelay, onBeforeRun, onCleanupSuccess, onCleanupFailure) {
            Core.SessionBase._overrideSession && (requestInfo = Core.SessionBase._overrideSession);
            var ctx, resultOrError, previousBatchMode, starterPromise = Core.CoreUtility.createPromise((function(resolve, reject) {
                resolve();
            })), succeeded = !1;
            return starterPromise.then((function() {
                if ((ctx = ctxRetriever(requestInfo))._autoCleanup) return new Promise((function(resolve, reject) {
                    ctx._onRunFinishedNotifiers.push((function() {
                        ctx._autoCleanup = !0, resolve();
                    }));
                }));
                ctx._autoCleanup = !0;
            })).then((function() {
                return "function" != typeof runBody ? ClientRequestContext.createErrorPromise(functionName) : (previousBatchMode = ctx.m_batchMode, 
                ctx.m_batchMode = batchMode, onBeforeRun && onBeforeRun(ctx), runBodyResult = runBody(1 == batchMode ? ctx.batch.bind(ctx) : ctx), 
                (Utility.isNullOrUndefined(runBodyResult) || "function" != typeof runBodyResult.then) && Utility.throwError(ResourceStrings.runMustReturnPromise), 
                runBodyResult);
                var runBodyResult;
            })).then((function(runBodyResult) {
                return 1 === batchMode ? runBodyResult : ctx.sync(runBodyResult);
            })).then((function(result) {
                succeeded = !0, resultOrError = result;
            })).catch((function(error) {
                resultOrError = error;
            })).then((function() {
                var itemsToRemove = ctx.trackedObjects._retrieveAndClearAutoCleanupList();
                for (var key in ctx._autoCleanup = !1, ctx.m_batchMode = previousBatchMode, itemsToRemove) itemsToRemove[key]._objectPath.isValid = !1;
                var cleanupCounter = 0;
                if (Utility._synchronousCleanup || ClientRequestContext.isRequestUrlAndHeaderInfoResolver(requestInfo)) return attemptCleanup();
                function attemptCleanup() {
                    cleanupCounter++;
                    var savedPendingRequest = ctx.m_pendingRequest, savedBatchMode = ctx.m_batchMode, request = new ClientRequest(ctx);
                    ctx.m_pendingRequest = request, ctx.m_batchMode = 0;
                    try {
                        for (var key in itemsToRemove) ctx.trackedObjects.remove(itemsToRemove[key]);
                    } finally {
                        ctx.m_batchMode = savedBatchMode, ctx.m_pendingRequest = savedPendingRequest;
                    }
                    return ctx.syncPrivate(request).then((function() {
                        onCleanupSuccess && onCleanupSuccess(cleanupCounter);
                    })).catch((function() {
                        onCleanupFailure && onCleanupFailure(cleanupCounter), cleanupCounter < numCleanupAttempts && setTimeout((function() {
                            attemptCleanup();
                        }), retryDelay);
                    }));
                }
                attemptCleanup();
            })).then((function() {
                ctx._onRunFinishedNotifiers && ctx._onRunFinishedNotifiers.length > 0 && ctx._onRunFinishedNotifiers.shift()();
                if (succeeded) return resultOrError;
                throw resultOrError;
            }));
        }, ClientRequestContext;
    }(Common.ClientRequestContextBase);
    exports.ClientRequestContext = ClientRequestContext;
    var RetrieveResultImpl = function() {
        function RetrieveResultImpl(m_proxy, m_shouldPolyfill) {
            this.m_proxy = m_proxy, this.m_shouldPolyfill = m_shouldPolyfill;
            var scalarPropertyNames = m_proxy[Constants.scalarPropertyNames], navigationPropertyNames = m_proxy[Constants.navigationPropertyNames], typeName = m_proxy[Constants.className], isCollection = m_proxy[Constants.isCollection];
            if (scalarPropertyNames) for (var i = 0; i < scalarPropertyNames.length; i++) Utility.definePropertyThrowUnloadedException(this, typeName, scalarPropertyNames[i]);
            if (navigationPropertyNames) for (i = 0; i < navigationPropertyNames.length; i++) Utility.definePropertyThrowUnloadedException(this, typeName, navigationPropertyNames[i]);
            isCollection && Utility.definePropertyThrowUnloadedException(this, typeName, Constants.itemsLowerCase);
        }
        return Object.defineProperty(RetrieveResultImpl.prototype, "$proxy", {
            get: function() {
                return this.m_proxy;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(RetrieveResultImpl.prototype, "$isNullObject", {
            get: function() {
                if (!this.m_isLoaded) throw new Core._Internal.RuntimeError({
                    code: ErrorCodes.valueNotLoaded,
                    httpStatusCode: 400,
                    message: Core.CoreUtility._getResourceString(ResourceStrings.valueNotLoaded),
                    debugInfo: {
                        errorLocation: "retrieveResult.$isNullObject"
                    }
                });
                return this.m_isNullObject;
            },
            enumerable: !1,
            configurable: !0
        }), RetrieveResultImpl.prototype.toJSON = function() {
            if (this.m_isLoaded) return this.m_isNullObject ? null : (Utility.isUndefined(this.m_json) && (this.m_json = Utility.purifyJson(this.m_value)), 
            this.m_json);
        }, RetrieveResultImpl.prototype.toString = function() {
            return JSON.stringify(this.toJSON());
        }, RetrieveResultImpl.prototype._handleResult = function(value) {
            this.m_isLoaded = !0, null === value || "object" == typeof value && value && value._IsNull ? (this.m_isNullObject = !0, 
            value = null) : this.m_isNullObject = !1, this.m_shouldPolyfill && (value = Utility.changePropertyNameToCamelLowerCase(value)), 
            this.m_value = value, this.m_proxy._handleRetrieveResult(value, this);
        }, RetrieveResultImpl;
    }(), Constants = function(_super) {
        function Constants() {
            return null !== _super && _super.apply(this, arguments) || this;
        }
        return __extends(Constants, _super), Constants.getItemAt = "GetItemAt", Constants.index = "_Index", 
        Constants.iterativeExecutor = "IterativeExecutor", Constants.isTracked = "_IsTracked", 
        Constants.eventMessageCategory = 65536, Constants.eventWorkbookId = "Workbook", 
        Constants.eventSourceRemote = "Remote", Constants.proxy = "$proxy", Constants.className = "_className", 
        Constants.isCollection = "_isCollection", Constants.collectionPropertyPath = "_collectionPropertyPath", 
        Constants.objectPathInfoDoNotKeepReferenceFieldName = "D", Constants.officeScriptEventId = "X-OfficeScriptEventId", 
        Constants.officeScriptFireRecordingEvent = "X-OfficeScriptFireRecordingEvent", Constants;
    }(Common.CommonConstants);
    exports.Constants = Constants;
    var ClientRequest = function(_super) {
        function ClientRequest(context) {
            var _this = _super.call(this, context) || this;
            return _this.m_context = context, _this.m_pendingProcessEventHandlers = [], _this.m_pendingEventHandlerActions = {}, 
            _this.m_traceInfos = {}, _this.m_responseTraceIds = {}, _this.m_responseTraceMessages = [], 
            _this;
        }
        return __extends(ClientRequest, _super), Object.defineProperty(ClientRequest.prototype, "traceInfos", {
            get: function() {
                return this.m_traceInfos;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ClientRequest.prototype, "_responseTraceMessages", {
            get: function() {
                return this.m_responseTraceMessages;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(ClientRequest.prototype, "_responseTraceIds", {
            get: function() {
                return this.m_responseTraceIds;
            },
            enumerable: !1,
            configurable: !0
        }), ClientRequest.prototype._setResponseTraceIds = function(value) {
            if (value) for (var i = 0; i < value.length; i++) {
                var traceId = value[i];
                this.m_responseTraceIds[traceId] = traceId;
                var message = this.m_traceInfos[traceId];
                Core.CoreUtility.isNullOrUndefined(message) || this.m_responseTraceMessages.push(message);
            }
        }, ClientRequest.prototype.addTrace = function(actionId, message) {
            this.m_traceInfos[actionId] = message;
        }, ClientRequest.prototype._addPendingEventHandlerAction = function(eventHandlers, action) {
            this.m_pendingEventHandlerActions[eventHandlers._id] || (this.m_pendingEventHandlerActions[eventHandlers._id] = [], 
            this.m_pendingProcessEventHandlers.push(eventHandlers)), this.m_pendingEventHandlerActions[eventHandlers._id].push(action);
        }, Object.defineProperty(ClientRequest.prototype, "_pendingProcessEventHandlers", {
            get: function() {
                return this.m_pendingProcessEventHandlers;
            },
            enumerable: !1,
            configurable: !0
        }), ClientRequest.prototype._getPendingEventHandlerActions = function(eventHandlers) {
            return this.m_pendingEventHandlerActions[eventHandlers._id];
        }, ClientRequest;
    }(Common.ClientRequestBase);
    exports.ClientRequest = ClientRequest;
    var EventHandlers = function() {
        function EventHandlers(context, parentObject, name, eventInfo) {
            var _this = this;
            this.m_id = context._nextId(), this.m_context = context, this.m_name = name, this.m_handlers = [], 
            this.m_registered = !1, this.m_eventInfo = eventInfo, this.m_callback = function(args) {
                _this.m_eventInfo.eventArgsTransformFunc(args).then((function(newArgs) {
                    return _this.fireEvent(newArgs);
                }));
            };
        }
        return Object.defineProperty(EventHandlers.prototype, "_registered", {
            get: function() {
                return this.m_registered;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(EventHandlers.prototype, "_id", {
            get: function() {
                return this.m_id;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(EventHandlers.prototype, "_handlers", {
            get: function() {
                return this.m_handlers;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(EventHandlers.prototype, "_context", {
            get: function() {
                return this.m_context;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(EventHandlers.prototype, "_callback", {
            get: function() {
                return this.m_callback;
            },
            enumerable: !1,
            configurable: !0
        }), EventHandlers.prototype.add = function(handler) {
            var action = ActionFactory.createTraceAction(this.m_context, null, !1);
            return this.m_context._pendingRequest._addPendingEventHandlerAction(this, {
                id: action.actionInfo.Id,
                handler: handler,
                operation: 0
            }), new EventHandlerResult(this.m_context, this, handler);
        }, EventHandlers.prototype.remove = function(handler) {
            var action = ActionFactory.createTraceAction(this.m_context, null, !1);
            this.m_context._pendingRequest._addPendingEventHandlerAction(this, {
                id: action.actionInfo.Id,
                handler: handler,
                operation: 1
            });
        }, EventHandlers.prototype.removeAll = function() {
            var action = ActionFactory.createTraceAction(this.m_context, null, !1);
            this.m_context._pendingRequest._addPendingEventHandlerAction(this, {
                id: action.actionInfo.Id,
                handler: null,
                operation: 2
            });
        }, EventHandlers.prototype._processRegistration = function(req) {
            var _this = this, ret = Core.CoreUtility._createPromiseFromResult(null), actions = req._getPendingEventHandlerActions(this);
            if (!actions) return ret;
            for (var handlersResult = [], i = 0; i < this.m_handlers.length; i++) handlersResult.push(this.m_handlers[i]);
            var hasChange = !1;
            for (i = 0; i < actions.length; i++) if (req._responseTraceIds[actions[i].id]) switch (hasChange = !0, 
            actions[i].operation) {
              case 0:
                handlersResult.push(actions[i].handler);
                break;

              case 1:
                for (var index = handlersResult.length - 1; index >= 0; index--) if (handlersResult[index] === actions[i].handler) {
                    handlersResult.splice(index, 1);
                    break;
                }
                break;

              case 2:
                handlersResult = [];
            }
            return hasChange && (!this.m_registered && handlersResult.length > 0 ? ret = ret.then((function() {
                return _this.m_eventInfo.registerFunc(_this.m_callback);
            })).then((function() {
                return _this.m_registered = !0;
            })) : this.m_registered && 0 == handlersResult.length && (ret = ret.then((function() {
                return _this.m_eventInfo.unregisterFunc(_this.m_callback);
            })).catch((function(ex) {
                Core.CoreUtility.log("Error when unregister event: " + JSON.stringify(ex));
            })).then((function() {
                return _this.m_registered = !1;
            }))), ret = ret.then((function() {
                return _this.m_handlers = handlersResult;
            }))), ret;
        }, EventHandlers.prototype.fireEvent = function(args) {
            for (var promises = [], i = 0; i < this.m_handlers.length; i++) {
                var handler = this.m_handlers[i], p = Core.CoreUtility._createPromiseFromResult(null).then(this.createFireOneEventHandlerFunc(handler, args)).catch((function(ex) {
                    Core.CoreUtility.log("Error when invoke handler: " + JSON.stringify(ex));
                }));
                promises.push(p);
            }
            Core.CoreUtility.Promise.all(promises);
        }, EventHandlers.prototype.createFireOneEventHandlerFunc = function(handler, args) {
            return function() {
                return handler(args);
            };
        }, EventHandlers;
    }();
    exports.EventHandlers = EventHandlers;
    var _Internal, EventHandlerResult = function() {
        function EventHandlerResult(context, handlers, handler) {
            this.m_context = context, this.m_allHandlers = handlers, this.m_handler = handler;
        }
        return Object.defineProperty(EventHandlerResult.prototype, "context", {
            get: function() {
                return this.m_context;
            },
            enumerable: !1,
            configurable: !0
        }), EventHandlerResult.prototype.remove = function() {
            this.m_allHandlers && this.m_handler && (this.m_allHandlers.remove(this.m_handler), 
            this.m_allHandlers = null, this.m_handler = null);
        }, EventHandlerResult;
    }();
    exports.EventHandlerResult = EventHandlerResult, function(_Internal) {
        var OfficeJsEventRegistration = function() {
            function OfficeJsEventRegistration() {}
            return OfficeJsEventRegistration.prototype.register = function(eventId, targetId, handler) {
                switch (eventId) {
                  case 4:
                    return Utility.promisify((function(callback) {
                        return Office.context.document.bindings.getByIdAsync(targetId, callback);
                    })).then((function(officeBinding) {
                        return Utility.promisify((function(callback) {
                            return officeBinding.addHandlerAsync(Office.EventType.BindingDataChanged, handler, callback);
                        }));
                    }));

                  case 3:
                    return Utility.promisify((function(callback) {
                        return Office.context.document.bindings.getByIdAsync(targetId, callback);
                    })).then((function(officeBinding) {
                        return Utility.promisify((function(callback) {
                            return officeBinding.addHandlerAsync(Office.EventType.BindingSelectionChanged, handler, callback);
                        }));
                    }));

                  case 2:
                    return Utility.promisify((function(callback) {
                        return Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler, callback);
                    }));

                  case 1:
                    return Utility.promisify((function(callback) {
                        return Office.context.document.settings.addHandlerAsync(Office.EventType.SettingsChanged, handler, callback);
                    }));

                  case 5:
                    return Utility.promisify((function(callback) {
                        return OSF.DDA.RichApi.richApiMessageManager.addHandlerAsync("richApiMessage", handler, callback);
                    }));

                  case 13:
                    return Utility.promisify((function(callback) {
                        return Office.context.document.addHandlerAsync(Office.EventType.ObjectDeleted, handler, {
                            id: targetId
                        }, callback);
                    }));

                  case 14:
                    return Utility.promisify((function(callback) {
                        return Office.context.document.addHandlerAsync(Office.EventType.ObjectSelectionChanged, handler, {
                            id: targetId
                        }, callback);
                    }));

                  case 15:
                    return Utility.promisify((function(callback) {
                        return Office.context.document.addHandlerAsync(Office.EventType.ObjectDataChanged, handler, {
                            id: targetId
                        }, callback);
                    }));

                  case 16:
                    return Utility.promisify((function(callback) {
                        return Office.context.document.addHandlerAsync(Office.EventType.ContentControlAdded, handler, {
                            id: targetId
                        }, callback);
                    }));

                  default:
                    throw Core._Internal.RuntimeError._createInvalidArgError({
                        argumentName: "eventId"
                    });
                }
            }, OfficeJsEventRegistration.prototype.unregister = function(eventId, targetId, handler) {
                switch (eventId) {
                  case 4:
                    return Utility.promisify((function(callback) {
                        return Office.context.document.bindings.getByIdAsync(targetId, callback);
                    })).then((function(officeBinding) {
                        return Utility.promisify((function(callback) {
                            return officeBinding.removeHandlerAsync(Office.EventType.BindingDataChanged, {
                                handler: handler
                            }, callback);
                        }));
                    }));

                  case 3:
                    return Utility.promisify((function(callback) {
                        return Office.context.document.bindings.getByIdAsync(targetId, callback);
                    })).then((function(officeBinding) {
                        return Utility.promisify((function(callback) {
                            return officeBinding.removeHandlerAsync(Office.EventType.BindingSelectionChanged, {
                                handler: handler
                            }, callback);
                        }));
                    }));

                  case 2:
                    return Utility.promisify((function(callback) {
                        return Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, {
                            handler: handler
                        }, callback);
                    }));

                  case 1:
                    return Utility.promisify((function(callback) {
                        return Office.context.document.settings.removeHandlerAsync(Office.EventType.SettingsChanged, {
                            handler: handler
                        }, callback);
                    }));

                  case 5:
                    return Utility.promisify((function(callback) {
                        return OSF.DDA.RichApi.richApiMessageManager.removeHandlerAsync("richApiMessage", {
                            handler: handler
                        }, callback);
                    }));

                  case 13:
                    return Utility.promisify((function(callback) {
                        return Office.context.document.removeHandlerAsync(Office.EventType.ObjectDeleted, {
                            id: targetId,
                            handler: handler
                        }, callback);
                    }));

                  case 14:
                    return Utility.promisify((function(callback) {
                        return Office.context.document.removeHandlerAsync(Office.EventType.ObjectSelectionChanged, {
                            id: targetId,
                            handler: handler
                        }, callback);
                    }));

                  case 15:
                    return Utility.promisify((function(callback) {
                        return Office.context.document.removeHandlerAsync(Office.EventType.ObjectDataChanged, {
                            id: targetId,
                            handler: handler
                        }, callback);
                    }));

                  case 16:
                    return Utility.promisify((function(callback) {
                        return Office.context.document.removeHandlerAsync(Office.EventType.ContentControlAdded, {
                            id: targetId,
                            handler: handler
                        }, callback);
                    }));

                  default:
                    throw Core._Internal.RuntimeError._createInvalidArgError({
                        argumentName: "eventId"
                    });
                }
            }, OfficeJsEventRegistration;
        }();
        _Internal.officeJsEventRegistration = new OfficeJsEventRegistration;
    }(_Internal = exports._Internal || (exports._Internal = {}));
    var EventRegistration = function() {
        function EventRegistration(registerEventImpl, unregisterEventImpl) {
            this.m_handlersByEventByTarget = {}, this.m_registerEventImpl = registerEventImpl, 
            this.m_unregisterEventImpl = unregisterEventImpl;
        }
        return EventRegistration.getTargetIdOrDefault = function(targetId) {
            return Utility.isNullOrUndefined(targetId) ? "" : targetId;
        }, EventRegistration.prototype.getHandlers = function(eventId, targetId) {
            targetId = EventRegistration.getTargetIdOrDefault(targetId);
            var handlersById = this.m_handlersByEventByTarget[eventId];
            handlersById || (handlersById = {}, this.m_handlersByEventByTarget[eventId] = handlersById);
            var handlers = handlersById[targetId];
            return handlers || (handlers = [], handlersById[targetId] = handlers), handlers;
        }, EventRegistration.prototype.callHandlers = function(eventId, targetId, argument) {
            for (var funcs = this.getHandlers(eventId, targetId), i = 0; i < funcs.length; i++) funcs[i](argument);
        }, EventRegistration.prototype.hasHandlers = function(eventId, targetId) {
            return this.getHandlers(eventId, targetId).length > 0;
        }, EventRegistration.prototype.register = function(eventId, targetId, handler) {
            if (!handler) throw Core._Internal.RuntimeError._createInvalidArgError({
                argumentName: "handler"
            });
            var handlers = this.getHandlers(eventId, targetId);
            return handlers.push(handler), 1 === handlers.length ? this.m_registerEventImpl(eventId, targetId) : Utility._createPromiseFromResult(null);
        }, EventRegistration.prototype.unregister = function(eventId, targetId, handler) {
            if (!handler) throw Core._Internal.RuntimeError._createInvalidArgError({
                argumentName: "handler"
            });
            for (var handlers = this.getHandlers(eventId, targetId), index = handlers.length - 1; index >= 0; index--) if (handlers[index] === handler) {
                handlers.splice(index, 1);
                break;
            }
            return 0 === handlers.length ? this.m_unregisterEventImpl(eventId, targetId) : Utility._createPromiseFromResult(null);
        }, EventRegistration;
    }();
    exports.EventRegistration = EventRegistration;
    var GenericEventRegistration = function() {
        function GenericEventRegistration() {
            this.m_eventRegistration = new EventRegistration(this._registerEventImpl.bind(this), this._unregisterEventImpl.bind(this)), 
            this.m_richApiMessageHandler = this._handleRichApiMessage.bind(this);
        }
        return GenericEventRegistration.prototype.ready = function() {
            var _this = this;
            return this.m_ready || (GenericEventRegistration._testReadyImpl ? this.m_ready = GenericEventRegistration._testReadyImpl().then((function() {
                _this.m_isReady = !0;
            })) : Core.HostBridge.instance ? this.m_ready = Utility._createPromiseFromResult(null).then((function() {
                _this.m_isReady = !0;
            })) : this.m_ready = _Internal.officeJsEventRegistration.register(5, "", this.m_richApiMessageHandler).then((function() {
                _this.m_isReady = !0;
            }))), this.m_ready;
        }, Object.defineProperty(GenericEventRegistration.prototype, "isReady", {
            get: function() {
                return this.m_isReady;
            },
            enumerable: !1,
            configurable: !0
        }), GenericEventRegistration.prototype.register = function(eventId, targetId, handler) {
            var _this = this;
            return this.ready().then((function() {
                return _this.m_eventRegistration.register(eventId, targetId, handler);
            }));
        }, GenericEventRegistration.prototype.unregister = function(eventId, targetId, handler) {
            var _this = this;
            return this.ready().then((function() {
                return _this.m_eventRegistration.unregister(eventId, targetId, handler);
            }));
        }, GenericEventRegistration.prototype._registerEventImpl = function(eventId, targetId) {
            return Utility._createPromiseFromResult(null);
        }, GenericEventRegistration.prototype._unregisterEventImpl = function(eventId, targetId) {
            return Utility._createPromiseFromResult(null);
        }, GenericEventRegistration.prototype._handleRichApiMessage = function(msg) {
            if (msg && msg.entries) for (var entryIndex = 0; entryIndex < msg.entries.length; entryIndex++) {
                var entry = msg.entries[entryIndex];
                if (entry.messageCategory == Constants.eventMessageCategory) {
                    Core.CoreUtility._logEnabled && Core.CoreUtility.log(JSON.stringify(entry));
                    var eventId = entry.messageType, targetId = entry.targetId;
                    if (this.m_eventRegistration.hasHandlers(eventId, targetId)) {
                        var arg = JSON.parse(entry.message);
                        entry.isRemoteOverride && (arg.source = Constants.eventSourceRemote), this.m_eventRegistration.callHandlers(eventId, targetId, arg);
                    }
                }
            }
        }, GenericEventRegistration.getGenericEventRegistration = function() {
            return GenericEventRegistration.s_genericEventRegistration || (GenericEventRegistration.s_genericEventRegistration = new GenericEventRegistration), 
            GenericEventRegistration.s_genericEventRegistration;
        }, GenericEventRegistration.richApiMessageEventCategory = 65536, GenericEventRegistration;
    }();
    exports.GenericEventRegistration = GenericEventRegistration, exports._testSetRichApiMessageReadyImpl = function(impl) {
        GenericEventRegistration._testReadyImpl = impl;
    }, exports._testTriggerRichApiMessageEvent = function(msg) {
        GenericEventRegistration.getGenericEventRegistration()._handleRichApiMessage(msg);
    };
    var GenericEventHandlers = function(_super) {
        function GenericEventHandlers(context, parentObject, name, eventInfo) {
            var _this = _super.call(this, context, parentObject, name, eventInfo) || this;
            return _this.m_genericEventInfo = eventInfo, _this;
        }
        return __extends(GenericEventHandlers, _super), GenericEventHandlers.prototype.add = function(handler) {
            var _this = this;
            return 0 == this._handlers.length && this.m_genericEventInfo.registerFunc && this.m_genericEventInfo.registerFunc(), 
            GenericEventRegistration.getGenericEventRegistration().isReady || this._context._pendingRequest._addPreSyncPromise(GenericEventRegistration.getGenericEventRegistration().ready()), 
            ActionFactory.createTraceMarkerForCallback(this._context, (function() {
                _this._handlers.push(handler), 1 == _this._handlers.length && GenericEventRegistration.getGenericEventRegistration().register(_this.m_genericEventInfo.eventType, _this.m_genericEventInfo.getTargetIdFunc(), _this._callback);
            })), new EventHandlerResult(this._context, this, handler);
        }, GenericEventHandlers.prototype.remove = function(handler) {
            var _this = this;
            1 == this._handlers.length && this.m_genericEventInfo.unregisterFunc && this.m_genericEventInfo.unregisterFunc(), 
            ActionFactory.createTraceMarkerForCallback(this._context, (function() {
                for (var handlers = _this._handlers, index = handlers.length - 1; index >= 0; index--) if (handlers[index] === handler) {
                    handlers.splice(index, 1);
                    break;
                }
                0 == handlers.length && GenericEventRegistration.getGenericEventRegistration().unregister(_this.m_genericEventInfo.eventType, _this.m_genericEventInfo.getTargetIdFunc(), _this._callback);
            }));
        }, GenericEventHandlers.prototype.removeAll = function() {}, GenericEventHandlers;
    }(EventHandlers);
    exports.GenericEventHandlers = GenericEventHandlers;
    var InstantiateActionResultHandler = function() {
        function InstantiateActionResultHandler(clientObject) {
            this.m_clientObject = clientObject;
        }
        return InstantiateActionResultHandler.prototype._handleResult = function(value) {
            this.m_clientObject._handleIdResult(value);
        }, InstantiateActionResultHandler;
    }(), ObjectPathFactory = function() {
        function ObjectPathFactory() {}
        return ObjectPathFactory.createGlobalObjectObjectPath = function(context) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 1,
                Name: ""
            };
            return new Common.ObjectPath(objectPathInfo, null, !1, !1, 1, 4);
        }, ObjectPathFactory.createNewObjectObjectPath = function(context, typeName, isCollection, flags) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 2,
                Name: typeName
            };
            return new Common.ObjectPath(objectPathInfo, null, isCollection, !1, 1, Utility._fixupApiFlags(flags));
        }, ObjectPathFactory.createPropertyObjectPath = function(context, parent, propertyName, isCollection, isInvalidAfterRequest, flags) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 4,
                Name: propertyName,
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id
            };
            return new Common.ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest, 1, Utility._fixupApiFlags(flags));
        }, ObjectPathFactory.createIndexerObjectPath = function(context, parent, args) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 5,
                Name: "",
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            return objectPathInfo.ArgumentInfo.Arguments = args, new Common.ObjectPath(objectPathInfo, parent._objectPath, !1, !1, 1, 4);
        }, ObjectPathFactory.createIndexerObjectPathUsingParentPath = function(context, parentObjectPath, args) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 5,
                Name: "",
                ParentObjectPathId: parentObjectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            return objectPathInfo.ArgumentInfo.Arguments = args, new Common.ObjectPath(objectPathInfo, parentObjectPath, !1, !1, 1, 4);
        }, ObjectPathFactory.createMethodObjectPath = function(context, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, flags) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 3,
                Name: methodName,
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            }, argumentObjectPaths = Utility.setMethodArguments(context, objectPathInfo.ArgumentInfo, args), ret = new Common.ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest, operationType, Utility._fixupApiFlags(flags));
            return ret.argumentObjectPaths = argumentObjectPaths, ret.getByIdMethodName = getByIdMethodName, 
            ret;
        }, ObjectPathFactory.createReferenceIdObjectPath = function(context, referenceId) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 6,
                Name: referenceId,
                ArgumentInfo: {}
            };
            return new Common.ObjectPath(objectPathInfo, null, !1, !1, 1, 4);
        }, ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt = function(hasIndexerMethod, context, parent, childItem, index) {
            var id = Utility.tryGetObjectIdFromLoadOrRetrieveResult(childItem);
            return hasIndexerMethod && !Utility.isNullOrUndefined(id) ? ObjectPathFactory.createChildItemObjectPathUsingIndexer(context, parent, childItem) : ObjectPathFactory.createChildItemObjectPathUsingGetItemAt(context, parent, childItem, index);
        }, ObjectPathFactory.createChildItemObjectPathUsingIndexer = function(context, parent, childItem) {
            var id = Utility.tryGetObjectIdFromLoadOrRetrieveResult(childItem), objectPathInfo = objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 5,
                Name: "",
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            return objectPathInfo.ArgumentInfo.Arguments = [ id ], new Common.ObjectPath(objectPathInfo, parent._objectPath, !1, !1, 1, 4);
        }, ObjectPathFactory.createChildItemObjectPathUsingGetItemAt = function(context, parent, childItem, index) {
            var indexFromServer = childItem[Constants.index];
            indexFromServer && (index = indexFromServer);
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 3,
                Name: Constants.getItemAt,
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            return objectPathInfo.ArgumentInfo.Arguments = [ index ], new Common.ObjectPath(objectPathInfo, parent._objectPath, !1, !1, 1, 4);
        }, ObjectPathFactory;
    }();
    exports.ObjectPathFactory = ObjectPathFactory;
    var OfficeJsRequestExecutor = function() {
        function OfficeJsRequestExecutor(context) {
            this.m_context = context;
        }
        return OfficeJsRequestExecutor.prototype.executeAsync = function(customData, requestFlags, requestMessage) {
            var _this = this, messageSafearray = Core.RichApiMessageUtility.buildMessageArrayForIRequestExecutor(customData, requestFlags, requestMessage, OfficeJsRequestExecutor.SourceLibHeaderValue);
            return new Promise((function(resolve, reject) {
                OSF.DDA.RichApi.executeRichApiRequestAsync(messageSafearray, (function(result) {
                    var response;
                    Core.CoreUtility.log("Response:"), Core.CoreUtility.log(JSON.stringify(result)), 
                    "succeeded" == result.status ? response = Core.RichApiMessageUtility.buildResponseOnSuccess(Core.RichApiMessageUtility.getResponseBody(result), Core.RichApiMessageUtility.getResponseHeaders(result)) : (response = Core.RichApiMessageUtility.buildResponseOnError(result.error.code, result.error.message), 
                    _this.m_context._processOfficeJsErrorResponse(result.error.code, response)), resolve(response);
                }));
            }));
        }, OfficeJsRequestExecutor.SourceLibHeaderValue = "officejs", OfficeJsRequestExecutor;
    }(), TrackedObjects = function() {
        function TrackedObjects(context) {
            this._autoCleanupList = {}, this.m_context = context;
        }
        return TrackedObjects.prototype.add = function(param) {
            var _this = this;
            Array.isArray(param) ? param.forEach((function(item) {
                return _this._addCommon(item, !0);
            })) : this._addCommon(param, !0);
        }, TrackedObjects.prototype._autoAdd = function(object) {
            this._addCommon(object, !1), this._autoCleanupList[object._objectPath.objectPathInfo.Id] = object;
        }, TrackedObjects.prototype._autoTrackIfNecessaryWhenHandleObjectResultValue = function(object, resultValue) {
            this.m_context._autoCleanup && !object[Constants.isTracked] && object !== this.m_context._rootObject && resultValue && !Utility.isNullOrEmptyString(resultValue[Constants.referenceId]) && (this._autoCleanupList[object._objectPath.objectPathInfo.Id] = object, 
            object[Constants.isTracked] = !0);
        }, TrackedObjects.prototype._addCommon = function(object, isExplicitlyAdded) {
            if (object[Constants.isTracked]) isExplicitlyAdded && this.m_context._autoCleanup && delete this._autoCleanupList[object._objectPath.objectPathInfo.Id]; else {
                var referenceId = object[Constants.referenceId];
                if (object._objectPath.objectPathInfo[Constants.objectPathInfoDoNotKeepReferenceFieldName]) throw Utility.createRuntimeError(Core.CoreErrorCodes.generalException, Core.CoreUtility._getResourceString(ResourceStrings.objectIsUntracked), null);
                Utility.isNullOrEmptyString(referenceId) && object._KeepReference && (object._KeepReference(), 
                ActionFactory.createInstantiateAction(this.m_context, object), isExplicitlyAdded && this.m_context._autoCleanup && delete this._autoCleanupList[object._objectPath.objectPathInfo.Id], 
                object[Constants.isTracked] = !0);
            }
        }, TrackedObjects.prototype.remove = function(param) {
            var _this = this;
            Array.isArray(param) ? param.forEach((function(item) {
                return _this._removeCommon(item);
            })) : this._removeCommon(param);
        }, TrackedObjects.prototype._removeCommon = function(object) {
            object._objectPath.objectPathInfo[Constants.objectPathInfoDoNotKeepReferenceFieldName] = !0, 
            object.context._pendingRequest._removeKeepReferenceAction(object._objectPath.objectPathInfo.Id);
            var referenceId = object[Constants.referenceId];
            if (!Utility.isNullOrEmptyString(referenceId)) {
                var rootObject = this.m_context._rootObject;
                rootObject._RemoveReference && rootObject._RemoveReference(referenceId);
            }
            delete object[Constants.isTracked];
        }, TrackedObjects.prototype._retrieveAndClearAutoCleanupList = function() {
            var list = this._autoCleanupList;
            return this._autoCleanupList = {}, list;
        }, TrackedObjects;
    }();
    exports.TrackedObjects = TrackedObjects;
    var RequestPrettyPrinter = function() {
        function RequestPrettyPrinter(globalObjName, referencedObjectPaths, actions, showDispose, removePII) {
            globalObjName || (globalObjName = "root"), this.m_globalObjName = globalObjName, 
            this.m_referencedObjectPaths = referencedObjectPaths, this.m_actions = actions, 
            this.m_statements = [], this.m_variableNameForObjectPathMap = {}, this.m_variableNameToObjectPathMap = {}, 
            this.m_declaredObjectPathMap = {}, this.m_showDispose = showDispose, this.m_removePII = removePII;
        }
        return RequestPrettyPrinter.prototype.process = function() {
            this.m_showDispose && ClientRequest._calculateLastUsedObjectPathIds(this.m_actions);
            for (var i = 0; i < this.m_actions.length; i++) this.processOneAction(this.m_actions[i]);
            return this.m_statements;
        }, RequestPrettyPrinter.prototype.processForDebugStatementInfo = function(actionIndex) {
            this.m_showDispose && ClientRequest._calculateLastUsedObjectPathIds(this.m_actions);
            this.m_statements = [];
            for (var statementIndex = -1, i = 0; i < this.m_actions.length && (this.processOneAction(this.m_actions[i]), 
            actionIndex == i && (statementIndex = this.m_statements.length - 1), !(statementIndex >= 0 && this.m_statements.length > statementIndex + 5 + 1)); i++) ;
            if (statementIndex < 0) return null;
            var startIndex = statementIndex - 5;
            startIndex < 0 && (startIndex = 0);
            var endIndex = statementIndex + 1 + 5;
            endIndex > this.m_statements.length && (endIndex = this.m_statements.length);
            var surroundingStatements = [];
            0 != startIndex && surroundingStatements.push("...");
            for (var i_1 = startIndex; i_1 < statementIndex; i_1++) surroundingStatements.push(this.m_statements[i_1]);
            surroundingStatements.push("// >>>>>"), surroundingStatements.push(this.m_statements[statementIndex]), 
            surroundingStatements.push("// <<<<<");
            for (var i_2 = statementIndex + 1; i_2 < endIndex; i_2++) surroundingStatements.push(this.m_statements[i_2]);
            return endIndex < this.m_statements.length && surroundingStatements.push("..."), 
            {
                statement: this.m_statements[statementIndex],
                surroundingStatements: surroundingStatements
            };
        }, RequestPrettyPrinter.prototype.processOneAction = function(action) {
            switch (action.actionInfo.ActionType) {
              case 1:
                this.processInstantiateAction(action);
                break;

              case 3:
                this.processMethodAction(action);
                break;

              case 2:
                this.processQueryAction(action);
                break;

              case 7:
                this.processQueryAsJsonAction(action);
                break;

              case 6:
                this.processRecursiveQueryAction(action);
                break;

              case 4:
                this.processSetPropertyAction(action);
                break;

              case 5:
                this.processTraceAction(action);
                break;

              case 8:
                this.processEnsureUnchangedAction(action);
                break;

              case 9:
                this.processUpdateAction(action);
            }
        }, RequestPrettyPrinter.prototype.processInstantiateAction = function(action) {
            var objId = action.actionInfo.ObjectPathId, objPath = this.m_referencedObjectPaths[objId], varName = this.getObjVarName(objId);
            if (this.m_declaredObjectPathMap[objId]) {
                statement = "// Instantiate {" + varName + "}";
                statement = this.appendDisposeCommentIfRelevant(statement, action), this.m_statements.push(statement);
            } else {
                var statement = "var " + varName + " = " + this.buildObjectPathExpressionWithParent(objPath) + ";";
                statement = this.appendDisposeCommentIfRelevant(statement, action), this.m_statements.push(statement), 
                this.m_declaredObjectPathMap[objId] = varName;
            }
        }, RequestPrettyPrinter.prototype.processMethodAction = function(action) {
            var methodName = action.actionInfo.Name;
            if ("_KeepReference" === methodName) {
                if (!Common._internalConfig.showInternalApiInDebugInfo) return;
                methodName = "track";
            }
            var statement = this.getObjVarName(action.actionInfo.ObjectPathId) + "." + Utility._toCamelLowerCase(methodName) + "(" + this.buildArgumentsExpression(action.actionInfo.ArgumentInfo) + ");";
            statement = this.appendDisposeCommentIfRelevant(statement, action), this.m_statements.push(statement);
        }, RequestPrettyPrinter.prototype.processQueryAction = function(action) {
            var queryExp = this.buildQueryExpression(action), statement = this.getObjVarName(action.actionInfo.ObjectPathId) + ".load(" + queryExp + ");";
            statement = this.appendDisposeCommentIfRelevant(statement, action), this.m_statements.push(statement);
        }, RequestPrettyPrinter.prototype.processQueryAsJsonAction = function(action) {
            var queryExp = this.buildQueryExpression(action), statement = this.getObjVarName(action.actionInfo.ObjectPathId) + ".retrieve(" + queryExp + ");";
            statement = this.appendDisposeCommentIfRelevant(statement, action), this.m_statements.push(statement);
        }, RequestPrettyPrinter.prototype.processRecursiveQueryAction = function(action) {
            var queryExp = "";
            action.actionInfo.RecursiveQueryInfo && (queryExp = JSON.stringify(action.actionInfo.RecursiveQueryInfo));
            var statement = this.getObjVarName(action.actionInfo.ObjectPathId) + ".loadRecursive(" + queryExp + ");";
            statement = this.appendDisposeCommentIfRelevant(statement, action), this.m_statements.push(statement);
        }, RequestPrettyPrinter.prototype.processSetPropertyAction = function(action) {
            var statement = this.getObjVarName(action.actionInfo.ObjectPathId) + "." + Utility._toCamelLowerCase(action.actionInfo.Name) + " = " + this.buildArgumentsExpression(action.actionInfo.ArgumentInfo) + ";";
            statement = this.appendDisposeCommentIfRelevant(statement, action), this.m_statements.push(statement);
        }, RequestPrettyPrinter.prototype.processTraceAction = function(action) {
            var statement = "context.trace();";
            statement = this.appendDisposeCommentIfRelevant(statement, action), this.m_statements.push(statement);
        }, RequestPrettyPrinter.prototype.processEnsureUnchangedAction = function(action) {
            var statement = this.getObjVarName(action.actionInfo.ObjectPathId) + ".ensureUnchanged(" + JSON.stringify(action.actionInfo.ObjectState) + ");";
            statement = this.appendDisposeCommentIfRelevant(statement, action), this.m_statements.push(statement);
        }, RequestPrettyPrinter.prototype.processUpdateAction = function(action) {
            var statement = this.getObjVarName(action.actionInfo.ObjectPathId) + ".update(" + JSON.stringify(action.actionInfo.ObjectState) + ");";
            statement = this.appendDisposeCommentIfRelevant(statement, action), this.m_statements.push(statement);
        }, RequestPrettyPrinter.prototype.appendDisposeCommentIfRelevant = function(statement, action) {
            var _this = this;
            if (this.m_showDispose) {
                var lastUsedObjectPathIds = action.actionInfo.L;
                if (lastUsedObjectPathIds && lastUsedObjectPathIds.length > 0) return statement + " // And then dispose {" + lastUsedObjectPathIds.map((function(item) {
                    return _this.getObjVarName(item);
                })).join(", ") + "}";
            }
            return statement;
        }, RequestPrettyPrinter.prototype.buildQueryExpression = function(action) {
            if (action.actionInfo.QueryInfo) {
                var option = {};
                return option.select = action.actionInfo.QueryInfo.Select, option.expand = action.actionInfo.QueryInfo.Expand, 
                option.skip = action.actionInfo.QueryInfo.Skip, option.top = action.actionInfo.QueryInfo.Top, 
                void 0 === option.top && void 0 === option.skip && void 0 === option.expand ? void 0 === option.select ? "" : JSON.stringify(option.select) : JSON.stringify(option);
            }
            return "";
        }, RequestPrettyPrinter.prototype.buildObjectPathExpressionWithParent = function(objPath) {
            return (5 == objPath.objectPathInfo.ObjectPathType || 3 == objPath.objectPathInfo.ObjectPathType || 4 == objPath.objectPathInfo.ObjectPathType) && objPath.objectPathInfo.ParentObjectPathId ? this.getObjVarName(objPath.objectPathInfo.ParentObjectPathId) + "." + this.buildObjectPathExpression(objPath) : this.buildObjectPathExpression(objPath);
        }, RequestPrettyPrinter.prototype.buildObjectPathExpression = function(objPath) {
            var expr = this.buildObjectPathInfoExpression(objPath.objectPathInfo), originalObjectPathInfo = objPath.originalObjectPathInfo;
            return originalObjectPathInfo && (expr = expr + " /* originally " + this.buildObjectPathInfoExpression(originalObjectPathInfo) + " */"), 
            expr;
        }, RequestPrettyPrinter.prototype.buildObjectPathInfoExpression = function(objectPathInfo) {
            switch (objectPathInfo.ObjectPathType) {
              case 1:
                return "context." + this.m_globalObjName;

              case 5:
                return "getItem(" + this.buildArgumentsExpression(objectPathInfo.ArgumentInfo) + ")";

              case 3:
                return Utility._toCamelLowerCase(objectPathInfo.Name) + "(" + this.buildArgumentsExpression(objectPathInfo.ArgumentInfo) + ")";

              case 2:
                return objectPathInfo.Name + ".newObject()";

              case 7:
                return "null";

              case 4:
                return Utility._toCamelLowerCase(objectPathInfo.Name);

              case 6:
                return "context." + this.m_globalObjName + "._getObjectByReferenceId(" + JSON.stringify(objectPathInfo.Name) + ")";
            }
        }, RequestPrettyPrinter.prototype.buildArgumentsExpression = function(args) {
            var ret = "";
            if (!args.Arguments || 0 === args.Arguments.length) return ret;
            if (this.m_removePII) return void 0 === args.Arguments[0] ? ret : "...";
            for (var i = 0; i < args.Arguments.length; i++) i > 0 && (ret += ", "), ret += this.buildArgumentLiteral(args.Arguments[i], args.ReferencedObjectPathIds ? args.ReferencedObjectPathIds[i] : null);
            return "undefined" === ret && (ret = ""), ret;
        }, RequestPrettyPrinter.prototype.buildArgumentLiteral = function(value, objectPathId) {
            return "number" == typeof value && value === objectPathId ? this.getObjVarName(objectPathId) : JSON.stringify(value);
        }, RequestPrettyPrinter.prototype.getObjVarNameBase = function(objectPathId) {
            var ret = "v", objPath = this.m_referencedObjectPaths[objectPathId];
            if (objPath) switch (objPath.objectPathInfo.ObjectPathType) {
              case 1:
                ret = this.m_globalObjName;
                break;

              case 4:
                ret = Utility._toCamelLowerCase(objPath.objectPathInfo.Name);
                break;

              case 3:
                var methodName = objPath.objectPathInfo.Name;
                methodName.length > 3 && "Get" === methodName.substr(0, 3) && (methodName = methodName.substr(3)), 
                ret = Utility._toCamelLowerCase(methodName);
                break;

              case 5:
                var parentName = this.getObjVarNameBase(objPath.objectPathInfo.ParentObjectPathId);
                ret = "s" === parentName.charAt(parentName.length - 1) ? parentName.substr(0, parentName.length - 1) : parentName + "Item";
            }
            return ret;
        }, RequestPrettyPrinter.prototype.getObjVarName = function(objectPathId) {
            if (this.m_variableNameForObjectPathMap[objectPathId]) return this.m_variableNameForObjectPathMap[objectPathId];
            var ret = this.getObjVarNameBase(objectPathId);
            if (!this.m_variableNameToObjectPathMap[ret]) return this.m_variableNameForObjectPathMap[objectPathId] = ret, 
            this.m_variableNameToObjectPathMap[ret] = objectPathId, ret;
            for (var i = 1; this.m_variableNameToObjectPathMap[ret + i.toString()]; ) i++;
            return ret += i.toString(), this.m_variableNameForObjectPathMap[objectPathId] = ret, 
            this.m_variableNameToObjectPathMap[ret] = objectPathId, ret;
        }, RequestPrettyPrinter;
    }(), ResourceStrings = function(_super) {
        function ResourceStrings() {
            return null !== _super && _super.apply(this, arguments) || this;
        }
        return __extends(ResourceStrings, _super), ResourceStrings.cannotRegisterEvent = "CannotRegisterEvent", 
        ResourceStrings.connectionFailureWithStatus = "ConnectionFailureWithStatus", ResourceStrings.connectionFailureWithDetails = "ConnectionFailureWithDetails", 
        ResourceStrings.propertyNotLoaded = "PropertyNotLoaded", ResourceStrings.runMustReturnPromise = "RunMustReturnPromise", 
        ResourceStrings.moreInfoInnerError = "MoreInfoInnerError", ResourceStrings.cannotApplyPropertyThroughSetMethod = "CannotApplyPropertyThroughSetMethod", 
        ResourceStrings.invalidOperationInCellEditMode = "InvalidOperationInCellEditMode", 
        ResourceStrings.objectIsUntracked = "ObjectIsUntracked", ResourceStrings.customFunctionDefintionMissing = "CustomFunctionDefintionMissing", 
        ResourceStrings.customFunctionImplementationMissing = "CustomFunctionImplementationMissing", 
        ResourceStrings.customFunctionNameContainsBadChars = "CustomFunctionNameContainsBadChars", 
        ResourceStrings.customFunctionNameCannotSplit = "CustomFunctionNameCannotSplit", 
        ResourceStrings.customFunctionUnexpectedNumberOfEntriesInResultBatch = "CustomFunctionUnexpectedNumberOfEntriesInResultBatch", 
        ResourceStrings.customFunctionCancellationHandlerMissing = "CustomFunctionCancellationHandlerMissing", 
        ResourceStrings.customFunctionInvalidFunction = "CustomFunctionInvalidFunction", 
        ResourceStrings.customFunctionInvalidFunctionMapping = "CustomFunctionInvalidFunctionMapping", 
        ResourceStrings.customFunctionWindowMissing = "CustomFunctionWindowMissing", ResourceStrings.customFunctionDefintionMissingOnWindow = "CustomFunctionDefintionMissingOnWindow", 
        ResourceStrings.pendingBatchInProgress = "PendingBatchInProgress", ResourceStrings.notInsideBatch = "NotInsideBatch", 
        ResourceStrings.cannotUpdateReadOnlyProperty = "CannotUpdateReadOnlyProperty", ResourceStrings;
    }(Common.CommonResourceStrings);
    exports.ResourceStrings = ResourceStrings, Core.CoreUtility.addResourceStringValues({
        CannotRegisterEvent: "The event handler cannot be registered.",
        PropertyNotLoaded: "The property '{0}' is not available. Before reading the property's value, call the load method on the containing object and call \"context.sync()\" on the associated request context.",
        RunMustReturnPromise: 'The batch function passed to the ".run" method didn\'t return a promise. The function must return a promise, so that any automatically-tracked objects can be released at the completion of the batch operation. Typically, you return a promise by returning the response from "context.sync()".',
        InvalidOrTimedOutSessionMessage: "Your Office Online session has expired or is invalid. To continue, refresh the page.",
        InvalidOperationInCellEditMode: "Excel is in cell-editing mode. Please exit the edit mode by pressing ENTER or TAB or selecting another cell, and then try again.",
        CustomFunctionDefintionMissing: "A property with the name '{0}' that represents the function's definition must exist on Excel.Script.CustomFunctions.",
        CustomFunctionDefintionMissingOnWindow: "A property with the name '{0}' that represents the function's definition must exist on the window object.",
        CustomFunctionImplementationMissing: "The property with the name '{0}' on Excel.Script.CustomFunctions that represents the function's definition must contain a 'call' property that implements the function.",
        CustomFunctionNameContainsBadChars: "The function name may only contain letters, digits, underscores, and periods.",
        CustomFunctionNameCannotSplit: "The function name must contain a non-empty namespace and a non-empty short name.",
        CustomFunctionUnexpectedNumberOfEntriesInResultBatch: "The batching function returned a number of results that doesn't match the number of parameter value sets that were passed into it.",
        CustomFunctionCancellationHandlerMissing: "The cancellation handler onCanceled is missing in the function. The handler must be present as the function is defined as cancelable.",
        CustomFunctionInvalidFunction: "The property with the name '{0}' that represents the function's definition is not a valid function.",
        CustomFunctionInvalidFunctionMapping: "The property with the name '{0}' on CustomFunctionMappings that represents the function's definition is not a valid function.",
        CustomFunctionWindowMissing: "The window object was not found.",
        PendingBatchInProgress: "There is a pending batch in progress. The batch method may not be called inside another batch, or simultaneously with another batch.",
        NotInsideBatch: "Operations may not be invoked outside of a batch method.",
        CannotUpdateReadOnlyProperty: "The property '{0}' is read-only and it cannot be updated.",
        ObjectIsUntracked: "The object is untracked."
    });
    var Utility = function(_super) {
        function Utility() {
            return null !== _super && _super.apply(this, arguments) || this;
        }
        return __extends(Utility, _super), Utility.fixObjectPathIfNecessary = function(clientObject, value) {
            clientObject && clientObject._objectPath && value && clientObject._objectPath.updateUsingObjectData(value, clientObject);
        }, Utility.load = function(clientObj, option) {
            return clientObj.context.load(clientObj, option), clientObj;
        }, Utility.loadAndSync = function(clientObj, option) {
            return clientObj.context.load(clientObj, option), clientObj.context.sync().then((function() {
                return clientObj;
            }));
        }, Utility.retrieve = function(clientObj, option) {
            var shouldPolyfill = Common._internalConfig.alwaysPolyfillClientObjectRetrieveMethod;
            shouldPolyfill || (shouldPolyfill = !Utility.isSetSupported("RichApiRuntime", "1.1"));
            var result = new RetrieveResultImpl(clientObj, shouldPolyfill);
            return clientObj._retrieve(option, result), result;
        }, Utility.retrieveAndSync = function(clientObj, option) {
            var result = Utility.retrieve(clientObj, option);
            return clientObj.context.sync().then((function() {
                return result;
            }));
        }, Utility.toJson = function(clientObj, scalarProperties, navigationProperties, collectionItemsIfAny) {
            var result = {};
            for (var prop in scalarProperties) {
                void 0 !== (value = scalarProperties[prop]) && (result[prop] = value);
            }
            for (var prop in navigationProperties) {
                var value;
                void 0 !== (value = navigationProperties[prop]) && (value[Utility.fieldName_isCollection] && void 0 !== value[Utility.fieldName_m__items] ? result[prop] = value.toJSON().items : result[prop] = value.toJSON());
            }
            return collectionItemsIfAny && (result.items = collectionItemsIfAny.map((function(item) {
                return item.toJSON();
            }))), result;
        }, Utility.throwError = function(resourceId, arg, errorLocation) {
            throw new Core._Internal.RuntimeError({
                code: resourceId,
                httpStatusCode: 400,
                message: Core.CoreUtility._getResourceString(resourceId, arg),
                debugInfo: errorLocation ? {
                    errorLocation: errorLocation
                } : void 0
            });
        }, Utility.createRuntimeError = function(code, message, location, httpStatusCode, data) {
            return new Core._Internal.RuntimeError({
                code: code,
                httpStatusCode: httpStatusCode,
                message: message,
                debugInfo: {
                    errorLocation: location
                },
                data: data
            });
        }, Utility.throwIfNotLoaded = function(propertyName, fieldValue, entityName, isNull) {
            if (!isNull && Core.CoreUtility.isUndefined(fieldValue) && propertyName.charCodeAt(0) != Utility.s_underscoreCharCode) throw Utility.createPropertyNotLoadedException(entityName, propertyName);
        }, Utility.createPropertyNotLoadedException = function(entityName, propertyName) {
            return new Core._Internal.RuntimeError({
                code: ErrorCodes.propertyNotLoaded,
                httpStatusCode: 400,
                message: Core.CoreUtility._getResourceString(ResourceStrings.propertyNotLoaded, propertyName),
                debugInfo: entityName ? {
                    errorLocation: entityName + "." + propertyName
                } : void 0
            });
        }, Utility.createCannotUpdateReadOnlyPropertyException = function(entityName, propertyName) {
            return new Core._Internal.RuntimeError({
                code: ErrorCodes.cannotUpdateReadOnlyProperty,
                httpStatusCode: 400,
                message: Core.CoreUtility._getResourceString(ResourceStrings.cannotUpdateReadOnlyProperty, propertyName),
                debugInfo: entityName ? {
                    errorLocation: entityName + "." + propertyName
                } : void 0
            });
        }, Utility.promisify = function(action) {
            return new Promise((function(resolve, reject) {
                action((function(result) {
                    "failed" == result.status ? reject(result.error) : resolve(result.value);
                }));
            }));
        }, Utility._addActionResultHandler = function(clientObj, action, resultHandler) {
            clientObj.context._pendingRequest.addActionResultHandler(action, resultHandler);
        }, Utility._handleNavigationPropertyResults = function(clientObj, objectValue, propertyNames) {
            for (var i = 0; i < propertyNames.length - 1; i += 2) Core.CoreUtility.isUndefined(objectValue[propertyNames[i + 1]]) || clientObj[propertyNames[i]]._handleResult(objectValue[propertyNames[i + 1]]);
        }, Utility._fixupApiFlags = function(flags) {
            return "boolean" == typeof flags && (flags = flags ? 1 : 0), flags;
        }, Utility.definePropertyThrowUnloadedException = function(obj, typeName, propertyName) {
            Object.defineProperty(obj, propertyName, {
                configurable: !0,
                enumerable: !0,
                get: function() {
                    throw Utility.createPropertyNotLoadedException(typeName, propertyName);
                },
                set: function() {
                    throw Utility.createCannotUpdateReadOnlyPropertyException(typeName, propertyName);
                }
            });
        }, Utility.defineReadOnlyPropertyWithValue = function(obj, propertyName, value) {
            Object.defineProperty(obj, propertyName, {
                configurable: !0,
                enumerable: !0,
                get: function() {
                    return value;
                },
                set: function() {
                    throw Utility.createCannotUpdateReadOnlyPropertyException(null, propertyName);
                }
            });
        }, Utility.processRetrieveResult = function(proxy, value, result, childItemCreateFunc) {
            if (!Core.CoreUtility.isNullOrUndefined(value)) if (childItemCreateFunc) {
                var data = value[Constants.itemsLowerCase];
                if (Array.isArray(data)) {
                    for (var itemsResult = [], i = 0; i < data.length; i++) {
                        var itemProxy = childItemCreateFunc(data[i], i), itemResult = {};
                        itemResult[Constants.proxy] = itemProxy, itemProxy._handleRetrieveResult(data[i], itemResult), 
                        itemsResult.push(itemResult);
                    }
                    Utility.defineReadOnlyPropertyWithValue(result, Constants.itemsLowerCase, itemsResult);
                }
            } else {
                var scalarPropertyNames = proxy[Constants.scalarPropertyNames], navigationPropertyNames = proxy[Constants.navigationPropertyNames], typeName = proxy[Constants.className];
                if (scalarPropertyNames) for (i = 0; i < scalarPropertyNames.length; i++) {
                    var propValue = value[propName = scalarPropertyNames[i]];
                    Core.CoreUtility.isUndefined(propValue) ? Utility.definePropertyThrowUnloadedException(result, typeName, propName) : Utility.defineReadOnlyPropertyWithValue(result, propName, propValue);
                }
                if (navigationPropertyNames) for (i = 0; i < navigationPropertyNames.length; i++) {
                    var propName;
                    propValue = value[propName = navigationPropertyNames[i]];
                    if (Core.CoreUtility.isUndefined(propValue)) Utility.definePropertyThrowUnloadedException(result, typeName, propName); else {
                        var propProxy = proxy[propName], propResult = {};
                        propProxy._handleRetrieveResult(propValue, propResult), propResult[Constants.proxy] = propProxy, 
                        Array.isArray(propResult[Constants.itemsLowerCase]) && (propResult = propResult[Constants.itemsLowerCase]), 
                        Utility.defineReadOnlyPropertyWithValue(result, propName, propResult);
                    }
                }
            }
        }, Utility.setMockData = function(clientObj, value, childItemCreateFunc, setItemsFunc) {
            if (Core.CoreUtility.isNullOrUndefined(value)) clientObj._handleResult(value); else {
                if (clientObj[Constants.scalarPropertyOriginalNames]) {
                    for (var result = {}, scalarPropertyOriginalNames = clientObj[Constants.scalarPropertyOriginalNames], scalarPropertyNames = clientObj[Constants.scalarPropertyNames], i = 0; i < scalarPropertyNames.length; i++) void 0 !== value[scalarPropertyNames[i]] && (result[scalarPropertyOriginalNames[i]] = value[scalarPropertyNames[i]]);
                    clientObj._handleResult(result);
                }
                if (clientObj[Constants.navigationPropertyNames]) {
                    var navigationPropertyNames = clientObj[Constants.navigationPropertyNames];
                    for (i = 0; i < navigationPropertyNames.length; i++) if (void 0 !== value[navigationPropertyNames[i]]) {
                        var navigationPropValue = clientObj[navigationPropertyNames[i]];
                        navigationPropValue.setMockData && navigationPropValue.setMockData(value[navigationPropertyNames[i]]);
                    }
                }
                if (clientObj[Constants.isCollection] && childItemCreateFunc) {
                    var itemsData = Array.isArray(value) ? value : value[Constants.itemsLowerCase];
                    if (Array.isArray(itemsData)) {
                        var items = [];
                        for (i = 0; i < itemsData.length; i++) {
                            var item = childItemCreateFunc(itemsData, i);
                            Utility.setMockData(item, itemsData[i]), items.push(item);
                        }
                        setItemsFunc(items);
                    }
                }
            }
        }, Utility.applyMixin = function(derived, base) {
            Object.getOwnPropertyNames(base.prototype).forEach((function(name) {
                "constructor" !== name && Object.defineProperty(derived.prototype, name, Object.getOwnPropertyDescriptor(base.prototype, name));
            }));
        }, Utility.fieldName_m__items = "m__items", Utility.fieldName_isCollection = "_isCollection", 
        Utility._synchronousCleanup = !1, Utility.s_underscoreCharCode = "_".charCodeAt(0), 
        Utility;
    }(Common.CommonUtility);
    exports.Utility = Utility;
}, function(module, exports) {
    var wrap, __assign = this && this.__assign || function() {
        return (__assign = Object.assign || function(t) {
            for (var s, i = 1, n = arguments.length; i < n; i++) for (var p in s = arguments[i]) Object.prototype.hasOwnProperty.call(s, p) && (t[p] = s[p]);
            return t;
        }).apply(this, arguments);
    };
    function getOfficeRuntimeImplementation(info) {
        return window[info.platform === Office.PlatformType.OfficeOnline ? "_OfficeRuntimeWeb" : "_OfficeRuntimeNative"];
    }
    function getWrapMethod(property) {
        return function(name) {
            return function(property, method) {
                return wrapMethod((function(impl) {
                    return impl[property][method];
                }));
            }(property, name);
        };
    }
    function wrapMethod(methodFetcher) {
        return function() {
            var _this = this, args = arguments;
            return Office.onReady().then((function(info) {
                return methodFetcher(getOfficeRuntimeImplementation(info)).apply(_this, args);
            }));
        };
    }
    Office.onReady((function(info) {
        window.OfficeRuntime = __assign(__assign({}, window.OfficeRuntime), getOfficeRuntimeImplementation(info));
    })), window.OfficeRuntime = {
        AsyncStorage: function() {
            return {
                getItem: wrapStorageMethod("getItem"),
                setItem: wrapStorageMethod("setItem"),
                removeItem: wrapStorageMethod("removeItem"),
                getAllKeys: wrapStorageMethod("getAllKeys"),
                multiSet: wrapStorageMethod("multiSet"),
                multiRemove: wrapStorageMethod("multiRemove"),
                multiGet: wrapStorageMethod("multiGet")
            };
            function wrapStorageMethod(methodName) {
                return wrapMethod((function(impl) {
                    return impl.storage[methodName];
                }));
            }
        }(),
        displayWebDialog: wrapMethod((function(impl) {
            return impl.displayWebDialog;
        })),
        storage: function() {
            return {
                getItem: wrapStorageMethod("getItem"),
                setItem: wrapStorageMethod("setItem"),
                removeItem: wrapStorageMethod("removeItem"),
                getKeys: wrapStorageMethod("getKeys"),
                setItems: wrapStorageMethod("setItems"),
                removeItems: wrapStorageMethod("removeItems"),
                getItems: wrapStorageMethod("getItems")
            };
            function wrapStorageMethod(methodName) {
                return wrapMethod((function(impl) {
                    return impl.storage[methodName];
                }));
            }
        }(),
        experimentation: function() {
            return {
                getBooleanFeatureGateAsync: wrapExperimentationMethod("getBooleanFeatureGateAsync"),
                getIntFeatureGateAsync: wrapExperimentationMethod("getIntFeatureGateAsync"),
                getStringFeatureGateAsync: wrapExperimentationMethod("getStringFeatureGateAsync")
            };
            function wrapExperimentationMethod(methodName) {
                return wrapMethod((function(impl) {
                    return impl.experimentation[methodName];
                }));
            }
        }(),
        apiInformation: {
            isSetSupported: function(capability, semVer) {
                return Office.context.requirements.isSetSupported(capability, Number(semVer));
            }
        },
        message: (wrap = getWrapMethod("message"), {
            on: wrap("on"),
            off: wrap("off"),
            emit: wrap("emit")
        }),
        auth: {
            getAccessToken: getWrapMethod("auth")("getAccessToken")
        },
        ui: {
            getRibbon: getWrapMethod("ui")("getRibbon")
        }
    };
}, function(module, exports) {
    !function(self) {
        "use strict";
        if (!self.fetch) {
            var support_searchParams = "URLSearchParams" in self, support_iterable = "Symbol" in self && "iterator" in Symbol, support_blob = "FileReader" in self && "Blob" in self && function() {
                try {
                    return new Blob, !0;
                } catch (e) {
                    return !1;
                }
            }(), support_formData = "FormData" in self, support_arrayBuffer = "ArrayBuffer" in self;
            if (support_arrayBuffer) var viewClasses = [ "[object Int8Array]", "[object Uint8Array]", "[object Uint8ClampedArray]", "[object Int16Array]", "[object Uint16Array]", "[object Int32Array]", "[object Uint32Array]", "[object Float32Array]", "[object Float64Array]" ], isDataView = function(obj) {
                return obj && DataView.prototype.isPrototypeOf(obj);
            }, isArrayBufferView = ArrayBuffer.isView || function(obj) {
                return obj && viewClasses.indexOf(Object.prototype.toString.call(obj)) > -1;
            };
            Headers.prototype.append = function(name, value) {
                name = normalizeName(name), value = normalizeValue(value);
                var oldValue = this.map[name];
                this.map[name] = oldValue ? oldValue + "," + value : value;
            }, Headers.prototype.delete = function(name) {
                delete this.map[normalizeName(name)];
            }, Headers.prototype.get = function(name) {
                return name = normalizeName(name), this.has(name) ? this.map[name] : null;
            }, Headers.prototype.has = function(name) {
                return this.map.hasOwnProperty(normalizeName(name));
            }, Headers.prototype.set = function(name, value) {
                this.map[normalizeName(name)] = normalizeValue(value);
            }, Headers.prototype.forEach = function(callback, thisArg) {
                for (var name in this.map) this.map.hasOwnProperty(name) && callback.call(thisArg, this.map[name], name, this);
            }, Headers.prototype.keys = function() {
                var items = [];
                return this.forEach((function(value, name) {
                    items.push(name);
                })), iteratorFor(items);
            }, Headers.prototype.values = function() {
                var items = [];
                return this.forEach((function(value) {
                    items.push(value);
                })), iteratorFor(items);
            }, Headers.prototype.entries = function() {
                var items = [];
                return this.forEach((function(value, name) {
                    items.push([ name, value ]);
                })), iteratorFor(items);
            }, support_iterable && (Headers.prototype[Symbol.iterator] = Headers.prototype.entries);
            var methods = [ "DELETE", "GET", "HEAD", "OPTIONS", "POST", "PUT" ];
            Request.prototype.clone = function() {
                return new Request(this, {
                    body: this._bodyInit
                });
            }, Body.call(Request.prototype), Body.call(Response.prototype), Response.prototype.clone = function() {
                return new Response(this._bodyInit, {
                    status: this.status,
                    statusText: this.statusText,
                    headers: new Headers(this.headers),
                    url: this.url
                });
            }, Response.error = function() {
                var response = new Response(null, {
                    status: 0,
                    statusText: ""
                });
                return response.type = "error", response;
            };
            var redirectStatuses = [ 301, 302, 303, 307, 308 ];
            Response.redirect = function(url, status) {
                if (-1 === redirectStatuses.indexOf(status)) throw new RangeError("Invalid status code");
                return new Response(null, {
                    status: status,
                    headers: {
                        location: url
                    }
                });
            }, self.Headers = Headers, self.Request = Request, self.Response = Response, self.fetch = function(input, init) {
                return new Promise((function(resolve, reject) {
                    var request = new Request(input, init), xhr = new XMLHttpRequest;
                    xhr.onload = function() {
                        var rawHeaders, headers, options = {
                            status: xhr.status,
                            statusText: xhr.statusText,
                            headers: (rawHeaders = xhr.getAllResponseHeaders() || "", headers = new Headers, 
                            rawHeaders.split(/\r?\n/).forEach((function(line) {
                                var parts = line.split(":"), key = parts.shift().trim();
                                if (key) {
                                    var value = parts.join(":").trim();
                                    headers.append(key, value);
                                }
                            })), headers)
                        };
                        options.url = "responseURL" in xhr ? xhr.responseURL : options.headers.get("X-Request-URL");
                        var body = "response" in xhr ? xhr.response : xhr.responseText;
                        resolve(new Response(body, options));
                    }, xhr.onerror = function() {
                        reject(new TypeError("Network request failed"));
                    }, xhr.ontimeout = function() {
                        reject(new TypeError("Network request failed"));
                    }, xhr.open(request.method, request.url, !0), "include" === request.credentials && (xhr.withCredentials = !0), 
                    "responseType" in xhr && support_blob && (xhr.responseType = "blob"), request.headers.forEach((function(value, name) {
                        xhr.setRequestHeader(name, value);
                    })), xhr.send(void 0 === request._bodyInit ? null : request._bodyInit);
                }));
            }, self.fetch.polyfill = !0;
        }
        function normalizeName(name) {
            if ("string" != typeof name && (name = String(name)), /[^a-z0-9\-#$%&'*+.\^_`|~]/i.test(name)) throw new TypeError("Invalid character in header field name");
            return name.toLowerCase();
        }
        function normalizeValue(value) {
            return "string" != typeof value && (value = String(value)), value;
        }
        function iteratorFor(items) {
            var iterator = {
                next: function() {
                    var value = items.shift();
                    return {
                        done: void 0 === value,
                        value: value
                    };
                }
            };
            return support_iterable && (iterator[Symbol.iterator] = function() {
                return iterator;
            }), iterator;
        }
        function Headers(headers) {
            this.map = {}, headers instanceof Headers ? headers.forEach((function(value, name) {
                this.append(name, value);
            }), this) : Array.isArray(headers) ? headers.forEach((function(header) {
                this.append(header[0], header[1]);
            }), this) : headers && Object.getOwnPropertyNames(headers).forEach((function(name) {
                this.append(name, headers[name]);
            }), this);
        }
        function consumed(body) {
            if (body.bodyUsed) return Promise.reject(new TypeError("Already read"));
            body.bodyUsed = !0;
        }
        function fileReaderReady(reader) {
            return new Promise((function(resolve, reject) {
                reader.onload = function() {
                    resolve(reader.result);
                }, reader.onerror = function() {
                    reject(reader.error);
                };
            }));
        }
        function readBlobAsArrayBuffer(blob) {
            var reader = new FileReader, promise = fileReaderReady(reader);
            return reader.readAsArrayBuffer(blob), promise;
        }
        function bufferClone(buf) {
            if (buf.slice) return buf.slice(0);
            var view = new Uint8Array(buf.byteLength);
            return view.set(new Uint8Array(buf)), view.buffer;
        }
        function Body() {
            return this.bodyUsed = !1, this._initBody = function(body) {
                if (this._bodyInit = body, body) if ("string" == typeof body) this._bodyText = body; else if (support_blob && Blob.prototype.isPrototypeOf(body)) this._bodyBlob = body; else if (support_formData && FormData.prototype.isPrototypeOf(body)) this._bodyFormData = body; else if (support_searchParams && URLSearchParams.prototype.isPrototypeOf(body)) this._bodyText = body.toString(); else if (support_arrayBuffer && support_blob && isDataView(body)) this._bodyArrayBuffer = bufferClone(body.buffer), 
                this._bodyInit = new Blob([ this._bodyArrayBuffer ]); else {
                    if (!support_arrayBuffer || !ArrayBuffer.prototype.isPrototypeOf(body) && !isArrayBufferView(body)) throw new Error("unsupported BodyInit type");
                    this._bodyArrayBuffer = bufferClone(body);
                } else this._bodyText = "";
                this.headers.get("content-type") || ("string" == typeof body ? this.headers.set("content-type", "text/plain;charset=UTF-8") : this._bodyBlob && this._bodyBlob.type ? this.headers.set("content-type", this._bodyBlob.type) : support_searchParams && URLSearchParams.prototype.isPrototypeOf(body) && this.headers.set("content-type", "application/x-www-form-urlencoded;charset=UTF-8"));
            }, support_blob && (this.blob = function() {
                var rejected = consumed(this);
                if (rejected) return rejected;
                if (this._bodyBlob) return Promise.resolve(this._bodyBlob);
                if (this._bodyArrayBuffer) return Promise.resolve(new Blob([ this._bodyArrayBuffer ]));
                if (this._bodyFormData) throw new Error("could not read FormData body as blob");
                return Promise.resolve(new Blob([ this._bodyText ]));
            }, this.arrayBuffer = function() {
                return this._bodyArrayBuffer ? consumed(this) || Promise.resolve(this._bodyArrayBuffer) : this.blob().then(readBlobAsArrayBuffer);
            }), this.text = function() {
                var blob, reader, promise, rejected = consumed(this);
                if (rejected) return rejected;
                if (this._bodyBlob) return blob = this._bodyBlob, reader = new FileReader, promise = fileReaderReady(reader), 
                reader.readAsText(blob), promise;
                if (this._bodyArrayBuffer) return Promise.resolve(function(buf) {
                    for (var view = new Uint8Array(buf), chars = new Array(view.length), i = 0; i < view.length; i++) chars[i] = String.fromCharCode(view[i]);
                    return chars.join("");
                }(this._bodyArrayBuffer));
                if (this._bodyFormData) throw new Error("could not read FormData body as text");
                return Promise.resolve(this._bodyText);
            }, support_formData && (this.formData = function() {
                return this.text().then(decode);
            }), this.json = function() {
                return this.text().then(JSON.parse);
            }, this;
        }
        function Request(input, options) {
            var method, upcased, body = (options = options || {}).body;
            if (input instanceof Request) {
                if (input.bodyUsed) throw new TypeError("Already read");
                this.url = input.url, this.credentials = input.credentials, options.headers || (this.headers = new Headers(input.headers)), 
                this.method = input.method, this.mode = input.mode, body || null == input._bodyInit || (body = input._bodyInit, 
                input.bodyUsed = !0);
            } else this.url = String(input);
            if (this.credentials = options.credentials || this.credentials || "omit", !options.headers && this.headers || (this.headers = new Headers(options.headers)), 
            this.method = (method = options.method || this.method || "GET", upcased = method.toUpperCase(), 
            methods.indexOf(upcased) > -1 ? upcased : method), this.mode = options.mode || this.mode || null, 
            this.referrer = null, ("GET" === this.method || "HEAD" === this.method) && body) throw new TypeError("Body not allowed for GET or HEAD requests");
            this._initBody(body);
        }
        function decode(body) {
            var form = new FormData;
            return body.trim().split("&").forEach((function(bytes) {
                if (bytes) {
                    var split = bytes.split("="), name = split.shift().replace(/\+/g, " "), value = split.join("=").replace(/\+/g, " ");
                    form.append(decodeURIComponent(name), decodeURIComponent(value));
                }
            })), form;
        }
        function Response(bodyInit, options) {
            options || (options = {}), this.type = "default", this.status = "status" in options ? options.status : 200, 
            this.ok = this.status >= 200 && this.status < 300, this.statusText = "statusText" in options ? options.statusText : "OK", 
            this.headers = new Headers(options.headers), this.url = options.url || "", this._initBody(bodyInit);
        }
    }("undefined" != typeof self ? self : this);
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var CFRuntime = __webpack_require__(8);
    exports.default = function(shouldInitRuntime) {
        function documentReadyCallback() {
            Office.onReady((function(hostInfo) {
                hostInfo.host === Office.HostType.Excel ? function initializeCustomFunctionsOrDelay() {
                    CustomFunctionMappings && CustomFunctionMappings.__delay__ ? setTimeout(initializeCustomFunctionsOrDelay, 50) : CFRuntime.CustomFunctions.initialize();
                }() : console.warn("Warning: Expected to be loaded inside of an Excel add-in.");
            }));
        }
        window.CustomFunctions = window.CustomFunctions || {}, window.CustomFunctions.setCustomFunctionInvoker = CFRuntime.setCustomFunctionInvoker, 
        window.CustomFunctions.Error = CFRuntime.CustomFunctionError, window.CustomFunctions.ErrorCode = CFRuntime.ErrorCode, 
        CFRuntime.setCustomFunctionAssociation(window.CustomFunctions._association), shouldInitRuntime && ("loading" === document.readyState ? document.addEventListener("DOMContentLoaded", documentReadyCallback) : documentReadyCallback());
    };
}, function(module, exports, __webpack_require__) {
    "use strict";
    var extendStatics, __extends = this && this.__extends || (extendStatics = function(d, b) {
        return (extendStatics = Object.setPrototypeOf || {
            __proto__: []
        } instanceof Array && function(d, b) {
            d.__proto__ = b;
        } || function(d, b) {
            for (var p in b) b.hasOwnProperty(p) && (d[p] = b[p]);
        })(d, b);
    }, function(d, b) {
        function __() {
            this.constructor = d;
        }
        extendStatics(d, b), d.prototype = null === b ? Object.create(b) : (__.prototype = b.prototype, 
        new __);
    });
    Object.defineProperty(exports, "__esModule", {
        value: !0
    }), exports.CustomFunctionsContainer = exports.CustomFunctions = exports.setCustomFunctionInvoker = exports.setCustomFunctionAssociation = exports.customFunctionProxy = exports.CustomFunctionProxy = exports.CustomFunctionError = exports.ErrorCode = exports.InvocationContext = exports.Script = void 0;
    var OfficeExtension = __webpack_require__(2), Core = __webpack_require__(0), _createPropertyObject = OfficeExtension.BatchApiHelper.createPropertyObject, _createRootServiceObject = (OfficeExtension.BatchApiHelper.createMethodObject, 
    OfficeExtension.BatchApiHelper.createIndexerObject, OfficeExtension.BatchApiHelper.createRootServiceObject), _createTopLevelServiceObject = OfficeExtension.BatchApiHelper.createTopLevelServiceObject, _invokeMethod = (OfficeExtension.BatchApiHelper.createChildItemObject, 
    OfficeExtension.BatchApiHelper.invokeMethod), _isNullOrUndefined = (OfficeExtension.BatchApiHelper.invokeEnsureUnchanged, 
    OfficeExtension.BatchApiHelper.invokeSetProperty, OfficeExtension.Utility.isNullOrUndefined), _throwIfApiNotSupported = (OfficeExtension.Utility.isUndefined, 
    OfficeExtension.Utility.throwIfNotLoaded, OfficeExtension.Utility.throwIfApiNotSupported), _load = OfficeExtension.Utility.load, _toJson = (OfficeExtension.Utility.retrieve, 
    OfficeExtension.Utility.toJson), _fixObjectPathIfNecessary = OfficeExtension.Utility.fixObjectPathIfNecessary, _handleNavigationPropertyResults = OfficeExtension.Utility._handleNavigationPropertyResults, _processRetrieveResult = (OfficeExtension.Utility.adjustToDateTime, 
    OfficeExtension.Utility.processRetrieveResult), CustomFunctionRequestContext = (OfficeExtension.Utility.setMockData, 
    OfficeExtension.CommonUtility.calculateApiFlags, function(_super) {
        function CustomFunctionRequestContext(requestInfo) {
            var _this = _super.call(this, requestInfo) || this;
            return _this.m_customFunctions = CustomFunctions.newObject(_this), _this.m_container = _createRootServiceObject(CustomFunctionsContainer, _this), 
            _this._rootObject = _this.m_container, _this._rootObjectPropertyName = "customFunctionsContainer", 
            _this._requestFlagModifier = 2176, _this;
        }
        return __extends(CustomFunctionRequestContext, _super), Object.defineProperty(CustomFunctionRequestContext.prototype, "customFunctions", {
            get: function() {
                return this.m_customFunctions;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(CustomFunctionRequestContext.prototype, "customFunctionsContainer", {
            get: function() {
                return this.m_container;
            },
            enumerable: !1,
            configurable: !0
        }), CustomFunctionRequestContext.prototype._processOfficeJsErrorResponse = function(officeJsErrorCode, response) {
            5004 === officeJsErrorCode && (response.ErrorCode = CustomFunctionRuntimeErrorCode.invalidOperationInCellEditMode, 
            response.ErrorMessage = OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidOperationInCellEditMode));
        }, CustomFunctionRequestContext;
    }(OfficeExtension.ClientRequestContext));
    exports.Script = {
        _CustomFunctionMetadata: {}
    };
    var CustomFunctionLoggingSeverity, InvocationContext = function() {
        function InvocationContext(functionName, address, parameterAddresses, setResultHandler, setErrorHandler) {
            this._functionName = functionName, _isNullOrUndefined(address) || (this._address = address), 
            _isNullOrUndefined(parameterAddresses) || (this._parameterAddresses = parameterAddresses), 
            this.setResult = setResultHandler, this.setError = setErrorHandler;
        }
        return Object.defineProperty(InvocationContext.prototype, "onCanceled", {
            get: function() {
                if (!_isNullOrUndefined(this._onCanceled) && "function" == typeof this._onCanceled) return this._onCanceled;
            },
            set: function(handler) {
                this._onCanceled = handler;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(InvocationContext.prototype, "functionName", {
            get: function() {
                return this._functionName;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(InvocationContext.prototype, "address", {
            get: function() {
                return this._address;
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(InvocationContext.prototype, "parameterAddresses", {
            get: function() {
                return this._parameterAddresses;
            },
            enumerable: !1,
            configurable: !0
        }), InvocationContext;
    }();
    exports.InvocationContext = InvocationContext, function(CustomFunctionLoggingSeverity) {
        CustomFunctionLoggingSeverity.Info = "Medium", CustomFunctionLoggingSeverity.Error = "Unexpected", 
        CustomFunctionLoggingSeverity.Verbose = "Verbose";
    }(CustomFunctionLoggingSeverity || (CustomFunctionLoggingSeverity = {}));
    var ErrorCode, CustomFunctionLog = function(Severity, Message) {
        this.Severity = Severity, this.Message = Message;
    }, CustomFunctionsLogger = function() {
        function CustomFunctionsLogger() {}
        return CustomFunctionsLogger.logEvent = function(log, data, data2) {
            if (CustomFunctionsLogger.s_shouldLog || OfficeExtension.CoreUtility._logEnabled) {
                var logMessage = log.Severity + " " + log.Message + data;
                if (data2 && (logMessage = logMessage + " " + data2), OfficeExtension.Utility.log(logMessage), 
                CustomFunctionsLogger.s_shouldLog) switch (log.Severity) {
                  case CustomFunctionLoggingSeverity.Verbose:
                    null !== console.log && console.log(logMessage);
                    break;

                  case CustomFunctionLoggingSeverity.Info:
                    null !== console.info && console.info(logMessage);
                    break;

                  case CustomFunctionLoggingSeverity.Error:
                    null !== console.error && console.error(logMessage);
                }
            }
        }, CustomFunctionsLogger.shouldLog = function() {
            try {
                return !_isNullOrUndefined(console) && !_isNullOrUndefined(window) && window.name && "string" == typeof window.name && JSON.parse(window.name)[CustomFunctionsLogger.CustomFunctionLoggingFlag];
            } catch (ex) {
                return OfficeExtension.Utility.log(JSON.stringify(ex)), !1;
            }
        }, CustomFunctionsLogger.CustomFunctionLoggingFlag = "CustomFunctionsRuntimeLogging", 
        CustomFunctionsLogger.s_shouldLog = CustomFunctionsLogger.shouldLog(), CustomFunctionsLogger;
    }();
    !function(ErrorCode) {
        ErrorCode.invalidValue = "#VALUE!", ErrorCode.notAvailable = "#N/A", ErrorCode.divisionByZero = "#DIV/0!", 
        ErrorCode.invalidNumber = "#NUM!", ErrorCode.nullReference = "#NULL!", ErrorCode.invalidName = "#NAME?", 
        ErrorCode.invalidReference = "#REF!";
    }(ErrorCode = exports.ErrorCode || (exports.ErrorCode = {}));
    var CustomFunctionError = function(_super) {
        function CustomFunctionError(errorCode, errorMessage) {
            var _this = _super.call(this, errorMessage || "") || this;
            return Object.setPrototypeOf(_this, CustomFunctionError.prototype), _this.valueType = "CustomFunctionError", 
            _this.code = errorCode || ErrorCode.invalidValue, _this.stack = void 0, _this;
        }
        return __extends(CustomFunctionError, _super), CustomFunctionError.prototype.toJSON = function() {
            return {
                valueType: this.valueType,
                code: this.code,
                message: this.message
            };
        }, CustomFunctionError;
    }(Error);
    exports.CustomFunctionError = CustomFunctionError;
    var CustomFunctionProxy = function() {
        function CustomFunctionProxy() {
            this._whenInit = void 0, this._isInit = !1, this._setResultsDelayMillis = 50, this._setResultsOverdueDelayMillis = 2e3, 
            this._maxContextSyncExecutionDurationMills = 15e3, this._minContextSyncIntervalMills = 500, 
            this._setResultsLifeMillis = 6e4, this._ensureInitRetryDelayMillis = 500, this._resultEntryBuffer = {}, 
            this._resultEntryRetryBuffer = {}, this._retryBufferBodySize = 0, this._isSetResultsTaskScheduled = !1, 
            this._setResultsTaskOverdueTime = 0, this._inProgressContextSyncExpectedFinishTime = 0, 
            this._batchQuotaMillis = 1e3, this._invocationContextMap = {}, this._maxResultSize = 2e6;
        }
        return CustomFunctionProxy.prototype._initSettings = function() {
            if ("object" == typeof exports.Script && "object" == typeof exports.Script._CustomFunctionSettings) {
                if ("number" == typeof exports.Script._CustomFunctionSettings.setResultsDelayMillis) {
                    var setResultsDelayMillis = exports.Script._CustomFunctionSettings.setResultsDelayMillis;
                    setResultsDelayMillis = Math.max(0, setResultsDelayMillis), setResultsDelayMillis = Math.min(1e3, setResultsDelayMillis), 
                    this._setResultsDelayMillis = setResultsDelayMillis;
                }
                if ("number" == typeof exports.Script._CustomFunctionSettings.ensureInitRetryDelayMillis) {
                    var ensureInitRetryDelayMillis = exports.Script._CustomFunctionSettings.ensureInitRetryDelayMillis;
                    ensureInitRetryDelayMillis = Math.max(0, ensureInitRetryDelayMillis), ensureInitRetryDelayMillis = Math.min(2e3, ensureInitRetryDelayMillis), 
                    this._ensureInitRetryDelayMillis = ensureInitRetryDelayMillis;
                }
                if ("number" == typeof exports.Script._CustomFunctionSettings.setResultsLifeMillis) {
                    var setResultsLifeMillis = exports.Script._CustomFunctionSettings.setResultsLifeMillis;
                    setResultsLifeMillis = Math.max(0, setResultsLifeMillis), setResultsLifeMillis = Math.min(6e5, setResultsLifeMillis), 
                    this._setResultsLifeMillis = setResultsLifeMillis;
                }
                if ("number" == typeof exports.Script._CustomFunctionSettings.batchQuotaMillis) {
                    var batchQuotaMillis = exports.Script._CustomFunctionSettings.batchQuotaMillis;
                    batchQuotaMillis = Math.max(0, batchQuotaMillis), batchQuotaMillis = Math.min(1e3, batchQuotaMillis), 
                    this._batchQuotaMillis = batchQuotaMillis;
                }
            }
        }, CustomFunctionProxy.prototype.ensureInit = function(context) {
            var _this = this;
            return this._initSettings(), void 0 === this._whenInit && (this._whenInit = OfficeExtension.Utility._createPromiseFromResult(null).then((function() {
                if (!_this._isInit) return context.eventRegistration.register(5, "", _this._handleMessage.bind(_this));
            })).then((function() {
                _this._isInit = !0;
            }))), this._isInit || context._pendingRequest._addPreSyncPromise(this._whenInit), 
            this._whenInit;
        }, CustomFunctionProxy.prototype.setCustomFunctionInvoker = function(invoker) {
            "object" == typeof CustomFunctionMappings && delete CustomFunctionMappings.__delay__, 
            this._invoker = invoker;
        }, CustomFunctionProxy.prototype.setCustomFunctionAssociation = function(association) {
            var _this = this;
            this._customFunctionMappingsUpperCase = void 0, this._association = association, 
            this._association && this._association.onchange((function() {
                _this._customFunctionMappingsUpperCase = void 0;
            }));
        }, CustomFunctionProxy.prototype._initFromHostBridge = function(hostBridge) {
            var _this = this;
            this._initSettings(), hostBridge.addHostMessageHandler((function(bridgeMessage) {
                3 === bridgeMessage.type ? _this._handleMessage(bridgeMessage.message) : 4 === bridgeMessage.type && _this._handleSettings(bridgeMessage.message);
            })), this._isInit = !0, this._whenInit = OfficeExtension.CoreUtility.Promise.resolve();
        }, CustomFunctionProxy.prototype._handleSettings = function(args) {
            OfficeExtension.Utility.log("CustomFunctionProxy._handleSettings:" + JSON.stringify(args)), 
            args && "object" == typeof args && (CustomFunctionsLogger.s_shouldLog = args[CustomFunctionsLogger.CustomFunctionLoggingFlag]);
        }, CustomFunctionProxy.prototype._handleMessage = function(args) {
            try {
                OfficeExtension.Utility.log("CustomFunctionProxy._handleMessage"), OfficeExtension.Utility.checkArgumentNull(args, "args");
                for (var entryArray = args.entries, invocationArray = [], cancellationArray = [], metadataArray = [], i = 0; i < entryArray.length; i++) 1 === entryArray[i].messageCategory && ("string" == typeof entryArray[i].message && (entryArray[i].message = JSON.parse(entryArray[i].message)), 
                1e3 === entryArray[i].messageType ? invocationArray.push(entryArray[i]) : 1001 === entryArray[i].messageType ? cancellationArray.push(entryArray[i]) : 1002 === entryArray[i].messageType ? metadataArray.push(entryArray[i]) : OfficeExtension.Utility.log("CustomFunctionProxy._handleMessage unknown message type " + entryArray[i].messageType));
                if (metadataArray.length > 0 && this._handleMetadataEntries(metadataArray), invocationArray.length > 0) {
                    var batchArray = this._batchInvocationEntries(invocationArray);
                    batchArray.length > 0 && this._invokeRemainingBatchEntries(batchArray, 0);
                }
                cancellationArray.length > 0 && this._handleCancellationEntries(cancellationArray);
            } catch (ex) {
                throw CustomFunctionProxy._tryLog(ex), ex;
            }
            return OfficeExtension.Utility._createPromiseFromResult(null);
        }, CustomFunctionProxy.toLogMessage = function(ex) {
            var ret = "Unknown Error";
            if (ex) try {
                ex.toString && (ret = ex.toString()), ret = ret + " " + JSON.stringify(ex);
            } catch (otherEx) {
                ret = "Unexpected Error";
            }
            return ret;
        }, CustomFunctionProxy._tryLog = function(ex) {
            var message = CustomFunctionProxy.toLogMessage(ex);
            OfficeExtension.Utility.log(message);
        }, CustomFunctionProxy.prototype._handleMetadataEntries = function(entryArray) {
            for (var i = 0; i < entryArray.length; i++) {
                var message = entryArray[i].message;
                if (_isNullOrUndefined(message)) throw OfficeExtension.Utility.createRuntimeError(CustomFunctionRuntimeErrorCode.generalException, "message", "CustomFunctionProxy._handleMetadataEntries");
                exports.Script._CustomFunctionMetadata[message.functionName] = {
                    options: {
                        stream: message.isStream,
                        cancelable: message.isCancelable
                    }
                };
            }
        }, CustomFunctionProxy.prototype._handleCancellationEntries = function(entryArray) {
            for (var i = 0; i < entryArray.length; i++) {
                var message = entryArray[i].message;
                if (_isNullOrUndefined(message)) throw OfficeExtension.Utility.createRuntimeError(CustomFunctionRuntimeErrorCode.generalException, "message", "CustomFunctionProxy._handleCancellationEntries");
                var invocationId = message.invocationId, invocationContext = this._invocationContextMap[invocationId];
                _isNullOrUndefined(invocationContext) || (delete this._invocationContextMap[invocationId], 
                CustomFunctionsLogger.logEvent(CustomFunctionProxy.CustomFunctionCancellation, invocationContext.functionName), 
                _isNullOrUndefined(invocationContext.onCanceled) || invocationContext.onCanceled());
            }
        }, CustomFunctionProxy.prototype._transferCustomFunctionError = function(entryArray) {
            var _this = this;
            return entryArray.map((function(value, index) {
                return null === value ? value : null != value.code && "CustomFunctionError" == value.type ? new CustomFunctionError(value.code) : value instanceof Array ? _this._transferCustomFunctionError(value) : value;
            }));
        }, CustomFunctionProxy.prototype._batchInvocationEntries = function(entryArray) {
            for (var _this = this, batchArray = [], _loop_1 = function(i) {
                var message, arrayOrObjectMessage = entryArray[i].message;
                if (Array.isArray(arrayOrObjectMessage) ? (message = {
                    invocationId: arrayOrObjectMessage[0],
                    functionName: arrayOrObjectMessage[1],
                    parameterValues: arrayOrObjectMessage[2],
                    address: arrayOrObjectMessage[3],
                    flags: arrayOrObjectMessage[4],
                    parameterAddresses: null
                }, _isNullOrUndefined(arrayOrObjectMessage[5]) || (message.parameterAddresses = arrayOrObjectMessage[5])) : message = arrayOrObjectMessage, 
                _isNullOrUndefined(message)) throw OfficeExtension.Utility.createRuntimeError(CustomFunctionRuntimeErrorCode.generalException, "message", "CustomFunctionProxy._batchInvocationEntries");
                if (_isNullOrUndefined(message.invocationId) || message.invocationId < 0) throw OfficeExtension.Utility.createRuntimeError(CustomFunctionRuntimeErrorCode.generalException, "invocationId", "CustomFunctionProxy._batchInvocationEntries");
                if (_isNullOrUndefined(message.functionName)) throw OfficeExtension.Utility.createRuntimeError(CustomFunctionRuntimeErrorCode.generalException, "functionName", "CustomFunctionProxy._batchInvocationEntries");
                var call = null, isCancelable = !1, isStreaming = !1;
                if ("number" == typeof message.flags) isCancelable = 0 != (1 & message.flags), isStreaming = 0 != (2 & message.flags), 
                0 != (4 & message.flags) && (message.parameterValues = this_1._transferCustomFunctionError(message.parameterValues)); else {
                    var metadata = exports.Script._CustomFunctionMetadata[message.functionName];
                    if (_isNullOrUndefined(metadata)) return CustomFunctionsLogger.logEvent(CustomFunctionProxy.CustomFunctionExecutionNotFoundLog, message.functionName), 
                    OfficeExtension.Utility.isSetSupported("CustomFunctions", "1.8") ? this_1._setError(message.invocationId, null, 9) : this_1._setError(message.invocationId, "N/A", 1), 
                    "continue";
                    isCancelable = metadata.options.cancelable, isStreaming = metadata.options.stream;
                }
                if (this_1._invoker && !this_1._customFunctionMappingsContains(message.functionName)) return this_1._invokeFunctionUsingInvoker(message), 
                "continue";
                try {
                    call = this_1._getFunction(message.functionName);
                } catch (ex) {
                    return CustomFunctionsLogger.logEvent(CustomFunctionProxy.CustomFunctionExecutionNotFoundLog, message.functionName), 
                    OfficeExtension.Utility.isSetSupported("CustomFunctions", "1.8") ? this_1._setError(message.invocationId, ex, 10) : this_1._setError(message.invocationId, ex, 1), 
                    "continue";
                }
                var invocationContext = void 0;
                if (isStreaming || isCancelable) {
                    var setResult = void 0, setError = void 0;
                    isStreaming && (setResult = function(result) {
                        _this._invocationContextMap[message.invocationId] ? _this._setResult(message.invocationId, result) : CustomFunctionsLogger.logEvent(CustomFunctionProxy.CustomFunctionAlreadyCancelled, message.functionName);
                    }, setError = function(error) {
                        _this._invocationContextMap[message.invocationId] ? _this._setError(message.invocationId, error.message, _this._getCustomFunctionResultErrorCodeFromErrorCode(error.code)) : CustomFunctionsLogger.logEvent(CustomFunctionProxy.CustomFunctionAlreadyCancelled, message.functionName);
                    }), invocationContext = new InvocationContext(message.functionName, message.address, message.parameterAddresses, setResult, setError), 
                    this_1._invocationContextMap[message.invocationId] = invocationContext;
                } else invocationContext = new InvocationContext(message.functionName, message.address, message.parameterAddresses);
                message.parameterValues.push(invocationContext), batchArray.push({
                    call: call,
                    isBatching: !1,
                    isStreaming: isStreaming,
                    invocationIds: [ message.invocationId ],
                    parameterValueSets: [ message.parameterValues ],
                    functionName: message.functionName
                });
            }, this_1 = this, i = 0; i < entryArray.length; i++) _loop_1(i);
            return batchArray;
        }, CustomFunctionProxy.prototype._invokeFunctionUsingInvoker = function(message) {
            var _this = this, isCancelable = 0 != (1 & message.flags), isStreaming = 0 != (2 & message.flags), invocationId = message.invocationId, setResult = void 0, setError = void 0;
            if (isStreaming) setResult = function(result) {
                _this._invocationContextMap[invocationId] && _this._setResult(invocationId, result);
            }, setError = function(error) {
                _this._invocationContextMap[invocationId] && _this._setError(invocationId, error.message, _this._getCustomFunctionResultErrorCodeFromErrorCode(error.code));
            }; else {
                var setResultCalled_1 = !1;
                setResult = function(result) {
                    setResultCalled_1 || _this._setResult(invocationId, result), setResultCalled_1 = !0;
                }, setError = function(error) {
                    setResultCalled_1 || _this._setError(invocationId, error.message, _this._getCustomFunctionResultErrorCodeFromErrorCode(error.code)), 
                    setResultCalled_1 = !0;
                };
            }
            var invocationContext = new InvocationContext(message.functionName, message.address, message.parameterAddresses, setResult, setError);
            (isStreaming || isCancelable) && (this._invocationContextMap[invocationId] = invocationContext), 
            this._invoker.invoke(message.functionName, message.parameterValues, invocationContext);
        }, CustomFunctionProxy.prototype._ensureCustomFunctionMappingsUpperCase = function() {
            if (_isNullOrUndefined(this._customFunctionMappingsUpperCase)) {
                if (this._customFunctionMappingsUpperCase = {}, "object" == typeof CustomFunctionMappings) for (var key in OfficeExtension.CoreUtility.log("CustomFunctionMappings.Keys=" + JSON.stringify(Object.keys(CustomFunctionMappings))), 
                CustomFunctionMappings) this._customFunctionMappingsUpperCase[key.toUpperCase()] && CustomFunctionsLogger.logEvent(CustomFunctionProxy.CustomFunctionDuplicatedName, key), 
                this._customFunctionMappingsUpperCase[key.toUpperCase()] = CustomFunctionMappings[key];
                if (this._association) for (var key in OfficeExtension.CoreUtility.log("CustomFunctionAssociateMappings.Keys=" + JSON.stringify(Object.keys(this._association.mappings))), 
                this._association.mappings) this._customFunctionMappingsUpperCase[key.toUpperCase()] && CustomFunctionsLogger.logEvent(CustomFunctionProxy.CustomFunctionDuplicatedName, key), 
                this._customFunctionMappingsUpperCase[key.toUpperCase()] = this._association.mappings[key];
            }
        }, CustomFunctionProxy.prototype._customFunctionMappingsContains = function(functionName) {
            this._ensureCustomFunctionMappingsUpperCase();
            var functionNameUpperCase = functionName.toUpperCase();
            if (!_isNullOrUndefined(this._customFunctionMappingsUpperCase[functionNameUpperCase])) return !0;
            if ("undefined" != typeof window) {
                for (var functionParent = window, functionNameSegments = functionName.split("."), i = 0; i < functionNameSegments.length - 1; i++) if (functionParent = functionParent[functionNameSegments[i]], 
                _isNullOrUndefined(functionParent) || "object" != typeof functionParent) return !1;
                if ("function" == typeof functionParent[functionNameSegments[functionNameSegments.length - 1]]) return !0;
            }
            return !1;
        }, CustomFunctionProxy.prototype._getCustomFunctionMappings = function(functionName) {
            this._ensureCustomFunctionMappingsUpperCase();
            var functionNameUpperCase = functionName.toUpperCase();
            if (!_isNullOrUndefined(this._customFunctionMappingsUpperCase[functionNameUpperCase])) {
                if ("function" == typeof this._customFunctionMappingsUpperCase[functionNameUpperCase]) return this._customFunctionMappingsUpperCase[functionNameUpperCase];
                throw OfficeExtension.Utility.createRuntimeError(CustomFunctionRuntimeErrorCode.invalidOperation, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.customFunctionInvalidFunctionMapping, functionName), "CustomFunctionProxy._getCustomFunctionMappings");
            }
        }, CustomFunctionProxy.prototype._getFunction = function(functionName) {
            var call = this._getCustomFunctionMappings(functionName);
            if (!_isNullOrUndefined(call)) return call;
            if (_isNullOrUndefined(window)) throw OfficeExtension.Utility.createRuntimeError(CustomFunctionRuntimeErrorCode.invalidOperation, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.customFunctionWindowMissing), "CustomFunctionProxy._getFunction");
            for (var functionParent = window, functionNameSegments = functionName.split("."), i = 0; i < functionNameSegments.length - 1; i++) if (functionParent = functionParent[functionNameSegments[i]], 
            _isNullOrUndefined(functionParent) || "object" != typeof functionParent) throw OfficeExtension.Utility.createRuntimeError(CustomFunctionRuntimeErrorCode.invalidOperation, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.customFunctionDefintionMissingOnWindow, functionName), "CustomFunctionProxy._getFunction");
            if ("function" != typeof (call = functionParent[functionNameSegments[functionNameSegments.length - 1]])) throw OfficeExtension.Utility.createRuntimeError(CustomFunctionRuntimeErrorCode.invalidOperation, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.customFunctionInvalidFunction, functionName), "CustomFunctionProxy._getFunction");
            return call;
        }, CustomFunctionProxy.prototype._invokeRemainingBatchEntries = function(batchArray, startIndex) {
            OfficeExtension.Utility.log("CustomFunctionProxy._invokeRemainingBatchEntries");
            for (var startTimeMillis = Date.now(), i = startIndex; i < batchArray.length; i++) {
                if (!(Date.now() - startTimeMillis < this._batchQuotaMillis)) {
                    OfficeExtension.Utility.log("setTimeout(CustomFunctionProxy._invokeRemainingBatchEntries)"), 
                    setTimeout(this._invokeRemainingBatchEntries.bind(this), 0, batchArray, i);
                    break;
                }
                this._invokeFunctionAndSetResult(batchArray[i]);
            }
        }, CustomFunctionProxy.prototype._invokeFunctionAndSetResult = function(batch) {
            var results, _this = this;
            CustomFunctionsLogger.logEvent(CustomFunctionProxy.CustomFunctionExecutionStartLog, batch.functionName);
            var startTime = Date.now();
            try {
                results = batch.isBatching ? batch.call.call(null, batch.parameterValueSets) : [ batch.call.apply(null, batch.parameterValueSets[0]) ];
            } catch (ex) {
                var outboundData_1 = [];
                outboundData_1.push(Date.now() - startTime);
                for (var i = 0; i < batch.invocationIds.length; i++) ex instanceof CustomFunctionError ? this._setErrorWithOutboundData(batch.invocationIds[i], ex.message, this._getCustomFunctionResultErrorCodeFromErrorCode(ex.code), outboundData_1) : this._setErrorWithOutboundData(batch.invocationIds[i], ex, 2, outboundData_1);
                return void CustomFunctionsLogger.logEvent(CustomFunctionProxy.CustomFunctionExecutionExceptionThrownLog, batch.functionName, CustomFunctionProxy.toLogMessage(ex));
            }
            var outboundData = [];
            if (outboundData.push(Date.now() - startTime), batch.isStreaming) ; else if (results.length === batch.parameterValueSets.length) {
                var _loop_2 = function(i) {
                    _isNullOrUndefined(results[i]) || "object" != typeof results[i] || "function" != typeof results[i].then ? (CustomFunctionsLogger.logEvent(CustomFunctionProxy.CustomFunctionExecutionFinishLog, batch.functionName), 
                    this_2._setResultWithOutboundData(batch.invocationIds[i], results[i], outboundData)) : results[i].then((function(value) {
                        CustomFunctionsLogger.logEvent(CustomFunctionProxy.CustomFunctionExecutionFinishLog, batch.functionName), 
                        _this._setResultWithOutboundData(batch.invocationIds[i], value, outboundData);
                    }), (function(reason) {
                        CustomFunctionsLogger.logEvent(CustomFunctionProxy.CustomFunctionExecutionRejectedPromoseLog, batch.functionName, CustomFunctionProxy.toLogMessage(reason)), 
                        reason instanceof CustomFunctionError ? _this._setErrorWithOutboundData(batch.invocationIds[i], reason, _this._getCustomFunctionResultErrorCodeFromErrorCode(reason.code), outboundData) : _this._setErrorWithOutboundData(batch.invocationIds[i], reason, 3, outboundData);
                    }));
                }, this_2 = this;
                for (i = 0; i < results.length; i++) _loop_2(i);
            } else {
                CustomFunctionsLogger.logEvent(CustomFunctionProxy.CustomFunctionExecutionBatchMismatchLog, batch.functionName);
                for (i = 0; i < batch.invocationIds.length; i++) this._setErrorWithOutboundData(batch.invocationIds[i], OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.customFunctionUnexpectedNumberOfEntriesInResultBatch), 4, outboundData);
            }
        }, CustomFunctionProxy.prototype._setResult = function(invocationId, result) {
            var outboundData = [];
            outboundData.push(0), this._setResultWithOutboundData(invocationId, result, outboundData);
        }, CustomFunctionProxy.prototype._setResultWithOutboundData = function(invocationId, result, outboundData) {
            var invocationResult = {
                id: invocationId,
                value: result,
                outboundData: outboundData
            };
            "number" == typeof result ? isNaN(result) ? (invocationResult.failed = !0, invocationResult.value = "NaN") : isFinite(result) || (invocationResult.failed = !0, 
            invocationResult.value = "Infinity", invocationResult.errorCode = 6) : result instanceof CustomFunctionError ? (invocationResult.failed = !1, 
            invocationResult.value = {
                valueType: result.valueType,
                code: result.code,
                message: result.message
            }) : result instanceof Error && (invocationResult.failed = !0, invocationResult.value = CustomFunctionProxy.toLogMessage(result), 
            invocationResult.errorCode = 0);
            var timeNow = Date.now();
            this._resultEntryBuffer[invocationId] = {
                timeCreated: timeNow,
                result: invocationResult
            }, this._ensureSetResultsTaskIsScheduled(timeNow);
        }, CustomFunctionProxy.prototype._setError = function(invocationId, error, errorCode) {
            var outboundData = [];
            outboundData.push(0), this._setErrorWithOutboundData(invocationId, error, errorCode, outboundData);
        }, CustomFunctionProxy.prototype._setErrorWithOutboundData = function(invocationId, error, errorCode, outboundData) {
            var message = "";
            _isNullOrUndefined(error) || (error instanceof CustomFunctionError && !_isNullOrUndefined(error.message) ? message = error.message : "string" == typeof error && (message = error));
            var result = {
                id: invocationId,
                failed: !0,
                value: message,
                errorCode: errorCode,
                outboundData: outboundData
            }, timeNow = Date.now();
            this._resultEntryBuffer[invocationId] = {
                timeCreated: timeNow,
                result: result
            }, this._ensureSetResultsTaskIsScheduled(timeNow);
        }, CustomFunctionProxy.prototype._getCustomFunctionResultErrorCodeFromErrorCode = function(errorCode) {
            var internalCode;
            switch (errorCode) {
              case ErrorCode.notAvailable:
                internalCode = 1;
                break;

              case ErrorCode.divisionByZero:
                internalCode = 5;
                break;

              case ErrorCode.invalidValue:
                internalCode = 7;
                break;

              case ErrorCode.invalidNumber:
                internalCode = 6;
                break;

              case ErrorCode.nullReference:
                internalCode = 8;
                break;

              default:
                internalCode = 7;
            }
            return internalCode;
        }, CustomFunctionProxy.prototype._ensureSetResultsTaskIsScheduled = function(timeNow) {
            if (this._setResultsTaskOverdueTime > 0 && timeNow > this._setResultsTaskOverdueTime) return OfficeExtension.Utility.log("SetResultsTask overdue"), 
            void this._executeSetResultsTask();
            this._isSetResultsTaskScheduled || (OfficeExtension.Utility.log("setTimeout(CustomFunctionProxy._executeSetResultsTask)"), 
            setTimeout(this._executeSetResultsTask.bind(this), this._setResultsDelayMillis), 
            this._isSetResultsTaskScheduled = !0, this._setResultsTaskOverdueTime = timeNow + this._setResultsDelayMillis + this._setResultsOverdueDelayMillis);
        }, CustomFunctionProxy.prototype._convertCustomFunctionInvocationResultToArray = function(result) {
            var ret = [];
            return ret.push(result.id), ret.push(!result.failed), OfficeExtension.CoreUtility.isUndefined(result.value) ? ret.push(null) : ret.push(result.value), 
            result.failed && (OfficeExtension.CoreUtility.isUndefined(result.errorCode) ? ret.push(0) : ret.push(result.errorCode)), 
            OfficeExtension.Utility.isSetSupported("CustomFunctions", "1.9") && ret.push(result.outboundData), 
            ret;
        }, CustomFunctionProxy.prototype._executeSetResultsTask = function() {
            var _this = this;
            OfficeExtension.Utility.log("CustomFunctionProxy._executeSetResultsTask");
            var timeNow = Date.now();
            if (this._inProgressContextSyncExpectedFinishTime > 0 && this._inProgressContextSyncExpectedFinishTime > timeNow) return OfficeExtension.Utility.log("context.sync() is in progress. setTimeout(CustomFunctionProxy._executeSetResultsTask)"), 
            setTimeout(this._executeSetResultsTask.bind(this), this._setResultsDelayMillis), 
            void (this._setResultsTaskOverdueTime = timeNow + this._setResultsDelayMillis + this._setResultsOverdueDelayMillis);
            this._isSetResultsTaskScheduled = !1, this._setResultsTaskOverdueTime = 0;
            var resultEntryBufferCopy = {}, isSizeRetry = !1;
            CustomFunctions.isFeatureEnabled(1) && Object.keys(this._resultEntryRetryBuffer).length > 0 ? (resultEntryBufferCopy = this._resultEntryRetryBuffer, 
            this._resultEntryRetryBuffer = {}, this._retryBufferBodySize = 2, isSizeRetry = !0) : (resultEntryBufferCopy = this._resultEntryBuffer, 
            this._resultEntryBuffer = {});
            var useArrayResult = OfficeExtension.Utility.isSetSupported("CustomFunctions", "1.7"), invocationResults = [];
            for (var key in resultEntryBufferCopy) useArrayResult ? invocationResults.push(this._convertCustomFunctionInvocationResultToArray(resultEntryBufferCopy[key].result)) : invocationResults.push(resultEntryBufferCopy[key].result);
            if (0 !== invocationResults.length) {
                var context = new CustomFunctionRequestContext;
                useArrayResult ? context.customFunctions.setInvocationArrayResults(invocationResults) : context.customFunctions.setInvocationResults(invocationResults);
                var contextSyncStartTime = Date.now();
                this._inProgressContextSyncExpectedFinishTime = contextSyncStartTime + this._maxContextSyncExecutionDurationMills, 
                context.sync().then((function(value) {
                    var timeNow = Date.now();
                    _this._clearInProgressContextSyncExpectedFinishTimeAfterMinInterval(timeNow - contextSyncStartTime), 
                    isSizeRetry && _this._ensureSetResultsTaskIsScheduled(timeNow);
                }), (function(reason) {
                    var timeNow = Date.now();
                    _this._clearInProgressContextSyncExpectedFinishTimeAfterMinInterval(timeNow - contextSyncStartTime), 
                    CustomFunctions.isFeatureEnabled(1) && reason.code == OfficeExtension.ErrorCodes.requestPayloadSizeLimitExceeded ? _this._restoreAndSliceResultEntries(timeNow, resultEntryBufferCopy) : _this._restoreResultEntries(timeNow, resultEntryBufferCopy), 
                    _this._ensureSetResultsTaskIsScheduled(timeNow);
                }));
            }
        }, CustomFunctionProxy.prototype._restoreResultEntries = function(timeNow, resultEntryBufferCopy) {
            for (var key in resultEntryBufferCopy) {
                var resultSetter = resultEntryBufferCopy[key];
                timeNow - resultSetter.timeCreated <= this._setResultsLifeMillis && (this._resultEntryBuffer[key] || (this._resultEntryBuffer[key] = resultSetter));
            }
        }, CustomFunctionProxy.prototype._restoreAndSliceResultEntries = function(timeNow, resultEntryBufferCopy) {
            for (var key in resultEntryBufferCopy) {
                var resultSetter = resultEntryBufferCopy[key], resultSize = void 0;
                if (timeNow - resultSetter.timeCreated <= this._setResultsLifeMillis && !this._resultEntryBuffer[key]) {
                    try {
                        resultSize = JSON.stringify(this._convertCustomFunctionInvocationResultToArray(resultSetter.result)).length + 1;
                    } catch (ex) {
                        this._resultEntryBuffer[key] = resultSetter, OfficeExtension.Utility.log("Result can't pass in JSON.stringify");
                        continue;
                    }
                    this._retryBufferBodySize + resultSize < this._maxResultSize ? (this._resultEntryRetryBuffer[key] = resultSetter, 
                    this._retryBufferBodySize += resultSize) : this._resultEntryBuffer[key] = resultSetter;
                }
            }
        }, CustomFunctionProxy.prototype._clearInProgressContextSyncExpectedFinishTimeAfterMinInterval = function(lastContextSyncDurationMills) {
            var _this = this, interval = Math.max(this._minContextSyncIntervalMills, 2 * lastContextSyncDurationMills);
            OfficeExtension.Utility.log("setTimeout(clearInProgressContestSyncExpectedFinishedTime," + interval + ")"), 
            setTimeout((function() {
                OfficeExtension.Utility.log("clearInProgressContestSyncExpectedFinishedTime"), _this._inProgressContextSyncExpectedFinishTime = 0;
            }), interval);
        }, CustomFunctionProxy.CustomFunctionExecutionStartLog = new CustomFunctionLog(CustomFunctionLoggingSeverity.Verbose, "CustomFunctions [Execution] [Begin] Function="), 
        CustomFunctionProxy.CustomFunctionExecutionFailureLog = new CustomFunctionLog(CustomFunctionLoggingSeverity.Error, "CustomFunctions [Execution] [End] [Failure] Function="), 
        CustomFunctionProxy.CustomFunctionExecutionRejectedPromoseLog = new CustomFunctionLog(CustomFunctionLoggingSeverity.Error, "CustomFunctions [Execution] [End] [Failure] [RejectedPromise] Function="), 
        CustomFunctionProxy.CustomFunctionExecutionExceptionThrownLog = new CustomFunctionLog(CustomFunctionLoggingSeverity.Error, "CustomFunctions [Execution] [End] [Failure] [ExceptionThrown] Function="), 
        CustomFunctionProxy.CustomFunctionExecutionBatchMismatchLog = new CustomFunctionLog(CustomFunctionLoggingSeverity.Error, "CustomFunctions [Execution] [End] [Failure] [BatchMismatch] Function="), 
        CustomFunctionProxy.CustomFunctionExecutionFinishLog = new CustomFunctionLog(CustomFunctionLoggingSeverity.Info, "CustomFunctions [Execution] [End] [Success] Function="), 
        CustomFunctionProxy.CustomFunctionExecutionNotFoundLog = new CustomFunctionLog(CustomFunctionLoggingSeverity.Error, "CustomFunctions [Execution] [NotFound] Function="), 
        CustomFunctionProxy.CustomFunctionCancellation = new CustomFunctionLog(CustomFunctionLoggingSeverity.Info, "CustomFunctions [Cancellation] Function="), 
        CustomFunctionProxy.CustomFunctionAlreadyCancelled = new CustomFunctionLog(CustomFunctionLoggingSeverity.Info, "CustomFunctions [AlreadyCancelled] Function="), 
        CustomFunctionProxy.CustomFunctionDuplicatedName = new CustomFunctionLog(CustomFunctionLoggingSeverity.Error, "CustomFunctions [DuplicatedName] Function="), 
        CustomFunctionProxy.CustomFunctionInvalidArg = new CustomFunctionLog(CustomFunctionLoggingSeverity.Error, "CustomFunctions [InvalidArg] Name="), 
        CustomFunctionProxy;
    }();
    exports.CustomFunctionProxy = CustomFunctionProxy, exports.customFunctionProxy = new CustomFunctionProxy, 
    exports.setCustomFunctionAssociation = exports.customFunctionProxy.setCustomFunctionAssociation.bind(exports.customFunctionProxy), 
    exports.setCustomFunctionInvoker = exports.customFunctionProxy.setCustomFunctionInvoker.bind(exports.customFunctionProxy), 
    Core.HostBridge.onInited((function(hostBridge) {
        exports.customFunctionProxy._initFromHostBridge(hostBridge);
    }));
    var CustomFunctions = function(_super) {
        function CustomFunctions() {
            return null !== _super && _super.apply(this, arguments) || this;
        }
        return __extends(CustomFunctions, _super), Object.defineProperty(CustomFunctions.prototype, "_className", {
            get: function() {
                return "CustomFunctions";
            },
            enumerable: !1,
            configurable: !0
        }), CustomFunctions.isFeatureEnabled = function(args) {
            return this._enabledFeature instanceof Array && this._enabledFeature.indexOf(args) >= 0;
        }, CustomFunctions.initialize = function() {
            var _this = this, context = new CustomFunctionRequestContext;
            return exports.customFunctionProxy.ensureInit(context).then((function() {
                var enabledFeature;
                return context.customFunctions._SetOsfControlContainerReadyForCustomFunctions(), 
                OfficeExtension.Utility.isSetSupported("CustomFunctions", "1.9") && (enabledFeature = context.customFunctions._GetEnabledFeatures()), 
                OfficeExtension.Utility.log("OsfControl activation lifecycle: Set OsfControlContainer ready for CustomFunctions"), 
                context._customData = "SetOsfControlContainerReadyForCustomFunctions", context.sync().then((function() {
                    _isNullOrUndefined(enabledFeature) || (_this._enabledFeature = enabledFeature.value);
                })).catch((function(error) {
                    !function(error, rethrowOtherError) {
                        var isCellEditModeError = error instanceof OfficeExtension.Error && error.code === CustomFunctionRuntimeErrorCode.invalidOperationInCellEditMode;
                        if (OfficeExtension.CoreUtility.log("Error on starting custom functions: " + error), 
                        isCellEditModeError) {
                            OfficeExtension.CoreUtility.log("Was in cell-edit mode, will try again");
                            var delay_1 = exports.customFunctionProxy._ensureInitRetryDelayMillis;
                            return new OfficeExtension.CoreUtility.Promise((function(resolve) {
                                return setTimeout(resolve, delay_1);
                            })).then((function() {
                                return CustomFunctions.initialize();
                            }));
                        }
                        if (rethrowOtherError) throw error;
                    }(error, !0);
                }));
            }));
        }, CustomFunctions.prototype.setInvocationArrayResults = function(results) {
            _throwIfApiNotSupported("CustomFunctions.setInvocationArrayResults", "CustomFunctions", "1.4", "Excel"), 
            _invokeMethod(this, "SetInvocationArrayResults", 0, [ results ], 2, 0);
        }, CustomFunctions.prototype.setInvocationResults = function(results) {
            _invokeMethod(this, "SetInvocationResults", 0, [ results ], 2, 0);
        }, CustomFunctions.prototype._GetEnabledFeatures = function() {
            return _throwIfApiNotSupported("CustomFunctions._GetEnabledFeatures", "CustomFunctions", "1.4", "Excel"), 
            _invokeMethod(this, "_GetEnabledFeatures", 1, [], 4, 0);
        }, CustomFunctions.prototype._SetInvocationError = function(invocationId, message) {
            _invokeMethod(this, "_SetInvocationError", 0, [ invocationId, message ], 2, 0);
        }, CustomFunctions.prototype._SetInvocationResult = function(invocationId, result) {
            _invokeMethod(this, "_SetInvocationResult", 0, [ invocationId, result ], 2, 0);
        }, CustomFunctions.prototype._SetOsfControlContainerReadyForCustomFunctions = function() {
            _invokeMethod(this, "_SetOsfControlContainerReadyForCustomFunctions", 0, [], 10, 0);
        }, CustomFunctions.prototype._handleResult = function(value) {
            (_super.prototype._handleResult.call(this, value), _isNullOrUndefined(value)) || _fixObjectPathIfNecessary(this, value);
        }, CustomFunctions.prototype._handleRetrieveResult = function(value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result), _processRetrieveResult(this, value, result);
        }, CustomFunctions.newObject = function(context) {
            return _createTopLevelServiceObject(CustomFunctions, context, "Microsoft.ExcelServices.CustomFunctions", !1, 4);
        }, CustomFunctions.prototype.toJSON = function() {
            return _toJson(this, {}, {});
        }, CustomFunctions;
    }(OfficeExtension.ClientObject);
    exports.CustomFunctions = CustomFunctions;
    var CustomFunctionRuntimeErrorCode, CustomFunctionsContainer = function(_super) {
        function CustomFunctionsContainer() {
            return null !== _super && _super.apply(this, arguments) || this;
        }
        return __extends(CustomFunctionsContainer, _super), Object.defineProperty(CustomFunctionsContainer.prototype, "_className", {
            get: function() {
                return "CustomFunctionsContainer";
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(CustomFunctionsContainer.prototype, "_navigationPropertyNames", {
            get: function() {
                return [ "customFunctions" ];
            },
            enumerable: !1,
            configurable: !0
        }), Object.defineProperty(CustomFunctionsContainer.prototype, "customFunctions", {
            get: function() {
                return _throwIfApiNotSupported("CustomFunctionsContainer.customFunctions", "CustomFunctions", "1.2", "Excel"), 
                this._C || (this._C = _createPropertyObject(CustomFunctions, this, "CustomFunctions", !1, 4)), 
                this._C;
            },
            enumerable: !1,
            configurable: !0
        }), CustomFunctionsContainer.prototype._handleResult = function(value) {
            if (_super.prototype._handleResult.call(this, value), !_isNullOrUndefined(value)) {
                var obj = value;
                _fixObjectPathIfNecessary(this, obj), _handleNavigationPropertyResults(this, obj, [ "customFunctions", "CustomFunctions" ]);
            }
        }, CustomFunctionsContainer.prototype.load = function(options) {
            return _load(this, options);
        }, CustomFunctionsContainer.prototype._handleRetrieveResult = function(value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result), _processRetrieveResult(this, value, result);
        }, CustomFunctionsContainer.prototype.toJSON = function() {
            return _toJson(this, {}, {});
        }, CustomFunctionsContainer;
    }(OfficeExtension.ClientObject);
    exports.CustomFunctionsContainer = CustomFunctionsContainer, function(CustomFunctionRuntimeErrorCode) {
        CustomFunctionRuntimeErrorCode.generalException = "GeneralException", CustomFunctionRuntimeErrorCode.invalidOperation = "InvalidOperation", 
        CustomFunctionRuntimeErrorCode.invalidOperationInCellEditMode = "InvalidOperationInCellEditMode";
    }(CustomFunctionRuntimeErrorCode || (CustomFunctionRuntimeErrorCode = {}));
} ]);



var oteljs = function(modules) {
    var installedModules = {};
    function __webpack_require__(moduleId) {
        if (installedModules[moduleId]) return installedModules[moduleId].exports;
        var module = installedModules[moduleId] = {
            i: moduleId,
            l: !1,
            exports: {}
        };
        return modules[moduleId].call(module.exports, module, module.exports, __webpack_require__), 
        module.l = !0, module.exports;
    }
    return __webpack_require__.m = modules, __webpack_require__.c = installedModules, 
    __webpack_require__.d = function(exports, name, getter) {
        __webpack_require__.o(exports, name) || Object.defineProperty(exports, name, {
            enumerable: !0,
            get: getter
        });
    }, __webpack_require__.r = function(exports) {
        "undefined" != typeof Symbol && Symbol.toStringTag && Object.defineProperty(exports, Symbol.toStringTag, {
            value: "Module"
        }), Object.defineProperty(exports, "__esModule", {
            value: !0
        });
    }, __webpack_require__.t = function(value, mode) {
        if (1 & mode && (value = __webpack_require__(value)), 8 & mode) return value;
        if (4 & mode && "object" == typeof value && value && value.__esModule) return value;
        var ns = Object.create(null);
        if (__webpack_require__.r(ns), Object.defineProperty(ns, "default", {
            enumerable: !0,
            value: value
        }), 2 & mode && "string" != typeof value) for (var key in value) __webpack_require__.d(ns, key, function(key) {
            return value[key];
        }.bind(null, key));
        return ns;
    }, __webpack_require__.n = function(module) {
        var getter = module && module.__esModule ? function() {
            return module.default;
        } : function() {
            return module;
        };
        return __webpack_require__.d(getter, "a", getter), getter;
    }, __webpack_require__.o = function(object, property) {
        return Object.prototype.hasOwnProperty.call(object, property);
    }, __webpack_require__.p = "", __webpack_require__(__webpack_require__.s = 19);
}([ function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    __webpack_require__.d(__webpack_exports__, "a", (function() {
        return makeBooleanDataField;
    })), __webpack_require__.d(__webpack_exports__, "d", (function() {
        return makeInt64DataField;
    })), __webpack_require__.d(__webpack_exports__, "b", (function() {
        return makeDoubleDataField;
    })), __webpack_require__.d(__webpack_exports__, "e", (function() {
        return makeStringDataField;
    })), __webpack_require__.d(__webpack_exports__, "c", (function() {
        return makeGuidDataField;
    }));
    var _DataFieldType__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(3), _DataClassification__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(4);
    function makeBooleanDataField(name, value) {
        return {
            name: name,
            dataType: _DataFieldType__WEBPACK_IMPORTED_MODULE_0__.a.Boolean,
            value: value,
            classification: _DataClassification__WEBPACK_IMPORTED_MODULE_1__.a.SystemMetadata
        };
    }
    function makeInt64DataField(name, value) {
        return {
            name: name,
            dataType: _DataFieldType__WEBPACK_IMPORTED_MODULE_0__.a.Int64,
            value: value,
            classification: _DataClassification__WEBPACK_IMPORTED_MODULE_1__.a.SystemMetadata
        };
    }
    function makeDoubleDataField(name, value) {
        return {
            name: name,
            dataType: _DataFieldType__WEBPACK_IMPORTED_MODULE_0__.a.Double,
            value: value,
            classification: _DataClassification__WEBPACK_IMPORTED_MODULE_1__.a.SystemMetadata
        };
    }
    function makeStringDataField(name, value) {
        return {
            name: name,
            dataType: _DataFieldType__WEBPACK_IMPORTED_MODULE_0__.a.String,
            value: value,
            classification: _DataClassification__WEBPACK_IMPORTED_MODULE_1__.a.SystemMetadata
        };
    }
    function makeGuidDataField(name, value) {
        return {
            name: name,
            dataType: _DataFieldType__WEBPACK_IMPORTED_MODULE_0__.a.Guid,
            value: value,
            classification: _DataClassification__WEBPACK_IMPORTED_MODULE_1__.a.SystemMetadata
        };
    }
}, function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    __webpack_require__.d(__webpack_exports__, "b", (function() {
        return LogLevel;
    })), __webpack_require__.d(__webpack_exports__, "a", (function() {
        return Category;
    })), __webpack_require__.d(__webpack_exports__, "e", (function() {
        return onNotification;
    })), __webpack_require__.d(__webpack_exports__, "d", (function() {
        return logNotification;
    })), __webpack_require__.d(__webpack_exports__, "c", (function() {
        return logError;
    }));
    var LogLevel, Category, onNotificationEvent = new (__webpack_require__(10).a);
    function onNotification() {
        return onNotificationEvent;
    }
    function logNotification(level, category, message) {
        onNotificationEvent.fireEvent({
            level: level,
            category: category,
            message: message
        });
    }
    function logError(category, message, error) {
        logNotification(LogLevel.Error, category, (function() {
            var errorMessage = error instanceof Error ? error.message : "";
            return message + ": " + errorMessage;
        }));
    }
    !function(LogLevel) {
        LogLevel[LogLevel.Error = 0] = "Error", LogLevel[LogLevel.Warning = 1] = "Warning", 
        LogLevel[LogLevel.Info = 2] = "Info", LogLevel[LogLevel.Verbose = 3] = "Verbose";
    }(LogLevel || (LogLevel = {})), function(Category) {
        Category[Category.Core = 0] = "Core", Category[Category.Sink = 1] = "Sink", Category[Category.Transport = 2] = "Transport";
    }(Category || (Category = {}));
}, function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    __webpack_require__.d(__webpack_exports__, "a", (function() {
        return addContractField;
    }));
    var _DataFieldHelper__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(0);
    function addContractField(dataFields, instanceName, contractName) {
        dataFields.push(Object(_DataFieldHelper__WEBPACK_IMPORTED_MODULE_0__.e)("zC." + instanceName, contractName));
    }
}, function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    var DataFieldType;
    __webpack_require__.d(__webpack_exports__, "a", (function() {
        return DataFieldType;
    })), function(DataFieldType) {
        DataFieldType[DataFieldType.String = 0] = "String", DataFieldType[DataFieldType.Boolean = 1] = "Boolean", 
        DataFieldType[DataFieldType.Int64 = 2] = "Int64", DataFieldType[DataFieldType.Double = 3] = "Double", 
        DataFieldType[DataFieldType.Guid = 4] = "Guid";
    }(DataFieldType || (DataFieldType = {}));
}, function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    var DataClassification;
    __webpack_require__.d(__webpack_exports__, "a", (function() {
        return DataClassification;
    })), function(DataClassification) {
        DataClassification[DataClassification.EssentialServiceMetadata = 1] = "EssentialServiceMetadata", 
        DataClassification[DataClassification.AccountData = 2] = "AccountData", DataClassification[DataClassification.SystemMetadata = 4] = "SystemMetadata", 
        DataClassification[DataClassification.OrganizationIdentifiableInformation = 8] = "OrganizationIdentifiableInformation", 
        DataClassification[DataClassification.EndUserIdentifiableInformation = 16] = "EndUserIdentifiableInformation", 
        DataClassification[DataClassification.CustomerContent = 32] = "CustomerContent", 
        DataClassification[DataClassification.AccessControl = 64] = "AccessControl";
    }(DataClassification || (DataClassification = {}));
}, function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    var SamplingPolicy, PersistencePriority, CostPriority, DataCategories, DiagnosticLevel;
    __webpack_require__.d(__webpack_exports__, "e", (function() {
        return SamplingPolicy;
    })), __webpack_require__.d(__webpack_exports__, "d", (function() {
        return PersistencePriority;
    })), __webpack_require__.d(__webpack_exports__, "a", (function() {
        return CostPriority;
    })), __webpack_require__.d(__webpack_exports__, "b", (function() {
        return DataCategories;
    })), __webpack_require__.d(__webpack_exports__, "c", (function() {
        return DiagnosticLevel;
    })), function(SamplingPolicy) {
        SamplingPolicy[SamplingPolicy.NotSet = 0] = "NotSet", SamplingPolicy[SamplingPolicy.Measure = 1] = "Measure", 
        SamplingPolicy[SamplingPolicy.Diagnostics = 2] = "Diagnostics", SamplingPolicy[SamplingPolicy.CriticalBusinessImpact = 191] = "CriticalBusinessImpact", 
        SamplingPolicy[SamplingPolicy.CriticalCensus = 192] = "CriticalCensus", SamplingPolicy[SamplingPolicy.CriticalExperimentation = 193] = "CriticalExperimentation", 
        SamplingPolicy[SamplingPolicy.CriticalUsage = 194] = "CriticalUsage";
    }(SamplingPolicy || (SamplingPolicy = {})), function(PersistencePriority) {
        PersistencePriority[PersistencePriority.NotSet = 0] = "NotSet", PersistencePriority[PersistencePriority.Normal = 1] = "Normal", 
        PersistencePriority[PersistencePriority.High = 2] = "High";
    }(PersistencePriority || (PersistencePriority = {})), function(CostPriority) {
        CostPriority[CostPriority.NotSet = 0] = "NotSet", CostPriority[CostPriority.Normal = 1] = "Normal", 
        CostPriority[CostPriority.High = 2] = "High";
    }(CostPriority || (CostPriority = {})), function(DataCategories) {
        DataCategories[DataCategories.NotSet = 0] = "NotSet", DataCategories[DataCategories.SoftwareSetup = 1] = "SoftwareSetup", 
        DataCategories[DataCategories.ProductServiceUsage = 2] = "ProductServiceUsage", 
        DataCategories[DataCategories.ProductServicePerformance = 4] = "ProductServicePerformance", 
        DataCategories[DataCategories.DeviceConfiguration = 8] = "DeviceConfiguration", 
        DataCategories[DataCategories.InkingTypingSpeech = 16] = "InkingTypingSpeech";
    }(DataCategories || (DataCategories = {})), function(DiagnosticLevel) {
        DiagnosticLevel[DiagnosticLevel.ReservedDoNotUse = 0] = "ReservedDoNotUse", DiagnosticLevel[DiagnosticLevel.BasicEvent = 10] = "BasicEvent", 
        DiagnosticLevel[DiagnosticLevel.FullEvent = 100] = "FullEvent", DiagnosticLevel[DiagnosticLevel.NecessaryServiceDataEvent = 110] = "NecessaryServiceDataEvent", 
        DiagnosticLevel[DiagnosticLevel.AlwaysOnNecessaryServiceDataEvent = 120] = "AlwaysOnNecessaryServiceDataEvent";
    }(DiagnosticLevel || (DiagnosticLevel = {}));
}, function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    __webpack_require__.d(__webpack_exports__, "a", (function() {
        return Contracts;
    }));
    var officeeventschema_tml_Result, officeeventschema_tml_Activity, Activity, officeeventschema_tml_Host, officeeventschema_tml_User, officeeventschema_tml_SDX, officeeventschema_tml_Funnel, officeeventschema_tml_UserAction, Office_System_Error_Error, DataFieldHelper = __webpack_require__(0), Contract = __webpack_require__(2);
    (officeeventschema_tml_Result || (officeeventschema_tml_Result = {})).getFields = function(instanceName, contract) {
        var dataFields = [];
        return dataFields.push(Object(DataFieldHelper.d)(instanceName + ".Code", contract.code)), 
        void 0 !== contract.type && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".Type", contract.type)), 
        void 0 !== contract.tag && dataFields.push(Object(DataFieldHelper.d)(instanceName + ".Tag", contract.tag)), 
        void 0 !== contract.isExpected && dataFields.push(Object(DataFieldHelper.a)(instanceName + ".IsExpected", contract.isExpected)), 
        Object(Contract.a)(dataFields, instanceName, "Office.System.Result"), dataFields;
    }, (Activity = officeeventschema_tml_Activity || (officeeventschema_tml_Activity = {})).contractName = "Office.System.Activity", 
    Activity.getFields = function(contract) {
        var dataFields = [];
        return void 0 !== contract.cV && dataFields.push(Object(DataFieldHelper.e)("Activity.CV", contract.cV)), 
        dataFields.push(Object(DataFieldHelper.d)("Activity.Duration", contract.duration)), 
        dataFields.push(Object(DataFieldHelper.d)("Activity.Count", contract.count)), dataFields.push(Object(DataFieldHelper.d)("Activity.AggMode", contract.aggMode)), 
        void 0 !== contract.success && dataFields.push(Object(DataFieldHelper.a)("Activity.Success", contract.success)), 
        void 0 !== contract.result && dataFields.push.apply(dataFields, officeeventschema_tml_Result.getFields("Activity.Result", contract.result)), 
        Object(Contract.a)(dataFields, "Activity", Activity.contractName), dataFields;
    }, (officeeventschema_tml_Host || (officeeventschema_tml_Host = {})).getFields = function(instanceName, contract) {
        var dataFields = [];
        return void 0 !== contract.id && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".Id", contract.id)), 
        void 0 !== contract.version && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".Version", contract.version)), 
        void 0 !== contract.sessionId && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".SessionId", contract.sessionId)), 
        Object(Contract.a)(dataFields, instanceName, "Office.System.Host"), dataFields;
    }, (officeeventschema_tml_User || (officeeventschema_tml_User = {})).getFields = function(instanceName, contract) {
        var dataFields = [];
        return void 0 !== contract.alias && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".Alias", contract.alias)), 
        void 0 !== contract.primaryIdentityHash && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".PrimaryIdentityHash", contract.primaryIdentityHash)), 
        void 0 !== contract.primaryIdentitySpace && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".PrimaryIdentitySpace", contract.primaryIdentitySpace)), 
        void 0 !== contract.tenantId && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".TenantId", contract.tenantId)), 
        void 0 !== contract.tenantGroup && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".TenantGroup", contract.tenantGroup)), 
        void 0 !== contract.isAnonymous && dataFields.push(Object(DataFieldHelper.a)(instanceName + ".IsAnonymous", contract.isAnonymous)), 
        Object(Contract.a)(dataFields, instanceName, "Office.System.User"), dataFields;
    }, (officeeventschema_tml_SDX || (officeeventschema_tml_SDX = {})).getFields = function(instanceName, contract) {
        var dataFields = [];
        return void 0 !== contract.id && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".Id", contract.id)), 
        void 0 !== contract.version && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".Version", contract.version)), 
        void 0 !== contract.instanceId && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".InstanceId", contract.instanceId)), 
        void 0 !== contract.name && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".Name", contract.name)), 
        void 0 !== contract.marketplaceType && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".MarketplaceType", contract.marketplaceType)), 
        void 0 !== contract.sessionId && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".SessionId", contract.sessionId)), 
        void 0 !== contract.browserToken && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".BrowserToken", contract.browserToken)), 
        void 0 !== contract.osfRuntimeVersion && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".OsfRuntimeVersion", contract.osfRuntimeVersion)), 
        void 0 !== contract.officeJsVersion && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".OfficeJsVersion", contract.officeJsVersion)), 
        void 0 !== contract.hostJsVersion && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".HostJsVersion", contract.hostJsVersion)), 
        void 0 !== contract.assetId && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".AssetId", contract.assetId)), 
        void 0 !== contract.providerName && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".ProviderName", contract.providerName)), 
        void 0 !== contract.type && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".Type", contract.type)), 
        Object(Contract.a)(dataFields, instanceName, "Office.System.SDX"), dataFields;
    }, (officeeventschema_tml_Funnel || (officeeventschema_tml_Funnel = {})).getFields = function(instanceName, contract) {
        var dataFields = [];
        return void 0 !== contract.name && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".Name", contract.name)), 
        void 0 !== contract.state && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".State", contract.state)), 
        Object(Contract.a)(dataFields, instanceName, "Office.System.Funnel"), dataFields;
    }, (officeeventschema_tml_UserAction || (officeeventschema_tml_UserAction = {})).getFields = function(instanceName, contract) {
        var dataFields = [];
        return void 0 !== contract.id && dataFields.push(Object(DataFieldHelper.d)(instanceName + ".Id", contract.id)), 
        void 0 !== contract.name && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".Name", contract.name)), 
        void 0 !== contract.commandSurface && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".CommandSurface", contract.commandSurface)), 
        void 0 !== contract.parentName && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".ParentName", contract.parentName)), 
        void 0 !== contract.triggerMethod && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".TriggerMethod", contract.triggerMethod)), 
        void 0 !== contract.timeOffsetMs && dataFields.push(Object(DataFieldHelper.d)(instanceName + ".TimeOffsetMs", contract.timeOffsetMs)), 
        Object(Contract.a)(dataFields, instanceName, "Office.System.UserAction"), dataFields;
    }, function(Error) {
        Error.getFields = function(instanceName, contract) {
            var dataFields = [];
            return dataFields.push(Object(DataFieldHelper.e)(instanceName + ".ErrorGroup", contract.errorGroup)), 
            dataFields.push(Object(DataFieldHelper.d)(instanceName + ".Tag", contract.tag)), 
            void 0 !== contract.code && dataFields.push(Object(DataFieldHelper.d)(instanceName + ".Code", contract.code)), 
            void 0 !== contract.id && dataFields.push(Object(DataFieldHelper.d)(instanceName + ".Id", contract.id)), 
            void 0 !== contract.count && dataFields.push(Object(DataFieldHelper.d)(instanceName + ".Count", contract.count)), 
            Object(Contract.a)(dataFields, instanceName, "Office.System.Error"), dataFields;
        };
    }(Office_System_Error_Error || (Office_System_Error_Error = {}));
    var Contracts, _Activity = officeeventschema_tml_Activity, _Result = officeeventschema_tml_Result, _Error = Office_System_Error_Error, _Funnel = officeeventschema_tml_Funnel, _Host = officeeventschema_tml_Host, _SDX = officeeventschema_tml_SDX, _UserAction = officeeventschema_tml_UserAction, _User = officeeventschema_tml_User;
    !function(Contracts) {
        !function(Office) {
            !function(System) {
                System.Activity = _Activity, System.Result = _Result, System.Error = _Error, System.Funnel = _Funnel, 
                System.Host = _Host, System.SDX = _SDX, System.User = _User, System.UserAction = _UserAction;
            }(Office.System || (Office.System = {}));
        }(Contracts.Office || (Contracts.Office = {}));
    }(Contracts || (Contracts = {}));
}, function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    function cloneEvent(event) {
        var localEvent = {
            eventName: event.eventName,
            eventFlags: event.eventFlags
        };
        return event.telemetryProperties && (localEvent.telemetryProperties = {
            ariaTenantToken: event.telemetryProperties.ariaTenantToken,
            nexusTenantToken: event.telemetryProperties.nexusTenantToken
        }), event.eventContract && (localEvent.eventContract = {
            name: event.eventContract.name,
            dataFields: event.eventContract.dataFields.slice()
        }), localEvent.dataFields = event.dataFields ? event.dataFields.slice() : [], localEvent;
    }
    __webpack_require__.d(__webpack_exports__, "a", (function() {
        return cloneEvent;
    }));
}, function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    __webpack_require__.d(__webpack_exports__, "b", (function() {
        return SuppressNexus;
    })), __webpack_require__.d(__webpack_exports__, "a", (function() {
        return SimpleTelemetryLogger_SimpleTelemetryLogger;
    }));
    var TokenType, TenantTokenManager_TenantTokenManager, TelemetryEvent = __webpack_require__(7), OTelNotifications = __webpack_require__(1);
    !function(TokenType) {
        TokenType[TokenType.Aria = 0] = "Aria", TokenType[TokenType.Nexus = 1] = "Nexus";
    }(TokenType || (TokenType = {})), function(TenantTokenManager) {
        var ariaTokenMap = {}, nexusTokenMap = {}, tenantTokens = {};
        function setTenantTokens(tokenTree) {
            if ("object" != typeof tokenTree) throw new Error("tokenTree must be an object");
            tenantTokens = function mergeTenantTokens(existingTokenTree, newTokenTree) {
                if ("object" != typeof newTokenTree) return newTokenTree;
                for (var _i = 0, _a = Object.keys(newTokenTree); _i < _a.length; _i++) {
                    var key = _a[_i];
                    key in existingTokenTree && (existingTokenTree[key], 1) ? existingTokenTree[key] = mergeTenantTokens(existingTokenTree[key], newTokenTree[key]) : existingTokenTree[key] = newTokenTree[key];
                }
                return existingTokenTree;
            }(tenantTokens, tokenTree);
        }
        function getAriaTenantToken(eventName) {
            if (ariaTokenMap[eventName]) return ariaTokenMap[eventName];
            var ariaToken = getTenantToken(eventName, TokenType.Aria);
            return "string" == typeof ariaToken ? (ariaTokenMap[eventName] = ariaToken, ariaToken) : void 0;
        }
        function getNexusTenantToken(eventName) {
            if (nexusTokenMap[eventName]) return nexusTokenMap[eventName];
            var nexusToken = getTenantToken(eventName, TokenType.Nexus);
            return "number" == typeof nexusToken ? (nexusTokenMap[eventName] = nexusToken, nexusToken) : void 0;
        }
        function getTenantToken(eventName, tokenType) {
            var pieces = eventName.split("."), node = tenantTokens, token = void 0;
            if (node) {
                for (var i = 0; i < pieces.length - 1; i++) node[pieces[i]] && (node = node[pieces[i]], 
                tokenType === TokenType.Aria && "string" == typeof node.ariaTenantToken ? token = node.ariaTenantToken : tokenType === TokenType.Nexus && "number" == typeof node.nexusTenantToken && (token = node.nexusTenantToken));
                return token;
            }
        }
        TenantTokenManager.setTenantToken = function(namespace, ariaTenantToken, nexusTenantToken) {
            var parts = namespace.split(".");
            if (parts.length < 2 || "Office" !== parts[0]) Object(OTelNotifications.d)(OTelNotifications.b.Error, OTelNotifications.a.Core, (function() {
                return "Invalid namespace: " + namespace;
            })); else {
                var leaf = Object.create(Object.prototype);
                ariaTenantToken && (leaf.ariaTenantToken = ariaTenantToken), nexusTenantToken && (leaf.nexusTenantToken = nexusTenantToken);
                var index, node = leaf;
                for (index = parts.length - 1; index >= 0; --index) {
                    var parentNode = Object.create(Object.prototype);
                    parentNode[parts[index]] = node, node = parentNode;
                }
                setTenantTokens(node);
            }
        }, TenantTokenManager.setTenantTokens = setTenantTokens, TenantTokenManager.getTenantTokens = function(eventName) {
            var ariaTenantToken = getAriaTenantToken(eventName), nexusTenantToken = getNexusTenantToken(eventName);
            if (!nexusTenantToken || !ariaTenantToken) throw new Error("Could not find tenant token for " + eventName);
            return {
                ariaTenantToken: ariaTenantToken,
                nexusTenantToken: nexusTenantToken
            };
        }, TenantTokenManager.getAriaTenantToken = getAriaTenantToken, TenantTokenManager.getNexusTenantToken = getNexusTenantToken, 
        TenantTokenManager.clear = function() {
            ariaTokenMap = {}, nexusTokenMap = {}, tenantTokens = {};
        };
    }(TenantTokenManager_TenantTokenManager || (TenantTokenManager_TenantTokenManager = {}));
    var TelemetryEventValidator_TelemetryEventValidator, DataFieldType = __webpack_require__(3);
    !function(TelemetryEventValidator) {
        var StartsWithCapitalRegex = /^[A-Z][a-zA-Z0-9]*$/, AlphanumericRegex = /^[a-zA-Z0-9_\.]*$/;
        function isNameValid(name) {
            return void 0 !== name && AlphanumericRegex.test(name);
        }
        function validateDataField(dataField) {
            if (!((dataFieldName = dataField.name) && isNameValid(dataFieldName) && dataFieldName.length + 5 < 100)) throw new Error("Invalid dataField name");
            var dataFieldName;
            dataField.dataType === DataFieldType.a.Int64 && validateInt(dataField.value);
        }
        function validateInt(value) {
            if ("number" != typeof value || !isFinite(value) || Math.floor(value) !== value || value < -9007199254740991 || value > 9007199254740991) throw new Error("Invalid integer " + JSON.stringify(value));
        }
        TelemetryEventValidator.validateTelemetryEvent = function(event) {
            if (!function(eventName) {
                if (!eventName || eventName.length > 98) return !1;
                var eventNamePieces = eventName.split("."), eventNodeName = eventNamePieces[eventNamePieces.length - 1];
                return function(eventNamePieces) {
                    return !!eventNamePieces && eventNamePieces.length >= 3 && "Office" === eventNamePieces[0];
                }(eventNamePieces) && (eventNode = eventNodeName, void 0 !== eventNode && StartsWithCapitalRegex.test(eventNode));
                var eventNode;
            }(event.eventName)) throw new Error("Invalid eventName");
            if (event.eventContract && !isNameValid(event.eventContract.name)) throw new Error("Invalid eventContract");
            if (null != event.dataFields) for (var i = 0; i < event.dataFields.length; i++) validateDataField(event.dataFields[i]);
        }, TelemetryEventValidator.validateInt = validateInt;
    }(TelemetryEventValidator_TelemetryEventValidator || (TelemetryEventValidator_TelemetryEventValidator = {}));
    var Event = __webpack_require__(10), DataFieldHelper = __webpack_require__(0), __assign = function() {
        return (__assign = Object.assign || function(t) {
            for (var s, i = 1, n = arguments.length; i < n; i++) for (var p in s = arguments[i]) Object.prototype.hasOwnProperty.call(s, p) && (t[p] = s[p]);
            return t;
        }).apply(this, arguments);
    }, SuppressNexus = -1, SimpleTelemetryLogger_SimpleTelemetryLogger = function() {
        function SimpleTelemetryLogger(parent, persistentDataFields, config) {
            var _a, _b;
            this.onSendEvent = new Event.a, this.persistentDataFields = [], this.config = config || {}, 
            parent ? (this.onSendEvent = parent.onSendEvent, (_a = this.persistentDataFields).push.apply(_a, parent.persistentDataFields), 
            this.config = __assign(__assign({}, parent.getConfig()), this.config)) : this.persistentDataFields.push(Object(DataFieldHelper.e)("OTelJS.Version", "3.1.70")), 
            persistentDataFields && (_b = this.persistentDataFields).push.apply(_b, persistentDataFields);
        }
        return SimpleTelemetryLogger.prototype.sendTelemetryEvent = function(event) {
            var localEvent;
            try {
                if (0 === this.onSendEvent.getListenerCount()) return void Object(OTelNotifications.d)(OTelNotifications.b.Warning, OTelNotifications.a.Core, (function() {
                    return "No telemetry sinks are attached.";
                }));
                localEvent = this.cloneEvent(event), this.processTelemetryEvent(localEvent);
            } catch (error) {
                return void Object(OTelNotifications.c)(OTelNotifications.a.Core, "SendTelemetryEvent", error);
            }
            try {
                this.onSendEvent.fireEvent(localEvent);
            } catch (_e) {}
        }, SimpleTelemetryLogger.prototype.processTelemetryEvent = function(event) {
            var _a;
            event.telemetryProperties || (event.telemetryProperties = TenantTokenManager_TenantTokenManager.getTenantTokens(event.eventName)), 
            event.dataFields && this.persistentDataFields && (_a = event.dataFields).unshift.apply(_a, this.persistentDataFields), 
            this.config.disableValidation || TelemetryEventValidator_TelemetryEventValidator.validateTelemetryEvent(event);
        }, SimpleTelemetryLogger.prototype.addSink = function(sink) {
            this.onSendEvent.addListener((function(event) {
                return sink.sendTelemetryEvent(event);
            }));
        }, SimpleTelemetryLogger.prototype.setTenantToken = function(namespace, ariaTenantToken, nexusTenantToken) {
            TenantTokenManager_TenantTokenManager.setTenantToken(namespace, ariaTenantToken, nexusTenantToken);
        }, SimpleTelemetryLogger.prototype.setTenantTokens = function(tokenTree) {
            TenantTokenManager_TenantTokenManager.setTenantTokens(tokenTree);
        }, SimpleTelemetryLogger.prototype.cloneEvent = function(event) {
            return Object(TelemetryEvent.a)(event);
        }, SimpleTelemetryLogger.prototype.getConfig = function() {
            return this.config;
        }, SimpleTelemetryLogger;
    }();
}, function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    var CorrelationVector;
    __webpack_require__.d(__webpack_exports__, "a", (function() {
        return Activity_ActivityScope;
    })), function(CorrelationVector) {
        var baseHash, baseId = 0;
        CorrelationVector.getNext = function() {
            return void 0 === baseHash && (baseHash = function() {
                for (var characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", result = [], i = 0; i < 22; i++) result.push(characters.charAt(Math.floor(Math.random() * characters.length)));
                return result.join("");
            }()), new CV(baseHash, ++baseId);
        }, CorrelationVector.getNextChild = function(parent) {
            return new CV(parent.getString(), ++parent.nextChild);
        };
        var CV = function() {
            function CV(base, id) {
                this.base = base, this.id = id, this.nextChild = 0;
            }
            return CV.prototype.getString = function() {
                return this.base + "." + this.id;
            }, CV;
        }();
        CorrelationVector.CV = CV;
    }(CorrelationVector || (CorrelationVector = {}));
    var OTelNotifications = __webpack_require__(1), __awaiter = function(thisArg, _arguments, P, generator) {
        return new (P || (P = Promise))((function(resolve, reject) {
            function fulfilled(value) {
                try {
                    step(generator.next(value));
                } catch (e) {
                    reject(e);
                }
            }
            function rejected(value) {
                try {
                    step(generator.throw(value));
                } catch (e) {
                    reject(e);
                }
            }
            function step(result) {
                var value;
                result.done ? resolve(result.value) : (value = result.value, value instanceof P ? value : new P((function(resolve) {
                    resolve(value);
                }))).then(fulfilled, rejected);
            }
            step((generator = generator.apply(thisArg, _arguments || [])).next());
        }));
    }, __generator = function(thisArg, body) {
        var f, y, t, g, _ = {
            label: 0,
            sent: function() {
                if (1 & t[0]) throw t[1];
                return t[1];
            },
            trys: [],
            ops: []
        };
        return g = {
            next: verb(0),
            throw: verb(1),
            return: verb(2)
        }, "function" == typeof Symbol && (g[Symbol.iterator] = function() {
            return this;
        }), g;
        function verb(n) {
            return function(v) {
                return function(op) {
                    if (f) throw new TypeError("Generator is already executing.");
                    for (;_; ) try {
                        if (f = 1, y && (t = 2 & op[0] ? y.return : op[0] ? y.throw || ((t = y.return) && t.call(y), 
                        0) : y.next) && !(t = t.call(y, op[1])).done) return t;
                        switch (y = 0, t && (op = [ 2 & op[0], t.value ]), op[0]) {
                          case 0:
                          case 1:
                            t = op;
                            break;

                          case 4:
                            return _.label++, {
                                value: op[1],
                                done: !1
                            };

                          case 5:
                            _.label++, y = op[1], op = [ 0 ];
                            continue;

                          case 7:
                            op = _.ops.pop(), _.trys.pop();
                            continue;

                          default:
                            if (!(t = _.trys, (t = t.length > 0 && t[t.length - 1]) || 6 !== op[0] && 2 !== op[0])) {
                                _ = 0;
                                continue;
                            }
                            if (3 === op[0] && (!t || op[1] > t[0] && op[1] < t[3])) {
                                _.label = op[1];
                                break;
                            }
                            if (6 === op[0] && _.label < t[1]) {
                                _.label = t[1], t = op;
                                break;
                            }
                            if (t && _.label < t[2]) {
                                _.label = t[2], _.ops.push(op);
                                break;
                            }
                            t[2] && _.ops.pop(), _.trys.pop();
                            continue;
                        }
                        op = body.call(thisArg, _);
                    } catch (e) {
                        op = [ 6, e ], y = 0;
                    } finally {
                        f = t = 0;
                    }
                    if (5 & op[0]) throw op[1];
                    return {
                        value: op[0] ? op[1] : void 0,
                        done: !0
                    };
                }([ n, v ]);
            };
        }
    }, getCurrentMicroseconds = function() {
        return 1e3 * Date.now();
    };
    "object" == typeof window && "object" == typeof window.performance && "now" in window.performance && (getCurrentMicroseconds = function() {
        return 1e3 * Math.floor(window.performance.now());
    });
    var Activity_ActivityScope = function() {
        function ActivityScope(telemetryLogger, activityName, parent) {
            this._optionalEventFlags = {}, this._ended = !1, this._telemetryLogger = telemetryLogger, 
            this._activityName = activityName, this._cv = parent ? CorrelationVector.getNextChild(parent._cv) : CorrelationVector.getNext(), 
            this._dataFields = [], this._success = void 0, this._startTime = getCurrentMicroseconds();
        }
        return ActivityScope.createNew = function(telemetryLogger, activityName) {
            return new ActivityScope(telemetryLogger, activityName);
        }, ActivityScope.prototype.createChildActivity = function(activityName) {
            return new ActivityScope(this._telemetryLogger, activityName, this);
        }, ActivityScope.prototype.setEventFlags = function(eventFlags) {
            this._optionalEventFlags = eventFlags;
        }, ActivityScope.prototype.addDataField = function(dataField) {
            this._dataFields.push(dataField);
        }, ActivityScope.prototype.addDataFields = function(dataFields) {
            var _a;
            (_a = this._dataFields).push.apply(_a, dataFields);
        }, ActivityScope.prototype.setSuccess = function(success) {
            this._success = success;
        }, ActivityScope.prototype.setResult = function(code, type, tag) {
            this._result = {
                code: code,
                type: type,
                tag: tag
            };
        }, ActivityScope.prototype.endNow = function() {
            if (!this._ended) {
                void 0 === this._success && void 0 === this._result && Object(OTelNotifications.d)(OTelNotifications.b.Warning, OTelNotifications.a.Core, (function() {
                    return "Activity does not have success or result set";
                }));
                var duration = getCurrentMicroseconds() - this._startTime;
                this._ended = !0;
                var activity = {
                    duration: duration,
                    count: 1,
                    aggMode: 0,
                    cV: this._cv.getString(),
                    success: this._success,
                    result: this._result
                };
                return this._telemetryLogger.sendActivity(this._activityName, activity, this._dataFields, this._optionalEventFlags);
            }
            Object(OTelNotifications.d)(OTelNotifications.b.Error, OTelNotifications.a.Core, (function() {
                return "Activity has already ended";
            }));
        }, ActivityScope.prototype.executeAsync = function(activityBody) {
            return __awaiter(this, void 0, void 0, (function() {
                var _this = this;
                return __generator(this, (function(_a) {
                    return [ 2, activityBody(this).then((function(result) {
                        return _this.endNow(), result;
                    })).catch((function(e) {
                        throw _this.endNow(), e;
                    })) ];
                }));
            }));
        }, ActivityScope.prototype.executeSync = function(activityBody) {
            try {
                var ret = activityBody(this);
                return this.endNow(), ret;
            } catch (e) {
                throw this.endNow(), e;
            }
        }, ActivityScope.prototype.executeChildActivityAsync = function(activityName, activityBody) {
            return __awaiter(this, void 0, void 0, (function() {
                return __generator(this, (function(_a) {
                    return [ 2, this.createChildActivity(activityName).executeAsync(activityBody) ];
                }));
            }));
        }, ActivityScope.prototype.executeChildActivitySync = function(activityName, activityBody) {
            return this.createChildActivity(activityName).executeSync(activityBody);
        }, ActivityScope;
    }();
}, function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    __webpack_require__.d(__webpack_exports__, "a", (function() {
        return Event;
    }));
    var Event = function() {
        function Event() {
            this._listeners = [];
        }
        return Event.prototype.fireEvent = function(args) {
            this._listeners.forEach((function(listener) {
                return listener(args);
            }));
        }, Event.prototype.addListener = function(listener) {
            listener && this._listeners.push(listener);
        }, Event.prototype.removeListener = function(listener) {
            this._listeners = this._listeners.filter((function(h) {
                return h !== listener;
            }));
        }, Event.prototype.getListenerCount = function() {
            return this._listeners.length;
        }, Event;
    }();
}, function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    __webpack_require__.r(__webpack_exports__);
    var _contracts_Contracts__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(6);
    __webpack_require__.d(__webpack_exports__, "Contracts", (function() {
        return _contracts_Contracts__WEBPACK_IMPORTED_MODULE_0__.a;
    }));
    var _Activity__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(9);
    __webpack_require__.d(__webpack_exports__, "ActivityScope", (function() {
        return _Activity__WEBPACK_IMPORTED_MODULE_1__.a;
    }));
    var _Contract__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(2);
    __webpack_require__.d(__webpack_exports__, "addContractField", (function() {
        return _Contract__WEBPACK_IMPORTED_MODULE_2__.a;
    }));
    var _CustomContract__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(12);
    __webpack_require__.d(__webpack_exports__, "getFieldsForContract", (function() {
        return _CustomContract__WEBPACK_IMPORTED_MODULE_3__.a;
    }));
    var _DataClassification__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(4);
    __webpack_require__.d(__webpack_exports__, "DataClassification", (function() {
        return _DataClassification__WEBPACK_IMPORTED_MODULE_4__.a;
    }));
    var _DataField__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(13);
    for (var __WEBPACK_IMPORT_KEY__ in _DataField__WEBPACK_IMPORTED_MODULE_5__) [ "default", "Contracts", "ActivityScope", "addContractField", "getFieldsForContract", "DataClassification" ].indexOf(__WEBPACK_IMPORT_KEY__) < 0 && function(key) {
        __webpack_require__.d(__webpack_exports__, key, (function() {
            return _DataField__WEBPACK_IMPORTED_MODULE_5__[key];
        }));
    }(__WEBPACK_IMPORT_KEY__);
    var _DataFieldHelper__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(0);
    __webpack_require__.d(__webpack_exports__, "makeBooleanDataField", (function() {
        return _DataFieldHelper__WEBPACK_IMPORTED_MODULE_6__.a;
    })), __webpack_require__.d(__webpack_exports__, "makeInt64DataField", (function() {
        return _DataFieldHelper__WEBPACK_IMPORTED_MODULE_6__.d;
    })), __webpack_require__.d(__webpack_exports__, "makeDoubleDataField", (function() {
        return _DataFieldHelper__WEBPACK_IMPORTED_MODULE_6__.b;
    })), __webpack_require__.d(__webpack_exports__, "makeStringDataField", (function() {
        return _DataFieldHelper__WEBPACK_IMPORTED_MODULE_6__.e;
    })), __webpack_require__.d(__webpack_exports__, "makeGuidDataField", (function() {
        return _DataFieldHelper__WEBPACK_IMPORTED_MODULE_6__.c;
    }));
    var _DataFieldType__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(3);
    __webpack_require__.d(__webpack_exports__, "DataFieldType", (function() {
        return _DataFieldType__WEBPACK_IMPORTED_MODULE_7__.a;
    }));
    var _EventFlagFiller__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(14);
    __webpack_require__.d(__webpack_exports__, "getEffectiveEventFlags", (function() {
        return _EventFlagFiller__WEBPACK_IMPORTED_MODULE_8__.a;
    }));
    var _EventFlagsProperties__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(5);
    __webpack_require__.d(__webpack_exports__, "SamplingPolicy", (function() {
        return _EventFlagsProperties__WEBPACK_IMPORTED_MODULE_9__.e;
    })), __webpack_require__.d(__webpack_exports__, "PersistencePriority", (function() {
        return _EventFlagsProperties__WEBPACK_IMPORTED_MODULE_9__.d;
    })), __webpack_require__.d(__webpack_exports__, "CostPriority", (function() {
        return _EventFlagsProperties__WEBPACK_IMPORTED_MODULE_9__.a;
    })), __webpack_require__.d(__webpack_exports__, "DataCategories", (function() {
        return _EventFlagsProperties__WEBPACK_IMPORTED_MODULE_9__.b;
    })), __webpack_require__.d(__webpack_exports__, "DiagnosticLevel", (function() {
        return _EventFlagsProperties__WEBPACK_IMPORTED_MODULE_9__.c;
    }));
    var _OptionalEventFlags__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(15);
    for (var __WEBPACK_IMPORT_KEY__ in _OptionalEventFlags__WEBPACK_IMPORTED_MODULE_10__) [ "default", "Contracts", "ActivityScope", "addContractField", "getFieldsForContract", "DataClassification", "makeBooleanDataField", "makeInt64DataField", "makeDoubleDataField", "makeStringDataField", "makeGuidDataField", "DataFieldType", "getEffectiveEventFlags", "SamplingPolicy", "PersistencePriority", "CostPriority", "DataCategories", "DiagnosticLevel" ].indexOf(__WEBPACK_IMPORT_KEY__) < 0 && function(key) {
        __webpack_require__.d(__webpack_exports__, key, (function() {
            return _OptionalEventFlags__WEBPACK_IMPORTED_MODULE_10__[key];
        }));
    }(__WEBPACK_IMPORT_KEY__);
    var _OTelNotifications__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(1);
    __webpack_require__.d(__webpack_exports__, "LogLevel", (function() {
        return _OTelNotifications__WEBPACK_IMPORTED_MODULE_11__.b;
    })), __webpack_require__.d(__webpack_exports__, "Category", (function() {
        return _OTelNotifications__WEBPACK_IMPORTED_MODULE_11__.a;
    })), __webpack_require__.d(__webpack_exports__, "onNotification", (function() {
        return _OTelNotifications__WEBPACK_IMPORTED_MODULE_11__.e;
    })), __webpack_require__.d(__webpack_exports__, "logNotification", (function() {
        return _OTelNotifications__WEBPACK_IMPORTED_MODULE_11__.d;
    })), __webpack_require__.d(__webpack_exports__, "logError", (function() {
        return _OTelNotifications__WEBPACK_IMPORTED_MODULE_11__.c;
    }));
    var _SimpleTelemetryLogger__WEBPACK_IMPORTED_MODULE_12__ = __webpack_require__(8);
    __webpack_require__.d(__webpack_exports__, "SuppressNexus", (function() {
        return _SimpleTelemetryLogger__WEBPACK_IMPORTED_MODULE_12__.b;
    })), __webpack_require__.d(__webpack_exports__, "SimpleTelemetryLogger", (function() {
        return _SimpleTelemetryLogger__WEBPACK_IMPORTED_MODULE_12__.a;
    }));
    var _TelemetryLogger__WEBPACK_IMPORTED_MODULE_13__ = __webpack_require__(16);
    __webpack_require__.d(__webpack_exports__, "TelemetryLogger", (function() {
        return _TelemetryLogger__WEBPACK_IMPORTED_MODULE_13__.a;
    }));
    var _TelemetryEvent__WEBPACK_IMPORTED_MODULE_14__ = __webpack_require__(7);
    __webpack_require__.d(__webpack_exports__, "cloneEvent", (function() {
        return _TelemetryEvent__WEBPACK_IMPORTED_MODULE_14__.a;
    }));
    var _TelemetryProperties__WEBPACK_IMPORTED_MODULE_15__ = __webpack_require__(17);
    for (var __WEBPACK_IMPORT_KEY__ in _TelemetryProperties__WEBPACK_IMPORTED_MODULE_15__) [ "default", "Contracts", "ActivityScope", "addContractField", "getFieldsForContract", "DataClassification", "makeBooleanDataField", "makeInt64DataField", "makeDoubleDataField", "makeStringDataField", "makeGuidDataField", "DataFieldType", "getEffectiveEventFlags", "SamplingPolicy", "PersistencePriority", "CostPriority", "DataCategories", "DiagnosticLevel", "LogLevel", "Category", "onNotification", "logNotification", "logError", "SuppressNexus", "SimpleTelemetryLogger", "TelemetryLogger", "cloneEvent" ].indexOf(__WEBPACK_IMPORT_KEY__) < 0 && function(key) {
        __webpack_require__.d(__webpack_exports__, key, (function() {
            return _TelemetryProperties__WEBPACK_IMPORTED_MODULE_15__[key];
        }));
    }(__WEBPACK_IMPORT_KEY__);
    var _TelemetrySink__WEBPACK_IMPORTED_MODULE_16__ = __webpack_require__(18);
    for (var __WEBPACK_IMPORT_KEY__ in _TelemetrySink__WEBPACK_IMPORTED_MODULE_16__) [ "default", "Contracts", "ActivityScope", "addContractField", "getFieldsForContract", "DataClassification", "makeBooleanDataField", "makeInt64DataField", "makeDoubleDataField", "makeStringDataField", "makeGuidDataField", "DataFieldType", "getEffectiveEventFlags", "SamplingPolicy", "PersistencePriority", "CostPriority", "DataCategories", "DiagnosticLevel", "LogLevel", "Category", "onNotification", "logNotification", "logError", "SuppressNexus", "SimpleTelemetryLogger", "TelemetryLogger", "cloneEvent" ].indexOf(__WEBPACK_IMPORT_KEY__) < 0 && function(key) {
        __webpack_require__.d(__webpack_exports__, key, (function() {
            return _TelemetrySink__WEBPACK_IMPORTED_MODULE_16__[key];
        }));
    }(__WEBPACK_IMPORT_KEY__);
}, function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    __webpack_require__.d(__webpack_exports__, "a", (function() {
        return getFieldsForContract;
    }));
    var _Contract__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(2);
    function getFieldsForContract(instanceName, contractName, contractFields) {
        var dataFields = contractFields.map((function(contractField) {
            return {
                name: instanceName + "." + contractField.name,
                value: contractField.value,
                dataType: contractField.dataType
            };
        }));
        return Object(_Contract__WEBPACK_IMPORTED_MODULE_0__.a)(dataFields, instanceName, contractName), 
        dataFields;
    }
}, function(module, exports) {}, function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    __webpack_require__.d(__webpack_exports__, "a", (function() {
        return getEffectiveEventFlags;
    }));
    var _EventFlagsProperties__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(5), ___WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(1);
    function getEffectiveEventFlags(telemetryEvent) {
        var eventFlags = {
            costPriority: _EventFlagsProperties__WEBPACK_IMPORTED_MODULE_0__.a.Normal,
            samplingPolicy: _EventFlagsProperties__WEBPACK_IMPORTED_MODULE_0__.e.Measure,
            persistencePriority: _EventFlagsProperties__WEBPACK_IMPORTED_MODULE_0__.d.Normal,
            dataCategories: _EventFlagsProperties__WEBPACK_IMPORTED_MODULE_0__.b.NotSet,
            diagnosticLevel: _EventFlagsProperties__WEBPACK_IMPORTED_MODULE_0__.c.FullEvent
        };
        return telemetryEvent.eventFlags && telemetryEvent.eventFlags.dataCategories || Object(___WEBPACK_IMPORTED_MODULE_1__.d)(___WEBPACK_IMPORTED_MODULE_1__.b.Error, ___WEBPACK_IMPORTED_MODULE_1__.a.Core, (function() {
            return "Event is missing DataCategories event flag";
        })), telemetryEvent.eventFlags ? (telemetryEvent.eventFlags.costPriority && (eventFlags.costPriority = telemetryEvent.eventFlags.costPriority), 
        telemetryEvent.eventFlags.samplingPolicy && (eventFlags.samplingPolicy = telemetryEvent.eventFlags.samplingPolicy), 
        telemetryEvent.eventFlags.persistencePriority && (eventFlags.persistencePriority = telemetryEvent.eventFlags.persistencePriority), 
        telemetryEvent.eventFlags.dataCategories && (eventFlags.dataCategories = telemetryEvent.eventFlags.dataCategories), 
        telemetryEvent.eventFlags.diagnosticLevel && (eventFlags.diagnosticLevel = telemetryEvent.eventFlags.diagnosticLevel), 
        eventFlags) : eventFlags;
    }
}, function(module, exports) {}, function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    __webpack_require__.d(__webpack_exports__, "a", (function() {
        return TelemetryLogger;
    }));
    var extendStatics, _SimpleTelemetryLogger__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(8), _Activity__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(9), _contracts_Contracts__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(6), __extends = (extendStatics = function(d, b) {
        return (extendStatics = Object.setPrototypeOf || {
            __proto__: []
        } instanceof Array && function(d, b) {
            d.__proto__ = b;
        } || function(d, b) {
            for (var p in b) b.hasOwnProperty(p) && (d[p] = b[p]);
        })(d, b);
    }, function(d, b) {
        function __() {
            this.constructor = d;
        }
        extendStatics(d, b), d.prototype = null === b ? Object.create(b) : (__.prototype = b.prototype, 
        new __);
    }), __awaiter = function(thisArg, _arguments, P, generator) {
        return new (P || (P = Promise))((function(resolve, reject) {
            function fulfilled(value) {
                try {
                    step(generator.next(value));
                } catch (e) {
                    reject(e);
                }
            }
            function rejected(value) {
                try {
                    step(generator.throw(value));
                } catch (e) {
                    reject(e);
                }
            }
            function step(result) {
                var value;
                result.done ? resolve(result.value) : (value = result.value, value instanceof P ? value : new P((function(resolve) {
                    resolve(value);
                }))).then(fulfilled, rejected);
            }
            step((generator = generator.apply(thisArg, _arguments || [])).next());
        }));
    }, __generator = function(thisArg, body) {
        var f, y, t, g, _ = {
            label: 0,
            sent: function() {
                if (1 & t[0]) throw t[1];
                return t[1];
            },
            trys: [],
            ops: []
        };
        return g = {
            next: verb(0),
            throw: verb(1),
            return: verb(2)
        }, "function" == typeof Symbol && (g[Symbol.iterator] = function() {
            return this;
        }), g;
        function verb(n) {
            return function(v) {
                return function(op) {
                    if (f) throw new TypeError("Generator is already executing.");
                    for (;_; ) try {
                        if (f = 1, y && (t = 2 & op[0] ? y.return : op[0] ? y.throw || ((t = y.return) && t.call(y), 
                        0) : y.next) && !(t = t.call(y, op[1])).done) return t;
                        switch (y = 0, t && (op = [ 2 & op[0], t.value ]), op[0]) {
                          case 0:
                          case 1:
                            t = op;
                            break;

                          case 4:
                            return _.label++, {
                                value: op[1],
                                done: !1
                            };

                          case 5:
                            _.label++, y = op[1], op = [ 0 ];
                            continue;

                          case 7:
                            op = _.ops.pop(), _.trys.pop();
                            continue;

                          default:
                            if (!(t = _.trys, (t = t.length > 0 && t[t.length - 1]) || 6 !== op[0] && 2 !== op[0])) {
                                _ = 0;
                                continue;
                            }
                            if (3 === op[0] && (!t || op[1] > t[0] && op[1] < t[3])) {
                                _.label = op[1];
                                break;
                            }
                            if (6 === op[0] && _.label < t[1]) {
                                _.label = t[1], t = op;
                                break;
                            }
                            if (t && _.label < t[2]) {
                                _.label = t[2], _.ops.push(op);
                                break;
                            }
                            t[2] && _.ops.pop(), _.trys.pop();
                            continue;
                        }
                        op = body.call(thisArg, _);
                    } catch (e) {
                        op = [ 6, e ], y = 0;
                    } finally {
                        f = t = 0;
                    }
                    if (5 & op[0]) throw op[1];
                    return {
                        value: op[0] ? op[1] : void 0,
                        done: !0
                    };
                }([ n, v ]);
            };
        }
    }, TelemetryLogger = function(_super) {
        function TelemetryLogger() {
            return null !== _super && _super.apply(this, arguments) || this;
        }
        return __extends(TelemetryLogger, _super), TelemetryLogger.prototype.executeActivityAsync = function(activityName, activityBody) {
            return __awaiter(this, void 0, void 0, (function() {
                return __generator(this, (function(_a) {
                    return [ 2, this.createNewActivity(activityName).executeAsync(activityBody) ];
                }));
            }));
        }, TelemetryLogger.prototype.executeActivitySync = function(activityName, activityBody) {
            return this.createNewActivity(activityName).executeSync(activityBody);
        }, TelemetryLogger.prototype.createNewActivity = function(activityName) {
            return _Activity__WEBPACK_IMPORTED_MODULE_1__.a.createNew(this, activityName);
        }, TelemetryLogger.prototype.sendActivity = function(activityName, activity, dataFields, optionalEventFlags) {
            return this.sendTelemetryEvent({
                eventName: activityName,
                eventContract: {
                    name: _contracts_Contracts__WEBPACK_IMPORTED_MODULE_2__.a.Office.System.Activity.contractName,
                    dataFields: _contracts_Contracts__WEBPACK_IMPORTED_MODULE_2__.a.Office.System.Activity.getFields(activity)
                },
                dataFields: dataFields,
                eventFlags: optionalEventFlags
            });
        }, TelemetryLogger.prototype.sendError = function(error) {
            var dataFields = _contracts_Contracts__WEBPACK_IMPORTED_MODULE_2__.a.Office.System.Error.getFields("Error", error.error);
            return null != error.dataFields && dataFields.push.apply(dataFields, error.dataFields), 
            this.sendTelemetryEvent({
                eventName: error.eventName,
                dataFields: dataFields,
                eventFlags: error.eventFlags
            });
        }, TelemetryLogger;
    }(_SimpleTelemetryLogger__WEBPACK_IMPORTED_MODULE_0__.a);
}, function(module, exports) {}, function(module, exports) {}, function(module, exports, __webpack_require__) {
    module.exports = __webpack_require__(11);
} ]);