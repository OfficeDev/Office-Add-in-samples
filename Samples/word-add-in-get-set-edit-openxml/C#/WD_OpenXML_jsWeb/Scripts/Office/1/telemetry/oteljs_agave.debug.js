var oteljs_agave = function(modules) {
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
    }, __webpack_require__.p = "", __webpack_require__(__webpack_require__.s = 31);
}([ function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    }), function(AWTPropertyType) {
        AWTPropertyType[AWTPropertyType.Unspecified = 0] = "Unspecified", AWTPropertyType[AWTPropertyType.String = 1] = "String", 
        AWTPropertyType[AWTPropertyType.Int64 = 2] = "Int64", AWTPropertyType[AWTPropertyType.Double = 3] = "Double", 
        AWTPropertyType[AWTPropertyType.Boolean = 4] = "Boolean", AWTPropertyType[AWTPropertyType.Date = 5] = "Date";
    }(exports.AWTPropertyType || (exports.AWTPropertyType = {})), function(AWTPiiKind) {
        AWTPiiKind[AWTPiiKind.NotSet = 0] = "NotSet", AWTPiiKind[AWTPiiKind.DistinguishedName = 1] = "DistinguishedName", 
        AWTPiiKind[AWTPiiKind.GenericData = 2] = "GenericData", AWTPiiKind[AWTPiiKind.IPV4Address = 3] = "IPV4Address", 
        AWTPiiKind[AWTPiiKind.IPv6Address = 4] = "IPv6Address", AWTPiiKind[AWTPiiKind.MailSubject = 5] = "MailSubject", 
        AWTPiiKind[AWTPiiKind.PhoneNumber = 6] = "PhoneNumber", AWTPiiKind[AWTPiiKind.QueryString = 7] = "QueryString", 
        AWTPiiKind[AWTPiiKind.SipAddress = 8] = "SipAddress", AWTPiiKind[AWTPiiKind.SmtpAddress = 9] = "SmtpAddress", 
        AWTPiiKind[AWTPiiKind.Identity = 10] = "Identity", AWTPiiKind[AWTPiiKind.Uri = 11] = "Uri", 
        AWTPiiKind[AWTPiiKind.Fqdn = 12] = "Fqdn", AWTPiiKind[AWTPiiKind.IPV4AddressLegacy = 13] = "IPV4AddressLegacy";
    }(exports.AWTPiiKind || (exports.AWTPiiKind = {})), function(AWTCustomerContentKind) {
        AWTCustomerContentKind[AWTCustomerContentKind.NotSet = 0] = "NotSet", AWTCustomerContentKind[AWTCustomerContentKind.GenericContent = 1] = "GenericContent";
    }(exports.AWTCustomerContentKind || (exports.AWTCustomerContentKind = {})), function(AWTEventPriority) {
        AWTEventPriority[AWTEventPriority.Low = 1] = "Low", AWTEventPriority[AWTEventPriority.Normal = 2] = "Normal", 
        AWTEventPriority[AWTEventPriority.High = 3] = "High", AWTEventPriority[AWTEventPriority.Immediate_sync = 5] = "Immediate_sync";
    }(exports.AWTEventPriority || (exports.AWTEventPriority = {})), function(AWTEventsDroppedReason) {
        AWTEventsDroppedReason[AWTEventsDroppedReason.NonRetryableStatus = 1] = "NonRetryableStatus", 
        AWTEventsDroppedReason[AWTEventsDroppedReason.QueueFull = 3] = "QueueFull", AWTEventsDroppedReason[AWTEventsDroppedReason.MaxRetryLimit = 4] = "MaxRetryLimit";
    }(exports.AWTEventsDroppedReason || (exports.AWTEventsDroppedReason = {})), function(AWTEventsRejectedReason) {
        AWTEventsRejectedReason[AWTEventsRejectedReason.InvalidEvent = 1] = "InvalidEvent", 
        AWTEventsRejectedReason[AWTEventsRejectedReason.SizeLimitExceeded = 2] = "SizeLimitExceeded", 
        AWTEventsRejectedReason[AWTEventsRejectedReason.KillSwitch = 3] = "KillSwitch";
    }(exports.AWTEventsRejectedReason || (exports.AWTEventsRejectedReason = {}));
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var Enums_1 = __webpack_require__(0);
    exports.AWTPropertyType = Enums_1.AWTPropertyType, exports.AWTPiiKind = Enums_1.AWTPiiKind, 
    exports.AWTEventPriority = Enums_1.AWTEventPriority, exports.AWTEventsDroppedReason = Enums_1.AWTEventsDroppedReason, 
    exports.AWTEventsRejectedReason = Enums_1.AWTEventsRejectedReason, exports.AWTCustomerContentKind = Enums_1.AWTCustomerContentKind;
    var Enums_2 = __webpack_require__(6);
    exports.AWTUserIdType = Enums_2.AWTUserIdType, exports.AWTSessionState = Enums_2.AWTSessionState;
    var DataModels_1 = __webpack_require__(12);
    exports.AWT_BEST_EFFORT = DataModels_1.AWT_BEST_EFFORT, exports.AWT_NEAR_REAL_TIME = DataModels_1.AWT_NEAR_REAL_TIME, 
    exports.AWT_REAL_TIME = DataModels_1.AWT_REAL_TIME;
    var AWTEventProperties_1 = __webpack_require__(7);
    exports.AWTEventProperties = AWTEventProperties_1.default;
    var AWTLogger_1 = __webpack_require__(13);
    exports.AWTLogger = AWTLogger_1.default;
    var AWTLogManager_1 = __webpack_require__(17);
    exports.AWTLogManager = AWTLogManager_1.default;
    var AWTTransmissionManager_1 = __webpack_require__(30);
    exports.AWTTransmissionManager = AWTTransmissionManager_1.default;
    var AWTSerializer_1 = __webpack_require__(15);
    exports.AWTSerializer = AWTSerializer_1.default;
    var AWTSemanticContext_1 = __webpack_require__(9);
    exports.AWTSemanticContext = AWTSemanticContext_1.default, exports.AWT_COLLECTOR_URL_UNITED_STATES = "https://us.pipe.aria.microsoft.com/Collector/3.0/", 
    exports.AWT_COLLECTOR_URL_GERMANY = "https://de.pipe.aria.microsoft.com/Collector/3.0/", 
    exports.AWT_COLLECTOR_URL_JAPAN = "https://jp.pipe.aria.microsoft.com/Collector/3.0/", 
    exports.AWT_COLLECTOR_URL_AUSTRALIA = "https://au.pipe.aria.microsoft.com/Collector/3.0/", 
    exports.AWT_COLLECTOR_URL_EUROPE = "https://eu.pipe.aria.microsoft.com/Collector/3.0/", 
    exports.AWT_COLLECTOR_URL_USGOV_DOD = "https://pf.pipe.aria.microsoft.com/Collector/3.0", 
    exports.AWT_COLLECTOR_URL_USGOV_DOJ = "https://tb.pipe.aria.microsoft.com/Collector/3.0";
}, , function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var microsoft_bond_primitives_1 = __webpack_require__(8), Enums_1 = __webpack_require__(0), GuidRegex = /[xy]/g;
    exports.EventNameAndTypeRegex = /^[a-zA-Z]([a-zA-Z0-9]|_){2,98}[a-zA-Z0-9]$/, exports.EventNameDotRegex = /\./g, 
    exports.PropertyNameRegex = /^[a-zA-Z](([a-zA-Z0-9|_|\.]){0,98}[a-zA-Z0-9])?$/, 
    exports.StatsApiKey = "a387cfcf60114a43a7699f9fbb49289e-9bceb9fe-1c06-460f-96c5-6a0b247358bc-7238";
    var beaconsSupported = null, uInt8ArraySupported = null, useXDR = null;
    function isString(value) {
        return "string" == typeof value;
    }
    function isNumber(value) {
        return "number" == typeof value;
    }
    function isBoolean(value) {
        return "boolean" == typeof value;
    }
    function isDate(value) {
        return value instanceof Date;
    }
    function msToTicks(timeInMs) {
        return 1e4 * (timeInMs + 621355968e5);
    }
    function isReactNative() {
        return !("undefined" == typeof navigator || !navigator.product) && "ReactNative" === navigator.product;
    }
    function isServiceWorkerGlobalScope() {
        return "object" == typeof self && "ServiceWorkerGlobalScope" === self.constructor.name;
    }
    function twoDigit(n) {
        return n < 10 ? "0" + n : n.toString();
    }
    function isNotDefined(value) {
        return null == value || "" === value;
    }
    exports.numberToBondInt64 = function(value) {
        var bond_value = new microsoft_bond_primitives_1.Int64("0");
        return bond_value.low = 4294967295 & value, bond_value.high = Math.floor(value / 4294967296), 
        bond_value;
    }, exports.newGuid = function() {
        return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(GuidRegex, (function(c) {
            var r = 16 * Math.random() | 0;
            return ("x" === c ? r : 3 & r | 8).toString(16);
        }));
    }, exports.isString = isString, exports.isNumber = isNumber, exports.isBoolean = isBoolean, 
    exports.isDate = isDate, exports.msToTicks = msToTicks, exports.getTenantId = function(apiKey) {
        var indexTenantId = apiKey.indexOf("-");
        return indexTenantId > -1 ? apiKey.substring(0, indexTenantId) : "";
    }, exports.isBeaconsSupported = function() {
        return null === beaconsSupported && (beaconsSupported = "undefined" != typeof navigator && Boolean(navigator.sendBeacon)), 
        beaconsSupported;
    }, exports.isUint8ArrayAvailable = function() {
        return null === uInt8ArraySupported && (uInt8ArraySupported = "undefined" != typeof Uint8Array && !function() {
            if ("undefined" != typeof navigator && navigator.userAgent) {
                var ua = navigator.userAgent.toLowerCase();
                if ((ua.indexOf("safari") >= 0 || ua.indexOf("firefox") >= 0) && ua.indexOf("chrome") < 0) return !0;
            }
            return !1;
        }() && !isReactNative()), uInt8ArraySupported;
    }, exports.isPriority = function(value) {
        return !(!isNumber(value) || !(value >= 1 && value <= 3 || 5 === value));
    }, exports.sanitizeProperty = function(name, property) {
        return !exports.PropertyNameRegex.test(name) || isNotDefined(property) ? null : (isNotDefined(property.value) && (property = {
            value: property,
            type: Enums_1.AWTPropertyType.Unspecified
        }), property.type = function(value, type) {
            switch (type = function(value) {
                if (isNumber(value) && value >= 0 && value <= 4) return !0;
                return !1;
            }(type) ? type : Enums_1.AWTPropertyType.Unspecified) {
              case Enums_1.AWTPropertyType.Unspecified:
                return function(value) {
                    switch (typeof value) {
                      case "string":
                        return Enums_1.AWTPropertyType.String;

                      case "boolean":
                        return Enums_1.AWTPropertyType.Boolean;

                      case "number":
                        return Enums_1.AWTPropertyType.Double;

                      case "object":
                        return isDate(value) ? Enums_1.AWTPropertyType.Date : null;
                    }
                    return null;
                }(value);

              case Enums_1.AWTPropertyType.String:
                return isString(value) ? type : null;

              case Enums_1.AWTPropertyType.Boolean:
                return isBoolean(value) ? type : null;

              case Enums_1.AWTPropertyType.Date:
                return isDate(value) && NaN !== value.getTime() ? type : null;

              case Enums_1.AWTPropertyType.Int64:
                return isNumber(value) && value % 1 == 0 ? type : null;

              case Enums_1.AWTPropertyType.Double:
                return isNumber(value) ? type : null;
            }
            return null;
        }(property.value, property.type), property.type ? (isDate(property.value) && (property.value = msToTicks(property.value.getTime())), 
        property.pii > 0 && property.cc > 0 ? null : property.pii ? function(value) {
            if (isNumber(value) && value >= 0 && value <= 13) return !0;
            return !1;
        }(property.pii) ? property : null : property.cc ? function(value) {
            if (isNumber(value) && value >= 0 && value <= 1) return !0;
            return !1;
        }(property.cc) ? property : null : property) : null);
    }, exports.getISOString = function(date) {
        return date.getUTCFullYear() + "-" + twoDigit(date.getUTCMonth() + 1) + "-" + twoDigit(date.getUTCDate()) + "T" + twoDigit(date.getUTCHours()) + ":" + twoDigit(date.getUTCMinutes()) + ":" + twoDigit(date.getUTCSeconds()) + "." + function(n) {
            if (n < 10) return "00" + n;
            if (n < 100) return "0" + n;
            return n.toString();
        }(date.getUTCMilliseconds()) + "Z";
    }, exports.useXDomainRequest = function() {
        if (null === useXDR) {
            var conn = new XMLHttpRequest;
            useXDR = "undefined" == typeof conn.withCredentials && "undefined" != typeof XDomainRequest;
        }
        return useXDR;
    }, exports.useFetchRequest = function() {
        return isReactNative() || isServiceWorkerGlobalScope();
    }, exports.isReactNative = isReactNative, exports.isServiceWorkerGlobalScope = isServiceWorkerGlobalScope;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var DataModels_1 = __webpack_require__(12), Enums_1 = __webpack_require__(0), AWTQueueManager_1 = __webpack_require__(11), AWTStatsManager_1 = __webpack_require__(14), AWTEventProperties_1 = __webpack_require__(7), AWTLogManager_1 = __webpack_require__(17), Utils = __webpack_require__(3), AWTTransmissionManagerCore = function() {
        function AWTTransmissionManagerCore() {}
        return AWTTransmissionManagerCore.setEventsHandler = function(eventsHandler) {
            this._eventHandler = eventsHandler;
        }, AWTTransmissionManagerCore.getEventsHandler = function() {
            return this._eventHandler;
        }, AWTTransmissionManagerCore.scheduleTimer = function() {
            var _this = this, timer = this._profiles[this._currentProfile][2];
            this._timeout < 0 && timer >= 0 && !this._paused && (this._eventHandler.hasEvents() ? (0 === timer && this._currentBackoffCount > 0 && (timer = 1), 
            this._timeout = setTimeout((function() {
                return _this._batchAndSendEvents();
            }), timer * (1 << this._currentBackoffCount) * 1e3)) : this._timerCount = 0);
        }, AWTTransmissionManagerCore.initialize = function(config) {
            var _this = this;
            this._newEventsAllowed = !0, this._config = config, this._eventHandler = new AWTQueueManager_1.default(config.collectorUri, config.cacheMemorySizeLimitInNumberOfEvents, config.httpXHROverride, config.clockSkewRefreshDurationInMins), 
            this._initializeProfiles(), AWTStatsManager_1.default.initialize((function(stats, tenantId) {
                if (_this._config.canSendStatEvent("awt_stats")) {
                    var event_1 = new AWTEventProperties_1.default("awt_stats");
                    for (var statKey in event_1.setEventPriority(Enums_1.AWTEventPriority.High), event_1.setProperty("TenantId", tenantId), 
                    stats) stats.hasOwnProperty(statKey) && event_1.setProperty(statKey, stats[statKey].toString());
                    AWTLogManager_1.default.getLogger(Utils.StatsApiKey).logEvent(event_1);
                }
            }));
        }, AWTTransmissionManagerCore.setTransmitProfile = function(profileName) {
            this._currentProfile !== profileName && void 0 !== this._profiles[profileName] && (this.clearTimeout(), 
            this._currentProfile = profileName, this.scheduleTimer());
        }, AWTTransmissionManagerCore.loadTransmitProfiles = function(profiles) {
            for (var profileName in this._resetTransmitProfiles(), profiles) if (profiles.hasOwnProperty(profileName)) {
                if (3 !== profiles[profileName].length) continue;
                for (var i = 2; i >= 0; --i) if (profiles[profileName][i] < 0) {
                    for (var j = i; j >= 0; --j) profiles[profileName][j] = -1;
                    break;
                }
                for (i = 2; i > 0; --i) if (profiles[profileName][i] > 0 && profiles[profileName][i - 1] > 0) {
                    var timerMultiplier = profiles[profileName][i - 1] / profiles[profileName][i];
                    profiles[profileName][i - 1] = Math.ceil(timerMultiplier) * profiles[profileName][i];
                }
                this._profiles[profileName] = profiles[profileName];
            }
        }, AWTTransmissionManagerCore.sendEvent = function(event) {
            this._newEventsAllowed && (this._currentBackoffCount > 0 && event.priority === Enums_1.AWTEventPriority.Immediate_sync && (event.priority = Enums_1.AWTEventPriority.High), 
            this._eventHandler.addEvent(event), this.scheduleTimer());
        }, AWTTransmissionManagerCore.flush = function(callback) {
            var currentTime = (new Date).getTime();
            !this._paused && this._lastUploadNowCall + 3e4 < currentTime && (this._lastUploadNowCall = currentTime, 
            this._timeout > -1 && (clearTimeout(this._timeout), this._timeout = -1), this._eventHandler.uploadNow(callback));
        }, AWTTransmissionManagerCore.pauseTransmission = function() {
            this._paused || (this.clearTimeout(), this._eventHandler.pauseTransmission(), this._paused = !0);
        }, AWTTransmissionManagerCore.resumeTransmision = function() {
            this._paused && (this._paused = !1, this._eventHandler.resumeTransmission(), this.scheduleTimer());
        }, AWTTransmissionManagerCore.flushAndTeardown = function() {
            AWTStatsManager_1.default.teardown(), this._newEventsAllowed = !1, this.clearTimeout(), 
            this._eventHandler.teardown();
        }, AWTTransmissionManagerCore.backOffTransmission = function() {
            this._currentBackoffCount < 4 && (this._currentBackoffCount++, this.clearTimeout(), 
            this.scheduleTimer());
        }, AWTTransmissionManagerCore.clearBackOff = function() {
            this._currentBackoffCount > 0 && (this._currentBackoffCount = 0, this.clearTimeout(), 
            this.scheduleTimer());
        }, AWTTransmissionManagerCore._resetTransmitProfiles = function() {
            this.clearTimeout(), this._initializeProfiles(), this._currentProfile = DataModels_1.AWT_REAL_TIME, 
            this.scheduleTimer();
        }, AWTTransmissionManagerCore.clearTimeout = function() {
            this._timeout > 0 && (clearTimeout(this._timeout), this._timeout = -1, this._timerCount = 0);
        }, AWTTransmissionManagerCore._batchAndSendEvents = function() {
            var priority = Enums_1.AWTEventPriority.High;
            this._timerCount++, this._timerCount * this._profiles[this._currentProfile][2] === this._profiles[this._currentProfile][0] ? (priority = Enums_1.AWTEventPriority.Low, 
            this._timerCount = 0) : this._timerCount * this._profiles[this._currentProfile][2] === this._profiles[this._currentProfile][1] && (priority = Enums_1.AWTEventPriority.Normal), 
            this._eventHandler.sendEventsForPriorityAndAbove(priority), this._timeout = -1, 
            this.scheduleTimer();
        }, AWTTransmissionManagerCore._initializeProfiles = function() {
            this._profiles = {}, this._profiles[DataModels_1.AWT_REAL_TIME] = [ 4, 2, 1 ], this._profiles[DataModels_1.AWT_NEAR_REAL_TIME] = [ 12, 6, 3 ], 
            this._profiles[DataModels_1.AWT_BEST_EFFORT] = [ 36, 18, 9 ];
        }, AWTTransmissionManagerCore._newEventsAllowed = !1, AWTTransmissionManagerCore._currentProfile = DataModels_1.AWT_REAL_TIME, 
        AWTTransmissionManagerCore._timeout = -1, AWTTransmissionManagerCore._currentBackoffCount = 0, 
        AWTTransmissionManagerCore._paused = !1, AWTTransmissionManagerCore._timerCount = 0, 
        AWTTransmissionManagerCore._lastUploadNowCall = 0, AWTTransmissionManagerCore;
    }();
    exports.default = AWTTransmissionManagerCore;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var AWTNotificationManager = function() {
        function AWTNotificationManager() {}
        return AWTNotificationManager.addNotificationListener = function(listener) {
            this.listeners.push(listener);
        }, AWTNotificationManager.removeNotificationListener = function(listener) {
            for (var index = this.listeners.indexOf(listener); index > -1; ) this.listeners.splice(index, 1), 
            index = this.listeners.indexOf(listener);
        }, AWTNotificationManager.eventsSent = function(events) {
            for (var _this = this, _loop_1 = function(i) {
                this_1.listeners[i].eventsSent && setTimeout((function() {
                    return _this.listeners[i].eventsSent(events);
                }), 0);
            }, this_1 = this, i = 0; i < this.listeners.length; ++i) _loop_1(i);
        }, AWTNotificationManager.eventsDropped = function(events, reason) {
            for (var _this = this, _loop_2 = function(i) {
                this_2.listeners[i].eventsDropped && setTimeout((function() {
                    return _this.listeners[i].eventsDropped(events, reason);
                }), 0);
            }, this_2 = this, i = 0; i < this.listeners.length; ++i) _loop_2(i);
        }, AWTNotificationManager.eventsRetrying = function(events) {
            for (var _this = this, _loop_3 = function(i) {
                this_3.listeners[i].eventsRetrying && setTimeout((function() {
                    return _this.listeners[i].eventsRetrying(events);
                }), 0);
            }, this_3 = this, i = 0; i < this.listeners.length; ++i) _loop_3(i);
        }, AWTNotificationManager.eventsRejected = function(events, reason) {
            for (var _this = this, _loop_4 = function(i) {
                this_4.listeners[i].eventsRejected && setTimeout((function() {
                    return _this.listeners[i].eventsRejected(events, reason);
                }), 0);
            }, this_4 = this, i = 0; i < this.listeners.length; ++i) _loop_4(i);
        }, AWTNotificationManager.listeners = [], AWTNotificationManager;
    }();
    exports.default = AWTNotificationManager;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    }), function(AWTUserIdType) {
        AWTUserIdType[AWTUserIdType.Unknown = 0] = "Unknown", AWTUserIdType[AWTUserIdType.MSACID = 1] = "MSACID", 
        AWTUserIdType[AWTUserIdType.MSAPUID = 2] = "MSAPUID", AWTUserIdType[AWTUserIdType.ANID = 3] = "ANID", 
        AWTUserIdType[AWTUserIdType.OrgIdCID = 4] = "OrgIdCID", AWTUserIdType[AWTUserIdType.OrgIdPUID = 5] = "OrgIdPUID", 
        AWTUserIdType[AWTUserIdType.UserObjectId = 6] = "UserObjectId", AWTUserIdType[AWTUserIdType.Skype = 7] = "Skype", 
        AWTUserIdType[AWTUserIdType.Yammer = 8] = "Yammer", AWTUserIdType[AWTUserIdType.EmailAddress = 9] = "EmailAddress", 
        AWTUserIdType[AWTUserIdType.PhoneNumber = 10] = "PhoneNumber", AWTUserIdType[AWTUserIdType.SipAddress = 11] = "SipAddress", 
        AWTUserIdType[AWTUserIdType.MUID = 12] = "MUID";
    }(exports.AWTUserIdType || (exports.AWTUserIdType = {})), function(AWTSessionState) {
        AWTSessionState[AWTSessionState.Started = 0] = "Started", AWTSessionState[AWTSessionState.Ended = 1] = "Ended";
    }(exports.AWTSessionState || (exports.AWTSessionState = {}));
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var Utils = __webpack_require__(3), Enums_1 = __webpack_require__(0), AWTEventProperties = function() {
        function AWTEventProperties(name) {
            this._event = {
                name: "",
                properties: {}
            }, name && this.setName(name);
        }
        return AWTEventProperties.prototype.setName = function(name) {
            this._event.name = name;
        }, AWTEventProperties.prototype.getName = function() {
            return this._event.name;
        }, AWTEventProperties.prototype.setType = function(type) {
            this._event.type = type;
        }, AWTEventProperties.prototype.getType = function() {
            return this._event.type;
        }, AWTEventProperties.prototype.setTimestamp = function(timestampInEpochMillis) {
            this._event.timestamp = timestampInEpochMillis;
        }, AWTEventProperties.prototype.getTimestamp = function() {
            return this._event.timestamp;
        }, AWTEventProperties.prototype.setEventPriority = function(priority) {
            this._event.priority = priority;
        }, AWTEventProperties.prototype.getEventPriority = function() {
            return this._event.priority;
        }, AWTEventProperties.prototype.setProperty = function(name, value, type) {
            void 0 === type && (type = Enums_1.AWTPropertyType.Unspecified);
            var property = {
                value: value,
                type: type,
                pii: Enums_1.AWTPiiKind.NotSet,
                cc: Enums_1.AWTCustomerContentKind.NotSet
            };
            null !== (property = Utils.sanitizeProperty(name, property)) ? this._event.properties[name] = property : delete this._event.properties[name];
        }, AWTEventProperties.prototype.setPropertyWithPii = function(name, value, pii, type) {
            void 0 === type && (type = Enums_1.AWTPropertyType.Unspecified);
            var property = {
                value: value,
                type: type,
                pii: pii,
                cc: Enums_1.AWTCustomerContentKind.NotSet
            };
            null !== (property = Utils.sanitizeProperty(name, property)) ? this._event.properties[name] = property : delete this._event.properties[name];
        }, AWTEventProperties.prototype.setPropertyWithCustomerContent = function(name, value, customerContent, type) {
            void 0 === type && (type = Enums_1.AWTPropertyType.Unspecified);
            var property = {
                value: value,
                type: type,
                pii: Enums_1.AWTPiiKind.NotSet,
                cc: customerContent
            };
            null !== (property = Utils.sanitizeProperty(name, property)) ? this._event.properties[name] = property : delete this._event.properties[name];
        }, AWTEventProperties.prototype.getPropertyMap = function() {
            return this._event.properties;
        }, AWTEventProperties.prototype.getEvent = function() {
            return this._event;
        }, AWTEventProperties;
    }();
    exports.default = AWTEventProperties;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var Int64 = function() {
        function Int64(numberStr) {
            this.low = 0, this.high = 0, this.low = parseInt(numberStr, 10), this.low < 0 && (this.high = -1);
        }
        return Int64.prototype._Equals = function(numberStr) {
            var tmp = new Int64(numberStr);
            return this.low === tmp.low && this.high === tmp.high;
        }, Int64;
    }();
    exports.Int64 = Int64;
    var UInt64 = function() {
        function UInt64(numberStr) {
            this.low = 0, this.high = 0, this.low = parseInt(numberStr, 10);
        }
        return UInt64.prototype._Equals = function(numberStr) {
            var tmp = new UInt64(numberStr);
            return this.low === tmp.low && this.high === tmp.high;
        }, UInt64;
    }();
    exports.UInt64 = UInt64;
    var Number = function() {
        function Number() {}
        return Number._ToByte = function(value) {
            return this._ToUInt8(value);
        }, Number._ToUInt8 = function(value) {
            return 255 & value;
        }, Number._ToInt32 = function(value) {
            return 2147483647 & value | 2147483648 & value;
        }, Number._ToUInt32 = function(value) {
            return 4294967295 & value;
        }, Number;
    }();
    exports.Number = Number;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var AWTAutoCollection_1 = __webpack_require__(10), Enums_1 = __webpack_require__(0), Enums_2 = __webpack_require__(6), AWTSemanticContext = function() {
        function AWTSemanticContext(_allowDeviceFields, _properties) {
            this._allowDeviceFields = _allowDeviceFields, this._properties = _properties;
        }
        return AWTSemanticContext.prototype.setAppId = function(appId) {
            this._addContext("AppInfo.Id", appId);
        }, AWTSemanticContext.prototype.setAppVersion = function(appVersion) {
            this._addContext("AppInfo.Version", appVersion);
        }, AWTSemanticContext.prototype.setAppLanguage = function(appLanguage) {
            this._addContext("AppInfo.Language", appLanguage);
        }, AWTSemanticContext.prototype.setDeviceId = function(deviceId) {
            this._allowDeviceFields && (AWTAutoCollection_1.default.checkAndSaveDeviceId(deviceId), 
            this._addContext("DeviceInfo.Id", deviceId));
        }, AWTSemanticContext.prototype.setDeviceOsName = function(deviceOsName) {
            this._allowDeviceFields && this._addContext("DeviceInfo.OsName", deviceOsName);
        }, AWTSemanticContext.prototype.setDeviceOsVersion = function(deviceOsVersion) {
            this._allowDeviceFields && this._addContext("DeviceInfo.OsVersion", deviceOsVersion);
        }, AWTSemanticContext.prototype.setDeviceBrowserName = function(deviceBrowserName) {
            this._allowDeviceFields && this._addContext("DeviceInfo.BrowserName", deviceBrowserName);
        }, AWTSemanticContext.prototype.setDeviceBrowserVersion = function(deviceBrowserVersion) {
            this._allowDeviceFields && this._addContext("DeviceInfo.BrowserVersion", deviceBrowserVersion);
        }, AWTSemanticContext.prototype.setDeviceMake = function(deviceMake) {
            this._allowDeviceFields && this._addContext("DeviceInfo.Make", deviceMake);
        }, AWTSemanticContext.prototype.setDeviceModel = function(deviceModel) {
            this._allowDeviceFields && this._addContext("DeviceInfo.Model", deviceModel);
        }, AWTSemanticContext.prototype.setUserId = function(userId, pii, userIdType) {
            if (!isNaN(userIdType) && null !== userIdType && userIdType >= 0 && userIdType <= 12) this._addContext("UserInfo.IdType", userIdType.toString()); else {
                var inferredUserIdType = void 0;
                switch (pii) {
                  case Enums_1.AWTPiiKind.SipAddress:
                    inferredUserIdType = Enums_2.AWTUserIdType.SipAddress;
                    break;

                  case Enums_1.AWTPiiKind.PhoneNumber:
                    inferredUserIdType = Enums_2.AWTUserIdType.PhoneNumber;
                    break;

                  case Enums_1.AWTPiiKind.SmtpAddress:
                    inferredUserIdType = Enums_2.AWTUserIdType.EmailAddress;
                    break;

                  default:
                    inferredUserIdType = Enums_2.AWTUserIdType.Unknown;
                }
                this._addContext("UserInfo.IdType", inferredUserIdType.toString());
            }
            if (isNaN(pii) || null === pii || pii === Enums_1.AWTPiiKind.NotSet || pii > 13) switch (userIdType) {
              case Enums_2.AWTUserIdType.Skype:
                pii = Enums_1.AWTPiiKind.Identity;
                break;

              case Enums_2.AWTUserIdType.EmailAddress:
                pii = Enums_1.AWTPiiKind.SmtpAddress;
                break;

              case Enums_2.AWTUserIdType.PhoneNumber:
                pii = Enums_1.AWTPiiKind.PhoneNumber;
                break;

              case Enums_2.AWTUserIdType.SipAddress:
                pii = Enums_1.AWTPiiKind.SipAddress;
                break;

              default:
                pii = Enums_1.AWTPiiKind.NotSet;
            }
            this._addContextWithPii("UserInfo.Id", userId, pii);
        }, AWTSemanticContext.prototype.setUserAdvertisingId = function(userAdvertisingId) {
            this._addContext("UserInfo.AdvertisingId", userAdvertisingId);
        }, AWTSemanticContext.prototype.setUserTimeZone = function(userTimeZone) {
            this._addContext("UserInfo.TimeZone", userTimeZone);
        }, AWTSemanticContext.prototype.setUserLanguage = function(userLanguage) {
            this._addContext("UserInfo.Language", userLanguage);
        }, AWTSemanticContext.prototype._addContext = function(key, value) {
            "string" == typeof value && this._properties.setProperty(key, value);
        }, AWTSemanticContext.prototype._addContextWithPii = function(key, value, pii) {
            "string" == typeof value && this._properties.setPropertyWithPii(key, value, pii);
        }, AWTSemanticContext;
    }();
    exports.default = AWTSemanticContext;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var Utils = __webpack_require__(3), DEVICE_ID_COOKIE = "MicrosoftApplicationsTelemetryDeviceId", FIRSTLAUNCHTIME_COOKIE = "MicrosoftApplicationsTelemetryFirstLaunchTime", BROWSERS_MSIE = "MSIE", BROWSERS_CHROME = "Chrome", BROWSERS_FIREFOX = "Firefox", BROWSERS_SAFARI = "Safari", BROWSERS_EDGE = "Edge", BROWSERS_ELECTRON = "Electron", BROWSERS_SKYPE_SHELL = "SkypeShell", BROWSERS_PHANTOMJS = "PhantomJS", BROWSERS_OPERA = "Opera", OPERATING_SYSTEMS_WINDOWS = "Windows", OPERATING_SYSTEMS_MACOSX = "Mac OS X", OPERATING_SYSTEMS_WINDOWS_PHONE = "Windows Phone", OPERATING_SYSTEMS_WINDOWS_RT = "Windows RT", OPERATING_SYSTEMS_IOS = "iOS", OPERATING_SYSTEMS_ANDROID = "Android", OPERATING_SYSTEMS_LINUX = "Linux", OPERATING_SYSTEMS_CROS = "Chrome OS", OSNAMEREGEX_WIN = /(windows|win32)/i, OSNAMEREGEX_WINRT = / arm;/i, OSNAMEREGEX_WINPHONE = /windows\sphone\s\d+\.\d+/i, OSNAMEREGEX_OSX = /(macintosh|mac os x)/i, OSNAMEREGEX_IOS = /(iPad|iPhone|iPod)(?=.*like Mac OS X)/i, OSNAMEREGEX_LINUX = /(linux|joli|[kxln]?ubuntu|debian|[open]*suse|gentoo|arch|slackware|fedora|mandriva|centos|pclinuxos|redhat|zenwalk)/i, OSNAMEREGEX_ANDROID = /android/i, OSNAMEREGEX_CROS = /CrOS/i, VERSION_MAPPINGS = {
        5.1: "XP",
        "6.0": "Vista",
        6.1: "7",
        6.2: "8",
        6.3: "8.1",
        "10.0": "10"
    }, AWTAutoCollection = function() {
        function AWTAutoCollection() {}
        return AWTAutoCollection.addPropertyStorageOverride = function(propertyStorage) {
            return !!propertyStorage && (this._propertyStorage = propertyStorage, !0);
        }, AWTAutoCollection.autoCollect = function(semanticContext, disableCookies, userAgent) {
            if (this._semanticContext = semanticContext, this._disableCookies = disableCookies, 
            this._autoCollect(), userAgent || "undefined" == typeof navigator || (userAgent = navigator.userAgent || ""), 
            this._autoCollectFromUserAgent(userAgent), this._disableCookies && !this._propertyStorage) return this._deleteCookie(DEVICE_ID_COOKIE), 
            void this._deleteCookie(FIRSTLAUNCHTIME_COOKIE);
            (this._propertyStorage || this._areCookiesAvailable && !this._disableCookies) && this._autoCollectDeviceId();
        }, AWTAutoCollection.checkAndSaveDeviceId = function(deviceId) {
            if (deviceId) {
                var oldDeviceId = this._getData(DEVICE_ID_COOKIE), flt = this._getData(FIRSTLAUNCHTIME_COOKIE);
                oldDeviceId !== deviceId && (flt = Utils.getISOString(new Date)), this._saveData(DEVICE_ID_COOKIE, deviceId), 
                this._saveData(FIRSTLAUNCHTIME_COOKIE, flt), this._setFirstLaunchTime(flt);
            }
        }, AWTAutoCollection._autoCollectDeviceId = function() {
            var deviceId = this._getData(DEVICE_ID_COOKIE);
            deviceId || (deviceId = Utils.newGuid()), this._semanticContext.setDeviceId(deviceId);
        }, AWTAutoCollection._autoCollect = function() {
            "undefined" != typeof document && document.documentElement && this._semanticContext.setAppLanguage(document.documentElement.lang), 
            "undefined" != typeof navigator && this._semanticContext.setUserLanguage(navigator.userLanguage || navigator.language);
            var timeZone = (new Date).getTimezoneOffset(), minutes = timeZone % 60, hours = (timeZone - minutes) / 60, timeZonePrefix = "+";
            hours > 0 && (timeZonePrefix = "-"), hours = Math.abs(hours), minutes = Math.abs(minutes), 
            this._semanticContext.setUserTimeZone(timeZonePrefix + (hours < 10 ? "0" + hours : hours.toString()) + ":" + (minutes < 10 ? "0" + minutes : minutes.toString()));
        }, AWTAutoCollection._autoCollectFromUserAgent = function(userAgent) {
            if (userAgent) {
                var browserName = this._getBrowserName(userAgent);
                this._semanticContext.setDeviceBrowserName(browserName), this._semanticContext.setDeviceBrowserVersion(this._getBrowserVersion(userAgent, browserName));
                var osName = this._getOsName(userAgent);
                this._semanticContext.setDeviceOsName(osName), this._semanticContext.setDeviceOsVersion(this._getOsVersion(userAgent, osName));
            }
        }, AWTAutoCollection._getBrowserName = function(userAgent) {
            return this._userAgentContainsString("OPR/", userAgent) ? BROWSERS_OPERA : this._userAgentContainsString(BROWSERS_PHANTOMJS, userAgent) ? BROWSERS_PHANTOMJS : this._userAgentContainsString(BROWSERS_EDGE, userAgent) || this._userAgentContainsString("Edg", userAgent) ? BROWSERS_EDGE : this._userAgentContainsString(BROWSERS_ELECTRON, userAgent) ? BROWSERS_ELECTRON : this._userAgentContainsString(BROWSERS_CHROME, userAgent) ? BROWSERS_CHROME : this._userAgentContainsString("Trident", userAgent) ? BROWSERS_MSIE : this._userAgentContainsString(BROWSERS_FIREFOX, userAgent) ? BROWSERS_FIREFOX : this._userAgentContainsString(BROWSERS_SAFARI, userAgent) ? BROWSERS_SAFARI : this._userAgentContainsString(BROWSERS_SKYPE_SHELL, userAgent) ? BROWSERS_SKYPE_SHELL : "Unknown";
        }, AWTAutoCollection._setFirstLaunchTime = function(flt) {
            if (!isNaN(flt)) {
                var fltDate = new Date;
                fltDate.setTime(parseInt(flt, 10)), flt = Utils.getISOString(fltDate);
            }
            this.firstLaunchTime = flt;
        }, AWTAutoCollection._userAgentContainsString = function(searchString, userAgent) {
            return userAgent.indexOf(searchString) > -1;
        }, AWTAutoCollection._getBrowserVersion = function(userAgent, browserName) {
            if (browserName === BROWSERS_MSIE) return this._getIeVersion(userAgent);
            if (browserName === BROWSERS_EDGE) {
                var version = this._getOtherVersion(browserName, userAgent);
                return "Unknown" === version ? this._getOtherVersion("Edg", userAgent) : version;
            }
            return this._getOtherVersion(browserName, userAgent);
        }, AWTAutoCollection._getIeVersion = function(userAgent) {
            var classicIeVersionMatches = userAgent.match(new RegExp(BROWSERS_MSIE + " ([\\d,.]+)"));
            if (classicIeVersionMatches) return classicIeVersionMatches[1];
            var ieVersionMatches = userAgent.match(new RegExp("rv:([\\d,.]+)"));
            return ieVersionMatches ? ieVersionMatches[1] : void 0;
        }, AWTAutoCollection._getOtherVersion = function(browserString, userAgent) {
            browserString === BROWSERS_SAFARI && (browserString = "Version");
            var matches = userAgent.match(new RegExp(browserString + "/([\\d,.]+)"));
            return matches ? matches[1] : "Unknown";
        }, AWTAutoCollection._getOsName = function(userAgent) {
            return userAgent.match(OSNAMEREGEX_WINPHONE) ? OPERATING_SYSTEMS_WINDOWS_PHONE : userAgent.match(OSNAMEREGEX_WINRT) ? OPERATING_SYSTEMS_WINDOWS_RT : userAgent.match(OSNAMEREGEX_IOS) ? OPERATING_SYSTEMS_IOS : userAgent.match(OSNAMEREGEX_ANDROID) ? OPERATING_SYSTEMS_ANDROID : userAgent.match(OSNAMEREGEX_LINUX) ? OPERATING_SYSTEMS_LINUX : userAgent.match(OSNAMEREGEX_OSX) ? OPERATING_SYSTEMS_MACOSX : userAgent.match(OSNAMEREGEX_WIN) ? OPERATING_SYSTEMS_WINDOWS : userAgent.match(OSNAMEREGEX_CROS) ? OPERATING_SYSTEMS_CROS : "Unknown";
        }, AWTAutoCollection._getOsVersion = function(userAgent, osName) {
            return osName === OPERATING_SYSTEMS_WINDOWS ? this._getGenericOsVersion(userAgent, "Windows NT") : osName === OPERATING_SYSTEMS_ANDROID ? this._getGenericOsVersion(userAgent, osName) : osName === OPERATING_SYSTEMS_MACOSX ? this._getMacOsxVersion(userAgent) : "Unknown";
        }, AWTAutoCollection._getGenericOsVersion = function(userAgent, osName) {
            var ntVersionMatches = userAgent.match(new RegExp(osName + " ([\\d,.]+)"));
            return ntVersionMatches ? VERSION_MAPPINGS[ntVersionMatches[1]] ? VERSION_MAPPINGS[ntVersionMatches[1]] : ntVersionMatches[1] : "Unknown";
        }, AWTAutoCollection._getMacOsxVersion = function(userAgent) {
            var macOsxVersionInUserAgentMatches = userAgent.match(new RegExp(OPERATING_SYSTEMS_MACOSX + " ([\\d,_,.]+)"));
            if (macOsxVersionInUserAgentMatches) {
                var versionString = macOsxVersionInUserAgentMatches[1].replace(/_/g, ".");
                if (versionString) {
                    var delimiter = this._getDelimiter(versionString);
                    return delimiter ? versionString.split(delimiter)[0] : versionString;
                }
            }
            return "Unknown";
        }, AWTAutoCollection._getDelimiter = function(versionString) {
            return versionString.indexOf(".") > -1 ? "." : versionString.indexOf("_") > -1 ? "_" : null;
        }, AWTAutoCollection._saveData = function(name, value) {
            if (this._propertyStorage) this._propertyStorage.setProperty(name, value); else if (this._areCookiesAvailable) {
                var date = new Date;
                date.setTime(date.getTime() + 31536e6);
                var expires = "expires=" + date.toUTCString();
                document.cookie = name + "=" + value + "; " + expires;
            }
        }, AWTAutoCollection._getData = function(name) {
            if (this._propertyStorage) return this._propertyStorage.getProperty(name) || "";
            if (this._areCookiesAvailable) {
                name += "=";
                for (var ca = document.cookie.split(";"), i = 0; i < ca.length; i++) {
                    for (var c = ca[i], j = 0; " " === c.charAt(j); ) j++;
                    if (0 === (c = c.substring(j)).indexOf(name)) return c.substring(name.length, c.length);
                }
            }
            return "";
        }, AWTAutoCollection._deleteCookie = function(name) {
            this._areCookiesAvailable && (document.cookie = name + "=;expires=Thu, 01 Jan 1970 00:00:01 GMT;");
        }, AWTAutoCollection._disableCookies = !1, AWTAutoCollection._areCookiesAvailable = "undefined" != typeof document && "undefined" != typeof document.cookie, 
        AWTAutoCollection;
    }();
    exports.default = AWTAutoCollection;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var Enums_1 = __webpack_require__(0), AWTHttpManager_1 = __webpack_require__(19), AWTTransmissionManagerCore_1 = __webpack_require__(4), AWTRecordBatcher_1 = __webpack_require__(29), AWTNotificationManager_1 = __webpack_require__(5), Utils = __webpack_require__(3), AWTQueueManager = function() {
        function AWTQueueManager(collectorUrl, _queueSizeLimit, xhrOverride, clockSkewRefreshDurationInMins) {
            this._queueSizeLimit = _queueSizeLimit, this._isCurrentlyUploadingNow = !1, this._uploadNowQueue = [], 
            this._shouldDropEventsOnPause = !1, this._paused = !1, this._queueSize = 0, this._outboundQueue = [], 
            this._inboundQueues = {}, this._inboundQueues[Enums_1.AWTEventPriority.High] = [], 
            this._inboundQueues[Enums_1.AWTEventPriority.Normal] = [], this._inboundQueues[Enums_1.AWTEventPriority.Low] = [], 
            this._addEmptyQueues(), this._batcher = new AWTRecordBatcher_1.default(this._outboundQueue, 500), 
            this._httpManager = new AWTHttpManager_1.default(this._outboundQueue, collectorUrl, this, xhrOverride, clockSkewRefreshDurationInMins);
        }
        return AWTQueueManager.prototype.addEvent = function(event) {
            Utils.isPriority(event.priority) || (event.priority = Enums_1.AWTEventPriority.Normal), 
            event.priority === Enums_1.AWTEventPriority.Immediate_sync ? this._httpManager.sendSynchronousRequest(this._batcher.addEventToBatch(event), event.apiKey) : this._queueSize < this._queueSizeLimit || this._dropEventWithPriorityOrLess(event.priority) ? this._addEventToProperQueue(event) : AWTNotificationManager_1.default.eventsDropped([ event ], Enums_1.AWTEventsDroppedReason.QueueFull);
        }, AWTQueueManager.prototype.sendEventsForPriorityAndAbove = function(priority) {
            this._batchEvents(priority), this._httpManager.sendQueuedRequests();
        }, AWTQueueManager.prototype.hasEvents = function() {
            return (this._inboundQueues[Enums_1.AWTEventPriority.High][0].length > 0 || this._inboundQueues[Enums_1.AWTEventPriority.Normal][0].length > 0 || this._inboundQueues[Enums_1.AWTEventPriority.Low][0].length > 0 || this._batcher.hasBatch()) && this._httpManager.hasIdleConnection();
        }, AWTQueueManager.prototype.addBackRequest = function(request) {
            if (!this._paused || !this._shouldDropEventsOnPause) {
                for (var token in request) if (request.hasOwnProperty(token)) for (var i = 0; i < request[token].length; ++i) request[token][i].sendAttempt < 6 ? this.addEvent(request[token][i]) : AWTNotificationManager_1.default.eventsDropped([ request[token][i] ], Enums_1.AWTEventsDroppedReason.MaxRetryLimit);
                AWTTransmissionManagerCore_1.default.scheduleTimer();
            }
        }, AWTQueueManager.prototype.teardown = function() {
            this._paused || (this._batchEvents(Enums_1.AWTEventPriority.Low), this._httpManager.teardown());
        }, AWTQueueManager.prototype.uploadNow = function(callback) {
            var _this = this;
            this._addEmptyQueues(), this._isCurrentlyUploadingNow ? this._uploadNowQueue.push(callback) : (this._isCurrentlyUploadingNow = !0, 
            setTimeout((function() {
                return _this._uploadNow(callback);
            }), 0));
        }, AWTQueueManager.prototype.pauseTransmission = function() {
            this._paused = !0, this._httpManager.pause(), this._shouldDropEventsOnPause && (this._queueSize -= this._inboundQueues[Enums_1.AWTEventPriority.High][0].length + this._inboundQueues[Enums_1.AWTEventPriority.Normal][0].length + this._inboundQueues[Enums_1.AWTEventPriority.Low][0].length, 
            this._inboundQueues[Enums_1.AWTEventPriority.High][0] = [], this._inboundQueues[Enums_1.AWTEventPriority.Normal][0] = [], 
            this._inboundQueues[Enums_1.AWTEventPriority.Low][0] = [], this._httpManager.removeQueuedRequests());
        }, AWTQueueManager.prototype.resumeTransmission = function() {
            this._paused = !1, this._httpManager.resume();
        }, AWTQueueManager.prototype.shouldDropEventsOnPause = function(shouldDropEventsOnPause) {
            this._shouldDropEventsOnPause = shouldDropEventsOnPause;
        }, AWTQueueManager.prototype._removeFirstQueues = function() {
            this._inboundQueues[Enums_1.AWTEventPriority.High].shift(), this._inboundQueues[Enums_1.AWTEventPriority.Normal].shift(), 
            this._inboundQueues[Enums_1.AWTEventPriority.Low].shift();
        }, AWTQueueManager.prototype._addEmptyQueues = function() {
            this._inboundQueues[Enums_1.AWTEventPriority.High].push([]), this._inboundQueues[Enums_1.AWTEventPriority.Normal].push([]), 
            this._inboundQueues[Enums_1.AWTEventPriority.Low].push([]);
        }, AWTQueueManager.prototype._addEventToProperQueue = function(event) {
            this._paused && this._shouldDropEventsOnPause || (this._queueSize++, this._inboundQueues[event.priority][this._inboundQueues[event.priority].length - 1].push(event));
        }, AWTQueueManager.prototype._dropEventWithPriorityOrLess = function(priority) {
            for (var currentPriority = Enums_1.AWTEventPriority.Low; currentPriority <= priority; ) {
                if (this._inboundQueues[currentPriority][this._inboundQueues[currentPriority].length - 1].length > 0) return AWTNotificationManager_1.default.eventsDropped([ this._inboundQueues[currentPriority][this._inboundQueues[currentPriority].length - 1].shift() ], Enums_1.AWTEventsDroppedReason.QueueFull), 
                !0;
                currentPriority++;
            }
            return !1;
        }, AWTQueueManager.prototype._batchEvents = function(priority) {
            for (var priorityToProcess = Enums_1.AWTEventPriority.High; priorityToProcess >= priority; ) {
                for (;this._inboundQueues[priorityToProcess][0].length > 0; ) {
                    var event_1 = this._inboundQueues[priorityToProcess][0].pop();
                    this._queueSize--, this._batcher.addEventToBatch(event_1);
                }
                priorityToProcess--;
            }
            this._batcher.flushBatch();
        }, AWTQueueManager.prototype._uploadNow = function(callback) {
            var _this = this;
            this.hasEvents() && this.sendEventsForPriorityAndAbove(Enums_1.AWTEventPriority.Low), 
            this._checkOutboundQueueEmptyAndSent((function() {
                _this._removeFirstQueues(), null != callback && callback(), _this._uploadNowQueue.length > 0 ? setTimeout((function() {
                    return _this._uploadNow(_this._uploadNowQueue.shift());
                }), 0) : (_this._isCurrentlyUploadingNow = !1, _this.hasEvents() && AWTTransmissionManagerCore_1.default.scheduleTimer());
            }));
        }, AWTQueueManager.prototype._checkOutboundQueueEmptyAndSent = function(callback) {
            var _this = this;
            this._httpManager.isCompletelyIdle() ? callback() : setTimeout((function() {
                return _this._checkOutboundQueueEmptyAndSent(callback);
            }), 250);
        }, AWTQueueManager;
    }();
    exports.default = AWTQueueManager;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    }), exports.AWT_REAL_TIME = "REAL_TIME", exports.AWT_NEAR_REAL_TIME = "NEAR_REAL_TIME", 
    exports.AWT_BEST_EFFORT = "BEST_EFFORT";
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var Enums_1 = __webpack_require__(0), Enums_2 = __webpack_require__(6), AWTEventProperties_1 = __webpack_require__(7), Utils = __webpack_require__(3), AWTStatsManager_1 = __webpack_require__(14), AWTNotificationManager_1 = __webpack_require__(5), AWTTransmissionManagerCore_1 = __webpack_require__(4), AWTLogManagerSettings_1 = __webpack_require__(18), Version = __webpack_require__(16), AWTSemanticContext_1 = __webpack_require__(9), AWTAutoCollection_1 = __webpack_require__(10), AWTLogger = function() {
        function AWTLogger(_apiKey) {
            this._apiKey = _apiKey, this._contextProperties = new AWTEventProperties_1.default, 
            this._semanticContext = new AWTSemanticContext_1.default(!1, this._contextProperties), 
            this._sessionStartTime = 0, this._createInitId();
        }
        return AWTLogger.prototype.setContext = function(name, value, type) {
            void 0 === type && (type = Enums_1.AWTPropertyType.Unspecified), this._contextProperties.setProperty(name, value, type);
        }, AWTLogger.prototype.setContextWithPii = function(name, value, pii, type) {
            void 0 === type && (type = Enums_1.AWTPropertyType.Unspecified), this._contextProperties.setPropertyWithPii(name, value, pii, type);
        }, AWTLogger.prototype.setContextWithCustomerContent = function(name, value, customerContent, type) {
            void 0 === type && (type = Enums_1.AWTPropertyType.Unspecified), this._contextProperties.setPropertyWithCustomerContent(name, value, customerContent, type);
        }, AWTLogger.prototype.getSemanticContext = function() {
            return this._semanticContext;
        }, AWTLogger.prototype.logEvent = function(event) {
            if (AWTLogManagerSettings_1.default.loggingEnabled) {
                this._apiKey || (this._apiKey = AWTLogManagerSettings_1.default.defaultTenantToken, 
                this._createInitId());
                var sanitizeProperties = !0;
                Utils.isString(event) ? event = {
                    name: event
                } : event instanceof AWTEventProperties_1.default && (event = event.getEvent(), 
                sanitizeProperties = !1), AWTStatsManager_1.default.eventReceived(this._apiKey), 
                AWTLogger._logEvent(AWTLogger._getInternalEvent(event, this._apiKey, sanitizeProperties), this._contextProperties);
            }
        }, AWTLogger.prototype.logSession = function(state, properties) {
            if (AWTLogManagerSettings_1.default.sessionEnabled) {
                var sessionEvent = {
                    name: "session",
                    type: "session",
                    properties: {}
                };
                if (AWTLogger._addPropertiesToEvent(sessionEvent, properties), sessionEvent.priority = Enums_1.AWTEventPriority.High, 
                state === Enums_2.AWTSessionState.Started) {
                    if (this._sessionStartTime > 0) return;
                    this._sessionStartTime = (new Date).getTime(), this._sessionId = Utils.newGuid(), 
                    this.setContext("Session.Id", this._sessionId), sessionEvent.properties["Session.State"] = "Started";
                } else {
                    if (state !== Enums_2.AWTSessionState.Ended) return;
                    if (0 === this._sessionStartTime) return;
                    var sessionDurationSec = Math.floor(((new Date).getTime() - this._sessionStartTime) / 1e3);
                    sessionEvent.properties["Session.Id"] = this._sessionId, sessionEvent.properties["Session.State"] = "Ended", 
                    sessionEvent.properties["Session.Duration"] = sessionDurationSec.toString(), sessionEvent.properties["Session.DurationBucket"] = AWTLogger._getSessionDurationFromTime(sessionDurationSec), 
                    this._sessionStartTime = 0, this.setContext("Session.Id", null), this._sessionId = void 0;
                }
                sessionEvent.properties["Session.FirstLaunchTime"] = AWTAutoCollection_1.default.firstLaunchTime, 
                this.logEvent(sessionEvent);
            }
        }, AWTLogger.prototype.getSessionId = function() {
            return this._sessionId;
        }, AWTLogger.prototype.logFailure = function(signature, detail, category, id, properties) {
            if (signature && detail) {
                var failureEvent = {
                    name: "failure",
                    type: "failure",
                    properties: {}
                };
                AWTLogger._addPropertiesToEvent(failureEvent, properties), failureEvent.properties["Failure.Signature"] = signature, 
                failureEvent.properties["Failure.Detail"] = detail, category && (failureEvent.properties["Failure.Category"] = category), 
                id && (failureEvent.properties["Failure.Id"] = id), failureEvent.priority = Enums_1.AWTEventPriority.High, 
                this.logEvent(failureEvent);
            }
        }, AWTLogger.prototype.logPageView = function(id, pageName, category, uri, referrerUri, properties) {
            if (id && pageName) {
                var pageViewEvent = {
                    name: "pageview",
                    type: "pageview",
                    properties: {}
                };
                AWTLogger._addPropertiesToEvent(pageViewEvent, properties), pageViewEvent.properties["PageView.Id"] = id, 
                pageViewEvent.properties["PageView.Name"] = pageName, category && (pageViewEvent.properties["PageView.Category"] = category), 
                uri && (pageViewEvent.properties["PageView.Uri"] = uri), referrerUri && (pageViewEvent.properties["PageView.ReferrerUri"] = referrerUri), 
                this.logEvent(pageViewEvent);
            }
        }, AWTLogger.prototype._createInitId = function() {
            !AWTLogger._initIdMap[this._apiKey] && this._apiKey && (AWTLogger._initIdMap[this._apiKey] = Utils.newGuid());
        }, AWTLogger._addPropertiesToEvent = function(event, propertiesEvent) {
            if (propertiesEvent) for (var name_1 in propertiesEvent instanceof AWTEventProperties_1.default && (propertiesEvent = propertiesEvent.getEvent()), 
            propertiesEvent.name && (event.name = propertiesEvent.name), propertiesEvent.priority && (event.priority = propertiesEvent.priority), 
            propertiesEvent.properties) propertiesEvent.properties.hasOwnProperty(name_1) && (event.properties[name_1] = propertiesEvent.properties[name_1]);
        }, AWTLogger._getSessionDurationFromTime = function(timeInSec) {
            return timeInSec < 0 ? "Undefined" : timeInSec <= 3 ? "UpTo3Sec" : timeInSec <= 10 ? "UpTo10Sec" : timeInSec <= 30 ? "UpTo30Sec" : timeInSec <= 60 ? "UpTo60Sec" : timeInSec <= 180 ? "UpTo3Min" : timeInSec <= 600 ? "UpTo10Min" : timeInSec <= 1800 ? "UpTo30Min" : "Above30Min";
        }, AWTLogger._logEvent = function(eventWithMetaData, contextProperties) {
            eventWithMetaData.name && Utils.isString(eventWithMetaData.name) ? (eventWithMetaData.name = eventWithMetaData.name.toLowerCase(), 
            eventWithMetaData.name = eventWithMetaData.name.replace(Utils.EventNameDotRegex, "_"), 
            eventWithMetaData.type && Utils.isString(eventWithMetaData.type) ? eventWithMetaData.type = eventWithMetaData.type.toLowerCase() : eventWithMetaData.type = "custom", 
            Utils.EventNameAndTypeRegex.test(eventWithMetaData.name) && Utils.EventNameAndTypeRegex.test(eventWithMetaData.type) ? ((!Utils.isNumber(eventWithMetaData.timestamp) || eventWithMetaData.timestamp < 0) && (eventWithMetaData.timestamp = (new Date).getTime()), 
            eventWithMetaData.properties || (eventWithMetaData.properties = {}), this._addContextIfAbsent(eventWithMetaData, contextProperties.getPropertyMap()), 
            this._addContextIfAbsent(eventWithMetaData, AWTLogManagerSettings_1.default.logManagerContext.getPropertyMap()), 
            this._setDefaultProperty(eventWithMetaData, "EventInfo.InitId", this._getInitId(eventWithMetaData.apiKey)), 
            this._setDefaultProperty(eventWithMetaData, "EventInfo.Sequence", this._getSequenceId(eventWithMetaData.apiKey)), 
            this._setDefaultProperty(eventWithMetaData, "EventInfo.SdkVersion", Version.FullVersionString), 
            this._setDefaultProperty(eventWithMetaData, "EventInfo.Name", eventWithMetaData.name), 
            this._setDefaultProperty(eventWithMetaData, "EventInfo.Time", new Date(eventWithMetaData.timestamp).toISOString()), 
            Utils.isPriority(eventWithMetaData.priority) || (eventWithMetaData.priority = Enums_1.AWTEventPriority.Normal), 
            this._sendEvent(eventWithMetaData)) : AWTNotificationManager_1.default.eventsRejected([ eventWithMetaData ], Enums_1.AWTEventsRejectedReason.InvalidEvent)) : AWTNotificationManager_1.default.eventsRejected([ eventWithMetaData ], Enums_1.AWTEventsRejectedReason.InvalidEvent);
        }, AWTLogger._addContextIfAbsent = function(event, contextProperties) {
            if (contextProperties) for (var name_2 in contextProperties) contextProperties.hasOwnProperty(name_2) && (event.properties[name_2] || (event.properties[name_2] = contextProperties[name_2]));
        }, AWTLogger._setDefaultProperty = function(event, name, value) {
            event.properties[name] = {
                value: value,
                pii: Enums_1.AWTPiiKind.NotSet,
                type: Enums_1.AWTPropertyType.String
            };
        }, AWTLogger._sendEvent = function(event) {
            AWTTransmissionManagerCore_1.default.sendEvent(event);
        }, AWTLogger._getInternalEvent = function(event, apiKey, sanitizeProperties) {
            if (event.properties = event.properties || {}, sanitizeProperties) for (var name_3 in event.properties) event.properties.hasOwnProperty(name_3) && (event.properties[name_3] = Utils.sanitizeProperty(name_3, event.properties[name_3]), 
            null === event.properties[name_3] && delete event.properties[name_3]);
            var internalEvent = event;
            return internalEvent.id = Utils.newGuid(), internalEvent.apiKey = apiKey, internalEvent;
        }, AWTLogger._getInitId = function(apiKey) {
            return AWTLogger._initIdMap[apiKey];
        }, AWTLogger._getSequenceId = function(apiKey) {
            return void 0 === AWTLogger._sequenceIdMap[apiKey] && (AWTLogger._sequenceIdMap[apiKey] = 0), 
            (++AWTLogger._sequenceIdMap[apiKey]).toString();
        }, AWTLogger._sequenceIdMap = {}, AWTLogger._initIdMap = {}, AWTLogger;
    }();
    exports.default = AWTLogger;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var Utils = __webpack_require__(3), AWTNotificationManager_1 = __webpack_require__(5), Enums_1 = __webpack_require__(0), AWTStatsManager = function() {
        function AWTStatsManager() {}
        return AWTStatsManager.initialize = function(sendStats) {
            var _this = this;
            this._sendStats = sendStats, this._isInitalized = !0, AWTNotificationManager_1.default.addNotificationListener({
                eventsSent: function(events) {
                    _this._addStat("records_sent_count", events.length, events[0].apiKey);
                },
                eventsDropped: function(events, reason) {
                    switch (reason) {
                      case Enums_1.AWTEventsDroppedReason.NonRetryableStatus:
                        _this._addStat("d_send_fail", events.length, events[0].apiKey), _this._addStat("records_dropped_count", events.length, events[0].apiKey);
                        break;

                      case Enums_1.AWTEventsDroppedReason.MaxRetryLimit:
                        _this._addStat("d_retry_limit", events.length, events[0].apiKey), _this._addStat("records_dropped_count", events.length, events[0].apiKey);
                        break;

                      case Enums_1.AWTEventsDroppedReason.QueueFull:
                        _this._addStat("d_queue_full", events.length, events[0].apiKey);
                    }
                },
                eventsRejected: function(events, reason) {
                    switch (reason) {
                      case Enums_1.AWTEventsRejectedReason.InvalidEvent:
                        _this._addStat("r_inv", events.length, events[0].apiKey);
                        break;

                      case Enums_1.AWTEventsRejectedReason.KillSwitch:
                        _this._addStat("r_kl", events.length, events[0].apiKey);
                        break;

                      case Enums_1.AWTEventsRejectedReason.SizeLimitExceeded:
                        _this._addStat("r_size", events.length, events[0].apiKey);
                    }
                    _this._addStat("r_count", events.length, events[0].apiKey);
                },
                eventsRetrying: null
            }), setTimeout((function() {
                return _this.flush();
            }), 6e4);
        }, AWTStatsManager.teardown = function() {
            this._isInitalized && (this.flush(), this._isInitalized = !1);
        }, AWTStatsManager.eventReceived = function(apiKey) {
            AWTStatsManager._addStat("records_received_count", 1, apiKey);
        }, AWTStatsManager.flush = function() {
            var _this = this;
            if (this._isInitalized) {
                for (var tenantId in this._stats) this._stats.hasOwnProperty(tenantId) && this._sendStats(this._stats[tenantId], tenantId);
                this._stats = {}, setTimeout((function() {
                    return _this.flush();
                }), 6e4);
            }
        }, AWTStatsManager._addStat = function(statName, value, apiKey) {
            if (this._isInitalized && apiKey !== Utils.StatsApiKey) {
                var tenantId = Utils.getTenantId(apiKey);
                this._stats[tenantId] || (this._stats[tenantId] = {}), this._stats[tenantId][statName] ? this._stats[tenantId][statName] = this._stats[tenantId][statName] + value : this._stats[tenantId][statName] = value;
            }
        }, AWTStatsManager._isInitalized = !1, AWTStatsManager._stats = {}, AWTStatsManager;
    }();
    exports.default = AWTStatsManager;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var Bond = __webpack_require__(20), Enums_1 = __webpack_require__(0), AWTNotificationManager_1 = __webpack_require__(5), Utils = __webpack_require__(3), AWTSerializer = function() {
        function AWTSerializer() {}
        return AWTSerializer.getPayloadBlob = function(requestDictionary, tokenCount) {
            var remainingRequest, requestFull = !1, stream = new Bond.IO.MemoryStream, writer = new Bond.CompactBinaryProtocolWriter(stream);
            for (var token in writer._WriteFieldBegin(Bond._BondDataType._BT_MAP, 3, null), 
            writer._WriteMapContainerBegin(tokenCount, Bond._BondDataType._BT_STRING, Bond._BondDataType._BT_LIST), 
            requestDictionary) if (requestFull) remainingRequest || (remainingRequest = {}), 
            remainingRequest[token] = requestDictionary[token], delete requestDictionary[token]; else if (requestDictionary.hasOwnProperty(token)) {
                writer._WriteString(token);
                var dataPackage = requestDictionary[token];
                writer._WriteContainerBegin(1, Bond._BondDataType._BT_STRUCT), writer._WriteFieldBegin(Bond._BondDataType._BT_STRING, 2, null), 
                writer._WriteString("act_default_source"), writer._WriteFieldBegin(Bond._BondDataType._BT_STRING, 5, null), 
                writer._WriteString(Utils.newGuid()), writer._WriteFieldBegin(Bond._BondDataType._BT_INT64, 6, null), 
                writer._WriteInt64(Utils.numberToBondInt64(Date.now())), writer._WriteFieldBegin(Bond._BondDataType._BT_LIST, 8, null);
                var dpSizePos = stream._GetBuffer().length + 1;
                writer._WriteContainerBegin(requestDictionary[token].length, Bond._BondDataType._BT_STRUCT);
                for (var dpSizeSerialized = stream._GetBuffer().length - dpSizePos, i = 0; i < dataPackage.length; ++i) {
                    var currentStreamPos = stream._GetBuffer().length;
                    if (this.writeEvent(dataPackage[i], writer), stream._GetBuffer().length - currentStreamPos > 2936012) AWTNotificationManager_1.default.eventsRejected([ dataPackage[i] ], Enums_1.AWTEventsRejectedReason.SizeLimitExceeded), 
                    dataPackage.splice(i--, 1), stream._GetBuffer().splice(currentStreamPos), this._addNewDataPackageSize(dataPackage.length, stream, dpSizeSerialized, dpSizePos); else if (stream._GetBuffer().length > 2936012) {
                        stream._GetBuffer().splice(currentStreamPos), remainingRequest || (remainingRequest = {}), 
                        requestDictionary[token] = dataPackage.splice(0, i), remainingRequest[token] = dataPackage, 
                        this._addNewDataPackageSize(requestDictionary[token].length, stream, dpSizeSerialized, dpSizePos), 
                        requestFull = !0;
                        break;
                    }
                }
                writer._WriteStructEnd(!1);
            }
            return writer._WriteStructEnd(!1), {
                payloadBlob: stream._GetBuffer(),
                remainingRequest: remainingRequest
            };
        }, AWTSerializer._addNewDataPackageSize = function(size, stream, oldDpSize, streamPos) {
            for (var newRecordCountSerialized = Bond._Encoding._Varint_GetBytes(Bond.Number._ToUInt32(size)), j = 0; j < oldDpSize; ++j) {
                if (!(j < newRecordCountSerialized.length)) {
                    stream._GetBuffer().slice(streamPos + j, oldDpSize - j);
                    break;
                }
                stream._GetBuffer()[streamPos + j] = newRecordCountSerialized[j];
            }
        }, AWTSerializer.writeEvent = function(eventData, writer) {
            writer._WriteFieldBegin(Bond._BondDataType._BT_STRING, 1, null), writer._WriteString(eventData.id), 
            writer._WriteFieldBegin(Bond._BondDataType._BT_INT64, 3, null), writer._WriteInt64(Utils.numberToBondInt64(eventData.timestamp)), 
            writer._WriteFieldBegin(Bond._BondDataType._BT_STRING, 5, null), writer._WriteString(eventData.type), 
            writer._WriteFieldBegin(Bond._BondDataType._BT_STRING, 6, null), writer._WriteString(eventData.name);
            var propsString = {}, propStringCount = 0, propsInt64 = {}, propInt64Count = 0, propsDouble = {}, propDoubleCount = 0, propsBool = {}, propBoolCount = 0, propsDate = {}, propDateCount = 0, piiProps = {}, piiPropCount = 0, ccProps = {}, ccPropCount = 0;
            for (var key in eventData.properties) {
                if (eventData.properties.hasOwnProperty(key)) if ((property = eventData.properties[key]).cc > 0) ccProps[key] = property, 
                ccPropCount++; else if (property.pii > 0) piiProps[key] = property, piiPropCount++; else switch (property.type) {
                  case Enums_1.AWTPropertyType.String:
                    propsString[key] = property.value, propStringCount++;
                    break;

                  case Enums_1.AWTPropertyType.Int64:
                    propsInt64[key] = property.value, propInt64Count++;
                    break;

                  case Enums_1.AWTPropertyType.Double:
                    propsDouble[key] = property.value, propDoubleCount++;
                    break;

                  case Enums_1.AWTPropertyType.Boolean:
                    propsBool[key] = property.value, propBoolCount++;
                    break;

                  case Enums_1.AWTPropertyType.Date:
                    propsDate[key] = property.value, propDateCount++;
                }
            }
            if (propStringCount) for (var key in writer._WriteFieldBegin(Bond._BondDataType._BT_MAP, 13, null), 
            writer._WriteMapContainerBegin(propStringCount, Bond._BondDataType._BT_STRING, Bond._BondDataType._BT_STRING), 
            propsString) if (propsString.hasOwnProperty(key)) {
                var value = propsString[key];
                writer._WriteString(key), writer._WriteString(value.toString());
            }
            if (piiPropCount) for (var key in writer._WriteFieldBegin(Bond._BondDataType._BT_MAP, 30, null), 
            writer._WriteMapContainerBegin(piiPropCount, Bond._BondDataType._BT_STRING, Bond._BondDataType._BT_STRUCT), 
            piiProps) if (piiProps.hasOwnProperty(key)) {
                var property = piiProps[key];
                writer._WriteString(key), writer._WriteFieldBegin(Bond._BondDataType._BT_INT32, 1, null), 
                writer._WriteInt32(1), writer._WriteFieldBegin(Bond._BondDataType._BT_INT32, 2, null), 
                writer._WriteInt32(property.pii), writer._WriteFieldBegin(Bond._BondDataType._BT_STRING, 3, null), 
                writer._WriteString(property.value.toString()), writer._WriteStructEnd(!1);
            }
            if (propBoolCount) for (var key in writer._WriteFieldBegin(Bond._BondDataType._BT_MAP, 31, null), 
            writer._WriteMapContainerBegin(propBoolCount, Bond._BondDataType._BT_STRING, Bond._BondDataType._BT_BOOL), 
            propsBool) if (propsBool.hasOwnProperty(key)) {
                value = propsBool[key];
                writer._WriteString(key), writer._WriteBool(value);
            }
            if (propDateCount) for (var key in writer._WriteFieldBegin(Bond._BondDataType._BT_MAP, 32, null), 
            writer._WriteMapContainerBegin(propDateCount, Bond._BondDataType._BT_STRING, Bond._BondDataType._BT_INT64), 
            propsDate) if (propsDate.hasOwnProperty(key)) {
                value = propsDate[key];
                writer._WriteString(key), writer._WriteInt64(Utils.numberToBondInt64(value));
            }
            if (propInt64Count) for (var key in writer._WriteFieldBegin(Bond._BondDataType._BT_MAP, 33, null), 
            writer._WriteMapContainerBegin(propInt64Count, Bond._BondDataType._BT_STRING, Bond._BondDataType._BT_INT64), 
            propsInt64) if (propsInt64.hasOwnProperty(key)) {
                value = propsInt64[key];
                writer._WriteString(key), writer._WriteInt64(Utils.numberToBondInt64(value));
            }
            if (propDoubleCount) for (var key in writer._WriteFieldBegin(Bond._BondDataType._BT_MAP, 34, null), 
            writer._WriteMapContainerBegin(propDoubleCount, Bond._BondDataType._BT_STRING, Bond._BondDataType._BT_DOUBLE), 
            propsDouble) if (propsDouble.hasOwnProperty(key)) {
                value = propsDouble[key];
                writer._WriteString(key), writer._WriteDouble(value);
            }
            if (ccPropCount) for (var key in writer._WriteFieldBegin(Bond._BondDataType._BT_MAP, 36, null), 
            writer._WriteMapContainerBegin(ccPropCount, Bond._BondDataType._BT_STRING, Bond._BondDataType._BT_STRUCT), 
            ccProps) if (ccProps.hasOwnProperty(key)) {
                property = ccProps[key];
                writer._WriteString(key), writer._WriteFieldBegin(Bond._BondDataType._BT_INT32, 1, null), 
                writer._WriteInt32(property.cc), writer._WriteFieldBegin(Bond._BondDataType._BT_STRING, 2, null), 
                writer._WriteString(property.value.toString()), writer._WriteStructEnd(!1);
            }
            writer._WriteStructEnd(!1);
        }, AWTSerializer.base64Encode = function(data) {
            return Bond._Encoding._Base64_GetString(data);
        }, AWTSerializer;
    }();
    exports.default = AWTSerializer;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    }), exports.Version = "1.8.7", exports.FullVersionString = "AWT-Web-JS-" + exports.Version;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var Enums_1 = __webpack_require__(0), Enums_2 = __webpack_require__(6), AWTLogManagerSettings_1 = __webpack_require__(18), AWTLogger_1 = __webpack_require__(13), AWTTransmissionManagerCore_1 = __webpack_require__(4), AWTNotificationManager_1 = __webpack_require__(5), AWTAutoCollection_1 = __webpack_require__(10), AWTLogManager = function() {
        function AWTLogManager() {}
        return AWTLogManager.initialize = function(tenantToken, configuration) {
            if (void 0 === configuration && (configuration = {}), !this._isInitialized) return this._isInitialized = !0, 
            AWTLogManagerSettings_1.default.defaultTenantToken = tenantToken, this._overrideValuesFromConfig(configuration), 
            this._config.disableCookiesUsage && !this._config.propertyStorageOverride && (AWTLogManagerSettings_1.default.sessionEnabled = !1), 
            AWTAutoCollection_1.default.addPropertyStorageOverride(this._config.propertyStorageOverride), 
            AWTAutoCollection_1.default.autoCollect(AWTLogManagerSettings_1.default.semanticContext, this._config.disableCookiesUsage, this._config.userAgent), 
            AWTTransmissionManagerCore_1.default.initialize(this._config), AWTLogManagerSettings_1.default.loggingEnabled = !0, 
            this._config.enableAutoUserSession && (this.getLogger().logSession(Enums_2.AWTSessionState.Started), 
            window.addEventListener("beforeunload", this.flushAndTeardown)), this.getLogger();
        }, AWTLogManager.getSemanticContext = function() {
            return AWTLogManagerSettings_1.default.semanticContext;
        }, AWTLogManager.flush = function(callback) {
            this._isInitialized && !this._isDestroyed && AWTTransmissionManagerCore_1.default.flush(callback);
        }, AWTLogManager.flushAndTeardown = function() {
            this._isInitialized && !this._isDestroyed && (this._config.enableAutoUserSession && this.getLogger().logSession(Enums_2.AWTSessionState.Ended), 
            AWTTransmissionManagerCore_1.default.flushAndTeardown(), AWTLogManagerSettings_1.default.loggingEnabled = !1, 
            this._isDestroyed = !0);
        }, AWTLogManager.pauseTransmission = function() {
            this._isInitialized && !this._isDestroyed && AWTTransmissionManagerCore_1.default.pauseTransmission();
        }, AWTLogManager.resumeTransmision = function() {
            this._isInitialized && !this._isDestroyed && AWTTransmissionManagerCore_1.default.resumeTransmision();
        }, AWTLogManager.setTransmitProfile = function(profileName) {
            this._isInitialized && !this._isDestroyed && AWTTransmissionManagerCore_1.default.setTransmitProfile(profileName);
        }, AWTLogManager.loadTransmitProfiles = function(profiles) {
            this._isInitialized && !this._isDestroyed && AWTTransmissionManagerCore_1.default.loadTransmitProfiles(profiles);
        }, AWTLogManager.setContext = function(name, value, type) {
            void 0 === type && (type = Enums_1.AWTPropertyType.Unspecified), AWTLogManagerSettings_1.default.logManagerContext.setProperty(name, value, type);
        }, AWTLogManager.setContextWithPii = function(name, value, pii, type) {
            void 0 === type && (type = Enums_1.AWTPropertyType.Unspecified), AWTLogManagerSettings_1.default.logManagerContext.setPropertyWithPii(name, value, pii, type);
        }, AWTLogManager.setContextWithCustomerContent = function(name, value, customerContent, type) {
            void 0 === type && (type = Enums_1.AWTPropertyType.Unspecified), AWTLogManagerSettings_1.default.logManagerContext.setPropertyWithCustomerContent(name, value, customerContent, type);
        }, AWTLogManager.getLogger = function(tenantToken) {
            var key = tenantToken;
            return key && key !== AWTLogManagerSettings_1.default.defaultTenantToken || (key = ""), 
            this._loggers[key] || (this._loggers[key] = new AWTLogger_1.default(key)), this._loggers[key];
        }, AWTLogManager.addNotificationListener = function(listener) {
            AWTNotificationManager_1.default.addNotificationListener(listener);
        }, AWTLogManager.removeNotificationListener = function(listener) {
            AWTNotificationManager_1.default.removeNotificationListener(listener);
        }, AWTLogManager._overrideValuesFromConfig = function(config) {
            config.collectorUri && (this._config.collectorUri = config.collectorUri), config.cacheMemorySizeLimitInNumberOfEvents > 0 && (this._config.cacheMemorySizeLimitInNumberOfEvents = config.cacheMemorySizeLimitInNumberOfEvents), 
            config.httpXHROverride && config.httpXHROverride.sendPOST && (this._config.httpXHROverride = config.httpXHROverride), 
            config.propertyStorageOverride && config.propertyStorageOverride.getProperty && config.propertyStorageOverride.setProperty && (this._config.propertyStorageOverride = config.propertyStorageOverride), 
            config.userAgent && (this._config.userAgent = config.userAgent), config.disableCookiesUsage && (this._config.disableCookiesUsage = config.disableCookiesUsage), 
            config.canSendStatEvent && (this._config.canSendStatEvent = config.canSendStatEvent), 
            config.enableAutoUserSession && "undefined" != typeof window && window.addEventListener && (this._config.enableAutoUserSession = config.enableAutoUserSession), 
            config.clockSkewRefreshDurationInMins > 0 && (this._config.clockSkewRefreshDurationInMins = config.clockSkewRefreshDurationInMins);
        }, AWTLogManager._loggers = {}, AWTLogManager._isInitialized = !1, AWTLogManager._isDestroyed = !1, 
        AWTLogManager._config = {
            collectorUri: "https://browser.pipe.aria.microsoft.com/Collector/3.0/",
            cacheMemorySizeLimitInNumberOfEvents: 1e4,
            disableCookiesUsage: !1,
            canSendStatEvent: function(eventName) {
                return !0;
            },
            clockSkewRefreshDurationInMins: 0
        }, AWTLogManager;
    }();
    exports.default = AWTLogManager;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var AWTEventProperties_1 = __webpack_require__(7), AWTSemanticContext_1 = __webpack_require__(9), AWTLogManagerSettings = function() {
        function AWTLogManagerSettings() {}
        return AWTLogManagerSettings.logManagerContext = new AWTEventProperties_1.default, 
        AWTLogManagerSettings.sessionEnabled = !0, AWTLogManagerSettings.loggingEnabled = !1, 
        AWTLogManagerSettings.defaultTenantToken = "", AWTLogManagerSettings.semanticContext = new AWTSemanticContext_1.default(!0, AWTLogManagerSettings.logManagerContext), 
        AWTLogManagerSettings;
    }();
    exports.default = AWTLogManagerSettings;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var Enums_1 = __webpack_require__(0), AWTSerializer_1 = __webpack_require__(15), AWTRetryPolicy_1 = __webpack_require__(26), AWTKillSwitch_1 = __webpack_require__(27), AWTClockSkewManager_1 = __webpack_require__(28), Version = __webpack_require__(16), Utils = __webpack_require__(3), AWTNotificationManager_1 = __webpack_require__(5), AWTTransmissionManagerCore_1 = __webpack_require__(4), AWTHttpManager = function() {
        function AWTHttpManager(_requestQueue, collectorUrl, _queueManager, _httpInterface, clockSkewRefreshDurationInMins) {
            var _this = this;
            this._requestQueue = _requestQueue, this._queueManager = _queueManager, this._httpInterface = _httpInterface, 
            this._urlString = "?qsp=true&content-type=application%2Fbond-compact-binary&client-id=NO_AUTH&sdk-version=" + Version.FullVersionString, 
            this._killSwitch = new AWTKillSwitch_1.default, this._paused = !1, this._useBeacons = !1, 
            this._activeConnections = 0, this._clockSkewManager = new AWTClockSkewManager_1.default(clockSkewRefreshDurationInMins), 
            Utils.isUint8ArrayAvailable() || (this._urlString += "&content-encoding=base64"), 
            this._urlString = collectorUrl + this._urlString, this._httpInterface || (this._useBeacons = !Utils.isReactNative(), 
            this._httpInterface = {
                sendPOST: function(urlString, data, ontimeout, onerror, onload, sync) {
                    try {
                        if (Utils.useFetchRequest()) fetch(urlString, {
                            body: data,
                            method: "POST"
                        }).then((function(response) {
                            var headerMap = {};
                            response.headers && response.headers.forEach((function(value, name) {
                                headerMap[name] = value;
                            })), onload(response.status, headerMap);
                        })).catch((function(error) {
                            onerror(0, {});
                        })); else if (Utils.useXDomainRequest()) {
                            var xdr = new XDomainRequest;
                            xdr.open("POST", urlString), xdr.onload = function() {
                                onload(200, null);
                            }, xdr.onerror = function() {
                                onerror(400, null);
                            }, xdr.ontimeout = function() {
                                ontimeout(500, null);
                            }, xdr.send(data);
                        } else {
                            var xhr_1 = new XMLHttpRequest;
                            xhr_1.open("POST", urlString, !sync), xhr_1.onload = function() {
                                onload(xhr_1.status, _this._convertAllHeadersToMap(xhr_1.getAllResponseHeaders()));
                            }, xhr_1.onerror = function() {
                                onerror(xhr_1.status, _this._convertAllHeadersToMap(xhr_1.getAllResponseHeaders()));
                            }, xhr_1.ontimeout = function() {
                                ontimeout(xhr_1.status, _this._convertAllHeadersToMap(xhr_1.getAllResponseHeaders()));
                            }, xhr_1.send(data);
                        }
                    } catch (e) {
                        onerror(400, null);
                    }
                }
            });
        }
        return AWTHttpManager.prototype.hasIdleConnection = function() {
            return this._activeConnections < 2;
        }, AWTHttpManager.prototype.sendQueuedRequests = function() {
            for (;this.hasIdleConnection() && !this._paused && this._requestQueue.length > 0 && this._clockSkewManager.allowRequestSending(); ) this._activeConnections++, 
            this._sendRequest(this._requestQueue.shift(), 0, !1);
            this.hasIdleConnection() && AWTTransmissionManagerCore_1.default.scheduleTimer();
        }, AWTHttpManager.prototype.isCompletelyIdle = function() {
            return 0 === this._activeConnections;
        }, AWTHttpManager.prototype.teardown = function() {
            for (;this._requestQueue.length > 0; ) this._sendRequest(this._requestQueue.shift(), 0, !0);
        }, AWTHttpManager.prototype.pause = function() {
            this._paused = !0;
        }, AWTHttpManager.prototype.resume = function() {
            this._paused = !1, this.sendQueuedRequests();
        }, AWTHttpManager.prototype.removeQueuedRequests = function() {
            this._requestQueue.length = 0;
        }, AWTHttpManager.prototype.sendSynchronousRequest = function(request, token) {
            this._paused && (request[token][0].priority = Enums_1.AWTEventPriority.High), this._activeConnections++, 
            this._sendRequest(request, 0, !1, !0);
        }, AWTHttpManager.prototype._sendRequest = function(request, retryCount, isTeardown, isSynchronous) {
            var _this = this;
            void 0 === isSynchronous && (isSynchronous = !1);
            try {
                if (this._paused) return this._activeConnections--, void this._queueManager.addBackRequest(request);
                var tokenCount_1 = 0, apikey_1 = "";
                for (var token in request) request.hasOwnProperty(token) && (this._killSwitch.isTenantKilled(token) ? (AWTNotificationManager_1.default.eventsRejected(request[token], Enums_1.AWTEventsRejectedReason.KillSwitch), 
                delete request[token]) : (apikey_1.length > 0 && (apikey_1 += ","), apikey_1 += token, 
                tokenCount_1++));
                if (tokenCount_1 > 0) {
                    var payloadResult = AWTSerializer_1.default.getPayloadBlob(request, tokenCount_1);
                    payloadResult.remainingRequest && this._requestQueue.push(payloadResult.remainingRequest);
                    var urlString = this._urlString + "&x-apikey=" + apikey_1 + "&client-time-epoch-millis=" + Date.now().toString();
                    this._clockSkewManager.shouldAddClockSkewHeaders() && (urlString = urlString + "&time-delta-to-apply-millis=" + this._clockSkewManager.getClockSkewHeaderValue());
                    var data = void 0;
                    for (var token in data = Utils.isUint8ArrayAvailable() ? new Uint8Array(payloadResult.payloadBlob) : AWTSerializer_1.default.base64Encode(payloadResult.payloadBlob), 
                    request) if (request.hasOwnProperty(token)) for (var i = 0; i < request[token].length; ++i) request[token][i].sendAttempt > 0 ? request[token][i].sendAttempt++ : request[token][i].sendAttempt = 1;
                    if (this._useBeacons && isTeardown && Utils.isBeaconsSupported() && navigator.sendBeacon(urlString, data)) return;
                    this._httpInterface.sendPOST(urlString, data, (function(status, headers) {
                        _this._retryRequestIfNeeded(status, headers, request, tokenCount_1, apikey_1, retryCount, isTeardown, isSynchronous);
                    }), (function(status, headers) {
                        _this._retryRequestIfNeeded(status, headers, request, tokenCount_1, apikey_1, retryCount, isTeardown, isSynchronous);
                    }), (function(status, headers) {
                        _this._retryRequestIfNeeded(status, headers, request, tokenCount_1, apikey_1, retryCount, isTeardown, isSynchronous);
                    }), isTeardown || isSynchronous);
                } else isTeardown || this._handleRequestFinished(!1, {}, isTeardown, isSynchronous);
            } catch (e) {
                this._handleRequestFinished(!1, {}, isTeardown, isSynchronous);
            }
        }, AWTHttpManager.prototype._retryRequestIfNeeded = function(status, headers, request, tokenCount, apikey, retryCount, isTeardown, isSynchronous) {
            var _this = this, shouldRetry = !0;
            if ("undefined" != typeof status) {
                if (headers) {
                    var killedTokens = this._killSwitch.setKillSwitchTenants(headers["kill-tokens"], headers["kill-duration-seconds"]);
                    this._clockSkewManager.setClockSkew(headers["time-delta-millis"]);
                    for (var i = 0; i < killedTokens.length; ++i) AWTNotificationManager_1.default.eventsRejected(request[killedTokens[i]], Enums_1.AWTEventsRejectedReason.KillSwitch), 
                    delete request[killedTokens[i]], tokenCount--;
                } else this._clockSkewManager.setClockSkew(null);
                if (200 === status) return void this._handleRequestFinished(!0, request, isTeardown, isSynchronous);
                (!AWTRetryPolicy_1.default.shouldRetryForStatus(status) || tokenCount <= 0) && (shouldRetry = !1);
            }
            if (shouldRetry) if (isSynchronous) this._activeConnections--, request[apikey][0].priority = Enums_1.AWTEventPriority.High, 
            this._queueManager.addBackRequest(request); else if (retryCount < 1) {
                for (var token in request) request.hasOwnProperty(token) && AWTNotificationManager_1.default.eventsRetrying(request[token]);
                setTimeout((function() {
                    return _this._sendRequest(request, retryCount + 1, !1);
                }), AWTRetryPolicy_1.default.getMillisToBackoffForRetry(retryCount));
            } else this._activeConnections--, AWTTransmissionManagerCore_1.default.backOffTransmission(), 
            this._queueManager.addBackRequest(request); else this._handleRequestFinished(!1, request, isTeardown, isSynchronous);
        }, AWTHttpManager.prototype._handleRequestFinished = function(success, request, isTeardown, isSynchronous) {
            for (var token in success && AWTTransmissionManagerCore_1.default.clearBackOff(), 
            request) request.hasOwnProperty(token) && (success ? AWTNotificationManager_1.default.eventsSent(request[token]) : AWTNotificationManager_1.default.eventsDropped(request[token], Enums_1.AWTEventsDroppedReason.NonRetryableStatus));
            this._activeConnections--, isSynchronous || isTeardown || this.sendQueuedRequests();
        }, AWTHttpManager.prototype._convertAllHeadersToMap = function(headersString) {
            var headers = {};
            if (headersString) for (var headersArray = headersString.split("\n"), i = 0; i < headersArray.length; ++i) {
                var header = headersArray[i].split(": ");
                headers[header[0]] = header[1];
            }
            return headers;
        }, AWTHttpManager;
    }();
    exports.default = AWTHttpManager;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var bond_const_1 = __webpack_require__(21);
    exports._BondDataType = bond_const_1._BondDataType;
    var _Encoding = __webpack_require__(22);
    exports._Encoding = _Encoding;
    var IO = __webpack_require__(25);
    exports.IO = IO;
    var microsoft_bond_primitives_1 = __webpack_require__(8);
    exports.Int64 = microsoft_bond_primitives_1.Int64, exports.UInt64 = microsoft_bond_primitives_1.UInt64, 
    exports.Number = microsoft_bond_primitives_1.Number;
    var CompactBinaryProtocolWriter = function() {
        function CompactBinaryProtocolWriter(stream) {
            this._stream = stream;
        }
        return CompactBinaryProtocolWriter.prototype._WriteBlob = function(blob) {
            this._stream._Write(blob, 0, blob.length);
        }, CompactBinaryProtocolWriter.prototype._WriteBool = function(value) {
            this._stream._WriteByte(value ? 1 : 0);
        }, CompactBinaryProtocolWriter.prototype._WriteContainerBegin = function(size, elementType) {
            this._WriteUInt8(elementType), this._WriteUInt32(size);
        }, CompactBinaryProtocolWriter.prototype._WriteMapContainerBegin = function(size, keyType, valueType) {
            this._WriteUInt8(keyType), this._WriteUInt8(valueType), this._WriteUInt32(size);
        }, CompactBinaryProtocolWriter.prototype._WriteDouble = function(value) {
            var array = _Encoding._Double_GetBytes(value);
            this._stream._Write(array, 0, array.length);
        }, CompactBinaryProtocolWriter.prototype._WriteFieldBegin = function(type, id, metadata) {
            id <= 5 ? this._stream._WriteByte(type | id << 5) : id <= 255 ? (this._stream._WriteByte(192 | type), 
            this._stream._WriteByte(id)) : (this._stream._WriteByte(224 | type), this._stream._WriteByte(id), 
            this._stream._WriteByte(id >> 8));
        }, CompactBinaryProtocolWriter.prototype._WriteInt32 = function(value) {
            value = _Encoding._Zigzag_EncodeZigzag32(value), this._WriteUInt32(value);
        }, CompactBinaryProtocolWriter.prototype._WriteInt64 = function(value) {
            this._WriteUInt64(_Encoding._Zigzag_EncodeZigzag64(value));
        }, CompactBinaryProtocolWriter.prototype._WriteString = function(value) {
            if ("" === value) this._WriteUInt32(0); else {
                var array = _Encoding._Utf8_GetBytes(value);
                this._WriteUInt32(array.length), this._stream._Write(array, 0, array.length);
            }
        }, CompactBinaryProtocolWriter.prototype._WriteStructEnd = function(isBase) {
            this._WriteUInt8(isBase ? bond_const_1._BondDataType._BT_STOP_BASE : bond_const_1._BondDataType._BT_STOP);
        }, CompactBinaryProtocolWriter.prototype._WriteUInt32 = function(value) {
            var array = _Encoding._Varint_GetBytes(microsoft_bond_primitives_1.Number._ToUInt32(value));
            this._stream._Write(array, 0, array.length);
        }, CompactBinaryProtocolWriter.prototype._WriteUInt64 = function(value) {
            var array = _Encoding._Varint64_GetBytes(value);
            this._stream._Write(array, 0, array.length);
        }, CompactBinaryProtocolWriter.prototype._WriteUInt8 = function(value) {
            this._stream._WriteByte(microsoft_bond_primitives_1.Number._ToUInt8(value));
        }, CompactBinaryProtocolWriter;
    }();
    exports.CompactBinaryProtocolWriter = CompactBinaryProtocolWriter;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    }), function(_BondDataType) {
        _BondDataType[_BondDataType._BT_STOP = 0] = "_BT_STOP", _BondDataType[_BondDataType._BT_STOP_BASE = 1] = "_BT_STOP_BASE", 
        _BondDataType[_BondDataType._BT_BOOL = 2] = "_BT_BOOL", _BondDataType[_BondDataType._BT_DOUBLE = 8] = "_BT_DOUBLE", 
        _BondDataType[_BondDataType._BT_STRING = 9] = "_BT_STRING", _BondDataType[_BondDataType._BT_STRUCT = 10] = "_BT_STRUCT", 
        _BondDataType[_BondDataType._BT_LIST = 11] = "_BT_LIST", _BondDataType[_BondDataType._BT_MAP = 13] = "_BT_MAP", 
        _BondDataType[_BondDataType._BT_INT32 = 16] = "_BT_INT32", _BondDataType[_BondDataType._BT_INT64 = 17] = "_BT_INT64";
    }(exports._BondDataType || (exports._BondDataType = {}));
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var microsoft_bond_primitives_1 = __webpack_require__(8), microsoft_bond_floatutils_1 = __webpack_require__(23), microsoft_bond_utils_1 = __webpack_require__(24);
    exports._Utf8_GetBytes = function(value) {
        for (var array = [], i = 0; i < value.length; ++i) {
            var char = value.charCodeAt(i);
            char < 128 ? array.push(char) : char < 2048 ? array.push(192 | char >> 6, 128 | 63 & char) : char < 55296 || char >= 57344 ? array.push(224 | char >> 12, 128 | char >> 6 & 63, 128 | 63 & char) : (char = 65536 + ((1023 & char) << 10 | 1023 & value.charCodeAt(++i)), 
            array.push(240 | char >> 18, 128 | char >> 12 & 63, 128 | char >> 6 & 63, 128 | 63 & char));
        }
        return array;
    }, exports._Base64_GetString = function(inArray) {
        for (var num, lookup = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", output = [], paddingBytes = inArray.length % 3, i = 0, length_1 = inArray.length - paddingBytes; i < length_1; i += 3) {
            var temp = (inArray[i] << 16) + (inArray[i + 1] << 8) + inArray[i + 2];
            output.push([ lookup.charAt((num = temp) >> 18 & 63), lookup.charAt(num >> 12 & 63), lookup.charAt(num >> 6 & 63), lookup.charAt(63 & num) ].join(""));
        }
        switch (paddingBytes) {
          case 1:
            temp = inArray[inArray.length - 1];
            output.push(lookup.charAt(temp >> 2)), output.push(lookup.charAt(temp << 4 & 63)), 
            output.push("==");
            break;

          case 2:
            var temp2 = (inArray[inArray.length - 2] << 8) + inArray[inArray.length - 1];
            output.push(lookup.charAt(temp2 >> 10)), output.push(lookup.charAt(temp2 >> 4 & 63)), 
            output.push(lookup.charAt(temp2 << 2 & 63)), output.push("=");
        }
        return output.join("");
    }, exports._Varint_GetBytes = function(value) {
        for (var array = []; 4294967168 & value; ) array.push(127 & value | 128), value >>>= 7;
        return array.push(127 & value), array;
    }, exports._Varint64_GetBytes = function(value) {
        for (var low = value.low, high = value.high, array = []; high || 4294967168 & low; ) array.push(127 & low | 128), 
        low = (127 & high) << 25 | low >>> 7, high >>>= 7;
        return array.push(127 & low), array;
    }, exports._Double_GetBytes = function(value) {
        if (microsoft_bond_utils_1.BrowserChecker._IsDataViewSupport()) {
            var view = new DataView(new ArrayBuffer(8));
            view.setFloat64(0, value, !0);
            for (var array = [], i = 0; i < 8; ++i) array.push(view.getUint8(i));
            return array;
        }
        return microsoft_bond_floatutils_1.FloatUtils._ConvertNumberToArray(value, !0);
    }, exports._Zigzag_EncodeZigzag32 = function(value) {
        return (value = microsoft_bond_primitives_1.Number._ToInt32(value)) << 1 ^ value >> 31;
    }, exports._Zigzag_EncodeZigzag64 = function(value) {
        var low = value.low, high = value.high, tmpH = high << 1 | low >>> 31, tmpL = low << 1;
        2147483648 & high && (tmpH = ~tmpH, tmpL = ~tmpL);
        var res = new microsoft_bond_primitives_1.UInt64("0");
        return res.low = tmpL, res.high = tmpH, res;
    };
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var FloatUtils = function() {
        function FloatUtils() {}
        return FloatUtils._ConvertNumberToArray = function(num, isDouble) {
            if (!num) return isDouble ? this._doubleZero : this._floatZero;
            var precisionBits = isDouble ? 52 : 23, bias = (1 << (isDouble ? 11 : 8) - 1) - 1, minExponent = 1 - bias, maxExponent = bias, sign = num < 0 ? 1 : 0;
            num = Math.abs(num);
            for (var intPart = Math.floor(num), floatPart = num - intPart, len = 2 * (bias + 2) + precisionBits, buffer = new Array(len), i = 0; i < len; ) buffer[i++] = 0;
            for (i = bias + 2; i && intPart; ) buffer[--i] = intPart % 2, intPart = Math.floor(intPart / 2);
            for (i = bias + 1; i < len - 1 && floatPart > 0; ) (floatPart *= 2) >= 1 ? (buffer[++i] = 1, 
            --floatPart) : buffer[++i] = 0;
            for (var firstBit = 0; firstBit < len && !buffer[firstBit]; ) firstBit++;
            var exponent = bias + 1 - firstBit, lastBit = firstBit + precisionBits;
            if (buffer[lastBit + 1]) {
                for (i = lastBit; i > firstBit && (buffer[i] = 1 - buffer[i], !buffer); --i) ;
                i === firstBit && ++exponent;
            }
            if (exponent > maxExponent || intPart) return sign ? isDouble ? this._doubleNegInifinity : this._floatNegInifinity : isDouble ? this._doubleInifinity : this._floatInifinity;
            if (exponent < minExponent) return isDouble ? this._doubleZero : this._floatZero;
            if (isDouble) {
                var high = 0;
                for (i = 0; i < 20; ++i) high = high << 1 | buffer[++firstBit];
                for (var low = 0; i < 52; ++i) low = low << 1 | buffer[++firstBit];
                return [ 255 & low, low >> 8 & 255, low >> 16 & 255, low >>> 24, 255 & (high = sign << 31 | 2147483647 & (high |= exponent + bias << 20)), high >> 8 & 255, high >> 16 & 255, high >>> 24 ];
            }
            var result = 0;
            for (i = 0; i < 23; ++i) result = result << 1 | buffer[++firstBit];
            return [ 255 & (result = sign << 31 | 2147483647 & (result |= exponent + bias << 23)), result >> 8 & 255, result >> 16 & 255, result >>> 24 ];
        }, FloatUtils._floatZero = [ 0, 0, 0, 0 ], FloatUtils._doubleZero = [ 0, 0, 0, 0, 0, 0, 0, 0 ], 
        FloatUtils._floatInifinity = [ 0, 0, 128, 127 ], FloatUtils._floatNegInifinity = [ 0, 0, 128, 255 ], 
        FloatUtils._doubleInifinity = [ 0, 0, 0, 0, 0, 0, 240, 127 ], FloatUtils._doubleNegInifinity = [ 0, 0, 0, 0, 0, 0, 240, 255 ], 
        FloatUtils;
    }();
    exports.FloatUtils = FloatUtils;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var BrowserChecker = function() {
        function BrowserChecker() {}
        return BrowserChecker._IsDataViewSupport = function() {
            return "undefined" != typeof ArrayBuffer && "undefined" != typeof DataView;
        }, BrowserChecker;
    }();
    exports.BrowserChecker = BrowserChecker;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var microsoft_bond_primitives_1 = __webpack_require__(8), MemoryStream = function() {
        function MemoryStream() {
            this._buffer = [];
        }
        return MemoryStream.prototype._WriteByte = function(byte) {
            this._buffer.push(microsoft_bond_primitives_1.Number._ToByte(byte));
        }, MemoryStream.prototype._Write = function(buffer, offset, count) {
            for (;count--; ) this._WriteByte(buffer[offset++]);
        }, MemoryStream.prototype._GetBuffer = function() {
            return this._buffer;
        }, MemoryStream;
    }();
    exports.MemoryStream = MemoryStream;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var AWTRetryPolicy = function() {
        function AWTRetryPolicy() {}
        return AWTRetryPolicy.shouldRetryForStatus = function(httpStatusCode) {
            return !(httpStatusCode >= 300 && httpStatusCode < 500 && 408 !== httpStatusCode || 501 === httpStatusCode || 505 === httpStatusCode);
        }, AWTRetryPolicy.getMillisToBackoffForRetry = function(retriesSoFar) {
            var waitDuration, randomBackoff = Math.floor(1200 * Math.random()) + 2400;
            return waitDuration = Math.pow(4, retriesSoFar) * randomBackoff, Math.min(waitDuration, 12e4);
        }, AWTRetryPolicy;
    }();
    exports.default = AWTRetryPolicy;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var AWTKillSwitch = function() {
        function AWTKillSwitch() {
            this._killedTokenDictionary = {};
        }
        return AWTKillSwitch.prototype.setKillSwitchTenants = function(killTokens, killDuration) {
            if (killTokens && killDuration) try {
                var killedTokens = killTokens.split(",");
                if ("this-request-only" === killDuration) return killedTokens;
                for (var durationMs = 1e3 * parseInt(killDuration, 10), i = 0; i < killedTokens.length; ++i) this._killedTokenDictionary[killedTokens[i]] = Date.now() + durationMs;
            } catch (ex) {
                return [];
            }
            return [];
        }, AWTKillSwitch.prototype.isTenantKilled = function(tenantToken) {
            return void 0 !== this._killedTokenDictionary[tenantToken] && this._killedTokenDictionary[tenantToken] > Date.now() || (delete this._killedTokenDictionary[tenantToken], 
            !1);
        }, AWTKillSwitch;
    }();
    exports.default = AWTKillSwitch;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var AWTClockSkewManager = function() {
        function AWTClockSkewManager(clockSkewRefreshDurationInMins) {
            this.clockSkewRefreshDurationInMins = clockSkewRefreshDurationInMins, this._reset();
        }
        return AWTClockSkewManager.prototype.allowRequestSending = function() {
            return this._isFirstRequest && !this._clockSkewSet ? (this._isFirstRequest = !1, 
            this._allowRequestSending = !1, !0) : this._allowRequestSending;
        }, AWTClockSkewManager.prototype.shouldAddClockSkewHeaders = function() {
            return this._shouldAddClockSkewHeaders;
        }, AWTClockSkewManager.prototype.getClockSkewHeaderValue = function() {
            return this._clockSkewHeaderValue;
        }, AWTClockSkewManager.prototype.setClockSkew = function(timeDeltaInMillis) {
            this._clockSkewSet || (timeDeltaInMillis ? this._clockSkewHeaderValue = timeDeltaInMillis : this._shouldAddClockSkewHeaders = !1, 
            this._clockSkewSet = !0, this._allowRequestSending = !0);
        }, AWTClockSkewManager.prototype._reset = function() {
            var _this = this;
            this._isFirstRequest = !0, this._clockSkewSet = !1, this._allowRequestSending = !0, 
            this._shouldAddClockSkewHeaders = !0, this._clockSkewHeaderValue = "use-collector-delta", 
            this.clockSkewRefreshDurationInMins > 0 && setTimeout((function() {
                return _this._reset();
            }), 6e4 * this.clockSkewRefreshDurationInMins);
        }, AWTClockSkewManager;
    }();
    exports.default = AWTClockSkewManager;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var Enums_1 = __webpack_require__(0), AWTRecordBatcher = function() {
        function AWTRecordBatcher(_outboundQueue, _maxNumberOfEvents) {
            this._outboundQueue = _outboundQueue, this._maxNumberOfEvents = _maxNumberOfEvents, 
            this._currentBatch = {}, this._currentNumEventsInBatch = 0;
        }
        return AWTRecordBatcher.prototype.addEventToBatch = function(event) {
            if (event.priority === Enums_1.AWTEventPriority.Immediate_sync) {
                var immediateBatch = {};
                return immediateBatch[event.apiKey] = [ event ], immediateBatch;
            }
            return this._currentNumEventsInBatch >= this._maxNumberOfEvents && this.flushBatch(), 
            void 0 === this._currentBatch[event.apiKey] && (this._currentBatch[event.apiKey] = []), 
            this._currentBatch[event.apiKey].push(event), this._currentNumEventsInBatch++, null;
        }, AWTRecordBatcher.prototype.flushBatch = function() {
            this._currentNumEventsInBatch > 0 && (this._outboundQueue.push(this._currentBatch), 
            this._currentBatch = {}, this._currentNumEventsInBatch = 0);
        }, AWTRecordBatcher.prototype.hasBatch = function() {
            return this._currentNumEventsInBatch > 0;
        }, AWTRecordBatcher;
    }();
    exports.default = AWTRecordBatcher;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var AWTTransmissionManagerCore_1 = __webpack_require__(4), AWTTransmissionManager = function() {
        function AWTTransmissionManager() {}
        return AWTTransmissionManager.setEventsHandler = function(eventsHandler) {
            AWTTransmissionManagerCore_1.default.setEventsHandler(eventsHandler);
        }, AWTTransmissionManager.getEventsHandler = function() {
            return AWTTransmissionManagerCore_1.default.getEventsHandler();
        }, AWTTransmissionManager.scheduleTimer = function() {
            AWTTransmissionManagerCore_1.default.scheduleTimer();
        }, AWTTransmissionManager;
    }();
    exports.default = AWTTransmissionManager;
}, function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    __webpack_require__.r(__webpack_exports__), __webpack_require__.d(__webpack_exports__, "AgaveSink", (function() {
        return AgaveSink_AgaveSink;
    })), __webpack_require__.d(__webpack_exports__, "TelemetryContext", (function() {
        return TelemetryContext_TelemetryContext;
    }));
    var LogLevel, Category, SamplingPolicy, PersistencePriority, CostPriority, DataCategories, DiagnosticLevel, DataClassification, DataFieldType, onNotificationEvent = new (function() {
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
    }());
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
    }(Category || (Category = {})), function(SamplingPolicy) {
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
    }(DiagnosticLevel || (DiagnosticLevel = {})), function(DataClassification) {
        DataClassification[DataClassification.EssentialServiceMetadata = 1] = "EssentialServiceMetadata", 
        DataClassification[DataClassification.AccountData = 2] = "AccountData", DataClassification[DataClassification.SystemMetadata = 4] = "SystemMetadata", 
        DataClassification[DataClassification.OrganizationIdentifiableInformation = 8] = "OrganizationIdentifiableInformation", 
        DataClassification[DataClassification.EndUserIdentifiableInformation = 16] = "EndUserIdentifiableInformation", 
        DataClassification[DataClassification.CustomerContent = 32] = "CustomerContent", 
        DataClassification[DataClassification.AccessControl = 64] = "AccessControl";
    }(DataClassification || (DataClassification = {})), function(DataFieldType) {
        DataFieldType[DataFieldType.String = 0] = "String", DataFieldType[DataFieldType.Boolean = 1] = "Boolean", 
        DataFieldType[DataFieldType.Int64 = 2] = "Int64", DataFieldType[DataFieldType.Double = 3] = "Double", 
        DataFieldType[DataFieldType.Guid = 4] = "Guid";
    }(DataFieldType || (DataFieldType = {}));
    var TelemetryContext_TelemetryContext = function() {
        function TelemetryContext() {}
        return TelemetryContext.setTelemetryContext = function(steContext) {
            this._steContext = steContext;
        }, TelemetryContext.getDataFieldsFromContext = function() {
            var _this = this, additionalDataFields = [];
            return Object.keys(this._steContext).forEach((function(key) {
                additionalDataFields.push({
                    name: key,
                    value: _this._steContext[key],
                    dataType: DataFieldType.String
                });
            })), additionalDataFields;
        }, TelemetryContext.getAppName = function() {
            return this._steContext["App.Name"];
        }, TelemetryContext.getAppPlatform = function() {
            return this._steContext["App.Platform"];
        }, TelemetryContext.getAppVersion = function() {
            return this._steContext["App.Version"];
        }, TelemetryContext.isWacAgave = function() {
            return "Web" === this._steContext["App.Platform"];
        }, TelemetryContext._steContext = {}, TelemetryContext;
    }(), RichApiHelper_RichApiHelper = function() {
        function RichApiHelper() {
            this._requestIsPending = !1, this._telemetryQueue = [], this._sentFirstEvent = !1;
        }
        return RichApiHelper.prototype.isSupportedByDeclaration = function() {
            return Office.context.requirements.isSetSupported("Telemetry", "1.2");
        }, RichApiHelper.prototype.isUnsupported = function() {
            return TelemetryContext_TelemetryContext.isWacAgave() || "undefined" == typeof OfficeCore;
        }, RichApiHelper.prototype.isSupportedAsync = function(callback) {
            (this._sentFirstEvent || "undefined" != typeof this._onSentFirstEvent) && (logNotification(LogLevel.Error, Category.Sink, (function() {
                return "isSupportedAsync may only be called once";
            })), callback(!1)), this._onSentFirstEvent = callback, this.sendTestEvent();
        }, RichApiHelper.prototype.sendTelemetryEvent = function(telemetryEvent) {
            this._telemetryQueue.push(telemetryEvent), this._requestIsPending || this.processWorkBacklog();
        }, RichApiHelper.prototype.sendTestEvent = function() {
            var testEvent = {
                telemetryProperties: {
                    nexusTenantToken: 1723,
                    ariaTenantToken: "f998cc5ba4d448d6a1e8e913ff18be94-dd122e0a-fcf8-4dc5-9dbb-6afac5325183-7405"
                },
                eventName: "Office.Telemetry.RichApi.TestForSupport",
                eventFlags: {
                    dataCategories: DataCategories.ProductServiceUsage,
                    diagnosticLevel: DiagnosticLevel.FullEvent
                }
            };
            this.sendTelemetryEvent(testEvent);
        }, RichApiHelper.prototype.processWorkBacklog = function() {
            var _this = this;
            this._requestIsPending = !0;
            var currentWork = this._telemetryQueue;
            this._telemetryQueue = [], this.pauseIfNecessary().then((function() {
                _this.processTelemetryEvents(currentWork), _this.waitAndProcessMore();
            })).catch((function(error) {
                logError(Category.Sink, "RichApiSink Error", error), _this.waitAndProcessMore();
            }));
        }, RichApiHelper.prototype.waitAndProcessMore = function() {
            var _this = this;
            pause(1e3).then((function() {
                _this._telemetryQueue.length > 0 && setTimeout((function() {
                    return _this.processWorkBacklog();
                }), 0), _this._requestIsPending = !1;
            })).catch((function() {}));
        }, RichApiHelper.prototype.processTelemetryEvents = function(telemetryEvents) {
            var _this = this, ctx = new OfficeCore.RequestContext;
            telemetryEvents.forEach((function(telemetryEvent) {
                if (telemetryEvent.telemetryProperties) {
                    var dataFields = [];
                    _this.addDataFields(dataFields, telemetryEvent.dataFields);
                    var contractName = telemetryEvent.eventContract ? telemetryEvent.eventContract.name : "";
                    telemetryEvent.eventContract && _this.addDataFields(dataFields, telemetryEvent.eventContract.dataFields), 
                    ctx.telemetry.sendTelemetryEvent(telemetryEvent.telemetryProperties, telemetryEvent.eventName, contractName, function(telemetryEvent) {
                        var eventFlags = {
                            costPriority: CostPriority.Normal,
                            samplingPolicy: SamplingPolicy.Measure,
                            persistencePriority: PersistencePriority.Normal,
                            dataCategories: DataCategories.NotSet,
                            diagnosticLevel: DiagnosticLevel.FullEvent
                        };
                        return telemetryEvent.eventFlags && telemetryEvent.eventFlags.dataCategories || logNotification(LogLevel.Error, Category.Core, (function() {
                            return "Event is missing DataCategories event flag";
                        })), telemetryEvent.eventFlags ? (telemetryEvent.eventFlags.costPriority && (eventFlags.costPriority = telemetryEvent.eventFlags.costPriority), 
                        telemetryEvent.eventFlags.samplingPolicy && (eventFlags.samplingPolicy = telemetryEvent.eventFlags.samplingPolicy), 
                        telemetryEvent.eventFlags.persistencePriority && (eventFlags.persistencePriority = telemetryEvent.eventFlags.persistencePriority), 
                        telemetryEvent.eventFlags.dataCategories && (eventFlags.dataCategories = telemetryEvent.eventFlags.dataCategories), 
                        telemetryEvent.eventFlags.diagnosticLevel && (eventFlags.diagnosticLevel = telemetryEvent.eventFlags.diagnosticLevel), 
                        eventFlags) : eventFlags;
                    }(telemetryEvent), dataFields);
                }
            })), this._sentFirstEvent ? ctx.sync().catch((function(e) {
                logError(Category.Sink, "RichApiError", e);
            })) : (this._sentFirstEvent = !0, ctx.sync().then((function() {
                _this._onSentFirstEvent && _this._onSentFirstEvent(!0);
            })).catch((function() {
                logNotification(LogLevel.Info, Category.Sink, (function() {
                    return "RichApiTelemetry not supported on the host.";
                })), _this._onSentFirstEvent && _this._onSentFirstEvent(!1);
            })));
        }, RichApiHelper.prototype.addDataFields = function(richApiDataFields, dataFields) {
            dataFields && dataFields.forEach((function(dataField) {
                richApiDataFields.push({
                    name: dataField.name,
                    value: dataField.value,
                    classification: dataField.classification ? dataField.classification : DataClassification.SystemMetadata,
                    type: dataField.dataType
                });
            }));
        }, RichApiHelper.prototype.pauseIfNecessary = function() {
            return this._sentFirstEvent ? Office.Promise.resolve(void 0) : pause(1e3);
        }, RichApiHelper;
    }();
    function pause(ms) {
        return new Office.Promise((function(resolve) {
            return setTimeout(resolve, ms);
        }));
    }
    var RichApiSink_RichApiSink = function() {
        function RichApiSink() {}
        return RichApiSink.prototype.sendTelemetryEvent = function(telemetryEvent) {
            RichApiSink._richApiHelper.sendTelemetryEvent(telemetryEvent);
        }, RichApiSink.isSupportedByDeclaration = function() {
            return this._richApiHelper.isSupportedByDeclaration();
        }, RichApiSink.isUnsupported = function() {
            return this._richApiHelper.isUnsupported();
        }, RichApiSink.isSupportedAsync = function(callback) {
            return this._richApiHelper.isSupportedAsync(callback);
        }, RichApiSink._richApiHelper = new RichApiHelper_RichApiHelper, RichApiSink;
    }();
    var SdxWacSink_SdxWacSink = function() {
        function SdxWacSink() {}
        return SdxWacSink.isSupported = function() {
            return ("undefined" != typeof Office && "undefined" != typeof Office.context && "undefined" != typeof Office.context.platform ? Office.context.platform === Office.PlatformType.OfficeOnline : "undefined" != typeof OfficeExt && "undefined" != typeof OfficeExt.HostName && "undefined" != typeof OfficeExt.HostName.Host && "function" == typeof OfficeExt.HostName.Host.getInstance && "function" == typeof OfficeExt.HostName.Host.getInstance().getPlatform && OfficeExt.HostName.Host.getInstance().getPlatform() === Office.PlatformType.OfficeOnline) && "object" == typeof OSF && "function" == typeof OSF.getClientEndPoint && "object" == typeof OSF._OfficeAppFactory && "function" == typeof OSF._OfficeAppFactory.getId && "object" == typeof OSF.AgaveHostAction && "number" == typeof OSF.AgaveHostAction.SendTelemetryEvent;
        }, SdxWacSink.prototype.sendTelemetryEvent = function(event, _timestamp) {
            try {
                if (event.dataFields && event.dataFields.filter((function(dataField) {
                    return dataField.classification && dataField.classification !== DataClassification.SystemMetadata;
                })).length > 0) return;
                var id = OSF._OfficeAppFactory.getId(), SendTelemetryEventId = OSF.AgaveHostAction.SendTelemetryEvent;
                OSF.getClientEndPoint().invoke("ContextActivationManager_notifyHost", null, [ id, SendTelemetryEventId, event ]);
            } catch (error) {
                logError(Category.Sink, "AgaveWacSink", error);
            }
        }, SdxWacSink;
    }(), OutlookSink_OutlookSink = function() {
        function OutlookSink() {
            this._supportsAllEvents = !1, this._contextDataFields = [], this._supportsAllEvents = Office.context.requirements.isSetSupported("OutlookTelemetry", "1.1");
            var platform = TelemetryContext_TelemetryContext.getAppPlatform();
            "iOS" !== platform && "Android" !== platform || (this._contextDataFields = TelemetryContext_TelemetryContext.getDataFieldsFromContext());
        }
        return OutlookSink.isSupported = function() {
            try {
                return Office.context.requirements.isSetSupported("OutlookTelemetry");
            } catch (_a) {
                return !1;
            }
        }, OutlookSink.prototype.sendTelemetryEvent = function(event) {
            this._supportsAllEvents || event.eventName.match(/^Office\.Extensibility\.OfficeJs\.[a-zA-Z]*$/) ? (this.addAdditionalDataFields(event), 
            Office.context.mailbox.logTelemetry(JSON.stringify(event))) : logNotification(LogLevel.Info, Category.Sink, (function() {
                return "This version of Outlook only accepts OfficeJS telemetry events";
            }));
        }, OutlookSink.prototype.addAdditionalDataFields = function(event) {
            var _a;
            event.dataFields = event.dataFields || [], (_a = event.dataFields).push.apply(_a, this._contextDataFields);
        }, OutlookSink;
    }(), MAX_SUPPORTED_WIN32_VERSION = [ 16, 0, 11599 ], MIN_SUPPORTED_WIN32_VERSION = [ 16, 0, 4266 ], MAX_SUPPORTED_MAC_VERSION = [ 16, 26 ], MAX_SUPPORTED_IOS_WXP_VERSION = [ 2, 29 ], MAX_SUPPORTED_WEB_OUTLOOK_VERSION = [ 16, 0, 99999 ], ALLOWED_APPS = [ "Excel", "Outlook", "PowerPoint", "Project", "Word" ], _sendEventEnabled = !1;
    function canSendToAria() {
        if (!_sendEventEnabled) return !1;
        var platform = TelemetryContext_TelemetryContext.getAppPlatform();
        return function(version, platform, appName) {
            if (!version) return !0;
            if (!appName || ALLOWED_APPS.indexOf(appName) < 0) return !1;
            var maxVersion, minVersion = [];
            if ("Win32" === platform) maxVersion = MAX_SUPPORTED_WIN32_VERSION, minVersion = MIN_SUPPORTED_WIN32_VERSION; else if ("Mac" === platform) maxVersion = MAX_SUPPORTED_MAC_VERSION; else if ("iOS" === platform) {
                if ("Outlook" === appName) return !0;
                maxVersion = MAX_SUPPORTED_IOS_WXP_VERSION;
            } else {
                if ("Web" !== platform || "Outlook" !== appName) return !1;
                maxVersion = MAX_SUPPORTED_WEB_OUTLOOK_VERSION;
            }
            var versionArray = String(version).split(".").map((function(x) {
                return parseInt(x, 10);
            }));
            return isVersionLessThanOrEqual(minVersion, versionArray) && isVersionLessThanOrEqual(versionArray, maxVersion);
        }(TelemetryContext_TelemetryContext.getAppVersion(), platform, TelemetryContext_TelemetryContext.getAppName());
    }
    function isVersionLessThanOrEqual(smallerVersion, largerVersion) {
        for (var i = 0; i < smallerVersion.length && i < largerVersion.length; i++) {
            if (isNaN(smallerVersion[i]) || isNaN(largerVersion[i])) return !1;
            if (smallerVersion[i] < largerVersion[i]) return !0;
            if (smallerVersion[i] > largerVersion[i]) return !1;
        }
        return !0;
    }
    var AriaSDK = __webpack_require__(1), Enums = __webpack_require__(0), EVENT_NAME_DOT_REPLACE_REGEX = /\./g, eventSequence = 0;
    function addDataFields(ariaEvent, fields, prependDataToken) {
        fields && fields.forEach((function(field) {
            if (!field.classification || field.classification === DataClassification.SystemMetadata || field.classification === DataClassification.EssentialServiceMetadata) {
                var _a = [ "", "", field.name ], metadataPrefix = _a[0], dataToken = _a[1], fieldName = _a[2], firstSeparator = field.name.indexOf(".");
                firstSeparator > 0 && "zC" === field.name.substr(0, firstSeparator) && (metadataPrefix = field.name.substring(0, firstSeparator + 1), 
                fieldName = field.name.substring(firstSeparator + 1)), prependDataToken && (dataToken = "Data.");
                var ariaFieldName = metadataPrefix + dataToken + fieldName;
                ariaEvent.properties[ariaFieldName] = {
                    value: field.value,
                    type: mapDataFieldTypeToAWTPropertyType(field.dataType)
                };
            }
        }));
    }
    function mapDataFieldTypeToAWTPropertyType(otelType) {
        switch (otelType) {
          case DataFieldType.String:
          case DataFieldType.Guid:
            return Enums.AWTPropertyType.String;

          case DataFieldType.Boolean:
            return Enums.AWTPropertyType.Boolean;

          case DataFieldType.Int64:
            return Enums.AWTPropertyType.Int64;

          case DataFieldType.Double:
            return Enums.AWTPropertyType.Double;

          default:
            throw new Error(otelType);
        }
    }
    var sendRequestActualMethod, sendPostActualMethod, statsCallback, payloadSizeCallback, AWTTransmissionManagerCore = __webpack_require__(4), AWTTransmissionManagerCore_default = __webpack_require__.n(AWTTransmissionManagerCore), AWTQueueManager = __webpack_require__(11), AWTQueueManager_default = __webpack_require__.n(AWTQueueManager);
    var Wrappers = {
        sendRequestWrapper: function(request, retryCount, isTeardown, isSynchronous) {
            void 0 === isSynchronous && (isSynchronous = !1);
            var numOfEvents = 0;
            for (var token in request) request.hasOwnProperty(token) && (numOfEvents += request[token].length);
            var startTime = performance.now();
            sendRequestActualMethod(request, retryCount, isTeardown, isSynchronous), statsCallback(performance.now() - startTime, numOfEvents);
        },
        sendPostWrapper: function(urlString, data, ontimeout, onerror, onload, sync) {
            sendPostActualMethod(urlString, data, ontimeout, onerror, onload, sync), payloadSizeCallback(data.length);
        }
    };
    var awtInitialized = !1;
    function sendEvent(telemetryEvent, additionalDataFields, timestamp) {
        var ariaEvent;
        if (initialize(), !telemetryEvent.telemetryProperties || !telemetryEvent.telemetryProperties.ariaTenantToken) throw new Error("Missing Aria Tenant Token");
        ariaEvent = function(telemetryEvent, additionalDataFields, timestamp) {
            var eventName, timestampLocal, ariaEvent = {
                name: (eventName = telemetryEvent.eventName, eventName.toLowerCase().replace(EVENT_NAME_DOT_REPLACE_REGEX, "_")),
                properties: {}
            };
            return ariaEvent.properties["Event.Sequence"] = {
                value: ++eventSequence,
                type: Enums.AWTPropertyType.Int64
            }, ariaEvent.properties["Event.Name"] = telemetryEvent.eventName, ariaEvent.properties["Event.Source"] = "OTelJS", 
            timestampLocal = timestamp ? new Date(timestamp) : new Date, ariaEvent.properties["Event.Time"] = {
                value: timestampLocal,
                type: Enums.AWTPropertyType.Date
            }, telemetryEvent.eventContract && (ariaEvent.properties["Event.Contract"] = telemetryEvent.eventContract.name, 
            addDataFields(ariaEvent, telemetryEvent.eventContract.dataFields, !1)), addDataFields(ariaEvent, additionalDataFields, !1), 
            addDataFields(ariaEvent, telemetryEvent.dataFields, !0), ariaEvent;
        }(telemetryEvent, additionalDataFields, timestamp), AriaSDK.AWTLogManager.getLogger(telemetryEvent.telemetryProperties.ariaTenantToken).logEvent(ariaEvent);
    }
    function initialize(configuration, ariaSinkProperties) {
        var notificationListener;
        awtInitialized || (AriaSDK.AWTLogManager.initialize("cd836626611c4caaa8fc5b2e728ee81d-3b6d6c45-6377-4bf5-9792-dbf8e1881088-7521", configuration), 
        ariaSinkProperties && (!function(uploadFrequency) {
            if (!uploadFrequency) return;
            var normalLatencyFrequency = uploadFrequency / 1e3, highLatencyFrequency = normalLatencyFrequency / 2, lowLatencyFrequency = 2 * normalLatencyFrequency, customProfiles = {};
            customProfiles.OTelCustomTransmissionProfile = [ lowLatencyFrequency, normalLatencyFrequency, highLatencyFrequency ], 
            AriaSDK.AWTLogManager.loadTransmitProfiles(customProfiles), AriaSDK.AWTLogManager.setTransmitProfile("OTelCustomTransmissionProfile");
        }(ariaSinkProperties.uploadFrequency), (notificationListener = ariaSinkProperties.notificationListener) && AriaSDK.AWTLogManager.addNotificationListener(notificationListener), 
        function(callbacks) {
            if (!callbacks) return;
            (function(callbackForStats, callbackForPayloadSize) {
                var eventHandler = AWTTransmissionManagerCore_default.a.getEventsHandler();
                if (!(eventHandler instanceof AWTQueueManager_default.a)) return !1;
                var httpManager = eventHandler._httpManager;
                if (!(httpManager && httpManager._sendRequest && httpManager._httpInterface)) return !1;
                sendRequestActualMethod = httpManager._sendRequest.bind(httpManager), httpManager._sendRequest = Wrappers.sendRequestWrapper;
                var httpInterface = httpManager._httpInterface;
                return sendPostActualMethod = httpInterface.sendPOST.bind(httpManager), httpInterface.sendPOST = Wrappers.sendPostWrapper, 
                statsCallback = callbackForStats, payloadSizeCallback = callbackForPayloadSize, 
                !0;
            })(callbacks.requestProcessingStats, callbacks.networkStats) || logNotification(LogLevel.Error, Category.Sink, (function() {
                return "Failed to instrument Aria delivery task";
            }));
        }(ariaSinkProperties.stats), ariaSinkProperties.disableStatsTracking || AriaSDK.AWTLogManager.addNotificationListener({
            eventsSent: function(events) {
                logNotification(LogLevel.Info, Category.Transport, (function() {
                    return "Successfully sent " + events.length + " event(s)";
                })), logNotification(LogLevel.Verbose, Category.Transport, (function() {
                    return "Sent event(s) details : " + JSON.stringify(events, null, 2);
                })), events.length;
            },
            eventsDropped: function(events, reason) {
                logNotification(LogLevel.Error, Category.Transport, (function() {
                    return "Dropped " + events.length + " event(s) because " + reason;
                })), logNotification(LogLevel.Verbose, Category.Transport, (function() {
                    return "Dropped event(s) details : " + JSON.stringify(events, null, 2);
                })), events.length;
            },
            eventsRejected: function(events, reason) {
                logNotification(LogLevel.Error, Category.Transport, (function() {
                    return "Rejected " + events.length + " event(s) because " + reason;
                })), logNotification(LogLevel.Verbose, Category.Transport, (function() {
                    return "Rejected event(s) details : " + JSON.stringify(events, null, 2);
                })), events.length;
            },
            eventsRetrying: function(events) {
                logNotification(LogLevel.Warning, Category.Transport, (function() {
                    return "Retrying " + events.length + " event(s)";
                })), logNotification(LogLevel.Verbose, Category.Transport, (function() {
                    return "Retrying event(s) details : " + JSON.stringify(events, null, 2);
                })), events.length;
            }
        })), awtInitialized = !0);
    }
    var AriaSinkType, FullEventProcessor_FullEventProcessor = function() {
        function FullEventProcessor() {
            this._fullEventsEnabled = !1;
        }
        return FullEventProcessor.prototype.processEvent = function(event) {
            return this._fullEventsEnabled || !!event.eventFlags && (event.eventFlags.diagnosticLevel === DiagnosticLevel.BasicEvent || event.eventFlags.diagnosticLevel === DiagnosticLevel.NecessaryServiceDataEvent || event.eventFlags.diagnosticLevel === DiagnosticLevel.AlwaysOnNecessaryServiceDataEvent);
        }, FullEventProcessor.prototype.setFullEventsEnabled = function(enabled) {
            this._fullEventsEnabled = enabled;
        }, FullEventProcessor;
    }();
    !function(AriaSinkType) {
        AriaSinkType[AriaSinkType.Aria = 0] = "Aria", AriaSinkType[AriaSinkType.AriaSE = 1] = "AriaSE";
    }(AriaSinkType || (AriaSinkType = {}));
    var AriaSink_AriaSink = function() {
        function AriaSink(additionalDataFields, ariaSinkProperties) {
            if (void 0 === additionalDataFields && (additionalDataFields = []), this._preprocessors = [], 
            this.additionalDataFields = additionalDataFields, this._fullEventProcessor = new FullEventProcessor_FullEventProcessor, 
            this.addPreprocessor(this._fullEventProcessor), void 0 === AriaSink.ariaSinkType) AriaSink.ariaSinkType = this.getSinkType(); else if (AriaSink.ariaSinkType !== this.getSinkType()) throw new Error("Multiple Aria Configurations are not allowed");
            initialize(this.getAWTLogConfiguration(ariaSinkProperties), ariaSinkProperties);
        }
        return AriaSink.prototype.getSinkType = function() {
            return AriaSinkType.Aria;
        }, AriaSink.prototype.getAWTLogConfiguration = function(ariaSinkProperties) {
            var awtLogConfiguration = {
                disableCookiesUsage: !0,
                canSendStatEvent: function() {
                    return !1;
                }
            };
            return ariaSinkProperties && (awtLogConfiguration.cacheMemorySizeLimitInNumberOfEvents = ariaSinkProperties.eventsLimitInMem, 
            awtLogConfiguration.collectorUri = ariaSinkProperties.endpointUrl), awtLogConfiguration;
        }, AriaSink.prototype.sendTelemetryEvent = function(event, timestamp) {
            try {
                for (var i = 0; i < this._preprocessors.length; i++) if (!this._preprocessors[i].processEvent(event)) return;
                sendEvent(event, this.additionalDataFields, timestamp);
            } catch (error) {
                logError(Category.Sink, "AriaSink", error);
            }
        }, AriaSink.prototype.getAdditionalDataFields = function() {
            return this.additionalDataFields;
        }, AriaSink.prototype.addPreprocessor = function(preprocessor) {
            this._preprocessors.push(preprocessor);
        }, AriaSink.prototype.setFullEventsEnabled = function(enabled) {
            this._fullEventProcessor.setFullEventsEnabled(enabled);
        }, AriaSink.prototype.flushAsync = function(callback) {
            AriaSDK.AWTLogManager.flush(callback);
        }, AriaSink.prototype.shutdown = function() {
            AriaSDK.AWTLogManager.flushAndTeardown();
        }, AriaSink;
    }();
    var AgaveSink_AgaveSink = function() {
        function AgaveSink(steContext, ariaSendEventEnabled, onReady) {
            this._isUsable = !0, this._awaitingInitialization = !1, this._eventQueue = [], this._onReady = onReady, 
            TelemetryContext_TelemetryContext.setTelemetryContext(steContext), _sendEventEnabled = ariaSendEventEnabled, 
            this.initialize();
        }
        return AgaveSink.createInstance = function(steContext, ariaSendEventEnabled) {
            void 0 === steContext && (steContext = {}), void 0 === ariaSendEventEnabled && (ariaSendEventEnabled = !0);
            var agaveSink = new AgaveSink(steContext, ariaSendEventEnabled);
            if (agaveSink._isUsable) return agaveSink;
        }, AgaveSink.prototype.initialize = function() {
            if (!this.isTelemetryEnabled()) return this.failToInitialize("AppTelemetry is disabled for this platform.");
            if (RichApiSink_RichApiSink.isSupportedByDeclaration()) this.connectRichApiSink(); else if (OutlookSink_OutlookSink.isSupported()) this.connectSink("OutlookSink", new OutlookSink_OutlookSink); else if (TelemetryContext_TelemetryContext.isWacAgave() && SdxWacSink_SdxWacSink.isSupported()) this.connectSink("SdxWacSink", new SdxWacSink_SdxWacSink); else if (RichApiSink_RichApiSink.isUnsupported()) {
                if (!canSendToAria()) return this.failToInitialize();
                this.connectAriaSink();
            } else this._awaitingInitialization = !0, RichApiSink_RichApiSink.isSupportedAsync(this.onRichApiSupportedAsync.bind(this));
        }, AgaveSink.prototype.onRichApiSupportedAsync = function(isSupported) {
            var _this = this;
            isSupported ? this.connectSink("RichApiSink", new RichApiSink_RichApiSink) : canSendToAria() ? this.connectAriaSink() : this.failToInitialize(), 
            this._awaitingInitialization = !1, this._eventQueue.forEach((function(event) {
                _this.sendTelemetryEvent(event);
            }));
        }, AgaveSink.prototype.failToInitialize = function(reason) {
            this._isUsable = !1, this._awaitingInitialization = !1;
            var errorMessage = reason || "AgaveSink has no available sink";
            logNotification(LogLevel.Error, Category.Sink, (function() {
                return errorMessage;
            })), this._onReady && this._onReady(!1);
        }, AgaveSink.prototype.sendTelemetryEvent = function(event) {
            if (this._awaitingInitialization && this._isUsable) this._eventQueue.push(event); else if (this._sink) try {
                this._sink.sendTelemetryEvent(event);
            } catch (error) {
                logError(Category.Sink, "AgaveSink", error);
            } else logNotification(LogLevel.Error, Category.Sink, (function() {
                return "AgaveSink has no available sink";
            }));
        }, AgaveSink.prototype.connectSink = function(sinkName, sink) {
            this._sink = sink, logNotification(LogLevel.Info, Category.Sink, (function() {
                return "AgaveSink is using " + sinkName;
            })), this._onReady && this._onReady(!0);
        }, AgaveSink.prototype.connectRichApiSink = function() {
            this.connectSink("RichApiSink", new RichApiSink_RichApiSink);
        }, AgaveSink.prototype.connectAriaSink = function() {
            var additionalDataFields, _this = this;
            additionalDataFields = TelemetryContext_TelemetryContext.getDataFieldsFromContext(), 
            function(ariaSink) {
                ariaSink ? _this.connectSink("AriaSink", ariaSink) : _this.failToInitialize();
            }(new AriaSink_AriaSink(additionalDataFields));
        }, AgaveSink.prototype.isTelemetryEnabled = function() {
            return !("undefined" == typeof OSF || "undefined" == typeof OSF.AppTelemetry || "boolean" != typeof OSF.AppTelemetry.enableTelemetry || !OSF.AppTelemetry.enableTelemetry) || (logNotification(LogLevel.Warning, Category.Core, (function() {
                return "AppTelemetry is disabled for this platform.";
            })), !1);
        }, AgaveSink;
    }();
} ]);