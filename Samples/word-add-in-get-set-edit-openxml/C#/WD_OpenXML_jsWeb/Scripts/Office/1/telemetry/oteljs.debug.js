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
            parent && (this.onSendEvent = parent.onSendEvent, (_a = this.persistentDataFields).push.apply(_a, parent.persistentDataFields), 
            this.config = __assign(__assign({}, parent.getConfig()), this.config)), persistentDataFields && (_b = this.persistentDataFields).push.apply(_b, persistentDataFields);
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
            event.dataFields && (event.dataFields.unshift(Object(DataFieldHelper.e)("OTelJS.Version", "3.1.74")), 
            this.persistentDataFields && (_a = event.dataFields).unshift.apply(_a, this.persistentDataFields)), 
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