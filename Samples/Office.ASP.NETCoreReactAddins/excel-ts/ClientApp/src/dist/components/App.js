"use strict";
/*
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
/// <reference types="office-js" />
/* global Excel */ //Required for these to be found.  see: https://github.com/OfficeDev/office-js-docs-pr/issues/691
var React = require("react");
var office_ui_fabric_react_1 = require("office-ui-fabric-react");
var Header_1 = require("./Header");
var HeroList_1 = require("./HeroList");
var Progress_1 = require("./Progress");
var logo = require('../assets/logo-filled.png');
var App = /** @class */ (function (_super) {
    __extends(App, _super);
    function App(props, context) {
        var _this = _super.call(this, props, context) || this;
        _this.click = function () { return __awaiter(_this, void 0, void 0, function () {
            var error_1;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, Excel.run(function (context) { return __awaiter(_this, void 0, void 0, function () {
                                var range;
                                return __generator(this, function (_a) {
                                    switch (_a.label) {
                                        case 0:
                                            range = context.workbook.getSelectedRange();
                                            // Read the range address
                                            range.load('address');
                                            // Update the fill color
                                            range.format.fill.color = 'red';
                                            return [4 /*yield*/, context.sync()];
                                        case 1:
                                            _a.sent();
                                            console.log("The range address was " + range.address + ".");
                                            return [2 /*return*/];
                                    }
                                });
                            }); })];
                    case 1:
                        _a.sent();
                        return [3 /*break*/, 3];
                    case 2:
                        error_1 = _a.sent();
                        console.error(error_1);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        }); };
        _this.state = {
            listItems: [],
        };
        return _this;
    }
    App.prototype.componentDidMount = function () {
        this.setState({
            listItems: [
                {
                    icon: 'Ribbon',
                    primaryText: 'Achieve more with Office integration',
                },
                {
                    icon: 'Unlock',
                    primaryText: 'Unlock features and functionality',
                },
                {
                    icon: 'Design',
                    primaryText: 'Create and visualize like a pro',
                },
            ],
        });
    };
    App.prototype.render = function () {
        var _a = this.props, title = _a.title, isOfficeInitialized = _a.isOfficeInitialized;
        if (!isOfficeInitialized) {
            return (React.createElement(Progress_1.default, { title: title, logo: logo, message: "Please sideload your addin to see app body." }));
        }
        return (React.createElement("div", { className: "ms-welcome" },
            React.createElement(Header_1.default, { logo: logo, title: this.props.title, message: "Welcome TypeScript" }),
            React.createElement(HeroList_1.default, { message: "Discover what Office .NET Core 3.1 Add-ins can do for you today!", items: this.state.listItems },
                React.createElement("p", { className: "ms-font-l" },
                    "Modify the source files, then click ",
                    React.createElement("b", null, "Run"),
                    "."),
                React.createElement(office_ui_fabric_react_1.Button, { className: "ms-welcome__action", buttonType: office_ui_fabric_react_1.ButtonType.hero, iconProps: { iconName: 'ChevronRight' }, onClick: this.click }, "Run"))));
    };
    return App;
}(React.Component));
exports.default = App;
//# sourceMappingURL=App.js.map