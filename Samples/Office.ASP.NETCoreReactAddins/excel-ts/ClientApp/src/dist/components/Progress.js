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
Object.defineProperty(exports, "__esModule", { value: true });
var React = require("react");
var Header_1 = require("./Header");
var HeroList_1 = require("./HeroList");
var office_ui_fabric_react_1 = require("office-ui-fabric-react");
var Progress = /** @class */ (function (_super) {
    __extends(Progress, _super);
    function Progress() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Progress.prototype.render = function () {
        var _a = this.props, logo = _a.logo, message = _a.message, title = _a.title;
        return (React.createElement("section", { className: "ms-welcome__progress ms-u-fadeIn500" },
            React.createElement(Header_1.default, { logo: logo, title: title, message: "Welcome TypeScript" }),
            React.createElement(HeroList_1.default, { message: "Discover what Office .NET Core 3.1 Add-ins can do for you today!", items: [] },
                React.createElement(office_ui_fabric_react_1.Spinner, { type: office_ui_fabric_react_1.SpinnerType.large, label: message }))));
    };
    return Progress;
}(React.Component));
exports.default = Progress;
//# sourceMappingURL=Progress.js.map