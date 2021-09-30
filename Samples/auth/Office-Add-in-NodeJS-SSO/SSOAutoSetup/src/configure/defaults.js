"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const path = require("path");
exports.azAddSecretCommandPath = path.resolve(`${__dirname}/scripts/azAddSecret.txt`);
exports.azCliInstallCommandPath = path.resolve(`${__dirname}/scripts/azCliInstallCmd.ps1`);
exports.azRestpCreateCommandPath = path.resolve(`${__dirname}/scripts/azRestAppCreateCmd.txt`);
exports.fallbackAuthDialogFilePath = path.resolve(`${process.cwd()}/src/public/javascripts/fallbackAuthDialog.js`);
exports.getApplicationInfoCommandPath = path.resolve(`${__dirname}/scripts/azGetApplicationInfoCmd.txt`);
exports.getInstalledAppsPath = path.resolve(`${__dirname}/scripts/getInstalledApps.ps1`);
exports.grantAdminConsentCommandPath = path.resolve(`${__dirname}/scripts/azGrantAdminConsentCmd.txt`);
exports.manifestPath = path.resolve(`${process.cwd()}/manifest.xml`);
exports.setIdentifierUriCommmandPath = path.resolve(`${__dirname}/scripts/azRestSetIdentifierUri.txt`);
exports.setSigninAudienceCommandPath = path.resolve(`${__dirname}/scripts/azSetSignInAudienceCmd.txt`);
exports.ssoDataFilePath = path.resolve(`${process.cwd()}/.ENV`);
exports.addSecretCommandPath = path.resolve(`${__dirname}/scripts/addAppSecret.ps1`);
exports.getSecretCommandPath = path.resolve(`${__dirname}/scripts/getAppSecret.ps1`);
//# sourceMappingURL=defaults.js.map