const childProcess = require('child_process');
const defaults = require('./defaults');
const fs = require('fs');
const os = require('os');

function addSecretToCredentialStore(ssoAppName, secret) {
    try {
        switch (process.platform) {
            case "win32":
                console.log(`Adding application secret for ${ssoAppName} to Windows Credential Store`);
                const addSecretToWindowsStoreCommand = `powershell -ExecutionPolicy Bypass -File "${defaults.addSecretCommandPath}" "${ssoAppName}" "${os.userInfo().username}" "${secret}"`;
                childProcess.execSync(addSecretToWindowsStoreCommand, { stdio: "pipe" });
                break;
            case "darwin":
                console.log(`Adding application secret for ${ssoAppName} to Mac OS Keychain`);
                const addSecretToMacStoreCommand = `sudo security add-generic-password -a ${os.userInfo().username} -s "${ssoAppName}" -w "${secret}"`;
                childProcess.execSync(addSecretToMacStoreCommand, { stdio: "pipe" });
                break;
            default:
                throw new Error(`Platform not supported: ${process.platform}`);
        }
    } catch (err) {
        throw new Error(`Unable to add secret for ${ssoAppName} to Windows Credential Store. \n${err}`);
    }
}

function getSecretFromCredentialStore(ssoAppName) {
    try {
        switch (process.platform) {
            case "win32":
                console.log(`Getting application secret for ${ssoAppName} from Windows Credential Store`);
                const getSecretFromWindowsStoreCommand = `powershell -ExecutionPolicy Bypass -File "${defaults.getSecretCommandPath}" "${ssoAppName}" "${os.userInfo().username}"`;
                return childProcess.execSync(getSecretFromWindowsStoreCommand, { stdio: "pipe" }).toString();
            case "darwin":
                console.log(`Getting application secret for ${ssoAppName} from Mac OS Keychain`);
                const getSecretFromMacStoreCommand = `sudo security find-generic-password -a ${os.userInfo().username} -s ${ssoAppName} -w`;
                return childProcess.execSync(getSecretFromMacStoreCommand, { stdio: "pipe" }).toString();;
            default:
                throw new Error(`Platform not supported: ${process.platform}`);
        }

    } catch (err) {
        throw new Error(`Unable to retrieve secret for ${ssoAppName} to Windows Credential Store. \n${err}`);
    }
}

function writeApplicationData(applicationId) {
    try {
        // Update .ENV file
        if (fs.existsSync(defaults.ssoDataFilePath)) {
            const appData = fs.readFileSync(defaults.ssoDataFilePath, 'utf8');
            const updatedAppData = appData.replace('CLIENT_ID=', `CLIENT_ID=${applicationId}`);
            fs.writeFileSync(defaults.ssoDataFilePath, updatedAppData);
        } else {
            throw new Error(`${defaults.ssoDataFilePath} does not exist`)
        }
    } catch (err) {
        throw new Error(`Unable to write SSO application data to ${defaults.ssoDataFilePath}. \n${err}`);
    }

    try {
        // Update fallbackAuthDialog.js
        if (fs.existsSync(defaults.fallbackAuthDialogFilePath)) {
            const srcFile = fs.readFileSync(defaults.fallbackAuthDialogFilePath, 'utf8');
            const updatedSrcFile = srcFile.replace('{application GUID here}', applicationId);
            fs.writeFileSync(defaults.fallbackAuthDialogFilePath, updatedSrcFile);
        } else {
            throw new Error(`${defaults.fallbackAuthDialogFilePath} does not exist`)
        }
    } catch (err) {
        throw new Error(`Unable to write SSO application data to ${defaults.fallbackAuthDialogFilePath}. \n${err}`);
    }
}

exports.addSecretToCredentialStore = addSecretToCredentialStore;
exports.getSecretFromCredentialStore = getSecretFromCredentialStore;
exports.writeApplicationData = writeApplicationData;