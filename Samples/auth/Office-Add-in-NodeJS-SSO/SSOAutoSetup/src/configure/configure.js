const ssoAppData = require('./ssoAppDataSetttings');
const childProcess = require('child_process');
const defaults = require('./defaults');
require('dotenv').config();
const fs = require('fs');
const manifest = require('office-addin-manifest');

configureSSOApplication();

async function configureSSOApplication() {
    // Check to see if Azure CLI is installed.  If it isn't installed then install it
    const cliInstalled = await azureCliInstalled();
    if(!cliInstalled) {
        console.log("Azure CLI is not installed.  Installing now before proceeding");
        await installAzureCli();
        console.log('Please close your command shell, reopen and run configure-sso again.  This is neccessary to register the path to the Azure CLI');
        return;
    }

    const userJson = await logIntoAzure();
    if (userJson) {
        console.log('Login was successful!');
        const manifestInfo = await manifest.readManifestFile(defaults.manifestPath);
        const applicationJson = await createNewApplication(manifestInfo.displayName);
        ssoAppData.writeApplicationData(applicationJson.appId);
        const secretJson = await setApplicationSecret(applicationJson);
        ssoAppData.addSecretToCredentialStore(manifestInfo.displayName, secretJson.secretText);
        updateProjectManifest(applicationJson.appId);
        await logoutAzure();
        console.log("Outputting Azure application info:\n");
        console.log(applicationJson);
        
    }
    else {
        throw new Error(`Login to Azure did not succeed.`);
    }
}

async function createNewApplication(ssoAppName) {
    try {
        console.log('Registering new application in Azure');
        let azRestNewAppCommand = await fs.readFileSync(defaults.azRestpCreateCommandPath, 'utf8');
        const re = new RegExp('{SSO-AppName}', 'g');
        azRestNewAppCommand = azRestNewAppCommand.replace(re, ssoAppName).replace('{PORT}', process.env.PORT);
        const applicationJson = await promiseExecuteCommand(azRestNewAppCommand, true /* returnJson */, true /* configureSSO */);
        if (applicationJson) {
            console.log('Application was successfully registered with Azure');
        }
        return applicationJson;
    } catch (err) {
        throw new Error(`Unable to register new application ${ssoAppName}. \n${err}`);
    }
}

async function applicationReady(applicationJson) {
    try {
        let azRestCommand = await fs.readFileSync(defaults.getApplicationInfoCommandPath, 'utf8');
        azRestCommand = azRestCommand.replace('<App_ID>', applicationJson.appId);
        const appJson = await promiseExecuteCommand(azRestCommand);
        return appJson !== "";
    } catch (err) {
        throw new Error(`Unable to get application info for ${applicationJson.displayName}. \n${err}`);
    }    
}

async function grantAdminContent(applicationJson) {
    try {
        console.log('Granting admin consent');
        // Check to see if the application is available before granting admin consent
        let appReady = false;
        while (appReady === false) {
            appReady = await applicationReady(applicationJson);
        }        
        let azRestCommand = fs.readFileSync(defaults.grantAdminConsentCommandPath, 'utf8');
        azRestCommand = azRestCommand.replace('<App_ID>', applicationJson.appId);
        await promiseExecuteCommand(azRestCommand);
    } catch (err) {
        throw new Error(`Unable to set grant admin consent for ${applicationJson.displayName}. \n${err}`);
    }
}

async function azureCliInstalled() {
    try {
        switch (process.platform) {
            case "win32":
                const appsInstalledWindowsCommand = `powershell -ExecutionPolicy Bypass -File "${defaults.getInstalledAppsPath}"`;
                const appsWindows = await promiseExecuteCommand(appsInstalledWindowsCommand);
                return appsWindows.filter(app => app.DisplayName === 'Microsoft Azure CLI').length > 0
            case "darwin": 
                const appsInstalledMacCommand = 'brew list';
                const appsMac = await promiseExecuteCommand(appsInstalledMacCommand, false /* returnJson */);
                return appsMac.includes('azure-cli');;;
            default:
                throw new Error(`Platform not supported: ${process.platform}`);
        }
    } catch (err) {
        throw new Error(`Unable to install Azure CLI. \n${err}`);
    }
}

async function installAzureCli() {
    try {
        console.log("Downloading and installing Azure CLI - this could take a minute or so");
        switch (process.platform) {
            case "win32":
                const windowsCliInstallCommand = `powershell -ExecutionPolicy Bypass -File "${defaults.azCliInstallCommandPath}"`;
                await promiseExecuteCommand(windowsCliInstallCommand, false /* returnJson */);
                break;
            case "darwin": // macOS
                const macCliInstallCommand = 'brew update && brew install azure-cli';
                await promiseExecuteCommand(macCliInstallCommand, false /* returnJson */);
                break;
            default:
                throw new Error(`Platform not supported: ${process.platform}`);
        }
    } catch (err) {
        throw new Error(`Unable to install Azure CLI. \n${err}`);
    }
}

async function logIntoAzure() {
    console.log('Opening browser for authentication to Azure. Enter valid Azure credentials');
    return await promiseExecuteCommand('az login --allow-no-subscriptions');
}

async function logoutAzure() {
    console.log('Logging out of Azure now');
    return await promiseExecuteCommand('az logout');
}


async function promiseExecuteCommand(cmd, returnJson = true, configureSSO = false) {
    return new Promise((resolve, reject) => {
        try {
            childProcess.exec(cmd, async (err, stdout, stderr) => {
                let results = stdout;              
                if (results !== '' && returnJson) {
                    results = JSON.parse(results);
                }
                if (configureSSO) {
                    await setIdentifierUri(results);
                    await setSignInAudience(results);
                    await grantAdminContent(results);
                }
                resolve(results);
            });
        } catch (err) {
            reject(err);
        }
    });
}

async function setApplicationSecret(applicationJson) {
    try {
        console.log('Setting identifierUri');
        let azRestCommand = await fs.readFileSync(defaults.azAddSecretCommandPath, 'utf8');
        azRestCommand = azRestCommand.replace('<App_Object_ID>', applicationJson.id);
        const secretJson = await promiseExecuteCommand(azRestCommand);
        return secretJson;
    } catch (err) {
        throw new Error(`Unable to set identifierUri for ${applicationJson.displayName}. \n${err}`);
    }
}

async function setIdentifierUri(applicationJson) {
    try {
        console.log('Setting identifierUri');
        let azRestCommand = await fs.readFileSync(defaults.setIdentifierUriCommmandPath, 'utf8');
        azRestCommand = azRestCommand.replace('<App_Object_ID>', applicationJson.id).replace('<App_Id>', applicationJson.appId).replace('{PORT}', process.env.PORT);
        await promiseExecuteCommand(azRestCommand);
    } catch (err) {
        throw new Error(`Unable to set identifierUri for ${applicationJson.displayName}. \n${err}`);
    }
}

async function setSignInAudience(applicationJson) {
    try {
        console.log('Setting signin audience');
        let azRestCommand = await fs.readFileSync(defaults.setSigninAudienceCommandPath, 'utf8');
        azRestCommand = azRestCommand.replace('<App_Object_ID>', applicationJson.id);
        await promiseExecuteCommand(azRestCommand);
    } catch (err) {
        throw new Error(`Unable to set signInAudience for ${applicationJson.displayName}. \n${err}`);
    }
}

async function updateProjectManifest(applicationId) {
    console.log('Updating manifest with application ID');
    try {
        // Update manifest with application guid and unique manifest id
        const manifestContent = await fs.readFileSync(defaults.manifestPath, 'utf8');
        const re = new RegExp('{application GUID here}', 'g');
        const updatedManifestContent = manifestContent.replace(re, applicationId);
        await fs.writeFileSync(defaults.manifestPath, updatedManifestContent);
        await manifest.modifyManifestFile(defaults.manifestPath, 'random');

    } catch (err) {
        throw new Error(`Unable to update ${defaults.manifestPath}. \n${err}`);
    }
}

exports.updateProjectManifest = updateProjectManifest;