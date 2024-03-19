import * as msalBrowser from "@azure/msal-browser";
import getGraphData from "../msgraph-helpers/msgraph-helper";

//msal config - using dev tenant registered app ID. 
const msalConfig = {
    auth: {
        clientId: "57e00eca-d992-4e1c-bef6-a238cd0236c4",
        authority: "https://login.microsoftonline.com/common",
        supportsNestedAppAuth: true
    }
}

const myloginhint = "davechuatest3.onmicrosoft.com"

let pca = undefined;
msalBrowser.PublicClientNext.createPublicClientApplication(msalConfig).then((result) => {
    pca = result;
    ssoGetToken();
});

export default async function ssoGetToken() {
    //const activeAccount = pca.getActiveAccount();  
    const tokenRequest = {
        scopes: ["User.Read", "Files.Read", "openid", "profile"],
        loginhint: myloginhint
    };

    try {
        const result = await pca.ssoSilent(tokenRequest);
        console.log(result);
        const requestString = "https://graph.microsoft.com/v1.0/me";
        const headersInit = { 'Authorization': result.accessToken };
        const requestInit = { 'headers': headersInit }
        // if (requestString !== undefined) {
        //     const result = await fetch(requestString, requestInit);
        //     if (result.ok) {
        //         const data = await result.text();
        //         console.log(data);

        //         //document.getElementById("userInfo").innerText = data;
        //     } else {
        //         //Handle whatever errors could happen that have nothing to do with identity
        //         console.log(result);
        //     }
        // } else {
        //     //throw this should never happen
        //     throw new Error("unexpected: no requestString");
        // }
        return result.accessToken;
    } catch (error) {
        console.log(error);
        let resultatpu = pca.acquireTokenPopup(tokenRequest);
        console.log("result: " + resultatpu);
        throw (error); //rethrow
    }
}


async function getuserfilenames() {
    try {
        const accessToken = ssoGetToken();
        // Minimize the data that must come from MS Graph by specifying only the property we need ("name")
        // and only the top 10 folder or file names.
        const rootUrl = '/me/drive/root/children';

        // Note that the last parameter, for queryParamsSegment, is hardcoded. If you reuse this code in
        // a production add-in and any part of queryParamsSegment comes from user input, be sure that it is
        // sanitized so that it cannot be used in a Response header injection attack.
        const params = '?$select=name&$top=10';

        const graphData = await getGraphData(
            accessToken,
            rootUrl,
            params
        );

        // If Microsoft Graph returns an error, such as invalid or expired token,
        // there will be a code property in the returned object set to a HTTP status (e.g. 401).
        // Return it to the client. On client side it will get handled in the fail callback of `makeWebServerApiCall`.
        if (graphData.code) {
            throw new Error("An error occurred while calling the Microsoft Graph API.");
        } else {
            // MS Graph data includes OData metadata and eTags that we don't need.
            // Send only what is actually needed to the client: the item names.
            const itemNames = [];
            const oneDriveItems = graphData["value"];
            for (let item of oneDriveItems) {
                itemNames.push(item["name"]);
            }
            console.log(itemNames);
        }
    } catch (err) {
        throw (err); //rethrow
    }
}
