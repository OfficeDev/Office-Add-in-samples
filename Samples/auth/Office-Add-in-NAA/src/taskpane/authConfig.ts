import * as msalBrowser from "@azure/msal-browser";
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
        scopes: ["User.Read", "openid", "profile"],
        loginhint: myloginhint
    };

    try {
        const result = await pca.ssoSilent(tokenRequest);
        console.log(result);
        const requestString = "https://graph.microsoft.com/v1.0/me";
        const headersInit = { 'Authorization': result.accessToken };
        const requestInit = { 'headers': headersInit }
        if (requestString !== undefined) {
            const result = await fetch(requestString, requestInit);
            if (result.ok) {
                const data = await result.text();
                console.log(data);
                //document.getElementById("userInfo").innerText = data;
            } else {
                //Handle whatever errors could happen that have nothing to do with identity
                console.log(result);
            }
        } else {
            //throw this should never happen
            throw new Error("unexpected: no requestString");
        }
    } catch (error) {
        console.log(error);
        let resultatpu = pca.acquireTokenPopup(tokenRequest);
        console.log("result: " + resultatpu);

    }
}
