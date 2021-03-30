// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const hard_reset_button = document.getElementById ('hard_reset_button');

function sig_check()    {
    if(sessionStorage.getItem('flag') === 'true')  {
        console.log("flag true");
    }
    else    {
        Office.onReady(() => {
            var value = window.Office.context.roamingSettings.get('myKey');
            if(value)    {
                console.log("Received value from roaming settings" + value);
                localStorage.setItem('store_setter', JSON.stringify(value.store_obj));
                console.log(JSON.parse(localStorage.getItem ('store_setter')));
                value = null;
                location.href = 'signature.html';
            }
            else    {
                location.href = 'no_signature.html';
            }
        })
    }
}

hard_reset_button.onclick = function() {
    Office.onReady(() => {
        Office.context.roamingSettings.remove('myKey');
        Office.context.roamingSettings.saveAsync(() => {
            console.log("Removed successfully from roaming settings : inside");
        }); 
        if(!Office.context.roamingSettings.get('myKey'))    {
            console.log("Removed successfully from roaming settings");
        }
        localStorage.removeItem('store_setter');
        sessionStorage.setItem('flag', 'true');
        location.href = 'no_signature.html';
    })
}