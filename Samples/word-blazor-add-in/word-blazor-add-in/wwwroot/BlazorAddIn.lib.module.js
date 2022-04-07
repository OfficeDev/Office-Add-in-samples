export async function beforeStart(wasmOptions, extensions) {
    console.log("We are now entering function: beforeStart");

    Office.onReady((info) => {
        // Check that we loaded into Word
        if (info.host === Office.HostType.Word) {
            console.log("We are now hosting in Word");
        }
        else {
            console.log("We are now hosting in The Browser (of your choice)");
        }
        console.log("Office onReady");
    });
}

export async function afterStarted(blazor) {
    console.log("We are now entering function: afterStarted");
}