/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */

/**
 * Determines whether the add-in is currently running within the Microsoft Word host application.
 * Initializes the Office JavaScript API and verifies the host type.
 * @returns {Promise<boolean>} A promise that resolves to true if the add-in is running in Word, false if running in a different Office host.
 * @throws {Error} Throws an error if Office initialization fails.
 */
export async function IsRunningInHost(): Promise<boolean> {
    console.log("We are now entering function: IsRunningInHost");

    try {
        const info = await Office.onReady();
        
        // Check that we loaded into Word.
        if (info.host === Office.HostType.Word) {
            console.log("We are hosting in Word.");
            console.log("Office onReady completed.");
            return true;
        }
        else {
            console.log("We are not hosting in Word.");
            console.log("Office onReady completed.");
            return false;
        }
    } catch (error) {
        console.error("Error initializing Office:", error);
        throw error;
    }
}
