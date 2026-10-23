/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */

console.log("Loading WorldClock.razor.js");

/**
 * Refreshes the page by triggering a navigation or state update.
 * This can be called from the ribbon to force a refresh of the World Clock page.
 */
export async function refreshPage(): Promise<void> {
    console.log("WorldClock: refreshPage called from ribbon");
    
    // Dispatch a custom event that can be listened to by the Blazor component
    const event = new CustomEvent('worldclock-refresh', { detail: { timestamp: Date.now() } });
    document.dispatchEvent(event);
}

/**
 * Sets up a visibility change listener to detect when the taskpane becomes visible.
 * This helps refresh the page when navigating via ribbon buttons.
 */
export function setupVisibilityListener(): void {
    console.log("WorldClock: Setting up visibility listener");
    
    document.addEventListener('visibilitychange', () => {
        if (document.visibilityState === 'visible') {
            console.log("WorldClock: Page became visible");
            const event = new CustomEvent('worldclock-refresh', { detail: { timestamp: Date.now() } });
            document.dispatchEvent(event);
        }
    });
}

// Initialize the visibility listener when the module loads
setupVisibilityListener();
