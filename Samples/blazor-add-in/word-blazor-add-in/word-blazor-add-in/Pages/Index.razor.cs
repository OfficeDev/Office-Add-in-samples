/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace BlazorAddIn.Pages
{
    public partial class Index
    {
        [Inject]
        public IJSRuntime JSRuntime { get; set; } = default!;

        public IJSObjectReference JSModule { get; set; } = default!;

        protected override async Task OnAfterRenderAsync(bool firstRender)
        {
            if (firstRender)
            {
                JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/Index.razor.js");
            }
        }

        /// <summary>
        /// Invoke JavaScript function to set up the document (insert a couple of paragraphs).
        /// </summary>
        internal async Task Setup() =>
            await JSModule.InvokeVoidAsync("setupDocument");

        /// <summary>
        /// Invoke JavaScript function to surround the Paragraphs with content controls.
        /// </summary>
        internal async Task InsertContentControls() =>
            await JSModule.InvokeVoidAsync("insertContentControls");

        /// <summary>
        /// Invoke JavaScript function to tag the content controls with "odd" or "even" tags.
        /// </summary>
        internal async Task TagContentControls() =>
            await JSModule.InvokeVoidAsync("tagContentControls");

        /// <summary>
        /// Invoke JavaScript function to modify the content controls and give some decoration colors and info texts.
        /// </summary>
        internal async Task ModifyContentControls() =>
            await JSModule.InvokeVoidAsync("modifyContentControls");

        /// <summary>
        /// Invoke JavaScript function to register events in case the content controls are deleted, or the selection was changed.
        /// </summary>
        internal async Task RegisterContentControlsEvents() =>
            await JSModule.InvokeVoidAsync("registerEvents");

        /// <summary>
        /// Invoke JavaScript function to revoke registration, remove event handlers from the content controls.
        /// </summary>
        internal async Task DeRegisterContentControlEvents() =>
            await JSModule.InvokeVoidAsync("deregisterEvents");

        /// <summary>
        /// Invoke JavaScript function delete the first 'even' content control.
        /// </summary>
        internal async Task DeleteContentControl() =>
            await JSModule.InvokeVoidAsync("deleteContentControl");
    }
}