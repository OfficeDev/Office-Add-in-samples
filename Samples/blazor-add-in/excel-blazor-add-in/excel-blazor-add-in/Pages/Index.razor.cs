/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace BlazorAddIn.Pages
{
    /// <summary>
    /// Starter class to demo how to insert `Hello world!` text.
    /// </summary>
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
        /// Basic function to invoke inserting `Hello world!` text.
        /// </summary>
        private async Task HelloButton() =>
            await JSModule.InvokeVoidAsync("helloButton");
    }
}
