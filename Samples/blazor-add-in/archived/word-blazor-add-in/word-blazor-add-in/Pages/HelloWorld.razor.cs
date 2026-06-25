/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace BlazorAddIn.Pages
{
    /// <summary>
    /// Starter class to demo how to insert a paragraph
    /// </summary>
    public partial class HelloWorld
    {
        [Inject]
        public IJSRuntime JSRuntime { get; set; } = default!;
        public IJSObjectReference JSModule { get; set; } = default!;

        protected override async Task OnAfterRenderAsync(bool firstRender)
        {
            if (firstRender)
            {
                JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/HelloWorld.razor.js");
            }
        }

        /// <summary>
        /// Basic function to invoke inserting a paragraph
        /// </summary>
        private async Task InsertParagraph() =>
            await JSModule.InvokeVoidAsync("insertParagraph");
    }
}