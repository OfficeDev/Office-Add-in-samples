/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace BlazorAddIn.Pages
{
    /// <summary>
    /// FirstAddin Class
    /// This will demo inserting 'famous' quotes at the current selection, start or end of the document.
    /// </summary>
    public partial class FirstAddin
    {
        [Inject]
        public IJSRuntime JSRuntime { get; set; } = default!;
        public IJSObjectReference JSModule { get; set; } = default!;

        protected override async Task OnAfterRenderAsync(bool firstRender)
        {
            if (firstRender)
            {
                JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/FirstAddin.razor.js");
            }
        }

        /// <summary>
        /// Will invoke JavaScript function to insert Emerson Quote
        /// </summary>
        private async Task InsertWaldo() =>
            await JSModule.InvokeVoidAsync("insertEmersonQuoteAtSelection");

        /// <summary>
        /// Will invoke JavaScript function to insert Chechov Quote
        /// </summary>
        private async Task InsertChekhov() =>
            await JSModule.InvokeVoidAsync("insertChekhovQuoteAtTheBeginning");

        /// <summary>
        /// Will invoke JavaScript function to insert Chinese Proverb
        /// </summary>
        private async Task InsertProverb() =>
            await JSModule.InvokeVoidAsync("insertChineseProverbAtTheEnd");
    }
}