/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace BlazorAddIn.Pages
{
    /// <summary>
    /// Quotes Class
    /// This will demo inserting 'famous' quotes at the current selection, start or end of the document.
    /// </summary>
    public partial class Quotes
    {
        [Inject]
        public IJSRuntime JSRuntime { get; set; } = default!;
        public IJSObjectReference JSModule { get; set; } = default!;

        protected override async Task OnAfterRenderAsync(bool firstRender)
        {
            if (firstRender)
            {
                JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/Quotes.razor.js");
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
        private async Task InsertBruceSchneier() =>
            await JSModule.InvokeVoidAsync("insertBruceSchneierQuoteAtTheBeginning");

        /// <summary>
        /// Will invoke JavaScript function to insert Steve Ballmer Quote
        /// </summary>
        private async Task InsertBallmer() =>
            await JSModule.InvokeVoidAsync("insertSteveBallmerAtTheEnd");
    }
}