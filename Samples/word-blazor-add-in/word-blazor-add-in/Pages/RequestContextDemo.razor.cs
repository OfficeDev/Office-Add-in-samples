/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace BlazorAddIn.Pages
{
    /// <summary>
    /// This class shows the use of RequestContext instead of using the Word.run option.
    /// Word.run is usually prefered to keep track of objects, but you can track it yourself if you like.
    /// </summary>
    public partial class RequestContextDemo
    {
        [Inject]
        public IJSRuntime JSRuntime { get; set; } = default!;

        public IJSObjectReference JSModule { get; set; } = default!;

        public string ParagraphCount { get; set; } = "<undefined>";

        protected override async Task OnAfterRenderAsync(bool firstRender)
        {
            if (firstRender)
            {
                JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/RequestContextDemo.razor.js");
            }
        }

        private async Task Setup()
        {
            await Clear();
            await RequestContext("One more paragraph.", "Start");

            // ... calling InsertParagraph, room to elaborate the use of the requested context.
        }

        private async Task Clear()
        {
            await JSModule.InvokeVoidAsync("clearDocument");
        }

        /// <summary>
        /// Triggers the Request Context demo
        /// </summary>
        /// <param name="text"></param>
        /// <param name="location"></param>
        private async Task RequestContext(string text, string location)
        {
            await JSModule.InvokeVoidAsync("requestContextDemo", text, location);
        }
    }
}