/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace BlazorAddIn.Pages
{
    /// <summary>
    /// Starter class to demo how to create a bubble chart.
    /// </summary>
    public partial class BubbleChart
    {
        [Inject]
        public IJSRuntime JSRuntime { get; set; } = default!;
        public IJSObjectReference JSModule { get; set; } = default!;

        protected override async Task OnAfterRenderAsync(bool firstRender)
        {
            if (firstRender)
            {
                JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/BubbleChart.razor.js");
            }
        }

        /// <summary>
        /// Function to create the starter table as source for the bubble chart.
        /// </summary>
        private async Task CreateTable() =>
            await JSModule.InvokeVoidAsync("createTable");

        /// <summary>
        /// Function to create the actual bubble chart.
        /// </summary>
        private async Task CreateBubbleChart() =>
            await JSModule.InvokeVoidAsync("createBubbleChart");

        [JSInvokable]
        public static string SayHelloBubble(string name)
        {
            return $"Hello Bubble, {name} from BlazorFunctionFile!";
        }
    }
}

