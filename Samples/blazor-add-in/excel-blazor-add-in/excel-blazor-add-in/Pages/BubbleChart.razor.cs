/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;

namespace BlazorAddIn.Pages
{
    /// <summary>
    /// Starter class to demo how to create a bubble chart.
    /// </summary>
    [SupportedOSPlatform("browser")]
    public partial class BubbleChart
    {
        [Inject, AllowNull]
        private IJSRuntime JSRuntime { get; set; }
        private IJSObjectReference? JSModule { get; set; }

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

        [JSImport("createTable", "BubbleChart")]
        internal static partial void RunCreateTable();

        [JSImport("createBubbleChart", "BubbleChart")]
        internal static partial Task RunCreateChart();

        [JSImport("createTable", "BubbleChart")]
        internal static partial Task CreateImportedTable();

        [JSInvokable]
        public static async Task CreateBubbles()
        {
            await JSHost.ImportAsync("BubbleChart", "../Pages/BubbleChart.razor.js");
            await CreateImportedTable();
            await RunCreateChart();
        }

        [JSInvokable]
        public static string SayHelloBubble(string name)
        {
            return $"Hello Bubble, {name} from BlazorFunctionFile!";
        }
    }
}

