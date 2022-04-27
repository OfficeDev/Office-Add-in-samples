/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace BlazorAddIn.Pages
{
    /// <summary>
    /// ReturnValue Class
    /// </summary>
    public class ReturnValue
    {
        public int Value { get; set; } = -1;
    }

    /// <summary>
    /// This demo will show how to be able to perform very small steps in the OfficeJS model
    /// </summary>
    public partial class InteractWithParagraphs
    {
        [Inject]
        public IJSRuntime JSRuntime { get; set; } = default!;

        public IJSObjectReference JSModule { get; set; } = default!;

        public ReturnValue ReturnValue { get; set; } = new ReturnValue();

        protected override async Task OnAfterRenderAsync(bool firstRender)
        {
            Console.WriteLine("OnAfterRenderAsync");

            if (firstRender)
            {
                JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/InteractWithParagraphs.razor.js");
            }

            if (ReturnValue.Value == -1)
            {
                await CountParagraphs();
                StateHasChanged();
            }
        }

        /// <summary>
        /// InsertParagraphs performs a lot of small steps in one run 
        /// </summary>
        /// <returns></returns>
        internal async Task InsertParagraps()
        {
            // VSTO Way of Working
            // Get ActiveDocument
            // Clear ActiveDocument (Get Document Body, Select All, Delete?)

            await Clear();

            // Get ActiveDocument
            // ActiveDocument (Get Document Start, Insert Paragraph at specified location?)

            await InsertParagraph("One more paragraph.", "Start");
            await InsertParagraph("Co-locating Index.razor.js Demo.", "Start");
            await InsertParagraph("Inserting another paragraph.", "Start");
            await InsertParagraph("Effective add-ins offer unique and compelling functionality that extends Office applications in a visually appealing way. To create a great add-in, provide an engaging first-time experience for your users, design a first-class UI experience, and optimize your add-in's performance.", "Start");
            await ReplaceParagraph("With Word add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to build a solution that can run in Word across multiple platforms, including on the web, Windows, Mac, and iPad. Learn how to build, test, debug, and publish Word add-ins.");
            await CountParagraphs();
        }

        /// <summary>
        /// Calls the JavaScript interop to insert Paragraphs at the given location
        /// </summary>
        /// <param name="text"></param>
        /// <param name="location"></param>
        private async Task InsertParagraph(string text, string location)
        {
            await JSModule.InvokeVoidAsync("insertParagraph", text, location);
        }

        /// <summary>
        /// Calls the JavaScript interop to replace the Paragraph text
        /// </summary>
        /// <param name="text"></param>
        private async Task ReplaceParagraph(string text)
        {
            await JSModule.InvokeVoidAsync("replaceParagraph", text);
        }

        private async Task Clear()
        {
            await JSModule.InvokeVoidAsync("clearDocument");
        }

        /// <summary>
        /// Calls the interop to count the paragraphs
        /// </summary>
        private async Task CountParagraphs()
        {
            this.ReturnValue = await JSModule.InvokeAsync<ReturnValue>("paragraphCount");

            Console.WriteLine("Paragraph Count C#: ");
            Console.WriteLine(ReturnValue.Value);
        }
    }
}