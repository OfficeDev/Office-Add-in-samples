using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace BlazorAddIn.Pages
{
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

            // ... calling InsertParagraph
        }

        private async Task Clear()
        {
            await JSModule.InvokeVoidAsync("clearDocument");
        }


        private async Task RequestContext(string text, string location)
        {
            await JSModule.InvokeVoidAsync("requestContextDemo", text, location);
        }
    }
}