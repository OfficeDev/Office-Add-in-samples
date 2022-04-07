using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace BlazorAddIn.Pages
{
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

        private async Task InsertWaldo() =>
            await JSModule.InvokeVoidAsync("insertEmersonQuoteAtSelection");
        private async Task InsertChekhov() =>
            await JSModule.InvokeVoidAsync("insertChekhovQuoteAtTheBeginning");
        private async Task InsertProverb() =>
            await JSModule.InvokeVoidAsync("insertChineseProverbAtTheEnd");
    }
}