using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace BlazorAddIn.Pages
{
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

        internal async Task Setup() =>
            await JSModule.InvokeVoidAsync("setupDocument");

        internal async Task InsertContentControls() =>
            await JSModule.InvokeVoidAsync("insertContentControls");

        internal async Task TagContentControls() =>
            await JSModule.InvokeVoidAsync("tagContentControls");

        internal async Task ModifyContentControls() =>
            await JSModule.InvokeVoidAsync("modifyContentControls");
    }
}