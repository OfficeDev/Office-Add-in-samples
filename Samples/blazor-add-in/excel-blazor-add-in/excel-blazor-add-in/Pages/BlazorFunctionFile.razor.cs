using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

using System.Diagnostics;


// TODO: .NET 8.0
// TODO: TypeScript
// Get Functions from Blazor Page
// Try Code by One1Lion 
// I follow this pattern: https://www.meziantou.net/javascript-isolation-in-blazor-components.htm
// See StickyNote
// TODO: Mix Server Code with Client WASM (new .NET 8 feature)
// TODO: Serverside Components (render client first, don't wait for async components to load)
// https://devblogs.microsoft.com/dotnet/asp-net-core-updates-in-dotnet-8-preview-3/

namespace BlazorAddIn.Pages
{
    public partial class BlazorFunctionFile
    {
        [Inject]
        public IJSRuntime JSRuntime { get; set; } = default!;
        public IJSObjectReference JSModule { get; set; } = default!;

        protected override async Task OnAfterRenderAsync(bool firstRender)
        {
            if (firstRender)
            {
                Debug.WriteLine("Hit OnAfterRenderAsync BlazorFunctionFile!");
                Console.WriteLine("Hit OnAfterRenderAsync BlazorFunctionFile in Console!");
                JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/BlazorFunctionFile.razor.js");
            }
        }

        //[JSInvokable]
        //public static string LocalStaticMethod() => $"{nameof(LocalStaticMethod)} called";

        //[JSInvokable]
        //public static Task<int[]> Localfunction()
        //{
        //    return Task.FromResult(new int[] { 1, 2, 3 });
        //}

        [JSInvokable]
        public static string SayHello(string name)
        {
            return $"Hello, {name} from BlazorFunctionFile!";
        }
    }
}
