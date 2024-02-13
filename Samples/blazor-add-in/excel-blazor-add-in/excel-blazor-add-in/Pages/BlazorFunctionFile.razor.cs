using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

using System.Diagnostics;

// TODO: TypeScript
// See StickyNote
// TODO: Mix Server Code with Client WASM (new .NET 8 feature)
// TODO: Serverside Components (render client first, don't wait for async components to load)
// https://devblogs.microsoft.com/dotnet/asp-net-core-updates-in-dotnet-8-preview-3/
// Think of a good example of a C# function to use here that makes the use of the Blazor load feasible 
// https://learn.microsoft.com/en-us/aspnet/core/blazor/javascript-interoperability/?view=aspnetcore-8.0
// restore telemetry before PR
// Do we still need CustomFunctionsRuntime in Manifest?

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
