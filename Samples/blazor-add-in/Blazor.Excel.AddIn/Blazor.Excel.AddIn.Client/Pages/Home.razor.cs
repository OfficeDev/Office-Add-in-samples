/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;

using Blazor.Excel.AddIn.Client.Services;

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace Blazor.Excel.AddIn.Client.Pages;

/// <summary>
/// Starter class to demo how to insert a paragraph
/// </summary>
[SupportedOSPlatform("browser")]
public partial class Home : ComponentBase
{
    private bool _hostInformation;

    protected override async Task OnAfterRenderAsync(bool firstRender)
    {
        if (firstRender)
        {
            try
            {
                await JSHost.ImportAsync("Home", "../Pages/Home.razor.js");
                Console.WriteLine($"Imported Home module");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error importing Home module: {ex.Message}");
            }

            _hostInformation = await OfficeUtilities.IsRunningInHostAsync();
            Console.WriteLine($"Home HostInformation: {_hostInformation}");

            if (_hostInformation)
            {
                StateHasChanged();
            }
        }
    }

    /// <summary>
    /// Basic function to invoke inserting 'Hello world!' text.
    /// </summary>
    [JSImport("insertText", "Home")]
    internal static partial Task InsertText();

    [JSInvokable]
    public static Task<string> SayHelloHome(string name)
    {
        Console.WriteLine("Invoking SayHelloHome");
        return Task.FromResult($"Hello, {name} from Home Page!");
    }
}