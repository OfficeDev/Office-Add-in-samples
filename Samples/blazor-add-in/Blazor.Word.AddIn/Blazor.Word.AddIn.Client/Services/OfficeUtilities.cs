/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;

namespace Blazor.Word.AddIn.Client.Services;

/// <summary>
/// Provides utility methods for interacting with Microsoft Office JavaScript APIs
/// from a Blazor WebAssembly application.
/// </summary>
/// <remarks>
/// This class serves as a bridge between .NET and JavaScript, enabling Office add-in
/// functionality within a Blazor context. It manages the lifecycle of JavaScript module
/// imports and provides type-safe access to Office host detection capabilities.
/// </remarks>
[SupportedOSPlatform("browser")]
public static partial class OfficeUtilities
{
    /// <summary>
    /// Indicates whether the SharedUtils JavaScript module has been imported.
    /// </summary>
    private static bool _isImported = false;

    /// <summary>
    /// Ensures the SharedUtils JavaScript module is imported before use.
    /// </summary>
    /// <remarks>
    /// This method uses a guard pattern to prevent multiple imports of the same module.
    /// The import is performed only once during the lifetime of the application.
    /// </remarks>
    /// <returns>A task that represents the asynchronous import operation.</returns>
    /// <exception cref="Exception">Thrown when the JavaScript module import fails.</exception>
    public static async Task EnsureImportedAsync()
    {
        if (!_isImported)
        {
            try
            {
                await JSHost.ImportAsync("SharedUtils", "/scripts/SharedUtils.js");
                Console.WriteLine("Imported SharedUtils module");
                _isImported = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error importing SharedUtils module: {ex.Message}");
                throw;
            }
        }
    }

    /// <summary>
    /// Checks if the add-in is running in Microsoft Word host application.
    /// </summary>
    /// <remarks>
    /// This method ensures the required JavaScript module is loaded before performing
    /// the host detection check. It delegates to the JavaScript IsRunningInHost function
    /// to determine the runtime environment.
    /// </remarks>
    /// <returns>
    /// A task that resolves to <see langword="true"/> if running in Word; 
    /// otherwise, <see langword="false"/>.
    /// </returns>
    public static async Task<bool> IsRunningInHostAsync()
    {
        await EnsureImportedAsync();
        return await IsRunningInHostInternal();
    }

    /// <summary>
    /// Internal JavaScript interop method that invokes the IsRunningInHost function
    /// from the SharedUtils module.
    /// </summary>
    /// <returns>
    /// A task that resolves to <see langword="true"/> if the add-in is running 
    /// in the Office host; otherwise, <see langword="false"/>.
    /// </returns>
    [JSImport("IsRunningInHost", "SharedUtils")]
    private static partial Task<bool> IsRunningInHostInternal();
}