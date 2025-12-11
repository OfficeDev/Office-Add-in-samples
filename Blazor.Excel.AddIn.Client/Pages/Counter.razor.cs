/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace Blazor.Excel.AddIn.Client.Pages;

public partial class Counter : ComponentBase
{
    private int currentCount = 0;

    private void IncrementCount()
    {
        currentCount++;
    }

    [JSInvokable]
    public static Task<string> SayHelloCounter(string name)
    {
        return Task.FromResult($"Hello, {name} from Counter Page!");
    }
}
