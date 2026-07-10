/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
using Blazor.PowerPoint.AddIn.Client.Services;

using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using Microsoft.FluentUI.AspNetCore.Components;

var builder = WebAssemblyHostBuilder.CreateDefault(args);

builder.Services.AddFluentUIComponents();
builder.Services.AddScoped<ClientCommandHandler>();

await builder.Build().RunAsync();