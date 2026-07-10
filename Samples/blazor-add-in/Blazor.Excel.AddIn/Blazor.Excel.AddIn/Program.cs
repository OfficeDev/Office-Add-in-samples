/* Copyright (c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
using Blazor.Excel.AddIn.Client.Services;
using Blazor.Excel.AddIn.Components;

using Microsoft.FluentUI.AspNetCore.Components;
using Microsoft.Extensions.FileProviders;
using System.IO;

var builder = WebApplication.CreateBuilder(args);
builder.WebHost.UseStaticWebAssets();

// Add services to the container.
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents()
    .AddInteractiveWebAssemblyComponents();

builder.Services.AddFluentUIComponents();
builder.Services.AddScoped<ServerCommandHandler>();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
    app.UseWebAssemblyDebugging();
    // Dev-time static file fallback: serve generated TypeScript outputs from client wwwroot
    var clientWwwroot = Path.Combine(app.Environment.ContentRootPath, "..", "Blazor.Excel.AddIn.Client", "wwwroot");
    if (Directory.Exists(clientWwwroot))
    {
        app.UseStaticFiles(new StaticFileOptions
        {
            FileProvider = new PhysicalFileProvider(clientWwwroot),
            RequestPath = "/generated-assets"
        });
    }
}
else
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseStatusCodePagesWithReExecute("/not-found", createScopeForStatusCodePages: true);
app.UseHttpsRedirection();

// Fallback for files not in the build-time manifest (e.g. TypeScript-compiled outputs)
app.UseStaticFiles();
app.UseAntiforgery();

// Serves fingerprinted/cached assets from the build-time static web assets manifest.
app.MapStaticAssets();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode()
    .AddInteractiveWebAssemblyRenderMode()
    .AddAdditionalAssemblies(typeof(Blazor.Excel.AddIn.Client._Imports).Assembly);

app.Run();
