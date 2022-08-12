// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Code for configuring authentication, authorization, and OBO copied from https://github.com/Azure-Samples/active-directory-dotnet-native-aspnetcore-v2/blob/master/2.%20Web%20API%20now%20calls%20Microsoft%20Graph/TodoListService/Startup.cs

using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.AspNetCore.Mvc.Authorization;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.UI;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddAuthentication(OpenIdConnectDefaults.AuthenticationScheme)
    .AddMicrosoftIdentityWebApp(builder.Configuration.GetSection("AzureAd"));
builder.Services.AddSession(options =>
{
    options.Cookie.IsEssential = true;
});

var config = builder.Configuration;

builder.Services.AddMicrosoftIdentityWebApiAuthentication(config)
                    .EnableTokenAcquisitionToCallDownstreamApi()
                        .AddMicrosoftGraph(config.GetSection("DownstreamApi"))
                        .AddInMemoryTokenCaches();

// Add services to the container.
builder.Services.AddControllersWithViews();
builder.Services.AddControllers();



var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseStaticFiles();
app.UseHttpsRedirection();

app.UseSession();

app.UseRouting();

app.UseAuthentication();
app.UseAuthorization();

//app.MapControllerRoute(
//    name: "default",
//    pattern: "{controller=Home}/{action=Index}/{id?}");
app.MapControllers();


app.Run();
