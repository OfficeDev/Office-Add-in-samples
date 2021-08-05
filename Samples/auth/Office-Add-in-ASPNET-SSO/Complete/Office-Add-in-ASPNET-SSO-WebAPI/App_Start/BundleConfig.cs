// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
using System.Web.Optimization;

namespace Office_Add_in_ASPNET_SSO_WebAPI
{
    public class BundleConfig
    {
        // For more information on bundling, visit http://go.microsoft.com/fwlink/?LinkId=301862
        public static void RegisterBundles(BundleCollection bundles)
        {
            bundles.Add(new StyleBundle("~/Content/css").Include(
                      "~/Content/Site.css"));

            bundles.Add(new ScriptBundle("~/bundles/OfficeAddinMicrosoftGraphASPNET")
                .Include("~/Scripts/Home.js")
                .Include("~/Scripts/Document.js")
                );

        }
    }
}
