// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Web.Mvc;

namespace AttachmentDemoWeb
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
