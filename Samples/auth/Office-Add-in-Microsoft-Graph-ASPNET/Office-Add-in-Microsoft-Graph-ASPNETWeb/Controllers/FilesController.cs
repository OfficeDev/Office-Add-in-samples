// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Web.Mvc;
using System.Net.Http;
using System.Web.Http;
using System.Net;
using OfficeAddinMicrosoftGraphASPNET.Helpers;
using OfficeAddinMicrosoftGraphASPNET.Models;


namespace OfficeAddinMicrosoftGraphASPNET.Controllers
{
    public class FilesController : Controller
    {
        /// <summary>
        /// Recursively searches OneDrive for Business.
        /// </summary>
        /// <returns>The names of the first three workbooks in OneDrive for Business.</returns>
        public async Task<JsonResult> OneDriveFiles()
        {
            // Get access token
            var token = Data.GetUserSessionToken(Settings.GetUserAuthStateId(ControllerContext.HttpContext), Settings.AzureADAuthority);

            // Get all the Excel files in OneDrive for Business by using the Microsoft Graph API. Select only properties needed.
            var fullWorkbooksSearchUrl = GraphApiHelper.GetWorkbookSearchUrl("?$select=name,id&top=3");
            var filesResult = await ODataHelper.GetItems<ExcelWorkbook>(fullWorkbooksSearchUrl, token.AccessToken);

            List<string> fileNames = new List<string>();
            foreach(ExcelWorkbook workbook in filesResult)
            {
                fileNames.Add(workbook.Name);
            }
            return Json(fileNames, JsonRequestBehavior.AllowGet); 
        }

        public ActionResult Index()
        {
            return View();
        }
    }
}
