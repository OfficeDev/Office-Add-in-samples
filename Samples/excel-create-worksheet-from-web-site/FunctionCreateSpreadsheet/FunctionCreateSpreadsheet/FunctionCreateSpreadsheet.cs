// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Net.Http;

namespace FunctionApp1
{
    public static class FunctionCreateSpreadsheet
    {

        [FunctionName("FunctionCreateSpreadsheet")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");
            string responseMessage = "";            
            try
            {
                string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
                TableData tableData = JsonConvert.DeserializeObject<TableData>(requestBody);

                SpreadsheetBuilder s = new SpreadsheetBuilder();
                var spreadsheetBytes = s.CreateSpreadsheet("Web Data", tableData);
                responseMessage = Convert.ToBase64String(spreadsheetBytes);

                return new FileContentResult(spreadsheetBytes, "application/octet-stream");
            }
            catch (Exception ex)
            {
                log.LogError("Error creating spreadsheet. Inner error: " + ex.Message);
            }

            return new OkObjectResult(responseMessage);
        }
    }
}
