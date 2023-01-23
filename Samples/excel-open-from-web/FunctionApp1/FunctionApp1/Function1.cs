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
    public static class Function1
    {

        [FunctionName("Function1")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string name = req.Query["name"];

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            TableData tableData = JsonConvert.DeserializeObject<TableData>(requestBody);

            //dynamic data = JsonConvert.DeserializeObject(requestBody);
            //name = name ?? data?.name;

            //          string responseMessage = string.IsNullOrEmpty(name)
            //            ? "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response."
            //          : $"Hello, {name}. This HTTP triggered function executed successfully.";


            string responseMessage = "Test";
            //try
            //{
            //    var products = new List<Product>()
            //{ new Product {ID=1, Name="Frames", Qtr1=5000, Qtr2=7000, Qtr3=6544, Qtr4=4377},
            //new Product {ID=2, Name="Saddles", Qtr1=400, Qtr2=323, Qtr3=276, Qtr4=651},
            //new Product {ID=3, Name="Brake levers", Qtr1=12000, Qtr2=8766, Qtr3=8456, Qtr4=9812},
            //new Product {ID=4, Name="Chains", Qtr1=1550, Qtr2=1088, Qtr3=692, Qtr4=853},
            //new Product {ID=5, Name="Mirrors", Qtr1=225, Qtr2=600, Qtr3=923, Qtr4=544},
            //new Product {ID=5, Name="Spokes", Qtr1=6005, Qtr2=7634, Qtr3=4589, Qtr4=8765}
            //};
                // Build the spreadsheet
                SpreadsheetBuilder s = new SpreadsheetBuilder();


            //var spreadsheetBytes = s.CreateSpreadsheet("ProductSales", products);
            try { 
            var spreadsheetBytes = s.CreateSpreadsheet("ProductSales", tableData);
                responseMessage = Convert.ToBase64String(spreadsheetBytes);
                return new FileContentResult(spreadsheetBytes, "application/octet-stream");
                
            }
            catch (Exception ex) {
                Console.WriteLine("oh no!!!");
                }
        

            return new OkObjectResult(responseMessage);
        }
    }
}
