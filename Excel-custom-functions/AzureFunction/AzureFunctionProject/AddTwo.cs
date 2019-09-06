using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace Contoso.Functions
{
    public static class AddTwo
    {
        [FunctionName("AddTwo")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            //retrieve parameters if passed on URL. They are passed in string format (convert them later)
            string first = req.Query["first"];
            string second = req.Query["second"];

            //Check if parameters were passed in body JSON.
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            first = first ?? data?.first; 
            second = second ?? data?.second;

            //convert strings to numbers
            //return an error if they were not numbers
            int n1,n2;
            if (!int.TryParse(first,out n1)||!int.TryParse(second,out n2))
            {
                 return new BadRequestObjectResult("Please pass two number parameters in the query string or in the request body");
            }

            //add and return the result as JSON
            return new OkObjectResult("{ \"answer\": "+(n1+n2).ToString()+"}");

        }
    }
}
