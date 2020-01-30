using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace CellAnalyzerRESTAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class AnalyzeUnicodeController : ControllerBase
    {
        [HttpGet]
        public ActionResult<string> AnalyzeUnicode(string value)
        {
            if (value == null)
            {
                return BadRequest();
            }
            return CellAnalyzerSharedLibrary.CellOperations.GetUnicodeFromText(value);
        }
    }
}