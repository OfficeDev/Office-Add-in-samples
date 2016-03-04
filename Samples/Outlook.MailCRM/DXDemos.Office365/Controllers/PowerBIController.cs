using DXDemos.Office365.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;

namespace DXDemos.Office365.Controllers
{
    public class PowerBIController : ApiController
    {
        [HttpGet]
        public async Task<List<PowerBIDataset>> GetDatasets()
        {
            return await PowerBIModel.GetDatasets();
        }

        [HttpGet]
        public async Task<PowerBIDataset> GetDataset(Guid id)
        {
            return await PowerBIModel.GetDataset(id);
        }

        [HttpPost]
        public async Task<Guid> CreateDataset(PowerBIDataset dataset)
        {
            return await PowerBIModel.CreateDataset(dataset);
        }

        [HttpDelete]
        public async Task<bool> DeleteDataset(Guid id)
        {
            //DELETE IS UNSUPPORTED
            return await PowerBIModel.DeleteDataset(id);
        }

        [HttpPost]
        public async Task<bool> ClearTable(PowerBITableRef tableRef)
        {
            return await PowerBIModel.ClearTable(tableRef.datasetId, tableRef.tableName);
        }

        [HttpPost]
        public async Task<bool> AddTableRows(PowerBITableRows rows)
        {
            return await PowerBIModel.AddTableRows(rows.datasetId, rows.tableName, rows.rows);
        }
    }
}
