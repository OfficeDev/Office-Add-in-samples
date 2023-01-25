using DocumentFormat.OpenXml.Office.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FunctionApp1
{
    public class TableData
    {
        public RowData[] rows { get; set; }
    }
    public class RowData
    {
        public ColumnData[] columns { get; set; }
    }
    public class ColumnData
    {
        public string Value { get; set; }
    }
}
