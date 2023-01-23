using DocumentFormat.OpenXml.Office.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FunctionApp1
{

        public class Product
        {
            public int ID { get; set; }
            public string Name { get; set; }
            public int Qtr1 { get; set; }
            public int Qtr2 { get; set; }
            public int Qtr3 { get; set; }
            public int Qtr4 { get; set; }
        }
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
