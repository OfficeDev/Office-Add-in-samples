using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Cell_Analyzer
{
    public partial class CellAnalyzerPane : UserControl
    {
        public CellAnalyzerPane()
        {
            InitializeComponent();
        }

        private void btnUnicode_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Range rangeCell;
            rangeCell = Globals.ThisAddIn.Application.ActiveCell;

            string cellValue = "";

            if (null != rangeCell.Value)
            {
                cellValue = rangeCell.Value.ToString();
            }

            //convert string to Unicode listing
            string result = "";
            foreach (char c in cellValue)
            {
                int unicode = c;

                result += $"{c}: {unicode}\r\n";
            }
            
            //Output the result
            txtResult.Text = result;
        }

    }

}
