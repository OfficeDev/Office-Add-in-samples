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

            //Output the result
            txtResult.Text = CellAnalyzerSharedLibrary.CellOperations.GetUnicodeFromText(cellValue);
        }
        private int spacecounter(string value)
        {
            int spaceCount = 0;
            foreach (char c in value)
            {
                if (c == ' ') spaceCount += 1;
            }
            return spaceCount;
        }

    }

}
