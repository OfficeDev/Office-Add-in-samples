using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace Cell_Analyzer
{
    public partial class ThisAddIn
    {
        private CellAnalyzerPane control;
        private Microsoft.Office.Tools.CustomTaskPane taskpane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.control = new CellAnalyzerPane();
            this.taskpane = this.CustomTaskPanes.Add(this.control, "Cell Analyzer");
            this.taskpane.Visible = true;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
