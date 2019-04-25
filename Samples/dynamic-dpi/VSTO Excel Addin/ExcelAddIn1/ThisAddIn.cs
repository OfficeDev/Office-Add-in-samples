using System;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools;
using SharedModule;
using System.Windows;

namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {
        public Excel.Application ExcelApp;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            ExcelApp = Globals.ThisAddIn.Application;

			SharedApp.HostApp = Globals.ThisAddIn.Application;
            SharedApp.InitAppTaskPanes(ref this.CustomTaskPanes);
            SharedApp.AppTaskPanes.CreateTaskpaneInstance();

			ExcelApp.WorkbookOpen += new Excel.AppEvents_WorkbookOpenEventHandler(Application_WorkbookOpen);
			//ExcelApp.WorkbookBeforeSave += new Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);
		}

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

		private void Application_WorkbookOpen(Excel.Workbook Wb)
		{
			MessageBox.Show(string.Format("WorkbookOpen {0} with DpiThreadAwarenessContext {1}", Wb.Name, DPIHelper.GetThreadDpiAwarenessContext().ToString()));
		}

		private void Application_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
		{
			MessageBox.Show(string.Format("WorkbookSave {0} with DpiThreadAwarenessContext {1}", Wb.Name, DPIHelper.GetThreadDpiAwarenessContext().ToString()));
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
