// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using SharedModule;
using System.Windows.Forms;

namespace WordAddIn1
{
	public partial class ThisAddIn
	{
		public Microsoft.Office.Interop.Word.Application WordApp;

		private void InitializeCustom()
		{
			WordApp = Globals.ThisAddIn.Application;

			WordApp.DocumentOpen += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentOpenEventHandler(Word_DocumentOpen);
		}

		private void ThisAddIn_Startup(object sender, System.EventArgs e)
		{
			SharedApp.HostApp = Globals.ThisAddIn.Application;
			SharedApp.InitAppTaskPanes(ref this.CustomTaskPanes);
			SharedApp.AppTaskPanes.CreateTaskpaneInstance();
		}

		private void Word_DocumentOpen(Microsoft.Office.Interop.Word.Document Doc)
		{
			MessageBox.Show(string.Format("DocumentOpen {0} with DpiThreadAwarenessContext {1}", Doc.ActiveWindow.Caption, DPIHelper.GetThreadDpiAwarenessContext().ToString()));
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
			InitializeCustom();
			this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
