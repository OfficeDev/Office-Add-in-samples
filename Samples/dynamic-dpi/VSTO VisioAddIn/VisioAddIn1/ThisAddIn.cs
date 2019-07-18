// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using SharedModule;
using System.Windows.Forms;

namespace VisioAddIn1
{
    public partial class ThisAddIn
    {
		public Microsoft.Office.Interop.Visio.Application VisioApp;
		private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
			VisioApp = Globals.ThisAddIn.Application;

			SharedApp.HostApp = Globals.ThisAddIn.Application;
			// Custom taskpanes not available in Visio
			// SharedApp.InitAppTaskPanes(ref this.CustomTaskPanes);
			// SharedApp.AppTaskPanes.CreateTaskpaneInstance();

			VisioApp.DocumentOpened += new Microsoft.Office.Interop.Visio.EApplication_DocumentOpenedEventHandler(Visio_DocOpened);
		}

		private void Visio_DocOpened(Microsoft.Office.Interop.Visio.Document doc)
		{
			MessageBox.Show(string.Format("DocOpened {0} with DpiThreadAwarenessContext {1}", doc.Name, DPIHelper.GetThreadDpiAwarenessContext().ToString()));
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
