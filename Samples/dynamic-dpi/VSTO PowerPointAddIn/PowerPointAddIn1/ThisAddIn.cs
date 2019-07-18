// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using SharedModule;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    public partial class ThisAddIn
    {
		public Microsoft.Office.Interop.PowerPoint.Application PowerPointApp;

		private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
			PowerPointApp = Globals.ThisAddIn.Application;

			SharedApp.HostApp = Globals.ThisAddIn.Application;
			SharedApp.InitAppTaskPanes(ref this.CustomTaskPanes);
            SharedApp.AppTaskPanes.CreateTaskpaneInstance();

			PowerPointApp.AfterPresentationOpen += new Microsoft.Office.Interop.PowerPoint.EApplication_AfterPresentationOpenEventHandler(PowerPoint_AfterOpen);
		}

		private void PowerPoint_AfterOpen(Microsoft.Office.Interop.PowerPoint.Presentation pres)
		{
			MessageBox.Show(string.Format("PresentationAfterOpen {0} with DpiThreadAwarenessContext {1}", pres.Name, DPIHelper.GetThreadDpiAwarenessContext().ToString()));
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
