using SharedModule;
using System.Windows.Forms;

namespace OutlookAddIn1
{
	public partial class ThisAddIn
    {
		public Microsoft.Office.Interop.Outlook.Application OutlookApp;

		private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
			OutlookApp = Globals.ThisAddIn.Application;

			SharedApp.HostApp = Globals.ThisAddIn.Application;
			SharedApp.InitAppTaskPanes(ref this.CustomTaskPanes);
            SharedApp.AppTaskPanes.CreateTaskpaneInstance();

			OutlookApp.ItemSend += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ItemSendEventHandler(Outlook_ItemSend);
		}

		private void Outlook_ItemSend(object Item, ref bool Cancel)
		{

			Microsoft.Office.Interop.Outlook.MailItem mail = Item as Microsoft.Office.Interop.Outlook.MailItem;
			if (mail != null)
			{
				MessageBox.Show(string.Format("ItemSend {0} with DpiThreadAwarenessContext {1}", mail.Subject, DPIHelper.GetThreadDpiAwarenessContext().ToString()));
			}
		}

		private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
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
