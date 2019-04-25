using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using SharedModule;

namespace WordAddIn1
{
    public partial class TaskPaneRibbon
    {
		public Microsoft.Office.Interop.Word.Application WordApp;

		private void TaskPaneRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonCloseAllTaskpanes_Click(object sender, RibbonControlEventArgs e)
        {
            SharedApp.AppTaskPanes.CloseAllTaskpanes();
        }

        private void buttonAddTaskpane_Click(object sender, RibbonControlEventArgs e)
        {
            SharedApp.AppTaskPanes.CreateTaskpaneInstance();
        }

		private void btnOpenHelpNewProcess_Click(object sender, RibbonControlEventArgs e)
		{
			SharedApp.View_Help(false);
		}
	}
}
