using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using SharedModule;

namespace SharedModule
{
    public partial class TaskPaneRibbon
    {
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
    }
}
