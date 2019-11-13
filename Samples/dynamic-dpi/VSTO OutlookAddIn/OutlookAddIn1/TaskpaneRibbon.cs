using Microsoft.Office.Tools.Ribbon;
using SharedModule;

namespace OutlookAddIn1
{
    public partial class TaskpaneRibbon
    {
        private void TaskpaneRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonAddTaskpane_Click(object sender, RibbonControlEventArgs e)
        {
            SharedApp.AppTaskPanes.CreateTaskpaneInstance();
        }

        private void buttonCloseAllTaskpanes_Click(object sender, RibbonControlEventArgs e)
        {
            SharedApp.AppTaskPanes.CloseAllTaskpanes();
        }
    }
}
