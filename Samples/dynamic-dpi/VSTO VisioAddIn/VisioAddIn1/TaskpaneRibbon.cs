using Microsoft.Office.Tools.Ribbon;
using SharedModule;
using Microsoft.Office.Interop.Visio;

namespace VisioAddIn1
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

        private void ribbonAddWindow_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveWindow.Windows.Add("Visio DDPI Window", VisWindowStates.visWSVisible & VisWindowStates.visWSDockedRight, VisWinTypes.visAnchorBarAddon, 0, 0, 300, 210);
        }
    }
}
