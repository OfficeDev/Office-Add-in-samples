using Microsoft.Office.Tools;
using System.Windows.Forms;

namespace SharedModule
{
    public partial class UserControlWPFHostElement : UserControl
    {
        private CustomTaskPane m_customTaskPane = null;

        public void SetCustomTaskpane(ref CustomTaskPane ctp)
        {
            m_customTaskPane = ctp;
        }

        public UserControlWPFHostElement()
        {
            InitializeComponent();
        }
    }
}
