using System;
using System.Reflection;
using System.Windows.Forms;
using static SharedModule.DPIHelper;
using Microsoft.Office.Tools;
using System.Windows.Interop;

namespace SharedModule
{
    public partial class Top_Level_Forms : UserControl
    {
        private CustomTaskPane m_customTaskPane = null;

        public Top_Level_Forms()
        {
            InitializeComponent();

            // Hook up DPIChanged event
            // this.Parent.dpi += new System.Windows.Forms.DpiChangedEventHandler();

            cboDpiContext.DataSource = DpiAwarenessContexts;
            cboHostingBehavior.DataSource = DpiHostingBehaviors;

            // Fill template combo
            Type formType = typeof(UserControl);
            foreach (Type type in Assembly.GetExecutingAssembly().GetTypes())
                if (formType.IsAssignableFrom(type))
                {
                    cboTemplate.Items.Add(type.Name);
                }

            // Set cbo defaults
            cboTemplate.Text = this.Name;

            if (this.Handle != null)
            {
                cboDpiContext.Text =
                    DPIHelper.GetWindowDpiAwarenessContext(this.Handle).ToString();
            }

            cboHostingBehavior.Text =
                DPIHelper.GetThreadDpiHostingBehavior().ToString();

        }

        public void SetCustomTaskpane(ref CustomTaskPane ctp)
        {
            m_customTaskPane = ctp;
        }

        private DPI_AWARENESS_CONTEXT GetSelectedDpiAwarenessContext()
        {
            int index = DpiAwarenessContexts.Length - 1;
            for (; index >= 0; index--)
            {
                if (DpiAwarenessContexts[index].ToString().Equals(cboDpiContext.SelectedValue.ToString()))
                {
                    break;
                }
            }

            if (index >= 0)
            {
                return DpiAwarenessContexts[index];
            }
            return DPI_AWARENESS_CONTEXT_SYSTEM_AWARE;
        }

        private DPI_HOSTING_BEHAVIOR GetSelectedHostingBehavior()
        {
            int index = DpiHostingBehaviors.Length - 1;
            for (; index >= 0; index--)
            {
                if (DpiHostingBehaviors[index].ToString().Equals(cboHostingBehavior.SelectedValue.ToString()))
                {
                    break;
                }
            }

            if (index >= 0)
            {
                return DpiHostingBehaviors[index];
            }
            return DPI_HOSTING_BEHAVIOR.DPI_HOSTING_BEHAVIOR_DEFAULT;

        }
        private void btnTopLevelForm_Click(object sender, EventArgs e)
        {
            DPIContextBlock context = new DPIContextBlock(GetSelectedDpiAwarenessContext());
            SetThreadDpiHostingBehavior(GetSelectedHostingBehavior());
            TopLevelWinForm f1 = new TopLevelWinForm(cboTemplate.Text);
            f1.Show();
        }
    }
}
