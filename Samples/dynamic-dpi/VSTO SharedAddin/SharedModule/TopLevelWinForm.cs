using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SharedModule
{
    public partial class TopLevelWinForm : DpiAwareWindowsForm
    {
        public TopLevelWinForm(string userControlName)
        {
            InitializeComponent();
            this.SuspendLayout();
            LoadUserControl(userControlName);
            this.ResumeLayout(true);
        }

        public void LoadUserControl(string userControlName)
        {
            UserControl userControlAdd = (UserControl)Activator.CreateInstance(null, String.Format("{0}.{1}", this.GetType().Namespace, userControlName)).Unwrap();
            UserControl userControlCopy = (UserControl)this.Controls["userControlWinForm1"];

            userControlCopy.Name = "userControlWinForm1Copy";
            userControlAdd.Name = "userControlWinForm1";
            userControlAdd.Anchor = userControlCopy.Anchor;
            userControlAdd.Location = userControlCopy.Location;
            userControlAdd.Dock = userControlCopy.Dock;
            userControlAdd.Visible = true;

            this.Controls.Add(userControlAdd);
            this.Controls.Remove(userControlCopy);
        }
    }
}
