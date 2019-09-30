using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VSTOSharedAddin
{
    public partial class UserControlWinForm : UserControl
    {
        // My Code
        private System.Windows.Forms.Timer refreshTimer = new System.Windows.Forms.Timer();

        private void RefreshValues()
        {
            this.txtThreadAwareness.Text = DPIHelper.GetThreadDpiAwareness().ToString();
            this.txtProcessAwareness.Text = DPIHelper.GetProcessDpi().ToString();

            if (this.Handle != null)
            {
                Debug.WriteLine("Toplevel context window hwnd={0}", this.Handle);
                this.txtTaskpaneWindowAwareness.Text =
                    DPIHelper.GetWindowDpiAwareness(this.Handle).ToString();
            }

            this.txtHostWindowAwareness.Text = 
                DPIHelper.GetWindowDpiAwareness((IntPtr)Globals.ThisAddIn.Application.Hwnd).ToString();
        }
        public UserControlWinForm()
        {
            InitializeComponent();
            // Setup timer callback
            //refreshTimer.Tick += (Object o, EventArgs e) => RefreshValues();
            //refreshTimer.Interval = 1000;
        }

        private void btnSetThreadSA_Click(object sender, EventArgs e)
        {
            SetThreadDPI(DPIHelper.DPI_AWARENESS_CONTEXT.DPI_AWARENESS_CONTEXT_SYSTEM_AWARE, false);
            Globals.ThisAddIn.CreateTaskpaneInstance();
        }

        private void btnSetThreadPMA_Click(object sender, EventArgs e)
        {
            SetThreadDPI(DPIHelper.DPI_AWARENESS_CONTEXT.DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE);
        }

        private void btnSetThreadPMAV2_Click(object sender, EventArgs e)
        {
            SetThreadDPI(DPIHelper.DPI_AWARENESS_CONTEXT.DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
        }

        private void btnOpenNonModalSA_Click(object sender, EventArgs e)
        {
            SetThreadDPI(DPIHelper.DPI_AWARENESS_CONTEXT.DPI_AWARENESS_CONTEXT_SYSTEM_AWARE, false);
            Form f1 = new Form1();
            f1.Show();
        }

        private void btnSetCWMM_Click(object sender, EventArgs e)
        {
            DPIHelper.SetChildWindowMixedMode(DPIHelper.DPI_HOSTING_BEHAVIOR.DPI_HOSTING_BEHAVIOR_MIXED);
            MessageBox.Show(String.Format("DPI Hosting Behavior is {0}", DPIHelper.GetChildWindowMixedMode(this.Handle).ToString()));
        }

        private void btnRefreshValues_Click(object sender, EventArgs e)
        {
            RefreshValues();
        }

        private void AutoRefreshValues(bool start)
        {
            if (start)
            {
                refreshTimer.Start();
            }
            else
            {
                refreshTimer.Stop();
            }
        }

        private void SetThreadDPI(DPIHelper.DPI_AWARENESS_CONTEXT newvalue)
        {
            SetThreadDPI(newvalue, true);
        }
        private void SetThreadDPI(DPIHelper.DPI_AWARENESS_CONTEXT newvalue, bool showMessage)
        {
            DPIHelper.DPI_AWARENESS_CONTEXT previous =
                DPIHelper.SetThreadDpiAwareness(newvalue);
            int processId = Process.GetCurrentProcess().Id;
            int threadId = Thread.CurrentThread.ManagedThreadId;
            if (showMessage)
            {
                MessageBox.Show(String.Format("DPI Awareness set to {0}, was {1}\nProcessId {2}, ThreadId {3}", newvalue, previous, processId, threadId));
            }
        }

        private void chkAutoRefresh_Click(object sender, EventArgs e)
        {
            AutoRefreshValues(this.chkAutoRefresh.Checked);
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            RefreshValues();
        }

        private void setCWMMNormal_Click(object sender, EventArgs e)
        {
            DPIHelper.SetChildWindowMixedMode(DPIHelper.DPI_HOSTING_BEHAVIOR.DPI_HOSTING_BEHAVIOR_DEFAULT);
            MessageBox.Show(String.Format("DPI Hosting Behavior is {0}", DPIHelper.GetChildWindowMixedMode(this.Handle).ToString()));

        }
    }
}
