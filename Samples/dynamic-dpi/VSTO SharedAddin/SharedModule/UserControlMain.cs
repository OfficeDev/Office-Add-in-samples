using Microsoft.Office.Tools;
using System;
using System.Diagnostics;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using static SharedModule.DPIHelper;

namespace SharedModule
{
    public partial class UserControlMain : UserControl
    {
        private System.Windows.Forms.Timer refreshTimer = new System.Windows.Forms.Timer();
        private CustomTaskPane m_customTaskPane = null;

        public void SetCustomTaskpane(ref CustomTaskPane ctp)
        {
            m_customTaskPane = ctp;
        }

        private void RefreshValues()
        {
            if (!this.Visible || this.Disposing || this.IsDisposed) return;

            IntPtr hWndHost = Process.GetCurrentProcess().MainWindowHandle;
            IntPtr hWndTaskpane = IntPtr.Zero;
            try
            {
	            hWndTaskpane = this.Handle;
            }
            catch(System.ObjectDisposedException)
            {
            }

            IntPtr hWndContainer = FindParentWithClassName(hWndTaskpane, "MsoCommandBar");
            IntPtr hWndTaskpaneHost = FindParentWithClassName(hWndTaskpane, "CMMOcxHostChildWindowMixedMode");

            this.txtThreadAwareness.Text = GetThreadDpiAwarenessContext().ToString();
            this.txtProcessAwareness.Text = GetProcessDpiAwareness().ToString();

            if (this.Handle != null)
            {
                this.txtTaskpaneWindowAwareness.Text =
                    GetWindowDpiAwarenessContext(this.Handle).ToString();
            }

            this.txtHostWindowAwareness.Text = 
                GetWindowDpiAwarenessContext(hWndHost).ToString();

            this.txtChildWindowMixedMode.Text =
                GetWindowDpiHostingBehavior(hWndTaskpaneHost).ToString();

            this.txtTaskpaneRect.Text = HwndInfoString(hWndTaskpane);
            this.txtContainerRect.Text = HwndInfoString(hWndContainer);
            uint dpiTaskpane = GetDpiForWindow(hWndTaskpane);
            uint dpiApp = GetDpiForWindow(hWndHost);
            this.txtTaskpaneWindowDpi.Text = String.Format("{0} ({1:P0})", dpiTaskpane, dpiTaskpane / 96.0);
            this.txtAppWindowDpi.Text = String.Format("{0} ({1:P0})", dpiApp, dpiApp / 96.0);

            if (m_customTaskPane != null)
            {
                this.txtGetWidthHeight.Text = string.Format("{0}, {1}", m_customTaskPane.Width, m_customTaskPane.Height);
            }
        }

        private string HwndInfoString(IntPtr hWnd)
        {
            RECT rSA;
            RECT rPMA;

            {
                DPIContextBlock saBlock = new DPIContextBlock(DPI_AWARENESS_CONTEXT_SYSTEM_AWARE);
                rSA = GetWindowRectangle(hWnd);
            }
            {
                DPIContextBlock pmaBlock = new DPIContextBlock(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE);
                rPMA = GetWindowRectangle(hWnd);
            }

            return String.Format("SA: {0}, {1} PMA: {2}, {3}",
                rSA.Width.ToString(),
                rSA.Height.ToString(),
                rPMA.Width.ToString(),
                rPMA.Height.ToString());
        }

        public UserControlMain()
        {
            InitializeComponent();
            // Setup timer callback
            refreshTimer.Tick += (Object o, EventArgs e) => RefreshValues();
            refreshTimer.Interval = 1000;

            cboNewDockLocation.DataSource = Enum.GetValues(typeof(Microsoft.Office.Core.MsoCTPDockPosition));
            cboDpiContext.DataSource = DpiAwarenessContexts;

            Type formType = typeof(UserControl);
            foreach (Type type in Assembly.GetExecutingAssembly().GetTypes())
                if (formType.IsAssignableFrom(type))
                {
                    cboTemplate.Items.Add(type.Name);
                }

            cboTemplate.Text = this.Name;
            if (this.Handle != null)
            {
                cboDpiContext.Text =
                    GetWindowDpiAwarenessContext(this.Handle).ToString();
            }

            AutoRefreshValues(true);
        }

        private DPI_AWARENESS_CONTEXT GetSelectedDpiAwarenessContext()
        {
            int index = DpiAwarenessContexts.Length - 1;
            for (; index >= 0; index-- )
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

        private void CreateNewTaskpane()
        {
            int width = 0;
            int height = 0;
            Microsoft.Office.Core.MsoCTPDockPosition dock;

            if (!Enum.TryParse<Microsoft.Office.Core.MsoCTPDockPosition>(cboNewDockLocation.SelectedValue.ToString(), out dock))
            {
                dock = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            }
            Int32.TryParse(txtSetWidth.Text, out width);
            Int32.TryParse(txtSetHeight.Text, out height);

            SharedApp.AppTaskPanes.CreateTaskpaneInstance(cboTemplate.Text, width, height, dock);
        }


        private void btnAddTaskpane_Click(object sender, EventArgs e)
        {
            CreateNewTaskpane();
        }

        private void btnTopLevelForm_Click(object sender, EventArgs e)
        {
            DPIContextBlock context = new DPIContextBlock(GetSelectedDpiAwarenessContext());
            TopLevelWinForm f1 = new TopLevelWinForm(cboTemplate.Text);
            f1.Show();
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

        private void SetThreadDPI(DPI_AWARENESS_CONTEXT newvalue)
        {
            SetThreadDPI(newvalue, true);
        }
        private void SetThreadDPI(DPI_AWARENESS_CONTEXT newvalue, bool showMessage)
        {
            DPI_AWARENESS_CONTEXT previous =
                SetThreadDpiAwarenessContext(newvalue);
            int processId = Process.GetCurrentProcess().Id;
            int threadId = Thread.CurrentThread.ManagedThreadId;
            if (showMessage)
            {
                MessageBox.Show(String.Format("DPI Awareness set to {0}, was {1}\nProcessId {2}, ThreadId {3}", newvalue, previous, processId, threadId));
            }
        }

        private void setCWMMNormal_Click(object sender, EventArgs e)
        {
            SetThreadDpiHostingBehavior(DPI_HOSTING_BEHAVIOR.DPI_HOSTING_BEHAVIOR_DEFAULT);
            // MessageBox.Show(String.Format("DPI Hosting Behavior is {0}", GetChildWindowMixedMode(this.Handle).ToString()));

        }

        private void UserControlWinForm_Load(object sender, EventArgs e)
        {
            RefreshValues();
        }

        private void UserControlWinForm_Resize(object sender, EventArgs e)
        {
            RefreshValues();
        }

        private void SetWidth(object sender, EventArgs e)
        {
            if (m_customTaskPane == null) return;

            int width = m_customTaskPane.Width;
            if (int.TryParse(txtSetWidth.Text, out width))
            {
                try
                {
                    m_customTaskPane.Width = width;
                }
                catch (System.Runtime.InteropServices.COMException except)
                {
                    MessageBox.Show(except.Message);
                }
            }
        }

        private void SetHeight(object sender, EventArgs e)
        {
            if (m_customTaskPane == null) return;

            int Height = m_customTaskPane.Height;
            if (int.TryParse(txtSetHeight.Text, out Height))
            {
                try
                {
                    m_customTaskPane.Height = Height;
                }
                catch (System.Runtime.InteropServices.COMException except)
                {
                    MessageBox.Show(except.Message);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DPIContextBlock context = new DPIContextBlock(GetSelectedDpiAwarenessContext());
            TempForm frm = new TempForm();
            frm.Show();
        }
    }
}
