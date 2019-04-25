using System.Windows.Forms;
using Microsoft.Office.Tools;

namespace SharedModule
{
    public partial class UserControlWBC : UserControl
    {
        private CustomTaskPane m_customTaskPane = null;

        public void SetCustomTaskpane(ref CustomTaskPane ctp)
        {
            m_customTaskPane = ctp;
        }
        public UserControlWBC()
        {
            InitializeComponent();
            webBrowser1.Navigate(txtUrl.Text);
        }

        private void btnBack_Click(object sender, System.EventArgs e)
        {
            if (webBrowser1.CanGoBack)
                webBrowser1.GoBack();
        }

        private void UserControlWBC_Load(object sender, System.EventArgs e)
        {
        }

        private void txtUrl_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                webBrowser1.Navigate(txtUrl.Text);
        }

        private void webBrowser1_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            txtUrl.Text = webBrowser1.Url.AbsoluteUri;
        }
    }
}
