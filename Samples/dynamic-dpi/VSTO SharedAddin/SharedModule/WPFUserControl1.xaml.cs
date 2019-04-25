using System;
using System.Windows;
using System.Windows.Controls;

namespace SharedModule
{
    /// <summary>
    /// Interaction logic for WPFUserControl1.xaml
    /// </summary>
    public partial class WPFUserControl1 : UserControl
    {
        public WPFUserControl1()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(string.Format("Button click.  Text box \"{0}\"", textBox.Text));
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Check box checked");
        }

        private void ComboBox_DropDownClosed(object sender, EventArgs e)
        {
            MessageBox.Show("Combo box dropdown closed");
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            MessageBox.Show("Combo box dropdown selection changed");
        }
    }
}
