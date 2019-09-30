using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SharedModule
{
	public partial class ShowHelp : UserControl
	{
		public ShowHelp()
		{
			InitializeComponent();
		}

		private void button1_Click(object sender, EventArgs e)
		{
			Help.ShowHelp(button1, SharedApp.HelpFileName());
		}
	}
}
