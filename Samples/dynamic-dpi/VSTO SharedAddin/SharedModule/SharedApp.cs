using Microsoft.Office.Tools;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;

namespace SharedModule
{
	public static class SharedApp
	{
		public static object HostApp { get; set; }

		public static TaskPanes AppTaskPanes { get; set; }

		public static void InitAppTaskPanes(ref CustomTaskPaneCollection value)
		{
			AppTaskPanes = new TaskPanes(ref value);
		}

		public static string HelpFileName()
		{
			string location = AppDomain.CurrentDomain.BaseDirectory;
			string helpFile = Path.Combine(location, "samplehelp.chm");
			return helpFile;
		}

		public static void View_Help(bool fAsNewThread)
		{
			string helpFile = HelpFileName();
			string exePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "hh.exe");
			string quotedHelpfile = String.Format("\"{0}\"", helpFile);
			string exeAndFile = String.Format("{0} {1}", exePath, quotedHelpfile);
			IntPtr mainHwnd = Process.GetCurrentProcess().MainWindowHandle;

			SharedModule.HtmlHelpHelper.HtmlHelp((IntPtr)0, helpFile, 0, 0);

			//ProcessStartInfo prs = new ProcessStartInfo();
			//prs.FileName = exePath;
			//prs.Arguments = quotedHelpfile;
			//if (!fAsNewThread)
			//{
			//	SharedModule.ProcessCreator.CreateProcess(null, exeAndFile, Process.GetCurrentProcess().Id);
			//}
			//else
			//{
			//	// Process pr = new Process();
			//	Process pr = Process.GetCurrentProcess();
			//	pr.StartInfo = prs;

			//	ThreadStart ths = new ThreadStart(() => pr.Start());
			//	Thread th = new Thread(ths);
			//	th.Start();
			//}
		}

	}
}
