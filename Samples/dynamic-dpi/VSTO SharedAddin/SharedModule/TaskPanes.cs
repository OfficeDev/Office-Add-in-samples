// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Office.Tools;
using System;
using System.Reflection;
using System.Windows;

namespace SharedModule
{
    public class TaskPanes
    {
        private CustomTaskPaneCollection m_tpc;

        public TaskPanes(ref CustomTaskPaneCollection taskPanes)
        {
            m_tpc = taskPanes;
        }

        public void CreateTaskpaneInstance()
        {
            CreateTaskpaneInstance("UserControlMain", 700, 0, Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight);
        }
        public void CreateTaskpaneInstance(string userControlName, int width, int height, Microsoft.Office.Core.MsoCTPDockPosition dockPosition)
        {
            CustomTaskPane customTaskpane;

            // Add a custom taskpane
            dynamic userControl = System.Activator.CreateInstance(null, String.Format("{0}.{1}", this.GetType().Namespace, userControlName)).Unwrap();
            customTaskpane = m_tpc.Add(userControl, string.Format("DDPI Custom Taskpane {0}", m_tpc.Count + 1));
            customTaskpane.DockPosition = dockPosition;
            try
            {
                if (width != 0 && 
                   (dockPosition == Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating ||
                    dockPosition == Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft || 
                    dockPosition == Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight))
                    customTaskpane.Width = width;
            }
            catch (System.Runtime.InteropServices.COMException except)
            {
                MessageBox.Show(except.Message);
            }
            try
            {
                if (height != 0 &&
                   (dockPosition == Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating ||
                    dockPosition == Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionBottom||
                    dockPosition == Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionTop))
                    customTaskpane.Height = height;
            }
            catch (System.Runtime.InteropServices.COMException except)
            {
                MessageBox.Show(except.Message);
            }

			customTaskpane.DockPositionChanged += new System.EventHandler(CustomTaskPaneDockChangeHandler);

			customTaskpane.Visible = true;
			// Set a ref to the custom taskpane if the method exists
			try
			{
				userControl.SetCustomTaskpane(ref customTaskpane);
			}
			catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
			{
			}

		}

		private void CustomTaskPaneDockChangeHandler(object sender, EventArgs e)
		{
			CustomTaskPane ctp = (CustomTaskPane)sender;

			MessageBox.Show(string.Format("CustomTaskPane DockChange {0} with DpiThreadAwarenessContext {1}", ctp.DockPosition, DPIHelper.GetThreadDpiAwarenessContext().ToString()));
		}


		public void CloseAllTaskpanes()
        {
            int count = m_tpc.Count;
            for (int i = 0; i < count; i++)
            {
                m_tpc[0].Visible = false;
                m_tpc[0].Dispose();
                m_tpc.RemoveAt(0);
            }
        }
    }
}
