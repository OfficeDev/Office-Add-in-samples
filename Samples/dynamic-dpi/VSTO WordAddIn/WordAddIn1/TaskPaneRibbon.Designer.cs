// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

namespace WordAddIn1
{
    partial class TaskPaneRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public TaskPaneRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
			this.tab1 = this.Factory.CreateRibbonTab();
			this.groupTaskpane = this.Factory.CreateRibbonGroup();
			this.buttonAddTaskpane = this.Factory.CreateRibbonButton();
			this.buttonCloseAllTaskpanes = this.Factory.CreateRibbonButton();
			this.btnOpenHelpNewProcess = this.Factory.CreateRibbonButton();
			this.menu1 = this.Factory.CreateRibbonMenu();
			this.tab1.SuspendLayout();
			this.groupTaskpane.SuspendLayout();
			this.SuspendLayout();
			// 
			// tab1
			// 
			this.tab1.Groups.Add(this.groupTaskpane);
			this.tab1.Label = "DDPI Add-in";
			this.tab1.Name = "tab1";
			// 
			// groupTaskpane
			// 
			this.groupTaskpane.Items.Add(this.buttonAddTaskpane);
			this.groupTaskpane.Items.Add(this.buttonCloseAllTaskpanes);
			this.groupTaskpane.Items.Add(this.menu1);
			this.groupTaskpane.Label = "Taskpane";
			this.groupTaskpane.Name = "groupTaskpane";
			// 
			// buttonAddTaskpane
			// 
			this.buttonAddTaskpane.Label = "Add Taskpane";
			this.buttonAddTaskpane.Name = "buttonAddTaskpane";
			this.buttonAddTaskpane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAddTaskpane_Click);
			// 
			// buttonCloseAllTaskpanes
			// 
			this.buttonCloseAllTaskpanes.Label = "Close All Taskpanes";
			this.buttonCloseAllTaskpanes.Name = "buttonCloseAllTaskpanes";
			this.buttonCloseAllTaskpanes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCloseAllTaskpanes_Click);
			// 
			// btnOpenHelpNewProcess
			// 
			this.btnOpenHelpNewProcess.Description = "Open help in new process";
			this.btnOpenHelpNewProcess.Label = "Help (new process)";
			this.btnOpenHelpNewProcess.Name = "btnOpenHelpNewProcess";
			this.btnOpenHelpNewProcess.ShowImage = true;
			this.btnOpenHelpNewProcess.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOpenHelpNewProcess_Click);
			// 
			// menu1
			// 
			this.menu1.Items.Add(this.btnOpenHelpNewProcess);
			this.menu1.Label = "menu1";
			this.menu1.Name = "menu1";
			// 
			// TaskPaneRibbon
			// 
			this.Name = "TaskPaneRibbon";
			this.RibbonType = "Microsoft.Word.Document";
			this.Tabs.Add(this.tab1);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.TaskPaneRibbon_Load);
			this.tab1.ResumeLayout(false);
			this.tab1.PerformLayout();
			this.groupTaskpane.ResumeLayout(false);
			this.groupTaskpane.PerformLayout();
			this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupTaskpane;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddTaskpane;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCloseAllTaskpanes;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpenHelpNewProcess;
		internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu1;
	}

	partial class ThisRibbonCollection
    {
        internal TaskPaneRibbon TaskPaneRibbon
        {
            get { return this.GetRibbon<TaskPaneRibbon>(); }
        }
    }
}
