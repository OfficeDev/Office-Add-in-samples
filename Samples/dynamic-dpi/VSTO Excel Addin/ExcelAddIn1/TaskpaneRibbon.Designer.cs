namespace ExcelAddIn1
{
    partial class TaskpaneRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public TaskpaneRibbon()
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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.buttonAddTaskpane = this.Factory.CreateRibbonButton();
            this.buttonCloseAllTaskpanes = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "DDPI Add-in";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.buttonAddTaskpane);
            this.group1.Items.Add(this.buttonCloseAllTaskpanes);
            this.group1.Label = "DDPI Taskpanes";
            this.group1.Name = "group1";
            // 
            // buttonAddTaskpane
            // 
            this.buttonAddTaskpane.Image = global::ExcelAddIn1.Properties.Resources.PlusIcon;
            this.buttonAddTaskpane.Label = "Add Taskpane";
            this.buttonAddTaskpane.Name = "buttonAddTaskpane";
            this.buttonAddTaskpane.ScreenTip = "Add a new taskpane";
            this.buttonAddTaskpane.ShowImage = true;
            this.buttonAddTaskpane.SuperTip = "Add a new default taskpane to the right dock, 700px wide.";
            this.buttonAddTaskpane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAddTaskpane_Click);
            // 
            // buttonCloseAllTaskpanes
            // 
            this.buttonCloseAllTaskpanes.Label = "Close All Taskpanes";
            this.buttonCloseAllTaskpanes.Name = "buttonCloseAllTaskpanes";
            this.buttonCloseAllTaskpanes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCloseAllTaskpanes_Click);
            // 
            // TaskpaneRibbon
            // 
            this.Name = "TaskpaneRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.TaskpaneRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddTaskpane;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCloseAllTaskpanes;
    }

    partial class ThisRibbonCollection
    {
        internal TaskpaneRibbon TaskpaneRibbon
        {
            get { return this.GetRibbon<TaskpaneRibbon>(); }
        }
    }
}
