namespace SharedModule
{
    partial class TaskPaneRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public TaskPaneRibbon()
            : base(SharedApp.RibbonFactory())
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
            this.tab1.SuspendLayout();
            this.groupTaskpane.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupTaskpane);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // groupTaskpane
            // 
            this.groupTaskpane.Items.Add(this.buttonAddTaskpane);
            this.groupTaskpane.Items.Add(this.buttonCloseAllTaskpanes);
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
    }

    partial class ThisRibbonCollection
    {
        internal TaskPaneRibbon TaskPaneRibbon
        {
            get
            {
                ThisRibbonCollection thisRibbonCollection = (ThisRibbonCollection)SharedApp.Ribbons();
                return thisRibbonCollection.GetRibbon<TaskPaneRibbon>();
            }
        }
    }
}
