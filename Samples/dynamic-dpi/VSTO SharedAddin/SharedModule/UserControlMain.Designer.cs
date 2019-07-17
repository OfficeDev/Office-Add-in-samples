using System.Timers;

namespace SharedModule
{
    partial class UserControlMain
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label15 = new System.Windows.Forms.Label();
            this.txtAppWindowDpi = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.txtTaskpaneWindowDpi = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txtGetWidthHeight = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtChildWindowMixedMode = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtContainerRect = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtTaskpaneRect = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtHostWindowAwareness = new System.Windows.Forms.TextBox();
            this.txtTaskpaneWindowAwareness = new System.Windows.Forms.TextBox();
            this.txtProcessAwareness = new System.Windows.Forms.TextBox();
            this.txtThreadAwareness = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.button1 = new System.Windows.Forms.Button();
            this.btnSetHeight = new System.Windows.Forms.Button();
            this.btnSetWidth = new System.Windows.Forms.Button();
            this.label13 = new System.Windows.Forms.Label();
            this.txtSetHeight = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.txtSetWidth = new System.Windows.Forms.TextBox();
            this.btnAddTaskpane = new System.Windows.Forms.Button();
            this.btnTopLevelForm = new System.Windows.Forms.Button();
            this.label11 = new System.Windows.Forms.Label();
            this.cboTemplate = new System.Windows.Forms.ComboBox();
            this.cboNewDockLocation = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.cboDpiContext = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label15);
            this.groupBox1.Controls.Add(this.txtAppWindowDpi);
            this.groupBox1.Controls.Add(this.label14);
            this.groupBox1.Controls.Add(this.txtTaskpaneWindowDpi);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.txtGetWidthHeight);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.txtChildWindowMixedMode);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.txtContainerRect);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.txtTaskpaneRect);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.txtHostWindowAwareness);
            this.groupBox1.Controls.Add(this.txtTaskpaneWindowAwareness);
            this.groupBox1.Controls.Add(this.txtProcessAwareness);
            this.groupBox1.Controls.Add(this.txtThreadAwareness);
            this.groupBox1.Location = new System.Drawing.Point(11, 173);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(410, 263);
            this.groupBox1.TabIndex = 40;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Dynamic Dpi Info";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(7, 71);
            this.label15.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(87, 13);
            this.label15.TabIndex = 49;
            this.label15.Text = "App Window Dpi";
            // 
            // txtAppWindowDpi
            // 
            this.txtAppWindowDpi.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtAppWindowDpi.Location = new System.Drawing.Point(178, 71);
            this.txtAppWindowDpi.Name = "txtAppWindowDpi";
            this.txtAppWindowDpi.Size = new System.Drawing.Size(226, 20);
            this.txtAppWindowDpi.TabIndex = 48;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(7, 96);
            this.label14.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(116, 13);
            this.label14.TabIndex = 47;
            this.label14.Text = "Taskpane Window Dpi";
            // 
            // txtTaskpaneWindowDpi
            // 
            this.txtTaskpaneWindowDpi.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtTaskpaneWindowDpi.Location = new System.Drawing.Point(178, 96);
            this.txtTaskpaneWindowDpi.Name = "txtTaskpaneWindowDpi";
            this.txtTaskpaneWindowDpi.Size = new System.Drawing.Size(226, 20);
            this.txtTaskpaneWindowDpi.TabIndex = 46;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(7, 232);
            this.label8.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(78, 13);
            this.label8.TabIndex = 45;
            this.label8.Text = ".Width, .Height";
            // 
            // txtGetWidthHeight
            // 
            this.txtGetWidthHeight.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtGetWidthHeight.Location = new System.Drawing.Point(178, 232);
            this.txtGetWidthHeight.Name = "txtGetWidthHeight";
            this.txtGetWidthHeight.Size = new System.Drawing.Size(226, 20);
            this.txtGetWidthHeight.TabIndex = 44;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(7, 165);
            this.label7.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(143, 13);
            this.label7.TabIndex = 43;
            this.label7.Text = "Taskpane Dpi Hosting Mode";
            // 
            // txtChildWindowMixedMode
            // 
            this.txtChildWindowMixedMode.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtChildWindowMixedMode.Location = new System.Drawing.Point(178, 165);
            this.txtChildWindowMixedMode.Name = "txtChildWindowMixedMode";
            this.txtChildWindowMixedMode.Size = new System.Drawing.Size(226, 20);
            this.txtChildWindowMixedMode.TabIndex = 42;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(7, 187);
            this.label6.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(110, 13);
            this.label6.TabIndex = 41;
            this.label6.Text = "MsoCommandBar w,h";
            // 
            // txtContainerRect
            // 
            this.txtContainerRect.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtContainerRect.Location = new System.Drawing.Point(178, 187);
            this.txtContainerRect.Name = "txtContainerRect";
            this.txtContainerRect.Size = new System.Drawing.Size(226, 20);
            this.txtContainerRect.TabIndex = 40;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(7, 210);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(75, 13);
            this.label5.TabIndex = 39;
            this.label5.Text = "Taskpane w,h";
            // 
            // txtTaskpaneRect
            // 
            this.txtTaskpaneRect.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtTaskpaneRect.Location = new System.Drawing.Point(178, 210);
            this.txtTaskpaneRect.Name = "txtTaskpaneRect";
            this.txtTaskpaneRect.Size = new System.Drawing.Size(226, 20);
            this.txtTaskpaneRect.TabIndex = 38;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(7, 142);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(123, 13);
            this.label4.TabIndex = 37;
            this.label4.Text = "App Window Awareness";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(7, 119);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(152, 13);
            this.label3.TabIndex = 36;
            this.label3.Text = "Taskpane Window Awareness";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(7, 24);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(119, 13);
            this.label2.TabIndex = 35;
            this.label2.Text = "Process Dpi Awareness";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 47);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(154, 13);
            this.label1.TabIndex = 34;
            this.label1.Text = "Thread Dpi Awareness Context";
            // 
            // txtHostWindowAwareness
            // 
            this.txtHostWindowAwareness.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtHostWindowAwareness.Location = new System.Drawing.Point(178, 142);
            this.txtHostWindowAwareness.Name = "txtHostWindowAwareness";
            this.txtHostWindowAwareness.Size = new System.Drawing.Size(226, 20);
            this.txtHostWindowAwareness.TabIndex = 33;
            // 
            // txtTaskpaneWindowAwareness
            // 
            this.txtTaskpaneWindowAwareness.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtTaskpaneWindowAwareness.Location = new System.Drawing.Point(178, 119);
            this.txtTaskpaneWindowAwareness.Name = "txtTaskpaneWindowAwareness";
            this.txtTaskpaneWindowAwareness.Size = new System.Drawing.Size(226, 20);
            this.txtTaskpaneWindowAwareness.TabIndex = 32;
            // 
            // txtProcessAwareness
            // 
            this.txtProcessAwareness.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtProcessAwareness.Location = new System.Drawing.Point(177, 24);
            this.txtProcessAwareness.Name = "txtProcessAwareness";
            this.txtProcessAwareness.Size = new System.Drawing.Size(227, 20);
            this.txtProcessAwareness.TabIndex = 31;
            // 
            // txtThreadAwareness
            // 
            this.txtThreadAwareness.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtThreadAwareness.Location = new System.Drawing.Point(178, 47);
            this.txtThreadAwareness.Name = "txtThreadAwareness";
            this.txtThreadAwareness.Size = new System.Drawing.Size(226, 20);
            this.txtThreadAwareness.TabIndex = 30;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.button1);
            this.groupBox2.Controls.Add(this.btnSetHeight);
            this.groupBox2.Controls.Add(this.btnSetWidth);
            this.groupBox2.Controls.Add(this.label13);
            this.groupBox2.Controls.Add(this.txtSetHeight);
            this.groupBox2.Controls.Add(this.label12);
            this.groupBox2.Controls.Add(this.txtSetWidth);
            this.groupBox2.Controls.Add(this.btnAddTaskpane);
            this.groupBox2.Controls.Add(this.btnTopLevelForm);
            this.groupBox2.Controls.Add(this.label11);
            this.groupBox2.Controls.Add(this.cboTemplate);
            this.groupBox2.Controls.Add(this.cboNewDockLocation);
            this.groupBox2.Controls.Add(this.label10);
            this.groupBox2.Controls.Add(this.cboDpiContext);
            this.groupBox2.Controls.Add(this.label9);
            this.groupBox2.Location = new System.Drawing.Point(11, 3);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(413, 164);
            this.groupBox2.TabIndex = 41;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Create";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(15, 80);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(101, 20);
            this.button1.TabIndex = 54;
            this.button1.Text = "Open Temp Form";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnSetHeight
            // 
            this.btnSetHeight.Location = new System.Drawing.Point(6, 132);
            this.btnSetHeight.Name = "btnSetHeight";
            this.btnSetHeight.Size = new System.Drawing.Size(124, 23);
            this.btnSetHeight.TabIndex = 53;
            this.btnSetHeight.Text = "Set Height";
            this.btnSetHeight.UseVisualStyleBackColor = true;
            this.btnSetHeight.Click += new System.EventHandler(this.SetHeight);
            // 
            // btnSetWidth
            // 
            this.btnSetWidth.Location = new System.Drawing.Point(6, 104);
            this.btnSetWidth.Name = "btnSetWidth";
            this.btnSetWidth.Size = new System.Drawing.Size(124, 23);
            this.btnSetWidth.TabIndex = 52;
            this.btnSetWidth.Text = "Set Width";
            this.btnSetWidth.UseVisualStyleBackColor = true;
            this.btnSetWidth.Click += new System.EventHandler(this.SetWidth);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(176, 132);
            this.label13.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(38, 13);
            this.label13.TabIndex = 51;
            this.label13.Text = "Height";
            // 
            // txtSetHeight
            // 
            this.txtSetHeight.Location = new System.Drawing.Point(245, 132);
            this.txtSetHeight.Name = "txtSetHeight";
            this.txtSetHeight.Size = new System.Drawing.Size(159, 20);
            this.txtSetHeight.TabIndex = 50;
            this.txtSetHeight.Text = "500";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(176, 106);
            this.label12.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(35, 13);
            this.label12.TabIndex = 49;
            this.label12.Text = "Width";
            // 
            // txtSetWidth
            // 
            this.txtSetWidth.Location = new System.Drawing.Point(245, 106);
            this.txtSetWidth.Name = "txtSetWidth";
            this.txtSetWidth.Size = new System.Drawing.Size(159, 20);
            this.txtSetWidth.TabIndex = 48;
            this.txtSetWidth.Text = "500";
            // 
            // btnAddTaskpane
            // 
            this.btnAddTaskpane.Location = new System.Drawing.Point(6, 25);
            this.btnAddTaskpane.Name = "btnAddTaskpane";
            this.btnAddTaskpane.Size = new System.Drawing.Size(124, 23);
            this.btnAddTaskpane.TabIndex = 40;
            this.btnAddTaskpane.Text = "Add Taskpane";
            this.btnAddTaskpane.UseVisualStyleBackColor = true;
            this.btnAddTaskpane.Click += new System.EventHandler(this.btnAddTaskpane_Click);
            // 
            // btnTopLevelForm
            // 
            this.btnTopLevelForm.Location = new System.Drawing.Point(6, 54);
            this.btnTopLevelForm.Name = "btnTopLevelForm";
            this.btnTopLevelForm.Size = new System.Drawing.Size(124, 23);
            this.btnTopLevelForm.TabIndex = 41;
            this.btnTopLevelForm.Text = "Open Top-level Form";
            this.btnTopLevelForm.UseVisualStyleBackColor = true;
            this.btnTopLevelForm.Click += new System.EventHandler(this.btnTopLevelForm_Click);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(175, 25);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(51, 13);
            this.label11.TabIndex = 47;
            this.label11.Text = "Template";
            // 
            // cboTemplate
            // 
            this.cboTemplate.FormattingEnabled = true;
            this.cboTemplate.Location = new System.Drawing.Point(245, 25);
            this.cboTemplate.Name = "cboTemplate";
            this.cboTemplate.Size = new System.Drawing.Size(159, 21);
            this.cboTemplate.TabIndex = 46;
            // 
            // cboNewDockLocation
            // 
            this.cboNewDockLocation.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboNewDockLocation.FormattingEnabled = true;
            this.cboNewDockLocation.Location = new System.Drawing.Point(245, 54);
            this.cboNewDockLocation.Margin = new System.Windows.Forms.Padding(1);
            this.cboNewDockLocation.Name = "cboNewDockLocation";
            this.cboNewDockLocation.Size = new System.Drawing.Size(159, 21);
            this.cboNewDockLocation.TabIndex = 42;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(175, 81);
            this.label10.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(43, 13);
            this.label10.TabIndex = 45;
            this.label10.Text = "Context";
            // 
            // cboDpiContext
            // 
            this.cboDpiContext.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboDpiContext.FormattingEnabled = true;
            this.cboDpiContext.Location = new System.Drawing.Point(245, 80);
            this.cboDpiContext.Margin = new System.Windows.Forms.Padding(1);
            this.cboDpiContext.Name = "cboDpiContext";
            this.cboDpiContext.Size = new System.Drawing.Size(159, 21);
            this.cboDpiContext.TabIndex = 44;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(176, 54);
            this.label9.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(33, 13);
            this.label9.TabIndex = 43;
            this.label9.Text = "Dock";
            // 
            // UserControlMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.MinimumSize = new System.Drawing.Size(317, 0);
            this.Name = "UserControlMain";
            this.Size = new System.Drawing.Size(439, 453);
            this.Load += new System.EventHandler(this.UserControlWinForm_Load);
            this.Resize += new System.EventHandler(this.UserControlWinForm_Resize);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txtGetWidthHeight;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtChildWindowMixedMode;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtContainerRect;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtTaskpaneRect;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtHostWindowAwareness;
        private System.Windows.Forms.TextBox txtTaskpaneWindowAwareness;
        private System.Windows.Forms.TextBox txtProcessAwareness;
        private System.Windows.Forms.TextBox txtThreadAwareness;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ComboBox cboTemplate;
        private System.Windows.Forms.ComboBox cboNewDockLocation;
        private System.Windows.Forms.Button btnAddTaskpane;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Button btnTopLevelForm;
        private System.Windows.Forms.ComboBox cboDpiContext;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox txtSetHeight;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox txtSetWidth;
        private System.Windows.Forms.Button btnSetHeight;
        private System.Windows.Forms.Button btnSetWidth;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.TextBox txtTaskpaneWindowDpi;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox txtAppWindowDpi;
        private System.Windows.Forms.Button button1;
    }
}
