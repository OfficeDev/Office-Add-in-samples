// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

namespace SharedModule
{
	partial class Top_Level_Forms
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
            this.label11 = new System.Windows.Forms.Label();
            this.cboTemplate = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.cboDpiContext = new System.Windows.Forms.ComboBox();
            this.btnTopLevelForm = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.cboHostingBehavior = new System.Windows.Forms.ComboBox();
            this.txtDpi = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(3, 11);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(51, 13);
            this.label11.TabIndex = 51;
            this.label11.Text = "Template";
            // 
            // cboTemplate
            // 
            this.cboTemplate.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboTemplate.FormattingEnabled = true;
            this.cboTemplate.Location = new System.Drawing.Point(92, 11);
            this.cboTemplate.Name = "cboTemplate";
            this.cboTemplate.Size = new System.Drawing.Size(182, 21);
            this.cboTemplate.TabIndex = 50;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(3, 42);
            this.label10.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(43, 13);
            this.label10.TabIndex = 49;
            this.label10.Text = "Context";
            // 
            // cboDpiContext
            // 
            this.cboDpiContext.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboDpiContext.FormattingEnabled = true;
            this.cboDpiContext.Location = new System.Drawing.Point(92, 42);
            this.cboDpiContext.Margin = new System.Windows.Forms.Padding(1);
            this.cboDpiContext.Name = "cboDpiContext";
            this.cboDpiContext.Size = new System.Drawing.Size(182, 21);
            this.cboDpiContext.TabIndex = 48;
            // 
            // btnTopLevelForm
            // 
            this.btnTopLevelForm.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnTopLevelForm.Location = new System.Drawing.Point(92, 236);
            this.btnTopLevelForm.Name = "btnTopLevelForm";
            this.btnTopLevelForm.Size = new System.Drawing.Size(124, 23);
            this.btnTopLevelForm.TabIndex = 52;
            this.btnTopLevelForm.Text = "Open Top-level Form";
            this.btnTopLevelForm.UseVisualStyleBackColor = true;
            this.btnTopLevelForm.Click += new System.EventHandler(this.btnTopLevelForm_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(2, 67);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(88, 13);
            this.label1.TabIndex = 54;
            this.label1.Text = "Hosting Behavior";
            // 
            // cboHostingBehavior
            // 
            this.cboHostingBehavior.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboHostingBehavior.FormattingEnabled = true;
            this.cboHostingBehavior.Items.AddRange(new object[] {
            "DPI_HOSTING_BEHAVIOR_DEFAULT",
            "DPI_HOSTING_BEHAVIOR_MIXED"});
            this.cboHostingBehavior.Location = new System.Drawing.Point(92, 67);
            this.cboHostingBehavior.Margin = new System.Windows.Forms.Padding(1);
            this.cboHostingBehavior.Name = "cboHostingBehavior";
            this.cboHostingBehavior.Size = new System.Drawing.Size(182, 21);
            this.cboHostingBehavior.TabIndex = 53;
            // 
            // txtDpi
            // 
            this.txtDpi.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtDpi.Location = new System.Drawing.Point(92, 91);
            this.txtDpi.Margin = new System.Windows.Forms.Padding(2);
            this.txtDpi.Multiline = true;
            this.txtDpi.Name = "txtDpi";
            this.txtDpi.Size = new System.Drawing.Size(178, 126);
            this.txtDpi.TabIndex = 55;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 91);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(23, 13);
            this.label2.TabIndex = 56;
            this.label2.Text = "Dpi";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label11);
            this.panel1.Controls.Add(this.btnTopLevelForm);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.cboTemplate);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.txtDpi);
            this.panel1.Controls.Add(this.label10);
            this.panel1.Controls.Add(this.cboDpiContext);
            this.panel1.Controls.Add(this.cboHostingBehavior);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(286, 274);
            this.panel1.TabIndex = 57;
            // 
            // Top_Level_Forms
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.Controls.Add(this.panel1);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "Top_Level_Forms";
            this.Size = new System.Drawing.Size(286, 274);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

		}

		#endregion
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.ComboBox cboTemplate;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.ComboBox cboDpiContext;
		private System.Windows.Forms.Button btnTopLevelForm;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.ComboBox cboHostingBehavior;
		private System.Windows.Forms.TextBox txtDpi;
		private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel1;
    }
}
