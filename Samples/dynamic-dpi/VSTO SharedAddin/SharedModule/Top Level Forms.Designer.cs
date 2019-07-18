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
			this.SuspendLayout();
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(4, 13);
			this.label11.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(75, 20);
			this.label11.TabIndex = 51;
			this.label11.Text = "Template";
			// 
			// cboTemplate
			// 
			this.cboTemplate.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.cboTemplate.FormattingEnabled = true;
			this.cboTemplate.Location = new System.Drawing.Point(165, 13);
			this.cboTemplate.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
			this.cboTemplate.Name = "cboTemplate";
			this.cboTemplate.Size = new System.Drawing.Size(260, 28);
			this.cboTemplate.TabIndex = 50;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(4, 53);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(64, 20);
			this.label10.TabIndex = 49;
			this.label10.Text = "Context";
			// 
			// cboDpiContext
			// 
			this.cboDpiContext.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.cboDpiContext.FormattingEnabled = true;
			this.cboDpiContext.Location = new System.Drawing.Point(165, 51);
			this.cboDpiContext.Margin = new System.Windows.Forms.Padding(2, 1, 2, 1);
			this.cboDpiContext.Name = "cboDpiContext";
			this.cboDpiContext.Size = new System.Drawing.Size(260, 28);
			this.cboDpiContext.TabIndex = 48;
			// 
			// btnTopLevelForm
			// 
			this.btnTopLevelForm.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnTopLevelForm.Location = new System.Drawing.Point(165, 341);
			this.btnTopLevelForm.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
			this.btnTopLevelForm.Name = "btnTopLevelForm";
			this.btnTopLevelForm.Size = new System.Drawing.Size(186, 35);
			this.btnTopLevelForm.TabIndex = 52;
			this.btnTopLevelForm.Text = "Open Top-level Form";
			this.btnTopLevelForm.UseVisualStyleBackColor = true;
			this.btnTopLevelForm.Click += new System.EventHandler(this.btnTopLevelForm_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(4, 94);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(130, 20);
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
			this.cboHostingBehavior.Location = new System.Drawing.Point(165, 92);
			this.cboHostingBehavior.Margin = new System.Windows.Forms.Padding(2, 1, 2, 1);
			this.cboHostingBehavior.Name = "cboHostingBehavior";
			this.cboHostingBehavior.Size = new System.Drawing.Size(260, 28);
			this.cboHostingBehavior.TabIndex = 53;
			// 
			// txtDpi
			// 
			this.txtDpi.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.txtDpi.Location = new System.Drawing.Point(165, 133);
			this.txtDpi.Multiline = true;
			this.txtDpi.Name = "txtDpi";
			this.txtDpi.Size = new System.Drawing.Size(260, 193);
			this.txtDpi.TabIndex = 55;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(4, 133);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(33, 20);
			this.label2.TabIndex = 56;
			this.label2.Text = "Dpi";
			// 
			// Top_Level_Forms
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtDpi);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cboHostingBehavior);
			this.Controls.Add(this.btnTopLevelForm);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.cboTemplate);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.cboDpiContext);
			this.Name = "Top_Level_Forms";
			this.Size = new System.Drawing.Size(429, 421);
			this.ResumeLayout(false);
			this.PerformLayout();

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
	}
}
