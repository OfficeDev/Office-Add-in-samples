using System.Timers;

namespace ExcelAddIn1
{
    partial class UserControlWinForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UserControlWinForm));
            this.txtThreadAwareness = new System.Windows.Forms.TextBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.btnSetThreadSA = new System.Windows.Forms.Button();
            this.btnSetThreadPMA = new System.Windows.Forms.Button();
            this.btnSetThreadPMAV2 = new System.Windows.Forms.Button();
            this.btnOpenNonModalSA = new System.Windows.Forms.Button();
            this.btnSetCWMM = new System.Windows.Forms.Button();
            this.txtProcessAwareness = new System.Windows.Forms.TextBox();
            this.txtTaskpaneWindowAwareness = new System.Windows.Forms.TextBox();
            this.txtHostWindowAwareness = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.chkAutoRefresh = new System.Windows.Forms.CheckBox();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // txtThreadAwareness
            // 
            this.txtThreadAwareness.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtThreadAwareness.Location = new System.Drawing.Point(237, 5);
            this.txtThreadAwareness.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtThreadAwareness.Name = "txtThreadAwareness";
            this.txtThreadAwareness.Size = new System.Drawing.Size(213, 26);
            this.txtThreadAwareness.TabIndex = 1;
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBox1.BackgroundImage")));
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.pictureBox1.Location = new System.Drawing.Point(28, 480);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(408, 397);
            this.pictureBox1.TabIndex = 2;
            this.pictureBox1.TabStop = false;
            // 
            // btnSetThreadSA
            // 
            this.btnSetThreadSA.Location = new System.Drawing.Point(12, 207);
            this.btnSetThreadSA.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnSetThreadSA.Name = "btnSetThreadSA";
            this.btnSetThreadSA.Size = new System.Drawing.Size(438, 35);
            this.btnSetThreadSA.TabIndex = 4;
            this.btnSetThreadSA.Text = "Create System Aware Taskpane";
            this.btnSetThreadSA.UseVisualStyleBackColor = true;
            this.btnSetThreadSA.Click += new System.EventHandler(this.btnSetThreadSA_Click);
            // 
            // btnSetThreadPMA
            // 
            this.btnSetThreadPMA.Location = new System.Drawing.Point(14, 252);
            this.btnSetThreadPMA.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnSetThreadPMA.Name = "btnSetThreadPMA";
            this.btnSetThreadPMA.Size = new System.Drawing.Size(438, 35);
            this.btnSetThreadPMA.TabIndex = 5;
            this.btnSetThreadPMA.Text = "Set Thread to Per Monitor Aware";
            this.btnSetThreadPMA.UseVisualStyleBackColor = true;
            this.btnSetThreadPMA.Click += new System.EventHandler(this.btnSetThreadPMA_Click);
            // 
            // btnSetThreadPMAV2
            // 
            this.btnSetThreadPMAV2.Location = new System.Drawing.Point(14, 297);
            this.btnSetThreadPMAV2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnSetThreadPMAV2.Name = "btnSetThreadPMAV2";
            this.btnSetThreadPMAV2.Size = new System.Drawing.Size(438, 35);
            this.btnSetThreadPMAV2.TabIndex = 6;
            this.btnSetThreadPMAV2.Text = "Set Thread to Per Monitor Aware V2";
            this.btnSetThreadPMAV2.UseVisualStyleBackColor = true;
            this.btnSetThreadPMAV2.Click += new System.EventHandler(this.btnSetThreadPMAV2_Click);
            // 
            // btnOpenNonModalSA
            // 
            this.btnOpenNonModalSA.Location = new System.Drawing.Point(14, 342);
            this.btnOpenNonModalSA.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnOpenNonModalSA.Name = "btnOpenNonModalSA";
            this.btnOpenNonModalSA.Size = new System.Drawing.Size(438, 35);
            this.btnOpenNonModalSA.TabIndex = 7;
            this.btnOpenNonModalSA.Text = "Open Form, non-modal, system aware";
            this.btnOpenNonModalSA.UseVisualStyleBackColor = true;
            this.btnOpenNonModalSA.Click += new System.EventHandler(this.btnOpenNonModalSA_Click);
            // 
            // btnSetCWMM
            // 
            this.btnSetCWMM.Location = new System.Drawing.Point(14, 387);
            this.btnSetCWMM.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnSetCWMM.Name = "btnSetCWMM";
            this.btnSetCWMM.Size = new System.Drawing.Size(438, 35);
            this.btnSetCWMM.TabIndex = 8;
            this.btnSetCWMM.Text = "Set Child Window Mixed Mode to Mixed";
            this.btnSetCWMM.UseVisualStyleBackColor = true;
            this.btnSetCWMM.Click += new System.EventHandler(this.btnSetCWMM_Click);
            // 
            // txtProcessAwareness
            // 
            this.txtProcessAwareness.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtProcessAwareness.Location = new System.Drawing.Point(237, 41);
            this.txtProcessAwareness.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtProcessAwareness.Name = "txtProcessAwareness";
            this.txtProcessAwareness.Size = new System.Drawing.Size(213, 26);
            this.txtProcessAwareness.TabIndex = 10;
            // 
            // txtTaskpaneWindowAwareness
            // 
            this.txtTaskpaneWindowAwareness.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtTaskpaneWindowAwareness.Location = new System.Drawing.Point(237, 77);
            this.txtTaskpaneWindowAwareness.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtTaskpaneWindowAwareness.Name = "txtTaskpaneWindowAwareness";
            this.txtTaskpaneWindowAwareness.Size = new System.Drawing.Size(213, 26);
            this.txtTaskpaneWindowAwareness.TabIndex = 11;
            // 
            // txtHostWindowAwareness
            // 
            this.txtHostWindowAwareness.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtHostWindowAwareness.Location = new System.Drawing.Point(237, 113);
            this.txtHostWindowAwareness.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtHostWindowAwareness.Name = "txtHostWindowAwareness";
            this.txtHostWindowAwareness.Size = new System.Drawing.Size(213, 26);
            this.txtHostWindowAwareness.TabIndex = 12;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 5);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(142, 20);
            this.label1.TabIndex = 13;
            this.label1.Text = "Thread Awareness";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(8, 41);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(149, 20);
            this.label2.TabIndex = 14;
            this.label2.Text = "Process Awareness";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(8, 77);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(222, 20);
            this.label3.TabIndex = 15;
            this.label3.Text = "Taskpane Window Awareness";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(8, 113);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(186, 20);
            this.label4.TabIndex = 16;
            this.label4.Text = "Host Window Awareness";
            // 
            // chkAutoRefresh
            // 
            this.chkAutoRefresh.AutoSize = true;
            this.chkAutoRefresh.Location = new System.Drawing.Point(12, 152);
            this.chkAutoRefresh.Name = "chkAutoRefresh";
            this.chkAutoRefresh.Size = new System.Drawing.Size(183, 24);
            this.chkAutoRefresh.TabIndex = 17;
            this.chkAutoRefresh.Text = "Auto Refresh Values";
            this.chkAutoRefresh.UseVisualStyleBackColor = true;
            this.chkAutoRefresh.Click += new System.EventHandler(this.chkAutoRefresh_Click);
            // 
            // btnRefresh
            // 
            this.btnRefresh.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRefresh.Location = new System.Drawing.Point(237, 151);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(213, 33);
            this.btnRefresh.TabIndex = 18;
            this.btnRefresh.Text = "Refresh Now";
            this.btnRefresh.UseVisualStyleBackColor = true;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(14, 432);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(438, 35);
            this.button1.TabIndex = 19;
            this.button1.Text = "Set Child Window Mixed Mode to Normal";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.setCWMMNormal_Click);
            // 
            // UserControlWinForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnRefresh);
            this.Controls.Add(this.chkAutoRefresh);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtHostWindowAwareness);
            this.Controls.Add(this.txtTaskpaneWindowAwareness);
            this.Controls.Add(this.txtProcessAwareness);
            this.Controls.Add(this.btnSetCWMM);
            this.Controls.Add(this.btnOpenNonModalSA);
            this.Controls.Add(this.btnSetThreadPMAV2);
            this.Controls.Add(this.btnSetThreadPMA);
            this.Controls.Add(this.btnSetThreadSA);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.txtThreadAwareness);
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "UserControlWinForm";
            this.Size = new System.Drawing.Size(464, 877);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox txtThreadAwareness;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button btnSetThreadSA;
        private System.Windows.Forms.Button btnSetThreadPMA;
        private System.Windows.Forms.Button btnSetThreadPMAV2;
        private System.Windows.Forms.Button btnOpenNonModalSA;
        private System.Windows.Forms.Button btnSetCWMM;
        private System.Windows.Forms.TextBox txtProcessAwareness;
        private System.Windows.Forms.TextBox txtTaskpaneWindowAwareness;
        private System.Windows.Forms.TextBox txtHostWindowAwareness;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.CheckBox chkAutoRefresh;
        private System.Windows.Forms.Button btnRefresh;
        private System.Windows.Forms.Button button1;
    }
}
