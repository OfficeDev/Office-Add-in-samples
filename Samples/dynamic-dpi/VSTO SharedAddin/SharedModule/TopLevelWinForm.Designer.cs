namespace SharedModule
{
    partial class TopLevelWinForm
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.userControlWinForm1 = new SharedModule.UserControlMain();
            this.SuspendLayout();
            // 
            // userControlWinForm1
            // 
            this.userControlWinForm1.AutoScroll = true;
            this.userControlWinForm1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.userControlWinForm1.Location = new System.Drawing.Point(0, 0);
            this.userControlWinForm1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.userControlWinForm1.MinimumSize = new System.Drawing.Size(317, 0);
            this.userControlWinForm1.Name = "userControlWinForm1";
            this.userControlWinForm1.Size = new System.Drawing.Size(439, 453);
            this.userControlWinForm1.TabIndex = 0;
            // 
            // TopLevelWinForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(439, 453);
            this.Controls.Add(this.userControlWinForm1);
            this.Name = "TopLevelWinForm";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "DDPI top-level form";
            this.ResumeLayout(false);

        }

        #endregion

        private UserControlMain userControlWinForm1;
    }
}