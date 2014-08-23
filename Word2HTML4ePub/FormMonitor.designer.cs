namespace Word2HTML4ePub
{
    partial class FormMonitor
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
            this.cmdCancel = new System.Windows.Forms.Button();
            this.cmdQuitter = new System.Windows.Forms.Button();
            this.txtLog = new System.Windows.Forms.RichTextBox();
            this.Progress = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // cmdCancel
            // 
            this.cmdCancel.Location = new System.Drawing.Point(116, 133);
            this.cmdCancel.Name = "cmdCancel";
            this.cmdCancel.Size = new System.Drawing.Size(75, 23);
            this.cmdCancel.TabIndex = 1;
            this.cmdCancel.Text = "Abandonner";
            this.cmdCancel.UseVisualStyleBackColor = true;
            this.cmdCancel.Click += new System.EventHandler(this.cmdCancel_Click);
            // 
            // cmdQuitter
            // 
            this.cmdQuitter.Location = new System.Drawing.Point(197, 133);
            this.cmdQuitter.Name = "cmdQuitter";
            this.cmdQuitter.Size = new System.Drawing.Size(75, 23);
            this.cmdQuitter.TabIndex = 2;
            this.cmdQuitter.Text = "Quitter";
            this.cmdQuitter.UseVisualStyleBackColor = true;
            this.cmdQuitter.Click += new System.EventHandler(this.cmdQuitter_Click);
            // 
            // txtLog
            // 
            this.txtLog.Location = new System.Drawing.Point(13, 13);
            this.txtLog.Name = "txtLog";
            this.txtLog.ReadOnly = true;
            this.txtLog.Size = new System.Drawing.Size(259, 105);
            this.txtLog.TabIndex = 3;
            this.txtLog.Text = "";
            // 
            // Progress
            // 
            this.Progress.Location = new System.Drawing.Point(13, 133);
            this.Progress.Name = "Progress";
            this.Progress.Size = new System.Drawing.Size(100, 23);
            this.Progress.TabIndex = 4;
            // 
            // FormMonitor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 184);
            this.ControlBox = false;
            this.Controls.Add(this.Progress);
            this.Controls.Add(this.txtLog);
            this.Controls.Add(this.cmdQuitter);
            this.Controls.Add(this.cmdCancel);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(300, 200);
            this.Name = "FormMonitor";
            this.ShowIcon = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "Moniteur de tâche";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FormMonitor_FormClosed);
            this.Load += new System.EventHandler(this.FormMonitor_Load);
            this.Resize += new System.EventHandler(this.FormMonitor_Resize);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button cmdCancel;
        private System.Windows.Forms.Button cmdQuitter;
        private System.Windows.Forms.RichTextBox txtLog;
        private System.Windows.Forms.ProgressBar Progress;
    }
}