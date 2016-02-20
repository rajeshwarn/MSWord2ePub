namespace Word2HTML4ePub
{
    partial class InternalEditor
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
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.webVisu = new System.Windows.Forms.WebBrowser();
            this.txtSrc = new System.Windows.Forms.TextBox();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.webVisu, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.txtSrc, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 82.14936F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 17.85064F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(710, 549);
            this.tableLayoutPanel1.TabIndex = 1;
            // 
            // webVisu
            // 
            this.webVisu.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webVisu.Location = new System.Drawing.Point(358, 3);
            this.webVisu.MinimumSize = new System.Drawing.Size(20, 20);
            this.webVisu.Name = "webVisu";
            this.webVisu.Size = new System.Drawing.Size(349, 444);
            this.webVisu.TabIndex = 0;
            // 
            // txtSrc
            // 
            this.txtSrc.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtSrc.Location = new System.Drawing.Point(3, 3);
            this.txtSrc.Multiline = true;
            this.txtSrc.Name = "txtSrc";
            this.txtSrc.Size = new System.Drawing.Size(349, 444);
            this.txtSrc.TabIndex = 1;
            // 
            // InternalEditor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(710, 549);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "InternalEditor";
            this.Text = "InternalEditor";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.WebBrowser webVisu;
        private System.Windows.Forms.TextBox txtSrc;
    }
}