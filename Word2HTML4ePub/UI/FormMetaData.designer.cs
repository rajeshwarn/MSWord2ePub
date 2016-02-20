namespace Word2HTML4ePub
{
    partial class FormMetaData
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
            this.txtTitre = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cmdOK = new System.Windows.Forms.Button();
            this.cmdCancel = new System.Windows.Forms.Button();
            this.txthtml = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.txtFolderOut = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.chkFileTemp = new System.Windows.Forms.CheckBox();
            this.cmdSauvegarde = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtTitre
            // 
            this.txtTitre.Location = new System.Drawing.Point(103, 10);
            this.txtTitre.Name = "txtTitre";
            this.txtTitre.Size = new System.Drawing.Size(176, 20);
            this.txtTitre.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(87, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Titre de la page :";
            // 
            // cmdOK
            // 
            this.cmdOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.cmdOK.Location = new System.Drawing.Point(197, 128);
            this.cmdOK.Name = "cmdOK";
            this.cmdOK.Size = new System.Drawing.Size(75, 23);
            this.cmdOK.TabIndex = 6;
            this.cmdOK.Text = "Conversion";
            this.cmdOK.UseVisualStyleBackColor = true;
            this.cmdOK.Click += new System.EventHandler(this.cmdOK_Click);
            // 
            // cmdCancel
            // 
            this.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cmdCancel.Location = new System.Drawing.Point(13, 128);
            this.cmdCancel.Name = "cmdCancel";
            this.cmdCancel.Size = new System.Drawing.Size(75, 23);
            this.cmdCancel.TabIndex = 7;
            this.cmdCancel.Text = "Annuler";
            this.cmdCancel.UseVisualStyleBackColor = true;
            // 
            // txthtml
            // 
            this.txthtml.Location = new System.Drawing.Point(122, 36);
            this.txthtml.Name = "txthtml";
            this.txthtml.Size = new System.Drawing.Size(157, 20);
            this.txthtml.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(10, 39);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(106, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "Nom du fichier html  :";
            // 
            // txtFolderOut
            // 
            this.txtFolderOut.Location = new System.Drawing.Point(103, 62);
            this.txtFolderOut.Name = "txtFolderOut";
            this.txtFolderOut.Size = new System.Drawing.Size(176, 20);
            this.txtFolderOut.TabIndex = 3;
            this.txtFolderOut.Click += new System.EventHandler(this.txtFolderOut_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(10, 65);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(91, 13);
            this.label3.TabIndex = 0;
            this.label3.Text = "Dossier de sortie :";
            // 
            // chkFileTemp
            // 
            this.chkFileTemp.AutoSize = true;
            this.chkFileTemp.Checked = true;
            this.chkFileTemp.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkFileTemp.Location = new System.Drawing.Point(13, 88);
            this.chkFileTemp.Name = "chkFileTemp";
            this.chkFileTemp.Size = new System.Drawing.Size(197, 17);
            this.chkFileTemp.TabIndex = 4;
            this.chkFileTemp.Text = "Suppression des fichiers temporaires";
            this.chkFileTemp.UseVisualStyleBackColor = true;
            // 
            // cmdSauvegarde
            // 
            this.cmdSauvegarde.DialogResult = System.Windows.Forms.DialogResult.Abort;
            this.cmdSauvegarde.Location = new System.Drawing.Point(103, 128);
            this.cmdSauvegarde.Name = "cmdSauvegarde";
            this.cmdSauvegarde.Size = new System.Drawing.Size(75, 23);
            this.cmdSauvegarde.TabIndex = 5;
            this.cmdSauvegarde.Text = "Sauvegarder";
            this.cmdSauvegarde.UseVisualStyleBackColor = true;
            this.cmdSauvegarde.Click += new System.EventHandler(this.cmdSauvegarde_Click);
            // 
            // FormMetaData
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 163);
            this.Controls.Add(this.cmdSauvegarde);
            this.Controls.Add(this.chkFileTemp);
            this.Controls.Add(this.cmdCancel);
            this.Controls.Add(this.cmdOK);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txthtml);
            this.Controls.Add(this.txtFolderOut);
            this.Controls.Add(this.txtTitre);
            this.Name = "FormMetaData";
            this.Text = "Configuration";
            this.Load += new System.EventHandler(this.FormMetaData_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtTitre;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button cmdOK;
        private System.Windows.Forms.Button cmdCancel;
        private System.Windows.Forms.TextBox txthtml;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.TextBox txtFolderOut;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox chkFileTemp;
        private System.Windows.Forms.Button cmdSauvegarde;
    }
}