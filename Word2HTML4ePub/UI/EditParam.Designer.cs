namespace Word2HTML4ePub
{
    partial class EditParam
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
            this.label1 = new System.Windows.Forms.Label();
            this.txtTitle = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtAuteur = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtEditeur = new System.Windows.Forms.TextBox();
            this.cmdUID = new System.Windows.Forms.Button();
            this.txtGUID = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.datTimePick = new System.Windows.Forms.DateTimePicker();
            this.label5 = new System.Windows.Forms.Label();
            this.txtSujet = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtDescription = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.cmbLangue = new System.Windows.Forms.ComboBox();
            this.PicBoxCover = new System.Windows.Forms.PictureBox();
            this.txtPackage = new System.Windows.Forms.TextBox();
            this.cmdCouv = new System.Windows.Forms.Button();
            this.cmdPackFolder = new System.Windows.Forms.Button();
            this.cmdSave = new System.Windows.Forms.Button();
            this.cmdCancel = new System.Windows.Forms.Button();
            this.cmdReloadPackage = new System.Windows.Forms.Button();
            this.label8 = new System.Windows.Forms.Label();
            this.cmbLicence = new System.Windows.Forms.ComboBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.chkGenCouv = new System.Windows.Forms.CheckBox();
            this.cmdEraseCover = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.PicBoxCover)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(28, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Titre";
            // 
            // txtTitle
            // 
            this.txtTitle.Location = new System.Drawing.Point(100, 10);
            this.txtTitle.Name = "txtTitle";
            this.txtTitle.Size = new System.Drawing.Size(270, 20);
            this.txtTitle.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 39);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(38, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "Auteur";
            // 
            // txtAuteur
            // 
            this.txtAuteur.Location = new System.Drawing.Point(100, 36);
            this.txtAuteur.Name = "txtAuteur";
            this.txtAuteur.Size = new System.Drawing.Size(270, 20);
            this.txtAuteur.TabIndex = 2;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 65);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(40, 13);
            this.label3.TabIndex = 0;
            this.label3.Text = "Editeur";
            // 
            // txtEditeur
            // 
            this.txtEditeur.Location = new System.Drawing.Point(100, 62);
            this.txtEditeur.Name = "txtEditeur";
            this.txtEditeur.Size = new System.Drawing.Size(270, 20);
            this.txtEditeur.TabIndex = 3;
            // 
            // cmdUID
            // 
            this.cmdUID.Location = new System.Drawing.Point(12, 86);
            this.cmdUID.Name = "cmdUID";
            this.cmdUID.Size = new System.Drawing.Size(73, 23);
            this.cmdUID.TabIndex = 4;
            this.cmdUID.Text = "GUID";
            this.cmdUID.UseVisualStyleBackColor = true;
            this.cmdUID.Click += new System.EventHandler(this.cmdUID_Click);
            // 
            // txtGUID
            // 
            this.txtGUID.Location = new System.Drawing.Point(100, 88);
            this.txtGUID.Name = "txtGUID";
            this.txtGUID.ReadOnly = true;
            this.txtGUID.Size = new System.Drawing.Size(270, 20);
            this.txtGUID.TabIndex = 1;
            this.txtGUID.TabStop = false;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(13, 118);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(30, 13);
            this.label4.TabIndex = 0;
            this.label4.Text = "Date";
            // 
            // datTimePick
            // 
            this.datTimePick.Location = new System.Drawing.Point(100, 114);
            this.datTimePick.Name = "datTimePick";
            this.datTimePick.Size = new System.Drawing.Size(270, 20);
            this.datTimePick.TabIndex = 5;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(13, 143);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(31, 13);
            this.label5.TabIndex = 0;
            this.label5.Text = "Sujet";
            // 
            // txtSujet
            // 
            this.txtSujet.Location = new System.Drawing.Point(100, 140);
            this.txtSujet.Name = "txtSujet";
            this.txtSujet.Size = new System.Drawing.Size(270, 20);
            this.txtSujet.TabIndex = 6;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(13, 169);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(60, 13);
            this.label6.TabIndex = 0;
            this.label6.Text = "Description";
            // 
            // txtDescription
            // 
            this.txtDescription.Location = new System.Drawing.Point(100, 166);
            this.txtDescription.Multiline = true;
            this.txtDescription.Name = "txtDescription";
            this.txtDescription.Size = new System.Drawing.Size(270, 88);
            this.txtDescription.TabIndex = 7;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(13, 263);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(43, 13);
            this.label7.TabIndex = 0;
            this.label7.Text = "Langue";
            // 
            // cmbLangue
            // 
            this.cmbLangue.FormattingEnabled = true;
            this.cmbLangue.Items.AddRange(new object[] {
            "Francais",
            "English",
            "German",
            "Spanich"});
            this.cmbLangue.Location = new System.Drawing.Point(100, 260);
            this.cmbLangue.Name = "cmbLangue";
            this.cmbLangue.Size = new System.Drawing.Size(270, 21);
            this.cmbLangue.TabIndex = 8;
            // 
            // PicBoxCover
            // 
            this.PicBoxCover.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.PicBoxCover.Location = new System.Drawing.Point(100, 310);
            this.PicBoxCover.Name = "PicBoxCover";
            this.PicBoxCover.Size = new System.Drawing.Size(270, 155);
            this.PicBoxCover.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.PicBoxCover.TabIndex = 9;
            this.PicBoxCover.TabStop = false;
            // 
            // txtPackage
            // 
            this.txtPackage.Location = new System.Drawing.Point(100, 499);
            this.txtPackage.Name = "txtPackage";
            this.txtPackage.ReadOnly = true;
            this.txtPackage.Size = new System.Drawing.Size(270, 20);
            this.txtPackage.TabIndex = 10;
            this.txtPackage.TabStop = false;
            // 
            // cmdCouv
            // 
            this.cmdCouv.Location = new System.Drawing.Point(12, 310);
            this.cmdCouv.Name = "cmdCouv";
            this.cmdCouv.Size = new System.Drawing.Size(73, 23);
            this.cmdCouv.TabIndex = 10;
            this.cmdCouv.Text = "Couverture";
            this.cmdCouv.UseVisualStyleBackColor = true;
            this.cmdCouv.Click += new System.EventHandler(this.cmdCouv_Click);
            // 
            // cmdPackFolder
            // 
            this.cmdPackFolder.Location = new System.Drawing.Point(12, 497);
            this.cmdPackFolder.Name = "cmdPackFolder";
            this.cmdPackFolder.Size = new System.Drawing.Size(73, 23);
            this.cmdPackFolder.TabIndex = 13;
            this.cmdPackFolder.Text = "Package";
            this.cmdPackFolder.UseVisualStyleBackColor = true;
            this.cmdPackFolder.Click += new System.EventHandler(this.cmdPackFolder_Click);
            // 
            // cmdSave
            // 
            this.cmdSave.Location = new System.Drawing.Point(295, 532);
            this.cmdSave.Name = "cmdSave";
            this.cmdSave.Size = new System.Drawing.Size(75, 23);
            this.cmdSave.TabIndex = 14;
            this.cmdSave.Text = "Enregistrer";
            this.cmdSave.UseVisualStyleBackColor = true;
            this.cmdSave.Click += new System.EventHandler(this.cmdSave_Click);
            // 
            // cmdCancel
            // 
            this.cmdCancel.Location = new System.Drawing.Point(214, 532);
            this.cmdCancel.Name = "cmdCancel";
            this.cmdCancel.Size = new System.Drawing.Size(75, 23);
            this.cmdCancel.TabIndex = 15;
            this.cmdCancel.Text = "Annuler";
            this.cmdCancel.UseVisualStyleBackColor = true;
            this.cmdCancel.Click += new System.EventHandler(this.cmdCancel_Click);
            // 
            // cmdReloadPackage
            // 
            this.cmdReloadPackage.Location = new System.Drawing.Point(12, 532);
            this.cmdReloadPackage.Name = "cmdReloadPackage";
            this.cmdReloadPackage.Size = new System.Drawing.Size(153, 23);
            this.cmdReloadPackage.TabIndex = 16;
            this.cmdReloadPackage.Text = "Relire les infos du Package";
            this.cmdReloadPackage.UseVisualStyleBackColor = true;
            this.cmdReloadPackage.Click += new System.EventHandler(this.cmdReloadPackage_Click);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(13, 474);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(45, 13);
            this.label8.TabIndex = 0;
            this.label8.Text = "Licence";
            // 
            // cmbLicence
            // 
            this.cmbLicence.FormattingEnabled = true;
            this.cmbLicence.Items.AddRange(new object[] {
            "Aucune",
            "CC BY-NC-SA 3.0 FR"});
            this.cmbLicence.Location = new System.Drawing.Point(100, 471);
            this.cmbLicence.Name = "cmbLicence";
            this.cmbLicence.Size = new System.Drawing.Size(270, 21);
            this.cmbLicence.TabIndex = 12;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // chkGenCouv
            // 
            this.chkGenCouv.AutoSize = true;
            this.chkGenCouv.Checked = true;
            this.chkGenCouv.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkGenCouv.Location = new System.Drawing.Point(12, 287);
            this.chkGenCouv.Name = "chkGenCouv";
            this.chkGenCouv.Size = new System.Drawing.Size(147, 17);
            this.chkGenCouv.TabIndex = 9;
            this.chkGenCouv.Text = "Génération de couverture";
            this.chkGenCouv.UseVisualStyleBackColor = true;
            this.chkGenCouv.CheckedChanged += new System.EventHandler(this.chkGenCouv_CheckedChanged);
            // 
            // cmdEraseCover
            // 
            this.cmdEraseCover.Location = new System.Drawing.Point(12, 339);
            this.cmdEraseCover.Name = "cmdEraseCover";
            this.cmdEraseCover.Size = new System.Drawing.Size(73, 42);
            this.cmdEraseCover.TabIndex = 11;
            this.cmdEraseCover.Text = "Effacer couverture";
            this.cmdEraseCover.UseVisualStyleBackColor = true;
            this.cmdEraseCover.Click += new System.EventHandler(this.cmdEraseCover_Click);
            // 
            // EditParam
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(384, 562);
            this.Controls.Add(this.chkGenCouv);
            this.Controls.Add(this.cmdReloadPackage);
            this.Controls.Add(this.cmdCancel);
            this.Controls.Add(this.cmdSave);
            this.Controls.Add(this.PicBoxCover);
            this.Controls.Add(this.cmbLicence);
            this.Controls.Add(this.cmbLangue);
            this.Controls.Add(this.datTimePick);
            this.Controls.Add(this.cmdPackFolder);
            this.Controls.Add(this.cmdEraseCover);
            this.Controls.Add(this.cmdCouv);
            this.Controls.Add(this.cmdUID);
            this.Controls.Add(this.txtGUID);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtDescription);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.txtSujet);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtEditeur);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtAuteur);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtPackage);
            this.Controls.Add(this.txtTitle);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "EditParam";
            this.ShowIcon = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "Paramètres";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.EditParam_FormClosed);
            ((System.ComponentModel.ISupportInitialize)(this.PicBoxCover)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtTitle;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtAuteur;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtEditeur;
        private System.Windows.Forms.Button cmdUID;
        private System.Windows.Forms.TextBox txtGUID;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DateTimePicker datTimePick;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtSujet;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtDescription;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox cmbLangue;
        private System.Windows.Forms.PictureBox PicBoxCover;
        private System.Windows.Forms.TextBox txtPackage;
        private System.Windows.Forms.Button cmdCouv;
        private System.Windows.Forms.Button cmdPackFolder;
        private System.Windows.Forms.Button cmdSave;
        private System.Windows.Forms.Button cmdCancel;
        private System.Windows.Forms.Button cmdReloadPackage;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox cmbLicence;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.CheckBox chkGenCouv;
        private System.Windows.Forms.Button cmdEraseCover;

    }
}