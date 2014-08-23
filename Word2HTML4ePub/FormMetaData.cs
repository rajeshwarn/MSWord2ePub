using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.IO;


namespace Word2HTML4ePub
{
    public partial class FormMetaData : Form
    {
        Microsoft.Office.Interop.Word.Document doc;
        Microsoft.Office.Core.DocumentProperties properties;

        public static DialogResult EditEpubParam(ref Microsoft.Office.Interop.Word.Document doc)
        {
            //Affiche le formulaire
            FormMetaData frm = new FormMetaData(ref doc);
            return frm.ShowDialog();
        }

        public FormMetaData(ref Microsoft.Office.Interop.Word.Document doc)
        {
            InitializeComponent();
            this.doc = doc;
            properties = (Microsoft.Office.Core.DocumentProperties)doc.CustomDocumentProperties;
        }

        private void cmdSauvegarde_Click(object sender, EventArgs e)
        {
            SetData("Titre", txtTitre.Text);
            SetData("htmlFile", txthtml.Text);
            SetFlag("tempFile", chkFileTemp.Checked);

            doc.Saved = false;
            //this.DialogResult = DialogResult.OK;

        }

        private void cmdOK_Click(object sender, EventArgs e)
        {
            SetData("Titre", txtTitre.Text);
            SetData("htmlFile", txthtml.Text);
            SetFlag("tempFile", chkFileTemp.Checked);

             doc.Saved = false;
            this.DialogResult = DialogResult.OK;
        }

        private void FormMetaData_Load(object sender, EventArgs e)
        {
            txtTitre.Text = LoadData("Titre");
            txthtml.Text = LoadData("htmlFile");
            txtFolderOut.Text = LoadData("DossierSortie");
            chkFileTemp.Checked = LoadFlag("tempFile");

            if (!Directory.Exists(txtFolderOut.Text))
                txtFolderOut.Text = Path.GetDirectoryName(doc.FullName);
        }

        private void txtFolderOut_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBD = new FolderBrowserDialog();
            folderBD.Description = "Dossier de sortie";
            folderBD.SelectedPath = Path.GetDirectoryName(doc.FullName);
            folderBD.ShowNewFolderButton = true;
            DialogResult dr =  folderBD.ShowDialog();
            if (dr == DialogResult.Cancel)
                return;

            SetData("DossierSortie", folderBD.SelectedPath);
            txtFolderOut.Text = folderBD.SelectedPath;
        }

        private string LoadData(string param)
        {
            try
            {
                return (string)properties[param].Value;
            }
            catch (Exception ex)
            {
                properties.Add(param, false, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, "", null);
            }
            return (string)properties[param].Value;
        }

        private void SetData(string param, string value)
        {
            try
            {
                properties[param].Value = value;
            }
            catch (Exception ex)
            {
                properties.Add(param, false, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, value, null);
            }
        }

        private bool LoadFlag(string param)
        {
            try
            {
                return (bool)properties[param].Value;
            }
            catch (Exception ex)
            {
                properties.Add(param, false, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeBoolean, true, null);
            }
            return (bool)properties[param].Value;
        }

        private void SetFlag(string param, bool value)
        {
            try
            {
                properties[param].Value = value;
            }
            catch (Exception ex)
            {
                properties.Add(param, false, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeBoolean, value, null);
            }
        }
    }
}
