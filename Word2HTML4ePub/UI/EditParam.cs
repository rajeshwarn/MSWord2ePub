using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Word2HTML4ePub
{
    public partial class EditParam : Form
    {
        bool change = false;
        Microsoft.Office.Interop.Word.Document doc;

        public EditParam(ref Microsoft.Office.Interop.Word.Document doc)
        {
            InitializeComponent();
            this.doc = doc;
            change = doc.Saved;

            txtTitle.Text = WordHTML2ePubHTML.GetDocProperty(doc, "Titre");
            txtAuteur.Text = WordHTML2ePubHTML.GetDocProperty(doc, "Auteur");
            txtEditeur.Text = WordHTML2ePubHTML.GetDocProperty(doc, "Editeur");
            txtGUID.Text = WordHTML2ePubHTML.GetDocProperty(doc, "GUID");
            datTimePick.Value = WordHTML2ePubHTML.GetDocDateTime(doc, "DateCreation");
            txtSujet.Text = WordHTML2ePubHTML.GetDocProperty(doc, "Sujet");
            txtDescription.Text  = WordHTML2ePubHTML.GetDocProperty(doc, "Description");
            string langue = WordHTML2ePubHTML.GetDocProperty(doc, "Language");
            for (int i = 0; i < cmbLangue.Items.Count; i++)
            {
                if (string.Equals((string)cmbLangue.Items[i], langue))
                {
                    cmbLangue.SelectedIndex = i;
                    break;
                }
            }

            chkGenCouv.Checked = WordHTML2ePubHTML.GetDocFlag(doc, "GenCover");
            cmdCouv.Tag = WordHTML2ePubHTML.GetDocProperty(doc, "CoverPath");
            if (File.Exists((string)cmdCouv.Tag))
                PicBoxCover.Load((string)cmdCouv.Tag);

            string licence = WordHTML2ePubHTML.GetDocProperty(doc, "Licence");
            for (int i = 0; i < cmbLicence.Items.Count; i++)
            {
                if (string.Equals((string)cmbLicence.Items[i], licence))
                {
                    cmbLicence.SelectedIndex = i;
                    break;
                }
            }

            txtPackage.Text = WordHTML2ePubHTML.GetDocProperty(doc, "PackagePath");
        }

        private void cmdUID_Click(object sender, EventArgs e)
        {
            if (txtGUID.Text.Length != 0)
            { 
                DialogResult dr = MessageBox.Show("Changer l'UID d'un ePub déjà généré peut entrainer une confusion des liseuses!\r\nEtes-vous sur de vouloir changer l'UID?", "Changement d'Unique ID?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr == System.Windows.Forms.DialogResult.No)
                    return;
            }
            txtGUID.Text = Guid.NewGuid().ToString().ToUpper();
            WordHTML2ePubHTML.SetDocProperty(doc, "GUID", txtGUID.Text);
        }

        private void cmdSave_Click(object sender, EventArgs e)
        {
            WordHTML2ePubHTML.SetDocProperty(doc, "Titre", txtTitle.Text);
            WordHTML2ePubHTML.SetDocProperty(doc, "Auteur", txtAuteur.Text);
            WordHTML2ePubHTML.SetDocProperty(doc, "Editeur", txtEditeur.Text);
            WordHTML2ePubHTML.SetDocDateTime(doc, "DateCreation", datTimePick.Value);
            WordHTML2ePubHTML.SetDocProperty(doc, "Sujet", txtSujet.Text);
            WordHTML2ePubHTML.SetDocProperty(doc, "Description", txtDescription.Text);
            WordHTML2ePubHTML.SetDocProperty(doc, "Language", (string)cmbLangue.SelectedItem);
            WordHTML2ePubHTML.SetDocProperty(doc, "PackagePath", txtPackage.Text);
            WordHTML2ePubHTML.SetDocProperty(doc, "Licence", (string)cmbLicence.SelectedItem);
            WordHTML2ePubHTML.SetDocFlag(doc, "GenCover", chkGenCouv.Checked);
            DialogResult = System.Windows.Forms.DialogResult.OK;
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            if (change)
                doc.Saved = true;
            DialogResult = System.Windows.Forms.DialogResult.Cancel;
        }

        private void cmdCouv_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            fd.CheckFileExists = true;
            fd.CheckPathExists = true;
            fd.Filter = "Image Files|*.bmp;*.jp*;*.gif;*.png";
            fd.Multiselect = false;
            fd.InitialDirectory = txtPackage.Text;
            fd.Title = "Choisir l'image de couverture";

            DialogResult dr = fd.ShowDialog();

            if (dr == System.Windows.Forms.DialogResult.Cancel)
                return;
            
            string cover = fd.FileName;
            WordHTML2ePubHTML.SetDocProperty(doc, "CoverPath", cover);
            cmdCouv.Tag = cover;
            PicBoxCover.Load(cover);

        }

        private void cmdPackFolder_Click(object sender, EventArgs e)
        {
            if (txtPackage.Text.Length != 0)
            {
                DialogResult dr = MessageBox.Show("Vous allez changer le répertoire de destination de l'ePub!\r\nEtes-vous sur de vous?", "Changement de répertoire de Package?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr == System.Windows.Forms.DialogResult.No)
                    return;
            }

            bool valid = false;
            FolderBrowserDialog fd = new FolderBrowserDialog();
            fd.Description = "Choisir un nouveau dossier pour le Package";
            fd.ShowNewFolderButton = false;
            if (Directory.Exists(txtPackage.Text))
                fd.SelectedPath = txtPackage.Text;
            else
                fd.SelectedPath=doc.Path;

            while (!valid)
            {
                DialogResult dr = fd.ShowDialog();
                if (dr != System.Windows.Forms.DialogResult.OK)
                    return;

                string newOpfPath = Path.Combine(fd.SelectedPath, "content");
                if (!Directory.Exists(newOpfPath))
                { 
                    dr = MessageBox.Show("Le dossier :\r\n" + fd.SelectedPath + "\r\nne peut être un dossier Package (il ne possède pas de sous-dossier \"content\"", "Mauvais dossier", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                    if (dr == System.Windows.Forms.DialogResult.Retry)
                        continue;
                    else
                        return;
                }

                if (!(Directory.GetFiles(newOpfPath, "*.opf", SearchOption.AllDirectories)).Any())
                {
                    dr = MessageBox.Show("Le dossier :\r\n" + fd.SelectedPath + "\r\nne peut être un dossier Package (il ne contient pas de fichier \".opf\"", "Mauvais dossier", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                    if (dr == System.Windows.Forms.DialogResult.Retry)
                        continue;
                    else
                        return;
                }
                else
                    valid = true;

            }

            txtPackage.Text = fd.SelectedPath;
            WordHTML2ePubHTML.SetDocProperty(doc, "PackagePath", txtPackage.Text);
        }

        private void cmdReloadPackage_Click(object sender, EventArgs e)
        {
            // Lecture de l'OPF
            List<OPFFile.OPFData> data = OPFFile.ReadPackage(txtPackage.Text);

            // Update GUI
            txtTitle.Text = (from OPFFile.OPFData d in data where d.id.Equals("dc:title") select d.valeur).FirstOrDefault();
            txtGUID.Text = (from OPFFile.OPFData d in data where d.id.Equals("dc:identifier") select d.valeur).FirstOrDefault();
            if (txtGUID.Text.Length>8)
                txtGUID.Text = txtGUID.Text.Substring(9);
            txtAuteur.Text = (from OPFFile.OPFData d in data where d.id.Equals("dc:creator") select d.valeur).FirstOrDefault();
            string date = (from OPFFile.OPFData d in data where d.id.Equals("dc:date") select d.valeur).FirstOrDefault();
            if (!string.IsNullOrEmpty(date))
                datTimePick.Value = Convert.ToDateTime(date);
            txtDescription.Text = (from OPFFile.OPFData d in data where d.id.Equals("dc:description") select d.valeur).FirstOrDefault();
            txtSujet.Text = (from OPFFile.OPFData d in data where d.id.Equals("dc:subject") select d.valeur).FirstOrDefault();
            txtEditeur.Text = (from OPFFile.OPFData d in data where d.id.Equals("dc:publisher") select d.valeur).FirstOrDefault();
            string lang = (from OPFFile.OPFData d in data where d.id.Equals("dc:language") select d.valeur).FirstOrDefault();
            
            if (string.Equals(lang.ToLower(), "fr"))
                lang = "Francais";
            else if (string.Equals(lang.ToLower(), "en"))
                lang = "English";
            else if (string.Equals(lang.ToLower(), "de"))
                lang = "German";
            else if (string.Equals(lang.ToLower(), "sp"))
                lang = "Spanich";

            for (int i = 0; i < cmbLangue.Items.Count; i++)
            {
                if (string.Equals((string)cmbLangue.Items[i], lang))
                {
                    cmbLangue.SelectedIndex = i;
                    break;
                }
            }

            chkGenCouv.Checked = WordHTML2ePubHTML.GetDocFlag(doc, "GenCover");
        }

        private void EditParam_FormClosed(object sender, FormClosedEventArgs e)
        {
            //Libère les ressources, notament le fichier image
            System.GC.Collect();
        }

        private void cmdEraseCover_Click(object sender, EventArgs e)
        {
            WordHTML2ePubHTML.SetDocProperty(doc, "CoverPath", "");
            cmdCouv.Tag = "";
            PicBoxCover.Image = null;
        }

        private void chkGenCouv_CheckedChanged(object sender, EventArgs e)
        {
            cmdCouv.Enabled = chkGenCouv.Checked;
            cmdEraseCover.Enabled = chkGenCouv.Checked;
            if (chkGenCouv.Checked)
            { 
                if (File.Exists((string)cmdCouv.Tag))
                    PicBoxCover.Load((string)cmdCouv.Tag);
            }
            else
                PicBoxCover.Image = null;

        }
    }
}
