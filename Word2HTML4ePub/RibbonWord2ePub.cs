using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace Word2HTML4ePub
{
    public partial class RibbonWord2ePub
    {
        private void RibbonWord2ePub_Load(object sender, RibbonUIEventArgs e)
        {
            galDecoupe.Items.Clear();
            foreach (FormMonitor.Decoupe s in (FormMonitor.Decoupe[])Enum.GetValues(typeof(FormMonitor.Decoupe)))
            {
                Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDDI = this.Factory.CreateRibbonDropDownItem();
                ribbonDDI.Label = s.ToString();
                galDecoupe.Items.Add(ribbonDDI);
            }
            galDecoupe.SelectedItemIndex= 1;
            galDecoupe_Click(null, null);
        }

        private void btnAPropos_Click(object sender, RibbonControlEventArgs e)
        {
            AboutBox frm = new AboutBox();
            frm.ShowDialog();
        }

        private void btnConfig_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.Application.Documents.Count == 0)
                return;

            //Récupérer le document en cours d'édition
            Microsoft.Office.Interop.Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            DialogResult dr = FormMetaData.EditEpubParam(ref doc);

            if (dr == DialogResult.Cancel)
                return;
            else if (dr == DialogResult.Abort)
            {
                try
                {

                    if (!doc.Saved)
                        doc.Save();
                }
                catch (Exception ex)
                {

                }
            }
            else if (dr == DialogResult.OK)
            {
                btnConvert_Click(sender, e);
            }
        }

        private void btnConvert_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.Application.Documents.Count == 0)
                return;

            //Récupérer le document en cours d'édition
            Microsoft.Office.Interop.Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            //Extraction des paramètres de conversion
            FormMonitor.Decoupe dec;
            Enum.TryParse<FormMonitor.Decoupe>(galDecoupe.SelectedItem.Label, out dec);
            int taillemax = 0;
            if (dec == FormMonitor.Decoupe.ChapitresTailleMax)
                try
                {
                    taillemax = Convert.ToInt32(txtSize.Text);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            

            //Lancement du process
            FormMonitor frmMon = new FormMonitor(doc, dec, taillemax);
            frmMon.Show();

            // WordHTML2ePubHTML.ProcessDoc(doc);
        }

        private void galDecoupe_Click(object sender, RibbonControlEventArgs e)
        {
            galDecoupe.Label = "Découpe : " + galDecoupe.SelectedItem.Label;
            FormMonitor.Decoupe dec;
            Enum.TryParse<FormMonitor.Decoupe>(galDecoupe.SelectedItem.Label, out dec);

            txtSize.Enabled = (dec == FormMonitor.Decoupe.ChapitresTailleMax);

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            string ret = WordHTML2ePubHTML.getHTMLHeader("C:\\Users\\Emilie\\Desktop\\epub\\Ferdinand fabre\\Abbé tigrane\\abbe_tigrane.html");

            //WordHTML2ePubHTML.SplitHTMLFile(
            //    "C:\\Users\\Emilie\\Desktop\\epub\\Ferdinand fabre\\Abbé tigrane\\abbe_tigrane.html",
            //    "C:\\Users\\Emilie\\Desktop\\epub\\Ferdinand fabre\\Abbé tigrane\\abbe_tigrane1.html",
            //    "<h1 id=\"id.1.0.0.0.0.0.0\">I. UNE VILLE DÉVOTE. </h1>",
            //    "<h1 id=\"id.2.0.0.0.0.0.0\">II. MONSEIGNEUR DE ROQUEBRUN. </h1>");
        }

        private void btnCreatePack_Click(object sender, RibbonControlEventArgs e)
        {
            //Choisir le dossier
            InternalEditor frminted = new InternalEditor();
            frminted.Show();
            //Décompresser le modèle dans un sous folder

            //Ecrire le paramètrage dans le fichier Word

        }
    }
}
