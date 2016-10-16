using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.IO;
using System.Reflection;

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
            txtSize.Text = Properties.Settings.Default.TailleMaxKo.ToString();
            galDecoupe.SelectedItemIndex= 1;
            galDecoupe_Click(null, null);

            galImages.Items.Clear();
            foreach (FormMonitor.TraitementImages s in (FormMonitor.TraitementImages[])Enum.GetValues(typeof(FormMonitor.TraitementImages)))
            {
                Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDDI = this.Factory.CreateRibbonDropDownItem();
                ribbonDDI.Label = s.ToString();
                galImages.Items.Add(ribbonDDI);
            }
            galImages.SelectedItemIndex = 3;
            galImages_Click(null, null);
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

            EditParam frmEdit = new EditParam(ref doc);
            DialogResult dr = frmEdit.ShowDialog();
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

            FormMonitor.TraitementImages traitImg;
            Enum.TryParse<FormMonitor.TraitementImages>(galImages.SelectedItem.Label, out traitImg);


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

        private void btnCreatePack_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.Application.Documents.Count == 0)
                return;

            //Récupérer le document en cours d'édition
            Microsoft.Office.Interop.Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;


            //Choisir le dossier
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            //folderDlg.ShowNewFolderButton = true;
            folderDlg.Description = "Sélectionner un dossier VIDE pour le package" ;

            //Lire le paramètrage dans le fichier Word actif
            folderDlg.SelectedPath = WordHTML2ePubHTML.GetDocProperty(doc, "PackagePath");
            
            if (string.IsNullOrEmpty(folderDlg.SelectedPath))
                folderDlg.SelectedPath = Path.GetDirectoryName(doc.FullName);
            else
                if (MessageBox.Show("Un package a déjà été créé pour ce fichier :\r\n" + folderDlg.SelectedPath + "\r\nVoulez-vous recréer un package?", "Attention!", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK)
                    return;

            folderDlg.RootFolder = System.Environment.SpecialFolder.MyComputer;
            
            DialogResult dr = folderDlg.ShowDialog();
            if (dr != DialogResult.OK)
                return;

            //Charger le modèle
            Assembly _assembly = Assembly.GetExecutingAssembly();
            Stream _ePubFileStream = _assembly.GetManifestResourceStream("Word2HTML4ePub.Modele.epub");
            Ionic.Zip.ZipFile zip = Ionic.Zip.ZipFile.Read(_ePubFileStream);

            //Décompresser le modèle dans un sous folder
            zip.ExtractAll(folderDlg.SelectedPath, Ionic.Zip.ExtractExistingFileAction.DoNotOverwrite);
            
            //Ecrire le paramètrage dans le fichier Word actif
            WordHTML2ePubHTML.SetDocProperty(doc, "PackagePath", folderDlg.SelectedPath);

        }

        private void btnExportToPack_Click(object sender, RibbonControlEventArgs e)
        {
            string PackagePath = Word2ePub.GetCurrentDocPackageFolder();
            if (string.IsNullOrEmpty(PackagePath))
                return;

            //Récupérer le document en cours d'édition
            Microsoft.Office.Interop.Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            //Maj des métadatas de l'opf
            OPFFile.UpdatePackage(doc);
 
            //Extraction des paramètres de conversion
            //string PackagePath = WordHTML2ePubHTML.GetDocProperty(doc, "PackagePath");
            string title = WordHTML2ePubHTML.GetDocProperty(doc, "Titre");
            string auteur = WordHTML2ePubHTML.GetDocProperty(doc, "Auteur");
            string couvPath = WordHTML2ePubHTML.GetDocProperty(doc, "CoverPath");

            FormMonitor.Decoupe dec;
            Enum.TryParse<FormMonitor.Decoupe>(galDecoupe.SelectedItem.Label, out dec);
            int taillemax = 0;

            FormMonitor.TraitementImages traitImg;
            Enum.TryParse<FormMonitor.TraitementImages>(galImages.SelectedItem.Label, out traitImg);


            //Maj de la couverture
            if (!File.Exists(couvPath))
            {
                DialogResult dr = MessageBox.Show("Continuer la génération sans image?", "L'image de couverture est absente!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr == DialogResult.No)
                    return;
                couvPath = null;
            }

            //Modif de l'OPF
            if (WordHTML2ePubHTML.GetDocFlag(doc, "GenCover"))
            { 
                OPFFile.ChangeCover(PackagePath, couvPath);

                //Extraction des paramètres de conversion
                string coverFile = OPFFile.GetCoverFile(PackagePath);
                string author = WordHTML2ePubHTML.GetDocProperty(doc, "Auteur");
                WordHTML2ePubHTML.UpdateCoverHTMLFile(coverFile, couvPath, title, null, author);
            }

            //Lancement du process
            FormMonitor frmMon = new FormMonitor(doc, dec, traitImg, taillemax, PackagePath);
            frmMon.Show();
        }

        private void cmdGeneEPub_Click(object sender, RibbonControlEventArgs e)
        {
            //1. Get Package folder
            string PackagePath = Word2ePub.GetCurrentDocPackageFolder();
            if (string.IsNullOrEmpty(PackagePath))
                return;

            //Récupérer le document en cours d'édition
            Microsoft.Office.Interop.Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            string title = WordHTML2ePubHTML.GetDocProperty(doc, "Titre");
            string auteur = WordHTML2ePubHTML.GetDocProperty(doc, "Auteur");
            
            ////2. Generate ePub
            string ePubFile = Word2ePub.GenerateEPub(PackagePath);
            
            FormMonitor.TraitementImages traitImg;
            Enum.TryParse<FormMonitor.TraitementImages>(galImages.SelectedItem.Label, out traitImg);
            if (traitImg == FormMonitor.TraitementImages.NoImage)
            {
                if (File.Exists(ePubFile.Remove(ePubFile.Length - 5) + "_NoImages.epub"))
                    File.Delete(ePubFile.Remove(ePubFile.Length - 5) + "_NoImages.epub");
                File.Move(ePubFile, ePubFile.Remove(ePubFile.Length - 5) + "_NoImages.epub");
            }
        }

        private void cmdEditCss_Click(object sender, RibbonControlEventArgs e)
        {
            ////1. Get Package folder
            string PackagePath = Word2ePub.GetCurrentDocPackageFolder();
            if (string.IsNullOrEmpty(PackagePath))
                return;

            //Récupérer le document en cours d'édition
            Microsoft.Office.Interop.Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;


            //2. Get Css file
            string[] cssFiles = Directory.GetFiles(Path.Combine(PackagePath, "content"), "*.css", SearchOption.AllDirectories);
            if (cssFiles.Length == 0)
            {
                File.Create(Path.Combine(PackagePath, "content", "style.css")).Close();
                cssFiles = Directory.GetFiles(Path.Combine(PackagePath, "content"), "*.css", SearchOption.AllDirectories);
                //throw new Exception("Pas de fichier \".css\" dans le package!");
            }
            else if (cssFiles.Length >2)
                throw new Exception("plus d'un fichier \".css\" dans le package!");

            //3. Assurer la présence des balises présentes dans le Package
            StyleDocList sdl = StyleDocList.ReadStylesFromStyleFile(Path.Combine(PackagePath, "temp", "styles.txt"));
            CssFile.UpdateCssFile(cssFiles[0], sdl);

            //4. Edition du fichier css dans une apppli tierce...
            Word2ePub.EditCSSinNotepad(cssFiles[0]);
        }

        private void cmdAbout_Click(object sender, RibbonControlEventArgs e)
        {
            AboutBox frmAB = new AboutBox();
            frmAB.ShowDialog();
        }

        private void btnLoadePub_Click(object sender, RibbonControlEventArgs e)
        {
            //1. Choisir un fichier ePub
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Choisir un fichier ePub";
            ofd.Multiselect = false;
            ofd.Filter = "Fichiers ePub(*.ePub)|*.ePub";
            DialogResult dr = ofd.ShowDialog();
            if (dr != DialogResult.OK)
                return;

            //2. Extraire le nécessaire de l'ePub
            string destFile = Open_ePub.ConvertEpub2HTML(ofd.FileName);

            //3. Ouvrir le fichier dans Word
            WordHTML2ePubHTML.OpenHTMLFile(destFile);
        }

        private void galImages_Click(object sender, RibbonControlEventArgs e)
        {
            galImages.Label = "Traitement : " + galImages.SelectedItem.Label;

            FormMonitor.TraitementImages trait;
            Enum.TryParse<FormMonitor.TraitementImages>(galImages.SelectedItem.Label, out trait);


        }
    }
}
