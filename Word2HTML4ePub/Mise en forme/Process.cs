using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Core;
using System.Reflection;
using System.IO;
using System.Xml;
using System.Xml.XPath;
using System.Windows.Forms;

namespace Word2HTML4ePub
{
    public partial class WordHTML2ePubHTML
    {
        public delegate void ProcessEventHandler(string message);
        public static event ProcessEventHandler ReportLog;

        public delegate void ProgressEventHandler(int current, int max);
        public static event ProgressEventHandler Progress;

        volatile static string titre = null; //titre de la page dans le fichier html
        volatile static string fileNameHTML = null; //Nom du fichier html temporaire
        volatile static bool tempFile; //Flag pour conserver les fichiers temporaires
        volatile static string htmlFileName = null; //Fichier final
        volatile static List<string> FileNameList = null; //Liste des Fichiers à copier dans le package
        volatile static List<string> Manifest = null; //Texte à ajouter au manifest dans le package
        volatile static List<string> Spine = null; //Texte à ajouter à la spine dans le package


        /// <summary>
        /// Partie qui doit être executée dans le thread principal...
        /// </summary>
        /// <param name="doc"></param>
        public static bool PreProcessDoc(Microsoft.Office.Interop.Word.Document doc)
        {
            string PackagePath = "";
            ReportLog("Contrôle des paramètres");

            //Impossible de continuer si le doc n'a pas de titre...
            titre = GetDocProperty(doc, "Titre");
            if (string.IsNullOrEmpty(titre))
            {
                DialogResult dr = FormMetaData.EditEpubParam(ref doc);
                if (dr == DialogResult.Cancel)
                    return false;

                titre = GetDocProperty(doc, "Titre");
            }

            fileNameHTML = GetDocProperty(doc, "htmlFile");
            if (string.IsNullOrEmpty(fileNameHTML))
            {
                string fileName = titre.ToLower();
                foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                {
                    fileName = fileName.Replace(c, '_');
                }
                fileName = fileName.Replace('é', 'e');
                fileName = fileName.Replace('è', 'e');
                fileName = fileName.Replace('ê', 'e');
                fileName = fileName.Replace('à', 'a');
                fileName = fileName.Replace('ô', 'o');
                fileName = fileName.Replace('ù', 'u');
                fileName = fileName.Replace(" ", null);
                fileNameHTML = fileName + ".html";
                //MessageBox.Show("Impossible de traiter un fichier si les paramètres ne sont pas configurés...", "Paramètres manquants", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //return false;
                SetDocProperty(doc, "htmlFile", fileNameHTML);
            }

            string DossierSortie = GetDocProperty(doc, "DossierSortie");
            if (string.IsNullOrEmpty(DossierSortie))
            {
                PackagePath = WordHTML2ePubHTML.GetDocProperty(doc, "PackagePath");
                if (string.IsNullOrEmpty(PackagePath))
                {
                    MessageBox.Show("Impossible de traiter un fichier si les paramètres ne sont pas configurés...", "Paramètres manquants", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                DossierSortie = Path.Combine(PackagePath, "temp");
                if (!Directory.Exists(DossierSortie))
                    Directory.CreateDirectory(DossierSortie);
            }

            fileNameHTML = Path.Combine(DossierSortie, fileNameHTML);
            tempFile = GetDocFlag(doc, "tempFile");

            //0. Sauvegarde du nom du doc et verif s'il faut sauver le doc
            ReportLog("Début de la sauvegarde éventuelle du Document");
            string WordDoc = doc.FullName;
            if (!doc.Saved)
            {
                doc.Save();
                ReportLog("Document sauvegardé");
            }

            //0.1 Suppression des signets avant l'exportation
            int sig = doc.Bookmarks.Count;
            ReportLog("Suppression de " + sig.ToString() + " signets");
            doc.Bookmarks.ShowHidden = true;
            foreach (Word.Bookmark b in doc.Bookmarks)
                b.Delete();

            //0.2 Clean des espaces insécables
            CleanInsecableGauche(doc, "!");
            CleanInsecableGauche(doc, "?");
            CleanInsecableGauche(doc, ";");
            CleanInsecableGauche(doc, "»");
            CleanInsecableDroit(doc, "«");

            //0.3 Desactiver le bold sur les images
            foreach (Word.InlineShape sh in doc.InlineShapes)
            {
                sh.Range.set_Style(Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleNormal);
            }

            //0.4 Pré-traitement des notes de bas de page en note de fin de doc
            doc.Footnotes.Convert();

            //1. Exportation en HTML filtré
            ReportLog("Exportation en HTML (Word)");
            htmlFileName = SaveAsHTML(doc);
            if (string.IsNullOrEmpty(htmlFileName))
                return false;

            //2. Fermeture du fichier s'il s'agit du fichier ouvert... et réouverture du fichier original
            ReportLog("Ré-ouverture du fichier Word");
            if (doc.FullName != WordDoc)
            {
                object save = false;
                doc.Close(ref save, Type.Missing, Type.Missing);
                OpenFile(WordDoc);
            }

            ////3. Copie des fichiers images dans le dossier content.
            //string[] folderImg = Directory.GetDirectories(Path.GetDirectoryName(htmlFileName), Path.GetFileNameWithoutExtension(htmlFileName) + "_fi*.");
            //CopyImages(folderImg, PackagePath);
            return true;
        }


//        public static void ProcessDoc(Microsoft.Office.Interop.Word.Document doc)
        public static void ProcessDoc(FormMonitor.Decoupe decoupe, FormMonitor.TraitementImages traitementImg, int TailleMax, string PackagePath)
        {
            //3. Nettoyage du fichier html pour le rendre ouvrable... Merci Tidy...
            ReportLog("Nettoyage du fichier via Tidy");
            string parsedFN = WordHTML2ePubHTML.CleanHTMLFile(htmlFileName, fileNameHTML);
            if (string.IsNullOrEmpty(parsedFN))
                return;

            //4. Copie des fichiers images dans le dossier content.
            string[] folderImg = Directory.GetDirectories(Path.GetDirectoryName(htmlFileName), Path.GetFileNameWithoutExtension(htmlFileName) + "_fi*.");
            if (traitementImg == FormMonitor.TraitementImages.NoImage)
            { // Suppression des images
                ReportLog("Suppression des Images");
                DeleteImages(folderImg, PackagePath);
            }
            else
            {
                ReportLog("Copie des Images");
                CopyImages(folderImg, PackagePath);
            }


            ReportLog("Début du nettoyage spécifique");
            //5. Clean spécifique aux epub 
            WordHTML2ePubHTML.Clean4ePub(PackagePath, parsedFN, titre, decoupe, TailleMax, traitementImg, out FileNameList, out Manifest, out Spine);
            if (FileNameList.Count == 0)
            {
                ReportLog("Erreur durant le nettoyage spécifique!");
                return;
            }
            ReportLog("Fin du nettoyage spécifique");

            //6. Nettoyage des fichiers temporaires
            ReportLog("Suppression des fichiers temporaires");
            if (tempFile)
                File.Delete(htmlFileName);

            //7. Affichage de la fin du process
            //System.Windows.Forms.MessageBox.Show("Fin de process !");

            //8. Mise à jour du Content.OPF
            OPFFile.UpdateRessources(PackagePath, FileNameList);
            
        }
    }
}