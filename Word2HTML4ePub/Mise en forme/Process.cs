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
        //static Timer ticker = null;

        /// <summary>
        /// Partie qui doit être executée dans le thread principal...
        /// </summary>
        /// <param name="doc"></param>
        public static bool PreProcessDoc(Microsoft.Office.Interop.Word.Document doc)
        {
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
                MessageBox.Show("Impossible de traiter un fichier si les paramètres ne sont pas configurés...", "Paramètres manquants", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            string DossierSortie = GetDocProperty(doc, "DossierSortie");
            if (string.IsNullOrEmpty(DossierSortie))
            {
                MessageBox.Show("Impossible de traiter un fichier si les paramètres ne sont pas configurés...", "Paramètres manquants", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
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
            foreach (Word.Bookmark b in doc.Bookmarks)
                b.Delete();

            //0.2 Clean des espaces insécables
            CleanInsecableGauche(doc, "!");
            CleanInsecableGauche(doc, "?");
            CleanInsecableGauche(doc, ";");
            CleanInsecableGauche(doc, "»");
            CleanInsecableDroit(doc, "«");

            //1. Exportation en HTML filtré
            ReportLog("Exportation en HTML (Word)");
            htmlFileName = SaveAsHTML(doc);
            if (string.IsNullOrEmpty(htmlFileName))
                return false;

            //2.Fermeture du fichier s'il s'agit du fichier ouvert... et réouverture du fichier original
            ReportLog("Ré-ouverture du fichier Word");
            if (doc.FullName != WordDoc)
            {
                object save = false;
                doc.Close(ref save, Type.Missing, Type.Missing);
                OpenFile(WordDoc);
            }
            return true;
        }

//        public static void ProcessDoc(Microsoft.Office.Interop.Word.Document doc)
        public static void ProcessDoc(FormMonitor.Decoupe decoupe, int TailleMax)
        {
            //3. Nettoyage du fichier html pour le rendre ouvrable... Merci Tidy...
            ReportLog("Nettoyage du fichier via Tidy");
            string parsedFN = WordHTML2ePubHTML.CleanHTMLFile(htmlFileName, fileNameHTML);
            if (string.IsNullOrEmpty(parsedFN))
                return;

            ReportLog("Début du nettoyage spécifique");
            //4. Clean spécifique aux epub 
            WordHTML2ePubHTML.Clean4ePub(parsedFN, titre, decoupe, TailleMax);
            ReportLog("Fin du nettoyage spécifique");

            //5. Nettoyage des fichiers temporaires
            ReportLog("Suppression des fichiers temporaires");
            if (tempFile)
                File.Delete(htmlFileName);

            //6. Affichage de la fin du process
            //System.Windows.Forms.MessageBox.Show("Fin de process !");
        }
    }
}