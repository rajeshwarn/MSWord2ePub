using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Word2HTML4ePub
{
    public class Word2ePub
    {
        /// <summary>
        /// Fonction pour aller chercher le PackageFolder dans le doc actif (fonction récurrente)
        /// </summary>
        /// <returns>PackagePath existant</returns>
        internal static string GetCurrentDocPackageFolder()
        {
            if (Globals.ThisAddIn.Application.Documents.Count == 0)
                return null;

            //Récupérer le document en cours d'édition
            Microsoft.Office.Interop.Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            //Extraction du Paramètre
            string PackagePath = WordHTML2ePubHTML.GetDocProperty(doc, "PackagePath");

            if (string.IsNullOrEmpty(PackagePath))
            {
                MessageBox.Show("Attention, les paramètres nécessaires ne sont pas correctement entrés!\r\n Créez un nouveau package avant de pouvoir utiliser cette commande.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            else
                PackagePath = PackagePath.ToLower();

            if (!Directory.Exists(PackagePath))
            {
                MessageBox.Show("Avez-vous réorganiser vos dossier?\r\nRé-assignez le Dossier Package en éditant la configuration.", "Le dossier Package n'existe pas!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return null;
            }
            return PackagePath;
        }


        /// <summary>
        /// Automation de la génération, pour appel depuis l'UI
        /// </summary>
        /// <param name="PackageFolder">Path vers le Package</param>
        /// <returns>true : ePub Well formed</returns>
        public static string GenerateEPub(string PackageFolder)
        {
            string epubFile = null;
            string ErrorLog = null;
            if (!ePubTools.CreateePub(new string[] {PackageFolder}, out epubFile, out ErrorLog))
            { // Echec de la génération : Affichage du message
                MessageBox.Show(ErrorLog, "Erreur lors de la génération", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return epubFile;
            }

            // Génération réussie!
            if (ePubTools.CheckPostGen)
            {// Check new file                
                string resultCheck = ePubTools.CheckEPub(epubFile);
                if (string.IsNullOrEmpty(resultCheck))
                {
                    MessageBox.Show("ePub Well Formed!\r\n" + epubFile, "Génération réussie", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    return epubFile;
                }
                else
                {
                    CustomMessageBox msg = new CustomMessageBox("Erreurs dans l'epub", resultCheck, new string[] { "ERROR:", "WARNING:", "Check finished" });
                    msg.ShowDialog();
                    return epubFile;
                }
            }
            return epubFile;
        }

        public static void EditCSSinNotepad(string CssFileFullPath)
        {
            if (!File.Exists(CssFileFullPath))
            { 
                return;
            }
            System.Diagnostics.Process.Start("notepad.exe", CssFileFullPath);
        }
    }
}
