using System;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Core;
using System.IO;
using System.Windows.Forms;

namespace Word2HTML4ePub
{
    public partial class WordHTML2ePubHTML
    {
        //public static DialogResult EditEpubParam(ref Microsoft.Office.Interop.Word.Document doc)
        //{
        //    //Affiche le formulaire
        //    FormMetaData frm = new FormMetaData(ref doc);
        //    return frm.ShowDialog();
        //}

        /// <summary>
        /// Ouvrir automatiquement un fichier word... fonction utilisée pour le debug.
        /// </summary>
        /// <param name="fileName">Path complet</param>
        private static void OpenFile(string fileName)
        {
            //ouvrir le doc de test
            object newFileName = fileName;
            
            Globals.ThisAddIn.Application.Documents.Open(ref newFileName,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }

        /// <summary>
        /// Enregistrer un fichier sous forme html avec des options pour l'export epub...
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        private static string SaveAsHTML(Microsoft.Office.Interop.Word.Document doc)
        {
            object newFileName = Path.Combine(Path.GetDirectoryName(doc.FullName), Path.GetFileNameWithoutExtension(doc.Name) + ".html");
            object htmlFileFormat = Word.WdSaveFormat.wdFormatFilteredHTML;
            object LockComments = false;
            object Encoding = MsoEncoding.msoEncodingUTF8;
            object InsertLineBreaks = false;
            object lineEnd = Word.WdLineEndingType.wdLFOnly;
            object AddToRecentFiles = false;
            object ReadOnlyRecommended = false;
            object EmbedTrueTypeFonts = false;
            object SaveNativePictureFormat = false;
            object SaveFormsData = false;
            object SaveAsAOCELetter = false;
            object AllowSubstitutions = false;
            object LineEnding = Word.WdLineEndingType.wdCRLF;
            object AddBiDiMarks = false;

            try
            {
                doc.SaveAs(ref newFileName, ref htmlFileFormat, ref LockComments,
                    Type.Missing, ref AddToRecentFiles, Type.Missing,
                    ref ReadOnlyRecommended, ref EmbedTrueTypeFonts,
                    ref SaveNativePictureFormat, ref SaveFormsData,
                    ref SaveAsAOCELetter, ref Encoding, ref InsertLineBreaks,
                    ref AllowSubstitutions, ref LineEnding, ref AddBiDiMarks);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de Convertir en HTML!");
                return null;
            }
            return (string)newFileName;
        }


        /// <summary>
        /// Recupérer une propriété dans un doc word, si elle existe...
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="PropName"></param>
        /// <returns></returns>
        private static string GetDocProperty(Microsoft.Office.Interop.Word.Document doc, string PropName)
        {
            try
            {
                Microsoft.Office.Core.DocumentProperties properties = (Microsoft.Office.Core.DocumentProperties)doc.CustomDocumentProperties;
                if (properties[PropName].Type != MsoDocProperties.msoPropertyTypeString)
                    throw new Exception("Pas de titre...");

                return (string)properties[PropName].Value;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        /// <summary>
        /// Recupérer un flag dans un fichier, s'il existe...
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="PropName"></param>
        /// <returns></returns>
        private static bool GetDocFlag(Microsoft.Office.Interop.Word.Document doc, string PropName)
        {
            try
            {
                Microsoft.Office.Core.DocumentProperties properties = (Microsoft.Office.Core.DocumentProperties)doc.CustomDocumentProperties;
                if (properties[PropName].Type != MsoDocProperties.msoPropertyTypeBoolean)
                    throw new Exception("Pas de titre...");

                return (bool)properties[PropName].Value;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private static void CleanInsecableGauche(Microsoft.Office.Interop.Word.Document doc, string car)
        {
            try
            {
                Word.Find findObj = doc.Application.Selection.Find;
                findObj.ClearFormatting();
                findObj.Replacement.ClearFormatting();
                findObj.Text = " " + car;
                findObj.Replacement.Text = car;

                doc.Application.Selection.Find.ClearFormatting();
                object replaceAll = Word.WdReplace.wdReplaceAll;
                findObj.Execute(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    ref replaceAll, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                findObj.Text = "^s" + car;
                doc.Application.Selection.Find.ClearFormatting();
                replaceAll = Word.WdReplace.wdReplaceAll;
                findObj.Execute(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    ref replaceAll, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                findObj.Text = car;
                findObj.Replacement.Text = "^s" + car;

                doc.Application.Selection.Find.ClearFormatting();
                replaceAll = Word.WdReplace.wdReplaceAll;
                findObj.Execute(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    ref replaceAll, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                return;
            }
            catch (Exception ex)
            {
                return;
            }
        }

        private static void CleanInsecableDroit(Microsoft.Office.Interop.Word.Document doc, string car)
        {
            try
            {
                Word.Find findObj = doc.Application.Selection.Find;
                findObj.ClearFormatting();
                findObj.Replacement.ClearFormatting();
                findObj.Text = car + " ";
                findObj.Replacement.Text = car;

                doc.Application.Selection.Find.ClearFormatting();
                object replaceAll = Word.WdReplace.wdReplaceAll;
                findObj.Execute(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    ref replaceAll, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                findObj.Text = car+"^s";
                doc.Application.Selection.Find.ClearFormatting();
                replaceAll = Word.WdReplace.wdReplaceAll;
                findObj.Execute(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    ref replaceAll, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                findObj.Text = car;
                findObj.Replacement.Text = car + "^s";

                doc.Application.Selection.Find.ClearFormatting();
                replaceAll = Word.WdReplace.wdReplaceAll;
                findObj.Execute(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    ref replaceAll, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                return;
            }
            catch (Exception ex)
            {
                return;
            }
        }
    }
}