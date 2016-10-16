using System;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.IO;
using System.Windows.Forms;

namespace Word2HTML4ePub
{
    public partial class WordHTML2ePubHTML
    {
        public static void OpenHTMLFile(string fileName)
        {
            if (!File.Exists(fileName))
                return;

            object newFileName = fileName;

            Globals.ThisAddIn.Application.Documents.Open(
                ref newFileName, (object)true, (object)false,
                (object)false, Type.Missing, Type.Missing, (object)false,Type.Missing,
                Type.Missing, Microsoft.Office.Interop.Word.WdOpenFormat.wdOpenFormatWebPages, 
                Type.Missing,(object)true, (object)false, Type.Missing, Type.Missing, Type.Missing);
            //Microsoft.Office.Interop.Word.Documents.Open(ref object, [ref object], [ref object], [ref object], [ref object], [ref object], [ref object], [ref object], [ref object], [ref object], [ref object], [ref object], [ref object], [ref object], [ref object], [ref object])

        }

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
            string htmlFileName = GetDocProperty(doc, "htmlFile");
            object newFileName = Path.Combine(Path.GetDirectoryName(doc.FullName), Path.GetFileNameWithoutExtension(doc.Name) + ".html");
            if (!string.IsNullOrEmpty(htmlFileName))
                newFileName = Path.Combine(Path.GetDirectoryName(doc.FullName), htmlFileName);

            object htmlFileFormat = Word.WdSaveFormat.wdFormatFilteredHTML;
            object LockComments = false;
            object Encoding = Office.MsoEncoding.msoEncodingUTF8;
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
                doc.WebOptions.AllowPNG = true;
                doc.WebOptions.BrowserLevel = Word.WdBrowserLevel.wdBrowserLevelV4;
                doc.WebOptions.RelyOnCSS = true;
                doc.WebOptions.TargetBrowser = Office.MsoTargetBrowser.msoTargetBrowserIE4;

                //Ajoutées pour la sauvegarde des images...
                doc.WebOptions.ScreenSize = Office.MsoScreenSize.msoScreenSize1024x768;
                doc.WebOptions.PixelsPerInch = 150;

                doc.WebOptions.Encoding = Office.MsoEncoding.msoEncodingUTF8;
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
        public static string GetDocProperty(Word.Document doc, string PropName)
        {
            Microsoft.Office.Core.DocumentProperties properties = (Microsoft.Office.Core.DocumentProperties)doc.CustomDocumentProperties;
            try
            {
                if (CheckPropertyExist(properties, PropName))
                {
                    if (properties[PropName].Type != Office.MsoDocProperties.msoPropertyTypeString)
                        throw new Exception("Pas une string...");
                }
                else
                {
                    properties.Add(PropName, false, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, "", null);
                    doc.Saved = false;
                }
            }
            catch (Exception ex)
            {
                return null;
            }
            return (string)properties[PropName].Value;
        }

        static private bool CheckPropertyExist(Microsoft.Office.Core.DocumentProperties properties, string PropName)
        {
            foreach (Office.DocumentProperty prop in properties)
            {
                if (prop.Name.Equals(PropName))
                    return true;
            }
            return false;
        }


        /// <summary>
        /// Ecrit un paramètre dans le fichier word
        /// </summary>
        /// <param name="doc">fichier word doc</param>
        /// <param name="PropName">le nom du paramètre</param>
        /// <param name="value">sa valeur</param>
        public static void SetDocProperty(Microsoft.Office.Interop.Word.Document doc, string PropName, string value)
        {
            try
            {
                Microsoft.Office.Core.DocumentProperties properties = (Microsoft.Office.Core.DocumentProperties)doc.CustomDocumentProperties;
                if (CheckPropertyExist(properties, PropName))
                {
                    if (!string.Equals(properties[PropName].Value, value))
                    {
                        properties[PropName].Value = value;
                        doc.Saved = false;
                    }
                }
                else
                {
                    properties.Add(PropName, false, Office.MsoDocProperties.msoPropertyTypeString, value);
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// Recupérer un flag dans un fichier, s'il existe...
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="PropName"></param>
        /// <returns></returns>
        public static bool GetDocFlag(Microsoft.Office.Interop.Word.Document doc, string PropName)
        {
            Microsoft.Office.Core.DocumentProperties properties = (Microsoft.Office.Core.DocumentProperties)doc.CustomDocumentProperties;
            try
            {
                if (CheckPropertyExist(properties, PropName))
                {
                    if (properties[PropName].Type != Office.MsoDocProperties.msoPropertyTypeBoolean)
                        throw new Exception("Pas un booléen...");
                }
                else
                {
                    properties.Add(PropName, false, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeBoolean, false, null);
                    doc.Saved = false;
                }
            }
            catch (Exception ex)
            {
                return false;
            }
            return (bool)properties[PropName].Value;
        }

        /// <summary>
        /// Ecrit un paramètre dans le fichier word
        /// </summary>
        /// <param name="doc">fichier word doc</param>
        /// <param name="PropName">le nom du paramètre</param>
        /// <param name="value">sa valeur</param>
        public static void SetDocFlag(Microsoft.Office.Interop.Word.Document doc, string PropName, bool value)
        {
            try
            {
                Microsoft.Office.Core.DocumentProperties properties = (Microsoft.Office.Core.DocumentProperties)doc.CustomDocumentProperties;
                if (CheckPropertyExist(properties, PropName))
                {
                    if (properties[PropName].Value != value)
                    { 
                        properties[PropName].Value = value;
                        doc.Saved = false;
                    }
                }
                else
                {
                    properties.Add(PropName, false, Office.MsoDocProperties.msoPropertyTypeBoolean, value);
                    doc.Saved = false;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// Récupère une propriété de date dans le fichier
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="PropName"></param>
        /// <returns></returns>
        public static DateTime GetDocDateTime(Microsoft.Office.Interop.Word.Document doc, string PropName)
        {
            Microsoft.Office.Core.DocumentProperties properties = (Microsoft.Office.Core.DocumentProperties)doc.CustomDocumentProperties;
            try
            {
                if (CheckPropertyExist(properties, PropName))
                {
                    if (properties[PropName].Type != Office.MsoDocProperties.msoPropertyTypeDate)
                        throw new Exception("Pas une date...");
                }
                else
                { 
                    properties.Add(PropName, false, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeDate, DateTime.Now, null);
                    doc.Saved = false;
                }
            }
            catch (Exception ex)
            {
                return DateTime.Now;
            }
            return properties[PropName].Value;
        }

        /// <summary>
        /// Créer ou mettre à jour un paramètre DateTime dans un fichier word
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="PropName"></param>
        /// <param name="value"></param>
        public static void SetDocDateTime(Microsoft.Office.Interop.Word.Document doc, string PropName, DateTime value)
        {
            try
            {
                Microsoft.Office.Core.DocumentProperties properties = (Microsoft.Office.Core.DocumentProperties)doc.CustomDocumentProperties;
                if (CheckPropertyExist(properties, PropName))
                {
                    if (properties[PropName].Value != value)
                    {
                        properties[PropName].Value = value;
                        doc.Saved = false;
                    }
                }
                else
                {
                    properties.Add(PropName, false, Office.MsoDocProperties.msoPropertyTypeDate, value);
                    doc.Saved = false;
                }
            }
            catch (Exception ex)
            {
                throw;
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