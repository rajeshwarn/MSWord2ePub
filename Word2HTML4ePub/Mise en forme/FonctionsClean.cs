using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Reflection;
using System.IO;
using System.Xml;
using System.Xml.XPath;

namespace Word2HTML4ePub
{
    public partial class WordHTML2ePubHTML
    {
        /// <summary>
        /// Nettoyer un fichier HTML en utilisant Tidy
        /// </summary>
        /// <param name="FullFileName"></param>
        /// <param name="FullFinalName"></param>
        /// <returns></returns>
        public static string CleanHTMLFile(string FullFileName, string FullFinalName)
        {
            string parsedFN = FullFinalName;
            try
            {

                if (string.IsNullOrEmpty(FullFinalName))
                    parsedFN = Path.Combine(Path.GetDirectoryName(FullFileName), "Parsed-" + Path.GetFileName(FullFileName));


                using (FileStream fsr = File.OpenRead(FullFileName))
                {
                    using (FileStream fsw = File.Open(parsedFN, FileMode.Create))
                    {
                        TidyNet.Tidy tidy = new TidyNet.Tidy();
                        TidyNet.TidyMessageCollection mess = new TidyNet.TidyMessageCollection();
                        tidy.Options.BreakBeforeBR = true;
                        tidy.Options.CharEncoding = TidyNet.CharEncoding.UTF8;
                        tidy.Options.DocType = TidyNet.DocType.Omit;
                        tidy.Options.DropEmptyParas = false;
                        tidy.Options.DropFontTags = true;
                        tidy.Options.EncloseBlockText = true;
                        tidy.Options.EncloseText = true;
                        tidy.Options.FixBackslash = true;
                        tidy.Options.FixComments = true;
                        tidy.Options.HideEndTags = true;
                        tidy.Options.IndentAttributes = false;
                        tidy.Options.IndentContent = true;
                        tidy.Options.LiteralAttribs = false;
                        tidy.Options.LogicalEmphasis = true;
                        tidy.Options.MakeClean = true;
                        tidy.Options.QuoteAmpersand = true;
                        tidy.Options.QuoteMarks = true;
                        tidy.Options.QuoteNbsp = false;
                        tidy.Options.TidyMark = true;
                        tidy.Options.Word2000 = false;
                        tidy.Options.Xhtml = true;

                        tidy.Parse(fsr, fsw, mess);

                        //TODO: lire les message de tidy...
                    }
                }
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de Cleaner le HTML!");
                return null;
            }
            return parsedFN;
        }

        public static void Clean4ePub(string ParsedFileName, string TitreDuDoc, FormMonitor.Decoupe decoupe, int TailleMaxKo)
        {
            DateTime LastUpdate = DateTime.Now;

            ReportLog("Chargement du fichier html exporté");
            //Ouverture du fichier html
            XmlDocument htmlFile = new XmlDocument();

            XmlElement root;
            try
            {
                htmlFile.Load(ParsedFileName);
                //XmlElement root = htmlFile.CreateElement(null, "html", "http://www.w3.org/1999/xhtml");
                //root.SetAttribute("epub", "http://www.idpf.org/2007/ops");
                //htmlFile.InsertAfter(root, htmlFile.FirstChild);
                htmlFile.InsertBefore(htmlFile.CreateXmlDeclaration("1.0", "UTF-8", null), htmlFile.FirstChild);
                htmlFile.InsertAfter(htmlFile.CreateDocumentType("html", null, null, null), htmlFile.FirstChild);
                root = htmlFile.DocumentElement;
                //root.SetAttribute("xmlns", "http://www.w3.org/1999/xhtml"); 
                //root.SetAttribute("epub", "http://www.idpf.org/2007/ops");
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de lire le fichier HTML purgé!");
                return;
            }

            System.Xml.XPath.XPathNavigator lir;
            System.Xml.XPath.XPathNodeIterator it;
            string exPath;
            try
            {
                //Creation d'un navigateur xpath
                lir = htmlFile.CreateNavigator();
                it = null;

                exPath = null;

                //Ajout du namespace epub
                it = lir.Select("./html");
                if (it.Count == 0)
                    return; // pas un doc html
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de naviguer dans le fichier HTML purgé!");
                return;
            }

            try
            {

                ReportLog("Suppression des balises meta");
                //Suppression des infos meta
                it = lir.Select("/html/head//meta");
                while (it.MoveNext())
                {
                    it.Current.DeleteSelf();
                    it = lir.Select("/html/head//meta");
                }

                ReportLog("Ajout d'une balise meta pour Word2ePub");
                //Ajout d'un meta pour identifier le générateur
                it = lir.Select("/html/head");
                it.MoveNext();
                it.Current.AppendChildElement(null, "meta", null, null);
                it = lir.Select("/html/head/meta");
                it.MoveNext();
                //    <meta name="Generator" value="Word2ePub_1.0.0.0"/>
                it.Current.CreateAttribute(null, "name", null, "Generateur");
                it.Current.CreateAttribute(null, "content", null, "Word2ePub_" + Assembly.GetExecutingAssembly().GetName().Version.ToString());

            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de supprimer les infos méta");
            }

            try
            {
                ReportLog("Copie des styles dans le fichier style.css");
                //Copie des styles dans un fichier css à posttraiter
                it = lir.Select("/html/head//style");
                it.MoveNext();
                File.WriteAllText(Path.Combine(Path.GetDirectoryName(ParsedFileName), "style.css"), it.Current.OuterXml);

                ReportLog("Suppression des balises html/head/styles");
                //Suppression de la section style 
                it = lir.Select("/html/head//style");
                while (it.MoveNext())
                {
                    it.Current.DeleteSelf();
                    it = lir.Select("/html/head//style");
                }
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de traiter la section style");
            }

            try
            {
                //Insertion de la balise style <link rel="Stylesheet" href="style.css"  type="text/css" />
                ReportLog("Ajout de la balise de style (style.css)");
                it = lir.Select("/html/head");
                it.MoveNext();
                it.Current.AppendChild("<link rel=\"Stylesheet\" href=\"style.css\"  type=\"text/css\" />");

                //Modification du titre du document
                ReportLog("Ajout du titre du document");
                it = lir.Select("/html/head/title");
                it.MoveNext();
                it.Current.SetValue(TitreDuDoc); //it.Current.SetValue("Le titre de la page");
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de modifier le titre et la mention au fichire css");
            }

            try
            {
                //suppression des balises contenant l'attribut style
                ReportLog("Purge des attributs style dans toutes les balises du body");
                it = lir.Select("/html/body//*[@style]");
                while (it.MoveNext())
                {
                    if (it.Current.Name.Equals("hr"))
                        it.Current.ReplaceSelf("<hr />");
                    else
                    {
                        //navigation jusqu'à l'attribut, et suppression
                        it.Current.MoveToFirstAttribute();
                        while (true)
                        {
                            if (string.Equals(it.Current.Name, "style"))
                            {
                                it.Current.DeleteSelf();
                                break;
                            }
                            else
                                it.Current.MoveToNextAttribute();
                        }
                        //it.Current.OuterXml = it.Current.InnerXml;
                    }
                }
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de supprimer certaines infos de style");
            }

            try
            {
                //suppression de l'attribut lang de la balise body
                ReportLog("Clean des attributs de la balise body");
                it = lir.Select("/html/body/@lang");
                it.MoveNext();
                it.Current.DeleteSelf();
                //Suppression de tout attribut de la balise body
                it = lir.Select("/html/body");
                it.MoveNext();
                while (it.Current.HasAttributes)
                {
                    it.Current.MoveToFirstAttribute();
                    it.Current.DeleteSelf();
                }

                //suppression de la balise div de la balise body
                ReportLog("Suppression de la balise div du body (uniquement la 1ère)");
                it = lir.Select("/html/body/div");
                if (it.MoveNext())
                {
                    it.Current.OuterXml = it.Current.InnerXml;
                }
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de supprimer les div en trop...");
            }

            try
            {

                int nbmax = 100;
                //suppression des attributs MsoNormal et MsoNoSpacing qui n'ont pas de raison d'être puisque ce sont les styles normaux
                ReportLog("Clean des attributs MsoNormal et MsoNoSpacing");
                exPath = "/html/body//*[@class='MsoNormal' or @class='MsoNoSpacing']";
                it = lir.Select(lir.Compile(exPath));
                nbmax = it.Count;
                while (it.MoveNext())
                {
                    if ((DateTime.Now - LastUpdate).Seconds > 1)
                    {
                        Progress(nbmax - it.Count, nbmax);
                        LastUpdate = DateTime.Now;
                    }
                    it.Current.MoveToAttribute("class", "");
                    it.Current.DeleteSelf();
                }
                Progress(nbmax, nbmax);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de supprimer MsoNormal et MsoNoSpacing");
            }

            try
            {
                //idem avec MsoPlainText 
                ReportLog("Clean des attributs MsoPlainText");
                exPath = "/html/body//*[@class='MsoPlainText']";
                it = lir.Select(lir.Compile(exPath));
                while (it.MoveNext())
                {
                    it.Current.MoveToAttribute("class", "");
                    it.Current.DeleteSelf();
                }
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de supprimer les MsoPlainText");
            }

            try
            {
                ReportLog("Remplacement des <quotes> par des <cites> (HTML5)");
                //Remplacement des MsoQuote par des cite
                exPath = "/html/body//*[@class='MsoQuote']";
                it = lir.Select(lir.Compile(exPath));
                while (it.MoveNext())
                {
                    if (it.Current.InnerXml.Length > 0)
                        it.Current.ReplaceSelf("<cite>" + it.Current.InnerXml + "</cite>");
                    else
                        it.Current.DeleteSelf();
                    it = lir.Select(lir.Compile(exPath));
                }
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de remplacer les quote et les <cite>");
            }

            try
            {
                ReportLog("Suppression des <span> vides");
                //suppression des span autre que ceux de class (qui ne devraient pas exister)
                exPath = "/html/body//span[not(@class)]";
                it = lir.Select(lir.Compile(exPath));
                while (it.MoveNext())
                {
                    if (it.Current.InnerXml.Length > 0)
                        it.Current.ReplaceSelf(it.Current.InnerXml);
                    else
                        it.Current.DeleteSelf();
                    it = lir.Select(lir.Compile(exPath));
                }
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de purger les span");
            }

            ReportLog("Suppression de l'indentation word");
            try
            {
                //Suppression des \r formatés par Word
                WordHTML2ePubHTML.RemoveIndent(ref lir);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de supprimer les indents");
            }

            ReportLog("Extraction des styles utilisés (styles.txt)");
            try
            {
                //Extraction des styles pour la mise en forme css
                ExtractStyleList(lir);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible d'extraire une toc");
            }

            ReportLog("Extraction de la TOC et création de la table de navigation");
            if (decoupe == FormMonitor.Decoupe.Aucun)
            {
                ReportLog("Extraction de la TOC");
                NavTable nav = null;
                try
                {
                    ////Extraction des titres pour la balise nav, puis ajout d'un id
                    nav = ExtractTOC(ref lir);
                    nav.ExportNavTable(ParsedFileName); //, ref lir
                }
                catch (Exception e)
                {
                    System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible d'extraire une toc");
                }

                try
                {
                    //Sauvegarde du fichier
                    htmlFile.Save(ParsedFileName);

                    //Modif du namespace dans le fichier (impossible à faire facilement en mode xml
                    {
                        string contenu = File.ReadAllText(ParsedFileName, Encoding.UTF8);
                        contenu = contenu.Replace("<html>",
                            "<html xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:epub=\"http://www.idpf.org/2007/ops\">");
                        File.WriteAllText(ParsedFileName, contenu, Encoding.UTF8);
                    }
                }
                catch (Exception e)
                {
                    System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de sauvegarder le fichier HTML purgé!");
                    return;
                }
            }
            else if (decoupe == FormMonitor.Decoupe.Chapitre)
            {
                ReportLog("Extraction de la TOC");
                NavTable nav = null;
                try
                {
                    ////Extraction des titres pour la balise nav, puis ajout d'un id
                    nav = ExtractTOC(ref lir);
                }
                catch (Exception e)
                {
                    System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible d'extraire une toc");
                }

                string ParsedFileNamePartial = null;
                try
                {
                    ReportLog("Split par Chapitre");
                    htmlFile.Save(ParsedFileName);

                    //Découpe du fichier en chapitres
                    XmlDocument htmlfinal = htmlFile;

                    string debutfichier = WordHTML2ePubHTML.getHTMLHeader(ParsedFileName);
                    string finfichier = "</body></html>";
                    ParsedFileNamePartial = Path.Combine(Path.GetDirectoryName(ParsedFileName), Path.GetFileNameWithoutExtension(ParsedFileName));
                    List<string> lofiles = new List<string>();

                    for (int i = 0; i <= nav.NbOfChap; i++)
                    {
                        string debut;
                        string fin;

                        nav.SplitTextes(i, out debut, out fin);

                        //Extraction des titres pour la balise nav, puis ajout d'un id
                        if (!string.IsNullOrEmpty(debut))
                        {
                            exPath = "/html/body/*[@id=\"" + debut + "\"]";
                            XPathNavigator nodeDebut = lir.SelectSingleNode(lir.Compile(exPath));
                            debut = nodeDebut.OuterXml;
                            int splitter = debut.IndexOf(">");
                            debut = debut.Substring(0, splitter);
                        }


                        if (!string.IsNullOrEmpty(fin))
                        {
                            exPath = "/html/body/*[@id=\"" + fin + "\"]";
                            XPathNavigator nodeFin = lir.SelectSingleNode(lir.Compile(exPath));
                            fin = nodeFin.OuterXml;
                            //                            lir.MoveTo(nodeFin);
                            int splitter = fin.IndexOf(">");
                            if (splitter < 0)
                            {

                            }
                            fin = fin.Substring(0, splitter);
                        }

                        string content = SplitHTMLFile(ParsedFileName, debut, fin);

                        if (string.IsNullOrEmpty(debut))
                        {
                            lofiles.Add(ParsedFileNamePartial + "-0.html");
                            File.WriteAllText(lofiles[lofiles.Count - 1], content + finfichier);
                        }
                        else if (string.IsNullOrEmpty(fin))
                        {
                            lofiles.Add(ParsedFileNamePartial + "-" + i.ToString() + ".html");
                            File.WriteAllText(lofiles[lofiles.Count - 1], debutfichier + content);
                        }
                        else
                        {
                            lofiles.Add(ParsedFileNamePartial + "-" + i.ToString() + ".html");
                            File.WriteAllText(lofiles[lofiles.Count - 1], debutfichier + content + finfichier);
                        }
                    }

                    ////Sauvegarde du fichier
                    //htmlFile.Save(ParsedFileName);

                    string opf = "<opf:manifest>\r\n";
                    string spine = "<opf:spine>\r\n";

                    //Modif du namespace dans le fichier (impossible à faire facilement en mode xml
                    for (int i = 0; i < lofiles.Count; i++)
                    {
                        string contenu = File.ReadAllText(lofiles[i], Encoding.UTF8);
                        //contenu = contenu.Replace("<html epub=\"http://www.idpf.org/2007/ops\">",
                        contenu = contenu.Replace("<html>",
                            "<html xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:epub=\"http://www.idpf.org/2007/ops\">");
                        File.WriteAllText(lofiles[i], contenu, Encoding.UTF8);

                        opf += "<opf:item id=\"Chap" + i.ToString() + "\" href=\"" + Path.GetFileName(lofiles[i]) + "\" media-type=\"application/xhtml+xml\" />\r\n";
                        spine += "<opf:itemref idref=\"Chap" + i.ToString() + "\" linear=\"yes\" />\r\n";
                    }
                    File.WriteAllText(ParsedFileNamePartial + "-opf.txt", opf + "\r\n" + spine);

                }
                catch (Exception e)
                {
                    System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de sauvegarder le fichier HTML purgé!");
                    return;
                }

                ReportLog("Création de la table de navigation");
                try
                {
                    if (nav != null)
                        nav.ExportNavTableSplittedbyChap(ParsedFileNamePartial); //, ref lir
                }
                catch (Exception e)
                {
                    System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible d'extraire une toc");
                }
            }

        }
    }
}