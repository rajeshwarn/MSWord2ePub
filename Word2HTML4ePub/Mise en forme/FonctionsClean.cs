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
                    parsedFN = Path.Combine(Path.GetDirectoryName(FullFileName).ToLower(), "Parsed-" + Path.GetFileName(FullFileName));


                using (FileStream fsr = File.OpenRead(FullFileName))
                {
                    using (FileStream fsw = File.Open(parsedFN, FileMode.Create))
                    {
                        ////Pour le package NuGet - ! Temps d'accès à la librairie trop long (7s...)
                        //Tidy.Core.TidyMessageCollection mess = new Tidy.Core.TidyMessageCollection();
                        //Tidy.Core.Tidy tidy = new Tidy.Core.Tidy();
                        //tidy.Options.BreakBeforeBr = true; //BreakBeforeBR = true;
                        //tidy.Options.CharEncoding = Tidy.Core.CharEncoding.Utf8; //TidyNet.CharEncoding.UTF8;
                        //tidy.Options.DocType = Tidy.Core.DocType.Omit; //TidyNet.DocType.Omit;
                        
                        //Pour la solution installée
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

        public static void Clean4ePub(string PackagePath, string ParsedFileName, string TitreDuDoc, 
            FormMonitor.Decoupe decoupe, int TailleMaxKo, 
            FormMonitor.TraitementImages traitementImg, 
            out List<string> FileName, 
            out List<string> Manifest, 
            out List<string> Spine)
        {
            DateTime LastUpdate = DateTime.Now;
            Manifest = new List<string>();
            Spine = new List<string>();
            FileName = new List<string>();

            ReportLog("Chargement du fichier html exporté");
            //Ouverture du fichier html
            XmlDocument htmlFile = new XmlDocument();

            XmlElement root;
            if (!MajHTMLHeaders(ParsedFileName, htmlFile, out root))
				return;
				
            System.Xml.XPath.XPathNavigator lir;
            System.Xml.XPath.XPathNodeIterator it;
            string exPath;


            /* 
                <aside epub:type="footnote" id="n1">
                <p>These have been corrected in this EPUB3 edition.</p>
                </aside>
             */

            if (!CreateNavigator(htmlFile, out lir, out it))
				return;

 			if (!DeleteMeta(ref lir, ref it))
				return;
			
 			if (!ExtractCssStyles(ParsedFileName, ref lir, ref it))
				return;

            if (!CleanScripts(ParsedFileName, ref lir, ref it))
                return;

            if (!AddStyleHeader(TitreDuDoc, ref lir, ref it))
				return;

            if (!RemoveDivClass(ref lir, ref it, "WordSection"))
                return;

            if (!RemoveDiv(ref lir, ref it))
                return;

            if (!PurgeStylesFromHTML(ref lir, ref it))
				return;
			
 			if (!RemoveLangFromBody(ref lir, ref it))
				return;
 
			if (!RemoveAttributeStartWith("MsoNormal", ref lir, ref it))
				return;

            if (!RemoveAttribute("MsoNoSpacing", ref lir, ref it))
				return;
 
            if (!RemoveAttribute("MsoPlainText", ref lir, ref it))
				return;
			
			if (!ReplaceQuotes(ref lir, ref it))
				return;

			if (!RemoveBalise("span", ref lir, ref it))
				return;

            if (!RemoveBaliseClass("span", "MsoCommentReference", ref lir, ref it))
                return;

            if (!RemoveCommentsFinalBlocs(ref lir, ref it))
                return;

            if (!WordHTML2ePubHTML.RemoveIndent(ref lir))
                return;
            
            if (traitementImg == FormMonitor.TraitementImages.NoImage)
            {
                if (!TraitementNoImages(ref lir, ref it))
                    return;
            }
            else if (traitementImg == FormMonitor.TraitementImages.Convert2SVG)
            {
                if (!TraitementImagesSVG(ref lir, ref it))
                    return;
            }
            else if (traitementImg == FormMonitor.TraitementImages.Resize600x800)
            {
                if (!TraitementImages600x800(PackagePath, ref lir, ref it))
                    return;
            }
            else 
            {
                if (!NoTraitementImages(PackagePath, ref lir, ref it))
                    return;
            }

            if (!ExtractStyleList(lir))
                return;

            ReportLog("Extraction de la TOC et création de la table de navigation");
            if (decoupe == FormMonitor.Decoupe.Aucun)
            {
                NavTable nav = ExtractTOC(ref lir);
                if (nav == null)
                    return;
                FileName.Add(nav.ExportNavTable(ParsedFileName)); //, ref lir

                try
                {
                    //Sauvegarde du fichier
                    htmlFile.Save(ParsedFileName);
                    FileName.Add(ParsedFileName);

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
                NavTable nav = ExtractTOC(ref lir);
                if (nav == null)
                    return;

                List<string> notes = new List<string>();

                //Extraction des Notes de bas de Page
                exPath = "/html/body//p[starts-with(@class,'MsoFootnoteText')]";
                XPathNavigator notesNode = lir.SelectSingleNode(lir.Compile(exPath));
                while (notesNode != null)
                {
                    //Garder pour copier dans un fichier de notes
                    notes.Add(notesNode.OuterXml);
                    //Suppression 
                    notesNode.DeleteSelf();

                    notesNode = lir.SelectSingleNode(lir.Compile(exPath));
                }
                
                //Extraction des Notes de Fin
                exPath = "/html/body//p[starts-with(@class,'MsoEndnoteText')]";
                notesNode = lir.SelectSingleNode(lir.Compile(exPath));
                while (notesNode != null)
                {
                    //Garder pour copier dans un fichier de notes
                    notes.Add(notesNode.OuterXml);
                    //Suppression 
                    notesNode.DeleteSelf();

                    notesNode = lir.SelectSingleNode(lir.Compile(exPath));
                }

                ////Extraction des Notes de Fin "Reference"
                //exPath = "/html/body//span[starts-with(@class,'MsoFootnoteReference')]";
                //notesNode = lir.SelectSingleNode(lir.Compile(exPath));
                //while (notesNode != null)
                //{
                //    //Garder pour copier dans un fichier de notes
                //    notes.Add("<p>" +notesNode.OuterXml + "</p>");
                //    //Suppression 

                //    notesNode.MoveToParent();
                //    notesNode.DeleteSelf();

                //    notesNode = lir.SelectSingleNode(lir.Compile(exPath));
                //}


                //Suppression du div clear des foot et end notes (s'il en existe...)
                exPath = "/html/body//*[@clear]";
                notesNode = lir.SelectSingleNode(lir.Compile(exPath));
                while (notesNode != null)
                {
//                    notesNode.MoveToParent();
                    notesNode.DeleteSelf();
                    notesNode = lir.SelectSingleNode(lir.Compile(exPath));
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
                    ParsedFileNamePartial = Path.Combine(Path.GetDirectoryName(ParsedFileName).ToLower(), Path.GetFileNameWithoutExtension(ParsedFileName));
                    List<string> lofiles = new List<string>();

                    if (string.IsNullOrEmpty(debutfichier))
                    { 
                        ReportLog("Impossible de spliter par chapitre!");
                        return;
                    }

                    string debut;
                    string fin;

                    for (int i = 0; i <= nav.NbOfChap; i++)
                    {
                        nav.SplitTextes(i, out debut, out fin);

                        //Extraction des titres pour la balise nav, puis ajout d'un id
                        if (!string.IsNullOrEmpty(debut))
                        {
                            exPath = "/html/body//*[@id=\"" + debut + "\"]";
                            XPathNavigator nodeDebut = lir.SelectSingleNode(lir.Compile(exPath));
                            debut = nodeDebut.OuterXml;
                            int splitter = debut.IndexOf(">");
                            debut = debut.Substring(0, splitter);
                        }


                        if (!string.IsNullOrEmpty(fin))
                        {
                            exPath = "/html/body//*[@id=\"" + fin + "\"]";
                            XPathNavigator nodeFin = lir.SelectSingleNode(lir.Compile(exPath));
                            if (nodeFin == null)
                                continue; // TODO : debug de ce cas bizare

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

                    //Création d'un fichier de notes
                    if (notes.Count != 0)
                    {
                        ReportLog("Création d'un fichier d'annotations");
                        string notesFile = ParsedFileNamePartial + "-notes.html";

                        //Chargement des notes
                        string notesFinales = "";
                        foreach (string note in notes)
                        {
                            XmlDocument notesDoc = new XmlDocument();
                            notesDoc.LoadXml(note);

                            //Creation d'un navigateur
                            XPathNavigator navNotes = notesDoc.CreateNavigator();

                            //recherche des notes
                            exPath = "//a[starts-with(@href,'#_')]";
                            XPathNavigator node = navNotes.SelectSingleNode(navNotes.Compile(exPath));
                            while (node != null)
                            {
                                string balise = node.GetAttribute("href", "");
                                string baliseID = node.GetAttribute("id", "");
                                bool balID = false;
                                if (string.IsNullOrEmpty(baliseID))
                                {
                                    if (balise.Contains("_ftnref"))
                                        baliseID = balise.Replace("ref", "");
                                    else if (balise.Contains("_ednref"))
                                        baliseID = balise.Replace("ref", "");
                                }
                                else
                                    balID = true;


                                baliseID = baliseID.Replace("#", "");
                                balise = balise.Replace("#", "");
                                node.MoveToAttribute("href", "");

                                //recherche de la cible
                                foreach (string s in lofiles)
                                {
                                    string cont = File.ReadAllText(s);
                                    int ind = cont.IndexOf(baliseID);
                                    int indIDsrc = cont.IndexOf(balise);
                                    if (ind > 0)
                                    {
                                        //Modif de la note
                                        node.SetValue(Path.GetFileName(s) + "#" + balise);
                                        node.MoveToParent();
                                        if (!balID)
                                            node.CreateAttribute("", "id", notesDoc.NamespaceURI, baliseID);
                                        if (node.MoveToAttribute("name", ""))
                                            node.DeleteSelf();
                                        notesFinales += notesDoc.OuterXml;

                                        //modif du fichier source
                                        if (indIDsrc != -1)
                                            cont = cont.Replace("#" + baliseID, Path.GetFileName(notesFile) + "#" + baliseID);
                                        else
                                            cont = cont.Replace("#" + baliseID, Path.GetFileName(notesFile) + "#" + baliseID + "\" id=\"" + balise);
                                        cont = cont.Replace("name=\"" + balise + "\" ", "");
                                        File.WriteAllText(s, cont);
                                        break;
                                    }
                                }
                                node = navNotes.SelectSingleNode(lir.Compile(exPath));
                            }

                        }
                        File.WriteAllText(notesFile, debutfichier + notesFinales + finfichier);
                        lofiles.Add(notesFile);

                    }

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
                        if (contenu.Contains("<svg"))
                            opf += "<opf:item id=\"Chap" + i.ToString() + "\" href=\"" + Path.GetFileName(lofiles[i]) + "\" media-type=\"application/xhtml+xml\" properties=\"svg\" />\r\n";
                        else
                            opf += "<opf:item id=\"Chap" + i.ToString() + "\" href=\"" + Path.GetFileName(lofiles[i]) + "\" media-type=\"application/xhtml+xml\" />\r\n";
                        spine += "<opf:itemref idref=\"Chap" + i.ToString() + "\" linear=\"yes\" />\r\n";
                        
                        Spine.Add("<opf:itemref idref=\"Chap" + i.ToString() + "\" linear=\"yes\" />");
                        FileName.Add(lofiles[i]);
                    }
                    File.WriteAllText(ParsedFileNamePartial + "-opf.txt", opf + "\r\n" + spine);
                }
                catch (Exception e)
                {
                    System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de sauvegarder le fichier HTML purgé!");
                    return;
                }

                if (nav != null)
                    FileName.Add(nav.ExportNavTableSplittedbyChap(ParsedFileNamePartial)); //, ref lir

                Manifest = FileName;
            }
        }
    }
}