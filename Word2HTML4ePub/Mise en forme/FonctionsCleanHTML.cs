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

		private static bool MajHTMLHeaders(string ParsedFileName, XmlDocument htmlFile, out XmlElement root)
		{
            try
            {
                //Clean <!--[if gte mso 9]>
                string filecontent = File.ReadAllText(ParsedFileName);
                int start = filecontent.IndexOf("<![if mso 9]>");
                if (start > 0)
                {
                    int end = filecontent.IndexOf("<![endif]>", start);
                    if (end > 0)
                        filecontent = filecontent.Remove(start, end + 10 - start);
                }
                File.WriteAllText(ParsedFileName, filecontent);
                
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
                root = null;
                return false;
            }
			return true;
		}

        private static bool CreateNavigator(XmlDocument htmlFile, out System.Xml.XPath.XPathNavigator lir, out System.Xml.XPath.XPathNodeIterator it)
		{
			try
			{
				//Creation d'un navigateur xpath
				lir = htmlFile.CreateNavigator();
				it = null;

				//Ajout du namespace epub
				it = lir.Select("./html");
				if (it.Count == 0)
					return false; // pas un doc html
			}
			catch (Exception e)
			{
				System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de naviguer dans le fichier HTML purgé!");
                lir = null;
                it = null;
                return false;
			}
			return true;
		}

        private static bool DeleteMeta(ref System.Xml.XPath.XPathNavigator lir, ref System.Xml.XPath.XPathNodeIterator it)
		{
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
				return true;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de supprimer les infos méta");
				return false;
            }
		}

        private static bool ExtractCssStyles(string ParsedFileName, ref System.Xml.XPath.XPathNavigator lir, ref System.Xml.XPath.XPathNodeIterator it)
		{
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
				return true;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de traiter la section style");
				return false;
            }
		}

        private static bool CleanScripts(string ParsedFileName, ref System.Xml.XPath.XPathNavigator lir, ref System.Xml.XPath.XPathNodeIterator it)
        {
            try
            {
                ReportLog("Suppresion des sections <script> du fichier généré");
                //Suppression de la section style 
                it = lir.Select("/html/head//script");
                while (it.MoveNext())
                {
                    it.Current.DeleteSelf();
                    it = lir.Select("/html/head//script");
                }
                return true;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de traiter la section script");
                return false;
            }
        }

        private static bool AddStyleHeader(string TitreDuDoc, ref System.Xml.XPath.XPathNavigator lir, ref System.Xml.XPath.XPathNodeIterator it)
		{
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
				return true;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de modifier le titre et la mention au fichire css");
				return false;
            }
		}

        private static bool PurgeStylesFromHTML(ref System.Xml.XPath.XPathNavigator lir, ref System.Xml.XPath.XPathNodeIterator it)
		{
			 try
            {
                //suppression des balises contenant l'attribut style
                ReportLog("Purge des attributs style dans toutes les balises du body");
                it = lir.Select("/html/body//*[@style]");
                while (it.MoveNext())
                {
                    if (it.Current.Name.Equals("hr"))
                    { // Conservation des sauts de page
                        if (it.Current.OuterXml.ToLower().Contains("page-break-before"))
                        {
                            it.Current.ReplaceSelf("<hr class=\"sautPage\" />"); 
                            //TODO : ajouter sautPage dans le css
                        }
                        else
                        {
                            it.Current.ReplaceSelf("<hr />");
                        }
                    }
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
				
				return true;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de supprimer certaines infos de style");
				return false;
            }
		}

        private static bool RemoveLangFromBody(ref System.Xml.XPath.XPathNavigator lir, ref System.Xml.XPath.XPathNodeIterator it)
		{
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
				return true;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de supprimer les div en trop...");
				return false;
            }
		}

        private static bool RemoveAttribute(string nom, ref System.Xml.XPath.XPathNavigator lir, ref System.Xml.XPath.XPathNodeIterator it)
		{
            DateTime LastUpdate = DateTime.Now;
			try
            {
                int nbmax = 100;
                //suppression des attributs MsoNormal et MsoNoSpacing qui n'ont pas de raison d'être puisque ce sont les styles normaux
                ReportLog("Clean de l'attribut : " + nom);
                string exPath = "/html/body//*[@class='" + nom + "']"; 
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
				return true;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de supprimer "+ nom );
				return false;
            }
		}

        private static bool RemoveAttributeStartWith(string nom, ref System.Xml.XPath.XPathNavigator lir, ref System.Xml.XPath.XPathNodeIterator it)
        {
            DateTime LastUpdate = DateTime.Now;
            try
            {
                int nbmax = 100;
                //suppression des attributs MsoNormal et MsoNoSpacing qui n'ont pas de raison d'être puisque ce sont les styles normaux
                ReportLog("Clean de l'attribut : " + nom);
                string exPath = "/html/body//*[starts-with(@class,'" + nom + "')]"; 
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
                return true;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de supprimer " + nom);
                return false;
            }
        }

        private static bool ReplaceQuotes(ref System.Xml.XPath.XPathNavigator lir, ref System.Xml.XPath.XPathNodeIterator it)
		{
			try
            {
                ReportLog("Remplacement des <quotes> par des <cites> (HTML5)");
                //Remplacement des MsoQuote par des cite
                string exPath = "/html/body//*[@class='MsoQuote']";
                it = lir.Select(lir.Compile(exPath));
                while (it.MoveNext())
                {
                    if (it.Current.InnerXml.Length > 0)
                        if (it.Current.Name.Equals("p"))
                            it.Current.ReplaceSelf("<p><cite>" + it.Current.InnerXml + "</cite></p>");
                        else
                            it.Current.ReplaceSelf("<cite>" + it.Current.InnerXml + "</cite>");
                    else
                        it.Current.DeleteSelf();
                    it = lir.Select(lir.Compile(exPath));
                }
				return true;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de remplacer les quote et les <cite>");
				return false;
            }
		}

        private static bool RemoveBalise(string LocalName, ref System.Xml.XPath.XPathNavigator lir, ref System.Xml.XPath.XPathNodeIterator it)
		{
            try
            {
                ReportLog("Suppression des <" + LocalName + "> vides");
                //suppression des span autre que ceux de class (qui ne devraient pas exister)
                string exPath = "/html/body//"+ LocalName + "[not(@class)]";
                it = lir.Select(lir.Compile(exPath));
                while (it.MoveNext())
                {
                    if (it.Current.InnerXml.Length > 0)
                        it.Current.ReplaceSelf(it.Current.InnerXml);
                    else
                        it.Current.DeleteSelf();
                    it = lir.Select(lir.Compile(exPath));
                }
				return true;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de purger les " + LocalName);
				return false;
            }
		}

        private static bool RemoveBaliseClass(string LocalName, string className, ref System.Xml.XPath.XPathNavigator lir, ref System.Xml.XPath.XPathNodeIterator it)
        {
            try
            {
                ReportLog("Suppression des <" + LocalName + " class=\"" + className + "\" >");
                //suppression des span avec une class que l'on ne veux pas garder (par exemple les commentaires de relecture)
                string exPath = "/html/body//" + LocalName + "[@class=\"" + className +"\"]";
                it = lir.Select(lir.Compile(exPath));
                while (it.MoveNext())
                {
                    it.Current.DeleteSelf();
                    it = lir.Select(lir.Compile(exPath));
                }
                return true;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de purger les <" + LocalName + " class=\"" + className + "\" >");
                return false;
            }
        }

        private static bool RemoveCommentsFinalBlocs(ref System.Xml.XPath.XPathNavigator lir, ref System.Xml.XPath.XPathNodeIterator it)
        {
            //<div><hr class="msocomoff"
            try
            {
                ReportLog("Suppression du dernier bloc <div><hr class=\"msocomoff\">");
                string exPath = "/html/body/div/hr[@class=\"msocomoff\"]";
                it = lir.Select(lir.Compile(exPath));
                while (it.MoveNext())
                {
                    it.Current.MoveToParent();
                    it.Current.DeleteSelf();
                    it = lir.Select(lir.Compile(exPath));
                }
                return true;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de purger le dernier bloc de commentaires");
                return false;
            }

        }

        private static bool RemoveDiv(ref System.Xml.XPath.XPathNavigator lir, ref System.Xml.XPath.XPathNodeIterator it)
        {
            try
            {
                ReportLog("Suppression des <div> en trop");
                string exPath = "/html/body//div[not(@class)]";
                it = lir.Select(lir.Compile(exPath));
                while (it.MoveNext())
                {
                    //<div style="border:none;border-bottom:solid windowtext 1.0pt;padding:0cm 0cm 1.0pt 0cm">
                    bool hr = false;
                    string stl = it.Current.GetAttribute("style", "");
                    if (!string.IsNullOrEmpty(stl))
                        if (stl.Contains("border-bottom:solid"))
                        {
                            hr = true;
                        }

                    int child = it.Current.SelectChildren(XPathNodeType.Element).Count;
                    switch (child)
                    {
                        case 0:
                            if (!hr)
                                it.Current.DeleteSelf();
                            else
                                it.Current.ReplaceSelf("<hr />");
                            break;
                        case 1:
                            if (!hr)
                                it.Current.ReplaceSelf(it.Current.InnerXml);
                            else
                                it.Current.ReplaceSelf(it.Current.InnerXml + "<hr />");
                            break;
                        default:
                            continue;
                    }
                        
                    it = lir.Select(lir.Compile(exPath));
                }
                return true;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de purger les <div>");
                return false;
            }
        }

        private static bool RemoveDivClass(ref System.Xml.XPath.XPathNavigator lir, ref System.Xml.XPath.XPathNodeIterator it, string ClassRef)
        {
            try
            {
                ReportLog("Suppression des <div> en trop");
                string exPath = "/html/body//div[starts-with(@class,'" + ClassRef + "')]";
                it = lir.Select(lir.Compile(exPath));
                while (it.MoveNext())
                {
                    it.Current.ReplaceSelf(it.Current.InnerXml);
                    it = lir.Select(lir.Compile(exPath));
                }
                return true;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de purger les <div>");
                return false;
            }
        }
	}
}