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
        public static bool RemoveIndent(ref XPathNavigator lir)
        {
            ReportLog("Suppression de l'indentation word");
            try
            {
                RemoveIndent(ref lir, "/html/body//p");
                //Suppression dans les titres...
                for (int i = 1; i < 8; i++)
                    RemoveIndent(ref lir, "/html/body//h" + i.ToString());

                return true;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de supprimer les indents");
                return false;
            }
        }

        private static void RemoveIndent(ref XPathNavigator lir, string XPath)
        {

            System.Xml.XPath.XPathNodeIterator it = lir.Select(lir.Compile(XPath));
            while (it.MoveNext())
            {
                if (!it.Current.MoveToChild(XPathNodeType.Text))
                    continue;
                do
                {
                    StringBuilder sb = new StringBuilder();
                    string[] sentence = it.Current.Value.Split('\r');

                    foreach (string s in sentence)
                    {
                        string s1 = s;

                        if (string.IsNullOrEmpty(s1))
                            continue;

                        if (s1[0].Equals('\n'))
                        {
                            s1 = s1.Replace('\n', ' ');
                            s1 = s1.TrimStart(null);
                        }

                        sb.Append(s1 + " ");
                    }
                    it.Current.SetValue(sb.ToString());
                } while (it.Current.MoveToNext(XPathNodeType.Text));
                it.Current.MoveToParent();
            }
        }

        public static void UpdateCoverHTMLFile(string filename, string ImageFileName, string newTitre, string newSousTitre, string newAuteur)
        {
            //ouverture du fichier html
            XmlDocument doc = new XmlDocument();
            doc.Load(filename); // Chargement du content doc

            string contentFileName = new Uri(doc.BaseURI).LocalPath;

            XmlElement nodeHTML = doc.DocumentElement;

            //Recherche du header
            XmlNode nodeHead = (from XmlNode node in nodeHTML.ChildNodes where node.Name.ToLower().Equals("head") select node).FirstOrDefault();
            if (nodeHead == null)
                return;

            //Recherche du titre
            XmlNode nodeTitre = (from XmlNode node in nodeHead.ChildNodes where node.Name.ToLower().Equals("title") select node).FirstOrDefault();
            if (nodeTitre == null)
                return;
            nodeTitre.InnerText = newTitre;

            //Recherche du body
            XmlNode nodeBody = (from XmlNode node in nodeHTML.ChildNodes where node.Name.ToLower().Equals("body") select node).FirstOrDefault();
            if (nodeBody == null)
                return;

            //Recherche du svg
            XmlNode nodeSVG = (from XmlNode node in nodeBody.ChildNodes where node.Name.ToLower().Equals("svg") select node).FirstOrDefault();
            if (nodeSVG == null)
                return;

            //Maj de l'image
            XmlNode nodeImg = (from XmlNode node in nodeSVG.ChildNodes where node.Name.ToLower().Equals("image") select node).FirstOrDefault();
            //if (nodeImg != null)
            //    nodeSVG.RemoveChild(nodeImg);
            List<XmlNode> lon = (from XmlNode node in nodeSVG select node).ToList();
            foreach (XmlNode node in lon)
                nodeSVG.RemoveChild(node);

            int width = 150;
            int height = width * 4 / 3;

            if (!string.IsNullOrEmpty(ImageFileName))
            {
                System.Drawing.Image img = System.Drawing.Image.FromFile(ImageFileName);
                if (img.Width > width)
                { 
                    width = img.Width;
                    height = width * 4 / 3;
                }

                if (img.Height > (height - 60))
                {
                    height = img.Height + 60;
                    width = height * 3 / 4;
                }

                XNamespace xlink = "http://www.w3.org/1999/xlink";
                XmlElement newElem = doc.CreateElement("image", nodeSVG.NamespaceURI);
                newElem.SetAttribute("id", "image");
                newElem.SetAttribute("x", ((width - img.Width) / 2).ToString());
                newElem.SetAttribute("y", "0");
                newElem.SetAttribute("width", img.Width.ToString());
                newElem.SetAttribute("height", img.Height.ToString());
                XmlAttribute newatt = doc.CreateAttribute("xlink", "href", xlink.NamespaceName);
                newatt.Value = Path.GetFileName(ImageFileName);
                newElem.Attributes.Append(newatt);
                nodeImg = nodeSVG.InsertBefore(newElem, nodeSVG.FirstChild);
            }
            
            int botImg = height -60;

            //le titre
            {
                int TitreHeight = 16;
                XmlElement newElem = doc.CreateElement("text", nodeSVG.NamespaceURI);
                newElem.SetAttribute("id", "titre");
                newElem.SetAttribute("x", width.ToString());
                newElem.SetAttribute("y", (botImg+TitreHeight+2).ToString());
                newElem.SetAttribute("text-anchor", "end");
                newElem.SetAttribute("font-family", "arial");
                newElem.SetAttribute("font-size", TitreHeight.ToString());
                newElem.InnerText = newTitre;
                nodeSVG.AppendChild(newElem);
                botImg += (TitreHeight +2);
            }

            //sous titre si présent
            if (!string.IsNullOrEmpty(newSousTitre))
            {
                int sousTitreHeight = 14;
                XmlElement newElem = doc.CreateElement("text", nodeSVG.NamespaceURI);
                newElem.SetAttribute("id", "sousstitre");
                newElem.SetAttribute("x", width.ToString());
                newElem.SetAttribute("y", (botImg + sousTitreHeight + 2).ToString());
                newElem.SetAttribute("text-anchor", "end");
                newElem.SetAttribute("font-family", "arial");
                newElem.SetAttribute("font-size", sousTitreHeight.ToString());
                newElem.InnerText = newSousTitre;
                nodeSVG.AppendChild(newElem);
                botImg += (sousTitreHeight+2);
            }

            //Auteur
            if (!string.IsNullOrEmpty(newAuteur))
            {
                int auteurHeight = 14;
                XmlElement newElem = doc.CreateElement("text", nodeSVG.NamespaceURI);
                newElem.SetAttribute("id", "auteur");
                newElem.SetAttribute("x", width.ToString());
                newElem.SetAttribute("y", (botImg + auteurHeight + 2).ToString());
                newElem.SetAttribute("text-anchor", "end");
                newElem.SetAttribute("font-family", "arial");
                newElem.SetAttribute("font-size", auteurHeight.ToString());
                newElem.InnerText = newAuteur;
                nodeSVG.AppendChild(newElem);
                botImg += (auteurHeight+2);
            }

            //Maj du viewport du SVG
            nodeSVG.Attributes["viewBox"].Value="0 0 " + width.ToString() + " " + height.ToString();

            //sauvegarde
            doc.Save(new Uri(doc.BaseURI).LocalPath);

        }
    }
}