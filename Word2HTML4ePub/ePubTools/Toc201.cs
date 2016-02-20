using System.Collections.Generic;
using System.Xml;
using System.IO;

namespace Word2HTML4ePub
{
    public partial class ePubTools
    {



        private static XmlDocument creerUneTocV201(string NavFilev3)
        {
            //Charger le navdoc v3
            XmlDocument navdoc3 = new XmlDocument();
            navdoc3.Load(NavFilev3);

            //Chercher les infos utiles
            XmlNode htmlnode = navdoc3.DocumentElement;
            while (htmlnode != null)
            {
                if (htmlnode.Name.Equals("html"))
                    break;
                htmlnode = htmlnode.NextSibling;
            }
            XmlNode bodynode = htmlnode.FirstChild;
            while (bodynode != null)
            {
                if (bodynode.Name.Contains("body"))
                    break;
                bodynode = bodynode.NextSibling;
            }

            XmlNode navnode = bodynode.FirstChild;
            while (navnode != null)
            {
                if (navnode.Name.Contains("nav"))
                {
                    if (!string.IsNullOrEmpty(((XmlElement)navnode).GetAttribute("epub:type")))
                        if (((XmlElement)navnode).GetAttribute("epub:type").Contains("toc"))
                            break;
                }
                navnode = navnode.NextSibling;
            }
            string titre = navnode.FirstChild.InnerText;
            XmlNode olnode = navnode.FirstChild;
            while (olnode != null)
            {
                if ((olnode.Name.Contains("ol")) || (olnode.Name.Contains("ul")))
                {
                    break;
                }
                olnode = olnode.NextSibling;
            }

            //Recup du UUID
            string uuid = null;
            XmlDocument contentdoc = new XmlDocument();
            contentdoc.Load(Path.Combine(Path.GetDirectoryName(NavFilev3), "content.opf"));
            System.Xml.XPath.XPathNavigator lir = contentdoc.CreateNavigator();
            System.Xml.XPath.XPathNodeIterator it = null;
            string exPath = "//*[local-name() = 'identifier']";
            it = lir.Select(lir.Compile(exPath));
            if (it.Count == 1)
            {
                it.MoveNext();
                uuid = it.Current.InnerXml;
            }

            //Créer le navdoc v2
            XmlDocument navdoc2 = new XmlDocument();
            {
                XmlElement root2 = navdoc2.CreateElement("toc", "ncx", "http://www.daisy.org/z3986/2005/ncx/");
                root2.SetAttribute("version", "2005-1");
                root2.SetAttribute("xml:lang", "FR");

                {

                    XmlElement headNode2 = navdoc2.CreateElement("toc", "head", root2.NamespaceURI);
                    XmlElement MetaElem = navdoc2.CreateElement("toc", "meta", root2.NamespaceURI);
                    MetaElem.SetAttribute("name", "dtb:uid");
                    if (string.IsNullOrEmpty(uuid))
                        MetaElem.SetAttribute("content", générerUuid());
                    else
                        MetaElem.SetAttribute("content", uuid);
                    headNode2.AppendChild(MetaElem);
                    root2.AppendChild(headNode2);
                }

                {
                    XmlElement titleNode2 = navdoc2.CreateElement("toc", "docTitle", root2.NamespaceURI);
                    XmlElement titletextNode2 = navdoc2.CreateElement("toc", "text", root2.NamespaceURI);
                    titletextNode2.InnerText = titre;
                    titleNode2.AppendChild(titletextNode2);
                    root2.AppendChild(titleNode2);
                }

                int playorder = 1;
                XmlElement navMap = navdoc2.CreateElement("toc", "navMap", root2.NamespaceURI);
                List<XmlNode> list = CreeNavv2(olnode, navdoc2, root2.NamespaceURI, ref playorder);
                if (list != null)
                { 
                    foreach (XmlNode node in list)
                        navMap.AppendChild(node);
                }

                root2.AppendChild(navMap);

                navdoc2.AppendChild(root2);

                XmlDeclaration decl = navdoc2.CreateXmlDeclaration("1.0", "UTF-8", "");
                navdoc2.InsertBefore(decl, root2);
            }
            return navdoc2;
        }

        private static List<XmlNode> CreeNavv2(XmlNode navListv3, XmlDocument doc, string NamespaceUri, ref int playorder)
        {
            List<XmlNode> retour = new List<XmlNode>();
            if (navListv3 == null)
                return null;

            foreach (XmlNode item in navListv3.ChildNodes)
            {
                if (!item.Name.Contains("li"))
                    continue; // pas un item -> pas de création.

                //Création du point
                XmlElement NavPoint = doc.CreateElement("toc", "navPoint", NamespaceUri);
                NavPoint.SetAttribute("id", générerUuid());

                XmlNode linknode = item.FirstChild;
                if (linknode.Name.Equals("a"))
                {

                    NavPoint.SetAttribute("playOrder", (playorder++).ToString());

                    XmlElement navLabel = doc.CreateElement("toc", "navLabel", NamespaceUri);
                    XmlElement text = doc.CreateElement("toc", "text", NamespaceUri);
                    text.InnerText = linknode.InnerText;
                    navLabel.AppendChild(text);
                    NavPoint.AppendChild(navLabel);
                    XmlElement content = doc.CreateElement("toc", "content", NamespaceUri);
                    content.SetAttribute("src", ((XmlElement)linknode).GetAttribute("href"));
                    NavPoint.AppendChild(content);
                }
                else if (linknode.Name.Equals("span"))
                {
                    XmlElement navLabel = doc.CreateElement("toc", "navLabel", NamespaceUri);
                    XmlElement text = doc.CreateElement("toc", "text", NamespaceUri);
                    text.InnerText = linknode.InnerText;
                    navLabel.AppendChild(text);
                    NavPoint.AppendChild(navLabel);

                    //Find next href content to apply
                    int indexd = item.InnerXml.IndexOf("href=") + 6;
                    int indexf = -1;
                    if (indexd > 6)
                    {
                        indexf = item.InnerXml.IndexOf("\"", indexd);
                        XmlElement content = doc.CreateElement("toc", "content", NamespaceUri);
                        content.SetAttribute("src", item.InnerXml.Substring(indexd, indexf - indexd));
                        NavPoint.AppendChild(content);
                    }

                }

                while (linknode != null)
                {
                    if ((linknode.Name.Contains("ol")) || (linknode.Name.Contains("ul"))) //il existe une sous liste -> Créons une sous liste
                    {
                        List<XmlNode> retoursub = CreeNavv2(linknode, doc, NamespaceUri, ref playorder);
                        foreach (XmlNode node in retoursub)
                            NavPoint.AppendChild(node);
                    }
                    linknode = linknode.NextSibling;
                }

                retour.Add(NavPoint);
            }
            return retour;
        }
    }
}