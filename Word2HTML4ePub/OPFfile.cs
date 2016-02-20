using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace Word2HTML4ePub
{
    /// <summary>
    /// Cette classe sert à accéder aux informations contenues dans les ".opf"
    /// </summary>
    internal partial class OPFFile
    {
        /// <summary>
        /// Relecture d'un Package
        /// </summary>
        /// <param name="PackagePath">Path vers un Package</param>
        /// <returns></returns>
        internal static List<OPFData> ReadPackage(string PackagePath)
        {
            List<OPFData> retour = new List<OPFData>();
            string opfPath = Path.Combine(PackagePath, "content");
            if (!Directory.Exists(opfPath))
                return null;
            string[] lof = Directory.GetFiles(opfPath, "*.opf", SearchOption.TopDirectoryOnly);
            if ((lof.Length == 0) | (lof.Length > 1))
                return null;

            XmlDocument doc = new XmlDocument();
            doc.Load(lof[0]);

            XmlNode packNode = doc.DocumentElement;
            if (packNode == null)
                throw new Exception("Ce fichier n'est pas compatible!");

            if (!packNode.LocalName.Equals("package"))
                throw new Exception("Ce fichier n'est pas compatible!");

            //Pour le traitement du metadata
            XmlNode nodeMetaData = packNode.FirstChild;
            while ((nodeMetaData != null) && (!nodeMetaData.Name.Contains("metadata")))
            {
                nodeMetaData = nodeMetaData.NextSibling;
            }

            XmlNode nodeDC = nodeMetaData.FirstChild;
            while (nodeDC != null)
            {
                retour.Add(new OPFData(nodeDC.Name, nodeDC.InnerText));
                nodeDC = nodeDC.NextSibling;
            }

            return retour;
        }

        internal static bool UpdatePackage(Microsoft.Office.Interop.Word.Document doc)
        {
            //Chargement du paramètre package et contrôle
            //string PackagePath = WordHTML2ePubHTML.GetDocProperty(doc, "PackagePath");
            //if (string.IsNullOrEmpty(PackagePath))
            //    return false;
            string PackagePath = Word2ePub.GetCurrentDocPackageFolder();

            string opfPath = Path.Combine(PackagePath, "content");
            if (!Directory.Exists(opfPath))
                return false;
            string[] lof = Directory.GetFiles(opfPath, "*.opf", SearchOption.TopDirectoryOnly);
            if ((lof.Length == 0) | (lof.Length > 1))
                return false;

            //Ouverture du doc
            XmlDocument docxml = new XmlDocument();
            docxml.Load(lof[0]);

            XmlNode packNode = docxml.DocumentElement;
            if (packNode == null)
                throw new Exception("Ce fichier n'est pas compatible!");

            if (!packNode.LocalName.Equals("package"))
                throw new Exception("Ce fichier n'est pas compatible!");

            //Pour le traitement du metadata
            XmlNode nodeMetaData = packNode.FirstChild;
            while ((nodeMetaData != null) && (!nodeMetaData.Name.Contains("metadata")))
            {
                nodeMetaData = nodeMetaData.NextSibling;
            }

            //Maj des balises obligatoires
            UpdateOrCreateNode(docxml, ref nodeMetaData, "dc:identifier", "urn:uuid:" + WordHTML2ePubHTML.GetDocProperty(doc, "GUID"));
            UpdateOrCreateNode(docxml, ref nodeMetaData, "dc:title", WordHTML2ePubHTML.GetDocProperty(doc, "Titre"));

            string lang = WordHTML2ePubHTML.GetDocProperty(doc, "Language");
            if (string.Equals(lang.ToLower(), "Francais"))
                lang = "fr";
            else if (string.Equals(lang.ToLower(), "English"))
                lang = "en";
            else if (string.Equals(lang.ToLower(), "German"))
                lang = "de";
            else if (string.Equals(lang.ToLower(), "Spanich"))
                lang = "sp";
            UpdateOrCreateNode(docxml, ref nodeMetaData, "dc:language", lang);
            
            //Maj des balises facultatives
            UpdateOrCreateNode(docxml, ref nodeMetaData, "dc:creator", WordHTML2ePubHTML.GetDocProperty(doc, "Auteur"));
            UpdateOrCreateNode(docxml, ref nodeMetaData, "dc:publisher", WordHTML2ePubHTML.GetDocProperty(doc, "Editeur"));
            UpdateOrCreateNode(docxml, ref nodeMetaData, "dc:description", WordHTML2ePubHTML.GetDocProperty(doc, "Description"));
            UpdateOrCreateNode(docxml, ref nodeMetaData, "dc:subject", WordHTML2ePubHTML.GetDocProperty(doc, "Sujet"));
            UpdateOrCreateNode(docxml, ref nodeMetaData, "dc:date", WordHTML2ePubHTML.GetDocDateTime(doc, "DateCreation").ToString("s") + "Z");

            /*
            cmdCouv.Tag = WordHTML2ePubHTML.GetDocProperty(doc, "CoverPath");
            string licence = WordHTML2ePubHTML.GetDocProperty(doc, "Licence");
            */
            docxml.Save(new Uri(docxml.BaseURI).LocalPath);
            return true;
        }

        private static bool UpdateOrCreateNode(XmlDocument doc, ref XmlNode ParentNode, string id, string valeur)
        {
            if (string.IsNullOrEmpty(id))
                return false;
            XmlNode child = ParentNode.FirstChild;
            while (child != null)
            {
                if (child.Name.Equals(id))
                {
                    child.InnerText = valeur;
                    return true;
                }
                child = child.NextSibling;
            }
            
            XmlElement newElem = doc.CreateElement(id, ParentNode.FirstChild.NamespaceURI);
            newElem.InnerText = valeur;
            ParentNode.AppendChild(newElem);
            return true;
        }

        /// <summary>
        /// Fonction pour mettre à jour le package avec les fichiers auto-générés
        /// </summary>
        /// <param name="PackagePath">Chemin du package</param>
        /// <param name="FilesToAddToManifest">Liste des fichiers à ajouter</param>
        public static void UpdateRessources(string PackagePath, List<string> FilesToAddToManifest)
        {
            if (string.IsNullOrEmpty(PackagePath))
                return;

            //recherche du opf
            string opfFilePath = Directory.GetFiles(PackagePath, "*.opf", SearchOption.AllDirectories).FirstOrDefault();
            if (opfFilePath.Length == 0)
                return;

            //string contentFileName = new Uri(doc.BaseURI).LocalPath;

            XmlDocument doc = new XmlDocument();
            doc.Load(opfFilePath); // Chargement du content doc

            string contentFileName = new Uri(doc.BaseURI).LocalPath;

            XmlElement nodePackage = doc.DocumentElement;

            //Pour le traitement du manifest
            XmlNode nodemanifest = nodePackage.FirstChild;
            while ((nodemanifest != null) && (!nodemanifest.Name.Contains("manifest")))
            {
                nodemanifest = nodemanifest.NextSibling;
            }

            //pour le traitement de la spine
            XmlNode nodeSpine = nodePackage.FirstChild;
            while ((nodeSpine != null) && (!nodeSpine.Name.Contains("spine")))
            {
                nodeSpine = nodeSpine.NextSibling;
            }

            //Recup des fichiers qui ne devraient pas être là, ou à effacer...
            List<XmlNode> listOfFileToDelete = (from XmlNode xml in nodemanifest.ChildNodes where xml.Attributes["id"].Value.Contains("Chap") select xml).ToList();
            //listOfFileToDelete.AddRange((from XmlNode xml in nodemanifest.ChildNodes where xml.Attributes["id"].Value.Contains("autores") select xml).ToList());
            listOfFileToDelete.AddRange((from XmlNode xml in nodemanifest.ChildNodes where xml.Attributes["id"].Value.Contains("navPage") select xml).ToList());

            //Suppression des ressources de l'OPF et du package
            foreach (XmlNode xml in listOfFileToDelete)
            {
                string oldFile = Path.Combine(PackagePath, "content", xml.Attributes["href"].Value.Replace("/", "\\"));
                if (File.Exists(oldFile))
                    File.Delete(oldFile);
                nodemanifest.RemoveChild(xml);

                //Clean de la spine actuelle (on part de listOfFileToDelete)
                XmlNode node = (from XmlNode xmlspine in nodeSpine.ChildNodes where xmlspine.Attributes["idref"].Value.Equals(xml.Attributes["id"].Value) select xmlspine).FirstOrDefault();
                if (node != null)
                    nodeSpine.RemoveChild(node);
            }


            //Copie des fichiers dans le package
            List<string> FinalFiles = new List<string>(FilesToAddToManifest.Count);
            foreach (string s in FilesToAddToManifest)
            {
                string newfile = Path.Combine(PackagePath, "content", Path.GetFileName(s));
                if (File.Exists(newfile))
                    File.Delete(newfile);
                File.Move(s, newfile);
                FinalFiles.Add(newfile);
            }

            string basefolder = contentFileName.Replace(Path.GetFileName(contentFileName), "");
            //Ajout de la table de nav
            {
                //manifest
                XmlElement newElem = doc.CreateElement("opf", "item", nodePackage.NamespaceURI);
                newElem.SetAttribute("properties", "nav");
                newElem.SetAttribute("id", "navPage");
                string src = FinalFiles.Last().Replace(basefolder, "");
                newElem.SetAttribute("href", src.Replace("\\", "/"));
                newElem.SetAttribute("media-type", "application/xhtml+xml");
                nodemanifest.AppendChild(newElem);

                //spine
                newElem = doc.CreateElement("opf", "itemref", nodePackage.NamespaceURI);
                newElem.SetAttribute("idref", "navPage");
                newElem.SetAttribute("linear", "no"); // n'apparait pas dans l'ePub
                nodeSpine.AppendChild(newElem);
            }

            //Ajout des fichiers au Manifest
            for (int i = 0; i < FilesToAddToManifest.Count - 1; i++)
            {
                //manifest
                XmlElement newElem = doc.CreateElement("opf", "item", nodePackage.NamespaceURI);
                newElem.SetAttribute("id", "Chap" + i);
                string src = FinalFiles[i].Replace(basefolder, "");
                newElem.SetAttribute("href", src.Replace("\\", "/"));
                newElem.SetAttribute("media-type", "application/xhtml+xml");
                nodemanifest.AppendChild(newElem);

                //spine
                newElem = doc.CreateElement("opf", "itemref", nodePackage.NamespaceURI);
                newElem.SetAttribute("idref", "Chap" + i);
                newElem.SetAttribute("linear", "yes"); // apparait dans l'ePub
                nodeSpine.AppendChild(newElem);

            }

            doc.Save(new Uri(doc.BaseURI).LocalPath);

            return;
        }

        /// <summary>
        /// Fonction utilisée pour changer le fichier image dans le package
        /// </summary>
        /// <param name="PackagePath"></param>
        /// <param name="newcoverPath"></param>
        /// <returns></returns>
        public static bool ChangeCover(string PackagePath, string newcoverPath)
        {
            //Chargement du paramètre package et contrôle
            if (string.IsNullOrEmpty(PackagePath))
                return false;

            string opfPath = Path.Combine(PackagePath, "content");
            if (!Directory.Exists(opfPath))
                return false;
            string[] lof = Directory.GetFiles(opfPath, "*.opf", SearchOption.TopDirectoryOnly);
            if ((lof.Length == 0) | (lof.Length > 1))
                return false;

            //Ouverture du doc
            XmlDocument docxml = new XmlDocument();
            docxml.Load(lof[0]);

            XmlNode nodePackage = docxml.DocumentElement;
            if (nodePackage == null)
                throw new Exception("Ce fichier n'est pas compatible!");

            if (!nodePackage.LocalName.Equals("package"))
                throw new Exception("Ce fichier n'est pas compatible!");

            //Pour le traitement du manifest
            XmlNode nodemanifest = nodePackage.FirstChild;
            while ((nodemanifest != null) && (!nodemanifest.Name.Contains("manifest")))
            {
                nodemanifest = nodemanifest.NextSibling;
            }

            //Recherche d'une éventuelle balise cover
            //XmlNode covernode = (from XmlNode items in nodemanifest.ChildNodes where items.Attributes["properties"].Equals("cover-image") select items).FirstOrDefault();
            
            XmlNode covernode = nodemanifest.FirstChild;
            while (covernode != null)
            {
                XmlNode test = covernode.Attributes["properties"];
                if (test != null)
                {
                    if (covernode.Attributes["properties"].Value.Equals("cover-image"))
                        break;
                }
                covernode = covernode.NextSibling;
            }
            
            string oldfile = "";
            if (covernode == null)
            {
                //Création du noeud
                XmlElement newElem = docxml.CreateElement("opf", "item", nodePackage.NamespaceURI);
                newElem.SetAttribute("properties", "cover-image");
                newElem.SetAttribute("id", "cover");
                newElem.SetAttribute("href", "tempo");
                newElem.SetAttribute("media-type", "image/jpeg");
                covernode = nodemanifest.AppendChild(newElem);
            }
            else
                oldfile = Path.Combine(PackagePath, "content",covernode.Attributes["href"].Value);

            //Suppression du fichier existant
            if (!oldfile.ToLower().Equals(newcoverPath.ToLower()))
                if (File.Exists(oldfile))
                    File.Delete(oldfile);

            //Copie de la nouvelle image
            if (string.IsNullOrEmpty(newcoverPath))
            {
                //Suppression du noeud
                nodemanifest = nodemanifest.RemoveChild(covernode);
            }
            else
            {
                if (!newcoverPath.ToLower().Equals(Path.Combine(PackagePath, "content", Path.GetFileName(newcoverPath)).ToLower()))
                    File.Copy(newcoverPath, Path.Combine(PackagePath, "content", Path.GetFileName(newcoverPath)), true);

                //Modif du noeud
                covernode.Attributes["href"].Value = Path.GetFileName(newcoverPath);
                switch (Path.GetExtension(newcoverPath).ToLower())
                {
                    case ".jpg":
                    case ".jpeg":
                        covernode.Attributes["media-type"].Value = "image/jpeg";
                        break;
                    case ".gif":
                        covernode.Attributes["media-type"].Value = "image/gif";
                        break;
                    case ".png":
                        covernode.Attributes["media-type"].Value = "image/png";
                        break;
                    case ".svg":
                        covernode.Attributes["media-type"].Value = "image/svg+xml";
                        break;
                    default:
                        throw new Exception("les fichiers de type: " + Path.GetExtension(newcoverPath) + " ne sont pas autorisés dans la spec epub.");
                }
            }

            //sauvegarde
            docxml.Save(new Uri(docxml.BaseURI).LocalPath);
            return true;
        }


        /// <summary>
        /// Fonction pour récupérer le nom du fichier html servant de couverture
        /// </summary>
        /// <param name="PackagePath"></param>
        /// <returns></returns>
        public static string GetCoverFile(string PackagePath)
        {
            if (string.IsNullOrEmpty(PackagePath))
                return null;

            //recherche du opf
            if (!Directory.GetFiles(PackagePath, "*.opf", SearchOption.AllDirectories).Any())
                return null;

            string opfFilePath = Directory.GetFiles(PackagePath, "*.opf", SearchOption.AllDirectories).FirstOrDefault();
            if (opfFilePath.Length == 0)
                return null;

            XmlDocument doc = new XmlDocument();
            doc.Load(opfFilePath); // Chargement du content doc

            string contentFileName = new Uri(doc.BaseURI).LocalPath;

            XmlElement nodePackage = doc.DocumentElement;

            //Pour le traitement du manifest
            XmlNode nodemanifest = nodePackage.FirstChild;
            while ((nodemanifest != null) && (!nodemanifest.Name.Contains("manifest")))
            {
                nodemanifest = nodemanifest.NextSibling;
            }

            //pour le traitement de la spine
            XmlNode nodeSpine = nodePackage.FirstChild;
            while ((nodeSpine != null) && (!nodeSpine.Name.Contains("spine")))
            {
                nodeSpine = nodeSpine.NextSibling;
            }

            //Recupération de la première ligne de la spine (qui doit logiquement être la couverture!)
            string coverIdRef = nodeSpine.FirstChild.Attributes["idref"].Value;

            //Recherche de la référence associée
            string result = (from XmlNode node in nodemanifest.ChildNodes where node.Attributes["id"].Value.Equals(coverIdRef) select node.Attributes["href"].Value).FirstOrDefault();
            if (!string.IsNullOrEmpty(result))
                result = Path.Combine(PackagePath, "content", result);

            return result;
        }


        /// <summary>
        /// Classe de transfert de paramètres
        /// </summary>
        public class OPFData
        {
            public string id;
            public string valeur;

            public OPFData(string id, string valeur)
            {
                this.id = id;
                this.valeur = valeur;
            }
        }

    }


}
