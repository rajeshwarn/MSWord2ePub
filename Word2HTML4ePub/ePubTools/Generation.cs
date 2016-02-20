using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Ionic.Zip;

namespace Word2HTML4ePub
{
    public partial class ePubTools
    {
        public static string générerUuid()
        {
            return Guid.NewGuid().ToString().ToUpper();
        }

        /// <summary>
        /// Fonction de génération d'un ePub
        /// </summary>
        /// <param name="filelist">La liste des fichiers ou des packages</param>
        /// <param name="ePubFile">Le path vers le fichier généré</param>
        /// <param name="ErrorLog">Un journal d'erreur</param>
        /// <returns>True = génération réussie</returns>
        public static bool CreateePub(string[] filelist, out string ePubFile, out string ErrorLog)
        {
            string tempfile = null;
            ePubFile = null;
            ErrorLog = "";

            //Etablir une liste de fichiers
            List<string> ListOfFiles = new List<string>();
            List<string> ListOfDir = new List<string>();
            List<string> ListOfDirRemoved = new List<string>();

            foreach (string f in filelist)
            {
                if (Directory.Exists(f))
                { // ajouter le contenu du dossier
                    ListOfFiles.AddRange(Directory.GetFiles(f, "*", SearchOption.AllDirectories));
                    
                    ListOfDir.Add(f);
                    ListOfDir.AddRange(Directory.GetDirectories(f));

                }
                else if (File.Exists(f))
                { // ajouter le fichier
                    ListOfFiles.Add(f);
                }
            }

            //Verif que les répertoires ne contiennent pas de cachés ou de systèmes, etc.
            for (int i = ListOfDir.Count-1; i > 0; i--)
            {
                FileAttributes fa = File.GetAttributes(ListOfDir[i]);
                if (((fa & FileAttributes.Hidden) == FileAttributes.Hidden) ||
                    ((fa & FileAttributes.System) == FileAttributes.System) ||
                    ((fa & FileAttributes.Temporary) == FileAttributes.Temporary))
                {
                    ListOfDir.RemoveAt(i);
                    continue;
                }
                if (Path.GetFileName(ListOfDir[i]).ToLower().Equals("temp"))
                {
                    ListOfDirRemoved.Add(ListOfDir[i]);
                    ListOfDir.RemoveAt(i);
                }
            }

            //Verif que la liste de fichiers ne contient pas de cachés ou de systèmes, etc.
            for (int i =ListOfFiles.Count-1; i>0 ; i--)
            {
                FileAttributes fa = File.GetAttributes(ListOfFiles[i]);
                if (((fa & FileAttributes.Hidden) == FileAttributes.Hidden) ||
                    ((fa & FileAttributes.System) == FileAttributes.System) ||
                    ((fa & FileAttributes.Temporary) == FileAttributes.Temporary))
                {
                    ListOfFiles.RemoveAt(i);
                    continue;
                }
                if ((from string s in ListOfDirRemoved where ListOfFiles[i].Contains(s) select s).Any())
                {
                    ListOfFiles.RemoveAt(i);
                    continue;
                }
            }

            //Creation d'un zip
            ZipFile zip = new ZipFile(Encoding.UTF8);
            zip.UseZip64WhenSaving = Zip64Option.Never;
            zip.AddDirectoryWillTraverseReparsePoints = false;
            zip.AlternateEncodingUsage = ZipOption.Never;
            zip.CompressionLevel = Ionic.Zlib.CompressionLevel.Level6;
            zip.CompressionMethod = CompressionMethod.Deflate;
            zip.EmitTimesInWindowsFormatWhenSaving = false;
            zip.EmitTimesInUnixFormatWhenSaving = false;
            zip.Encryption = EncryptionAlgorithm.None;
            zip.FullScan = false;
            zip.Strategy = Ionic.Zlib.CompressionStrategy.Default;
            //zip.UseUnicodeAsNecessary = true; //En a-t-on besoin dans ePubTool?
            zip.UseZip64WhenSaving = Zip64Option.Default;

            //Chercher le fichier mimetype
            string mime = SearchFileByName(ListOfFiles, "mimetype");
            if (string.IsNullOrEmpty(mime))
            {
                ErrorLog += "mimetype file not found\r\n";
                return false;
            }
            //L'ajouter et le retirer de la liste
            zip.AddFile(mime, "");
            ListOfFiles.Remove(mime);
            string basefolder = Path.GetDirectoryName(mime);

            //Chercher le fichier container.xml et vérifier qu'il est dans un dossier META-INF
            string container = SearchFileByName(ListOfFiles, "container.xml");
            if (string.IsNullOrEmpty(container))
            {
                ErrorLog += "container.xml not founded\r\n";
                return false;
            }
            if (!container.Contains("META-INF"))
            {
                ErrorLog += "container.xml not in META-INF\r\n";
                return false;
            }
            zip.AddFile(container, (container.Replace(Path.GetFileName(container), "").Replace(basefolder, "")));
            ListOfFiles.Remove(container);

            //Rechercher le fichier opf, l'ajouter et le retirer de la liste.
            string opffile = SearchOPFDoc(container); // aller chercher le fichier opf dans container.xml
            opffile = opffile.Replace('/', '\\');
            opffile = SearchFileByPartOfThePath(ListOfFiles, opffile);

            if (string.IsNullOrEmpty(opffile))
            {
                ErrorLog += "Impossible de trouver le fichier de contenu (\".opf\").\r\n";
                return false;
            }
            ListOfFiles.Remove(opffile);
            XmlDocument doc = new XmlDocument();
            doc.Load(opffile); // Chargement du content doc

            //Modif de la date de modification (versionning du fichier)
            MajDateModifiers(ref doc);

            if (TocNcx)
            {
                //Recherche de la page de navigation
                string navdoc = SearchNavDoc(doc);
                if (string.IsNullOrEmpty(navdoc))
                {
                    ErrorLog += "Erreur TOC : Impossible de trouver le document de navigation\r\n Vérifiez qu'il existe et est bien présent dans le content document.\r\n"; 
                    return false;
                }
                navdoc = SearchFileByName(ListOfFiles, navdoc);
                if (string.IsNullOrEmpty(navdoc))
                {
                    ErrorLog += "Erreur TOC : Impossible de trouver le document de navigation\r\n Vérifiez qu'il existe et est bien présent dans le content document.\r\n"; 
                    return false;
                }
                if (!File.Exists(navdoc))
                {
                    ErrorLog += "Erreur TOC : Impossible de trouver le document de navigation\r\n Vérifiez qu'il existe et est bien présent dans le content document.\r\n";
                    return false;
                }

                //Creation de la table v2.0.1
                XmlDocument tocv2 = creerUneTocV201(navdoc);
                //ajouter au zip
                tempfile = Path.GetTempFileName();
                tocv2.Save(tempfile);
                string packagefilename = navdoc.Replace(Path.GetFileName(navdoc), "toc.ncx");
                zip.AddEntry(packagefilename.Replace(basefolder, ""), File.ReadAllBytes(tempfile));
                File.Delete(tempfile);
                string oldtoc = SearchFileByName(ListOfFiles, "toc.ncx");
                if (!string.IsNullOrEmpty(oldtoc))
                    ListOfFiles.Remove(oldtoc);

                //Update content.doc
                XmlElement nodePackage = doc.DocumentElement;
                XmlNode nodemanifest = nodePackage.FirstChild;
                while ((nodemanifest != null) && (!nodemanifest.Name.Contains("manifest")))
                {
                    nodemanifest = nodemanifest.NextSibling;
                }

                XmlElement newElem = doc.CreateElement("opf", "item", nodePackage.NamespaceURI);

                newElem.SetAttribute("id", "autoToc");
                newElem.SetAttribute("href", "toc.ncx");
                newElem.SetAttribute("media-type", "application/x-dtbncx+xml");
                nodemanifest.AppendChild(newElem);

                XmlNode nodespine = nodemanifest.NextSibling;
                while ((nodemanifest != null) && (!nodemanifest.Name.Contains("spine")))
                {
                    nodemanifest = nodemanifest.NextSibling;
                }

                ((XmlElement)nodespine).SetAttribute("toc", "autoToc");
            }

            if (AutoRessources)
                doc = UpdateRessources(ListOfFiles, doc); //Mise à jour du fichier .opf

            /*
            //En Passant par un Fichier temporaire
            tempfile = Path.GetTempFileName();
            doc.Save(tempfile);
            zip.AddEntry(opffile.Replace(basefolder, ""), File.ReadAllBytes(tempfile));
            File.Delete(tempfile);
            */

            //En passant par une écriture en mémoire
            MemoryStream ms = new MemoryStream();
            doc.Save(ms);
            ms.Capacity = (int)ms.Length;
            zip.AddEntry(opffile.Replace(basefolder, ""), ms.GetBuffer()); 

            //Ajouter la liste des fichiers
            string fichiersIgnores = "";
            foreach (string f in ListOfFiles)
            {
                if (!f.Contains(basefolder))
                {
                    fichiersIgnores += f + "\r\n";
                    ListOfFiles.Remove(f);
                    continue;
                }
                if ((File.GetAttributes(f) & FileAttributes.Hidden) == FileAttributes.Hidden)
                {
                    fichiersIgnores += f + "\r\n";
                    ListOfFiles.Remove(f);
                    continue;
                }
                string FolderinZip = f.Replace(Path.GetFileName(f), "").Replace(basefolder, "");
                zip.AddFile(f, FolderinZip);
            }

            ePubFile = basefolder + ".epub";
            zip.Save(ePubFile);

            if (!string.IsNullOrEmpty(fichiersIgnores))
            {
                ErrorLog += "Des fichiers n'ont pas été ajoutés à l'epub : \r\n";
                ErrorLog += fichiersIgnores;
                return false;
            }
            else
            {
                return true;
            }
        }

        private static XmlDocument UpdateRessources(List<string> lof, XmlDocument doc)
        {
            string contentFileName = new Uri(doc.BaseURI).LocalPath;

            XmlElement nodePackage = doc.DocumentElement;
            XmlNode nodemanifest = nodePackage.FirstChild;
            while ((nodemanifest != null) && (!nodemanifest.Name.Contains("manifest")))
            {
                nodemanifest = nodemanifest.NextSibling;
            }

            int indexres = (from XmlElement xml in nodemanifest where xml.Attributes["id"].Value.Contains("autores") select xml).Count();
            string basefolder = contentFileName.Replace(Path.GetFileName(contentFileName), "");
            foreach (string file in lof)
            {
                string fn = Path.GetFileName(file);
                if ((from XmlElement xml in nodemanifest where xml.Attributes["href"].Value.Contains(fn) select xml).Any())
                    continue;

                string ext = Path.GetExtension(file).ToLower();

                XmlElement newElem = doc.CreateElement("opf", "item", nodePackage.NamespaceURI);

                newElem.SetAttribute("id", "autores" + indexres);
                indexres++;
                string src = file.Replace(basefolder, "");
                newElem.SetAttribute("href", src.Replace("\\", "/"));

                if (ext.Contains("html"))
                    newElem.SetAttribute("media-type", "application/xhtml+xml");
                else if (ext.Contains("xml"))
                    newElem.SetAttribute("media-type", "application/xhtml+xml");
                else if (ext.Contains("css"))
                    newElem.SetAttribute("media-type", "text/css");
                else if (ext.Contains("png"))
                    newElem.SetAttribute("media-type", "image/png");
                else if (ext.Contains("jpg"))
                    newElem.SetAttribute("media-type", "image/jpeg");
                else if (ext.Contains("jpeg"))
                    newElem.SetAttribute("media-type", "image/jpeg");
                else if (ext.Contains("svg"))
                    newElem.SetAttribute("media-type", "image/svg+xml");
                else if (ext.Contains("gif"))
                    newElem.SetAttribute("media-type", "image/gif");
                else
                    newElem.SetAttribute("media-type", "/octet-stream");
                nodemanifest.AppendChild(newElem);
            }

            XmlNode nodespine = nodePackage.FirstChild;
            while ((nodespine != null) && (!nodespine.Name.Contains("spine")))
            {
                nodespine = nodespine.NextSibling;
            }

            return doc;
        }

        /// <summary>
        /// Mise à jour de la date de génération
        /// </summary>
        /// <param name="doc"></param>
        private static void MajDateModifiers(ref XmlDocument doc)
        {
            XmlElement nodePackage = doc.DocumentElement;

            XmlNode nodemetadata = nodePackage.FirstChild;
            while ((nodemetadata != null) && (!nodemetadata.Name.Contains("metadata")))
            {
                nodemetadata = nodemetadata.NextSibling;
            }

            if (nodemetadata == null)
                return;

            XmlNode nodemeta = nodemetadata.FirstChild;
            while ((nodemeta != null) && (!nodemeta.Name.Contains("meta")))
            {
                nodemeta = nodemeta.NextSibling;
            }
            if (nodemeta == null)
                return;

            nodemeta.InnerText = DateTime.Now.ToString("s") + "Z";
            return;
        }
    }
}