using Ionic.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace Word2HTML4ePub
{
    public class Open_ePub
    {
        XmlDocument document = null;
        XmlElement nodePackage = null;
        XmlElement nodeMetadata = null;
        XmlElement nodeManifest = null;
        XmlElement nodeSpine = null;
        List<ePubDocItem> lePubItem = null;
        List<ePubDocItem> lePubRess = null;

        string opfFilePath = null;

        string pFileName = null;
        /// <summary>
        /// FilePath to the ePub file. If null => not an ePubFile or no file provided...
        /// </summary>
        public string FileName
        {
            get { return pFileName; }
            set
            {
                if (string.IsNullOrEmpty(value))
                    pFileName = null;
                else
                {
                    if (!File.Exists(value))
                        pFileName = null;
                    else
                    {
                        using (MemoryStream ms = new MemoryStream(File.ReadAllBytes(value)))
                        {
                            ms.Seek(0, SeekOrigin.End);
                            if (!ZipFile.IsZipFile(ms, false))
                                pFileName = value;
                            else
                                pFileName = null;

                        }
                    }
                }
            }
        }

        string pOutputFileName = null;
        /// <summary>
        /// FilePath to the ePub file. If null => not an ePubFile or no file provided...
        /// </summary>
        public string OutputFileName
        {
            get { return pOutputFileName; }
            private set
            {
                pOutputFileName = value;
            }
        }

        public static string ConvertEpub2HTML(string FileName)
        {
            Open_ePub converter = new Open_ePub(FileName);
            return converter.OutputFileName;
        }

        public Open_ePub(string FileName)
        {
            this.FileName = FileName;

            IsEPubFile(); //TODO : test et décision
            document = ReadOPFFile();
            if (document == null)
                return;

            ReadOPFData(document);
            XmlDocument htmlFinalFile = RebuildHTMLFile(lePubItem);
            string fullHTMLPath = FileName.Substring(0, FileName.Length - Path.GetExtension(FileName).Length) + ".html";
            htmlFinalFile.Save(fullHTMLPath);
            OutputFileName = fullHTMLPath;

            string fullpath = Path.GetDirectoryName(fullHTMLPath);

            //extraction of the ressources files (only images?)
            foreach (ePubDocItem item in lePubRess)
            {
                if (item.mediaType.Contains("image"))
                    ExtractFileFromZip(item.href, fullpath);
                else if (item.mediaType.Contains("css"))
                    ExtractFileFromZip(item.href, fullpath);
            }
        }

        /// <summary>
        /// Check that the file contains a valid ePubFile
        /// </summary>
        /// <returns>True if the file is a wellFormed ePubFile</returns>
        bool IsEPubFile()
        {
            if (string.IsNullOrEmpty(FileName))
                return false;

            //1. Read zip file
            ZipFile ePubPack = ZipFile.Read(FileName);
            
            //2. retreive .opf file from the Package
            List<ZipEntry> loz = (List<ZipEntry>)(ePubPack.SelectEntries("*.opf"));
            if (loz.Count !=1)
                return false;

            //3. Check with epubCheck that the file is wellFomed
            //TODO

            opfFilePath = Path.GetDirectoryName(loz[0].FileName);

            return true;
        }

        /// <summary>
        /// Retreive a file zipped
        /// </summary>
        /// <param name="ZippedFileName">The name of the file to get</param>
        /// <returns>a MemoryStream containing the file</returns>
        MemoryStream GetFileFromZip(string ZippedFileNameWithoutDirectory)
        {
            //1. Read zip file
            ZipFile ePubPack = ZipFile.Read(FileName);

            //2. retreive .opf file from the Package
            List<ZipEntry> loz = (List<ZipEntry>)(ePubPack.SelectEntries(ZippedFileNameWithoutDirectory));
            if (loz.Count != 1)
                return null;

            //3. Read opf file
            MemoryStream ms = new MemoryStream();
            loz[0].Extract(ms);
            ms.Position = 0;
            return ms; 
        }

        /// <summary>
        /// Retreive a file zipped
        /// </summary>
        /// <param name="ZippedFileName">The name of the file to get</param>
        /// <returns>a MemoryStream containing the file</returns>
        MemoryStream GetFileFromNameZip(string ZippedFileName)
        {
            //1. Read zip file
            ZipFile ePubPack = ZipFile.Read(FileName);

            //2. retreive file from the Package
            ZipEntry zipEnt = null;
            foreach (ZipEntry z in ePubPack.Entries)
            {
                
                //if (!z.FileName.Contains(ZippedFileName))
                if (!z.FileName.Contains(Path.GetFileName(ZippedFileName)))
                    continue;
                zipEnt = z;
                break;
            }

            if (zipEnt == null)
                return null;
            //List<ZipEntry> loz = (List<ZipEntry>)(ePubPack.SelectEntries(Path.GetFileName(ZippedFileName),ZippedFileName.Substring(0, ZippedFileName.Length- Path.GetFileName(ZippedFileName).Length)));
            //if (loz.Count != 1)
            //    return null;

            //3. Read opf file
            MemoryStream ms = new MemoryStream();
            //loz[0].Extract(ms);
            zipEnt.Extract(ms);
            ms.Position = 0;
            return ms;
        }

        void ExtractFileFromZip(string ZippedFileName, string DestinationPath)
        {
            MemoryStream ms = GetFileFromNameZip(ZippedFileName);

            if (ms == null)
                return;

            string currentPath = Path.GetDirectoryName(Path.Combine(DestinationPath, ZippedFileName));
            if (!Directory.Exists(currentPath))
                Directory.CreateDirectory(currentPath);

            File.WriteAllBytes(Path.Combine(currentPath, Path.GetFileName(ZippedFileName)), ms.GetBuffer());
            //zipEnt.Extract(DestinationPath, ExtractExistingFileAction.OverwriteSilently);
            return;
        }

        /// <summary>
        /// Read the opfFile contained in the ePub File
        /// </summary>
        /// <returns>XmlDocument object</returns>
        XmlDocument ReadOPFFile()
        {
            if (string.IsNullOrEmpty(FileName))
                return null;
            XmlDocument doc = null;
            using (MemoryStream ms = GetFileFromZip("*.opf"))
            {
                doc = new XmlDocument();
                ms.Position = 0;
                doc.Load(ms);
            }
            return doc; 
        }

        bool ReadOPFData(XmlDocument doc)
        {
            nodePackage = doc.DocumentElement;
            if (nodePackage == null)
                throw new Exception("Fichier non compatible (pas de racine dans le xml)!");

            if (!nodePackage.LocalName.Equals("package"))
                throw new Exception("Fichier non compatible (la racine ne se nomme pas \"package\"!");

            try
            {
                //Recup des metadata
                nodeMetadata = (XmlElement)(nodePackage.FirstChild);
                if (!nodeMetadata.LocalName.Equals("metadata"))
                    throw new Exception("Fichier non compatible (pas de section \"metadata\"!");

                //Recup du manifest
                nodeManifest = (XmlElement)(nodeMetadata.NextSibling);
                if (!nodeManifest.LocalName.Equals("manifest"))
                    throw new Exception("Fichier non compatible (pas de section \"manifest\"!");

                //Recup de la spine
                nodeSpine = (XmlElement)(nodeManifest.NextSibling);
                if (!nodeSpine.LocalName.Equals("spine"))
                    throw new Exception("Fichier non compatible (pas de section \"spine\"!");
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de naviguer dans le fichier OPF!");
                return false;
            }

            lePubItem = new List<ePubDocItem>();
            XmlNode att = null;
            foreach (XmlNode spineChildNode in nodeSpine.ChildNodes)
            {
                if (!spineChildNode.LocalName.Equals("itemref"))
                    continue;

                ePubDocItem item = new ePubDocItem();
                att = null;
                att = spineChildNode.Attributes["idref"];
                if (att == null)
                    continue;
                item.idref = att.Value;

                att = spineChildNode.Attributes["linear"];
                if (att != null)
                    item.linear = att.Value;
                else
                    item.linear = "yes";

                //Get Manifest ress
                XmlNode ManifestChild = GetManifestRess(item.idref);
                item.href = ManifestChild.Attributes["href"].Value;
                item.mediaType = ManifestChild.Attributes["media-type"].Value;
                lePubItem.Add(item);
            }
            
            lePubRess = new List<ePubDocItem>();
            foreach (XmlNode manifestChildNode in nodeManifest.ChildNodes)
            {
                if (!manifestChildNode.LocalName.Equals("item"))
                    continue;

                ePubDocItem item = new ePubDocItem();
                att = null;
                att = manifestChildNode.Attributes["id"];
                if (att == null)
                    continue;
                item.idref = att.Value;

                item.href = manifestChildNode.Attributes["href"].Value;
                item.mediaType = manifestChildNode.Attributes["media-type"].Value;
                lePubRess.Add(item);
            }


            return true;
        }

        /// <summary>
        /// Search in Manifest Node the specific ID from the spine
        /// </summary>
        /// <param name="IDRef">ID from the spine</param>
        /// <returns>Node from the Manifest (1 item)</returns>
        XmlNode GetManifestRess(string IDRef)
        {
            XmlNode att = null;
            foreach (XmlNode ManifestChildNode in nodeManifest)
            {
                if (ManifestChildNode.Attributes == null)
                    continue;

                att = ManifestChildNode.Attributes["id"];
                if (att == null)
                    continue;

                if (IDRef.Equals(att.Value))
                    return ManifestChildNode;
                else continue;
            }
            return null;
        }

        XmlDocument RebuildHTMLFile(List<ePubDocItem> ePubItemsList)
        {
            XmlDocument htmlFinal = null;
            XmlDocument html = new XmlDocument();
            XmlNode bodyFinal = null;
            foreach (ePubDocItem item in ePubItemsList)
            {
                if (string.IsNullOrEmpty(item.linear))
                    continue;

                if (item.linear.Equals("no"))
                    continue;

                //Retreive href file in ePub file
                opfFilePath = opfFilePath.Replace("\\", "/");
                MemoryStream ms = GetFileFromNameZip(opfFilePath + "/" + item.href);
                if (ms == null)
                    continue; // TODO handle exception

                //Should be an html Doc => Read it
                ms.Position = 0;

                //Clean of DOCTYPE (in order to read the file)
                StreamReader sr = new StreamReader(ms);
                string file = sr.ReadToEnd();
                sr.Close();
                int st = file.IndexOf("<!DOCTYPE");
                int fn = file.IndexOf(">", st + 1);
                if ((st > 0) && (fn > 0))
                {
                    file = file.Remove(st, fn + 1 - st+1);
                }

                //Clean of html namespace (in order to read the file)
                st = file.IndexOf("<html");
                fn = file.IndexOf(">", st + 1);
                if ((st < 0) || (fn < 0))
                {
                    continue;
                }
                string htmltag = file.Substring(st, fn+1 -st);
                file = file.Replace(htmltag, "<html>");
                //ms = new MemoryStream();
                //StreamWriter sw = new StreamWriter(ms);
                //sw.Write(file);
                //sw.BaseStream.Position = 0;

                if (htmlFinal == null)
                { //Copie intégrale
                    htmlFinal  = new XmlDocument();
                    htmlFinal.LoadXml(file);
                    //htmlFinal.Load(ms);
                    XmlNode htmlTag = htmlFinal.DocumentElement;
                    XmlNode body = htmlTag.FirstChild;
                    while (!body.LocalName.Equals("body"))
                    {
                        body = body.NextSibling;
                    }
                    bodyFinal = body;
                }
                else
                { //intégration du body
                    //html.Load(ms);
                    html.LoadXml(file);
                    XmlNode htmlTag = html.DocumentElement;
                    XmlNode body = htmlTag.FirstChild;
                    while (!body.LocalName.Equals("body"))
                    {
                        body = body.NextSibling;
                    }
                    bodyFinal.InnerXml += body.InnerXml;
                }
                
            }

            return htmlFinal;
        }
    }

    class ePubDocItem
    {
        public string idref;
        public string linear;
        public string href;
        public string mediaType;
    }
}
