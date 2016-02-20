using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;

namespace Word2HTML4ePub
{
    public partial class ePubTools
    {
        private static string SearchFileByName(List<string> files, string filename)
        {
            return (from str in files where Path.GetFileName(str).Equals(filename, StringComparison.OrdinalIgnoreCase) select str).FirstOrDefault();
        }

        private static string SearchFileByPartOfTheName(List<string> files, string filename)
        {
            return (from str in files where Path.GetFileName(str).Contains(filename) select str).FirstOrDefault();
        }

        private static string SearchFileByPartOfThePath(List<string> files, string filename)
        {
            return (from str in files where str.Contains(filename) select str).FirstOrDefault();
        }

        private static string SearchNavDoc(XmlDocument doc)
        {
            XmlElement nodePackage = doc.DocumentElement;
            XmlNode nodemanifest = nodePackage.FirstChild;
            while ((nodemanifest != null) && (!nodemanifest.Name.Contains("manifest")))
            {
                nodemanifest = nodemanifest.NextSibling;
            }

            foreach (XmlElement nodeitem in nodemanifest.ChildNodes)
            {
                string properties = nodeitem.GetAttribute("properties");
                if (string.IsNullOrEmpty(properties))
                    continue;

                if (properties.Contains("nav"))
                    return nodeitem.GetAttribute("href");
            }
            return null;
        }

        private static string SearchOPFDoc(string containerxml)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(containerxml);

            XmlNode root = doc.DocumentElement;
            XmlNode rootfiles = root.FirstChild;
            while (rootfiles != null)
            {
                if (rootfiles.Name.Contains("rootfiles"))
                    break;
    
                rootfiles = rootfiles.NextSibling;
            }

            if (rootfiles == null)
                return null;

            XmlNode rootfile = rootfiles.FirstChild;
            while (rootfile != null)
            {
                if (rootfile.Name.Contains("rootfile"))
                    break;

                rootfile = rootfile.NextSibling;
            }
            
            if (rootfile == null)
                return null;

            return ((XmlElement)rootfile).GetAttribute("full-path");
        }
    }
}