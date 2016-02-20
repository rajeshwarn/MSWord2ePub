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

        private static string ComputeImageFolder(string[] folderImg, string PackagePath)
        {
            if (folderImg.Length == 1)
            {
                string dirName = Path.GetFileName(folderImg[0]);
                string destFolder = Path.Combine(PackagePath, "content", dirName);
                return destFolder;
            }
            return null;
        }

        private static bool CopyImages(string[] folderImg, string PackagePath)
        {
            string destFolder = ComputeImageFolder(folderImg, PackagePath);
            if (destFolder != null)
            {
                string[] lof = Directory.GetFiles(folderImg[0]);

                if (!Directory.Exists(destFolder))
                    Directory.CreateDirectory(destFolder);

                foreach (string f in lof)
                {
                    if (File.Exists(Path.Combine(destFolder, Path.GetFileName(f))))
                        File.Delete(Path.Combine(destFolder, Path.GetFileName(f)));
                    
                    File.Copy(f, Path.Combine(destFolder, Path.GetFileName(f)));
                }
            }
            return true;
        }

        private static bool DeleteImages(string[] folderImg, string PackagePath)
        {
            string destFolder = ComputeImageFolder(folderImg, PackagePath);
            if (destFolder != null)
            {
                if (!Directory.Exists(destFolder))
                    return true;

                try
                {
                    string[] lof = Directory.GetFiles(destFolder);
                    foreach (string s in lof)
                    {
                        File.Delete(s);
                    }
                    Directory.Delete(destFolder);
                    return true;
                }
                catch (Exception e)
                {
                    System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de supprimer les images");
                    return false;
                }
            }
            return true;
        }

        private static void DownSizeImages(string folderImg, int SizeX, int SizeY)
        {
            if (folderImg != null)
            {
                string[] lof = Directory.GetFiles(folderImg);

                foreach (string f in lof)
                {
                    Traitement_Images.ReduceBitmapDim(f, SizeX, SizeY);
                }
            }
        }

		private static bool TraitementImagesSVG(ref System.Xml.XPath.XPathNavigator lir, ref System.Xml.XPath.XPathNodeIterator it)
		{
		try
            {
                ReportLog("Traitement des images");
                //contrôle de l'ID, et extraction de la liste des fichiers
                string exPath = "/html/body//p//img";
                it = lir.Select(lir.Compile(exPath));
                while (it.MoveNext())
                {
                    //Recup des paramètres de l'image
                    string w = it.Current.GetAttribute("width", it.Current.NamespaceURI);
                    string h = it.Current.GetAttribute("height", it.Current.NamespaceURI);
                    string id = it.Current.GetAttribute("id", it.Current.NamespaceURI);
                    string src = it.Current.GetAttribute("src", it.Current.NamespaceURI);

                    //verif d'une eventuelle légende
                    it.Current.MoveToParent(); // retour dans le paragraphe
                    XPathNodeIterator it1 = it.Clone();
                    it1.Current.MoveToNext(); // deplacement sur le paragraphe suivant
                    string cap = "";
                    if (it1.Current.GetAttribute("class", it.Current.NamespaceURI).Contains("MsoCaption"))
                    {
                        string[] cap1 = it1.Current.Value.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string s in cap1)
                        {
                            cap += s.Trim() + " ";
                        }
                        it1.Current.DeleteSelf();
                    }

                    //Creation d'une balise figure
                    string bal = "<figure><svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"1.1\" width=\"100%\" height=\"90%\" viewBox=\"0 0 ";
                    bal += (w + " " + h + "\" preserveAspectRatio=\"meet\">");
                    bal += "<image width=\"" + w + "\" height=\"" + h + "\" xlink:href=\"" + src + "\" />";
                    bal += "</svg>";
                    bal += 	"<figcaption>" + cap.Trim() + "</figcaption></figure>";
                    it.Current.ReplaceSelf(bal);
                    //it.Current.ReplaceSelf("<figure><img src=\"" + src + "\" id=\"" + id.Replace(" ", null) + "\"/><figcaption>" + cap.Trim() + "</figcaption></figure>");
                    it = lir.Select(lir.Compile(exPath));
                }
				return true;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de traiter les images");
				return false;
            }
		}
		
		private static bool TraitementNoImages(ref System.Xml.XPath.XPathNavigator lir, ref System.Xml.XPath.XPathNodeIterator it)
		{
		try
            {
                ReportLog("Traitement des images : supression ");
                //contrôle de l'ID, et extraction de la liste des fichiers
                string exPath = "/html/body//p//img";
                it = lir.Select(lir.Compile(exPath));
                while (it.MoveNext())
                {
                    //Recup des paramètres de l'image
                    string w = it.Current.GetAttribute("width", it.Current.NamespaceURI);
                    string h = it.Current.GetAttribute("height", it.Current.NamespaceURI);
                    string id = it.Current.GetAttribute("id", it.Current.NamespaceURI);
                    string src = it.Current.GetAttribute("src", it.Current.NamespaceURI);
                    XPathNodeIterator it2 = it.Clone();

                    //verif d'une eventuelle légende
                    it.Current.MoveToParent(); // retour dans le paragraphe
                    XPathNodeIterator it1 = it.Clone();
                    it1.Current.MoveToNext(); // deplacement sur le paragraphe suivant
                    string cap = "";
                    if (it1.Current.GetAttribute("class", it.Current.NamespaceURI).Contains("MsoCaption"))
                    {
                        string[] cap1 = it1.Current.Value.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string s in cap1)
                        {
                            cap += s.Trim() + " ";
                        }
                        it1.Current.DeleteSelf();
                    }

                    //Suppression de la balise img
                    it2.Current.DeleteSelf();
                    it = lir.Select(lir.Compile(exPath));
                }
				return true;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de supprimer les images");
				return false;
            }
		}

        private static bool TraitementImages600x800(string PackagePath, ref System.Xml.XPath.XPathNavigator lir, ref System.Xml.XPath.XPathNodeIterator it)
        {
            try
            {
                ReportLog("Traitement des images");
                //contrôle de l'ID, et extraction de la liste des fichiers
                string exPath = "/html/body//p/img";
                it = lir.Select(lir.Compile(exPath));
                while (it.MoveNext())
                {
                    //Recup des paramètres de l'image
                    string w = it.Current.GetAttribute("width", it.Current.NamespaceURI);
                    string h = it.Current.GetAttribute("height", it.Current.NamespaceURI);
                    string id = it.Current.GetAttribute("id", it.Current.NamespaceURI);
                    string src = it.Current.GetAttribute("src", it.Current.NamespaceURI);

                    //verif d'une eventuelle légende
                    it.Current.MoveToParent(); // retour dans le paragraphe
                    XPathNodeIterator it1 = it.Clone();
                    it1.Current.MoveToNext(); // deplacement sur le paragraphe suivant
                    string cap = "";
                    if (it1.Current.GetAttribute("class", it.Current.NamespaceURI).Contains("MsoCaption"))
                    {
                        string[] cap1 = it1.Current.Value.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string s in cap1)
                        {
                            cap += s.Trim() + " ";
                        }
                        it1.Current.DeleteSelf();
                    }

                    
                    //Recherche du fichier image correspondant (ajustement de la taille)
                    string imfile = Path.Combine(PackagePath, "content", src);
                    if (File.Exists(imfile))
                    {
                        System.Drawing.Size newDim =  Traitement_Images.ReduceBitmapDim(imfile, 600, 800);
                        w = newDim.Width.ToString();
                        h = newDim.Height.ToString();
                    }

                    //Creation d'une balise figure
                    string bal = "<figure><svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"1.1\" width=\"100%\" height=\"90%\" viewBox=\"0 0 ";
                    bal += (w + " " + h + "\" preserveAspectRatio=\"meet\">");
                    bal += "<image width=\"" + w + "\" height=\"" + h + "\" xlink:href=\"" + src + "\" />";
                    bal += "</svg>";
                    bal += "<figcaption>" + cap.Trim() + "</figcaption></figure>";
                    it.Current.ReplaceSelf(bal);
                    //it.Current.ReplaceSelf("<figure><img src=\"" + src + "\" id=\"" + id.Replace(" ", null) + "\"/><figcaption>" + cap.Trim() + "</figcaption></figure>");
                    it = lir.Select(lir.Compile(exPath));
                }
                return true;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible de traiter les images");
                return false;
            }
        }
	}
}