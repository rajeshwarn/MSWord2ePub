using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Core;
using System.Reflection;
using System.IO;
using System.Xml;
using System.Xml.XPath;
using System.Windows.Forms;

namespace Word2HTML4ePub
{
    public partial class WordHTML2ePubHTML
    {
        /// <summary>
        /// Fonction utilisée pour extraire une TOC.
        /// </summary>
        /// <param name="lir"></param>
        /// <returns></returns>
        private static NavTable ExtractTOC(ref XPathNavigator lir)
        {
            ReportLog("Extraction de la TOC");

            try
            {
                //Extraction des titres pour la balise nav, puis ajout d'un id
                string exPath = "/html/body";
                XPathNodeIterator it = lir.Select(lir.Compile(exPath));
                NavTable nav = new NavTable();
                exPath = ".//h1|.//h2|.//h3|.//h4|.//h5|.//h6|.//h7";
                it = lir.Select(lir.Compile(exPath));
                while (it.MoveNext())
                {
                    if (it.Current.HasAttributes)
                    {
                        //Check s'il existe un attribut id
                    }
                    NavTable.Level curlev = (NavTable.Level)Enum.Parse(typeof(NavTable.Level), it.Current.Name);
                    string id = nav.AddTitre(curlev, it.Current.Value);
                    it.Current.CreateAttribute(null, "id", null, id);
                }

                return nav;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible d'extraire une toc");
                return null;
            }

        }

        /// <summary>
        /// Classe servant pour l'extraction d'une TOC à partir du doc original.
        /// Cette classe et ses fonctions ne sont pas censées être appelées par une autre fonction que extractTOC.
        /// </summary>
        private class NavTable
        {
            int[] currentLevel;
            //Level LastLevel = 0;
            Level maxlevel = 0;
            List<Titre> loTitre;

            public NavTable()
            {
                currentLevel = new int[7];
                loTitre = new List<Titre>();
            }

            public enum Level { h1 = 0, h2, h3, h4, h5, h6, h7 };

            public string ExportNavTable(string FileName)
            {

                int lastlevel = 0;
                string navstr = "<?xml version='1.0' encoding='UTF-8'?>\r";
                navstr += "<!DOCTYPE html>\r";
                navstr += "<html xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:epub=\"http://www.idpf.org/2007/ops\">\r";
                navstr += "<head><title>Navigation</title>\r";
                navstr += "<meta name=\"Generator\" content=\"Word2ePub_" + Assembly.GetExecutingAssembly().GetName().Version.ToString() + "\" />\r";
                navstr += "<link rel=\"Stylesheet\" href=\"style.css\"  type=\"text/css\" />";
                navstr += "</head>\r<body>\r";
                navstr += "<nav epub:type='toc' >\r<h1>Table des matières</h1>\r<ol>\r";
                for (int i = 0; i < loTitre.Count; i++)
                {
                    string nodestr = "<li><a href=\"" + Path.GetFileName(FileName) + "#" + loTitre[i].GetID() + "\">" + loTitre[i].getTitre() + "</a></li>\r";
                    int newLevel = loTitre[i].getLevel();
                    if (lastlevel < newLevel)
                    {
                        navstr += "<ol>\r" + nodestr;
                        lastlevel = newLevel;
                    }
                    else if (lastlevel == newLevel)
                        navstr += nodestr;
                    else
                    {
                        for (int j = newLevel; j < lastlevel; j++)
                            navstr += "</ol>\r";

                        navstr += nodestr;
                    }
                }
                navstr += "</ol>\r</nav>\r";
                navstr += "</body>\r</html>\r";

                string newFileName = Path.Combine(Path.GetDirectoryName(FileName), Path.GetFileNameWithoutExtension(FileName) + "-nav.html");
                File.WriteAllText(newFileName, navstr, Encoding.UTF8);
                return newFileName;
            }

            public string ExportNavTableSplittedbyChap(string FileNameBase)
            {
                ReportLog("Création de la table de navigation");

                try
                {

                    int lastlevel = -1;
                    string navstr = "<?xml version='1.0' encoding='UTF-8'?>\r";
                    navstr += "<!DOCTYPE html>\r";
                    navstr += "<html xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:epub=\"http://www.idpf.org/2007/ops\">\r";
                    navstr += "<head><title>Navigation</title>\r";
                    navstr += "<meta name=\"Generator\" content=\"Word2ePub_" + Assembly.GetExecutingAssembly().GetName().Version.ToString() + "\" />\r";
                    navstr += "<link rel=\"Stylesheet\" href=\"style.css\"  type=\"text/css\" />";
                    navstr += "</head>\r<body>\r";
                    navstr += "<nav epub:type='toc' >\r<h1>Table des matières</h1>\r";
                    for (int i = 0; i < loTitre.Count; i++)
                    {
                        string nodestr = "<li><a href=\"" + Path.GetFileName(FileNameBase) + "-" + loTitre[i].NumeroDuTitre.ToString() + ".html#" + loTitre[i].GetID() + "\">" + loTitre[i].getTitre() + "</a>";
                        int newLevel = loTitre[i].getLevel();
                        if (lastlevel < newLevel)
                        {
                            navstr += "\r<ol>\r";
						    navstr+= nodestr;
                            lastlevel = newLevel;
                        }
                        else if (lastlevel == newLevel)
					    {
						    navstr += "</li>\r";
                            navstr += nodestr;
					    }
                        else
                        {
                            navstr += "</li>\r";
                            for (int j = newLevel; j < lastlevel; j++)
                                navstr += "</ol>\r</li>\r";
                            navstr += nodestr;
						    lastlevel = newLevel;
                        }
                    }
				    for (int j = -1; j < lastlevel; j++) //-1 initialement...
					    navstr += "</li>\r</ol>\r";

                    navstr += "</nav>\r";
                    navstr += "</body>\r</html>\r";

                    string newFileName = Path.Combine(Path.GetDirectoryName(FileNameBase), Path.GetFileNameWithoutExtension(FileNameBase) + "-nav.html");
                    File.WriteAllText(newFileName, navstr, Encoding.UTF8);
                    return newFileName;
                }
                catch (Exception e)
                {
                    System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible d'extraire une toc");
                    return null;
                }

            }

            public string AddTitre(Level level, string titre)
            {
                currentLevel[(int)level]++;
                for (int i = (int)(level) + 1; i < 7; i++)
                    currentLevel[i] = 0;

                if (maxlevel < level)
                    maxlevel = level;

                Titre t = new Titre(currentLevel, titre, loTitre.Count+1);
                loTitre.Add(t);
                return t.GetID();
            }

            public int NbOfChap
            {
                get
                {
                    return loTitre.Count;
                }
            }

            public bool SplitTextes(int Chap ,out string debut, out string fin)
            {
                debut = null;
                fin = null;
                if (Chap > loTitre.Count)
                    return false;

                if (Chap < loTitre.Count)
                    fin = loTitre[Chap].GetID();

                if (Chap >0)
                    debut = loTitre[Chap-1].GetID();

                return true;
            }   
        }
    }
}