using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.XPath;
using System.IO;

namespace Word2HTML4ePub
{
    public partial class WordHTML2ePubHTML
    {
        public static bool ExtractStyleList(XPathNavigator lir)
        {
            ReportLog("Extraction des styles utilisés (styles.txt)");
            try
            {                //Extraction des styles pour la mise en forme css

                System.Xml.XPath.XPathNodeIterator it;
                List<BaliseClass> ListeDesBalises = new List<BaliseClass>();

                //Suppression de l'en tête de l'epub pour le traitement du fichier...
                lir.InnerXml = lir.InnerXml.Replace("<html xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:epub=\"http://www.idpf.org/2007/ops\">", "<html>");
                //On sélectionne toutes les balises du body
                it = lir.Select(lir.Compile("/html/body//*"));
                if (it.Count == 0)
                    return true;
                it.MoveNext();

                do
                {
                    //On va parcourrir toutes les balises en analysant les balises et les classes associées.
                    BaliseClass bal = null;
                    foreach (BaliseClass b in ListeDesBalises)
                    {
                        if (!b.Name.Equals(it.Current.Name))
                            continue;

                        bal = b;
                    }
                    // si bal est null, alors la balise n'existe pas encore dans la liste...
                    if (bal == null)
                    {
                        bal = new BaliseClass(it.Current.Name);
                        ListeDesBalises.Add(bal);
                    }

                    //Verif des attributs
                    if (it.Current.HasAttributes)
                    {
                        it.Current.MoveToFirstAttribute();
                        do
                        {
                            if (it.Current.Name.Equals("class"))
                                bal.AddClass(it.Current.Value);
                        } while (it.Current.MoveToNextAttribute());
                    }
                    else
                    {
                        if (!bal.classes.Contains("BALISE_SEULE"))
                            bal.AddClass("BALISE_SEULE");
                    }

                } while (it.MoveNext());

                string styles = BaliseClass.ToStringHeader();
                foreach (BaliseClass b in ListeDesBalises)
                {
                    styles += b.ToString();
                }
                Uri urifile = new Uri(lir.BaseURI);
                string styleFile = Path.Combine(Path.GetDirectoryName(urifile.LocalPath), "styles.txt");
                File.WriteAllText(styleFile, styles);
                return true;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Message d'erreur:\r" + e.Message, "Impossible d'extraire une toc");
                return false;
            }

        }

        class BaliseClass
        {
            public List<string> classes;
            string pName;
            public string Name
            {
                get
                {
                    return pName;
                }
            }

            public BaliseClass(string BaliseName)
            {
                pName = BaliseName;
                classes = new List<string>();
            }

            public void AddClass(string Classe)
            {
                foreach (string cl in classes)
                {
                    if (cl.Equals(Classe))
                        return;
                }
                classes.Add(Classe);
            }

            public override string ToString()
            {
                string retour = Name + "\r\n";
                foreach (string cl in classes)
                    retour += ("\t" + cl + "\r\n");
                return (retour + "\r\n");
            }

            public static string ToStringHeader()
            {
                return "Balise\tClasse\r\n";
            }

        }
    }
}
