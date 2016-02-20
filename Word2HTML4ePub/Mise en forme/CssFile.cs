using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Word2HTML4ePub
{
    internal class StyleDocList
    {
        internal List<Style> Styles;
        
        public StyleDocList()
        {
            Styles = new List<Style>();
        }

        public void AddNewStyle(string style)
        {
            if (!(from Style s in Styles where s.Balise.ToLower().Equals(style) select s).Any())
                Styles.Add(new Style(style));
        }

        public void AddNewStyleClass(string style, string classe)
        {
            if (!(from Style s in Styles where s.Balise.ToLower().Equals(style) select s).Any())
                Styles.Add(new Style(style));
            Style st = (from Style s in Styles where s.Balise.ToLower().Equals(style) select s).First();
            st.AddClass(classe);
        }

        public static StyleDocList ReadStylesFromStyleFile(string Filename)
        {
            if (!File.Exists(Filename))
                return null;

            StyleDocList sdl = new StyleDocList();

            string[] filelines = File.ReadAllLines(Filename);
            //if (filelines.Length < 2)
            //    return sdl;

            string LastStyle = "";
            for (int i = 1; i < filelines.Length; i++)
            {
                string[] parsedLine = filelines[i].Split(new char[] { '\t' });
                switch (parsedLine.Length)
                {
                    case 0:
                        break;
                    case 1:
                        LastStyle = parsedLine[0];
                        sdl.AddNewStyle(LastStyle);
                        break;
                    case 2:
                        sdl.AddNewStyleClass(LastStyle, parsedLine[1]);
                        break;
                    default:
                        break;
                }
            }

            return sdl;
        }

        internal class Style
        {
            internal string Balise;
            internal List<string> Classes;

            public Style(string Balise)
            {
                this.Balise = Balise;
                this.Classes = new List<string>();
            }

            public void AddClass(string classe)
            {
                if (!(from string c in Classes where c.ToLower().Equals(classe) select c).Any())
                    Classes.Add(classe);
            }
        }
    }

    internal class CssFile
    {
        public static string chaineCommentaire = "Word2ePub";

        public static void UpdateCssFile(string CssFileName, StyleDocList ListOfStyles)
        {
            if (!File.Exists(CssFileName))
                return;

            string css = File.ReadAllText(CssFileName);
            string[] styles = css.Split(new char[] { '}' });
            List<cssRule> rules = new List<cssRule>();

            foreach (string s in styles)
            {
                string[] Parsed = s.Split(new char[] { '{' });
                if (Parsed.Length < 2)
                    continue;

                cssRule rule = new cssRule(Parsed[0]);

                //Recherche d'un éventuel commentaire
                int debcom = Parsed[1].IndexOf("/*" + chaineCommentaire + "*/");
                if (debcom > 0)
                { 
                    rule.comment = chaineCommentaire;
                    rule.styles = Parsed[1].Substring(debcom+("/*" + chaineCommentaire + "*/").Length);
                }
                else
                    rule.styles = Parsed[1];

                rules.Add(rule);
            }

            foreach (StyleDocList.Style s in ListOfStyles.Styles)
            {
                foreach (string c in s.Classes)
                {
                    if (c.Equals("BALISE_SEULE"))
                    {
                        //Recherche de la balise seule
                        if ((from cssRule cssr in rules where cssr.name.Contains(s.Balise) && cssr.classes.Count == 0 select cssr).Any())
                        {
                            //Ajout du commentaire, si elle existe
                            List<cssRule> cssrl = (from cssRule cssr in rules where cssr.name.Contains(s.Balise) && cssr.classes.Count == 0 select cssr).ToList();
                            foreach (cssRule rule in cssrl)
                                rule.comment = chaineCommentaire;
                        }
                        else
                        {
                            //Création si elle n'existe pas
                            cssRule newRule = new cssRule(s.Balise);
                            newRule.comment = chaineCommentaire;
                            rules.Add(newRule);
                        }
                    }
                    else
                    {
                        if ((from cssRule cssr in rules where cssr.classes.Contains(c) select cssr).Any())
                        {
                            List<cssRule> cssrl = (from cssRule cssr in rules where cssr.classes.Contains(c) select cssr).ToList();
                            foreach (cssRule singlerule in cssrl)
                                singlerule.comment = chaineCommentaire;
                        }
                        else
                        {
                            cssRule newRule = new cssRule();
                            newRule.comment = chaineCommentaire;
                            newRule.classes.Add(c);
                            if (s.Balise.Length > 0)
                                newRule.name.Add(s.Balise);
                            rules.Add(newRule);
                        }
                    }
                }
            }

            StringBuilder sb = new StringBuilder();
            foreach (cssRule rule in rules)
            {
                sb.Append(rule.ToString());
            }
            File.WriteAllText(CssFileName, sb.ToString());
        }

        internal class cssRule
        {
            public List<string> name;
            public List<string> classes;
            public List<string> ID;
            public string styles;
            public string comment;

            public cssRule()
            {
                name = new List<string>();
                classes = new List<string>();
                ID = new List<string>();
            }

            public cssRule(string inlineNames)
            {
                name = new List<string>();
                classes = new List<string>();
                ID = new List<string>();

                inlineNames = inlineNames.Replace("\r", string.Empty);
                inlineNames = inlineNames.Replace("\n", string.Empty);
                inlineNames = inlineNames.Replace("\t", string.Empty);

                string[] balises = inlineNames.Split(new char[] { ',' });
                for (int i = 0; i < balises.Length; i++)
                {
                    balises[i] = balises[i].Trim();
                    switch (balises[i][0])
                    {
                        case '.': //Classe
                            classes.Add(balises[i].Substring(1));
                            break;
                        case '#': //ID
                            ID.Add(balises[i].Substring(1));
                            break;
                        default: //Balise
                            name.Add(balises[i]);
                            break;
                    }
                }
            }

            public override string ToString()
            {
                string retour = "";
                bool init = true;
                foreach (string item in name)
                {
                    if (retour.Length > 0)
                        retour += ", ";

                    retour += item;
                    init = false;
                }

                foreach (string item in classes)
                {
                    
                    if ((init) && (retour.Length > 0))
                        retour += ", ";
                    
                    retour += ("." + item);
                    init = false;
                    
                }

                foreach (string item in ID)
                {
                    if (retour.Length > 0)
                        retour += ", ";

                    retour += ("#" + item);
                }

                retour += "\r\n{";
                if (!string.IsNullOrEmpty(comment))
                    retour += ("\r\n\t/*" + comment + "*/");
                if (!string.IsNullOrEmpty(styles))
                    retour += styles;
                else
                    retour += "\r\n";
                retour += "}\r\n\r\n";


                return retour;
            }

        }
    }
}
