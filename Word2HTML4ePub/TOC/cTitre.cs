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
        /// Classe pour mettre en place les liens de la TOC
        /// </summary>
        private class Titre
        {
            int[] level;
            string titre;
            string ID;
            int nDuTitre;

            public Titre(int[] level, string titre, int NumDuTitre)
            {
                this.titre = titre;
                this.level = new int[7];
                for (int i = 0; i < 7; i++)
                    this.level[i] = level[i];
                nDuTitre= NumDuTitre;
            }

            public string GetID()
            {
                ID = "id";
                for (int i = 0; i < 7; i++)
                    ID += "." + level[i].ToString();
                return ID;

            }

            public int getLevel()
            {
                for (int i = 6; i > 0; i--)
                {
                    if (level[i] != 0)
                        return i;
                }
                return 0;
            }

            public string getTitre()
            {
                return titre;
            }

            public int NumeroDuTitre
            {
                get
                {
                    return nDuTitre;
                }
            }

        }
    }
}