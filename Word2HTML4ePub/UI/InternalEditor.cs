using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml;

namespace Word2HTML4ePub
{
    public partial class InternalEditor : Form
    {
        string pFilePath = null;
        public string FilePath
        {
            get
            {
                return pFilePath;
            }
        }

        public InternalEditor()
        {
            InitializeComponent();

            if (string.IsNullOrEmpty(FilePath))
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Title = "Choisir le fichier à éditer";
                ofd.Filter = "Fichiers htm*|*.htm*|Tous|*.*";
                ofd.Multiselect = false;

                DialogResult dr = ofd.ShowDialog();
                if (dr != System.Windows.Forms.DialogResult.OK)
                    this.Close();

                if (File.Exists(ofd.FileName))
                    pFilePath = ofd.FileName;
            }
            
            
            using (MemoryStream ms = new MemoryStream())
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(FilePath);
                StringBuilder builder = new StringBuilder();
                using (XmlTextWriter writer = new XmlTextWriter(new StringWriter(builder)))
                {
                    writer.Formatting = Formatting.Indented;
                    doc.Save(ms);
                }
                
                ms.Position = 0;
                StreamReader sr = new StreamReader(ms);
                txtSrc.Text= sr.ReadToEnd();

            }
            
            webVisu.Url = new Uri(FilePath);

        }

    }
}
