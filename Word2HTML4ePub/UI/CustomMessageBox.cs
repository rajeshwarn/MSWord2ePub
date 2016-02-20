using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Word2HTML4ePub
{
    public partial class CustomMessageBox : Form
    {
        public CustomMessageBox()
        {
            InitializeComponent();
        }

        public CustomMessageBox(string Titre, string[] messages)
        {
            InitializeComponent();
            this.Titre = Titre;
            
            StringBuilder sb = new StringBuilder();
            foreach (string s in messages)
                sb.AppendLine(s);

            txtMsg.Text = sb.ToString();
        }

        public CustomMessageBox(string Titre, string messages, string[] separateurs)
        {
            InitializeComponent();
            this.Titre = Titre;

            foreach (string separateur in separateurs)
            {
                int index = 0;
                while (true)
                {
                    index = messages.IndexOf(separateur, index);
                    if (index < 0)
                        break;

                    messages = messages.Insert(index, "\r\n");
                    index += (separateur.Length + 1);
                }
            }
            txtMsg.Text = messages;
        }

        string Titre
        {
            get
            {
                return this.Text;
            }
            set
            {
                this.Text = value;
            }
        }

        private void cmdOK_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
