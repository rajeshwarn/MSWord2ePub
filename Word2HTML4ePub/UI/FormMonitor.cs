using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace Word2HTML4ePub
{
    public partial class FormMonitor : Form
    {
        Thread tr = null;
        volatile Microsoft.Office.Interop.Word.Document documentToProcess;
        DateTime debut;
        volatile Decoupe decoupe;
        volatile TraitementImages traitementImg;
        volatile int SizeMaxKo;
        volatile string PackagePath = null;

        //public FormMonitor()
        //{
        //    InitializeComponent();
        //}

        /// <summary>
        /// Type de découpe. Les valeurs apparaitront dans la boite à choisir
        /// </summary>
        public enum Decoupe { Aucun, Chapitre, ChapitresTailleMax };

        /// <summary>
        /// Type de Traitement pour les Images. Les valeurs apparaitront dans la boite à choisir
        /// </summary>
        public enum TraitementImages { SansTraitement, Convert2SVG, NoImage, Resize600x800 };

        public FormMonitor(Microsoft.Office.Interop.Word.Document doc)
        {
            InitializeComponent();
            documentToProcess = doc;
            this.decoupe = Decoupe.Aucun;
            this.traitementImg = TraitementImages.SansTraitement;
            SizeMaxKo = 0;
        }

        public FormMonitor(Microsoft.Office.Interop.Word.Document doc, Decoupe decoupe)
        {
            InitializeComponent();
            documentToProcess = doc;
            this.decoupe = decoupe;
            this.traitementImg = TraitementImages.SansTraitement;
            SizeMaxKo = 40;
        }

        public FormMonitor(Microsoft.Office.Interop.Word.Document doc, Decoupe decoupe, int TailleMax)
        {
            InitializeComponent();
            documentToProcess = doc;
            this.decoupe = decoupe;
            this.traitementImg = TraitementImages.SansTraitement;
            SizeMaxKo = TailleMax;
        }

        public FormMonitor(Microsoft.Office.Interop.Word.Document doc, Decoupe decoupe, int TailleMax, string PackagePath)
        {
            InitializeComponent();
            documentToProcess = doc;
            this.decoupe = decoupe;
            this.traitementImg = TraitementImages.SansTraitement;
            SizeMaxKo = TailleMax;
            this.PackagePath = PackagePath;
        }

        public FormMonitor(Microsoft.Office.Interop.Word.Document doc, Decoupe decoupe, TraitementImages traitementImg, int TailleMax, string PackagePath)
        {
            InitializeComponent();
            documentToProcess = doc;
            this.decoupe = decoupe;
            this.traitementImg = traitementImg;
            SizeMaxKo = TailleMax;
            this.PackagePath = PackagePath;
        }

        private void cmdQuitter_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        delegate void SetTextCallback(string message);

        private void PrintLog(string message)
        {
            if (this.txtLog.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(PrintLog);
                this.Invoke(d, new object[] { message });
            }
            else
            {
                this.txtLog.Text += (" (" + (DateTime.Now-debut).TotalSeconds.ToString("F2") + "s)\r\n" + message );
                this.txtLog.SelectionStart = this.txtLog.Text.Length;
                this.txtLog.ScrollToCaret();
            }
        }

        delegate void SetButtonEnableCallback(Button bouton, bool EnableState);
        private void SetButtonEnable(Button bouton, bool State)
        {
            if (bouton.InvokeRequired)
            {
                SetButtonEnableCallback d = new SetButtonEnableCallback(SetButtonEnable);
                this.Invoke(d, new object[] { bouton, State});
            }
            else
            {
                bouton.Enabled = State;
            }
        }

        private void FormMonitor_Load(object sender, EventArgs e)
        {
            Cursor = System.Windows.Forms.Cursors.WaitCursor;
            this.Size = new Size(400, 300);
            SetButtonEnable(cmdQuitter, false);
            SetButtonEnable(cmdCancel, true);

            WordHTML2ePubHTML.ReportLog += PrintLog;
            WordHTML2ePubHTML.Progress += SetProgress;
            
            debut = DateTime.Now;

            txtLog.Text = "Début de process";

            Thread.Sleep(0);

            if (!WordHTML2ePubHTML.PreProcessDoc(documentToProcess))
                return;

            tr = new Thread(StartProcess);
            tr.Start();

            Thread.Sleep(0);
            //tr.Join();
        }

        void StartProcess()
        {
            //WordHTML2ePubHTML.ProcessDoc(documentToProcess);
            WordHTML2ePubHTML.ProcessDoc(decoupe, traitementImg, SizeMaxKo, PackagePath);
            SetButtonEnable(cmdCancel, false);
            SetButtonEnable(cmdQuitter, true);
            PrintLog("Fin de l'execution");
            SetCursor(System.Windows.Forms.Cursors.Default);
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            if (tr != null)
            {
                if (tr.IsAlive)
                {
                    tr.Abort();
                    PrintLog("Arrêt demandé");
                }
                while (tr.IsAlive)
                {
                    Thread.Sleep(50);
                }
            }
            SetButtonEnable(cmdQuitter, true);
            SetButtonEnable(cmdCancel, false);
            PrintLog("Fin prématurée de l'execution");
            SetCursor(System.Windows.Forms.Cursors.Default);
        }

        private void FormMonitor_Resize(object sender, EventArgs e)
        {
            int marges = 10;
            txtLog.Top = marges;
            txtLog.Left = marges;
            txtLog.Width = this.ClientRectangle.Width - 2 * marges;
            txtLog.Height = this.ClientRectangle.Height - 3 * marges - cmdCancel.Height;
            txtLog.ScrollToCaret();

            Progress.Top = txtLog.Bottom + marges;
            Progress.Left = marges;
            cmdQuitter.Top = Progress.Top;
            cmdCancel.Top = Progress.Top;
            cmdQuitter.Left = txtLog.Right - cmdQuitter.Width;
            cmdCancel.Left = cmdQuitter.Left - cmdCancel.Width - marges;
            Progress.Width = cmdCancel.Left - 2 * marges;
        }

        delegate void SetCursorCallback(Cursor curseur);
        private void SetCursor(Cursor curseur)
        {

            if (this.InvokeRequired)
            {
                SetCursorCallback d = new SetCursorCallback(SetCursor);
                this.Invoke(d, new object[] { curseur });
            }
            else
            {
                this.Cursor = curseur; //System.Windows.Forms.Cursors.WaitCursor;
            }
        }

        delegate void SetProgressCallback(int current, int max);
        private void SetProgress(int current, int max)
        {
            if (this.Progress.InvokeRequired)
            {
                SetProgressCallback d = new SetProgressCallback(SetProgress);
                this.Invoke(d, new object[] { current, max });
            }
            else
            {
                this.Progress.Maximum = max;
                this.Progress.Value = current;
                this.Progress.Refresh();
            }
        }

        private void FormMonitor_FormClosed(object sender, FormClosedEventArgs e)
        {
            WordHTML2ePubHTML.ReportLog -= PrintLog;
            WordHTML2ePubHTML.Progress -= SetProgress;
        }
        


    }
}
