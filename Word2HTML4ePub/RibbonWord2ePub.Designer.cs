namespace Word2HTML4ePub
{
    partial class RibbonWord2ePub : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonWord2ePub()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur de composants

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            this.TabWord2ePub = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnConvert = this.Factory.CreateRibbonButton();
            this.galDecoupe = this.Factory.CreateRibbonGallery();
            this.txtSize = this.Factory.CreateRibbonEditBox();
            this.grpePub = this.Factory.CreateRibbonGroup();
            this.btnCreatePack = this.Factory.CreateRibbonButton();
            this.btnEditCfg = this.Factory.CreateRibbonButton();
            this.btnExportToPack = this.Factory.CreateRibbonButton();
            this.cmdEditCss = this.Factory.CreateRibbonButton();
            this.cmdGeneEPub = this.Factory.CreateRibbonButton();
            this.grpePub2Word = this.Factory.CreateRibbonGroup();
            this.btnLoadePub = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.galImages = this.Factory.CreateRibbonGallery();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.cmdAbout = this.Factory.CreateRibbonButton();
            this.cmdTuto = this.Factory.CreateRibbonButton();
            this.TabWord2ePub.SuspendLayout();
            this.group1.SuspendLayout();
            this.grpePub.SuspendLayout();
            this.grpePub2Word.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // TabWord2ePub
            // 
            this.TabWord2ePub.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabWord2ePub.Groups.Add(this.group1);
            this.TabWord2ePub.Groups.Add(this.grpePub);
            this.TabWord2ePub.Groups.Add(this.grpePub2Word);
            this.TabWord2ePub.Groups.Add(this.group3);
            this.TabWord2ePub.Label = "Word2ePub";
            this.TabWord2ePub.Name = "TabWord2ePub";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnConvert);
            this.group1.Items.Add(this.txtSize);
            this.group1.Label = "Word2HTML";
            this.group1.Name = "group1";
            this.group1.Visible = false;
            // 
            // btnConvert
            // 
            this.btnConvert.Label = "Convertir en HTML5 (pour ePub)";
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConvert_Click);
            // 
            // galDecoupe
            // 
            this.galDecoupe.ColumnCount = 1;
            ribbonDropDownItemImpl4.Label = "Aucun";
            ribbonDropDownItemImpl4.Tag = "No";
            ribbonDropDownItemImpl5.Label = "Chapitre";
            ribbonDropDownItemImpl5.Tag = "Chap";
            this.galDecoupe.Items.Add(ribbonDropDownItemImpl4);
            this.galDecoupe.Items.Add(ribbonDropDownItemImpl5);
            this.galDecoupe.Label = "Découpe";
            this.galDecoupe.Name = "galDecoupe";
            this.galDecoupe.RowCount = 3;
            this.galDecoupe.ShowItemSelection = true;
            this.galDecoupe.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.galDecoupe_Click);
            // 
            // txtSize
            // 
            this.txtSize.Enabled = false;
            this.txtSize.Label = "Taille Max (ko):";
            this.txtSize.Name = "txtSize";
            this.txtSize.Text = "40";
            // 
            // grpePub
            // 
            this.grpePub.Items.Add(this.btnCreatePack);
            this.grpePub.Items.Add(this.btnEditCfg);
            this.grpePub.Items.Add(this.btnExportToPack);
            this.grpePub.Items.Add(this.cmdEditCss);
            this.grpePub.Items.Add(this.cmdGeneEPub);
            this.grpePub.Label = "Word=>ePub";
            this.grpePub.Name = "grpePub";
            // 
            // btnCreatePack
            // 
            this.btnCreatePack.Label = "Créer un Package ePub";
            this.btnCreatePack.Name = "btnCreatePack";
            this.btnCreatePack.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreatePack_Click);
            // 
            // btnEditCfg
            // 
            this.btnEditCfg.Label = "Editer la configuration";
            this.btnEditCfg.Name = "btnEditCfg";
            this.btnEditCfg.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConfig_Click);
            // 
            // btnExportToPack
            // 
            this.btnExportToPack.Label = "Exportation vers le Package ePub";
            this.btnExportToPack.Name = "btnExportToPack";
            this.btnExportToPack.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportToPack_Click);
            // 
            // cmdEditCss
            // 
            this.cmdEditCss.Label = "Editer feuilles de style";
            this.cmdEditCss.Name = "cmdEditCss";
            this.cmdEditCss.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cmdEditCss_Click);
            // 
            // cmdGeneEPub
            // 
            this.cmdGeneEPub.Label = "Générer ePub";
            this.cmdGeneEPub.Name = "cmdGeneEPub";
            this.cmdGeneEPub.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cmdGeneEPub_Click);
            // 
            // grpePub2Word
            // 
            this.grpePub2Word.Items.Add(this.btnLoadePub);
            this.grpePub2Word.Label = "ePub=>Word";
            this.grpePub2Word.Name = "grpePub2Word";
            // 
            // btnLoadePub
            // 
            this.btnLoadePub.Label = "Charger un ePub";
            this.btnLoadePub.Name = "btnLoadePub";
            this.btnLoadePub.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadePub_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.label1);
            this.group3.Items.Add(this.galImages);
            this.group3.Items.Add(this.galDecoupe);
            this.group3.Items.Add(this.separator1);
            this.group3.Items.Add(this.cmdAbout);
            this.group3.Items.Add(this.cmdTuto);
            this.group3.Label = "Word2ePub";
            this.group3.Name = "group3";
            // 
            // label1
            // 
            this.label1.Label = "Configuration";
            this.label1.Name = "label1";
            // 
            // galImages
            // 
            this.galImages.ColumnCount = 1;
            ribbonDropDownItemImpl1.Label = "Images";
            ribbonDropDownItemImpl2.Label = "NoImages";
            ribbonDropDownItemImpl3.Label = "Resized";
            this.galImages.Items.Add(ribbonDropDownItemImpl1);
            this.galImages.Items.Add(ribbonDropDownItemImpl2);
            this.galImages.Items.Add(ribbonDropDownItemImpl3);
            this.galImages.Label = "Traitement des images";
            this.galImages.Name = "galImages";
            this.galImages.RowCount = 3;
            this.galImages.ShowItemSelection = true;
            this.galImages.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.galImages_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // cmdAbout
            // 
            this.cmdAbout.Label = "A propos de Word2ePub";
            this.cmdAbout.Name = "cmdAbout";
            this.cmdAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cmdAbout_Click);
            // 
            // cmdTuto
            // 
            this.cmdTuto.Label = "Tutorial Word2ePub";
            this.cmdTuto.Name = "cmdTuto";
            this.cmdTuto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cmdTuto_Click);
            // 
            // RibbonWord2ePub
            // 
            this.Name = "RibbonWord2ePub";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.TabWord2ePub);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonWord2ePub_Load);
            this.TabWord2ePub.ResumeLayout(false);
            this.TabWord2ePub.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.grpePub.ResumeLayout(false);
            this.grpePub.PerformLayout();
            this.grpePub2Word.ResumeLayout(false);
            this.grpePub2Word.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConvert;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery galDecoupe;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txtSize;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpePub;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreatePack;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabWord2ePub;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEditCfg;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportToPack;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton cmdGeneEPub;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton cmdEditCss;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton cmdAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpePub2Word;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadePub;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery galImages;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton cmdTuto;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonWord2ePub RibbonWord2ePub
        {
            get { return this.GetRibbon<RibbonWord2ePub>(); }
        }
    }
}
