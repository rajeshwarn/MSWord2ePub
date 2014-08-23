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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            this.TabWord2ePub = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnConvert = this.Factory.CreateRibbonButton();
            this.btnConfig = this.Factory.CreateRibbonButton();
            this.btnAPropos = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.galDecoupe = this.Factory.CreateRibbonGallery();
            this.txtSize = this.Factory.CreateRibbonEditBox();
            this.grpePub = this.Factory.CreateRibbonGroup();
            this.btnCreatePack = this.Factory.CreateRibbonButton();
            this.TabWord2ePub.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.grpePub.SuspendLayout();
            // 
            // TabWord2ePub
            // 
            this.TabWord2ePub.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabWord2ePub.Groups.Add(this.group1);
            this.TabWord2ePub.Groups.Add(this.group2);
            this.TabWord2ePub.Groups.Add(this.grpePub);
            this.TabWord2ePub.Label = "Word2ePub";
            this.TabWord2ePub.Name = "TabWord2ePub";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnConvert);
            this.group1.Items.Add(this.btnConfig);
            this.group1.Items.Add(this.btnAPropos);
            this.group1.Label = "Word2HTML";
            this.group1.Name = "group1";
            // 
            // btnConvert
            // 
            this.btnConvert.Label = "Convertir en HTML5 (pour ePub)";
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConvert_Click);
            // 
            // btnConfig
            // 
            this.btnConfig.Label = "Configuration";
            this.btnConfig.Name = "btnConfig";
            this.btnConfig.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConfig_Click);
            // 
            // btnAPropos
            // 
            this.btnAPropos.Label = "A Propos de Word2HTML4ePub";
            this.btnAPropos.Name = "btnAPropos";
            this.btnAPropos.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAPropos_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.galDecoupe);
            this.group2.Items.Add(this.txtSize);
            this.group2.Label = "Découpe des fichiers";
            this.group2.Name = "group2";
            // 
            // galDecoupe
            // 
            this.galDecoupe.ColumnCount = 1;
            ribbonDropDownItemImpl1.Label = "Aucun";
            ribbonDropDownItemImpl1.Tag = "No";
            ribbonDropDownItemImpl2.Label = "Chapitre";
            ribbonDropDownItemImpl2.Tag = "Chap";
            ribbonDropDownItemImpl3.Label = "Chapitre et Taille max";
            ribbonDropDownItemImpl3.Tag = "Size";
            this.galDecoupe.Items.Add(ribbonDropDownItemImpl1);
            this.galDecoupe.Items.Add(ribbonDropDownItemImpl2);
            this.galDecoupe.Items.Add(ribbonDropDownItemImpl3);
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
            this.grpePub.Label = "ePub";
            this.grpePub.Name = "grpePub";
            // 
            // btnCreatePack
            // 
            this.btnCreatePack.Label = "Créer un Package ePub";
            this.btnCreatePack.Name = "btnCreatePack";
            this.btnCreatePack.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreatePack_Click);
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
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.grpePub.ResumeLayout(false);
            this.grpePub.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConvert;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConfig;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAPropos;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery galDecoupe;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txtSize;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpePub;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreatePack;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabWord2ePub;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonWord2ePub RibbonWord2ePub
        {
            get { return this.GetRibbon<RibbonWord2ePub>(); }
        }
    }
}
