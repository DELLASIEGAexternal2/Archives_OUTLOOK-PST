namespace OutlookPST
{
    partial class RibbonMesArchivesExplorer : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonMesArchivesExplorer()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonMesArchivesExplorer));
            this.tabRestorePST = this.Factory.CreateRibbonTab();
            this.grRestauration_av = this.Factory.CreateRibbonGroup();
            this.btn1 = this.Factory.CreateRibbonButton();
            this.btn2 = this.Factory.CreateRibbonButton();
            this.btn3 = this.Factory.CreateRibbonButton();
            this.btn4 = this.Factory.CreateRibbonButton();
            this.btn5 = this.Factory.CreateRibbonButton();
            this.btn6 = this.Factory.CreateRibbonButton();
            this.btn7 = this.Factory.CreateRibbonButton();
            this.btn8 = this.Factory.CreateRibbonButton();
            this.btn9 = this.Factory.CreateRibbonButton();
            this.btn10 = this.Factory.CreateRibbonButton();
            this.btAccueil = this.Factory.CreateRibbonButton();
            this.btModop = this.Factory.CreateRibbonButton();
            this.btAPropos = this.Factory.CreateRibbonButton();
            this.btCryptMdp = this.Factory.CreateRibbonButton();
            this.bgwTreadDwnLoad = new System.ComponentModel.BackgroundWorker();
            this.tabRestorePST.SuspendLayout();
            this.grRestauration_av.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabRestorePST
            // 
            this.tabRestorePST.Groups.Add(this.grRestauration_av);
            this.tabRestorePST.Label = "Mes archives";
            this.tabRestorePST.Name = "tabRestorePST";
            // 
            // grRestauration_av
            // 
            this.grRestauration_av.Items.Add(this.btn1);
            this.grRestauration_av.Items.Add(this.btn2);
            this.grRestauration_av.Items.Add(this.btn3);
            this.grRestauration_av.Items.Add(this.btn4);
            this.grRestauration_av.Items.Add(this.btn5);
            this.grRestauration_av.Items.Add(this.btn6);
            this.grRestauration_av.Items.Add(this.btn7);
            this.grRestauration_av.Items.Add(this.btn8);
            this.grRestauration_av.Items.Add(this.btn9);
            this.grRestauration_av.Items.Add(this.btn10);
            this.grRestauration_av.Items.Add(this.btAccueil);
            this.grRestauration_av.Items.Add(this.btModop);
            this.grRestauration_av.Items.Add(this.btAPropos);
            this.grRestauration_av.Items.Add(this.btCryptMdp);
            this.grRestauration_av.Label = "Mes archives mails ";
            this.grRestauration_av.Name = "grRestauration_av";
            // 
            // btn1
            // 
            this.btn1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn1.Description = "Restauration de l\'archive mail pour la boite séléctionnée, pour 2015";
            this.btn1.Image = global::OutlookPST.Properties.Resources.ico_Blanc;
            this.btn1.ImageName = "2015";
            this.btn1.Label = "2015";
            this.btn1.Name = "btn1";
            this.btn1.ScreenTip = "Restauration de l\'archive mail pour la boite séléctionnée, pour 2015";
            this.btn1.ShowImage = true;
            this.btn1.Visible = false;
            this.btn1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn1_Click);
            // 
            // btn2
            // 
            this.btn2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn2.Description = "Restauration de l\'archive mail pour la boite séléctionnée, pour 2016";
            this.btn2.Image = global::OutlookPST.Properties.Resources.ico_Blanc;
            this.btn2.ImageName = "2016";
            this.btn2.Label = "2016";
            this.btn2.Name = "btn2";
            this.btn2.ScreenTip = "Restauration de l\'archive mail pour la boite séléctionnée, pour 2016";
            this.btn2.ShowImage = true;
            this.btn2.Visible = false;
            this.btn2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn2_Click);
            // 
            // btn3
            // 
            this.btn3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn3.Description = "Restauration de l\'archive mail pour la boite séléctionnée, pour 2017";
            this.btn3.Image = global::OutlookPST.Properties.Resources.ico_Blanc;
            this.btn3.ImageName = "2017";
            this.btn3.Label = "2017";
            this.btn3.Name = "btn3";
            this.btn3.ScreenTip = "Restauration de l\'archive mail pour la boite séléctionnée, pour 2017";
            this.btn3.ShowImage = true;
            this.btn3.Visible = false;
            this.btn3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn3_Click);
            // 
            // btn4
            // 
            this.btn4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn4.Description = "Restauration de l\'archive mail pour la boite séléctionnée, pour 2018";
            this.btn4.Image = global::OutlookPST.Properties.Resources.ico_Blanc;
            this.btn4.ImageName = "2018";
            this.btn4.Label = "2018";
            this.btn4.Name = "btn4";
            this.btn4.ScreenTip = "Restauration de l\'archive mail pour la boite séléctionnée, pour 2018";
            this.btn4.ShowImage = true;
            this.btn4.Visible = false;
            this.btn4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn4_Click);
            // 
            // btn5
            // 
            this.btn5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn5.Description = "Restauration de l\'archive mail pour la boite séléctionnée, pour 2019";
            this.btn5.Image = global::OutlookPST.Properties.Resources.ico_Blanc;
            this.btn5.ImageName = "2019";
            this.btn5.Label = "2019";
            this.btn5.Name = "btn5";
            this.btn5.ScreenTip = "Restauration de l\'archive mail pour la boite séléctionnée, pour 2019";
            this.btn5.ShowImage = true;
            this.btn5.Visible = false;
            this.btn5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn5_Click);
            // 
            // btn6
            // 
            this.btn6.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn6.Description = "Restauration de l\'archive mail pour la boite séléctionnée, pour 2020";
            this.btn6.Image = global::OutlookPST.Properties.Resources.ico_Blanc;
            this.btn6.ImageName = "2020";
            this.btn6.Label = "2020";
            this.btn6.Name = "btn6";
            this.btn6.ScreenTip = "Restauration de l\'archive mail pour la boite séléctionnée, pour 2020";
            this.btn6.ShowImage = true;
            this.btn6.Visible = false;
            this.btn6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn6_Click);
            // 
            // btn7
            // 
            this.btn7.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn7.Description = "Restauration de l\'archive mail pour la boite séléctionnée, pour 2021";
            this.btn7.Image = global::OutlookPST.Properties.Resources.ico_Blanc;
            this.btn7.ImageName = "2021";
            this.btn7.Label = "2021";
            this.btn7.Name = "btn7";
            this.btn7.ScreenTip = "Restauration de l\'archive mail pour la boite séléctionnée, pour 2021";
            this.btn7.ShowImage = true;
            this.btn7.Visible = false;
            this.btn7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn7_Click);
            // 
            // btn8
            // 
            this.btn8.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn8.Description = "Restauration de l\'archive mail pour la boite séléctionnée, pour 2022";
            this.btn8.Image = global::OutlookPST.Properties.Resources.ico_Blanc;
            this.btn8.ImageName = "2022";
            this.btn8.Label = "2022";
            this.btn8.Name = "btn8";
            this.btn8.ScreenTip = "Restauration de l\'archive mail pour la boite séléctionnée, pour 2022";
            this.btn8.ShowImage = true;
            this.btn8.Visible = false;
            this.btn8.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn8_Click);
            // 
            // btn9
            // 
            this.btn9.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn9.Description = "Restauration de l\'archive mail pour la boite séléctionnée, pour 2023";
            this.btn9.Image = global::OutlookPST.Properties.Resources.ico_Blanc;
            this.btn9.ImageName = "2023";
            this.btn9.Label = "2023";
            this.btn9.Name = "btn9";
            this.btn9.ScreenTip = "Restauration de l\'archive mail pour la boite séléctionnée, pour 2023";
            this.btn9.ShowImage = true;
            this.btn9.Visible = false;
            this.btn9.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn9_Click);
            // 
            // btn10
            // 
            this.btn10.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn10.Description = "Restauration de l\'archive mail pour la boite séléctionnée, pour 2024";
            this.btn10.Image = global::OutlookPST.Properties.Resources.ico_Blanc;
            this.btn10.ImageName = "2024";
            this.btn10.Label = "2024";
            this.btn10.Name = "btn10";
            this.btn10.ScreenTip = "Restauration de l\'archive mail pour la boite séléctionnée, pour 2024";
            this.btn10.ShowImage = true;
            this.btn10.Visible = false;
            this.btn10.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn10_Click);
            // 
            // btAccueil
            // 
            this.btAccueil.Image = ((System.Drawing.Image)(resources.GetObject("btAccueil.Image")));
            this.btAccueil.Label = "Accueil";
            this.btAccueil.Name = "btAccueil";
            this.btAccueil.ShowImage = true;
            this.btAccueil.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btAccueil_Click);
            // 
            // btModop
            // 
            this.btModop.Image = ((System.Drawing.Image)(resources.GetObject("btModop.Image")));
            this.btModop.Label = "Modop";
            this.btModop.Name = "btModop";
            this.btModop.ShowImage = true;
            this.btModop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btModop_Click);
            // 
            // btAPropos
            // 
            this.btAPropos.Image = ((System.Drawing.Image)(resources.GetObject("btAPropos.Image")));
            this.btAPropos.Label = "À propos";
            this.btAPropos.Name = "btAPropos";
            this.btAPropos.ShowImage = true;
            this.btAPropos.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btAPropos_Click);
            // 
            // btCryptMdp
            // 
            this.btCryptMdp.Label = "btChiffre";
            this.btCryptMdp.Name = "btCryptMdp";
            this.btCryptMdp.Visible = false;
            this.btCryptMdp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btCryptMdp_Click);
            // 
            // bgwTreadDwnLoad
            // 
            this.bgwTreadDwnLoad.WorkerReportsProgress = true;
            this.bgwTreadDwnLoad.WorkerSupportsCancellation = true;
            this.bgwTreadDwnLoad.DoWork += new System.ComponentModel.DoWorkEventHandler(this.BgwTreadDwnLoad_DoWork);
            this.bgwTreadDwnLoad.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.BgwTreadDwnLoad_ProgressChanged);
            this.bgwTreadDwnLoad.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.BgwTreadDwnLoad_RunWorkerCompleted);
            // 
            // RibbonMesArchivesExplorer
            // 
            this.Name = "RibbonMesArchivesExplorer";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tabRestorePST);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonDownloadExplorer_Load);
            this.tabRestorePST.ResumeLayout(false);
            this.tabRestorePST.PerformLayout();
            this.grRestauration_av.ResumeLayout(false);
            this.grRestauration_av.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabRestorePST;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grRestauration_av;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn5;
        private System.ComponentModel.BackgroundWorker bgwTreadDwnLoad;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn9;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn10;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btCryptMdp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btAPropos;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btAccueil;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btModop;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonMesArchivesExplorer RibbonDownloadExplorer
        {
            get { return this.GetRibbon<RibbonMesArchivesExplorer>(); }
        }
    }
}
