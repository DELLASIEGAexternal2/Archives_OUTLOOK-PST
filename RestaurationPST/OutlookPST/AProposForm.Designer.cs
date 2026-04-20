namespace OutlookPST
{
    partial class AProposForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AProposForm));
            this.grpBoxAPropos = new System.Windows.Forms.GroupBox();
            this.lkLblMailToADR = new System.Windows.Forms.LinkLabel();
            this.label1 = new System.Windows.Forms.Label();
            this.picBoxSprint = new System.Windows.Forms.PictureBox();
            this.lblNumeroVersionValeur = new System.Windows.Forms.Label();
            this.lblNumeroVersion = new System.Windows.Forms.Label();
            this.grpBoxAPropos.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picBoxSprint)).BeginInit();
            this.SuspendLayout();
            // 
            // grpBoxAPropos
            // 
            this.grpBoxAPropos.Controls.Add(this.lkLblMailToADR);
            this.grpBoxAPropos.Controls.Add(this.label1);
            this.grpBoxAPropos.Controls.Add(this.picBoxSprint);
            this.grpBoxAPropos.Controls.Add(this.lblNumeroVersionValeur);
            this.grpBoxAPropos.Controls.Add(this.lblNumeroVersion);
            this.grpBoxAPropos.Location = new System.Drawing.Point(17, 19);
            this.grpBoxAPropos.Margin = new System.Windows.Forms.Padding(4);
            this.grpBoxAPropos.Name = "grpBoxAPropos";
            this.grpBoxAPropos.Padding = new System.Windows.Forms.Padding(4);
            this.grpBoxAPropos.Size = new System.Drawing.Size(697, 391);
            this.grpBoxAPropos.TabIndex = 2;
            this.grpBoxAPropos.TabStop = false;
            this.grpBoxAPropos.Text = "Information";
            // 
            // lkLblMailToADR
            // 
            this.lkLblMailToADR.AutoSize = true;
            this.lkLblMailToADR.Location = new System.Drawing.Point(17, 368);
            this.lkLblMailToADR.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lkLblMailToADR.Name = "lkLblMailToADR";
            this.lkLblMailToADR.Size = new System.Drawing.Size(259, 16);
            this.lkLblMailToADR.TabIndex = 4;
            this.lkLblMailToADR.TabStop = true;
            this.lkLblMailToADR.Text = "SPRINT 2116 POLE-ADR (Unité de travail)";
            this.lkLblMailToADR.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lkLblMailToADR_LinkClicked);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(13, 335);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(493, 17);
            this.label1.TabIndex = 3;
            this.label1.Text = "DGSI/DIEM/SPRINT, Pôle Aglité et Développements Rapides (ADR)";
            // 
            // picBoxSprint
            // 
            this.picBoxSprint.Image = ((System.Drawing.Image)(resources.GetObject("picBoxSprint.Image")));
            this.picBoxSprint.Location = new System.Drawing.Point(13, 46);
            this.picBoxSprint.Margin = new System.Windows.Forms.Padding(4);
            this.picBoxSprint.Name = "picBoxSprint";
            this.picBoxSprint.Size = new System.Drawing.Size(583, 281);
            this.picBoxSprint.TabIndex = 2;
            this.picBoxSprint.TabStop = false;
            // 
            // lblNumeroVersionValeur
            // 
            this.lblNumeroVersionValeur.AutoSize = true;
            this.lblNumeroVersionValeur.Location = new System.Drawing.Point(153, 26);
            this.lblNumeroVersionValeur.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblNumeroVersionValeur.Name = "lblNumeroVersionValeur";
            this.lblNumeroVersionValeur.Size = new System.Drawing.Size(121, 16);
            this.lblNumeroVersionValeur.TabIndex = 1;
            this.lblNumeroVersionValeur.Text = "Numéro de version";
            // 
            // lblNumeroVersion
            // 
            this.lblNumeroVersion.AutoSize = true;
            this.lblNumeroVersion.Location = new System.Drawing.Point(9, 25);
            this.lblNumeroVersion.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblNumeroVersion.Name = "lblNumeroVersion";
            this.lblNumeroVersion.Size = new System.Drawing.Size(127, 16);
            this.lblNumeroVersion.TabIndex = 0;
            this.lblNumeroVersion.Text = "Numéro de version :";
            // 
            // AProposForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(729, 422);
            this.Controls.Add(this.grpBoxAPropos);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "AProposForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "À propos";
            this.grpBoxAPropos.ResumeLayout(false);
            this.grpBoxAPropos.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picBoxSprint)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox grpBoxAPropos;
        public System.Windows.Forms.LinkLabel lkLblMailToADR;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox picBoxSprint;
        private System.Windows.Forms.Label lblNumeroVersionValeur;
        private System.Windows.Forms.Label lblNumeroVersion;
    }
}