namespace OutlookPST
{
    partial class ToastForm
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
            this.components = new System.ComponentModel.Container();
            this.lifeTimer = new System.Windows.Forms.Timer(this.components);
            this.lblMsgImportant = new System.Windows.Forms.Label();
            this.pgbProgret = new System.Windows.Forms.ProgressBar();
            this.lblMessage = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lifeTimer
            // 
            this.lifeTimer.Tick += new System.EventHandler(this.lifeTimer_Tick);
            // 
            // lblMsgImportant
            // 
            this.lblMsgImportant.BackColor = System.Drawing.Color.WhiteSmoke;
            this.lblMsgImportant.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblMsgImportant.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.lblMsgImportant.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMsgImportant.ForeColor = System.Drawing.Color.Red;
            this.lblMsgImportant.Location = new System.Drawing.Point(0, 0);
            this.lblMsgImportant.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblMsgImportant.Name = "lblMsgImportant";
            this.lblMsgImportant.Size = new System.Drawing.Size(820, 29);
            this.lblMsgImportant.TabIndex = 0;
            this.lblMsgImportant.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // pgbProgret
            // 
            this.pgbProgret.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.pgbProgret.Location = new System.Drawing.Point(0, 89);
            this.pgbProgret.Margin = new System.Windows.Forms.Padding(4);
            this.pgbProgret.Name = "pgbProgret";
            this.pgbProgret.Size = new System.Drawing.Size(820, 26);
            this.pgbProgret.TabIndex = 1;
            this.pgbProgret.Click += new System.EventHandler(this.pgbProgret_Click);
            // 
            // lblMessage
            // 
            this.lblMessage.BackColor = System.Drawing.Color.WhiteSmoke;
            this.lblMessage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblMessage.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.lblMessage.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMessage.ForeColor = System.Drawing.Color.Black;
            this.lblMessage.Location = new System.Drawing.Point(0, 29);
            this.lblMessage.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblMessage.Name = "lblMessage";
            this.lblMessage.Size = new System.Drawing.Size(820, 60);
            this.lblMessage.TabIndex = 2;
            this.lblMessage.Text = "Début traitement.";
            this.lblMessage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // ToastForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(820, 115);
            this.Controls.Add(this.lblMessage);
            this.Controls.Add(this.lblMsgImportant);
            this.Controls.Add(this.pgbProgret);
            this.ForeColor = System.Drawing.SystemColors.Highlight;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "ToastForm";
            this.Text = "Mon archive mail";
            this.TopMost = true;
            this.Activated += new System.EventHandler(this.ToastForm_Activated);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ToastForm_FormClosing);
            this.Load += new System.EventHandler(this.ToastForm_Load);
            this.Shown += new System.EventHandler(this.ToastForm_Shown);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Timer lifeTimer;
        private volatile System.Windows.Forms.Label lblMsgImportant;
        private System.Windows.Forms.ProgressBar pgbProgret;
        private System.Windows.Forms.Label lblMessage;
    }
}