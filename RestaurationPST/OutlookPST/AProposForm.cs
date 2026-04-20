using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookPST
{
    public partial class AProposForm : Form
    {
        public AProposForm()
        {
            InitializeComponent();
            this.lblNumeroVersionValeur.Text = Tools.GetAppVersion();
        }

        private void lkLblMailToADR_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Close();
            System.Diagnostics.Process.Start(string.Format("mailto:{0}", this.lkLblMailToADR.Text));
        }
    }
}
