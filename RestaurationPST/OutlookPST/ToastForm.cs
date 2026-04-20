using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace OutlookPST
{
    public partial class ToastForm : Form
    {
        public bool stopThread;
        // Alert sur le traitement
        //public string MsgAlert;

        /// <summary> Indicates whether the form can receive focus or not.
        /// </summary>
        private bool allowFocus = false;

        /// <summary> The object that creates the sliding animation.
        /// </summary>
        private FormAnimator animator;

        /// <summary> The handle of the window that currently has focus.
        /// </summary>
        private IntPtr currentForegroundWindow;

        /// <summary> Gets the handle of the window that currently has focus.
        /// </summary>
        /// <returns> The handle of the window that currently has focus.</returns>
        [DllImport("user32")]
        private static extern IntPtr GetForegroundWindow();

        /// <summary> Activates the specified window.
        /// </summary>
        /// <param name="hWnd"> The handle of the window to be focused.
        /// </param>
        /// <returns> True if the window was focused; False otherwise.</returns>
        [DllImport("user32")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        /// <summary> Creates a new ToastForm object that is displayed for the specified length of time.
        /// </summary>
        /// <param name="lifeTime"> The length of time, in milliseconds, that the form will be displayed.
        /// </param>
        internal ToastForm(string boiteDemande, string AnneeDemande, bool Archive, string samBtDemande)
        {
            this.InitializeComponent();
            this.Height = 100;

            // Display the form by sliding up.
            //this.animator = new FormAnimator(this, FormAnimator.AnimationMethod.Slide, FormAnimator.AnimationDirection.Up, 500);
            this.animator = new FormAnimator(this, FormAnimator.AnimationMethod.Slide, FormAnimator.AnimationDirection.Up, 200);
            this.pgbProgret.Maximum = 100;
            this.pgbProgret.Minimum = 0;

            if (boiteDemande.StartsWith(@"\\"))
                boiteDemande = boiteDemande.Remove(0, 2);
            if (Archive)
            {
                this.Text = string.Format("Téléchargement : {0} pour l'année {1} - Archive.", boiteDemande, AnneeDemande);
            }
            else
            {
                this.Text = string.Format("Téléchargement : {0} pour l'année {1}.", boiteDemande, AnneeDemande);
            }
        }

        internal void SetMessage(string message, int progression)
        {
            //if (!string.IsNullOrEmpty(MsgAlert))
            //    this.lblMsgImportant.Text = MsgAlert;
            if (!string.IsNullOrEmpty(Const.ToastsMessage))
                this.lblMsgImportant.Text = Const.ToastsMessage;

            this.lblMessage.Text = message;
            if (progression <= this.pgbProgret.Maximum)
                this.pgbProgret.Value = progression;
        }

        internal void SetTitre(string message)
        {
            this.Text = message;
        }

        /// <summary> Displays the form.
        /// </summary>
        /// <remarks> Required to allow the form to determine the current foreground window     before being displayed.
        /// </remarks>
        internal new void Show()
        {
            // Determine the current foreground window so it can be reactivated each time this form tries to get the focus.
            this.currentForegroundWindow = GetForegroundWindow();
            
            // Display the form.
            base.Show();
        }

        internal void PrepareClose()
        {            
            this.lifeTimer.Stop();
            this.lifeTimer.Interval = 20; // 2000;
            this.lifeTimer.Start();
        }

        private void ToastForm_Load(object sender, EventArgs e)
        {
            // Display the form just above the system tray.
            this.Location = new Point(Screen.PrimaryScreen.WorkingArea.Width - this.Width - 5,
                                      Screen.PrimaryScreen.WorkingArea.Height - this.Height - 5);
        }

        private void ToastForm_Activated(object sender, EventArgs e)
        {
            // Prevent the form taking focus when it is initially shown.
            if (!this.allowFocus)
            {
                // Activate the window that previously had the focus.
                SetForegroundWindow(this.currentForegroundWindow);
            }
        }

        private void ToastForm_Shown(object sender, EventArgs e)
        {
            // Once the animation has completed the form can receive focus.
            this.allowFocus = true;

            // Close the form by sliding down.
            this.animator.Direction = FormAnimator.AnimationDirection.Down;
        }

        /// <summary> La durée de vie de la forme a expiré
        /// </summary>
        private void lifeTimer_Tick(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ToastForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            Const.stopThread = true;
            if (Const.cancellationTokenSourceForDownloads != null && Const.downloadsInUseThread)
            {
                if (!Const.cancellationTokenSourceForDownloads.IsCancellationRequested)
                    Const.cancellationTokenSourceForDownloads.Cancel(); 
            }
            
        }

        private void pgbProgret_Click(object sender, EventArgs e)
        {

        }
    }
}
