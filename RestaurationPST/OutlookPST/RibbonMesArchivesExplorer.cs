using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Reflection;
using System.IO;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar;
using System.Threading;
using static OutlookPST.Tools;
using System.Security.Cryptography;
using System.Windows.Forms;
using System.Security.Cryptography.X509Certificates;

namespace OutlookPST
{
    public partial class RibbonMesArchivesExplorer
    {
        //private ToastForm toast;
        public ToastForm toast;

        /// <summary>
        /// Chargement du Ruban
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RibbonDownloadExplorer_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(Const.GetAccueil))
                {
                    btAccueil.Visible = false;
                }
                if (string.IsNullOrEmpty(Const.GetModop))
                {
                    btModop.Visible = false;
                }

                // Création des dossiers pour les PST et fichier Flg (Ruban) si ils n'existe pas
                Traitement.createFolderForPSTandFlag();

                // Chargement des lst Flag Ruban pour la boite courante
                Traitement.readFlagPST(Const.currentMailBox, Const.lstFlagPST);

                // Recupération des PST dans la session si isConnected
                if (Tools.isConnected())
                {
                    Const.isConnected = true;
                    Traitement.PSTInSession();
                }

                //Actualisation bt du ruban
                Traitement.refreshRuban(this, Const.currentMailBox, Const.lstFlagRuban);
            }
            catch (Exception ex)
            {
                Tools.LogMessage("Erreur lors du chargement du ruban." + MethodBase.GetCurrentMethod().Name, ex, true);
            }
        }

        /// <summary>
        /// tache de lancement du traitement
        /// </summary>
        private void TaskTreatement(object sender)
        {
            try
            {
                //Const.isConnnectMode=Tools.ModeConnected();

                //// Limite OUTLOOK a 15 pst
                //if (Const.lstPST_InSession.Count >= Const.limitNbPST)
                //{
                //    System.Windows.Forms.MessageBox.Show(MessageClass.LimiteNombrePst,"Alert - Limite OUTLOOK",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Warning);
                //    return;
                //}


                // Création de la tâche de téléchargement si un téléchargement n'est pas déja lancé
                if (this.bgwTreadDwnLoad.IsBusy)
                {
                    // Téléchargement en cours
                    System.Windows.Forms.MessageBox.Show(MessageClass.message_Traitement_En_Cours, "Information :", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }
                else
                {
                    Const.stopThread = false;
                    // Init les informations de la demande
                    bool res = Traitement.initDemande(((RibbonButton)sender).Name, Const.param);
                    if (res)
                    {
                        // Doit Ont Lancer le traitement ou exploiter les fichier Local
                        // retour du Choix = Oui => Utilisation du Fichier Local + refresh
                        // retour du choix = Annuler -> Refresh
                        Tools.ChoixUseLocalPST choix = Traitement.useLocalPST(Const.param.btDemande, Const.param.anneeDemande);
                        
                        // ANNULATION
                        if (choix == Tools.ChoixUseLocalPST.Annule)
                        {
                            Traitement.refreshRuban(this, Const.param.btDemande, Const.lstFlagRuban);
                            return;
                        }

                        // UTILISATION LOCAL 
                        if (choix == Tools.ChoixUseLocalPST.Oui)
                        {
                            Traitement.refreshRuban(this, Const.param.btDemande, Const.lstFlagRuban);
                            return;
                        }

                        // TELECHARGEMENT
                        // Si Choix = Non pour téléchargement -> sauf en Offligne
                        if (choix == Tools.ChoixUseLocalPST.Non && !Tools.isConnected())
                        {
                            Traitement.refreshRuban(this, Const.param.btDemande, Const.lstFlagRuban);
                            System.Windows.Forms.MessageBox.Show("Fonction non disponible en mode Hors ligne.", "Mode Hors Ligne", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }

                        if (choix == Tools.ChoixUseLocalPST.Non && Tools.isConnected())
                        {
                            if (Traitement.IsMountInSession(Const.param.btDemande,Const.param.anneeDemande,true) || Traitement.IsMountInSession(Const.param.btDemande, Const.param.anneeDemande, false))
                            {
                                System.Windows.Forms.MessageBox.Show(MessageClass.message_PSTCharger, "PST de l'année "+ Const.param.anneeDemande,  MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                Traitement.refreshRuban(this, Const.param.btDemande, Const.lstFlagRuban);
                                return;
                            }
                        }

                        // Sinon Tools.ChoixUseLocalPST.Non -> Téléchargement 

                        // Recherche le type de retention et l'enregistre si absent
                        Traitement.isTenYear(Const.param.btDemande);

                        string commentaire = string.Format(MessageClass.message_Commentaire_Start, DateTime.Now.ToString());

                        // Passe le statut du flag pour la boite principal en Encours
                        // Création ou Modification du Flags Ruban pour la boite demandé et l'année choisie
                        if (!Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban,Const.param.btDemande,Const.param.anneeDemande,Const.param.archiveDemande, Statut_ruban.Encours, commentaire))
                            return;

                        // Passe le statut du flag pour la boite Archive en Encours
                        Const.param.archiveDemande = true;
                        // Création ou Modification du Flags Ruban pour la boite demandé et l'année choisie
                        if (!Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_ruban.Encours, commentaire))
                            return;

                        // Passe le statut du flag PST pour la boite principal en Encours
                        // Création ou Modification du Flags Ruban pour la boite demandé et l'année choisie
                        //if (!Traitement.createOrUpdateFlagPST(Const.param.btDemande, Const.lstFlagPST, Const.param, Statut_PST.Encours, commentaire))
                        if (!Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_PST.Encours, commentaire))

                                return;
                        Const.param.archiveDemande = true;
                        // Création ou Modification du Flags Ruban pour la boite demandé et l'année choisie
                        //if (!Traitement.createOrUpdateFlagPST(Const.param.btDemande, Const.lstFlagPST, Const.param, Statut_PST.Encours, commentaire))
                        if (!Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_PST.Encours, commentaire))
                                return;

                        Const.param.archiveDemande = false;
                        // Lancement en tache de fonds du traitement de download et info a l'utilisateur
                        this.toast = new ToastForm(Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Const.param.btDemande_samAccountName);
                        //Tools.ToastsMessage = "Téléchargement en cours.";
                        this.toast.Show();
                        this.bgwTreadDwnLoad.RunWorkerAsync();

                    }
                }

            }
            catch 
            {
                throw ;
            }
        }


        public RibbonButton GetBtn(string Annee)
        {
            RibbonButton bt = null;
            try
            {
                switch (Annee)
                {
                    case "2015":
                        bt = btn1;
                        break;
                    case "2016":
                        bt = btn2;
                        break;
                    case "2017":
                        bt = btn3;
                        break;
                    case "2018":
                        bt = btn4;
                        break;
                    case "2019":
                        bt = btn5;
                        break;
                    case "2020":
                        bt = btn6;
                        break;
                    case "2021":
                        bt = btn7;
                        break;
                    case "2022":
                        bt = btn8;
                        break;
                    case "2023":
                        bt = btn9;
                        break;
                    case "2024":
                        bt = btn10;
                        break;
                     default:
                        bt = btn1;
                        break;
                }
            }
            catch (Exception ex)
            {
                Tools.LogMessage("Erreur lors de la récupération du bouton suivant la date.", ex, true, MethodBase.GetCurrentMethod().Name);
            }
            return bt;
        }

        private void BgwTreadDwnLoad_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            try
            {
                var resultat=Download.MainDownload(sender,toast);                
            }
            catch (Exception ex)
            {
                Tools.LogMessage("Erreur de la tâche de téléchargement.", ex,true, MethodBase.GetCurrentMethod().Name);
            }
        }

        private void BgwTreadDwnLoad_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            this.toast.SetMessage((string)e.UserState, e.ProgressPercentage);
        }

        private void BgwTreadDwnLoad_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
                this.toast.SetMessage("Erreur : " + e.Error.Message, 0);
            else if (e.Cancelled  || Const.stopThread == true)
            {
                this.toast.SetMessage("Annulation des traitements en tâches de fond.", 0);

                Const.param.statutRubanDemande = Statut_ruban.Annule;
                Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, false, Statut_ruban.Annule, MessageClass.message_StocObj_Download_Cancel);
                Const.param.statutPSTDemande = Statut_PST.Annule;
                Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, false, Statut_PST.Annule, MessageClass.message_StocObj_Download_Cancel);

                Const.param.statutRubanDemande = Statut_ruban.Annule;
                Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, true, Statut_ruban.Annule, MessageClass.message_StocObj_Download_Cancel);
                Const.param.statutPSTDemande = Statut_PST.Annule;
                Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, true, Statut_PST.Annule, MessageClass.message_StocObj_Download_Cancel);

            }

            Tools.LogMessage(string.Format("Terminé : Boîte {0}_{1} - statut {2}. ", Const.param.btDemande, Const.param.anneeDemande,Const.param.statutPSTDemande), null, false, MethodBase.GetCurrentMethod().Name);

            //Avant d'actualiser 
            Traitement.useLocalPST(Const.param.btDemande, Const.param.anneeDemande,true);


            //Actualisation ruban
            Traitement.refreshRuban(this,Const.param.btDemande, Const.lstFlagRuban);

            Globals.ThisAddIn.Application.Session.RefreshRemoteHeaders();
            this.toast.PrepareClose();

            // Nettoyage du repertoire de Download
            if (Directory.Exists(Const.tmp7z))
                Directory.Delete(Const.tmp7z, true);

            Const.stopThread = false;
        }

        private void btn1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                TaskTreatement(sender);
            }
            catch (Exception ex)
            {
                Tools.LogMessage("Erreur lors du traitement de l'année 2015.", ex, true, MethodBase.GetCurrentMethod().Name);
            }
        }

        private void btn2_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                TaskTreatement(sender);
            }
            catch (Exception ex)
            {
                Tools.LogMessage("Erreur lors du traitement de l'année 2016.", ex, true, MethodBase.GetCurrentMethod().Name);
            }

        }

        private void btn3_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                TaskTreatement(sender);
            }
            catch (Exception ex)
            {
                Tools.LogMessage("Erreur lors du traitement de l'année 2017.", ex, true, MethodBase.GetCurrentMethod().Name);
            }

        }

        private void btn4_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                TaskTreatement(sender);
            }
            catch (Exception ex)
            {
                Tools.LogMessage("Erreur lors du traitement de l'année 2018.", ex, true, MethodBase.GetCurrentMethod().Name);
            }

        }

        private void btn5_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                TaskTreatement(sender);
            }
            catch (Exception ex)
            {
                Tools.LogMessage("Erreur lors du traitement de l'année 2019.", ex, true, MethodBase.GetCurrentMethod().Name);
            }

        }

        private void btn6_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                TaskTreatement(sender);
            }
            catch (Exception ex)
            {
                Tools.LogMessage("Erreur lors du traitement de l'année 2020.", ex, true, MethodBase.GetCurrentMethod().Name);
            }

        }

        private void btn7_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                TaskTreatement(sender);
            }
            catch (Exception ex)
            {
                Tools.LogMessage("Erreur lors du traitement de l'année 2021.", ex, true, MethodBase.GetCurrentMethod().Name);
            }

        }

        private void btn8_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                TaskTreatement(sender);
            }
            catch (Exception ex)
            {
                Tools.LogMessage("Erreur lors du traitement de l'année 2022.", ex, true, MethodBase.GetCurrentMethod().Name);
            }

        }

        private void btn9_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                TaskTreatement(sender);
            }
            catch (Exception ex)
            {
                Tools.LogMessage("Erreur lors du traitement de l'année 2023.", ex, true, MethodBase.GetCurrentMethod().Name);
            }

        }

        private void btn10_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                TaskTreatement(sender);
            }
            catch (Exception ex)
            {
                Tools.LogMessage("Erreur lors du traitement de l'année 2024.", ex, true, MethodBase.GetCurrentMethod().Name);
            }

        }

        private void btCryptMdp_Click(object sender, RibbonControlEventArgs e)
        {
            string str = "zZ@US3(q8AkRXaF)6+O,!KV9:";
            byte[] encData_byte = new byte[str.Length];
            encData_byte = System.Text.Encoding.UTF8.GetBytes(str);
            string encodedData = Convert.ToBase64String(encData_byte);
            //La clé publique est utilisée pour chiffrer les données
        }

        private void btAPropos_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                AProposForm aProposForm = new AProposForm();
                aProposForm.ShowDialog();
            }
            catch (Exception ex)
            {
                Tools.LogMessage("Erreur lors de l'ouverture de la fenêtre \"À propos\" - Methode : " + MethodBase.GetCurrentMethod().Name, ex, true);
                MessageBox.Show("Erreur au chargement de À PROPOS" + ex.InnerException.ToString(), "Attention / Erreur ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btAccueil_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(Const.GetAccueil);
            }
            catch (Exception ex)
            {
                Tools.LogMessage("Erreur lors de l'ouverture de la fenêtre \"Accueil\" - Methode : " + MethodBase.GetCurrentMethod().Name, ex, true);
            }
        }

        private void btModop_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(Const.GetModop);
            }
            catch (Exception ex)
            {
                Tools.LogMessage("Erreur lors de l'ouverture de la fenêtre \"Modop\" - Methode : " + MethodBase.GetCurrentMethod().Name, ex, true);
            }
        }
    }
}
